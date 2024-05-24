using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Security;
using DataBotV5.Data.Root;
using DataBotV5.Data.Projects.Freelance;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.Freelance;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;
using System.Text;
using System.IO;
using DataBotV5.Logical.MicrosoftTools;
using System.Security;
using System.Net;

namespace DataBotV5.Automation.WEB.Freelance
{
    /// <summary>
    /// Clase WEB Automation encargada de reporte de horas y aprobaciones Freelance.
    /// </summary>
    class FreelanceSS
    {
        Rooting root = new Rooting();
        ProcessInteraction process = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        CRUD crud = new CRUD();
        FreelanceSqlSS freelanceSql = new FreelanceSqlSS();
        FreelanceFiSS freelanceFi = new FreelanceFiSS();
        string mandante = "ERP";
        string ssMandante = "PRD";
        Credentials cred = new Credentials();
        SecureAccess sec = new SecureAccess();
        Database db = new Database();
        MailInteraction mail = new MailInteraction();
        WebexTeams webex = new WebexTeams();
        MsExcel excel = new MsExcel();
        /// <summary>
        /// Metodo Main de la clase, ejecuta la creacion de tiempos, modificacion de tiempos, aprobaciones, creaciones de hojas
        /// </summary>
        public void Main()
        {
            // Reporte de horas
            console.WriteLine("Reportando Horas");
            ReportTimes();
            // Envio de factura a contabilidad
            console.WriteLine("Enviar factura a contabilidad");
            sendBill();
            // Crear Hoja
            //console.WriteLine("Creando hojas");
            //GenerateSheet();

        }
        /// <summary>
        /// Método de reporte de tiempos.
        /// </summary>
        private void ReportTimes()
        {
            ///Extrae las filas de la tabla, que se encuentre en estado procesando aprobacion (significa que ya el coordinador aprobo)
            DataTable tiempos = freelanceSql.GetxState("13"); //PROCESANDO-AP 
            string cats = "";
            if (tiempos != null)
            {
                if (tiempos.Rows.Count > 0)
                {

                    Dictionary<string, string> msjErrs = new Dictionary<string, string>();

                    for (int i = 0; i < tiempos.Rows.Count; i++)
                    {
                        string errorMsj = "";
                        string id = tiempos.Rows[i]["id"].ToString();
                        string empleado = tiempos.Rows[i]["createdBy"].ToString();
                        string po = tiempos.Rows[i]["purchaseOrder"].ToString();
                        string item = tiempos.Rows[i]["item"].ToString();
                        string fecha = tiempos.Rows[i]["reportDate"].ToString();
                        string ticket = tiempos.Rows[i]["ticket"].ToString();
                        string detalle = sec.DecodePass(tiempos.Rows[i]["details"].ToString());
                        string horas = tiempos.Rows[i]["hours"].ToString();
                        string cats_rec = tiempos.Rows[i]["catsId"].ToString();
                        string razon = (tiempos.Rows[i]["colabReason"].ToString() == "null") ? "" : sec.DecodePass(tiempos.Rows[i]["colabReason"].ToString());
                        string responsable = tiempos.Rows[i]["responsible"].ToString();
                        if (cats_rec != "" && cats_rec != null)
                        {
                            //es reprocesar las horas
                            cats = freelanceFi.CatsM(horas, "ZGBM02", cats_rec, ticket, detalle);
                            if (cats == "OK")
                            {
                                console.WriteLine(empleado + "\t" + po + "\t" + item + "\t" + fecha + "\t" + ticket + "\t" + horas + "\t" + cats);
                                string update_query = "UPDATE hourReportFreelance SET status = '1', sapAt = CURRENT_TIMESTAMP WHERE id = '" + id + "'";
                                crud.Update(update_query, "freelance_db");

                                string aprobacion = freelanceFi.CatsAppr("A", cats_rec, razon, "ZGBM02");

                                if (aprobacion == "OK" || aprobacion == "NO-CHANGE")
                                {
                                    //Cat creada y aprobado 
                                    console.WriteLine(cats_rec + "\t" + "Aprobado");
                                    string uQuery = "UPDATE hourReportFreelance SET catsId = '" + cats_rec + "', status = '1', sapAt = CURRENT_TIMESTAMP WHERE id = '" + id + "'"; //APROBADO
                                    crud.Update(uQuery, "freelance_db");
                                    //webex.SendNotification("hogonzalez@gbm.net", "Reporte de Horas", $"Horas cargadas en SAP, número de Cats: {reporte.CatId}");
                                }
                                else
                                {
                                    //error al aprobar
                                    console.WriteLine(cats_rec + "\t" + "Error de aprobacion");
                                    string sapError = Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(aprobacion));
                                    string updateQuery = $"UPDATE hourReportFreelance SET  catsId = '{cats_rec}', status = '7', sapAt = CURRENT_TIMESTAMP, sapError = '{sapError}' WHERE id = '" + id + "'"; //ERROR CATS
                                    crud.Update(updateQuery, "freelance_db");
                                    if (msjErrs.ContainsKey(responsable))
                                    {
                                        msjErrs[responsable] = msjErrs[responsable] + "<br><br>" + $"Error en PO {po}/{item} en la fecha {fecha}: " + aprobacion;

                                    }
                                    else
                                    {
                                        msjErrs[responsable] = $"Error en PO {po}/{item} en la fecha {fecha}: " + aprobacion;
                                    }
                                }
                            }
                            else
                            {
                                string sapError = Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(cats));
                                string update_query = $"UPDATE hourReportFreelance SET status = '4', sapAt = CURRENT_TIMESTAMP, sapError = '{sapError}' WHERE id = '" + id + "'";
                                crud.Update(update_query, "freelance_db");
                                if (msjErrs.ContainsKey(responsable))
                                {
                                    msjErrs[responsable] = msjErrs[responsable] + "<br><br>" + $"Error en PO {po}/{item} en la fecha {fecha}: " + cats;

                                }
                                else
                                {
                                    msjErrs[responsable] = $"Error en PO {po}/{item} en la fecha {fecha}: " + cats;
                                }
                            }
                        }
                        else
                        {
                            //es un reporte normal de horas
                            CatInfo reporte = freelanceFi.Cats(po, item, horas, fecha, empleado, detalle, ticket, "ZGBM02");
                            console.WriteLine(reporte.RespError);
                            if (reporte.CatId != null && reporte.CatId != "")
                            {
                                //todo salio bien
                                console.WriteLine(empleado + "\t" + po + "\t" + item + "\t" + fecha + "\t" + ticket + "\t" + horas + "\t" + reporte.CatId);

                                //aprobar Cats
                                string aprobacion = freelanceFi.CatsAppr("A", reporte.CatId, razon, "ZGBM02");

                                if (aprobacion == "OK" || aprobacion == "NO-CHANGE")
                                {
                                    //Cat creada y aprobado 
                                    console.WriteLine(reporte.CatId + "\t" + "Aprobado");
                                    string update_query = "UPDATE hourReportFreelance SET catsId = '" + reporte.CatId + "', status = '1', sapAt = CURRENT_TIMESTAMP WHERE id = '" + id + "'"; //APROBADO
                                    crud.Update(update_query, "freelance_db");
                                    //webex.SendNotification("hogonzalez@gbm.net", "Reporte de Horas", $"Horas cargadas en SAP, número de Cats: {reporte.CatId}");
                                }
                                else
                                {
                                    //error al aprobar
                                    console.WriteLine(reporte.CatId + "\t" + "Error de aprobacion");
                                    string sapError = Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(aprobacion));
                                    string updateQuery = $"UPDATE hourReportFreelance SET  catsId = '{reporte.CatId}', status = '7', sapAt = CURRENT_TIMESTAMP, sapError = '{sapError}' WHERE id = '" + id + "'"; //ERROR CATS
                                    crud.Update(updateQuery, "freelance_db");
                                    string html = Properties.Resources.emailtemplate1;
                                    html = html.Replace("{subject}", "Portal Proveedores: error al aprobar CATS");
                                    html = html.Replace("{cuerpo}", $"Se dio un error al aprobar el CATS {reporte.CatId} debido a: {aprobacion}");
                                    html = html.Replace("{contenido}", "");
                                    mail.SendHTMLMail(html, new string[] { responsable }, "Portal Proveedores: error al aprobar CATS", new string[] { "dmeza@gbm.net" }, null);
                                }
                            }
                            else
                            {
                                //error al crear el CATS
                                string sapError = Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(reporte.RespError));
                                console.WriteLine(empleado + "\t" + po + "\t" + item + "\t" + fecha + "\t" + ticket + "\t" + horas + "\t" + "Error de reporte en CATS");
                                string update_query = $"UPDATE hourReportFreelance SET status = '4', sapAt = CURRENT_TIMESTAMP, sapError = '{sapError}' WHERE id = '" + id + "'"; //ERROR CATS
                                crud.Update(update_query, "freelance_db");

                                if (msjErrs.ContainsKey(responsable))
                                {
                                    msjErrs[responsable] = msjErrs[responsable] + "<br><br>" + $"Error en PO {po}/{item} en la fecha {fecha}: " + reporte.RespError;

                                }
                                else
                                {
                                    msjErrs[responsable] = $"Error en PO {po}/{item} en la fecha {fecha}: " + reporte.RespError;
                                }

                            }
                        }
                    } //for
                    foreach (var kvp in msjErrs)
                    {
                        string key = kvp.Key;
                        string value = kvp.Value;
                        string html = Properties.Resources.emailtemplate1;
                        html = html.Replace("{subject}", "Portal Proveedores: error al crear CATS de las siguientes PO");
                        html = html.Replace("{cuerpo}", value);
                        html = html.Replace("{contenido}", "");
                        mail.SendHTMLMail(html, new string[] { key + "@gbm.net" }, "Portal Proveedores: error al crear CATS", new string[] { "dmeza@gbm.net" }, null);
                    }

                }
            }

        }


        /// <summary>
        /// Creacion de HES y Envio de factura aprobada por el L2 al departamento de contabilidad
        /// </summary>
        public void sendBill()
        {
            if (!sap.CheckLogin(mandante))
            {
                DataTable aprobadosL2 = freelanceSql.GetBillingsL2();
                //List<FreelanceBillings> listado = new List<FreelanceBillings>();

                if (aprobadosL2 != null)
                {
                    if (aprobadosL2.Rows.Count > 0)
                    {
                        foreach (DataRow row in aprobadosL2.Rows)
                        {
                            string id = row["id"].ToString();
                            string billId = row["BillNumber"].ToString();
                            try
                            {
                                string aprobadorL2 = row["approverL2"].ToString();
                                string createdBy = row["createdBy"].ToString();


                                //Primero determinar la Organizacion y correos
                                DataTable poInfo = freelanceSql.getBillPoInfo(id);

                                //CREAR HES
                                //bloquea mandante
                                sap.BlockUser(mandante, 1);
                                sap.KillSAP(); // Si quedara abierto revise y matelo
                                sap.LogSAP(mandante.ToString());
                                bool hesResponse = false;
                                foreach (DataRow hesRow in poInfo.Rows)
                                {
                                    //var v = new { Amount = 108, Message = "Hello" };
                                    if (hesRow["hesNumber"].ToString() == "")
                                    {

                                        bool createHoja = freelanceFi.CreateSheet(hesRow); //procesar la hoja y liberarla
                                        if (!hesResponse)
                                        {
                                            hesResponse = createHoja;
                                        }

                                    }
                                    else
                                    {
                                        hesResponse = true;
                                    }
                                    crud.Update($"UPDATE hourReportFreelance SET status = 9 WHERE handleUnitId = '{hesRow["hesId"].ToString()}'", "freelance_db");
                                    //break;
                                }
                                sap.KillSAP();
                                sap.BlockUser(mandante, 0);

                                if (hesResponse)
                                {
                                    try
                                    {

                                        poInfo = freelanceSql.getBillPoInfo(id); //extraer de nuevo la info pero ya con las HES creadas

                                        //determinar los emails de conta
                                        billEmails bases = freelanceSql.getBillEmails(poInfo.Rows[0]["companyCode"].ToString()); //esta bien que tome el pais de la primera que encuentre?

                                        string table = freelanceFi.BillingTable(poInfo, createdBy);
                                        string mensaje = $@"Estimado encargado financiero, favor tramitar el pago de las facturas adjuntas en este correo, tomando en cuenta el detalle de las horas aprobadas por: {aprobadorL2} mediante el Portal Freelance de GBM.";
                                        string subject = $"Portal Proveedores: Solicitud de pago a proveedor, número de factura: {billId}";
                                        string html = Properties.Resources.emailtemplate1;
                                        html = html.Replace("{subject}", "Portal Proveedores: Solicitud de pago a proveedor");
                                        html = html.Replace("{cuerpo}", mensaje);
                                        html = html.Replace("{contenido}", table);

                                        string mysql = $"SELECT name, path FROM uploadFiles WHERE idRequest = '{id}' AND motherTable IN (2)";
                                        //if (poInfo.Rows[0]["byHito"].ToString() == "1")
                                        //{
                                        //    mysql = $"({mysql}) UNION ALL (SELECT name, path FROM uploadFiles WHERE idRequest = '{poInfo.Rows[0]["poId"].ToString()}' AND motherTable IN (4))";
                                        //}
                                        //descargar y adjuntar attachs
                                        DataTable dt = crud.Select(mysql, "freelance_db");

                                        List<string> list = new List<string>();
                                        string[] attachs = null;


                                        if (dt.Rows.Count > 0)
                                        {
                                            foreach (DataRow arch in dt.Rows)
                                            {
                                                string filePath = arch["path"].ToString();
                                                freelanceSql.downloadFile(filePath);
                                                string filePathLocal = root.FilesDownloadPath + "\\" + Path.GetFileName(filePath);
                                                list.Add(filePathLocal);
                                            }

                                            attachs = list.ToArray();
                                        }

                                        string[] newCopies = new string[bases.copies.Length + 1];
                                        Array.Copy(bases.copies, newCopies, bases.copies.Length);
                                        newCopies[newCopies.Length - 1] = poInfo.Rows[0]["responsible"].ToString() + "@gbm.net";
                                        //mail.SendHTMLMail(html, bases.senders, subject, bases.copies, attachs);
                                        mail.SendHTMLMail(html, bases.senders, subject, newCopies, attachs);

                                        crud.Update($"UPDATE billsFreelance SET status = '9' WHERE ID = '{id}'", "freelance_db"); //ACEPTADO
                                    }
                                    catch (Exception ex)
                                    {
                                        string mensaje = $"Estimado, dio error al mandar el correo a facturación de la factura {billId} por el siguiente error: </br>{ex.Message}";
                                        string subject = $"Portal Freelance: Error al enviar correo a facturación";
                                        string html = Properties.Resources.emailtemplate1;
                                        html = html.Replace("{subject}", "Error al enviar correo a facturación");
                                        html = html.Replace("{cuerpo}", mensaje);
                                        html = html.Replace("{contenido}", "");

                                        //mail.SendNotificationErrorSheet($"", mensaje, "Portal Freelance: Error de HES (Creacion)");
                                        mail.SendHTMLMail(html, new string[] { "hogonzalez@gbm.net" }, subject, new string[] { "dmeza@gbm.net" }, null);
                                        crud.Update($"UPDATE billsFreelance SET status = '6' WHERE ID = '{id}'", "freelance_db"); //ERROR
                                    }
                                }
                                else
                                {
                                    crud.Update($"UPDATE billsFreelance SET status = '6' WHERE ID = '{id}'", "freelance_db"); //ERROR
                                }
                            }
                            catch (Exception EXS)
                            {
                                sap.BlockUser(mandante, 0);
                                sap.KillSAP();
                                string mensaje = $"Estimado coordinador, dio error al intentar inesperado al crear la HES de la factura {billId},  </br> {EXS.Message} </br> {EXS.StackTrace}";
                                string subject = $"Portal Proveedores: Error inesperado de facturación";
                                string html = Properties.Resources.emailtemplate1;
                                html = html.Replace("{subject}", "Portal Proveedores: Error inesperado de facturación");
                                html = html.Replace("{cuerpo}", mensaje);
                                html = html.Replace("{contenido}", "");
                                mail.SendHTMLMail(html, new string[] { "hogonzalez@gbm.net" }, subject, new string[] { "dmeza@gbm.net" }, null);
                                crud.Update($"UPDATE billsFreelance SET status = '6' WHERE ID = '{id}'", "freelance_db"); //ERROR
                            }

                            break; //solo haga el primero que encuentre

                        }
                    }
                }


            }
        }

        /// <summary>
        /// Metodo para crear Hojas
        /// </summary>
        public void GenerateSheet()
        {

            DataTable hojas = freelanceSql.GetxSheet();

            //DataTable Dproveedores = crud.Select("Databot", "SELECT USUARIO,CORREO,PROVEEDOR,CONSULTORES FROM freelance_a WHERE ACTIVO ='X'", "automation");
            //List<HESAP> hes = new List<HESAP>();
            //List<FreelanceVendors> proveedores = new List<FreelanceVendors>();

            if (hojas != null)
            {
                if (hojas.Rows.Count > 0)
                {
                    //CREAR HES 
                    while (sap.CheckLogin(mandante))
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    sap.BlockUser(mandante, 1); //bloquea mandante
                    foreach (DataRow hesRow in hojas.Rows)
                    {
                        freelanceFi.CreateSheet(hesRow); //procesar la hoja y liberarla
                        crud.Update($"UPDATE hourReportFreelance SET status = 9 WHERE handleUnitId = '{hesRow["hesId"].ToString()}'", "freelance_db");
                    }
                    sap.KillSAP(); //mate sap y pase al siguiente
                    sap.BlockUser(mandante, 1);
                }
            }
            //if (hojas != null)
            //{
            //    if (!sap.CheckLogin(mandante))
            //    {
            //        sap.BlockUser(mandante, 1);

            //        if (hojas.Rows.Count > 0)
            //        {
            //            for (int i = 0; i < hojas.Rows.Count; i++)
            //            {
            //                HESAP eSAP = new HESAP
            //                {
            //                    Id = $"{hojas.Rows[i][0]}",
            //                    Item = $"{hojas.Rows[i][3]}",
            //                    PO = $"{hojas.Rows[i][2]}",
            //                    Horas = $"{hojas.Rows[i][4]}",
            //                    Cliente = $"{hojas.Rows[i][8]}",
            //                    Proveedor = $"{hojas.Rows[i][1]}",
            //                    Responsable = $"{hojas.Rows[i][9]}"
            //                };
            //                hes.Add(eSAP);
            //            }

            //        }

            //        if (Dproveedores != null)
            //        {
            //            if (Dproveedores.Rows.Count > 0)
            //            {
            //                for (int i = 0; i < Dproveedores.Rows.Count; i++)
            //                {
            //                    FreelanceVendors prove = new FreelanceVendors
            //                    {

            //                        Correo = $"{Dproveedores.Rows[i][1]}",
            //                        Nombre = $"{Dproveedores.Rows[i][2]}",
            //                        Consultores = ($"{Dproveedores.Rows[i][3]}").Split(',').ToList(),
            //                        Usuario = $"{Dproveedores.Rows[i][0]}"

            //                    };
            //                    proveedores.Add(prove);
            //                }
            //            }
            //        }

            //        hes.ForEach(x =>
            //        {
            //            using (SapVariants sap = new SapVariants())
            //            {
            //                sap.KillSAP(); // Si quedara abierto revise y matelo
            //                sap.LogSAP("300"); //Logearse
            //                freelanceFi.CreateSheet(x.PO, x.Item, x.Id, proveedores, x.Proveedor, x.Responsable); //procesar la hoja y liberarla
            //                sap.KillSAP(); //mate sap y pase al siguiente
            //            }
            //        });

            //        sap.BlockUser(mandante, 0);
            //    }
            //}

        }

        public void approvedCats()
        {
            MsExcel ex = new MsExcel();
            DataTable dt = ex.GetExcel(root.FilesDownloadPath + "\\" + "cats.xlsx");
            dt.Columns.Add("response");
            foreach (DataRow row in dt.Rows)
            {
                string cats = row["catsId"].ToString();
                cats = "00000" + cats;
                string aprobacion = freelanceFi.CatsAppr("A", cats, "", "ZGBM02");
                row["response"] = aprobacion;
                console.WriteLine(aprobacion);
            }
            dt.AcceptChanges();
            ex.CreateExcel(dt, "responses", root.FilesDownloadPath + "\\" + "responseCats.xlsx");

        }
    }


    #region Analytics Freelance Classes

    //public class PFreelance
    //{
    //    /// <summary>
    //    /// Pos por area
    //    /// </summary>
    //    public List<POArea> Areas { get; set; }
    //    public List<ReporteState> Estados_Area { get; set; }
    //    public double CHRP { get; set; }
    //    public double CHAPRDS { get; set; }
    //    public double CHAPROB { get; set; }
    //    public int CFREE { get; set; }
    //}
    ///// <summary>
    ///// Pos por area
    ///// </summary>
    //public class POArea
    //{
    //    public string Area { get; set; }
    //    public double Cantidad { get; set; }
    //}
    //public class ReporteState
    //{
    //    public string Area { get; set; }
    //    public int Aprobacion { get; set; }
    //    public int Aprobado { get; set; }
    //    public int Rechazado { get; set; }
    //    public int Devolucion { get; set; }
    //    public int Error { get; set; }
    //    public int Completado { get; set; }
    //}
    //public class HESAP
    //{
    //    public string PO { get; set; }
    //    public string Item { get; set; }
    //    public string Id { get; set; }
    //    public string Cliente { get; set; }
    //    public string Proveedor { get; set; }
    //    public string Responsable { get; set; }
    //    public string Horas { get; set; }
    //}
    //public class ProcessHES
    //{
    //    public string HES { get; set; }
    //    public string Po { get; set; }
    //    public string Item { get; set; }
    //    public byte[] Captura { get; set; }
    //}
    //public class CATSRecords
    //{
    //    public string Id { get; set; }
    //    public string Records { get; set; }
    //    public string HES { get; set; }
    //}
    //public class FreeHES
    //{
    //    public bool Estado { get; set; }
    //    public string HES { get; set; }
    //    public byte[] Captura { get; set; }
    //}
    //public class FreelanceVendors
    //{
    //    public string Usuario { get; set; }
    //    public string Nombre { get; set; }
    //    public string Correo { get; set; }
    //    public List<string> Consultores { get; set; }
    //}
    //public class DeterminateVendor
    //{
    //    public string CorreoProveedor { get; set; }
    //    public string CorreoResponsable { get; set; }
    //    public string NombreProveedor { get; set; }

    //    DeterminateVendor(List<FreelanceVendors> proveedores, string consultor, string responsable)
    //    {
    //        int indx = proveedores.FindIndex(x =>
    //        {
    //            return x.Consultores.Contains(consultor);
    //        });
    //        if (indx == -1)
    //        {
    //            //No existe una relacion del freelance con algun proveedor
    //        }

    //    }

    //}
    #endregion
}
