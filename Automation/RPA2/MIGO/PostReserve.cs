using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;

namespace DataBotV5.Automation.RPA2.MIGO
{
    /// <summary>
    /// Clase RPA Automation encargada de "postear" reservas en MIGO.   
    /// </summary>
    class PostReserve 
    {
        public string response = "";
        public string response_failure = "";
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        int a;
        string mandante = "ERP";
        string material = "", serie = "", respuesta = "", validacion = "", reserva = "", asset_serie = "", asset_material = "", asset_placa = "";
        string respFinal = "";

        public void Main()
        {
            //leer correo y descargar archivo
            if (mail.GetAttachmentEmail("Solicitudes Reservas", "Procesados", "Procesados Reservas"))
            {
                ProcessPostReserve(root.FilesDownloadPath + "\\" + root.ExcelFile);
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void ProcessPostReserve(string route)
        {
            #region Variables Privadas
            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;
            respuesta = "";
            response = "";
            #endregion

            #region abrir excel
            Excel.Range xlRango;
            Excel.Range xlRangoDuplicate;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            rows = xlWorkSheet.UsedRange.Rows.Count;
            #endregion

            validacion = xlWorkSheet.Cells[1, 1].text.ToString().Trim();
            //Planntilla correcta, continue las validaciones
            for (int i = 2; i <= rows; i++)
            {

                reserva = xlWorkSheet.Cells[i, 2].text.ToString().Trim();
                if (reserva == "")
                {
                    continue;
                }
                else
                {
                    asset_material = xlWorkSheet.Cells[i, 3].text.ToString().Trim();
                    asset_serie = xlWorkSheet.Cells[i, 4].text.ToString().Trim();
                    asset_placa = xlWorkSheet.Cells[i, 5].text.ToString().Trim();
                    asset_serie = asset_serie.ToUpper();
                    asset_material = asset_material.ToUpper();
                    asset_placa = asset_placa.ToUpper();

                    material = xlWorkSheet.Cells[i, 3].text.ToString().Trim();


                    #region SAP
                    Console.WriteLine(DateTime.Now + " > > > " + "Corriendo RFC de SAP: " + root.BDProcess);
                    try
                    {
                        RfcDestination destination = new SapVariants().GetDestRFC(mandante);

                        RfcRepository repo = destination.Repository;
                        IRfcFunction func = repo.CreateFunction("ZFI_POST_RESERVA");
                        IRfcTable it_mat = func.GetTable("IT_MATERIALES");

                        #region Parametros de SAP

                        func.SetValue("RESERVA_DOC_NUM", reserva);

                        a = 4;
                        while (material != "")
                        {
                            it_mat.Append();

                            it_mat.SetValue("MATERIAL", material);
                            if (a == 4)
                            { it_mat.SetValue("SERIE", asset_serie); }
                            else
                            { it_mat.SetValue("SERIE", serie); }

                            a = a + 2;
                            material = xlWorkSheet.Cells[i, a].text.ToString().Trim();
                            serie = xlWorkSheet.Cells[i, a + 1].text.ToString().Trim();
                        }

                        func.SetValue("ASSET_MATERIAL", asset_material);
                        func.SetValue("ASSET_PLACA", asset_placa);
                        func.SetValue("ASSET_SERIE", asset_serie);
                        #endregion
                        #region Invocar FM
                        func.Invoke(destination);
                        #endregion
                        #region Procesar Salidas del FM
                        respuesta = func.GetValue("RESPUESTA").ToString();
                        string id_migo = func.GetValue("ID_DOC").ToString();
                        //log de base de datos
                        console.WriteLine(material + ": " + func.GetValue("RESPUESTA").ToString());
                        if (!String.IsNullOrEmpty(id_migo))
                        {
                            xlWorkSheet.Cells[i, 1].value = id_migo;
                        }

                        if (respuesta.Contains("error"))
                        {
                            validar_lineas = false;
                            xlWorkSheet.Cells[i, 1].value = func.GetValue("RESPUESTA").ToString();
                        }
                        response = (!String.IsNullOrEmpty(id_migo)) ? response + reserva + " - " + id_migo + ": " + respuesta + "<br>" : response + reserva + ": " + respuesta + "<br>";
                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Post Reserva MIGO", reserva + ": " + func.GetValue("RESPUESTA").ToString(), func.GetValue("ID_DOC").ToString());
                        respFinal = respFinal + "\\n" + reserva + ": " + func.GetValue("RESPUESTA").ToString();

                        #endregion
                    }
                    catch (Exception ex)
                    {
                        response_failure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                        console.WriteLine(" Finishing process " + response_failure);
                        response = respuesta + material + ": " + ex.ToString() + "<br>";
                        response_failure = ex.ToString();
                        xlWorkSheet.Cells[i, 1].value = response;
                        validar_lineas = false;

                    }

                    #endregion

                }
            } //for de cada linea del excel
            console.WriteLine("Respondiendo solicitud");

            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(route);
            xlWorkBook.Close();
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL",true);

            string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };

            if (validar_lineas == false)
            {
                //enviar email de repuesta de error
                string[] cc = { "appmanagement@gbm.net" };
                mail.SendHTMLMail(response + "<br>" + "<br>" + response_failure, new string[] { root.BDUserCreatedBy }, root.Subject, cc, adjunto);
            }
            else
            {
                //enviar email de repuesta de exito
                mail.SendHTMLMail(response, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC, adjunto);
            }


            root.requestDetails = respFinal;





        }

    }
}
