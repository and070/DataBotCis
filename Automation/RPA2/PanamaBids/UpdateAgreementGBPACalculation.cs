using System;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Projects.PanamaBids;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.Data.Database;
using DataBotV5.App.Global;
using System.Data;
using System.Collections.Generic;
using DataBotV5.Data.Projects.PanamaBids;
using System.Linq;

namespace DataBotV5.Automation.RPA2.PanamaBids
{
    /// <summary>
    /// Clase RPA para calcular y actualizar multa por quote (mediante envío de correo electrónico).
    /// </summary>
    class UpdateAgreementGBPACalculation
    {
        #region variables_globales
        PanamaPurchase pa_compra = new PanamaPurchase();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        MailInteraction mail = new MailInteraction();
        ProcessAdmin padmin = new ProcessAdmin();
        MsExcel MsExcel = new MsExcel();
        ValidateData val = new ValidateData();
        WebInteraction sel = new WebInteraction();
        WebexTeams wt = new WebexTeams();
        Database wb2 = new Database();
        Log log = new Log();
        ConsoleFormat console = new ConsoleFormat();
        ProcessInteraction proc = new ProcessInteraction();
        CRUD crud = new CRUD();
        string mandante = "QAS";

        string respFinal = "";

        #endregion

        public void Main() //convenio_update()
        {

            if (mail.GetAttachmentEmail("Solicitudes convenio gbpa update", "Procesados", "Procesados convenio gbpa update"))
            {
                console.WriteLine(DateTime.Now + " > > > " + "Procesando...");
                string respuesta = UpdateAgreement(root.FilesDownloadPath + "\\" + root.ExcelFile);

                console.WriteLine(DateTime.Now + " > > > " + "Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        /// <summary>
        /// Calcular y actualizar multa, sales order, forecast y fecha de entrega mediante quote
        /// </summary>
        /// <param name="route">El archivo que se descarga desde el correo electronico (ver carpeta Solicitudes convenio gbpa update)</param>
        /// <returns></returns>
        private string UpdateAgreement(string route)
        {
            #region Variables Privadas
            int rows;
            string mensaje_devolucion = "";
            string rerror = "";
            string estadogbm = "";
            string subtotal = "";
            string fecha_maxima = "";
            string entidad = "";
            bool validar_lineas = true;
            bool sqlerror = true;
            string respuesta = ""; string validacion = ""; string respuesta_final = ""; string multa = "";
            DataTable bidsInfo = crud.Select( "SELECT singleOrderRecord, entity, quote, salesOrder, forecast, maximumDeliveryDate, orderSubtotal, gbmStatus FROM `purchaseOrderMacro`", "panama_bids_db");
            #endregion

            console.WriteLine("Abrir Excel y modificando");
            #region abrir excel
            DataTable xlWorkSheet = MsExcel.GetExcel(route);

            xlWorkSheet.Columns.Add("Entidad");
            xlWorkSheet.Columns.Add("Fecha Máxima de Entrega");
            xlWorkSheet.Columns.Add("Subtotal");
            xlWorkSheet.Columns.Add("Multa");
            xlWorkSheet.Columns.Add("Respuesta");
            #endregion

            validacion = xlWorkSheet.Rows[1][1].ToString();
            if (validacion != "")
            { validacion = validacion.ToString().Trim(); }
            if (validacion != "Billing Document")
            {
                mensaje_devolucion = "No es plantilla de Actualización de Registros, por favor volver a enviar la correcta";
                mail.SendHTMLMail(mensaje_devolucion, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
                return "";
            }

            foreach (DataRow rRow in xlWorkSheet.Rows)
            {

                try
                {

                    respuesta = "";
                    rerror = "";
                    multa = "";
                    subtotal = "";
                    fecha_maxima = "";
                    entidad = "";

                    string quote = rRow[""].ToString().Trim();
                    if (quote == "")
                    {
                        continue;
                    }
                    if (quote.Substring(0, 3) != "733")
                    {
                        rerror = quote + " : Quote invalido.";
                        mensaje_devolucion += rerror + "<br>";
                        validar_lineas = false;
                        rRow["Respuesta"] = rerror;
                        continue;
                    }
                    //extraer la fecha maxima de entrega y el subtotal del quote (si no tiene significa que el quote no existe en la BD) notificar a Andreina  validar_lineas = false; no continue
                    DataRow[] bidInfo = bidsInfo.Select($"quote ='{quote}'");

                    // en caso que no tenga no hace nada

                    if (bidInfo.Count() > 0)
                    {
                        string reunico = bidInfo[0]["singleOrderRecord"].ToString();
                        entidad = bidInfo[0]["entity"].ToString();

                        string sales_order = rRow["Sales document"].ToString().Trim();
                        if (String.IsNullOrEmpty(sales_order))
                        { sales_order = bidInfo[0]["salesOrder"].ToString(); }

                        string forecast = "";
                        DateTime forecastdate = DateTime.MinValue;
                        if (DateTime.TryParse(rRow["Billing Date"].ToString().Trim(), out forecastdate))
                        {
                            forecast = forecastdate.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            forecast = bidInfo[0]["forecast"].ToString();
                        }


                        DateTime edate = DateTime.MinValue;
                        DateTime fentrega = DateTime.MinValue;
                        string fecha_entrega = "";
                        if (DateTime.TryParse(rRow["Fecha de Entrega"].ToString().Trim(), out fentrega))
                        {
                            fecha_entrega = fentrega.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            fecha_entrega = rRow["Fecha de Entrega"].ToString().Trim();
                        }

                        #region calculo de multa
                        if (DateTime.TryParse(fecha_entrega, out edate)) //si es fecha calcula la multa
                        {
                            estadogbm = "1"; // "Facturado";
                            fecha_maxima = bidInfo[0]["maximumDeliveryDate"].ToString();
                            subtotal = bidInfo[0]["orderSubtotal"].ToString();
                            DateTime fmdate = DateTime.MinValue;
                            if (DateTime.TryParse(fecha_maxima, out fmdate)) //si es fecha calcula la multa
                            {
                                if (edate >= fmdate)
                                {
                                    //si hay multa
                                    fecha_entrega = edate.ToString("yyyy-MM-dd");
                                    TimeSpan differ = edate - fmdate;
                                    int dias_multa = differ.Days;
                                    subtotal = subtotal.Replace(".", ",");
                                    float subtotal_itbms = float.Parse(subtotal) * (float)1.07;
                                    //float itbms = (float)1.07;
                                    //subtotal_itbms *= itbms;
                                    //float porcen = (float)0.04;
                                    float multa_diaria = ((float)0.04 * subtotal_itbms) / 30;
                                    float mc = (float)Math.Round(multa_diaria * 100f) / 100f;
                                    float multa_total = dias_multa * mc;
                                    multa = multa_total.ToString();
                                    multa = multa.Replace(",", ".");
                                }
                                else
                                {
                                    multa = "0";
                                }
                            }
                            else
                            {
                                mensaje_devolucion = mensaje_devolucion + quote + " : Error al convertir fecha maxima de entrega, por favor verificar y/o actualizar.<br>";
                                validar_lineas = false;
                                sqlerror = false;
                                continue;
                            }
                        }
                        else
                        {
                            multa = "0";
                            fecha_entrega = "";
                            estadogbm = rRow["gbmStatus"].ToString().Trim();
                        }
                        #endregion
                        console.WriteLine("  Actualizando información a la base de datos");

                        //bool update = lpsql.UpdateRegister(quote, infoupd, 2);
                        string sql = $@"UPDATE `purchaseOrderMacro` SET actualDeliveryDate = '{fecha_entrega}', salesOrder = '{sales_order}', forecast = '{forecast}', fineAmount = '{multa}', gbmStatus = {estadogbm} WHERE quote = '{quote}'";
                        bool update = crud.Update(sql, "panama_bids_db");
                        respuesta = reunico + " - " + entidad + ". Multa: $" + multa + "<br>";
                        console.WriteLine("  " + respuesta);
                        console.WriteLine("  Resultado de actualización: " + update);
                        
                        log.LogDeCambios("Modificacion", root.BDProcess, "Ventas Panama", "Actualizar Registro Unico por quote", respuesta, update.ToString());
                        respFinal = respFinal + "\\n" + "Actualizar registro único por quote: " + respuesta;




                        if (update == false)
                        {
                            //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                            //para enviarlo por email y agregarla
                            rerror = respuesta + "error al actualizar.<br>";
                            mensaje_devolucion += rerror;
                            sqlerror = false;
                        }
                        else
                        {
                            respuesta_final += respuesta;
                        }

                    }
                    else
                    {
                        rerror = quote + " : Quote no existe en la base de datos.";
                        mensaje_devolucion += rerror + "<br>";
                        validar_lineas = false;
                        rRow["Respuesta"] = rerror;
                        continue;
                    }

                    #region guardar info en excel
                    rRow["Entidad"] = entidad;
                    rRow["Fecha Máxima de Entrega"] = fecha_maxima;
                    rRow["Subtotal"] = subtotal;
                    rRow["Multa"] = multa;
                    respuesta = respuesta.Replace("<br>", "");
                    rerror = rerror.Replace("<br>", "");
                    rRow["Respuesta"] = respuesta + rerror;
                    #endregion

                }
                catch (Exception ex)
                {
                    sqlerror = false;
                    rerror = ex.Message;
                    mensaje_devolucion += rerror;
                    rRow["Respuesta"] = rerror;
                }
            } //for
            xlWorkSheet.AcceptChanges();
            MsExcel.CreateExcel(xlWorkSheet, "Resultados", route);

            string[] adjunto = { route };

            if (validar_lineas == false)
            {
                if (sqlerror == false)
                {
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail("Se le notifica que los siquientes quotes no se actualizarón: " + "<br>" + mensaje_devolucion, new string[] {"appmanagement@gbm.net"}, root.Subject, cc);
                }
                string[] cc2 = { root.BDUserCreatedBy };
                mail.SendHTMLMail("Se le notifica que se actualizarón los siguientes quotes: <br><br>" + respuesta_final + "<br><br>" + "Con excepción de: <br><br>" + mensaje_devolucion, new string[] { "kvanegas@gbm.net" }, root.Subject, cc2, attachments: adjunto);
            }
            else
            {
                //todo salio bien, se envia notificación.
                string[] cc2 = { root.BDUserCreatedBy };
                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Calculo de multa - ordenes Convenio Macro");
                html = html.Replace("{cuerpo}", "Se le notifica que se actualizarón los siguientes quotes");
                html = html.Replace("{contenido}", respuesta_final);
                mail.SendHTMLMail(html, new string[] { "kvanegas@gbm.net" }, root.Subject, cc2, adjunto);


            }

            root.requestDetails = respFinal;

            return "";
        }


    }
}

