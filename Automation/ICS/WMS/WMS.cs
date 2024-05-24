using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System;
using System.Data;

namespace DataBotV5.Automation.ICS.WMS

{
    /// <summary>
    /// Clase ICS Automation encargada de la gestión WMS en ICS.
    /// </summary>
    class WMS
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        MsExcel excel = new MsExcel();
        Stats stats = new Stats();
        Log log = new Log();
        string respFinal = "";
        string erpSystem = "ERP";

        public void Main()
        {
            //leer correo y descargar archivo
            if (mail.GetAttachmentEmail("Solicitudes WMS", "Procesados", "Procesados WMS"))
            {
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessWMS(excelDt);
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessWMS(DataTable excelDt)
        {
            bool validateLines = true;
            string responseFailure = "", response = "";

            string validacion = excelDt.Columns[0].ColumnName;

            if (validacion.Substring(0, 1) != "x")
            {
                response = "Utilizar la plantilla oficial de datos maestros";
                mail.SendHTMLMail(response, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }
            else
            {
                //Plantilla correcta, continúe las validaciones
                foreach (DataRow item in excelDt.Rows)
                {
                    string material = item[0].ToString().Trim();
                    if (material != "")
                    {
                        material = material.ToUpper();
                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            #region Parámetros de SAP
                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["MATNR01"] = material
                            };

                            IRfcFunction func = sap.ExecuteRFC(erpSystem, "ZDM_AMP_WMS", parameters);
                            #endregion

                            #region Procesar Salidas del FM
                            response = response + material + ": " + func.GetValue("RESPUESTA").ToString() + "<br>";

                            //log de base de datos
                            console.WriteLine(material + ": " + func.GetValue("RESPUESTA").ToString());
                            log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Ampliar Material a WMS", material + ": " + response, root.Subject);
                            respFinal = respFinal + "\\n" + "Ampliar Material a WMS " + material + ": " + response;

                            if (response.Contains("Error"))
                                validateLines = false;
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            responseFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                            console.WriteLine(" Finishing process " + responseFailure);
                            response = response + material + ": " + ex.ToString() + "<br>";
                            responseFailure = ex.ToString();
                            validateLines = false;
                        }

                        #endregion
                    }
                }
                console.WriteLine("Respondiendo solicitud");
                if (validateLines == false)
                    mail.SendHTMLMail(response + "<br>" + responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);  //enviar email de repuesta de error
                else
                    mail.SendHTMLMail(response, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC); //enviar email de repuesta de éxito

                root.requestDetails = respFinal;

            }
        }

    }
}
