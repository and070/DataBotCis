using DataBotV5.Logical.MicrosoftTools;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.HumanCapital
{
    /// <summary>
    /// Clase ICS Automation encargada de la sincronización de colaboradores en human capital.
    /// </summary>
    class SyncPerson
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly SapVariants sap = new SapVariants();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        const string mand = "ERP";
        string respFinal = "";

        public void Main()
        {
            //leer correo
            mail.GetAttachmentEmail("Solicitudes Sincronizar colaborador", "Procesados", "Procesados Sincronizar colaborador");
            if (!string.IsNullOrWhiteSpace(root.Email_Body) && (root.BDUserCreatedBy.ToLower() == "bpm@mailgbm.com" || root.BDUserCreatedBy.ToLower().Contains("databot")))
            {
                ProcessSyncPerson(false);//POR BPM

                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
            else if (root.ExcelFile != null && root.ExcelFile != "")
            {
                ProcessSyncPerson(true);//POR CORREO
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="reqType">false: BPM , true: EMAIL</param>
        private void ProcessSyncPerson(bool reqType)
        {
            #region Variables Privadas
            bool validateLines = true;
            Regex htmlFix = new Regex("[*'\"_&+^><@]");
            Regex alphanum = new Regex(@"[^\p{L}0-9 ]");
            string res1 = "", response, employeeID, employeeName = "", responseFailure = "", body = root.Email_Body;
            string[] error;
            #endregion

            if (reqType == false)//bpm
            {
                if (body.Contains("Número Colaborador"))
                {
                    //VALIDACIONES                
                    body = htmlFix.Replace(body, string.Empty);
                    string[] Separator = new string[] { "Número Colaborador" };
                    string[] bodySplit = body.Split(Separator, StringSplitOptions.None);
                    bodySplit[1] = bodySplit[1].Replace('\r', ' ');
                    bodySplit = bodySplit[1].Split('\n');
                    employeeID = alphanum.Replace(bodySplit[0], "").Trim().ToUpper();

                    Separator = new string[] { "Nombre Colaborador" };
                    bodySplit = body.Split(Separator, StringSplitOptions.None);
                    bodySplit[1] = bodySplit[1].Replace('\r', ' ');
                    bodySplit = bodySplit[1].Split('\n');
                    employeeName = alphanum.Replace(bodySplit[0], "").Trim();

                    if (employeeID != "")
                    {
                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["ID"] = employeeID,
                                ["IDTYPE"] = "2"
                            };

                            IRfcFunction func = sap.ExecuteRFC(mand, "ZHR_SYNC_PERSON_FM", parameters);

                            #region Procesar Salidas del FM

                            //arreglar respuesta
                            if (func.GetValue("RESPONSE").ToString().Contains("|") == true)
                            {
                                validateLines = false;
                                error = func.GetValue("RESPONSE").ToString().Split('|');
                                response = error[4].Trim();
                            }
                            else
                                response = func.GetValue("RESPONSE").ToString();

                            //log de base de datos
                            console.WriteLine("Sincronizar colaborador: " + employeeID + ": " + response);
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Sincronizar colaborador", employeeID + ": " + response, root.Subject);
                            res1 = res1 + "Nombre colaborador: " + employeeName + "<br><br>" + "ID: " + employeeID + ": " + response + "<br>";
                            respFinal = respFinal + "Sincronizar colaborador: " + employeeID + ": " + response;

                            #endregion
                        }
                        catch (Exception ex)
                        {
                            responseFailure = ex.ToString();
                            console.WriteLine(" Finalizando Proceso " + responseFailure);
                            res1 = res1 + employeeID + ": " + responseFailure + "<br>";
                            validateLines = false;
                        }
                        #endregion
                    }

                    console.WriteLine("Respondiendo solicitud");
                    if (validateLines == false)//enviar email de repuesta de error
                    {
                        if (responseFailure != "")
                            mail.SendHTMLMail(res1 + "<br>" + responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                        else
                            mail.SendHTMLMail(res1 + "<br>" + responseFailure, new string[] { "gvillalobos@gbm.net" }, root.Subject,  new string[] { "internalcustomersrvs@gbm.net" });
                    }
                    else//enviar email de repuesta de éxito
                        mail.SendHTMLMail(res1, new string[] { "gvillalobos@gbm.net" }, root.Subject, root.CopyCC);
                }
            }
            else//correo manual
            {
                #region abrir excel
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile, false);
                #endregion

                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["IDTYPE"] = "2"
                };

                foreach (DataRow item in excelDt.Rows)
                {
                    employeeID = item[0].ToString().Trim();
                    if (employeeID != "")
                    {
                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            parameters["ID"] = employeeID;
                            IRfcFunction func = sap.ExecuteRFC(mand, "ZHR_SYNC_PERSON_FM", parameters);
                            #region Procesar Salidas del FM
                            //arreglar respuesta
                            if (func.GetValue("RESPONSE").ToString().Contains("|") == true)
                            {
                                validateLines = false;
                                error = func.GetValue("RESPONSE").ToString().Split('|');
                                response = error[4].Trim();
                            }
                            else
                                response = func.GetValue("RESPONSE").ToString();

                            //log de base de datos
                            console.WriteLine("Sincronizar colaborador: " + employeeID + ": " + response);
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Sincronizar colaborador", employeeID + ": " + response, root.Subject);
                            res1 = res1 + "Nombre colaborador: " + employeeName + "<br><br>" + "ID: " + employeeID + ": " + response + "<br>";
                            respFinal = respFinal + "Sincronizar colaborador: " + employeeID + ": " + response;

                            #endregion
                        }
                        catch (Exception ex)
                        {
                            responseFailure = ex.ToString();
                            console.WriteLine(" Finalizando Proceso " + responseFailure);
                            res1 = res1 + employeeID + ": " + responseFailure + "<br>";
                            validateLines = false;
                        }
                        #endregion
                    }
                }
                if (validateLines == false)
                    //enviar email de repuesta de error
                    mail.SendHTMLMail(res1 + "<br>" + responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                else
                    //enviar email de repuesta de exito
                    mail.SendHTMLMail(res1.Replace("Nombre colaborador:", ""), new string[] { root.BDUserCreatedBy }, root.Subject);

                root.requestDetails = respFinal;

            }
        }
    }
}
