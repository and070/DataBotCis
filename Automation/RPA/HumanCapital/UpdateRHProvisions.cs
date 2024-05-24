using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.RPA.HumanCapital
{
    /// <summary>
    /// Clase RPA Automation encargada de la actualización de provisiones de Human Capital.
    /// </summary>
    class UpdateRHProvitions
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        string respFinal = "";
        Log log = new Log();


        public void Main()
        {
            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Provisiones HR", "Procesados", "Procesados Provisiones HR"))
            {
                console.WriteLine("Procesando...");
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessProvisionsRH(excelDt);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void ProcessProvisionsRH(DataTable excelDt)
        {
            DataTable proc110 = new DataTable();
            DataTable proc260 = new DataTable();
            DataTable proc300 = new DataTable();

            string validation = excelDt.Columns[2].ColumnName;

            if (validation.Substring(0, 1) == "x")
            {
                proc110 = UpdateProv(110, excelDt);
                proc260 = UpdateProv(260, excelDt);
                proc300 = UpdateProv(300, excelDt);
            }
            else
                mail.SendHTMLMail("Utilizar la plantilla oficial de datos maestros", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

            string res110 = "Resultados DEV:<br>" + val.ConvertDataTableToHTML(proc110) + "<br><br>";
            string res260 = "Resultados QAS:<br>" + val.ConvertDataTableToHTML(proc260) + "<br><br>";
            string res300 = "Resultados PRD:<br>" + val.ConvertDataTableToHTML(proc300) + "<br><br>";

            if (proc110.Select("`Resultado` LIKE '%FAILURE%'").Length > 0 || proc260.Select("`Resultado` LIKE '%FAILURE%'").Length > 0 || proc300.Select("`Resultado` LIKE '%FAILURE%'").Length > 0)
                mail.SendHTMLMail(res110 + res260 + res300, new string[] { root.BDUserCreatedBy }, root.Subject, new string[] { "internalcustomersrvs@gbm.net" });
            else//éxito
                mail.SendHTMLMail(res110 + res260 + res300, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

            root.requestDetails = respFinal;

        }

        private DataTable UpdateProv(int mandante, DataTable excelDt)
        {
            string response = "";

            DataTable newDt = excelDt.Copy();

            try { newDt.Columns.Add("Resultado"); } catch (DuplicateNameException) { }

            foreach (DataRow item in newDt.Rows)
            {
                string cons = item[0].ToString().Trim();
                if (cons != "")
                {
                    string ratio = item[1].ToString().Trim();
                    string sDate = item[2].ToString().Trim();

                    //Validaciones

                    if (cons.Length > 5)
                        cons = cons.Substring(0, 5);

                    if (ratio.Contains(","))
                        ratio = ratio.Replace(",", ".");

                    if (sDate.Contains("."))
                    {
                        sDate = sDate.ToString().Trim();

                        string[] dmy3 = sDate.Split(new char[1] { '.' });
                        string day = int.Parse(dmy3[0]).ToString();
                        if (day.Length == 1)
                            day = "0" + day;
                        string month = int.Parse(dmy3[1]).ToString();
                        if (month.Length == 1)
                            month = "0" + month;
                        sDate = int.Parse(dmy3[2]) + "-" + month + "-" + day;

                        #region SAP
                        try
                        {
                            console.WriteLine("Corriendo RFC de SAP: " + mandante + " - " + root.BDProcess);
                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["CONSTANTE"] = cons,
                                ["PERCENT"] = ratio,
                                ["FECHAINICIAL"] = sDate
                            };

                            IRfcFunction func = new SapVariants().ExecuteRFC("", "ZRPA_HR_PROVISION", parameters, mandante);

                            string respfm = func.GetValue("RESPUESTA").ToString();
                            console.WriteLine("Corriendo RFC de SAP: " + respfm);
                            if (respfm == "OK")
                                response = "La provisión ha sido actualizada";
                            else if (respfm == "")
                                response = "FAILURE: La información es incorrecta";
                            else if (respfm.Contains("ERROR"))
                                response = "FAILURE: La provisión dio error a la hora de actualizarse";

                            //log de cambios base de datos
                            log.LogDeCambios("Modificar", root.BDProcess, root.BDUserCreatedBy , "Actualizar Retencion", mandante + " - " + cons + ": " + respfm, root.Subject);
                            respFinal = respFinal + "\\n" + mandante + " - " + cons + ": " + respfm;

                        }
                        catch (Exception ex)
                        {
                            string responseFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                            console.WriteLine(" Finishing process " + responseFailure);
                            response = "FAILURE: " + ex.Message;
                        }
                        #endregion
                    }

                    item["Resultado"] = response;
                }
            }
            return newDt;
        }
    }
}
