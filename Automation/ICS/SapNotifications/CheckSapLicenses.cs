using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using System.Globalization;
using DataBotV5.Data.SAP;
using System.Threading;
using System.Data;
using System;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;

namespace DataBotV5.Automation.ICS.SapNotifications
{
    internal class CheckSapLicenses
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();
        Log log = new Log();

        string respFinal = "";

        bool executeStats = false;



        const string mand = "QAS";

        public void Main()
        {
            CultureInfo OriginalCulture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            int[] mands = { 110, 120, 260, 300, 410, 420, 460, 500, 100, 400, 4002 }; //4002 es fiori 400 QAS

            bool warning = false;

            DataTable warningFewDaysResponse = new DataTable();
            warningFewDaysResponse.Columns.Add("SYSNAME");
            warningFewDaysResponse.Columns.Add("PRODUCTID");
            warningFewDaysResponse.Columns.Add("INSTNUMBER");
            warningFewDaysResponse.Columns.Add("SYSTEMID");
            warningFewDaysResponse.Columns.Add("EXP_DATE");
            warningFewDaysResponse.Columns.Add("LKEY");
            warningFewDaysResponse.Columns.Add("DÍAS FALTANTES");

            DataTable warningWeeklyResponse = warningFewDaysResponse.Clone();

            DataTable lastCheckDateDt = crud.Select("SELECT * FROM `checklLastLicense` WHERE `id` = 1", "check_sap_licenses_db");

            DateTime lastCheckDate = DateTime.Parse(lastCheckDateDt.Rows[0][1].ToString());

            foreach (int mand in mands)
            {
                CultureInfo culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
                culture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd";
                culture.DateTimeFormat.LongTimePattern = "";
                Thread.CurrentThread.CurrentCulture = culture;

                Dictionary<string, string> parameters = new Dictionary<string, string>();
                IRfcFunction zfmCheckLicense;
                string sysId = "";
                try
                {
                    console.WriteLine("Revisando mandante: " + mand);
                    zfmCheckLicense = sap.ExecuteRFC("", "/SDF/SLIC_READ_LICENSES_700", parameters, mand); //En caso de que especifique el mandante de la FM no hace falta poner ERP o CRM en el system
                    IRfcTable resultFM = zfmCheckLicense.GetTable("LICENSES");
                    sysId = zfmCheckLicense.GetValue("SYSTEMID").ToString().TrimStart('0');
                    DataTable resultLicense = sap.GetDataTableFromRFCTable(resultFM);

                    resultLicense.Columns.Add("DÍAS FALTANTES");
                    resultLicense.Columns.Remove("USERLIMIT");
                    resultLicense.Columns.Remove("FPRINT");
                    resultLicense.Columns.Remove("CRE_DATE");
                    resultLicense.Columns.Remove("LCHK_DATE");
                    resultLicense.Columns.Remove("CUSTKEY");

                    foreach (DataRow dr in resultLicense.Rows)
                    {
                        executeStats = true;
                        dr["SYSNAME"] = mand;
                        dr["SYSTEMID"] = sysId;
                        if (dr["EXP_DATE"].ToString().Trim() != "9999-12-31")  // No importan las licencias que no se vencen
                        {
                            dr["DÍAS FALTANTES"] = (Convert.ToDateTime(dr["EXP_DATE"]) - DateTime.Today).Days;

                            warningWeeklyResponse.Rows.Add(dr.ItemArray);

                            if (Convert.ToInt32(dr["DÍAS FALTANTES"]) <= 7)
                            {
                               
                                warning = true;
                                warningFewDaysResponse.Rows.Add(dr.ItemArray);

                                string respo = $"Quedan pocos días para el vencimiento de la siguiente licencia en SAP: {sysId} - {mand}";

                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Generar alerta licencia", respo, "");
                                respFinal = respFinal + "\\n" + respo;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DataRow dr = warningFewDaysResponse.NewRow();
                    dr["SYSNAME"] = mand;
                    dr["SYSTEMID"] = sysId;
                    dr["DÍAS FALTANTES"] = ex.Message;
                    warningFewDaysResponse.Rows.Add(dr);
                }
            }

            if (warning)
                mail.SendHTMLMail("Quedan pocos días para el vencimiento de las siguientes Licencias en SAP: " + "<br>" + val.ConvertDataTableToHTML(warningFewDaysResponse), new string[] { "internalcustomersrvs@gbm.net" }, "Información sobre expiración de licencias SAP: ",  new string[] { "lrrojas@gbm.net" });
            else if ((DateTime.Today - lastCheckDate).Days >= 7)
            {
                mail.SendHTMLMail("Resultado Verificación de licencias SAP: " + "<br>" + val.ConvertDataTableToHTML(warningWeeklyResponse), new string[] { "internalcustomersrvs@gbm.net" }, "Información sobre expiración de licencias SAP: ", new string[] { "lrrojas@gbm.net" });
                crud.Update("UPDATE `checklLastLicense` SET `date` = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' WHERE `checklLastLicense`.`id` = 1;", "check_sap_licenses_db");
            }

            if (executeStats == true)
            {
                root.BDUserCreatedBy = "internalcustomersrv";
                root.requestDetails = respFinal;

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

            Thread.CurrentThread.CurrentCulture = OriginalCulture;

        }
    }
}
