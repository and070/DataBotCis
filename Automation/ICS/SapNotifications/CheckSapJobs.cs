using DataBotV5.Logical.Processes;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using System.Globalization;
using DataBotV5.Data.SAP;
using System.Threading;
using System.Data;
using System;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;

namespace DataBotV5.Automation.ICS.SapNotifications
{
    internal class CheckSapJobs
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        Log log = new Log();

        const string erpMand = "ERP";

        string respFinal = "";


        public void Main()
        {

            DataTable jobsTable = new DataTable();
            jobsTable.Columns.Add("Job Name");
            jobsTable.Columns.Add("Program");
            jobsTable.Columns.Add("Variant");
            jobsTable.Columns.Add("Start Date");
            jobsTable.Columns.Add("Status");

            CultureInfo cultureOriginal = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            CultureInfo culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            culture.DateTimeFormat.ShortDatePattern = "yyyyMMdd";
            culture.DateTimeFormat.LongTimePattern = "";
            Thread.CurrentThread.CurrentCulture = culture;
            string date = DateTime.Today.AddDays(-1).ToString();
            Thread.CurrentThread.CurrentCulture = cultureOriginal;

  

            #region Traer datos de ERP - Tabla TBTCP
            RfcDestination destErp = sap.GetDestRFC(erpMand);

            IRfcFunction fmTableErp = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            fmTableErp.SetValue("USE_ET_DATA_4_RETURN", "X");
            fmTableErp.SetValue("QUERY_TABLE", "TBTCP");
            fmTableErp.SetValue("DELIMITER", "|");


            IRfcTable fieldsErp = fmTableErp.GetTable("FIELDS");
            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "JOBNAME");
            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "PROGNAME");
            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "VARIANT");
            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "SDLDATE");
            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "STATUS");


            IRfcTable optionTableErp = fmTableErp.GetTable("OPTIONS");
            optionTableErp.Append();
            optionTableErp.SetValue("TEXT", "SDLDATE IN ('" + date + "' ) AND STATUS EQ 'A'");


            fmTableErp.Invoke(destErp);

            DataTable reportErp = sap.GetDataTableFromRFCTable(fmTableErp.GetTable("ET_DATA"));
            foreach (DataRow rowErp in reportErp.Rows)
            {
                DataRow rowReport = jobsTable.NewRow();
                rowReport["Job Name"] = rowErp["LINE"].ToString().Split(new char[] { '|' })[0].Trim().TrimStart(new char[] { '0' });
                rowReport["Program"] = rowErp["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                rowReport["Start Date"] = rowErp["LINE"].ToString().Split(new char[] { '|' })[2].Trim();
                rowReport["Variant"] = rowErp["LINE"].ToString().Split(new char[] { '|' })[3].Trim();
                rowReport["Status"] = rowErp["LINE"].ToString().Split(new char[] { '|' })[4].Trim().Replace("A", "Canceled");

                jobsTable.Rows.Add(rowReport);

                string response =
                    "Se verifica el Job en SAP: \\n" +
                    "Job Name: " + rowErp["LINE"].ToString().Split(new char[] { '|' })[0].Trim().TrimStart(new char[] { '0' }) +", "+
                    "Program: " + rowErp["LINE"].ToString().Split(new char[] { '|' })[1].Trim() + ", " +
                    "Start Date: " + rowErp["LINE"].ToString().Split(new char[] { '|' })[2].Trim() + ", " +
                    "Variant: " + rowErp["LINE"].ToString().Split(new char[] { '|' })[3].Trim() + ", " +
                    "Status: " + rowErp["LINE"].ToString().Split(new char[] { '|' })[4].Trim().Replace("A", "Canceled");


                log.LogDeCambios("Revisión", root.BDProcess,root.BDUserCreatedBy, "Verificar Job SAP", response, "");
                respFinal = respFinal + "\\n" + "Verificar Job SAP: " + response;

            }

            #endregion


            if (jobsTable.Rows.Count != 0)
            {

                mail.SendHTMLMail("Resultado Verificación de Jobs SAP: " + "<br>" + val.ConvertDataTableToHTML(jobsTable), new string[] { "internalcustomersrvs@gbm.net" }, "Información sobre Jobs No Ejecutados");

                root.BDUserCreatedBy = "internalcustomersrvs";
                root.requestDetails = respFinal;

                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
            else
            {
                mail.SendHTMLMail("No existen Jobs Cancelados del día de ayer.", new string[] { "internalcustomersrvs@gbm.net" }, "Información sobre Jobs No Ejecutados");
            }


      
        }
    }
}
