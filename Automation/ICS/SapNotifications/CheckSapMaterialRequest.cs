using DataBotV5.Logical.Processes;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using System.Globalization;
using DataBotV5.Data.SAP;
using System.Threading;
using System.Linq;
using System.Data;
using System;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.ICS.SapNotifications
{
    internal class CheckSapMaterialRequest
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        ConsoleFormat console = new ConsoleFormat();


        Log log = new Log();
        Rooting root = new Rooting();


        const string crmMand = "CRM";
        const string erpMand = "ERP";

        string respFinal = "";


        public void Main()
        {

            DataTable dataErp = new DataTable();
            dataErp.Columns.Add("MR");
            dataErp.Columns.Add("FECHA");

            DataTable dataCrm = dataErp.Clone();

            CultureInfo cultureOriginal = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            CultureInfo culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            culture.DateTimeFormat.ShortDatePattern = "yyyyMMdd";
            culture.DateTimeFormat.LongTimePattern = "";
            Thread.CurrentThread.CurrentCulture = culture;
            string today = DateTime.Today.ToString();
            Thread.CurrentThread.CurrentCulture = cultureOriginal;




            #region Traer datos de ERP - VBAK
            RfcDestination destErp = sap.GetDestRFC(erpMand);

            IRfcFunction readTableVbak = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            readTableVbak.SetValue("USE_ET_DATA_4_RETURN", "X");
            readTableVbak.SetValue("QUERY_TABLE", "VBAK");
            readTableVbak.SetValue("DELIMITER", "|");

            IRfcTable fieldsErp = readTableVbak.GetTable("FIELDS");

            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "VBELN");
            fieldsErp.Append();
            fieldsErp.SetValue("FIELDNAME", "AUDAT");


            IRfcTable optionsVbak = readTableVbak.GetTable("OPTIONS");


            optionsVbak.Append();
            optionsVbak.SetValue("TEXT", "AUDAT IN ('" + today + "' )");

            readTableVbak.Invoke(destErp);

            DataTable reportErp = sap.GetDataTableFromRFCTable(readTableVbak.GetTable("ET_DATA"));
            foreach (DataRow row in reportErp.Rows)
            {
                DataRow rowReport = dataErp.NewRow();
                rowReport["MR"] = row["LINE"].ToString().Split('|')[0].Trim().TrimStart('0');
                rowReport["FECHA"] = row["LINE"].ToString().Split('|')[1].Trim();

                dataErp.Rows.Add(rowReport);
            }

            for (int i = dataErp.Rows.Count - 1; i >= 0; i--)
            {
                if (!dataErp.Rows[i]["MR"].ToString().StartsWith("52")) //El numero para el MR se maneja un consecutivo de momento está en 52
                {
                    dataErp.Rows[i].Delete();
                }
            }

            #endregion


            #region Traer datos de CRM - CRMD_ORDERADM_H
            RfcDestination destCrm = sap.GetDestRFC(crmMand);

            IRfcFunction readTableCrm = destCrm.Repository.CreateFunction("RFC_READ_TABLE");
            readTableCrm.SetValue("USE_ET_DATA_4_RETURN", "X");
            readTableCrm.SetValue("QUERY_TABLE", "CRMD_ORDERADM_H");
            readTableCrm.SetValue("DELIMITER", "|");
            readTableCrm.SetValue("USE_ET_DATA_4_RETURN", "X");

            IRfcTable fieldsCrm = readTableCrm.GetTable("FIELDS");

            fieldsCrm.Append();
            fieldsCrm.SetValue("FIELDNAME", "OBJECT_ID");
            fieldsCrm.Append();
            fieldsCrm.SetValue("FIELDNAME", "POSTING_DATE");

            IRfcTable optionsTableCrm = readTableCrm.GetTable("OPTIONS");


            optionsTableCrm.Append();
            optionsTableCrm.SetValue("TEXT", "POSTING_DATE IN ('" + today + "' )");

            readTableCrm.Invoke(destCrm);

            DataTable reportCrm = sap.GetDataTableFromRFCTable(readTableCrm.GetTable("ET_DATA"));
            foreach (DataRow rowCrm in reportCrm.Rows)
            {
                DataRow rowReportCrm = dataCrm.NewRow();
                rowReportCrm["MR"] = rowCrm["LINE"].ToString().Split('|')[0].Trim().TrimStart('0');
                rowReportCrm["FECHA"] = rowCrm["LINE"].ToString().Split('|')[1].Trim();

                dataCrm.Rows.Add(rowReportCrm);
            }

            for (int i = dataCrm.Rows.Count - 1; i >= 0; i--)
            {
                if (!dataCrm.Rows[i]["MR"].ToString().StartsWith("52")) //El número para el MR se maneja un consecutivo de momento está en 52
                {
                    dataCrm.Rows[i].Delete();
                }
            }

            #endregion


            if (dataCrm.Rows.Count < dataErp.Rows.Count)
            {
                mail.SendHTMLMail("Resultado Verificación de MRs SAP: " + "<br>" + "Se encontraron más MR's en ERP que en CRM, favor revisar", new string[] { "internalcustomersrvs@gbm.net" }, "Advertencia: No se pudo completar la revisión");

                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Reporte", "Resultado Verificación de MRs SAP: " + " <br> " + "Se encontraron más MR's en ERP que en CRM, favor revisar", root.Subject);
                respFinal = respFinal + "\\n" + "Crear Reporte resultado Verificación de MRs SAP: " + " <br> " + "Se encontraron más MR's en ERP que en CRM, favor revisar";

            }
            else if (dataCrm.Rows.Count > dataErp.Rows.Count)
            {
                DataTable dataResult = dataCrm.AsEnumerable().Except(dataErp.AsEnumerable(), DataRowComparer.Default).ToArray().CopyToDataTable();
                mail.SendHTMLMail("Resultado Verificación de MRs SAP: " + "<br>" + val.ConvertDataTableToHTML(dataResult), new string[] { "internalcustomersrvs@gbm.net" }, "Información sobre MRs que no viajaron:");

                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Reporte", "Resultado Verificación de MRs SAP: " + "<br>" + val.ConvertDataTableToHTML(dataResult), root.Subject);
                respFinal = respFinal + "\\n" + "Crear Reporte Resultado Verificación de MRs SAP: " + "<br>" + val.ConvertDataTableToHTML(dataResult);

            }
            else
            {
                mail.SendHTMLMail("Resultado Verificación de MRs SAP: " + "<br>" + "No hay MRs sin viajar.", new string[] { "internalcustomersrvs@gbm.net" }, "Información sobre MRs que no viajaron:");

                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Reporte", "Resultado Verificación de MRs SAP: " + "<br>" + "No hay MRs sin viajar.", root.Subject);
                respFinal = respFinal + "\\n" + "Crear Reporte Resultado Verificación de MRs SAP: " + "<br>" + "No hay MRs sin viajar.";

            }

            root.BDUserCreatedBy = "internalcustomersrvs";
            root.requestDetails = respFinal;


            console.WriteLine("Creando estadísticas...");
            //Se pone en el main debido a que trabaja con planner.
            using (Stats stats = new Stats())
            {
                stats.CreateStat();
            }

        }

    }
}
