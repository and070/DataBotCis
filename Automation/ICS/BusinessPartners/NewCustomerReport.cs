using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using System.Windows.Forms;
using DataBotV5.Data.Stats;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Threading;
using System.Data;
using System;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.ICS.BusinessPartners
{
    /// <summary>
    /// Clase ICS que genera un Reporte Semanal Clientes.
    /// </summary>
    class NewCustomerReport
    {
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ActiveDirectory ad = new ActiveDirectory();
        readonly PowerAutomate flow = new PowerAutomate();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        ConsoleFormat console = new ConsoleFormat();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Stats stats = new Stats();
        readonly CRUD crud = new CRUD();
        Settings sett = new Settings();
        Log log = new Log();

        const string mand = "ERP";

        string respFinal = "";


        public void Main()
        {
            #region Enviar Reporte de clientes
            if (!sap.CheckLogin(mand))
            {
                sap.BlockUser(mand, 1);
                ProcessReport();
     
                sap.BlockUser(mand, 0);
            }
            else
            {
                sett.setPlannerAgain();
            }
            #endregion

            #region Procesar Respuesta de los aprobadores
            GetApproval();
            #endregion
        }

        private void GetApproval()
        {
            Dictionary<string, string> resAppr = flow.GetApprovalRequests(root.BDProcess);

            if (resAppr.Count > 0)
            {
                string apprStatus = JObject.Parse(resAppr["ResponseJson"])["RESPONSE"].ToString();

                string subject = JObject.Parse(resAppr["OriginalJson"])["appr_title"].ToString();
                //string approver = JObject.Parse(resAppr["OriginalJson"])["approver"].ToString();
                string customersTable = JObject.Parse(JObject.Parse(resAppr["OriginalJson"])["specific_data"].ToString())["customer_table"].ToString();
                string comments = JObject.Parse(resAppr["ResponseJson"])["COMMENTS"].ToString();
                string copyEmails = JObject.Parse(JObject.Parse(resAppr["OriginalJson"])["specific_data"].ToString())["cc"].ToString();
                string endDate = JObject.Parse(JObject.Parse(resAppr["OriginalJson"])["specific_data"].ToString())["endDate"].ToString();
                string startDate = JObject.Parse(JObject.Parse(resAppr["OriginalJson"])["specific_data"].ToString())["startDate"].ToString();
                string managerFullName = JObject.Parse(JObject.Parse(resAppr["OriginalJson"])["specific_data"].ToString())["managerFullName"].ToString();

                if (apprStatus.ToLower() == "approve")
                {
                    string msg = "Se aprobaron los siguiente clientes creados desde el " + startDate + " al " + endDate + " por " + managerFullName + ":<br><br>" + customersTable;
                    mail.SendHTMLMail(msg + "<br>Con los siguientes comentarios: " + comments, new string[] { "internalcustomersrvs@gbm.net" }, subject);

                    string resp = "Se aprobaron los siguiente clientes creados desde el " + startDate + " al " + endDate + " por " + managerFullName + ": " + customersTable + "Con los siguientes comentarios: " + comments;
                   
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Reporte Clientes", resp, root.Subject);
                    respFinal = respFinal + "\\n" + "Aprobación clientes: " + resp;

                    //enviar correos a las copias
                    string[] senders = copyEmails.Split(',');
                    mail.SendHTMLMail(msg, senders, subject);

                    root.BDUserCreatedBy = senders[0];

                }
                else
                {
                    mail.SendHTMLMail("Se rechazaron los siguiente clientes:<br><br>" + customersTable + "<br>Con los siguientes comentarios: " + comments, new string[] { "internalcustomersrvs@gbm.net" }, subject);

                    string resp = "Se rechazaron los siguiente clientes: " + customersTable + "Con los siguientes comentarios: " + comments;

                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Reporte Clientes", resp, root.Subject);
                    respFinal = respFinal + "\\n" + "Rechazo clientes: " + resp;

                    root.BDUserCreatedBy = "internalcustomersrvs";
                }

                root.requestDetails = respFinal;

                console.WriteLine(DateTime.Now + " > > > " + "Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }


            }
        }
        private void ProcessReport()
        {
            string[] startDate = GetDate(true);
            string[] endDate = GetDate(false);

            RfcDestination destErp = new SapVariants().GetDestRFC(mand);
            IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
            fmMg.SetValue("QUERY_TABLE", "KNA1");
            fmMg.SetValue("DELIMITER", "");

            IRfcTable fields = fmMg.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "KUNNR");

            IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "( ERDAT BETWEEN '" + startDate[2] + startDate[1] + startDate[0] + "' AND '" + endDate[2] + endDate[1] + endDate[0] + "' ) AND KTOKD = 'ZGBM'");

            fmMg.Invoke(destErp);

            DataTable tableSap = sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA"));

            string weekCustonmer = ConvertDataTableToString(tableSap); //trasforma el valor a string

            Thread thread = new Thread(() => Clipboard.SetText(weekCustonmer)); // copia en clipboard los clientes
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            GetTableFromSapGui();

            DataTable report = excel.GetExcel(root.FilesDownloadPath + "\\Reporte Semanal Clientes.XLSX");

            DataView viewReport = new DataView(report);
            DataTable customerReport = viewReport.ToTable("Selected", false, "Customer", "Deletion Flag", "Sales Organization", "Territory", "Name", "Personnel number", "Last Name", "2do. Last Name", "First name", "Industry", "Description"); // selecciona las filas que nos interesan

            DataView viewSo = new DataView(report); // toma las Sales Organization contenidas en el reporte
            DataTable selectedSO = viewSo.ToTable("Selected", false, "Sales Organization", "Territory");

            DataView soNotRepeated = new DataView(selectedSO);
            DataTable soUnique = soNotRepeated.ToTable(true, "Sales Organization", "Territory"); // elimina las SO duplicadas

            DataTable customerReportFinal = customerReport;

            DataSet final = new DataSet("Reporte");

            for (int j = 0; j < soUnique.Rows.Count; j++)
            {
                DataRow dr1 = soUnique.Rows[j];
                DataTable finalTable = final.Tables.Add(dr1["Territory"].ToString() + " " + dr1["Sales Organization"].ToString());

                finalTable.Columns.Add("Cliente");
                finalTable.Columns.Add("Territorio");
                finalTable.Columns.Add("Razón social");
                finalTable.Columns.Add("ID representante");
                finalTable.Columns.Add("Nombre");
                finalTable.Columns.Add("Industria");


                for (int k = 0; k < customerReportFinal.Rows.Count; k++)
                {
                    DataRow dr2 = customerReportFinal.Rows[k];

                    if (dr2["Territory"].ToString() + " " + dr2["Sales Organization"].ToString() == dr1["Territory"].ToString() + " " + dr1["Sales Organization"].ToString())
                    {
                        DataRow x = final.Tables[j].NewRow();
                        x["Cliente"] = customerReportFinal.Rows[k]["Customer"].ToString();
                        x["Territorio"] = customerReportFinal.Rows[k]["Territory"].ToString();
                        x["Razón social"] = customerReportFinal.Rows[k]["Name"].ToString();
                        x["ID representante"] = customerReportFinal.Rows[k]["Personnel number"].ToString();
                        x["Nombre"] = customerReportFinal.Rows[k]["First name"].ToString() + " " + customerReportFinal.Rows[k]["Last Name"].ToString();
                        x["Industria"] = customerReportFinal.Rows[k]["Industry"].ToString() + "(" + customerReportFinal.Rows[k]["Description"].ToString() + ")";
                        final.Tables[j].Rows.Add(x);
                    }
                }

            }

            for (int i = final.Tables.Count - 1; i >= 0; i--) // elimina wtc0, itc0 y los Md01 que el representante no empiezan con 1
            {
                DataTable dt1 = final.Tables[i];

                if (dt1.ToString() == "GBM Direct WTC0" || dt1.ToString() == "GBM Direct IT01" || dt1.ToString() == "Premium Account IT01" || dt1.ToString() == "Premium Account WTC0")
                    final.Tables.Remove(dt1);

                else if (dt1.ToString() == "GBM Direct MD01" || dt1.ToString() == "Premium Account MD01")
                {
                    for (int j = dt1.Rows.Count - 1; j >= 0; j--)
                    {
                        DataRow dr3 = dt1.Rows[j];

                        if (dr3["ID representante"].ToString().Substring(0, 1) != "1")
                            dt1.Rows.Remove(dr3);
                    }

                    if (dt1.Rows.Count.ToString() == "0")
                        final.Tables.Remove(dt1);
                }
            }

            DataTable mailsDt = crud.Select("SELECT * FROM `emails`", "customer_report_db");

            for (int k = 0; k < final.Tables.Count; k++)
            {
                DataTable dt2 = final.Tables[k];
                string customersCountry = dt2.TableName;
                try
                {
                    string[] mailDirection = GetMails(customersCountry, "manager", mailsDt);

                    Dictionary<string, string> infoManagerAd = ad.GetAdData(mailDirection[0]);

                    string[] mailCC = GetMails(customersCountry, "cc", mailsDt);
                    string managerName = infoManagerAd["Name"];
                    string managerFullName = infoManagerAd["FullName"];
                    string mailSubject = "Validar industria y representante de ventas. " + customersCountry + ".";
                    string formattedStartDate = startDate[0] + "/" + startDate[1] + "/" + startDate[2];
                    string formattedEndDate = endDate[0] + "/" + endDate[1] + "/" + endDate[2];
                    string mailContent = GetEmailBodyMarkdown(dt2, formattedStartDate, formattedEndDate, managerName);

                    string json = "{" +
                                    "\"customer_table\":\"" + val.ConvertDataTableToHTML(dt2).Replace("\"", "\\\"") + "\"," +
                                    "\"cc\":\"" + string.Join(",", mailCC) + "\"," +
                                    "\"startDate\":\"" + formattedStartDate + "\"," +
                                    "\"endDate\":\"" + formattedEndDate + "\"," +
                                    "\"managerFullName\":\"" + managerFullName + "\"" +
                                  "}";
                    flow.SendApproval(mailSubject, mailDirection[0], mailContent,  "internalcustomersrvs@gbm.net" , json, root.BDProcess);
                }
                catch (Exception ex)
                {
                    mail.SendHTMLMail("Error enviando reporte de clientes para " + customersCountry + ".<br><br>Exception: " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                }
            }
        }
        private void GetTableFromSapGui()
        {
            proc.KillProcess("saplogon", false);
            sap.LogSAP(mand.ToString());
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
            ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "Y_BP_LIST_001";
            ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(8);
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btn%_CUSTO_ID_%_APP_%-VALU_PUSH")).Press(); //pega los clientes que estaban en clipboard
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[24]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[8]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
            ((SAPFEWSELib.GuiMenu)SapVariants.session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]")).Select();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = root.FilesDownloadPath;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "Reporte Semanal Clientes.XLSX";
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nex";
            ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);

            Thread.Sleep(3000);
            proc.KillProcess("saplogon", false);
            proc.KillProcess("EXCEL", true);
        }
        private string[] GetDate(bool monday)
        {
            DateTime dt = DateTime.Now;
            if (monday)
            {
                int diff = (7 + (dt.DayOfWeek - DayOfWeek.Monday)) % 7;
                DateTime lastMonday = dt.AddDays(-1 * diff).Date;
                return new string[] { lastMonday.Day.ToString().PadLeft(2, '0'), lastMonday.Month.ToString().PadLeft(2, '0'), lastMonday.Year.ToString().PadLeft(4, '0') };
            }
            else
                return new string[] { dt.Day.ToString().PadLeft(2, '0'), dt.Month.ToString().PadLeft(2, '0'), dt.Year.ToString().PadLeft(4, '0') };
        }
        private string[] GetMails(string territory, string type, DataTable mailsDt)
        {
            DataRow[] result = mailsDt.Select("territory = '" + territory + "'");
            if (type == "manager")
                return new string[] { result[0]["manager"].ToString() };
            else
                return result[0]["copies"].ToString().Split(',');
        }
        private string ConvertDataTableToString(DataTable dataTable)
        {
            string data = string.Empty;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                DataRow row = dataTable.Rows[i];
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    data += row[j] + Environment.NewLine;
                    if (j == dataTable.Columns.Count - 1)
                    {
                        if (i != (dataTable.Rows.Count - 1))
                            data += "";
                    }
                    else
                        data += "";
                }
            }
            return data;
        }
        private string GetEmailBodyMarkdown(DataTable dt2, string startDate, string endDate, string managerName)
        {
            string mailContent = "Buenas días " + managerName;
            mailContent += "\n" + "\n";
            mailContent += "Con el objeto de no retrasar el proceso de la creación de clientes, se envía esta lista de clientes a su cargo creados desde el " + startDate + " al " + endDate + ", donde debemos validar la industria, el representante de ventas y el territorio.";
            mailContent += "\n" + "\n";
            mailContent += "Favor su visto bueno, e indicarnos si se requiere actualizar alguno de los representantes en SAP.";
            mailContent += "\n" + "\n";
            mailContent += "Por favor enviarnos su respuesta, es de suma importancia como proceso en Internal Customer Services, y como proceso de Auditoria para tener actualizados los representantes en SAP.";
            mailContent += "\n" + "\n";
            mailContent += val.ConvertDataTableToMarkdown(dt2);
            mailContent += "\n" + "\n";
            mailContent += "Saludos";
            return mailContent;
        }
    }
}
