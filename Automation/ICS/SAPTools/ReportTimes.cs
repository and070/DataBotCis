using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using System.Globalization;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Threading;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.SAPTools
{
    internal class ReportTimes
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        public void Main()
        {
            mail.GetAttachmentEmail("Solicitudes Reporte Tiempos", "Procesados", "Procesados Reporte Tiempos");
            if (root.ExcelFile != "")
            {
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessReportTime(excelDt);
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }

        private void ProcessReportTime(DataTable excelDt)
        {
            string respFinal = "";
            string errorMsg = "";

            bool valExcel = false;
            bool sendIcs = false;
            bool sendUser = false;

            DataTable errorTable = new DataTable();
            errorTable.Columns.Add("MENSAJE");
            errorTable.Columns.Add("MENSAJE1");
            errorTable.Columns.Add("MENSAJE2");
            errorTable.Columns.Add("FILA");

            try
            {
                if (excelDt.Columns[excelDt.Columns.Count - 1].ColumnName == "x")
                    valExcel = true;
            }
            catch (Exception) { }

            if (valExcel)
            {
                try
                {
                    excelDt.Columns.Remove("x");
                    excelDt.Columns.Add("Resultado");

                    string sender = root.BDUserCreatedBy.ToUpper();
                    string employeeID = val.GetEmployeeID(sender.Replace("@GBM.NET", ""));

                    RfcDestination destErp = new SapVariants().GetDestRFC("ERP");
                    IRfcFunction cat2Fm = destErp.Repository.CreateFunction("BAPI_CATIMESHEETMGR_INSERT");

                    IRfcTable catsRecordsIn = cat2Fm.GetTable("CATSRECORDS_IN");
                    IRfcTable returnTable = cat2Fm.GetTable("RETURN");

                    foreach (DataRow timeRow in excelDt.Rows)
                    {
                        string error315 = "";
                        string ceCo = "";
                        string actType = "";

                        CultureInfo cultureOriginal = (CultureInfo)CultureInfo.CurrentCulture.Clone();
                        CultureInfo cultureNew = (CultureInfo)CultureInfo.CurrentCulture.Clone();
                        cultureNew.DateTimeFormat.LongTimePattern = "HH:mm:ss";
                        Thread.CurrentThread.CurrentCulture = cultureNew;

                        string start = timeRow["Hora Inicial"].ToString().Trim();
                        string end = timeRow["Hora Final"].ToString().Trim();

                        Thread.CurrentThread.CurrentCulture = cultureOriginal;

                        string date = timeRow["Fecha"].ToString().Trim();
                        string order = timeRow["Orden"].ToString().Trim();
                        string type = timeRow["Tipo"].ToString().Trim();
                        string description = timeRow["Descripción"].ToString().Trim();
                        string network = timeRow["Network"].ToString().Trim();
                        string netItem = timeRow["Network Item"].ToString().Trim();

                        timeRow["Hora Inicial"] = Convert.ToDateTime(start).TimeOfDay.ToString();
                        timeRow["Hora Final"] = Convert.ToDateTime(end).TimeOfDay.ToString();

                        //Validaciones
                        employeeID = employeeID.ToUpper().Replace("AA", "");
                        start = Convert.ToDateTime(start).TimeOfDay.ToString().Replace(":", "");//hhmmss
                        end = Convert.ToDateTime(end).TimeOfDay.ToString().Replace(":", "");
                        type = type.Split('-')[0].Trim();

                        //Validar si tiene o no Orden
                        if (order == "" && network == "")
                        {
                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["PERNR"] = employeeID,
                                ["DATE"] = date
                            };
                            try
                            {
                                IRfcFunction catsGetInfoType0315 = sap.ExecuteRFC("ERP", "CATS_GET_INFOTYPE_0315", parameters);
                                IRfcStructure i0315 = catsGetInfoType0315.GetStructure("I0315");
                                DataTable i0315Dt = sap.GetDataTableFromRFCStructure(i0315);

                                ceCo = i0315Dt.Rows[0]["KOSTL"].ToString();
                                actType = i0315Dt.Rows[0]["LSTAR"].ToString();
                            }
                            catch (RfcAbapException ex)
                            {
                                sendIcs = true;
                                error315 = ex.Message;
                            }

                            order = "";
                        }
                        else
                        {
                            if (order != "")
                                order = order.PadLeft(12, '0'); //tiene que llenarse con 12 dígitos
                            else
                            {
                                if (network != "")
                                {
                                    network = network.PadLeft(12, '0'); //tiene que llenarse con 12 dígitos
                                    netItem = netItem.PadLeft(4, '0'); //tiene que llenarse con 12 dígitos
                                }
                            }
                        }

                        if (error315 == "")
                        {
                            catsRecordsIn.Append();
                            catsRecordsIn.SetValue("EMPLOYEENUMBER", employeeID);
                            catsRecordsIn.SetValue("WORKDATE", date);
                            catsRecordsIn.SetValue("TASKTYPE", type);
                            catsRecordsIn.SetValue("STARTTIME", start);
                            catsRecordsIn.SetValue("ENDTIME", end);
                            catsRecordsIn.SetValue("SHORTTEXT", description);

                            catsRecordsIn.SetValue("SEND_CCTR", ceCo);
                            catsRecordsIn.SetValue("ACTTYPE", actType);
                            catsRecordsIn.SetValue("REC_CCTR", ceCo);

                            if (order != "")
                                catsRecordsIn.SetValue("REC_ORDER", order);
                            if (network != "")
                            {
                                catsRecordsIn.SetValue("NETWORK", network);
                                catsRecordsIn.SetValue("ACTIVITY", netItem);
                            }
                        }
                        else
                            timeRow["Resultado"] = "infotipo 315: " + error315;

                    }
                    console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                    cat2Fm.SetValue("PROFILE", "ZGBM");
                    cat2Fm.Invoke(destErp);

                    DataTable returnDt = sap.GetDataTableFromRFCTable(returnTable);

                    if (returnDt.Rows.Count != 0)
                    {
                        sendUser = true;
                        List<string> ignoredErrorsIds = new List<string> { "207", "335", "201" };
                        foreach (DataRow item in returnDt.Rows)
                        {
                            DataRow errorRow = errorTable.NewRow();
                            errorRow[0] = item["MESSAGE"].ToString();
                            errorRow[1] = item["MESSAGE_V1"].ToString();
                            errorRow[2] = item["MESSAGE_V2"].ToString();
                            errorRow[3] = item["ROW"].ToString();
                            errorTable.Rows.Add(errorRow);

                            if (!ignoredErrorsIds.Contains(item["NUMBER"].ToString()))
                                sendIcs = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorMsg = ex.Message;
                    sendIcs = true;
                }

                if (sendUser)
                {
                    string[] cc = sendIcs ? new string[] { "internalcustomersrvs@gbm.net" } : null;
                    mail.SendHTMLMail("Falló Carga de horas<br><br>Solicitud: <br><br>" + val.ConvertDataTableToHTML(excelDt) + "<br><br>Error:<br><b>" + errorMsg + val.ConvertDataTableToHTML(errorTable) + "</b>", new string[] { root.BDUserCreatedBy }, root.BDClass, cc);
                }
                else
                {
                    mail.SendHTMLMail("Carga Satisfactoria de horas<br><br>" + val.ConvertDataTableToHTML(excelDt), new string[] { root.BDUserCreatedBy }, root.Subject);
                    log.LogDeCambios("", root.BDClass, root.BDUserCreatedBy, root.BDClass, val.ConvertDataTableToHTML(excelDt), "");
                    respFinal = respFinal + "\\n" + $"Carga Satisfactoria de horas {root.BDUserCreatedBy}: " + val.ConvertDataTableToHTML(excelDt);
                }
            }
            else
                mail.SendHTMLMail("Por favor utilizar la plantilla oficial de carga de horas", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

            root.requestDetails = respFinal;
        }
    }
}
