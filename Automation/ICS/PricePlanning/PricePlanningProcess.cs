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

namespace DataBotV5.Automation.ICS.PricePlanning
{
    /// <summary>
    /// Clase ICS Automation encargada de agregar tarifas a los colaboradores.
    /// </summary>
    class PricePlanningProcess
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();

        string mand = "ERP";
        string respFinal = "";



        public void Main()
        {
            mail.GetAttachmentEmail("Solicitudes Tarifas Actualizadas", "Procesados", "Procesados Tarifas Actualizadas");
            if (root.ExcelFile != "")
            {
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessPricePlanning(excelDt);
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        /// <summary>Leer el correo con el personal number y robot para leer excel de tarifas.</summary>
        private void ProcessPricePlanning(DataTable excelDt)
        {
            bool valExcel = false;
            try
            {
                if (excelDt.Columns[2].ColumnName == "x")
                    valExcel = true;
            }
            catch (Exception) { }

            if (valExcel)//excel válido
            {
                excelDt.Columns.Remove("x");

                excelDt.Columns.Add("Respuesta Clase Actividad");
                excelDt.Columns.Add("Respuesta Tarifa");

                int fiscalYear = DateTime.Now.Year;
                int initialPeriod = DateTime.Now.Month;
                bool sendIcs = false;

                string reportPath = mail.GetLastPriceReportMail();
                DataTable excelPrices = excel.GetExcel(reportPath);

                foreach (DataRow reqPricesRow in excelDt.Rows)
                {
                    string employeeID = reqPricesRow[0].ToString().Trim(); //personal number
                    string getPriceResponse = reqPricesRow[1].ToString().Trim(); //tarifa 

                    console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                    Dictionary<string, string> parameters = new Dictionary<string, string>
                    {
                        ["PERSONAL_NUMBER"] = employeeID
                    };

                    IRfcFunction insertActivClass = sap.ExecuteRFC(mand, "ZFM_INSERT_ACTIVCLASS", parameters);

                    #region Procesar Salidas del FM
                    string ceCo = insertActivClass.GetValue("CECO").ToString();
                    string classActv = insertActivClass.GetValue("IV_CLAACT").ToString();
                    string activClassResponse = insertActivClass.GetValue("RESPUESTA").ToString();
                    string jobDesc = insertActivClass.GetValue("IV_STEXT").ToString();

                    console.WriteLine("Ejecución de Proceso CECO y Clase de Actividad: " + ceCo + ": " + classActv + " Colaborador: " + employeeID);
                    #endregion

                    if (activClassResponse != "No se encontro infotipo vigente")
                    {
                        if (getPriceResponse == "")
                            getPriceResponse = GetPriceFromReport(excelPrices, classActv, ceCo, jobDesc);

                        if (getPriceResponse.Contains("No se encontró") || getPriceResponse == "0")
                        {
                            if (getPriceResponse == "0")
                                getPriceResponse = "La tarifa esta en blanco en el reporte";

                            sendIcs = true;
                            reqPricesRow["Respuesta Tarifa"] = "ERROR <b>" + getPriceResponse + "</b>. CLASSACTV: <b>" + classActv + "</b> CECO: <b>" + ceCo + "</b> JOBDESC: <b>" + jobDesc + "</b>";
                            reqPricesRow["Respuesta Clase Actividad"] = "";

                        }
                        else
                        {
                            if (classActv != "")
                            {
                                string[] priceInsertResponse = InsertPricePlanning(fiscalYear, initialPeriod, ceCo, getPriceResponse, classActv, mand);

                                if (priceInsertResponse[0] == "Se han registrado las tarifas correctamente" || priceInsertResponse[0] == "Ya existen registros para ese Periodo")
                                {
                                    log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Proceso Completado de Tarifa para: " + ceCo + ": " + classActv + " Colaborador: " + employeeID + " En el Periodo: " + fiscalYear, root.Subject, "");
                                    respFinal = respFinal + "\\n" + "Proceso Completado de Tarifa para: " + ceCo + ": " + classActv + " Colaborador: " + employeeID + " En el Periodo: " + fiscalYear;

                                    //Carga de segundo periodo para tarifa
                                    fiscalYear++;
                                    priceInsertResponse = InsertPricePlanning(fiscalYear, 1, ceCo, getPriceResponse, classActv, mand);
                                    if (priceInsertResponse[0] == "Se han registrado las tarifas correctamente" || priceInsertResponse[0] == "Ya existen registros para ese Periodo")
                                    {
                                        log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Proceso Completado de Tarifa para: " + ceCo + ": " + classActv + " Colaborador: " + employeeID + " En el Periodo: " + fiscalYear, root.Subject, "");
                                        respFinal = respFinal + "\\n" + "Proceso Completado de Tarifa para: " + ceCo + ": " + classActv + " Colaborador: " + employeeID + " En el Periodo: " + fiscalYear;

                                        reqPricesRow["Respuesta Tarifa"] = priceInsertResponse[0];
                                        reqPricesRow["Respuesta Clase Actividad"] = activClassResponse;
                                    }
                                    else if (priceInsertResponse[0] == "Carga de datos Fallida")
                                    {
                                        sendIcs = true;
                                        reqPricesRow["Respuesta Tarifa"] = "ERROR";
                                        reqPricesRow["Respuesta Clase Actividad"] = "ERROR";
                                    }
                                }
                                else if (priceInsertResponse[0] == "Carga de datos Fallida")
                                {
                                    sendIcs = true;
                                    reqPricesRow["Respuesta Tarifa"] = "ERROR: <b>" + priceInsertResponse[1] + "</b>";
                                    reqPricesRow["Respuesta Clase Actividad"] = "ERROR";
                                }
                            }
                            else
                            {
                                sendIcs = true;
                                reqPricesRow["Respuesta Tarifa"] = "Sin procesar";
                                reqPricesRow["Respuesta Clase Actividad"] = activClassResponse;
                            }
                        }

                    }
                    else
                    {
                        sendIcs = true;
                        reqPricesRow["Respuesta Tarifa"] = "";
                        reqPricesRow["Respuesta Clase Actividad"] = "ERROR <b>" + activClassResponse + "</b>. CECO: <b>" + ceCo + "</b> JOBDESC: <b>" + jobDesc + "</b>";
                    }
                    excelDt.AcceptChanges();

                }

                if (sendIcs)
                    mail.SendHTMLMail("Fallo Carga de tarifas<br><br>" + val.ConvertDataTableToHTML(excelDt), new string[] { "internalcustomersrvs@gbm.net" }, root.BDProcess);
                else
                    mail.SendHTMLMail("Carga Satisfactoria de tarifas<br><br>" + val.ConvertDataTableToHTML(excelDt), new string[] { root.BDUserCreatedBy }, root.Subject);
                
                root.requestDetails = respFinal;

            }
            else
                mail.SendHTMLMail("Por favor utilizar la plantilla oficial de carga de tarifas", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
        }

        private string GetPriceFromReport(DataTable excelPrices, string sapClaact, string cencos, string job)
        {
            for (int i = 0; i < excelPrices.Columns.Count; i++)
                excelPrices.Columns[i].ColumnName = "col" + i;

            DataTable infoSap = new DataTable();
            infoSap.Columns.Add("JOB");
            infoSap.Columns.Add("CENCOS");
            infoSap.Columns.Add("CLAACT");
            infoSap.Columns.Add("ReportCLAACT");

            RfcDestination destErp = new SapVariants().GetDestRFC(mand);
            IRfcFunction fmTable = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            fmTable.SetValue("USE_ET_DATA_4_RETURN", "X");
            fmTable.SetValue("QUERY_TABLE", "zhr_update_0315");
            fmTable.SetValue("DELIMITER", "|");

            IRfcTable fields = fmTable.GetTable("FIELDS");

            fields.Append();
            fields.SetValue("FIELDNAME", "JOB"); //campos a traer
            fields.Append();
            fields.SetValue("FIELDNAME", "CENCOS");
            fields.Append();
            fields.SetValue("FIELDNAME", "CLAACT");

            fmTable.Invoke(destErp);

            DataTable report = sap.GetDataTableFromRFCTable(fmTable.GetTable("ET_DATA"));
            foreach (DataRow row in report.Rows)
            {
                DataRow newInfoRow = infoSap.NewRow();
                newInfoRow["JOB"] = row["LINE"].ToString().Split(new char[] { '|' })[0].Trim().TrimStart(new char[] { '0' });
                newInfoRow["CENCOS"] = row["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                newInfoRow["CLAACT"] = row["LINE"].ToString().Split(new char[] { '|' })[2].Trim();

                infoSap.Rows.Add(newInfoRow);
            }

            foreach (DataRow dr in infoSap.Rows)
            {
                string claAct = dr.ItemArray[2].ToString();
                string jobDesc = dr.ItemArray[0].ToString().Trim();

                if (claAct.Contains("BA") && !jobDesc.EndsWith("ADV") || jobDesc.EndsWith("BAS"))
                    dr["ReportCLAACT"] = "BAS";
                else if (claAct.Contains("J") || claAct.Contains("JUN") || jobDesc.EndsWith("JNR"))
                    dr["ReportCLAACT"] = "JNR";
                else if (claAct.Contains("ST") && !jobDesc.EndsWith("ADV") && !jobDesc.EndsWith("SNR") || jobDesc.EndsWith("STD"))
                    dr["ReportCLAACT"] = "STD";
                else if (claAct.Contains("AD") && !jobDesc.EndsWith("SNR") || jobDesc.EndsWith("ADV"))
                    dr["ReportCLAACT"] = "ADV";
                else if (claAct.Contains("SN") || claAct.Contains("SR") || jobDesc.EndsWith("SNR"))
                    dr["ReportCLAACT"] = "SNR";
                else
                    dr["ReportCLAACT"] = "No se pudo clasificar el activity";
            }

            DataRow[] selectInfoSap = infoSap.Select("CLAACT ='" + sapClaact + "' AND CENCOS = '" + cencos + "' AND JOB = '" + job + "'");

            string price;

            string reportClaact = "No se encontró claact en SAP";

            if (selectInfoSap.Length >= 1)
            {
                reportClaact = selectInfoSap[0]["ReportCLAACT"].ToString();

                DataRow[] selectExcelPrices = excelPrices.Select("col2 ='" + reportClaact + "' AND col1 = '" + cencos + "' AND col0 = '" + job + "'");
                if (selectExcelPrices.Length == 1)
                {
                    price = selectExcelPrices[0]["col3"].ToString();

                    //arreglar price

                    char sepDec = Convert.ToChar(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                    price = price.Replace(',', sepDec);
                    price = price.Replace('.', sepDec);
                    double.TryParse(price, out double priceDouble);
                    priceDouble = Math.Round(priceDouble, 2);
                    price = priceDouble.ToString();
                    price = price.Replace(',', '.');

                }
                else
                    price = "No se encontró tarifa en el Reporte";
            }
            else
                price = reportClaact;

            return price;
        }

        /// <summary> Método de inserción de tarifa, con dos string respuesta y errores.</summary>
        string[] InsertPricePlanning(int fiscalYear, int initialPeriod, string ceCo, string price, string classActv, string mand)
        {
            string[] ret = new string[2];

            price = price.Replace(",", ".");  //investigar mejor como hacerlo mejor

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FISCALYEAR"] = fiscalYear.ToString(),
                ["CECO"] = ceCo,
                ["CLASSACTIVITY"] = classActv,
                ["INITIALPERIOD"] = initialPeriod.ToString(),
                ["PRICE"] = price
            };

            IRfcFunction func = sap.ExecuteRFC(mand, "ZFM_INSERT_TARIFA", parameters);

            #region Procesar Salidas del FM
            ret[0] = func.GetValue("RESPUESTA").ToString();

            DataTable retDt = sap.GetDataTableFromRFCTable(func.GetTable("RET"));

            foreach (DataRow row in retDt.Rows)
                ret[1] += row["MESSAGE"].ToString() + " :-: ";
            #endregion

            return ret;
        }
    }
}