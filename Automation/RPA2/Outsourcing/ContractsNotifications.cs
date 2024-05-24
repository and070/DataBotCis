using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Linq;
using System.Data;
using System;
using DataBotV5.Data.Database;

namespace DataBotV5.Automation.RPA2.Outsourcing
{
    /// <summary>
    /// Clase RPA dedicada a varias notificaciones de contratos nuevos ContractsNotifications

    /// de Outsourcing con External reference.
    /// </summary>
    class ContractsNotifications
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();
        string respFinal = "";

        bool executeStats = false;

        const string mandErp = "ERP";
        const string mandCrm = "CRM";

        public void Main()
        {

            List<DateTime> pendingDates = GetPendingDates("MatGroupsContracts");
            ProcessMatGroupsContracts(pendingDates);

            pendingDates = GetPendingDates("OutsourcingContracts");
            ProcessOutsourcingContracts(pendingDates);

            //pendingDates = GetPendingDates("OnHoldAndClosedContracts");
            //ProcessOnHoldAndClosedContracts(pendingDates);

            if (executeStats)
            {
                root.requestDetails = respFinal;
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }

        private void ProcessOnHoldAndClosedContracts(List<DateTime> pendingDates)
        {
            foreach (DateTime date in pendingDates)
            {
                DataSet consInfo = GetNewOnHoldAndClosedContracts(date.ToString("yyyy-MM-dd"), date.ToString("yyyy-MM-dd"));

                DataTable resDt = new DataTable();
                resDt.Columns.Add("Contrato");
                resDt.Columns.Add("Cliente");
                resDt.Columns.Add("Fecha de Inicio");
                resDt.Columns.Add("Fecha de Fin");
                resDt.Columns.Add("País");

                string sign = "<br><p><strong><span style=\"font-size: 12.0pt; font-family: 'Arial',sans-serif; color: #005898;\">Robotic Process Automation</span>" +
                "</strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\"><br />GBM Costa Rica</span><br />" +
                "<strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">GBM as a Service </span>" +
                "</strong><a href=\"http://www.gbm.net/\"><span style=\"font-size: 7.5pt; font-family: 'Arial',sans-serif;\">GBM</span></a>" +
                "<span style=\"font-size: 7.5pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">&nbsp;| </span><a href=\"https://www.facebook.com/GBMCorp?fref=ts\">" +
                "<span style=\"font-size: 7.5pt; font-family: 'Arial',sans-serif;\">Facebook</span></a><br /><strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">" +
                "<br />---------------------------------------------------</span>" +
                "</strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">&nbsp;</span>" +
                "<br /><em><span style=\"font-size: 12.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">Please do not reply this message, favor no responda sobre este mensaje.</span></em></p>";

                for (int i = 0; i < consInfo.Tables.Count; i++)
                {
                    executeStats = true;
                    resDt.Clear();
                    string consType = "";
                    DataTable contracts = consInfo.Tables[i];

                    if (i == 0)
                        consType = "OnHold";
                    else if (i == 1)
                        consType = "Closed";

                    foreach (DataRow contract in contracts.Rows)
                    {
                        DataRow resRow = resDt.NewRow();
                        resRow["Contrato"] = contract["CONTRACT"].ToString();
                        resRow["Cliente"] = contract["CUSTOMER_DESC"].ToString();
                        resRow["País"] = contract["SALES_ORG"].ToString();
                        resRow["Fecha de Inicio"] = contract["CON_START"].ToString().Substring(0, 4) + "-" + contract["CON_START"].ToString().Substring(4, 2) + "-" + contract["CON_START"].ToString().Substring(6, 2);
                        resRow["Fecha de Fin"] = contract["CON_END"].ToString().Substring(0, 4) + "-" + contract["CON_END"].ToString().Substring(4, 2) + "-" + contract["CON_END"].ToString().Substring(6, 2);

                        resDt.Rows.Add(resRow);
                    }

                    string dtHtml = val.ConvertDataTableToHTML(resDt);
                    string body = "Hola, se le envían los contratos que pasaron a estado <b>" + consType + "</b> el " + date.ToString("D") + "<br><br>" + dtHtml + sign;

                    mail.SendHTMLMail(body, new string[] { "cmanagement@gbm.net", "krodriguez@gbm.net", "jublanco@gbm.net" }, "NOTIFICACIÓN DE CONTRATOS QUE PASAN A ESTADO " + consType.ToUpper());
                    log.LogDeCambios("Notificacion", root.BDProcess, "Automatico", "Reporte de contratos", dtHtml, "");
                    respFinal = respFinal + "\\n" + "Reporte de contratos" + dtHtml;


                }
                root.BDUserCreatedBy = "CMANAGEMENT, KRODRIGUEZ, JUBLANCO";
            }
            SetExecutedTime("OnHoldAndClosedContracts");
        }
        private void ProcessMatGroupsContracts(List<DateTime> pendingDates)
        {
            foreach (DateTime date in pendingDates)
            {
                //Lista de MGs Validos
                List<string> evalList = new List<string>();
                evalList.AddRange(new string[] {
                "3010101",
                "3010102",
                "3010103",
                "3010104",
                "301010402",
                "3010105",
                "3010106",
                "301010701",
                "301010702",
                "301010703",
                "301010704",
                "301010705",
                "3010108",
                "3010109",
                "3010110",
                "3010111",
                "301050501",
                "3010401",
                "30104",
                "3010402",
                "30101",
                "30406",
                "3040601",
                "3040602",
                "3040603",
                "30407"
            });

                //Tabla de los contratos nuevos
                DataTable newCons = GetNewContracts(date.ToString("yyyy-MM-dd"), date.ToString("yyyy-MM-dd"));

                //Tabla con los Items de los contratos
                DataTable conItems = new DataTable();
                conItems.Columns.Add("Contrato");
                conItems.Columns.Add("Cliente");
                conItems.Columns.Add("Fecha de Inicio");
                conItems.Columns.Add("País");
                conItems.Columns.Add("items", typeof(List<string>));

                //Lista con todos los items para buscar los MG en ERP
                List<string> totalItems = new List<string>();

                foreach (DataRow contract in newCons.Rows)
                {
                    executeStats = true;
                    List<string> itemsList = new List<string>();
                    DataRow conItemRow = conItems.NewRow();
                    conItemRow["Contrato"] = contract["CONTRACT"].ToString();
                    conItemRow["Cliente"] = contract["CUSTOMER_DESC"].ToString();
                    conItemRow["País"] = contract["SALES_ORG"].ToString();
                    conItemRow["Fecha de Inicio"] = contract["CON_START"].ToString().Substring(0, 4) + "-" + contract["CON_START"].ToString().Substring(4, 2) + "-" + contract["CON_START"].ToString().Substring(6, 2);

                    DataTable equi = (DataTable)contract["EQUI"];
                    DataTable items = (DataTable)contract["ITEMS"];

                    foreach (DataRow equipment in equi.Rows)
                    {
                        totalItems.Add(equipment["ITEM_PROD"].ToString());
                        itemsList.Add(equipment["ITEM_PROD"].ToString());
                    }
                    foreach (DataRow item in items.Rows)
                    {
                        totalItems.Add(item["ITEM_PROD"].ToString());
                        itemsList.Add(item["ITEM_PROD"].ToString());
                    }

                    conItemRow["items"] = itemsList;

                    conItems.Rows.Add(conItemRow);
                }

                //Diccionarios con los MG de los items
                Dictionary<string, string> materialGroups = GetMg(totalItems);
                Dictionary<string, string> materialGroupsDesc = GetMgDescription(materialGroups);


                #region Cambiar la tabla de con->item a con->mg
                foreach (DataRow contract in conItems.Rows)
                {
                    List<string> temp = new List<string>();
                    List<string> materials = (List<string>)contract["items"];

                    foreach (string material in materials)
                    {
                        try
                        {
                            temp.Add(materialGroups[material]);
                        }
                        catch (Exception)
                        {
                            temp.Add("");
                        }
                    }

                    contract["items"] = temp;
                }
                #endregion

                #region Eliminar los contracts que no tienen mg valido
                for (int i = conItems.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow contractRow = conItems.Rows[i];

                    List<string> contractMaterialGroups = (List<string>)contractRow["items"];

                    if (!contractMaterialGroups.Any(x => evalList.Any(y => y == x))) //la lista contractMaterialGroups tiene algún elemento de evalList?
                        conItems.Rows.Remove(contractRow);
                }
                #endregion

                #region Enviar la respuestas
                if (conItems.Rows.Count > 0)
                {
                    #region Formatear la tabla para enviar
                    conItems.Columns.Add("Material Groups");
                    foreach (DataRow item in conItems.Rows)
                    {
                        List<string> mgIds = (List<string>)item["items"];
                        mgIds = mgIds.Distinct().ToList();

                        #region mgIds -> mgDesc
                        List<string> mgDescs = new List<string> { };
                        if (mgIds.Count > 0)
                            foreach (string mgId in mgIds)
                                mgDescs.Add(materialGroupsDesc[mgId]);
                        #endregion

                        item["Material Groups"] = String.Join("<br>", mgDescs.ToArray());
                    }
                    conItems.Columns.Remove("items");
                    #endregion

                    string sign = "<br><p><strong><span style=\"font-size: 12.0pt; font-family: 'Arial',sans-serif; color: #005898;\">Robotic Process Automation</span>" +
                                  "</strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\"><br />GBM Costa Rica</span><br />" +
                                  "<strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">GBM as a Service </span>" +
                                  "</strong><a href=\"http://www.gbm.net/\"><span style=\"font-size: 7.5pt; font-family: 'Arial',sans-serif;\">GBM</span></a>" +
                                  "<span style=\"font-size: 7.5pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">&nbsp;| </span><a href=\"https://www.facebook.com/GBMCorp?fref=ts\">" +
                                  "<span style=\"font-size: 7.5pt; font-family: 'Arial',sans-serif;\">Facebook</span></a><br /><strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">" +
                                  "<br />---------------------------------------------------</span>" +
                                  "</strong><span style=\"font-size: 10.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">&nbsp;</span>" +
                                  "<br /><em><span style=\"font-size: 12.0pt; font-family: 'Arial',sans-serif; color: #6d6e71;\">Please do not reply this message, favor no responda sobre este mensaje.</span></em></p>";

                    string dtHtml = val.ConvertDataTableToHTML(conItems);
                    string body = "Buenas, se le envían los contratos de <b>Datacenter Hybrid Cloud Tower</b> y <b>System Mgt Tower</b> que se crearon el " + date.ToString("D") + "<br><br>" + dtHtml + sign;

                    mail.SendHTMLMail(body, new string[] { "dcfm@gbm.net", "jvarela@gbm.net", "cmanagement@gbm.net" }, "Nuevos contratos Datacenter Hybrid Cloud Tower y System Mgt Tower");

                    log.LogDeCambios("Notificacion", root.BDProcess, "Automatico", "Reporte de contratos Datacenter Hybrid Cloud Tower y System Mgt Tower", dtHtml, "");
                    respFinal = respFinal + "\\n" + "Reporte de contratos Datacenter Hybrid Cloud Tower y System Mgt Tower" + dtHtml;

                }
                root.BDUserCreatedBy = "dcfm, jvarela, cmanagement";
                #endregion

            }
            SetExecutedTime("MatGroupsContracts");
        }
        private void ProcessOutsourcingContracts(List<DateTime> pendingDates)
        {
            string[] sender = { "calopez@gbm.net", "dsolis@gbm.net", "rsaborio@gbm.net", "kruilova@gbm.net", "cbarria@gbm.net", "aarguedas@gbm.net", "gbarahona@gbm.net", "jublanco@gbm.net", "KGonzalez@gbm.net" };

            foreach (DateTime pendingDate in pendingDates)
            {
                #region SAP
                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["FECHA_INI"] = pendingDate.ToString("yyyy-MM-dd"),
                    ["FECHA_FIN"] = pendingDate.ToString("yyyy-MM-dd")
                };

                IRfcFunction FM = sap.ExecuteRFC(mandCrm, "ZICS_GET_OUTSRC_NEW_CON", parameters);

                DataTable responseExt = sap.GetDataTableFromRFCTable(FM.GetTable("RESPONSE_EXT"));
                DataTable responseInfo = sap.GetDataTableFromRFCTable(FM.GetTable("RESPONSE_INFO"));

                #endregion

                if (responseExt.Rows.Count > 0)
                {
                    responseExt.Columns["ID_CON"].ColumnName = "Id del contrato";
                    responseExt.Columns["EXT_REF"].ColumnName = "External Reference";
                    responseExt.Columns.Add("Cliente");
                    responseExt.Columns.Add("Descripcion");
                    responseExt.Columns.Add("Pais");

                    responseExt.Columns["Descripcion"].SetOrdinal(1);

                    responseExt.AcceptChanges();
                    foreach (DataRow contract in responseExt.Rows)
                    {
                        if (contract["External Reference"].ToString() == "")
                            contract.Delete();
                        else
                        {
                            DataRow[] select = responseInfo.Select("CONTRACT = '" + contract["Id del contrato"] + "'");
                            contract["Cliente"] = select[0]["CUSTOMER_DESC"];
                            contract["Descripcion"] = select[0]["CONTRACT_DESC"];
                            contract["Pais"] = select[0]["SALES_ORG"];
                        }
                    }
                    responseExt.AcceptChanges();

                    string dtHtml = val.ConvertDataTableToHTML(responseExt);

                    if (responseExt.Rows.Count > 0)
                    {
                        executeStats = true;

                        //Generar el body del correo
                        string body = "Buenas, se le envían los contratos de outsourcing que se crearon el " + pendingDate.ToString("D") + "<br><br>" + dtHtml;

                        //Enviar las notificaciones
                        mail.SendHTMLMail(body, sender, "Nuevos contratos Outsourcing con External reference " + pendingDate.ToString("D"));

                        log.LogDeCambios("Notificacion", root.BDProcess, "Automatico", "Reporte de contratos OnHold", dtHtml, "");
                        respFinal = respFinal + "\\n" + "Reporte de contratos OnHold" + dtHtml;
                        root.BDUserCreatedBy = "KGONZALES";
                    }
                }
            }
            SetExecutedTime("OutsourcingContracts");
        }

        private DataSet GetNewOnHoldAndClosedContracts(string startDate, string endDate)
        {
            console.WriteLine("Leyendo nuevos contratos On Hold y Closed");

            DataSet result = new DataSet();

            List<string> onholdCon = new List<string>();
            List<string> closeCon = new List<string>();

            RfcDestination destination = sap.GetDestRFC(mandCrm);

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FECHA_INI"] = startDate,
                ["FECHA_FIN"] = endDate
            };

            IRfcFunction fm = sap.ExecuteRFC(mandCrm, "ZDM_GET_CONTRACT_RENEWAL", parameters);

            DataTable fmResponse = sap.GetDataTableFromRFCTable(fm.GetTable("RESPONSE_STATUS"));

            if (fmResponse.Rows.Count > 0)
            {
                foreach (DataRow con in fmResponse.Rows)
                {
                    if (con["STAT"].ToString() == "E0011")//onHold
                        onholdCon.Add(con["GUID"].ToString());
                    else if (con["STAT"].ToString() == "I1005")//Closed
                        closeCon.Add(con["GUID"].ToString());
                }
            }
            onholdCon = onholdCon.Distinct().ToList();
            closeCon = closeCon.Distinct().ToList();

            result.Tables.Add(GetContractData(onholdCon, destination));
            result.Tables.Add(GetContractData(closeCon, destination));

            return result;
        }
        private DataTable GetContractData(List<string> onholdCon, RfcDestination destination)
        {
            IRfcFunction fm = destination.Repository.CreateFunction("ZICS_GET_CONTRACT_DATA");
            IRfcTable zcontract = fm.GetTable("ZCONTRACT");

            foreach (string guid in onholdCon)
            {
                zcontract.Append();
                zcontract.SetValue("GUID", guid);
            }

            fm.Invoke(destination);
            IRfcTable sapResult = fm.GetTable("RESPONSE");
            return sap.GetDataTableFromRFCTable(sapResult);
        }
        private DataTable GetNewContracts(string sdate, string edate)
        {
            IRfcTable response;

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FECHA_INI"] = sdate,
                ["FECHA_FIN"] = edate
            };

            IRfcFunction func = sap.ExecuteRFC(mandCrm, "ZDM_GET_NEW_CONTRACT", parameters);

            response = func.GetTable("RESPONSE");

            DataTable tableReponse = sap.GetDataTableFromRFCTable(response);

            for (int i = 0; i < tableReponse.Rows.Count; i++)
            {
                tableReponse.Rows[i]["EQUI"] = sap.GetDataTableFromRFCTable(response[i].GetTable("EQUI"));
                tableReponse.Rows[i]["ITEMS"] = sap.GetDataTableFromRFCTable(response[i].GetTable("ITEMS"));
            }

            #region Eliminar los documentos 801 ya que no deben ir a CD(lo ideal seria hacerlo en la FM pero bueno)
            for (int i = tableReponse.Rows.Count - 1; i >= 0; i--)
            {
                DataRow fila = tableReponse.Rows[i];
                if (fila["CONTRACT"].ToString().StartsWith("801"))
                    tableReponse.Rows.Remove(fila);
            }
            #endregion

            return tableReponse;
        }
        private Dictionary<string, string> GetMg(List<string> totalItems)
        {
            Dictionary<string, string> matGroups = new Dictionary<string, string>();

            RfcDestination destErp = sap.GetDestRFC(mandErp);
            IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
            fmMg.SetValue("QUERY_TABLE", "MARA");
            fmMg.SetValue("DELIMITER", "|");

            IRfcTable fields = fmMg.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "MATNR");
            fields.Append();
            fields.SetValue("FIELDNAME", "MATKL");

            IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "MATNR IN (");

            foreach (string item in totalItems)
            {
                string prod = item;

                fmOptions.Append();
                fmOptions.SetValue("TEXT", "'" + prod + "',");

            }
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "'' )");
            fmMg.Invoke(destErp);

            foreach (DataRow row in sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA")).Rows)
            {
                string prod = row["LINE"].ToString().Split(new char[] { '|' })[0].Trim();
                if (!matGroups.ContainsKey(prod))
                {
                    string MG = row["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                    try { matGroups.Add(prod, MG); } catch (Exception) { }
                }
            }

            return matGroups;
        }
        private Dictionary<string, string> GetMgDescription(Dictionary<string, string> materialGroups)
        {
            Dictionary<string, string> matGroups = new Dictionary<string, string>();

            RfcDestination destErp = sap.GetDestRFC(mandErp);
            IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
            fmMg.SetValue("QUERY_TABLE", "T023T");
            fmMg.SetValue("DELIMITER", "|");

            IRfcTable fields = fmMg.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "MATKL");
            fields.Append();
            fields.SetValue("FIELDNAME", "WGBEZ60");

            IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "MATKL IN (");

            foreach (KeyValuePair<string, string> materialGroup in materialGroups)
            {
                //string prod = item;

                fmOptions.Append();
                fmOptions.SetValue("TEXT", "'" + materialGroup.Value + "',");

            }
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "'' )");
            fmMg.Invoke(destErp);

            foreach (DataRow row in sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA")).Rows)
            {
                string mgId = row["LINE"].ToString().Split(new char[] { '|' })[0].Trim();
                if (!matGroups.ContainsKey(mgId))
                {
                    string mgDesc = row["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                    try { matGroups.Add(mgId, mgDesc); } catch (Exception) { }
                }
            }

            return matGroups;
        }
        private DateTime GetLastExecutedTime(string method)
        {
            DataTable resDt = crud.Select("SELECT date FROM `last_executed_time` WHERE `method` = '" + method + "'", "contracts_notifications_db");
            string resDate = resDt.Rows[0]["date"].ToString();
            DateTime resDateTime = DateTime.Parse(resDate);
            return resDateTime;
        }
        private void SetExecutedTime(string method)
        {
            crud.Update("UPDATE `last_executed_time` SET `date` = '" + DateTime.Now.ToString("yyyy-MM-dd") + "' WHERE `last_executed_time`.`method` = '" + method + "'", "contracts_notifications_db");
        }
        private List<DateTime> GetPendingDates(string method)
        {
            DateTime lastExecDate = GetLastExecutedTime(method);
            TimeSpan daysSinceLastRunning = DateTime.Now.Date - lastExecDate.Date;
            List<DateTime> ret = new List<DateTime>();

            for (int i = 0; i < daysSinceLastRunning.Days; i++)
            {
                DateTime toCheck = DateTime.Now.Date.AddDays(-1 * i);
                ret.Add(toCheck);
            }

            return ret;
        }
    }
}
