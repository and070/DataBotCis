using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.CrmOrders
{
    /// <summary>
    /// Clase ICS Automation encargada de la creación de Ordenes de Servicio en CRM.
    /// </summary>
    class CreateZSOM
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();

        public void Main()
        {
            //leer correo
            mail.GetAttachmentEmail("Solicitudes ZSOM", "Procesados", "Procesados ZSOM");
            if (root.ExcelFile != null && root.ExcelFile != "")
            {
                bool isValidMail = IsValidMail(root.BDUserCreatedBy.ToUpper());
                if (isValidMail)
                {
                    ProcessCreateZSOM(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    using (Stats stats = new Stats()) { stats.CreateStat(); }
                }
            }
        }
        private void ProcessCreateZSOM(string excelPath)
        {
            DataTable excelDt = excel.GetExcel(excelPath);

            bool sendIcs = false;

            DataTable responseTable = new DataTable();
            responseTable.Columns.Add("Descripción");
            responseTable.Columns.Add("Ticket");
            responseTable.Columns.Add("Id de la Orden");
            responseTable.Columns.Add("Resultado");

            DataTable distinctIds = excelDt.DefaultView.ToTable(true, "ID");

            RfcDestination destCrm = sap.GetDestRFC("CRM");
            IRfcFunction zicsCreateZsom = destCrm.Repository.CreateFunction("ZICS_CREATE_ZSOM");
            IRfcTable itLinesData = zicsCreateZsom.GetTable("IT_LINES_DATA");

            foreach (DataRow row in distinctIds.Rows)
            {
                DataRow responseTableRow = responseTable.NewRow();

                string id = row["Id"].ToString();

                itLinesData.Clear();
                DataRow[] servOrder = new DataRow[0];
                servOrder = excelDt.Select("Id ='" + id + "'");
                if (servOrder.Length == 0)
                    servOrder = excelDt.Select("Id =" + id + "");

                string customerId = servOrder[0]["cliente"].ToString().Trim();
                string description = servOrder[0]["desc"].ToString().Trim();
                string ticketCd = servOrder[0]["ticket"].ToString().Trim();

                string serviceEmployee = servOrder[0]["tematica"].ToString().Trim();
                serviceEmployee = serviceEmployee.Split('-')[0].Trim();

                string salesOrg = servOrder[0]["pais(SALES_ORG)"].ToString().Trim();
                salesOrg = salesOrg.Split('-')[0].Trim();

                string serviceOrg = servOrder[0]["servicio(SERVICE_ORG)"].ToString().Trim();
                serviceOrg = serviceOrg.Split('-')[0].Trim();

                if (!customerId.StartsWith("00"))
                    customerId = "00" + customerId;

                zicsCreateZsom.SetValue("IV_CUSTOMER_ID", customerId);
                zicsCreateZsom.SetValue("IV_SALES_ORG", salesOrg);
                zicsCreateZsom.SetValue("IV_SERVICE_ORG", serviceOrg);
                zicsCreateZsom.SetValue("IV_DESCRIPTION", description);
                zicsCreateZsom.SetValue("IV_TICKET_CD", ticketCd);
                zicsCreateZsom.SetValue("IV_SERVICE_EMPLOYEE", serviceEmployee);

                foreach (DataRow items in servOrder)
                {

                    string product = items["product"].ToString().Trim();
                    //string quantity = items["quantity"].ToString();
                    string quantity = "2";
                    string partnerNoResponsible = GetEmployeeID(items["responsible"].ToString());
                    string contractId = items["Contract"].ToString().Trim();
                    string contractItemNo = items["contract_item"].ToString().Trim();

                    itLinesData.Append();
                    itLinesData.SetValue("PRODUCT", product);
                    itLinesData.SetValue("QUANTITY", quantity);
                    itLinesData.SetValue("CONTRACT_ID", contractId);
                    itLinesData.SetValue("CONTRACT_ITEM_NO", contractItemNo);
                    itLinesData.SetValue("PARTNER_NO", partnerNoResponsible);
                }

                string result = "";

                try
                {
                    console.WriteLine("Corriendo RFC de SAP: ZICS_CREATE_ZSOM");
                    zicsCreateZsom.Invoke(destCrm);

                    string zsomId = zicsCreateZsom.GetValue("ZSOM_ID").ToString();                          //se llena si la orden se creo, pero con errores
                    string ret = zicsCreateZsom.GetValue("RET").ToString();                                 //se llena si la FM no se pudo ejecutar
                    DataTable messages = sap.GetDataTableFromRFCTable(zicsCreateZsom.GetTable("MESSAGES")); //se llena con el ID de la orden generada

                    if (ret == "")
                    {
                        //Todo good
                        responseTableRow["Id de la Orden"] = zsomId;
                        responseTableRow["Descripción"] = description;
                        responseTableRow["Ticket"] = ticketCd;
                        foreach (DataRow message in messages.Rows)
                            result += message["ID"].ToString() + " - " + message["NUMBER"].ToString() + " - " + message["MESSAGE"].ToString() + Environment.NewLine;

                        responseTableRow["Resultado"] = result;
                    }
                    else
                    {
                        //Error en la FM
                        sendIcs = true;
                        responseTableRow["Descripción"] = description;
                        responseTableRow["Resultado"] = ret;
                    }
                }
                catch (Exception sapEx)
                {
                    //Error de SAP
                    sendIcs = true;
                    responseTableRow["Descripción"] = description;
                    responseTableRow["Resultado"] = sapEx.Message;
                }

                foreach (object item in responseTableRow.ItemArray)
                    console.WriteLine(item.ToString());

                responseTable.Rows.Add(responseTableRow);

                //log de base de datos

                log.LogDeCambios("", "", root.BDUserCreatedBy , root.BDProcess, responseTableRow["Id de la Orden"] + ": " + result, root.Subject);
                root.requestDetails += root.BDProcess + ": " + result;
            }

            if (sendIcs)
                //enviar email de repuesta de error
                mail.SendHTMLMail("Error al crear ordenes de servicio" + val.ConvertDataTableToHTML(responseTable), new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, attachments: new string[] { excelPath });
            else
                //enviar email de repuesta de éxito
                mail.SendHTMLMail("Se procesaron las ordenes de servicio:<br><br>" + val.ConvertDataTableToHTML(responseTable), new string[] { root.BDUserCreatedBy }, root.Subject, attachments: new string[] { excelPath });
        }
        private string GetEmployeeID(string email)
        {
            string userId = email.ToUpper().Trim().Replace("@GBM.NET", "");
            string partnerNoResponsible = val.GetEmployeeID(userId);

            if (!partnerNoResponsible.StartsWith("AA"))
                partnerNoResponsible = "AA" + partnerNoResponsible.PadLeft(8, '0');

            return partnerNoResponsible;
        }
        private bool IsValidMail(string email)
        {
            List<string> validMails = new List<string>();
            DataTable validMailsDt = crud.Select("SELECT `mail` FROM `authMails`", "create_zsom_db");

            foreach (DataRow fila in validMailsDt.Rows)
                validMails.Add(fila["mail"].ToString().ToUpper());

            if (validMails.Contains(email))
                return true;
            else
                return false;
        }
    }
}
