using System;
using SAP.Middleware.Connector;
using System.Data;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.SAP;

namespace DataBotV5.Automation.ICS.BusinessPartners
{
    /// <summary>
    /// Clase ICS Automation encargada de modificar las cuentas bancarias BP.
    /// </summary>
    class BankAccounts
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();

        string mandErp = "ERP";
        int mand = 260;

        string respFinal = "";


        public void Main()
        {
            if (mail.GetAttachmentEmail("Solicitudes Cuentas Bancarias", "Procesados", "Procesados Cuentas Bancarias"))
            {
                console.WriteLine("Procesando...");
                UpdateAccounts(root.FilesDownloadPath + "\\" + root.ExcelFile);
                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        private void UpdateAccounts(string path)
        {
            string partner, cty, bank, account, iban, currency, holder, action, response;
            string[] bankId, cc = { "hlherrera@gbm.net" };

            DataTable respTemplate = new DataTable();
            respTemplate.Columns.Add("PARTNER");
            respTemplate.Columns.Add("PAIS");
            respTemplate.Columns.Add("BANCO");
            respTemplate.Columns.Add("CUENTA");
            respTemplate.Columns.Add("MONEDA");
            respTemplate.Columns.Add("IBAN");
            respTemplate.Columns.Add("HOLDER");
            respTemplate.Columns.Add("ACCION");
            respTemplate.Columns.Add("RESPUESTA");

            #region abrir excel

            DataTable ws2 = excel.GetExcel(path);
            int rows = ws2.Rows.Count;

            #endregion

            if (ws2.Columns[0].ToString() != "PARTNER" || ws2.Columns[1].ToString() != "PAÍS" || ws2.Columns[2].ToString() != "BANCO" || ws2.Columns[3].ToString().ToUpper() != "CUENTA" || rows > 31)
            {
                if (rows > 31)
                {
                    string returnMsg = "Para la actualización masiva de cuentas favor enviar la gestión directamente a Internal Customer Services";
                    console.WriteLine(returnMsg);
                    console.WriteLine("Devolviendo solicitud");
                    mail.SendHTMLMail(returnMsg, new string[] { root.BDUserCreatedBy }, root.Subject, cc);
                }
                else
                {
                    console.WriteLine("Devolviendo Solicitud");
                    response = "Favor utilizar la plantilla oficial de Internal Customer Services, sin modificar." + "<br>";
                    mail.SendHTMLMail(response, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
                }
            }
            else
            {
                foreach (DataRow row in ws2.Rows)
                {
                    partner = row[0].ToString().ToUpper();
                    cty = row[1].ToString().ToUpper();
                    bank = row[2].ToString().ToUpper();
                    account = row[3].ToString().ToUpper();
                    iban = row[4].ToString().ToUpper();
                    currency = row[5].ToString().ToUpper();
                    holder = row[6].ToString().ToUpper();
                    action = row[7].ToString().ToUpper();
                    bankId = bank.Split(' ');
                    bank = bankId[0].ToUpper();

                    try
                    {
                        if ((partner.Substring(0, 1).ToUpper() == "1" && partner.Length != 10) || (partner.Substring(1, 1).ToUpper() == "1" && partner.Length == 9))
                            partner = partner.PadLeft(10, '0');
                        else
                        {
                            partner = partner.PadLeft(8, '0');
                            partner = "AA" + partner;
                        }

                        #region SAP
                        Dictionary<string, string> parameters = new Dictionary<string, string>
                        {
                            ["VENTA"] = partner,
                            ["PAIS"] = cty,
                            ["COD_BANCO"] = bank,
                            ["CUENTA_BCO"] = account,
                            ["MONEDA"] = currency,
                            ["IBAN"] = iban,
                            ["CTA_HOLDER"] = holder,
                            ["ACCION"] = action
                        };

                        IRfcFunction zicsChangeCtaBanco = sap.ExecuteRFC(mandErp, "ZICS_CHANGE_CTA_BANCO", parameters, mand);
                        #endregion

                        response = zicsChangeCtaBanco.GetValue("MENSAJE").ToString();

                        if (partner != "")
                        {
                            DataRow respTemplRow = respTemplate.NewRow();
                            respTemplRow["PARTNER"] = partner;
                            respTemplRow["PAIS"] = cty;
                            respTemplRow["BANCO"] = bank;
                            respTemplRow["CUENTA"] = account;
                            respTemplRow["MONEDA"] = currency;
                            respTemplRow["IBAN"] = iban;
                            respTemplRow["HOLDER"] = holder;
                            respTemplRow["ACCION"] = action;
                            respTemplRow["RESPUESTA"] = response;
                            respTemplate.Rows.Add(respTemplRow);

                            log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, partner, response, cty + " / " + bank + " / " + account + " / " + currency + " / " + iban + " / " + holder + " / " + action);
                            respFinal = respFinal + "\\n" + partner + " / " + cty + " / " + bank + " / " + account + " / " + currency + " / " + iban + " / " + holder + " / " + action + ": " + response;

                        }
                    }
                    catch (Exception) { continue; }
                }
            }
            console.WriteLine("Respondiendo solicitud");

            mail.SendHTMLMail("Se adjuntan los resultados de la actualización: <br>" + val.ConvertDataTableToHTML(respTemplate), new string[] { root.BDUserCreatedBy }, root.Subject, cc);
            root.requestDetails = respFinal;

        }
    }
}
