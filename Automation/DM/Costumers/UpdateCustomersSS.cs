using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Automation.DM.Costumers
{
    public class UpdateCustomersSS
    {
        string system = "ERP";
        string mand = "QAS";
        int mandSap = 260;
        public void Main()
        {
            ConsoleFormat console = new ConsoleFormat();
            try
            {
                DataTable resp = new DataTable();
                resp.Columns.Add("cliente");
                resp.Columns.Add("sql");
                resp.Columns.Add("resp");
                CRUD crud = new CRUD();
                MailInteraction mailCc = new MailInteraction();
                SapVariants sap = new SapVariants();
                Rooting root = new Rooting();
                MsExcel ms = new MsExcel();
                string respFinal = "";
                DataTable dtRows = crud.Select("SELECT COUNT(*) FROM clients", "databot_db", mand);
                int rows = int.Parse(dtRows.Rows[0]["COUNT(*)"].ToString());
                int halfRow = rows / 4;
                bool valLines = true;
                DataTable dt1 = crud.Select($"SELECT * FROM clients WHERE id < {halfRow}", "databot_db", mand);
                DataTable dt2 = crud.Select($"SELECT * FROM clients WHERE id >= {halfRow} AND id < {halfRow * 2}", "databot_db", mand);
                DataTable dt3 = crud.Select($"SELECT * FROM clients WHERE id >= {halfRow * 2} AND id < {halfRow * 3}", "databot_db", mand);
                DataTable dt4 = crud.Select($"SELECT * FROM clients WHERE id >= {halfRow * 3}  AND id < {halfRow * 4}", "databot_db", mand);

                DataTable dt = dt1.Clone();
                dt.Merge(dt1);
                dt.Merge(dt2);
                dt.Merge(dt3);
                dt.Merge(dt4);
                foreach (DataRow item in dt.Rows)
                {
                    DataRow rRow = resp.Rows.Add();
                    string id = item["idClient"].ToString();
                    try
                    {
                        Dictionary<string, string> parameters = new Dictionary<string, string>
                        {
                            ["BP"] = id,
                        };
                        console.WriteLine($"Run: {id}");
                        IRfcFunction fm = sap.ExecuteRFC(system, "ZDM_READ_BP", parameters, mandSap);

                        string name = fm.GetValue("NOMBRE").ToString();
                        string add = fm.GetValue("ADDRESS").ToString() + " " + fm.GetValue("COMPLADDRESS").ToString();
                        string salesRep = fm.GetValue("SALESREP").ToString().Replace("AA", "");
                        string phone = fm.GetValue("PHONE").ToString();
                        string mail = fm.GetValue("EMAIL").ToString();
                        string terro = fm.GetValue("CG1").ToString();

                        string upQuery = $@"UPDATE clients SET
                                            name = '{name}',
                                            {((salesRep != "") ? $"accountManagerId = '{salesRep}'," : "")}
                                            {((salesRep != "") ? $"accountManagerUser = (SELECT MIS.digital_sign.user from MIS.digital_sign WHERE MIS.digital_sign.UserID = '{salesRep}')," : "")}
                                            address = '{add}', 
                                            telephone = '{phone}', 
                                            email = '{mail}', 
                                            {((salesRep != "") ? $"employeeResponsible = (SELECT MIS.digital_sign.id from MIS.digital_sign WHERE MIS.digital_sign.UserID = '{salesRep}')," : "")}
                                            territory = (SELECT databot_db.valueTeam.id from databot_db.valueTeam WHERE databot_db.valueTeam.code = '{terro.Substring(0, 3)}'),
                                            updatedAt = CURRENT_TIMESTAMP,
                                            updatedBy = 'Databot'
                                            WHERE id = {item["id"]}";
                        bool upda = crud.Update(upQuery, "databot_db", mand);
                        if (!upda)
                        {
                            valLines = false;

                        }
                        rRow["cliente"] = item["idClient"].ToString();
                        rRow["sql"] = upQuery;
                        rRow["resp"] = upda;

                    }
                    catch (Exception ex)
                    {
                        rRow["cliente"] = id;
                        rRow["sql"] = "";
                        rRow["resp"] = ex.Message;
                    }
                }
                resp.AcceptChanges();
                string dtResponseRoute = root.FilesDownloadPath + "\\" + "respClients.xlsx";
                ms.CreateExcel(resp, "sheet1", dtResponseRoute);
                //send email
                console.WriteLine("Send Email...");
                string[] attachments = new string[] { dtResponseRoute };
                string htmlEmail = Properties.Resources.emailtemplate1.Replace("{subject}", "Actualización de Clientes S&S").Replace("{cuerpo}", "Adjunto encontrará un excel con los datos actualizados").Replace("{contenido}", "");
                mailCc.SendHTMLMail(htmlEmail, new string[] { "dmeza@gbm.net" }, $"Actualización de clientes en databot_db {DateTime.Now.ToString("dd/MM/yyyy")}", null, attachments);

                root.requestDetails = respFinal;
                root.BDUserCreatedBy = "DMEZA";

                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
            }

        }
    }
}
