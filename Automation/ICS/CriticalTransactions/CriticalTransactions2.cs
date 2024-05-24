using DataBotV5.Logical.Projects.CriticalTransactions;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.CriticalTransactions
{
    internal class CriticalTransactions2
    {
        readonly CriticalTransactionsLogical ct = new CriticalTransactionsLogical();
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        readonly Settings sett = new Settings();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        const string mand = "ERP";

        public void Main()
        {
            if (!sap.CheckLogin(mand))
            {
                sap.BlockUser(mand, 1);
                ProcessCriticalTransacs(mand);
                sap.BlockUser(mand, 0);

                root.requestDetails = "\\n" + "Checkeo de las transacciones criticas diarias";
                root.BDUserCreatedBy = "internalcustomersrvs";

                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
            else
            {
                sett.setPlannerAgain();
            }
        }
        private void ProcessCriticalTransacs(string mand)
        {
            string[] transactions = { "SU01", "SU10", "PFCG", "SU02", "SU1", "SU2", "SU3", "SARA", "SM02", "SCC4", "STMS" };
            string[] allowedUser = { "HLHERRERA", "CLGARCIA", "KASANCHEZ", "JEARAYA", "CCASCANTE", "ROFERNANDEZ", "SMARIN", "KPADILLA", "ATRIGUEROS" };
            string[] senders = { "internalcustomersrvs@gbm.net" };

            string body = "";

            try
            {
                proc.KillProcess("saplogon", false);
                sap.LogSAP(mand);

                console.WriteLine("Corriendo SAP GUI: " + root.BDProcess);

                #region script

                ct.IniST03("5");

                foreach (string transaction in transactions)
                {
                    DataTable res = ct.GetTransaction(transaction, false, true).DtResult;

                    if (res.Rows.Count > 0)
                    {
                        //borrar los usuarios de excepción
                        string colunm0 = res.Columns[0].ColumnName;
                        DataRow[] result = res.Select("[" + colunm0 + "] IN ('" + string.Join("', '", allowedUser) + "')");

                        foreach (DataRow row in result)
                            row.Delete();

                        res.AcceptChanges();

                        if (res.Rows.Count > 0)
                        {
                            //si la tabla tiene fila ir construyendo la respuesta
                            body += transaction + ": <br><br>" + val.ConvertDataTableToHTML(res) + "<br><br>";
                        }
                    }
                }

                ct.CloseGui();

                #endregion

                if (body != "")
                {
                    #region Notificar

                    string msg = "Buen día<br><br>Adjunto evidencia de ejecución de transacciones críticas durante del día:  " + DateTime.Today.ToString("dd-MM-yyyy");
                    msg += "<br><br>" + body;

                    mail.SendHTMLMail(msg,  senders , "Ejecución Transacciones Criticas " + DateTime.Today.ToString("dd-MM-yyyy"));
                    #endregion
                    log.LogDeCambios("Notificacion", root.BDProcess, "Databot", root.BDProcess, "Transacciones Criticas2", string.Join(",", transactions));
                }
            }
            catch (Exception ex)
            {
                console.WriteLine("Error en el script: " + ex.Message);
                console.WriteLine("Respondiendo solicitud");
                mail.SendHTMLMail("Error: " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, root.BDProcess);
            }
            sap.KillSAP();
        }
    }
}
