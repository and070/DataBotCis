using DataBotV5.Logical.Projects.CriticalTransactions;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;
using System.Globalization;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System;

namespace DataBotV5.Automation.ICS.CriticalTransactions
{
    /// <summary>
    /// Clase ICS Automation encargada de reportar las transacciones criticas. 
    /// </summary>
    class CriticalTransactions
    {
        readonly CriticalTransactionsLogical ct = new CriticalTransactionsLogical();
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly SapVariants sap = new SapVariants();
        readonly Settings sett = new Settings();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        public void Main()
        {
            const string mand = "ERP";
            string respFinal = "";

            if (!sap.CheckLogin(mand))
            {
                sap.BlockUser(mand, 1);

                //Transacciones HC
                string[] transactionsHc = { "PA20", "PA30" };
                string[] senderHc = { "MSalas@gbm.net", "jearaya@gbm.net", "POcampo@gbm.net", "GVillalobos@gbm.net" };
                ProcessCriticalTransacs(mand, transactionsHc,  senderHc);
                respFinal = respFinal + "\\n" + "Envío de correo con las transacciones criticas realizadas";

                //Transacciones DM
                string[] transactionsDm = { "XD02", "BP", "PFAL", "MM17", "MM01", "MM02", "XK02", "MASS", "COMMPR01", "SCC4", "STMS" };
                string[] senderDm = { "JEAraya@gbm.net" };
                ProcessCriticalTransacs(mand, transactionsDm, senderDm);
                respFinal = respFinal + "\\n" + "Auditoría transacciones DM: Se envió correo con las transacciones DM realizadas";

                //Transacciones FI
                string[] transactionsFi = { "IDCP", "ME23N", "MIGO", "VA01", "VA02", "VF01", "VF02", "VF04", "VFX3", "VL01N" };
                string[] senderFi = { "CCASTRO@gbm.net", "jearaya@gbm.net" };
                ProcessCriticalTransacs(mand, transactionsFi,senderFi);
                respFinal = respFinal + "\\n" + "Auditoría transacciones FI: Se envió correo con las transacciones FI realizadas";

                sap.BlockUser(mand, 0);
                root.requestDetails = respFinal;
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

        private void ProcessCriticalTransacs(string mand, string[] transactions, string[] senders)
        {
            byte[] image;
            List<string> tags = new List<string>();

            try
            {
                proc.KillProcess("saplogon", false);
                sap.LogSAP(mand);

                console.WriteLine("Corriendo SAP GUI: " + root.BDProcess);

                //como sacar screenshot de ventana //no se uso pero que quede de ejemplo
                //Process[] processes = Process.GetProcessesByName("saplogon");
                //Process saplogon = processes[0];
                //IntPtr saplogon_handle = saplogon.MainWindowHandle;
                //string tag = screen.CaptureWindowToHtmlTag(saplogon_handle) + i;
                //screen.CaptureWindowToFile(saplogon_handle, @"C:\" + DateTime.Now.ToFileTime() + ".jpg", ImageFormat.Jpeg);

                #region script

                ct.IniST03();

                foreach (string transaction in transactions)
                {
                    image = ct.GetTransaction(transaction, true, false).ImageResult;
                    if (image != null)
                        tags.Add("Transacción: " + transaction + "<br>" + ct.ByteToHtmlTag(image));
                }

                ct.CloseGui();

                #endregion

                #region Notificar

                string month = "", year = "";
                if (DateTime.Now.Month == 1)
                {
                    year = (DateTime.Now.Year - 1).ToString();
                    month = CultureInfo.GetCultureInfo("es-CR").DateTimeFormat.GetMonthName(12);
                }
                else
                {
                    month = CultureInfo.GetCultureInfo("es-CR").DateTimeFormat.GetMonthName(DateTime.Now.Month - 1);
                    year = DateTime.Now.Year.ToString();
                }

                month = CultureInfo.GetCultureInfo("es-CR").TextInfo.ToTitleCase(month.ToLower());

                string msg = "Buen día<br><br>Adjunto evidencia de ejecución de transacciones críticas durante el mes de  " + month + " del " + year;
                string tagsText = "";

                foreach (string tag in tags)
                    tagsText = tagsText + tag + "<br><br>";

                msg = msg + "<br><br>" + tagsText;

                mail.SendHTMLMail(msg, senders , "Ejecución Transacciones Criticas " + month + " del " + year);
                #endregion

                log.LogDeCambios("Notificacion", root.BDProcess, "Databot", root.BDProcess, "Transacciones Criticas", string.Join(",", transactions));
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
