using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Linq;
using System;

namespace DataBotV5.Automation.ICS.MRS
{
    /// <summary>
    /// Clase ICS Automation encargada del mantenimiento de MRS.
    /// </summary>
    class PerformanceMRS
    {
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();

        string mand = "ERP";

        public void Main()
        {
            ProcessUnplannedDemands();
            
        }
        private void ProcessUnplannedDemands()
        {
            string[] orgUnits = { "O50003350", "O50003580", "O50003579", "O70010729" };
            string[] results = (string[])orgUnits.Clone();
            string endDa = DateTime.Now.AddDays(-5).ToString("yyyy-MM-dd");
            string begDa = DateTime.Now.Year + "-01-01";
            string contend = "";

            try
            {
                RfcDestination destination = sap.GetDestRFC(mand);
                IRfcFunction func = destination.Repository.CreateFunction("ZRPA_MRS_PERFORMANCE_BACKLOG");

                for (int i = 0; i < orgUnits.Length; i++)
                {
                    func.SetValue("OBJID", orgUnits[i]);
                    func.SetValue("BEGDA", begDa); //primer dia del año
                    func.SetValue("ENDDA", endDa); //hoy - 5 dias

                    func.Invoke(destination);

                    results[i] = func.GetValue("RETURN").ToString();
                    contend = contend + orgUnits[i] + " (" + results[i] + ")<br>";
                    console.WriteLine(orgUnits[i] + " (" + results[i] + ")");
                }

                if (results.Contains("ERROR"))
                    mail.SendHTMLMail("Error en SAP: " + contend, new string[] { "internalcustomersrvs@gbm.net" }, "Error en Desempeño MRS");
                else
                    log.LogDeCambios("Procesar", root.BDProcess, "Automatico", "Procesar Unplanned Demands MRS", contend, "");

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(" > > > Error en el script: " + ex.Message);
                console.WriteLine(" > > > Respondiendo solicitud");
                mail.SendHTMLMail("Error : " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error en Desempeño MRS");
            }
        }
    }
}
