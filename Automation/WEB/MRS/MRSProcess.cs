using DataBotV5.Data.Database;

using DataBotV5.Data.SAP;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.EasyLDR;
using DataBotV5.Logical.Webex;
using System;
using System.Data;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;

namespace DataBotV5.Automation.WEB.MRS

{
    /// <summary>
    /// Clase WEB Automation encargada del Proces MRS.
    /// </summary>
    class MRSProcess
    {
        public void Main()
        {
            CRUD crud = new CRUD();
            //DataTable dt = crud.Select("Databot", "SELECT * FROM test_oauth_ms WHERE status = 1", "licitaciones_cr");
            //if (dt.Rows.Count > 0)
            //{

            MicrosoftTeams ms = new MicrosoftTeams();
            MailInteraction mail = new MailInteraction();
            WebexTeams wb = new WebexTeams();
            Rooting root = new Rooting();
            try
            {
                mail.GetAttachmentEmail("Solicitudes Test", "Procesados", "Procesados Tests");
                mail.SendHTMLMail("Hola Mundo", new string[] { "dmeza@gbm.net" }, "subject de prueba exchange", new string[] { "epiedra@gbm.net" }, new string[] { root.FilesDownloadPath + "\\test.xlsx" });
                ms.EnviarChatMS();
                //crud.Update("Databot", $"UPDATE test_oauth_ms SET status = '0' response = '1' WHERE id = {dt.Rows[0]["id"]}", "licitaciones_cr");

            }
            catch (Exception ex)
            {
                //crud.Update("Databot", $"UPDATE test_oauth_ms SET status = '0', response = '0' WHERE id = {dt.Rows[0]["id"]}", "licitaciones_cr");
                //mail.SendHTMLMail("Error al intentar enviar un correo con OAuth, debido a: " + ex.Message, new string[] { "dmeza@gbm.net" }, "Error al enviar correo Databot Autentificacion Moderna", 2, null, null, 0);
                wb.SendNotification("dvillalobos@gbm.net", "Error OAuth", "Error al intentar enviar un correo con OAuth, debido a: " + ex.Message);
                wb.SendNotification("dmeza@gbm.net", "Error OAuth", "Error al intentar enviar un correo con OAuth, debido a: " + ex.Message);
            }
            using (Stats stats = new Stats())
            {
                stats.CreateStat();
            }
            //}
            ////Job de MRS todos los dias a las 10PM
            //if (Rank())
            //{
            //    ProcessInteraction proc = new ProcessInteraction();
            //    SapVariants sap = new SapVariants(); 
            //    try
            //    {
            //        //si se encuentra en rango programe el job y corralo
            //        sap.LogSAP("300");
            //        EasyLDR con = new EasyLDR();
            //        con.MrsJob();
            //        sap.KillSAP();
            //    }
            //    catch (Exception ex)
            //    {
            //        //proc.MatarProceso("saplogon",false);
            //        sap.KillSAP();
            //        //System.Threading.Thread.Sleep(200000);
            //        string n_ar = DateTime.Now.ToString("yyyyMMddHHmmss");
            //        PushNotificacion push_n = new PushNotificacion();
            //        push_n.PushNotification(n_ar, new string[] { "dmeza@gbm.net" }, "El job de MRS ha fallado listado de errores: " + ex.Message);

            //    }
            //}
        }
        private bool Rank()
        {
            bool r = false;
            DateTime t1 = DateTime.Now;
            DateTime t2 = Convert.ToDateTime("10:00:00 PM");
            DateTime t3 = Convert.ToDateTime("10:04:00 PM");
            if (t1 > t2 && t1 < t3)
            {
                r = true;
            }
            return r;
        }

    }
}
