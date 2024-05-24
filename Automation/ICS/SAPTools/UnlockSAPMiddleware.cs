using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System;

namespace DataBotV5.Automation.ICS.SAPTools
{
    /// <summary>
    ///Clase ICS Automation encargada del desbloqueo del middleware de SAP.
    /// </summary>
    class UnlockSAPMiddleware
    {
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();
        Settings sett = new Settings();

        string mand = "ERP";
        string respFinal = "";

        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
         
            if (!sap.CheckLogin(mand))
            {
                sap.BlockUser(mand, 1);
                ProcessMID();
    
                sap.BlockUser(mand, 0);
            }
            else
            {
                sett.setPlannerAgain();
            }
        }
        private void ProcessMID()
        {
            proc.KillProcess("saplogon", false);
            sap.LogSAP(mand.ToString());

            console.WriteLine("Corriendo SAP GUI: " + root.BDProcess);
            try
            {
                #region Script

                SapVariants.frame.Maximize();
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nsmq1";
                SapVariants.frame.SendVKey(0);
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtQNAME")).Text = "R3AD_CS_EQU_*";
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                string list = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[36,2]")).Text.Trim();
                int queues = int.Parse(((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[36,3]")).Text.Trim());
                string[] queueNames = new string[queues];

                if (list != "0")
                {
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[5]")).Press(); //select all
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[7]")).Press(); //displ
                    int fila = 3;
                    for (int i = 0; i < queues; i++)
                    {
                        queueNames[i] = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[5," + fila + "]")).Text.Trim();
                        fila += 2;
                    }

                    for (int i = 0; i < queues; i++)
                    {
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[5]")).Press(); //unlock
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]")).Text = queueNames[i];
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press(); //ok
                        if (SapVariants.session.ActiveWindow.Name == "wnd[1]")
                        {
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnSPOP-OPTION1")).Press();
                        }
                    }
                    console.WriteLine(" > > > Se desbloquearon: " + list + " equipos");
                    log.LogDeCambios("Modificacion", root.BDProcess, "Datos Maestros", "Desbloqueo Middleware", "Se desbloquearon: " + list + " equipos", "");
                    respFinal = respFinal + "\\n"+ "Se desbloquearon: " + list + " equipos Middleware";

                    root.BDUserCreatedBy = "internalcustomersrvs";
                    root.requestDetails = respFinal;

                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
                else
                    console.WriteLine(" > > > No hay equipos bloqueados");
                #endregion
            }
            catch (Exception ex)
            {
                console.WriteLine(" > > > Error en el script: " + ex.Message);
                console.WriteLine(" > > > Respondiendo solicitud");
                mail.SendHTMLMail("Error en el script: " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Desbloquear equipos del Middleware");
            }
            sap.KillSAP();
 

        }
    }
}
