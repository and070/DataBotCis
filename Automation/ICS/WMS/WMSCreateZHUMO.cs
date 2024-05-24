using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System;

namespace DataBotV5.Automation.ICS.WMS
{
    /// <summary>
    /// Clase ICS Automation encargada de la creación de HUMO en WMS de ICS.
    /// </summary>
    class WMSCreateZHUMO
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();
        Settings sett = new Settings();
        string mand = "ERP";
        public void Main()
        {
            bool validateLines = true;
            string response = "";
            
            string respFinal = "";


            //revisa si el usuario RPAUSER esta abierto
            if (!sap.CheckLogin(mand))
            {
                sap.BlockUser(mand, 1);
                string sapMsg = "";
                int a = 0;
                while (sapMsg == "")
                {
                    console.WriteLine("Procesando... " + a);
                    sap.LogSAP(mand);
                    try
                    {
                        SapVariants.frame.Iconify();
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nzhumo";
                        SapVariants.frame.SendVKey(0);
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                        try { sapMsg = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString(); } catch (Exception) { }

                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
                        SapVariants.frame.SendVKey(0);
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "ZUPDATE_ZLX02";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                    }
                    catch (Exception) { }

                    sap.KillSAP();

                    if (sapMsg == "")
                    {
                        string date = DateTime.Now.ToString("dd.MM.yyyy");

                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            #region Parámetros de SAP
                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["TODAY"] = date
                            };

                            IRfcFunction func = sap.ExecuteRFC(mand, "ZDM_HUMO", parameters);
                            #endregion

                            #region Procesar Salidas del FM
                            response = date + ": " + func.GetValue("RESPUESTA").ToString() + "<br>";

                            //log de base de datos
                            console.WriteLine(date + ": " + func.GetValue("RESPUESTA").ToString());
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            console.WriteLine(" Finalizando proceso " + ex.ToString());
                            response = response + date + ": " + ex.ToString() + "<br>";
                            validateLines = false;
                        }

                        if (response.Contains("No hay errores"))
                            sapMsg = "Se cargo la ZHUMO con éxito";
                        else if (response.Contains("Error:"))
                            sapMsg = "Se cargo la ZHUMO, sin embargo dio error en Z_SALDOS" + "<br>" + "<br>" + response;
                        else if (response.Contains("No se encontró errores por saldos"))
                            sapMsg = "Se cargo la ZHUMO, sin embargo algún documento dio error, pero no por saldos";
                        else
                            sapMsg = "";

                        if (validateLines == false)
                            break;
                    }

                    a++;
                    if (a >= 3)
                        break;
                }

                if (sapMsg == "")
                    sapMsg = "Se cargo la ZHUMO, sin embargo algún documento dio error, verificar la tabla ZWMAF_TSTOCK";

                //enviar email a datos maestros
                mail.SendHTMLMail(sapMsg, new string[] { "internalcustomersrvs@gbm.net" }, "Cargar ZHUMO en SAP",  new string[] { "hlherrera@gbm.net" });
                System.Threading.Thread.Sleep(20000);
                log.LogDeCambios("Creacion", root.BDProcess, "internalcustomersrvs@gbm.net", "Crear Humos", sapMsg, "ZHUMO");
                respFinal = respFinal + "\\n" + "Crear Humos" + sapMsg;

                root.requestDetails = respFinal;
                root.BDUserCreatedBy = "internalcustomersrvs";

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }

                sap.BlockUser(mand, 0);
            }
            else
            {
                sett.setPlannerAgain();
            }
        }
    }
}
