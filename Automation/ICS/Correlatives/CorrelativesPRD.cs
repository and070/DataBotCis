using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.ICS.Correlatives
{
    /// <summary>
    /// Clase ICS Automation encargada de la gestión de correlativos PRD.
    /// </summary>
    /// 
    class CorrelativesPRD
    {
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        Log log = new Log();

        int[] mand = { 110, 260, 300 };
        string sapSystem = "ERP";
        string respFinal = "";


        /// <summary> </summary>
        public void Main()
        {
            if (!sap.CheckLogin(sapSystem, mand[1]))
            {
                sap.BlockUser(sapSystem, 1, mand[1]);

                //leer aprobaciones pendientes.
                Dictionary<string, string> resAppr = mail.GetApprovalRequests("Correlatives");

                try
                {
                    if (resAppr.Count > 0)
                    {
                        //Tomar datos generales
                        JObject originalJson = JObject.Parse(resAppr["OriginalJson"]);
                        string requestedBy = originalJson["reported_by"].Value<string>().ToUpper();
                        string approverEmail = originalJson["approver"].Value<string>().ToUpper();

                        //Tomar resultado de la aprobación
                        JObject responseApprJson = JObject.Parse(resAppr["ResponseJson"]);
                        string status = responseApprJson["RESPONSE"].Value<string>().ToUpper();
                        string comment = responseApprJson["COMMENTS"].Value<string>().ToUpper();

                        //Tomar datos específicos
                        JObject specificJson = JObject.Parse(originalJson["specific_data"].ToString());
                        string devkNum = specificJson["DEVK_num"].Value<string>().ToUpper();
                        string pais = specificJson["pais"].Value<string>().ToUpper();
                        string output = specificJson["output"].Value<string>().ToUpper();
                        string lot = specificJson["lot"].Value<string>().ToUpper();
                        string book = specificJson["book"].Value<string>().ToUpper();
                        string from = specificJson["from"].Value<string>().ToUpper();
                        string to = specificJson["to"].Value<string>().ToUpper();


                        if (status == "APPROVE")
                        {
                            //PASAR A PRD EL TRANSPORTE devkNum

                            string res = Transport(devkNum, "PRD");

                            if (res == "OK")
                            {
                                mail.SendHTMLMail("Se actualizaron los correlativos rango " + from + " al " + to, new string[] { requestedBy }, "CORRELATIVO APROBADO");
                                log.LogDeCambios("Modificacion", "Correlativos", requestedBy, "Aprobado", originalJson["specific_data"].ToString(), comment);
                                respFinal = respFinal + "\\n" + "Correlativos solicitado por " + requestedBy + " aprobado: " + originalJson["specific_data"].ToString() + comment;

                            }
                            else
                                mail.SendHTMLMail("Error al aplicar el transporte correlativos: " + devkNum, new string[] { requestedBy }, root.Subject);
                        }
                        else if (status == "REJECT")
                        {
                            mail.SendHTMLMail("SE RECHAZO EL TRANSPORTE " + devkNum + " POR EL SIGUIENTE COMENTARIO:<BR>" + comment, new string[] { "SMARIN@GBM.NET" }, "CORRELATIVO RECHAZADO");/////Deberia enviar algun correo?
                            log.LogDeCambios("Modificacion", "Correlativos", requestedBy, "Rechazado", originalJson["specific_data"].ToString(), comment);
                            respFinal = respFinal + "\\n" + "Correlativos solicitado por " + requestedBy + " Rechazado: " + originalJson["specific_data"].ToString() + comment;

                        }
                        root.BDUserCreatedBy = requestedBy;
                        root.requestDetails = respFinal;


                        console.WriteLine("Creando estadísticas...");
                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                    }

                }
                catch (Exception ex)
                {
                    mail.SendHTMLMail("Error correlativos<br>" + ex, new string[] { "smarin@gbm.net" }, root.Subject, new string[] { "smarin@gbm.net" });
                }

                proc.KillProcess("saplogon", false);


                sap.BlockUser(sapSystem, 0, mand[1]);
            }
        }

        private string Transport(string devkNum, string mandTarget)
        {
            string res;
            int client;
            try
            {
                switch (mandTarget)
                {
                    case "PRD":
                        client = mand[2];
                        sap.LogSAP(sapSystem, client);
                        //add transport to queue
                        ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).Maximize();
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse37";
                        ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38L-NAME")).Text = "TMS_MGR_FORWARD_TR_REQUEST";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,7]")).Text = devkNum;
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,8]")).Text = mandTarget;
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,10]")).Text = client.ToString();
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,15]")).Text = "X";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,16]")).Text = "X";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                        break;
                    case "QAS":
                        client = mand[1];
                        sap.LogSAP(sapSystem, client);
                        break;
                    default:
                        client = 0;
                        break;
                }

                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).Maximize();
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse37";
                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38L-NAME")).Text = "TMS_MGR_IMPORT_TR_REQUEST";
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,7]")).Text = mandTarget;
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,8]")).Text = "DOMAIN_DEV";
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,9]")).Text = devkNum;
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,10]")).Text = client.ToString();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,14]")).Text = "X"; //Originality
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,15]")).Text = "X"; //Repairs
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,16]")).Text = "X"; //Transtype
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,17]")).Text = "X"; //Tabletype
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,20]")).Text = "X"; //Cvers
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();//pasar
                res = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[37,45]")).Text.Trim();

                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nex";
                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);

                return res;
            }
            catch (Exception)
            {
                res = "Error en el Script de transporte";
                return res;
            }
        }
    }
}
