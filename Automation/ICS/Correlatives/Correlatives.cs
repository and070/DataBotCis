using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.Correlatives
{
    /// <summary>
    /// Clase ICS Automation encargada de la gestión de correlativos Dev y QAS.
    /// </summary>
    /// 
    class Correlatives
    {
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        PowerAutomate flow = new PowerAutomate();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();

        string respFinal = "";

        int[] mand = { 110, 260, 300 };
        string sapSystem = "ERP";

        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto 
            bool checkDev = sap.CheckLogin(sapSystem, mand[0]);
            bool checkQas = sap.CheckLogin(sapSystem, mand[1]);
            if (checkDev == false && checkQas == false)
            {
                if (mail.GetAttachmentEmail("Solicitudes correlativos", "Procesados", "Procesados Correlativos"))
                {
                    sap.BlockUser(sapSystem, 1, mand[0]);
                    sap.BlockUser(sapSystem, 1, mand[1]);
                    DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    CreateCorrelativos(excelDt);
                    sap.BlockUser(sapSystem, 0, mand[1]);
                    sap.BlockUser(sapSystem, 0, mand[0]);

                    console.WriteLine("Creando estadísticas...");

                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }

                }
            }
        }

        /// <summary> </summary>
        private void CreateCorrelativos(DataTable excel)
        {
            console.WriteLine("Cambiando correlativos");

            #region Excel
            //pass excel: correlativos

            string country = excel.Rows[0][0].ToString().Trim();
            string output = excel.Rows[0][1].ToString().Trim();
            string lot = excel.Rows[0][2].ToString().Split('-')[0].Trim();
            string book = excel.Rows[0][3].ToString().Trim();
            string from = excel.Rows[0][4].ToString().Trim();
            string to = excel.Rows[0][5].ToString().Trim();
            string validation = excel.Columns[0].ColumnName.Trim();

            #endregion

            if (validation.Substring(0, 1) == "x")
            {
                proc.KillProcess("saplogon", false);
                sap.LogSAP(sapSystem, mand[0]);

                int.TryParse(from, out int intFrom);
                intFrom--;
                string lastDocIssued = intFrom.ToString();
                string msgAppr = "Por favor su aprobación para ampliar correlativo país: " + country + /*" Output: " + output +*/ " Lot: " + lot + " Book: " + book + " Del rango " + from + " al " + to;
                string requestDesc = root.Subject;
                string lastOfficialDoc = to;

                string requestID = ExecuteIdlb(country, lot, lastOfficialDoc, lastDocIssued, requestDesc);

                //PASAR a QAS
                if (!requestID.Contains("ERROR"))
                {
                    console.WriteLine("  Enviar transporte " + requestID + " a QAS");
                    string res = Transport(requestID, "QAS");
                    if (res != "OK")
                    {
                        //Error al pasar a QA
                        mail.SendHTMLMail("Error al pasar a QA:<br>" + res, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, new string[] { "smarin@gbm.net" });
                    }
                    else
                    {
                        console.WriteLine("  Enviar solicitud de aprobación");
                        string json = "{\"pais\":\"" + country + "\",\"output\":\"" + output + "\",\"lot\":\"" + lot + "\",\"book\":\"" + book + "\",\"from\":\"" + from + "\",\"to\":\"" + to + "\",\"DEVK_num\":\"" + requestID + "\"}";
                        flow.SendApproval("Ampliar Correlativo", "BSolano@gbm.net", msgAppr,  root.BDUserCreatedBy , json, root.BDProcess);

                        //string resp=""
                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Enviar solicitud de aprobación", msgAppr + " Datos: " + json.Replace(@"\", ""), root.Subject);
                        respFinal = respFinal + "\\n" + "Crear Proveedor: " + "Enviar solicitud de aprobación - " + msgAppr + " Datos: " + json.Replace(@"\", "");

                    }
                }
                else //error
                    mail.SendHTMLMail(requestID, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, new string[] { "smarin@gbm.net" });
            }
            else
                mail.SendHTMLMail("Utilizar la plantilla oficial de cambios", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);


            root.requestDetails = respFinal;


            proc.KillProcess("saplogon", false);
        }
        /// <summary> </summary>
        private string ExecuteIdlb(string coCode, string lot, string lastOfficialDoc, string lastDocIssued, string requestDesc)
        {
            Dictionary<string, int> lista = new Dictionary<string, int>();
            string respuesta;

            try
            {
                #region Actualizar los rangos
                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).Maximize();
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nIDLB";
                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/sub:SAPLSVIX:0100/ctxtD0100_FIELD_TAB-LOWER_LIMIT[0,37]")).Text = coCode;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[25]")).Press();


                for (int i = 0; i <= 27; i++)
                {
                    try
                    {
                        string temp = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tblSAPLIDCNCUSTTCTRL_V_IDCN_LOMA/ctxtV_IDCN_LOMA-LOTNO[0," + i + "]")).Text;
                        if (temp != "____")
                            lista.Add(temp, i);
                    }
                    catch (Exception) { }
                }


                ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/tblSAPLIDCNCUSTTCTRL_V_IDCN_LOMA")).GetAbsoluteRow(lista[lot]).Selected = true;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[38]")).Press();

                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtV_IDCN_LOMA-INVTO")).Text = lastOfficialDoc;

                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell")).SelectItem("02", "Column1");
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell")).DoubleClickItem("02", "Column1");
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtV_IDCN_BOMA-LIINV")).Text = lastDocIssued;
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtV_IDCN_BOMA-INVTO")).Text = lastOfficialDoc;//aqui tambien???

                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();//guardar 
                #endregion

                #region El Transporte
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[8]")).Press();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[2]/usr/txtKO013-AS4TEXT")).Text = requestDesc;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[2]/tbar[0]/btn[0]")).Press(); // 0 es save // 12 es cancel
                string requestId = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtKO008-TRKORR")).Text;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press(); // 0 es save // 12 es cancel

                string devkTask = requestId.Replace("DEVK", "");
                int.TryParse(devkTask, out int devkTaskInt);
                devkTaskInt++;
                devkTask = "DEVK" + devkTaskInt;
                #endregion

                #region Liberar
                console.WriteLine("Liberando el transporte");

                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).Maximize();
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse37";
                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38L-NAME")).Text = "TR_RELEASE_REQUEST";
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,7]")).Text = devkTask;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,7]")).Text = requestId;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                System.Threading.Thread.Sleep(5000);
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[14]")).Press();

                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nex";
                ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0);
                #endregion

                respuesta = requestId;
            }
            catch (Exception err1)
            {
                try
                {
                    respuesta = "ERROR: " + ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                }
                catch (Exception ex)
                {
                    respuesta = "ERROR: " + ex.Message;
                }

                if (respuesta == "")
                    respuesta = "ERROR: " + err1.Message;
            }
            return respuesta;
        }
        /// <summary> </summary>
        private string Transport(string devkNum, string mandTarget)
        {
            int client;
            string res;
            try
            {
                switch (mandTarget)
                {
                    case "PRD":
                        client = 300;
                        sap.LogSAP(mandTarget, client);
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
                        client = 260;
                        sap.LogSAP(mandTarget, client);
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
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,14]")).Text = "X";
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,15]")).Text = "X";
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,16]")).Text = "X";
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,17]")).Text = "X";
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txt[34,20]")).Text = "X"; //CVERS
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
