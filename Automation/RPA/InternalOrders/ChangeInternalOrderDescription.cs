using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.RPA.InternalOrders
{
    /// <summary>
    /// Clase RPA Automation encargada del cambio de descripción de ordenes internas.
    /// </summary>
    class ChangeInternalOrderDescription
    {
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        string respFinal = "";
        Log log = new Log();

        string mand = "ERP";


        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
            bool checkLogin = sap.CheckLogin(mand);
            if (!checkLogin)
            {
                //leer correo y descargar archivo
                if (mail.GetAttachmentEmail("Solicitudes Internal order", "Procesados", "Procesados Internal order"))
                {
                    sap.BlockUser(mand, 1);
                    DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    ProcessInternalOrder(excelDt);
                    sap.BlockUser(mand, 0);
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }

        }
        public void ProcessInternalOrder(DataTable excelDt)
        {
            string response = "", responFailure = "", guiResponse = "";

            try { excelDt.Columns.Add("Resultado"); } catch (DuplicateNameException) { }

            string validation = excelDt.Columns[1].ColumnName;

            if (validation.Substring(0, 1) == "x")//Plantilla correcta, continúe las validaciones
            {
                foreach (DataRow item in excelDt.Rows)
                {
                    //Validaciones
                    string order = item[0].ToString().Trim();
                    if (order.Length > 12)
                    {
                        response = response + order + ": " + "La orden es mayor a 12 caracteres";
                        break;
                    }

                    string desc = item[1].ToString().Trim();
                    desc = val.RemoveSpecialChars(desc, 1);
                    if (desc.Length > 40)
                    {
                        response = response + order + ": " + "La descripción es mayor a 40 caracteres";
                        break;
                    }

                    if (order != "")
                    {
                        console.WriteLine("Corriendo SAP GUI: " + root.BDProcess);

                        guiResponse = ChangeIntOrderDescScript(order, desc);

                        if (!guiResponse.Contains("FAILURE"))
                        {
                            //log de base de datos
                            console.WriteLine(order + ": " + guiResponse);
                            log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Cambio descripcion Orden Interna", order + ": " + guiResponse, root.Subject);
                            respFinal = respFinal + "\\n" + order + ": " + guiResponse;

                        }
                        else
                            responFailure = guiResponse;
                    }
                    item["Resultado"] = guiResponse;
                }

                console.WriteLine("Respondiendo solicitud");

                string htmlTable = val.ConvertDataTableToHTML(excelDt);

                if (responFailure.Contains("FAILURE"))//enviar email de repuesta de error
                    mail.SendHTMLMail(htmlTable, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                else
                    mail.SendHTMLMail(htmlTable, new string[] { root.BDUserCreatedBy }, root.Subject);

                sap.KillSAP();
            }
            else
                mail.SendHTMLMail("Utilizar la plantilla oficial de cambios", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

            root.requestDetails = respFinal;

        }

        private string ChangeIntOrderDescScript(string order, string description)
        {
            string guiResponse;

            try
            {
                proc.KillProcess("saplogon", false);
                sap.LogSAP(mand);

                SapVariants.frame.Iconify();
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nko02";
                SapVariants.frame.SendVKey(0);
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtCOAS-AUFNR")).Text = order;
                SapVariants.frame.SendVKey(0);

                if (((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString() != "")
                {
                    guiResponse = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                    if (guiResponse.Contains("does not exist"))
                        guiResponse = "La orden no existe";
                    else
                        guiResponse = "FAILURE";
                }
                else
                {
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtCOAS-KTEXT")).Text = description;
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                    if (SapVariants.session.ActiveWindow.Name == "wnd[1]")
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                    guiResponse = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();

                    if (guiResponse.Contains("Data was not changed"))
                        guiResponse = "La orden ya tiene esta descripción :" + description;
                    else if (guiResponse.Contains("has been changed"))
                        guiResponse = "La orden se ha actualizado :" + description;
                    else
                        guiResponse = "FAILURE";
                }
            }
            catch (Exception ex)
            {
                console.WriteLine("Error " + val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0));
                guiResponse = "FAILURE  ||" + ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString() + "   ||" + ex.Message;
            }

            return guiResponse;
        }
    }
}
