using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Automation.MASS.RevenueRecognition
{
    class CollectiveProcessing
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        CRUD crud = new CRUD();
        Log log = new Log();
        string mandante = "ERP";

        string respFinal = "";

        /// <summary>
        /// Disponer de un RPA que permita la carga de información en SAP/ERP en la transacción VF44 de las lineas a reconocer de forma automatica, según la información del excel enviada por cada país.
        /// </summary>
        public void Main()
        {
            //utilizar en caso de que el robot utilice SAP Logon GUI
            if (!sap.CheckLogin(mandante))
            {
                //Leer email y extraer adjunto
                if (mail.GetAttachmentEmail("Solicitudes Revenue", "Procesados", "Procesados Revenue"))
                {
                    //convertir el excel en un DataTable para ser procesada
                    //root.ExcelFile = "RECONOCIMIENTO.xlsx";
                    DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    console.WriteLine("Processing...");
                    //utilizar en caso de que el robot utilice SAP Logon GUI
                    sap.BlockUser(mandante, 1);
                    ///Procesar------------------
                    PutRecognition(excelDt);
                    ///--------------------------
                    //utilizar en caso de que el robot utilice SAP Logon GUI
                    sap.BlockUser(mandante, 0);

                    //crear estadisticas
                    console.WriteLine("Creando estadísticas...");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }
        /// <summary>
        /// Cargar el item de cada sales order en la VF44 para confirmarlo
        /// </summary>
        /// <param name="salesOrderInfo">el excel que envía el país por email outlook</param>
        private void PutRecognition(DataTable salesOrderInfo)
        {
            #region private variables

            //columnas necesarias para cargar en SAP
            string[] readColumns = {
                "Sales Document",
                "Sales Document Item",
                "Company Code",
                "Post. year & post. period",
                "Acc. Deferred Revenues/Costs",
            };
            //columnas del excel
            DataColumnCollection columns = salesOrderInfo.Columns;
            //contador para verificar el excel
            int contTrue = 0;
            //variable de respuesta en caso de que se necesite
            string response = "";
            //DataTable furuto excel de respuesta
            DataTable dtResponse = new DataTable();
            //variable nombre de la hoja de resultados
            string dtResponseSheetName = "Revenue Recognition Results";
            //variable nombre del libro de resultados + extension
            string dtResponseBookName = "Results " + root.ExcelFile;
            //ruta + nombre donde se guardará el excel de resultado
            string dtResponseRoute = root.FilesDownloadPath + "\\" + dtResponseBookName;
            //PLantilla en html para el envío de email
            string htmlEmail = Properties.Resources.emailtemplate1;
            //variable titulo del cuerpo del correo
            string htmlSubject = root.Subject + " " + DateTime.Now.ToString("yyyyMMddHHmmss");
            //variable contenido del correo: texto, cuadros, tablas, imagenes, etc
            string htmlContents = "";
            //se agregan los parametros anterior al html del cuerpo del email de respuesta
            htmlContents = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", response).Replace("{contenido}", "");
            //variable remitente del email de respuesta
            string sender = root.BDUserCreatedBy;
            //variable copias del email de respuesta
            string[] cc = new string[] { "DMEZA@gbm.net" };
            //variable ruta de adjunto
            string[] attachments = new string[] { dtResponseRoute };
            //variable cambio para log
            string logText = "";
            //variable del body
            string body = root.Email_Body;
            //fecha de posteo
            string postDate = body.Split(new string[] { "Posting Date: " }, StringSplitOptions.None)[1].Substring(0, 10).Replace("\r\n", "").Trim();
            #endregion
            #region excel verification and Molded of the Results Excel
            console.WriteLine("Checking...");
            string columnsReadString = String.Join("<br>", readColumns.ToArray());
            string columnsResponseString = "";
            foreach (string columnName in readColumns)
            {
                //verifica si la columna esta en el excel
                if (columns.Contains(columnName))
                {
                    contTrue++;
                    //crea la columna en el excel de resultados
                    dtResponse.Columns.Add(columnName);
                    columnsResponseString = columnsResponseString + columnName + "<br>";
                }

            }
            //agrega la columna de resultado en el Excel
            dtResponse.Columns.Add("Response");
            //si es diferente a 4 significa que no encontro una de las columnas necesarias para cargar la reconocimiento
            if (contTrue != readColumns.Length)
            {

                response = $"No se encontró las columnas necesarias para procesar la orden. Por favor verifique e intente nuevamente<br><br>{columnsReadString}<br><br>Columnas encontradas:<br>{columnsResponseString}";
                htmlContents = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", response).Replace("{contenido}", "");
                mail.SendHTMLMail(htmlContents, new string[] { sender }, "Error: " + root.Subject, cc, new string[] { root.FilesDownloadPath + "\\" + root.ExcelFile });
                return;
            }
            #endregion
            #region loop each excel row
            console.WriteLine("Foreach Excel row...");
            sap.LogSAP(mandante.ToString());
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nZVF44";
            SapVariants.frame.SendVKey(0);
            foreach (DataRow rRow in salesOrderInfo.Rows)
            {
                #region robot Process
                string salesDocument = "";
                string item = "";
                string Cocode = "";
                string postYear = "";
                string account = "";
                response = "";
                try
                {
                    salesDocument = rRow["Sales Document"].ToString();
                    item = rRow["Sales Document Item"].ToString();
                    Cocode = rRow["Company Code"].ToString();
                    postYear = rRow["Post. year & post. period"].ToString();
                    account = rRow["Acc. Deferred Revenues/Costs"].ToString();
                    if (!string.IsNullOrWhiteSpace(salesDocument))
                    {
                        response = zvff4Recognition(salesDocument, item, Cocode, postYear, account, postDate);

                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.ToString());
                    response = ex.ToString();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nZVF44";
                    SapVariants.frame.SendVKey(0);
                }
                if (!string.IsNullOrWhiteSpace(salesDocument))
                {

                    DataRow rRowResult = dtResponse.Rows.Add();
                    rRowResult["Sales Document"] = salesDocument;
                    rRowResult["Sales Document Item"] = item;
                    rRowResult["Company Code"] = Cocode;
                    rRowResult["Post. year & post. period"] = postYear;
                    rRowResult["Acc. Deferred Revenues/Costs"] = account;
                    rRowResult["Response"] = response;

                }
                //log de cambios
                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, $"{salesDocument}/{item}", response, $"{postYear}/{account}");
                respFinal = respFinal + "\\n" + $"{salesDocument}/{item}:" + response + ", "+ $"{postYear}/{account}";


                #endregion
            }
            sap.KillSAP();
            dtResponse.AcceptChanges();
            #endregion
            #region Create results Excel
            console.WriteLine("Save Excel...");
            excel.CreateExcel(dtResponse, dtResponseSheetName, dtResponseRoute);
            #endregion
            #region SendEmail
            console.WriteLine("Send Email...");
            mail.SendHTMLMail(htmlContents, new string[] { sender }, root.Subject, cc, attachments);
            #endregion

            root.requestDetails = respFinal;

        }
        private string zvff4Recognition(string salesDocument, string item, string Cocode, string postYear, string account, string postDate)
        {
            string sMessage = "";
            string yearP = postYear.Substring(0, 4);
            string monthP = postYear.Substring(postYear.Length - 3);
            console.WriteLine(DateTime.Now + " > > > " + $"Corriendo Script de SAP, Sales Document: {salesDocument}, {item}");
            //((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nZVF44";
            //SapVariants.frame.SendVKey(0);
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSBUKRS-LOW")).Text = Cocode;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSVBELN-LOW")).Text = salesDocument;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSPOSNR-LOW")).Text = item;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtP_SAKDR-LOW")).Text = account;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtPPOPER_L")).Text = monthP;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtPGJAHR_L")).Text = yearP;
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtPPOSTDAT")).Text = postDate;
            ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkPBLKZ")).Selected = true;
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
            try
            {
                SAPFEWSELib.GuiFrameWindow frame1 = (SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[1]");
                //frame1.Iconify();
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
            }
            catch (Exception)
            {

            }

            try
            {

                //si intenta clickear pero no existe se va al catch
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[0]/shell")).SelectedRows = "0";
                //Delete Blocking ID
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[0]/shell")).PressToolbarContextButton("REVUNLOCK");

                System.Threading.Thread.Sleep(2000);

                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[0]/shell")).SelectedRows = "0";
                //Collective processing
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[0]/shell")).PressToolbarContextButton("SAMD");

                //tomar el stauts de la fila
                sMessage = ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[0]/shell")).GetCellValue(0, "STAPO");
            }
            catch (Exception)
            {
                sMessage = "No existe la línea buscada, verifique los datos y envíe nuevamente";
            }

            if (sMessage == "@5B\\QProcessed Without Errors@")
            {
                sMessage = sMessage.Replace("@5B\\Q", "").Replace("@", "");
            }
            else if (sMessage == @"@5C\\QProcessed with Errors@")
            {
                sMessage = sMessage.Replace("@5C\\Q", "").Replace("@", "");
            }

            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();

            return sMessage;
        }
    }
}
