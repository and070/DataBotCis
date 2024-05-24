using Excel = Microsoft.Office.Interop.Excel;
using SP = Microsoft.SharePoint.Client;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Security;
using System.Security;
using System.Data;
using System.IO;
using System;
using System.Collections.Generic;

namespace DataBotV5.Automation.RPA.Invoice
{
    /// <summary>
    /// Clase RPA Automation encargada de la gestión de Invoice de GBM.
    /// </summary>
    class GBMInvoice
    {
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly SecureAccess login = new SecureAccess();
        readonly Rooting root = new Rooting();
        readonly Database db = new Database();
        readonly Stats stats = new Stats();
        readonly SharePoint sp = new SharePoint();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();
        string respFinal = "";


        public void Main()
        {
            int count = 0;

            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\invoice\\";
            string response = "";
            DirectoryInfo directory = new DirectoryInfo(path);
            foreach (FileInfo file in directory.GetFiles())
                file.Delete();

            if (mail.GetAttachmentEmail("Solicitudes IBMINVOICES", "Procesados", "Procesados IBMINVOICES"))
            {
                //pausar el robot 1 minuto a la espera de mas correos
                System.Threading.Thread.Sleep(60000);

                while (root.filesList != null || count > 200)
                {
                    if (root.BDUserCreatedBy.ToLower() != "")
                    {
                        //buscar archivos ver si no hay attachments
                        for (int i = 0; i < root.filesList.Length; i++)
                        {
                            if (root.filesList[i].Contains(".data"))
                            {
                                response = ConvertFileToPdfAndUpload(root.FilesDownloadPath + "\\" + root.filesList[i], path);
                                if (response != "Se cargo correctamente")
                                    response = "Error procesando el archivo " + root.filesList[i] + " " + response + "<br>";

                                File.Delete(root.FilesDownloadPath + "\\" + root.filesList[i]);
                                
                            }
                        }

                        string[] mails = GetInvoiceUsers();
                        if (mails[0].Contains("Error"))
                            response = mails[0];

                        if (response != "Se cargo correctamente")//algo dio error
                            mail.SendHTMLMail(response, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, new string[] { "clgarcia@gbm.net" });
                        else//nada dio error
                        {
                            string[] attach = Directory.GetFiles(path);
                            string body = "La factura se cargo correctamente<br><br> Se puede visualizar en: https://gbmcorp.sharepoint.com/sites/IBMINVOICES/Documentos%20compartidos/Forms/AllItems.aspx";
                            mail.SendHTMLMail(body, mails, root.Subject, null, attachments: attach);
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "GBM invoice", body, Path.GetFileName(attach[0]) );
                            respFinal = respFinal + "\\n" + "GBM invoice, la factura se cargó correctamente, el nombre del archivo: " + Path.GetFileName(attach[0]);


                            foreach (FileInfo file in directory.GetFiles())
                                file.Delete();
                        }
                    }
                    //volver a buscar mas correos
                    root.filesList = null;
                    mail.GetAttachmentEmail("Solicitudes IBMINVOICES", "Procesados", "Procesados IBMINVOICES");
                    count++;
                }
                
                root.requestDetails = respFinal;

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        private string[] GetInvoiceUsers()
        {
            List<string> response = new List<string>();

            try
            {
                DataTable emailUsersDt = crud.Select("select * from users", "invoice");

                foreach (DataRow row in emailUsersDt.Rows)
                    response.Add(row[1].ToString());
            }
            catch (Exception e)
            {
                response.Clear();
                response.Add(e.Message);
            }

            return response.ToArray();
        }
        private string ConvertFileToPdfAndUpload(string file, string path)
        {
            string filename;

            file = File.ReadAllText(file);

            #region Tomar datos del archivo de texto
            string[] separator = { "THESE COMMODITIES LICENSED BY THE UNITED STATES FOR ULTIMATE DESTINATION" };
            string[] separatorEnter = { "\r\n" };
            string[] separatorBar = { "|" };
            string[] separatorSpace = { " " };
            string[] separatorValue = { "TOTAL VALUE:" };
            string country;
            string shipToCode;

            string[] splitInvoice = file.Split(separator, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < splitInvoice.Length; i++)
            {
                try
                {
                    splitInvoice[i] = splitInvoice[i].Substring(splitInvoice[i].IndexOf("REMIT TO:"));
                }
                catch (Exception)
                {
                    //eliminar el elemento del array
                }
            }

            string[] invoiceLines = splitInvoice[0].Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);

            //sacar la direccion del remit to
            string remitTo = invoiceLines[1].Trim();

            //eliminar el elemento del array

            //Sacar el Invoice to
            //tomar el string que va desde "invoice to" a "ship to" guardarlo en invoice
            string invoice = splitInvoice[0].Substring(splitInvoice[0].IndexOf("INVOICE TO"));
            invoice = invoice.Remove(invoice.LastIndexOf("SHIP TO"));

            //hacer un array con las lineas del invoice y guardarlo en lineas
            invoiceLines = invoice.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);

            //quitar todo lo que esta despues de "|" en cada linea.
            for (int i = 0; i < invoiceLines.Length; i++)
            {
                try
                {
                    invoiceLines[i] = invoiceLines[i].Remove(invoiceLines[i].IndexOf("|"));
                    invoiceLines[i] = invoiceLines[i].Trim() + "\r\n";
                }
                catch (Exception) { }
            }
            //unimos todas las lineas en invoice
            //invoice = string.Concat(lineas);
            string[] invoiceTo = invoiceLines;

            //remover la palabra"invoice to"
            invoice = invoice.Replace("INVOICE TO\r\n", "");

            //extractar la fecha
            //tomar el string que va desde "invoice date" a "invoice type" guardarlo en invoicedate
            string invoiceDate = file.Substring(file.IndexOf("INVOICE DATE"));

            //hacer un array con las lineas del invoicedate y guardarlo en lineas
            invoiceLines = invoiceDate.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);
            invoiceLines = invoiceLines[1].Split(separatorBar, StringSplitOptions.RemoveEmptyEntries);
            invoiceDate = invoiceLines[1].Trim();
            invoiceDate = invoiceDate.Replace(" ", "/");

            //extractar el country y ship to code
            //tomar el string que va desde "invoice date" a "invoice type" guardarlo en invoiceshipcountry
            string invoiceShipCountry = file.Substring(file.IndexOf("SHIP TO CTY/LOC"));

            //hacer un array con las lineas del invoiceshipcountry y guardarlo en lineas
            invoiceLines = invoiceShipCountry.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);
            invoiceLines = invoiceLines[2].Split(separatorBar, StringSplitOptions.RemoveEmptyEntries);

            if (invoiceLines[2][3] == ' ')
                country = "";
            else
                country = invoiceLines[2].Substring(3, 5);

            if (invoiceLines[2][9] == ' ')
                shipToCode = "";
            else
                shipToCode = invoiceLines[2].Substring(9, 5);


            //extractar el invoice N
            //tomar el string que va desde "INVÑ" a XX guardarlo en invoicen
            string invoiceN = file.Substring(file.IndexOf("| INV"));

            //hacer un array con las lineas del invoicen y guardarlo en lineas
            invoiceLines = invoiceN.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);
            invoiceLines = invoiceLines[1].Split(separatorBar, StringSplitOptions.RemoveEmptyEntries);
            string invoiceNum = invoiceLines[3].Trim();

            //Extraer el Shipto
            //tomar el string que va desde "ship to" a "-----" guardarlo en invoiceshipto
            string invoiceShipTo = splitInvoice[0].Substring(splitInvoice[0].IndexOf(" SHIP TO"));
            invoiceShipTo = invoiceShipTo.Remove(invoiceShipTo.IndexOf("-"));

            //hacer un array con las lineas del invoiceshipt y guardarlo en lineas
            invoiceLines = invoiceShipTo.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);

            //quitar todo lo que esta despues de "|" en cada linea.
            for (int i = 0; i < invoiceLines.Length; i++)
            {
                try
                {
                    invoiceLines[i] = invoiceLines[i].Remove(invoiceLines[i].IndexOf("|"));
                    invoiceLines[i] = invoiceLines[i].Trim() + "\r\n";
                }
                catch (Exception) { }
            }
            //unimos todas las lineas en invoiceshipto
            //invoiceshipto = string.Concat(lineas);

            //remover la palabra" SHIP TO"
            invoiceShipTo = invoiceShipTo.Replace("SHIP TO\r\n", "");

            string[] invoiceShipTo1 = invoiceLines;

            //extractar la INVOICE TYPE
            //tomar el string que va desde "INVOICE TYPE" a "invoice type" guardarlo en INVOICE TYPE
            string invoiceType = file.Substring(file.IndexOf("INVOICE TYPE"));

            //hacer un array con las lineas del invoicedate y guardarlo en lineas
            invoiceLines = invoiceType.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);
            invoiceLines = invoiceLines[0].Split(separatorBar, StringSplitOptions.RemoveEmptyEntries);
            invoiceType = invoiceLines[0].Replace("INVOICE TYPE:", "");
            invoiceType = invoiceType.Trim();

            //extractar la CONTRACT
            //tomar el string que va desde "invoice TYPE" a "CONTRACT" guardarlo en invoiceCONTRACT
            string invoicecontract = file.Substring(file.IndexOf("CONTRACT"));

            //hacer un array con las lineas del invoicedate y guardarlo en lineas
            invoiceLines = invoicecontract.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);
            invoiceLines = invoiceLines[0].Split(separatorBar, StringSplitOptions.RemoveEmptyEntries);
            invoicecontract = invoiceLines[0].Replace("CONTRACT:", "");
            invoicecontract = invoicecontract.Trim();

            //extractar TEXTOFINAL
            //tomar el string que va desde "---------" a "" guardarlo en invoicetextofinal
            string invoiceFormatted = file.Substring(file.IndexOf("-------------------------------------------------------------------------------"));

            for (int i = 0; i < splitInvoice.Length; i++)
            {
                try
                {
                    splitInvoice[i] = splitInvoice[i].Substring(splitInvoice[i].IndexOf("-------------------------------------------------------------------------------"));
                    splitInvoice[i] = splitInvoice[i].Replace("------------------------------------------------------------------------------- \r\nORD�       PRODUCT      SERIAL�        NT/NM  C/O MES QTY UNIT PRICE    EXT.AMT \r\nCUST�      DESCRIPTION                                      (USA$)       (USA$) \r\n------------------------------------------------------------------------------- \r\n", "");
                }
                catch (Exception) { }
            }

            invoiceFormatted = string.Concat(splitInvoice);

            //TOTAL VALUE:
            //1. hacer split de invoicetextofinal donde diga  Total value:

            invoiceLines = invoiceFormatted.Split(separatorValue, StringSplitOptions.RemoveEmptyEntries);

            //2. Al split1 de eso hay qye hacerle un indexoff enter,
            invoiceLines[1] = invoiceLines[1].Substring(0, invoiceLines[1].IndexOf("\r\n"));
            invoiceLines[1] = "TOTAL VALUE:" + invoiceLines[1];

            //cuando esta lista esa al indexoff le concatenamos linea0
            invoiceFormatted = string.Concat(invoiceLines);
            #endregion

            #region Manejo del Excel
            int lastRow;
            int firstRow;

            Excel.Application excelWin = new Excel.Application { Visible = false };
            Excel.Workbook excelBook = excelWin.Workbooks.Add();

            #region Cambiar fuente
            excelWin.Cells.Font.Name = "Consolas";
            excelWin.Cells.Font.Size = 10;
            #endregion

            #region Cambiar ancho de columna  
            excelWin.Columns["C:C"].ColumnWidth = 12.64;
            excelWin.Columns["D:D"].ColumnWidth = 14.18;
            excelWin.Columns["G:G"].ColumnWidth = 5.82;
            excelWin.Columns["F:F"].ColumnWidth = 6.27;
            excelWin.Columns["H:H"].ColumnWidth = 13;
            #endregion

            excelWin.Range["A1"].Value = "IBM Invoice N: " + invoiceNum;
            excelWin.Range["A1"].Font.Color = -4165632;
            excelWin.Range["A1"].Font.Bold = true;

            excelWin.Range["A2"].Value = "REMIT TO:";
            BoldStyle("A2", excelWin);

            excelWin.Range["A3"].Value = remitTo;

            //colocar invoice to
            excelWin.Range["A4"].Value = "Ivoice To:";
            for (int i = 1; i < invoiceTo.Length; i++)
            {
                string row = (4 + i).ToString();
                excelWin.Range["A" + row].Value = invoiceTo[i].Trim();
            }

            lastRow = excelWin.Cells[excelWin.Rows.Count, "A"].End(Excel.XlDirection.xlUp).Row; //busca la ultima fila
            BorderStyle("A4:C" + (lastRow + 1), excelWin);

            excelWin.Range["D4"].Value = "Invoice Date:";
            excelWin.Range["D5"].Value = invoiceDate.Trim();
            excelWin.Range["D5"].NumberFormat = "dd/mm/yyyy;@";

            excelWin.Range["E4"].Value = "Country:";
            excelWin.Range["E5"].Value = country.Trim();

            excelWin.Range["F4"].Value = "Ship To Code:";
            excelWin.Range["F5"].Value = shipToCode.Trim();

            excelWin.Range["H4"].Value = "Inv N";
            excelWin.Range["H5"].Value = "'" + invoiceNum.Trim();

            BackColorStyle("A4:H4", excelWin);
            GridStyle("d4:H5", excelWin);
            BoldStyle("A4:H4", excelWin);
            CenterText("D4:H5", excelWin);
            BorderStyle("A4:C4", excelWin);

            //colocar ship to
            firstRow = lastRow + 2;
            excelWin.Range["A" + firstRow].Value = "Ship To:";
            BackColorStyle("A" + firstRow + ":C" + firstRow, excelWin);
            BorderStyle("A" + firstRow + ":C" + firstRow, excelWin);

            for (int i = 1; i < invoiceShipTo1.Length; i++)
            {
                string row = (firstRow + i).ToString();
                excelWin.Range["A" + row].Value = invoiceShipTo1[i].Trim();
            }
            lastRow = excelWin.Cells[excelWin.Rows.Count, "A"].End(Excel.XlDirection.xlUp).Row;
            BorderStyle("A" + (firstRow + 1) + ":C" + lastRow, excelWin);
            BoldStyle("A" + firstRow, excelWin);

            excelWin.Range["D7"].Value = "INVOICE TYPE: " + invoiceType.Trim();
            excelWin.Range["D8"].Value = "CONTRACT: " + invoicecontract.Trim();
            BorderStyle("d6:H" + lastRow, excelWin);

            lastRow = excelWin.Cells[excelWin.Rows.Count, "A"].End(Excel.XlDirection.xlUp).Row;
            excelWin.Range["A" + (lastRow + 1)].Value = "-------------------------------------------------------------------------------";
            excelWin.Range["A" + (lastRow + 2)].Value = "ORD        PRODUCT      SERIAL         NT/NM  C/O MES QTY UNIT PRICE    EXT.AMT";
            excelWin.Range["A" + (lastRow + 3)].Value = "CUST       DESCRIPTION                                      (USA$)       (USA$)";
            excelWin.Range["A" + (lastRow + 4)].Value = "-------------------------------------------------------------------------------";

            invoiceFormatted = invoiceFormatted.Replace("�", " ");
            invoiceLines = invoiceFormatted.Split(separatorEnter, StringSplitOptions.RemoveEmptyEntries);

            lastRow = 6 + lastRow;

            for (int i = 0; i < invoiceLines.Length; i++)
            {
                string row = (lastRow + i).ToString();
                excelWin.Range["A" + row].Value = invoiceLines[i];
            }

            excelWin.Range["F4:G4"].Merge();
            excelWin.Range["F5:G5"].Merge();
            PageNumbering(excelWin);

            filename = country.Trim() + "-" + invoiceNum.Trim();

            excelWin.ActiveSheet.ExportAsFixedFormat(Type: Excel.XlFixedFormatType.xlTypePDF,
                                                     Filename: path + filename,
                                                     Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                                                     IncludeDocProperties: true,
                                                     IgnorePrintAreas: false,
                                                     OpenAfterPublish: false);

            proc.KillProcess("EXCEL", true);
            #endregion

            #region Subir a Sharepoint

            string passSharepoint = crud.Select("SELECT `pwd` FROM `pass`", "invoice").Rows[0]["pwd"].ToString();
            passSharepoint = login.DecodePass(passSharepoint);

            string spResponse = UploadFileToSharePoint(path + filename + ".pdf", "invoicesgbm@gbm.net", passSharepoint);

            #endregion

            console.WriteLine(spResponse);

            return spResponse;
        }
        private string UploadFileToSharePoint(string FileName, string Login, string Password)
        {
            string SiteUrl = "https://gbmcorp.sharepoint.com/sites/IBMINVOICES";
            string DocLibrary = "Documentos";

            try
            {
                #region ConnectToSharePoint
                var securePassword = new SecureString();
                foreach (char c in Password)
                { securePassword.AppendChar(c); }
                var onlineCredentials = new SP.SharePointOnlineCredentials(Login, securePassword);
                #endregion

                #region Insert the data
                using (SP.ClientContext CContext = new SP.ClientContext(SiteUrl))
                {
                    CContext.Credentials = onlineCredentials;
                    SP.Web web = CContext.Web;
                    SP.FileCreationInformation newFile = new SP.FileCreationInformation();
                    byte[] FileContent = System.IO.File.ReadAllBytes(FileName);
                    newFile.ContentStream = new MemoryStream(FileContent);
                    newFile.Url = Path.GetFileName(FileName);

                    SP.List DocumentLibrary = web.Lists.GetByTitle(DocLibrary);

                    SP.File uploadFile = DocumentLibrary.RootFolder.Files.Add(newFile);

                    CContext.Load(DocumentLibrary);
                    CContext.Load(uploadFile);
                    CContext.ExecuteQuery();

                    return "Se cargo correctamente";
                }
                #endregion
            }
            catch (Exception exp)
            {
                return "Error al subir a Sharepoint: " + exp.Message;
            }
        }
        private void BorderStyle(string range, Excel.Application excel)
        {
            excel.Range[range].Select();

            excel.Selection.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone;
            excel.Selection.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.Constants.xlNone;
            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone;
        }
        private void GridStyle(string range, Excel.Application excel)
        {
            excel.Range[range].Select();

            excel.Selection.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone;
            excel.Selection.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium;

            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = false;
            excel.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium;

        }
        private void BackColorStyle(string range, Excel.Application excel)
        {
            excel.Range[range].Select();
            excel.Selection.Interior.Pattern = Excel.Constants.xlSolid;
            excel.Selection.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
            excel.Selection.Interior.Color = 11711154;
        }
        private void BoldStyle(string range, Excel.Application excel)
        {
            excel.Range[range].Select();
            excel.Range[range].Font.Bold = true;
        }
        private void CenterText(string range, Excel.Application excel)
        {

            excel.Range[range].HorizontalAlignment = Excel.Constants.xlCenter;
            excel.Range[range].VerticalAlignment = Excel.Constants.xlBottom;
            excel.Range[range].WrapText = false;
            excel.Range[range].Orientation = false;
            excel.Range[range].AddIndent = false;
            excel.Range[range].IndentLevel = false;
            excel.Range[range].ShrinkToFit = false;
            excel.Range[range].MergeCells = false;
        }
        private void PageNumbering(Excel.Application excel)
        {
            excel.ActiveSheet.PageSetup.LeftHeader = "";
            excel.ActiveSheet.PageSetup.CenterHeader = "";
            excel.ActiveSheet.PageSetup.RightHeader = "PAGINA &P";
            excel.ActiveSheet.PageSetup.LeftFooter = "";
            excel.ActiveSheet.PageSetup.CenterFooter = "";
            excel.ActiveSheet.PageSetup.RightFooter = "";
            excel.ActiveSheet.PageSetup.LeftMargin = 79.3700787401575;
            excel.ActiveSheet.PageSetup.RightMargin = 51.0236220472441;
            excel.ActiveSheet.PageSetup.TopMargin = 53.8582677165354;
            excel.ActiveSheet.PageSetup.BottomMargin = 53.8582677165354;
            excel.ActiveSheet.PageSetup.HeaderMargin = 22.6771653543307;
            excel.ActiveSheet.PageSetup.FooterMargin = 22.6771653543307;
            excel.ActiveSheet.PageSetup.PrintHeadings = false;
            excel.ActiveSheet.PageSetup.PrintGridlines = false;
            excel.ActiveSheet.PageSetup.PrintComments = Excel.XlPrintLocation.xlPrintNoComments;
            excel.ActiveSheet.PageSetup.CenterHorizontally = false;
            excel.ActiveSheet.PageSetup.CenterVertically = false;
            excel.ActiveSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            excel.ActiveSheet.PageSetup.Draft = false;
            excel.ActiveSheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLetter;
            excel.ActiveSheet.PageSetup.FirstPageNumber = Excel.Constants.xlAutomatic;
            excel.ActiveSheet.PageSetup.Order = Excel.XlOrder.xlDownThenOver;
            excel.ActiveSheet.PageSetup.BlackAndWhite = false;
            excel.ActiveSheet.PageSetup.Zoom = 100;
            excel.ActiveSheet.PageSetup.PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed;
            excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter = false;
            excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter = false;
            excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter = true;
            excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter = true;
            excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text = "";
            excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text = "";
            excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text = "";
            excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text = "";
            excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text = "";
            excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text = "";
            excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text = "";
            excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text = "";
            excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text = "";
            excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text = "";
            excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text = "";
            excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text = "";
        }
    }
}
