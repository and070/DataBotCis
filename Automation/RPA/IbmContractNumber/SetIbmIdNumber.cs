using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.BusinessSystem;
using DataBotV5.Logical.Web;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SAP.Middleware.Connector;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DataBotV5.Automation.RPA.IbmContractNumber
{
    public class SetIbmIdNumber
    {
        BusinessSystemLogical bsLogical = new BusinessSystemLogical();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        WebInteraction web = new WebInteraction();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        BsSQL bsql = new BsSQL();
        CRUD crud = new CRUD();
        Log log = new Log();
        int mandante = 260;
        string SSMandante = "QAS";
        string erpSystem = "ERP";
        string respFinal = "";

        /// <summary>
        /// 
        /// </summary>
        public void Main()
        {

            if (mail.GetAttachmentEmail("Requests IBM Contract Number", "Procesados", "Processed IBM Contract Number"))
            {
                console.WriteLine(" Procesando....");
                Process(root.Email_Body);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ExcelFile">el excel que envía el país por email outlook</param>
        private void Process(string body)
        {
            #region private variables
            //variable de respuesta en caso de que se necesite
            string response = "";
            //DataTable furuto excel de respuesta
            DataTable dtResponse = new DataTable();
            //variable nombre de la hoja de resultados
            string dtResponseSheetName = "Results";
            //variable nombre del libro de resultados + extension
            string dtResponseBookName = $"ResultsBook{DateTime.Now.ToString("yyyyMMdd")}" + root.ExcelFile + ".xlsx";
            //ruta + nombre donde se guardará el excel de resultado
            string dtResponseRoute = root.FilesDownloadPath + "\\" + dtResponseBookName;
            //PLantilla en html para el envío de email
            string htmlEmail = Properties.Resources.emailtemplate1;
            //variable titulo del cuerpo del correo
            string htmlSubject = "Resultados";
            //variable contenido del correo: texto, cuadros, tablas, imagenes, etc
            string htmlContents = "";
            //variable remitente del email de respuesta
            string sender = root.BDUserCreatedBy;
            //variable copias del email de respuesta
            string[] cc = new string[] { "dmeza@gbm.net" }; // root.CopyCC;
            //variable ruta de adjunto
            string[] attachments = new string[] { dtResponseRoute };
            //variable cambio para log
            string logText = "";
            #endregion

            #region robot Process
            try
            {
                string documentId = root.Subject.Split(new string[] { "PO #: " }, StringSplitOptions.None)[1];
                string countryEmail = documentId.Substring(0, 2);
                documentId = documentId.Substring(2, documentId.Length - 2);
                if (documentId != "")
                {
                    //DataTable dt = crud.Select($"SELECT IBMContractNumber FROM document_system WHERE documentId = {documentId}", "document_system");
                    //if (dt.Rows.Count == 0)
                    //{
                    //    return; //ya se proceso
                    //}
                    Regex reg;
                    reg = new Regex("[*'\"_&+^><@]");
                    string cleanBody = reg.Replace(body, string.Empty);
                    int startIndex = root.Subject.IndexOf("order for ") + "order for ".Length;
                    int endIndex = root.Subject.IndexOf(", PO #");
                    string countryName = "";
                    if (startIndex >= 0 && endIndex >= 0)
                    {
                        countryEmail = root.Subject.Substring(startIndex, endIndex - startIndex);
                        countryName = countryEmail;
                    }
                    else
                    {
                        countryName = getCountryName(countryEmail);
                    }

                    string ibmContractNumber = "";
                    if (root.Subject.Contains("[External]We have received your order for"))
                    {
                        ibmContractNumber = body.Split(new string[] { "<https://engage-support.ibm.com/staticResources/images/requestType/other.png>" }, StringSplitOptions.None)[0];

                        ibmContractNumber = ibmContractNumber.Split(new string[] { countryName }, StringSplitOptions.None)[1].Trim();
                    }

                    string salesOrder = "";
                    string country = "";
                    string countryNameDb = "";
                    string countryAdmin = "";
                    string customer = "";
                    string documentType = (documentId.Substring(0, 2) == "10") ? "PO" : "SO";
                    string specialBidNumber = "";

                    #region buscar PO/SO en el sistema de documentos

                    DataTable dataTable = crud.Select($@"SELECT 
                                                            document_system.salesOrderOn, 
                                                            document_system.discountNumber,
                                                            document_system.country, 
                                                            document_system.createdBy, 
                                                            document_system.customerName,
                                                            document_country.countryName
                                                            FROM `document_system` 
                                                            INNER JOIN document_country ON document_country.countryCode = document_system.country
                                                            WHERE document_system.documentId = '{documentId}'", "document_system");
                    if (dataTable.Rows.Count > 0)
                    {
                        if (documentType == "PO")
                        {
                            salesOrder = dataTable.Rows[0]["salesOrderOn"].ToString();
                        }
                        else if (documentType == "SO")
                        {
                            salesOrder = documentId;
                        }

                        country = dataTable.Rows[0]["country"].ToString();
                        countryNameDb = dataTable.Rows[0]["countryName"].ToString();
                        countryAdmin = dataTable.Rows[0]["createdBy"].ToString();
                        customer = dataTable.Rows[0]["customerName"].ToString();
                        specialBidNumber = dataTable.Rows[0]["discountNumber"].ToString();

                        //Proceso para guardar el SB de la orden 
                        if (specialBidNumber != "")
                        {
                            pdfResponse pdfRes = getSbPdf(specialBidNumber, documentId, countryNameDb, customer);
                            if (!pdfRes.response)
                            {
                                string htmlEmail2 = Properties.Resources.emailtemplate1;

                                string[] cCopy = bsql.EmailAddress(3);
                                Array.Resize(ref cCopy, cCopy.Length + 1);
                                cCopy[cCopy.Length - 1] = "dmeza@gbm.net";
                                htmlEmail2 = htmlEmail2
                                    .Replace("{subject}", "Error al Descargar SB PDF HW")
                                    .Replace("{cuerpo}", $"No se pudo descargar el SB {specialBidNumber} para el documento {documentId}")
                                    .Replace("{contenido}", pdfRes.error);
                                mail.SendHTMLMail(htmlEmail2, new string[] { root.f_sender }, root.Subject, cCopy, null);
                            }
                        }
                    }

                    #endregion
                    if (salesOrder == "")
                    {
                        #region Buscar Sales Order en SAP
                        Dictionary<string, string> dic = new Dictionary<string, string>
                        {
                            ["TIPO_DOC"] = documentType,
                            ["ID_DOC"] = documentId,

                        };
                        IRfcFunction func = sap.ExecuteRFC(erpSystem, "ZFI_READ_PO", dic);
                        if (func.GetValue("RESPUESTA").ToString() != "NA")
                        {
                            if (documentType == "PO")
                            {
                                salesOrder = func.GetValue("SO_PAIS").ToString();
                            }
                            else if (documentType == "SO")
                            {
                                salesOrder = documentId;
                            }
                            country = func.GetValue("COCODE_PAIS").ToString();
                            string storageLocation = (country == "DR") ? "DO01" : country + "01";
                            customer = func.GetValue("CUSTOMER_NAME").ToString();
                            //buscar por medio del país el admin support en una tabla Z o bien de S&S
                            //ZIBM_CONTRA_CONF
                            DataTable people = sap.GetSapTable("ZIBM_CONTRA_CONF", "ERP");

                            RfcDestination destErp = new SapVariants().GetDestRFC("ERP");
                            IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
                            fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
                            fmMg.SetValue("QUERY_TABLE", "ZIBM_CONTRA_CONF");
                            fmMg.SetValue("DELIMITER", "");

                            IRfcTable fields = fmMg.GetTable("FIELDS");
                            fields.Append();
                            fields.SetValue("FIELDNAME", "EMAIL");

                            IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
                            fmOptions.Append();
                            fmOptions.SetValue("TEXT", $"STORAGE_LOCATION = '{storageLocation}'");

                            fmMg.Invoke(destErp);

                            DataTable tableSap = sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA"));
                            if (tableSap.Rows.Count > 0)
                            {
                                countryAdmin = tableSap.Rows[0]["LINE"].ToString();
                            }


                            //---------------------
                        }


                        #endregion
                    }


                    if (salesOrder == "")
                    {
                        htmlEmail = htmlEmail.Replace("{subject}", "").Replace("{cuerpo}", $"No se encontró la Orden de Venta para el documento {documentId}, Coloque el número de contrato de IBM {ibmContractNumber} manualmente para que se envíe la notificación correspondiente al cliente {customer} al día siguiente").Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlEmail, new string[] { countryAdmin }, root.Subject, cc, null);
                        return;
                    }

                    crud.Update($"UPDATE document_system SET IBMContractNumber = '{ibmContractNumber}' WHERE documentId = {documentId}", "document_system");

                    //Insertar el contrato de IBM en el Sales Order de SAP
                    Dictionary<string, string> parameters = new Dictionary<string, string>
                    {
                        ["SALES_ORDER"] = salesOrder,
                        ["IBM_CONTRACT_NUMBER"] = ibmContractNumber,

                    };
                    IRfcFunction function = sap.ExecuteRFC(erpSystem, "ZFI_BS_ADD_IBM_CONTRACT_NUM", parameters); //transporte: DEVK952625

                    response = function.GetValue("RESPONSE").ToString();

                    if (response != "ok")
                    {
                        htmlEmail = htmlEmail.Replace("{subject}", "Error al cargar el número de contrato de IBM en la orden").Replace("{cuerpo}", $"No se pudo cargar el número de contrato de IBM {ibmContractNumber} en la Orden de Venta {salesOrder} para el documento {documentId}, por favor notifique al cliente {customer}").Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlEmail, new string[] { countryAdmin }, root.Subject, cc, null);
                        return;
                    }
                    //--------------------------

                }

            }
            catch (Exception ex)
            {
                htmlEmail = htmlEmail.Replace("{subject}", "Error no esperado").Replace("{cuerpo}", ex.Message).Replace("{contenido}", "");
                mail.SendHTMLMail(htmlEmail, new string[] { "dmeza@gbm.net" }, root.Subject, cc, null);
                return;
            }
            //log de cambios
            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, logText, "", "");
            respFinal = respFinal + "\\n" + "Creación Ibm Number: " + logText;


            #endregion

            //#region Create results Excel
            //dtResponse.AcceptChanges();
            //console.WriteLine("Save Excel...");
            //excel.CreateExcel(dtResponse, dtResponseSheetName, dtResponseRoute);
            //#endregion
            #region SendEmail
            //console.WriteLine("Send Email...");
            //response = (response == "") ? "error" : "success";
            //htmlEmail = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", response).Replace("{contenido}", htmlContents);
            //mail.SendHTMLMail(htmlEmail, "jealor@gbm.net", root.Subject, cc, null);

            #endregion

            root.requestDetails = respFinal;

        }
        private string getCountryName(string country)
        {
            string countryName = "";
            switch (country)
            {
                case "CR":
                    countryName = "GBM COSTA RICA";
                    break;
                case "GT":
                    countryName = "GBM DE GUATEMALA S.A.";
                    break;
                case "NI":
                    countryName = "GBM DE NICARAGUA";
                    break;
                case "DR":
                    countryName = "GBM DOMINICANA";
                    break;
                case "SV":
                    countryName = "GBM EL SALVADOR";
                    break;
                case "HN":
                    countryName = "GBM HONDURAS";
                    break;
                case "PA":
                    countryName = "GBM PANAMA";
                    break;
                default:
                    break;
            }
            return countryName;
        }

        private pdfResponse getSbPdf(string sb, string documentId, string countryNameDb, string customerName)
        {
            pdfResponse pdfResp = new pdfResponse();

            IWebDriver chrome = web.NewSeleniumChromeDriver(root.FilesDownloadPath + "\\");
            try
            {
                string filePdfName = "SpecialBidAddendum" + sb + ".pdf";
                string filePdfPath = root.FilesDownloadPath + "\\" + filePdfName;
                chrome.Navigate().GoToUrl("https://www.ibm.com/services/partners/epricer/v2/directLogin.do");
                new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("username")));
                chrome.FindElement(By.Id("username")).SendKeys("ePricer_CR@gbm.net");
                chrome.FindElement(By.Id("continue-button")).Click();
                new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("password")));

                chrome.FindElement(By.Id("password")).SendKeys("ePricer20");
                chrome.FindElement(By.Id("signinbutton")).Submit();

                bool auth = bsLogical.doubleAuth(chrome, documentId);
                if (!auth)
                {
                    chrome.Close();
                    proc.KillProcess("chromedriver", true);
                    pdfResp.response = false;
                    pdfResp.error = "Error en la doble autentificación de IBM";
                    return pdfResp;
                }

                WebDriverWait wait = new WebDriverWait(chrome, TimeSpan.FromSeconds(35));

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 35)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ibm-columns-main']/div/main-indicator"))); }
                catch
                { }



                //wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//*[@id='ibm-columns-main']/div/main-indicator")));
                IWebElement loader = chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/main-indicator"));
                // Define a custom ExpectedCondition to wait for the element's class attribute to change
                Func<IWebDriver, bool> elementIsVisible = (driver) =>
                {
                    try
                    {
                        return loader.GetAttribute("class").Contains("ng-hide");
                    }
                    catch (StaleElementReferenceException ex)
                    {
                        return false;
                    }
                };

                wait.Until(elementIsVisible);

                //SelectElement resol = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ibm-pagetitle-h1']/div[1]/div[5]/div[2]/div[2]/select")));
                //resol.SelectByValue("object:91"); //10air0kb - GBM Corp - Distributor

                // Assuming you have a WebDriver instance named "driver"
                IWebElement dropdown = chrome.FindElement(By.XPath("//*[@id='ibm-pagetitle-h1']/div[1]/div[5]/div[2]/div[2]/select"));
                // Wait for the dropdown element to be present
                wait.Until(ExpectedConditions.ElementToBeClickable(dropdown));

                // Create a SelectElement for the dropdown
                SelectElement resol = new SelectElement(dropdown);

                // Wait for the options to be populated (you can adjust the timeout as needed)
                wait.Until(driver => resol.Options.Count > 0);

                // Now you can select an option by value
                resol.SelectByValue("object:92"); // 10air0kb - GBM Corp - Distributor


                System.Threading.Thread.Sleep(1000);

                //SelectElement role = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ibm - pagetitle - h1'']/div[1]/div[5]/div[3]/div[2]/select")));
                //System.Threading.Thread.Sleep(1000);
                //role.SelectByValue("string:LACR_PTRDSTA01"); 

                chrome.FindElement(By.XPath("//*[@id='ibm-pagetitle-h1']/div[1]/div[5]/div[4]/div[2]/button")).Click(); //start

                //Esperar a que termine el portal de cargar


                //wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//*[@id='loader']")));
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 35)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='loader']"))); }
                catch
                { }


                wait.Until(elementIsVisible);

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 35)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='quotesearch']"))); }
                catch
                { }

                //chrome.FindElement(By.XPath("//*[@id='quotesearch']")).Click(); //find quote

                IWebElement quotesearchElement = chrome.FindElement(By.XPath("//*[@id='quotesearch']")); //find quote
                wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                wait.Until(ExpectedConditions.ElementToBeClickable(quotesearchElement));
                ((IJavaScriptExecutor)chrome).ExecuteScript("arguments[0].click();", quotesearchElement);
                //quotesearchElement.Click();

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 35)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='loader']"))); }
                catch
                { }

                wait.Until(elementIsVisible);

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 20)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quotesearch-page/div/div[2]/div[1]/div[2]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div[8]/input"))); }
                catch
                { }

                //buscar el SB number
                chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quotesearch-page/div/div[2]/div[1]/div[2]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div[8]/input")).SendKeys(sb);

                chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quotesearch-page/div/div[2]/div[2]/table/tbody/tr/td[2]/a")).Click(); //find quote

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 35)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='loader']"))); }
                catch
                { }

                // Wait for the element to become visible
                wait.Until(elementIsVisible);

                chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quoteedit-page/div/div/quotedetails-page/div/div/div/ul/li[2]/a")).Click(); //addendum

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 15)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='+@@totalcom@@']/td[4]"))); }
                catch
                { }

                //moverse hasta el boton de descargar 
                try { new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quoteedit-page/div/div/quotedetails-page/div/quoteapprovedaddendum-page/div/div/a[1]"))).Perform(); }
                catch { }

                //*[@id='ibm-content-main']/div[2]/div[1]/div[3]/div/table/tbody/tr[2]/td[3]

                IWebElement href_download = chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quoteedit-page/div/div/quotedetails-page/div/quoteapprovedaddendum-page/div/div/a[1]"));

                if (href_download.GetAttribute("href").Length > 0)
                {
                    href_download.Click();
                    //var descargar_link = chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quoteedit-page/div/div/quotedetails-page/div/quoteapprovedaddendum-page/div/div/a[1]")).GetAttribute("href");

                    //chrome.Navigate().GoToUrl(descargar_link);

                    for (var x = 0; x < 40; x++)
                    {
                        if (File.Exists(filePdfPath))
                        { break; }
                        //if (File.Exists(root.FilesDownloadPath + "\\" + "customer quote PDF " + sbId + ".pdf"))
                        //{ break; }
                        System.Threading.Thread.Sleep(1000);
                    }

                }

                chrome.Close();
                proc.KillProcess("chromedriver", true);

                //crear carpeta en folder compartido y guardar PDF
                string fldrpath = "";
                string ruta = @"\\Fs01\bs\SB\SB IBM HW";

                #region folder año
                //busca si la carpeta del año existe

                fldrpath = ruta + "\\" + DateTime.Now.Year.ToString();
                if (!Directory.Exists(fldrpath))
                {
                    Directory.CreateDirectory(fldrpath);
                }
                #endregion

                #region folder mes
                //"busca si la carpeta del mes existe"
                string nombreMes = new CultureInfo("es-ES").DateTimeFormat.GetMonthName(DateTime.Now.Month);
                nombreMes = char.ToUpper(nombreMes[0]) + nombreMes.Substring(1);

                fldrpath = fldrpath + "\\" + nombreMes;
                if (!Directory.Exists(fldrpath))
                {
                    Directory.CreateDirectory(fldrpath);
                }
                #endregion

                #region folder pais

                fldrpath = fldrpath + "\\" + countryNameDb;
                if (!Directory.Exists(fldrpath))
                {
                    Directory.CreateDirectory(fldrpath);
                }
                #endregion

                #region folder Cliente
                string folder_name = "SB" + sb + " - " + customerName;
                Regex reg2;
                reg2 = new Regex("[*:'\"_&+^><@]");
                folder_name = reg2.Replace(folder_name, string.Empty);

                fldrpath = fldrpath + "\\" + folder_name;
                if (!Directory.Exists(fldrpath))
                {
                    Directory.CreateDirectory(fldrpath);
                }
                #endregion

                #region copiar y pegar el PDF
                try
                {
                    console.WriteLine("Copiando PDF y enviando respuesta");
                    string destinationPath = fldrpath + "\\" + filePdfName;
                    File.Copy(filePdfPath, destinationPath);

                }
                catch (Exception ex)
                {
                    console.WriteLine("Error: " + ex.ToString());
                    string[] cCopy = bsql.EmailAddress(3);
                    //enviar email de error al copiar el archivo
                    mail.SendHTMLMail($"Error al copiar el SB {sb} de HW en el folder Z: " + ex.ToString(), new string[] { "dmeza@gbm.net" }, $"Error al guardar el SB {sb}", cCopy);
                }
                #endregion

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                pdfResp.response = false;
                pdfResp.error = ex.Message;
                return pdfResp;
            }

            pdfResp.response = true;
            pdfResp.error = "OK";
            return pdfResp;

        }
    }
}
public class pdfResponse
{
    public string error { get; set; }
    public bool response { get; set; }
}

