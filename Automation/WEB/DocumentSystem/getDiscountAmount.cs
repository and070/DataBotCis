using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.BusinessSystem;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using Microsoft.Exchange.WebServices.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SAP.Middleware.Connector;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DataBotV5.Automation.WEB.DocumentSystem
{
    public class getDiscountAmount
    {
        BusinessSystemLogical bsLogical = new BusinessSystemLogical();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        WebInteraction web = new WebInteraction();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        WebexTeams wx = new WebexTeams();
        Rooting roots = new Rooting();
        Rooting root = new Rooting();
        BsSQL bsql = new BsSQL();
        CRUD crud = new CRUD();

        int mand = 260;
        string syste = "QAS";
        public void Main()
        {
            DataTable request = crud.Select($@"SELECT 
                                                document_system.*,
                                                document_country.countryName,
                                                Group_Concat(vendors.vendorName) as vendors,
                                                Group_Concat(document_trading.poTrading) as poTradings
                                                FROM `document_system` 
                                                LEFT JOIN document_vendors ON document_vendors.fk_documentId = document_system.documentId
                                                INNER JOIN vendors ON vendors.vendorId = document_vendors.fk_vendor
                                                INNER JOIN document_country ON document_country.countryCode = document_system.country
                                                LEFT JOIN document_trading ON document_trading.documentId = document_system.documentId
                                                LEFT JOIN discountType ON discountType.id = document_system.discountType
                                                WHERE (vendors.vendorName LIKE '%IBM%')
                                                AND document_system.statusRobot = 1
                                                AND document_system.status = 'CD'
                                                AND discountType.discountTypeName = 'Special Bid'
                                                GROUP BY document_system.documentId
                                            UNION ALL
                                               SELECT 
                                                document_system.*,
                                                document_country.countryName,
                                                Group_Concat(vendors.vendorName) as vendors,
                                                Group_Concat(document_trading.poTrading) as poTradings
                                                FROM `document_system` 
                                                LEFT JOIN document_vendors ON document_vendors.fk_documentId = document_system.documentId
                                                INNER JOIN vendors ON vendors.vendorId = document_vendors.fk_vendor
                                                INNER JOIN document_country ON document_country.countryCode = document_system.country
                                                LEFT JOIN document_trading ON document_trading.documentId = document_system.documentId
                                                WHERE vendors.vendorName LIKE '%CISCO%'
                                                AND document_system.statusRobot = 1
                                                AND document_system.status = 'CD'
                                                GROUP BY document_system.documentId
                                            UNION ALL
                                               SELECT 
                                                document_system.*,
                                                document_country.countryName,
                                                Group_Concat(vendors.vendorName) as vendors,
                                                Group_Concat(document_trading.poTrading) as poTradings
                                                FROM `document_system` 
                                                LEFT JOIN document_vendors ON document_vendors.fk_documentId = document_system.documentId
                                                INNER JOIN vendors ON vendors.vendorId = document_vendors.fk_vendor
                                                INNER JOIN document_country ON document_country.countryCode = document_system.country
                                                LEFT JOIN document_trading ON document_trading.documentId = document_system.documentId
                                                LEFT JOIN discountType ON discountType.id = document_system.discountType
                                                WHERE (vendors.vendorName LIKE '%LOTUS%')
                                                AND document_system.statusRobot = 1
                                                AND document_system.status = 'CD'
                                                AND discountType.discountTypeName = 'Special Bid'
                                                GROUP BY document_system.documentId", "document_system", syste);
            if (request.Rows.Count > 0)
            {
                getDiscount(request);
            }
            #region MyRegion
            //string sourceEmail = "databot@gbm.net";
            //string sourcePassword = "X1c$vmjk.5LmN";
            //string destinationEmail = "databotqa@gbm.net";
            //string destinationPassword = "IrhrqW$V0e4A7";

            //ExchangeService sourceService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            //sourceService.Credentials = new WebCredentials(sourceEmail, sourcePassword);
            //sourceService.AutodiscoverUrl(sourceEmail, url => true);

            //RuleCollection sourceRules = sourceService.GetInboxRules(sourceEmail);

            //ExchangeService destinationService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            //destinationService.Credentials = new WebCredentials(destinationEmail, destinationPassword);
            //destinationService.AutodiscoverUrl(destinationEmail, url => true);

            //List<RuleOperation> ruleOperations = new List<RuleOperation>();

            //foreach (Rule sourceRule in sourceRules)
            //{
            //    RuleOperation ruleOperation = new CreateRuleOperation(sourceRule);
            //    ruleOperations.Add(ruleOperation);
            //}

            //// Apply rule changes to the destination service
            //destinationService.UpdateInboxRules(ruleOperations.ToArray(), false);

            //Console.WriteLine("Rules exported and imported successfully.");
            #endregion

        }
        private void getDiscount(DataTable requestInfo)
        {
            foreach (DataRow row in requestInfo.Rows)
            {
                if (row["documentId"].ToString() != "")
                {

                    string vendors = row["vendors"].ToString();
                    console.WriteLine($"Procesando... {row["documentId"]} - {vendors}");
                    bool getAmount = true;
                    AmountResult amountResult = new AmountResult();
                    if (vendors.Contains("IBM"))
                    {
                        amountResult = getDMIBM(row);
                    }
                    else if (vendors.Contains("LOTUS"))
                    {
                        amountResult = getDMLotusSap(row);
                    }
                    else if (vendors.Contains("CISCO"))
                    {
                        amountResult = getDMCisco(row);
                    }

                    if (!amountResult.result)
                    {
                        console.WriteLine("Error al ingresar el Monto");
                        //send error email
                        mail.SendHTMLMail($"Error: no se pudo descargar el monto del número de descuento/Smart Account de la PO {row["documentId"]}", new string[] { "dmeza@gbm.net" }, "Notificación Sistema de Documentos: Error al descargar el monto de descuento", new string[] { "dmeza@gbm.net" });
                    }
                    else
                    {
                        console.WriteLine("Monto Agregado");
                        //wx.SendNotification(row["createdBy"].ToString() + "@gbm.net", "", $"Hola! Ya se agregó el monto en la PO {row["documentId"]}");
                        if (vendors.Contains("IBM") || vendors.Contains("LOTUS"))
                        {
                            //ENVIAR CORREO A PRICING
                            string[] cc = bsql.EmailAddress(11);

                            string body = $@"<table class='myCustomTable' width='100 %'>
                                                <thead>
                                                    <tr>
                                                        <th>Special Bid</th>
                                                        <th>Monto del Special Bid</th>
                                                        <th>Cliente</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                 <tr>
                                                    <td>{row["discountNumber"]}</td>
                                                    <td>{amountResult.discountAmount}</td>
                                                    <td>{row["customerName"]}</td>
                                                  </tr>";

                            string quotesQuery = $@"SELECT
                                                    document_quotes.*
                                                    FROM document_system
                                                    LEFT JOIN document_quotes on document_system.id = document_quotes.fk_documentId
                                                    WHERE document_system.documentId={row["documentId"]}
                                                    GROUP BY document_system.id;";

                            DataTable quotesInfo = crud.Select(quotesQuery, "document_system", syste);
                            string quotesRows = "";
                            foreach (DataRow item in quotesInfo.Rows)
                            {
                                quotesRows += $@"<tr>
                                                  <td colspan='3'>{item["quotes"]}</td>
                                                </tr>";
                            }

                            string quotesTable = $@"<tr>
                                                        <th colspan='3'>Quotes</th>
                                                    </tr>
                                                    {quotesRows}";
                            
                            body += $@"</tbody>
                                        {((quotesInfo.Rows.Count > 0) ? quotesTable : "")}
                                        </table>";
                           
                            string html = Properties.Resources.emailtemplate1;
                            html = html.Replace("{subject}", $"Special Bid {row["discountNumber"]}, {row["quote"]} - {row["countryName"]}");
                            html = html.Replace("{cuerpo}", $"Se le notifica la colocación de la Orden de Compra {row["documentId"]}, cuya información del Special Bid es:");
                            html = html.Replace("{contenido}", body);
                            string subject = $"Notificación Sistema de Documentos: Special Bid {row["discountNumber"]}, {row["quote"]} - {row["countryName"]}";
                            mail.SendHTMLMail(html, new string[] { root.f_sender }, subject, cc, null);
                        }
                    }

                    crud.Update($"UPDATE document_system SET statusRobot = 0 WHERE id = {row["id"]}", "document_system", syste);
                }
            }
        }

        private AmountResult getDMIBM(DataRow row)
        {
            AmountResult resultAmount = new AmountResult();
            string discountAmount = "";
            IWebDriver chrome = web.NewSeleniumChromeDriver();
            try
            {
                chrome.Navigate().GoToUrl("https://www.ibm.com/services/partners/epricer/v2/directLogin.do");
                new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("username")));
                chrome.FindElement(By.Id("username")).SendKeys("ePricer_CR@gbm.net");
                chrome.FindElement(By.Id("continue-button")).Click();
                new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("password")));

                chrome.FindElement(By.Id("password")).SendKeys("ePricer20");
                chrome.FindElement(By.Id("signinbutton")).Submit();

                bool auth = bsLogical.doubleAuth(chrome, row["documentId"].ToString());
                if (!auth)
                {
                    chrome.Close();
                    proc.KillProcess("chromedriver", true);
                    resultAmount.result = false;
                    resultAmount.discountAmount = "Error al tratar de realizar la doble autentificación en IBM";
                    return resultAmount;
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
                chrome.FindElement(By.XPath("//*[@id='ibm-columns-main']/div/div[2]/div/div/quotesearch-page/div/div[2]/div[1]/div[2]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div[8]/input")).SendKeys(row["discountNumber"].ToString());

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

                discountAmount = chrome.FindElement(By.XPath("//*[@id='+@@totalcom@@']/td[4]")).Text;

                chrome.Close();
                proc.KillProcess("chromedriver", true);
                discountAmount = discountAmount.Replace(" USD", "").Trim(); //.Replace(",", "").Replace(".", ",");
                bool up = crud.Update($"UPDATE document_system SET discountAmount = '{discountAmount}' WHERE id = '{row["id"]}'", "document_system", syste);

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                resultAmount.result = false;
                resultAmount.discountAmount = ex.Message;
                return resultAmount;
            }

            resultAmount.result = true;
            resultAmount.discountAmount = discountAmount;
            return resultAmount;
        }

        private AmountResult getDMLotus(DataRow row)
        {
            AmountResult resultAmount = new AmountResult();
            IWebDriver chrome = web.NewSeleniumChromeDriver(roots.FilesDownloadPath);
            string discountAmount = "";
            try
            {
                string ibm_user = "vaarrieta@gbm.net";
                string ibm_pass = cred.password_pdf_sb;

                #region Ingreso al website
                string urlLink = "https://www-112.ibm.com/software/howtobuy/passportadvantage/paoreseller/guidedselling/quotePartner.wss?jadeAction=DISPLAY_FIND_QUOTE_BY_NUM";
                chrome.Navigate().GoToUrl(urlLink);
                IJavaScriptExecutor jse = (IJavaScriptExecutor)chrome;
                #endregion


                #region Login In
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("username"))); }
                catch { }
                IWebElement User = chrome.FindElement(By.Id("username"));
                if (User.Displayed)
                {
                    User.SendKeys(ibm_user);
                    System.Threading.Thread.Sleep(1000);
                    IWebElement botoncontinuar = chrome.FindElement(By.XPath("//*[@id='continue-button']"));
                    botoncontinuar.Submit();
                    System.Threading.Thread.Sleep(1000);

                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("password"))); }
                    catch { }
                    IWebElement Pass = chrome.FindElement(By.Id("password"));
                    Pass.SendKeys(ibm_pass);
                    System.Threading.Thread.Sleep(1000);
                    IWebElement signinbutton = chrome.FindElement(By.XPath("//*[@id='signinbutton']"));
                    signinbutton.Submit();
                }

                #endregion

                bool auth = bsLogical.doubleAuth(chrome, row["documentId"].ToString());
                if (!auth)
                {
                    chrome.Close();
                    proc.KillProcess("chromedriver", true);
                    resultAmount.result = false;
                    resultAmount.discountAmount = "Error al tratar de realizar la doble autentificación en IBM";
                    return resultAmount;
                }

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 300)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='number']"))); }
                catch { }

                //aceptar las cookies
                try
                {
                    chrome.FindElement(By.XPath("//*[@id='truste-consent-button']")).Click();
                }
                catch (Exception)
                {

                }


                IWebElement sbnumber = chrome.FindElement(By.XPath("//*[@id='number']"));
                jse.ExecuteScript("arguments[0].value='" + row["discountNumber"].ToString() + "';", sbnumber);

                //buscar
                try
                {
                    chrome.FindElement(By.XPath("//*[@id='findByNumForm']/div/div[7]/input")).Click();
                }
                catch (Exception)
                {
                }
                System.Threading.Thread.Sleep(2000);

                discountAmount = "";

                try
                {
                    discountAmount = chrome.FindElement(By.XPath("//*[@id='findResultByNumForm']/div/table[2]/tbody/tr[1]/td[8]")).Text;

                }
                catch (Exception)
                {
                    console.WriteLine("No posee total value");
                }
                
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                MatchCollection matches = Regex.Matches(discountAmount, "USD");
                if (matches.Count >= 2)
                {
                    double totalAmount = 0.0;

                    // Use a regular expression to extract the amounts
                    MatchCollection amountMatches = Regex.Matches(discountAmount, @"\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?");

                    foreach (Match match in amountMatches)
                    {
                        double amount;
                        string val = match.Value.ToString().Trim().Replace(",", "").Replace(".", ",");
                        if (double.TryParse(val, out amount))
                        {
                            totalAmount += amount;
                        }
                    }
                    discountAmount = totalAmount.ToString();
                }
                else
                {
                    discountAmount = discountAmount.Replace(" USD", "").Trim(); //.Replace(",", "").Replace(".", ",");

                }
                bool up = crud.Update($"UPDATE document_system SET discountAmount = '{discountAmount}' WHERE id = '{row["id"]}'", "document_system", syste);

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                resultAmount.result = false;
                resultAmount.discountAmount = ex.Message;
                return resultAmount;
            }
            resultAmount.result = true;
            resultAmount.discountAmount = discountAmount;
            return resultAmount;
        }

        private AmountResult getDMLotusSap(DataRow row)
        {
            AmountResult resultAmount = new AmountResult();
            string discountAmount = "";
            try
            {
                string poStringTradings = row["poTradings"].ToString();
                string[] potradings = poStringTradings.Split(',');
                double netTotalValue = 0;
                foreach (string potrading in potradings)
                {
                    RfcDestination destErp = sap.GetDestRFC("ERP", mand);

                    IRfcFunction fmTableErp = destErp.Repository.CreateFunction("RFC_READ_TABLE");

                    fmTableErp.SetValue("QUERY_TABLE", "EKPO");
                    fmTableErp.SetValue("USE_ET_DATA_4_RETURN", "X");
                    //fmTableErp.SetValue("DELIMITER", "|");


                    IRfcTable fieldsErp = fmTableErp.GetTable("FIELDS");
                    fieldsErp.Append();
                    fieldsErp.SetValue("FIELDNAME", "NETWR");


                    IRfcTable optionTableErp = fmTableErp.GetTable("OPTIONS");
                    optionTableErp.Append();
                    optionTableErp.SetValue("TEXT", $"EBELN EQ '{potrading}'");


                    fmTableErp.Invoke(destErp);

                    DataTable reportErp = sap.GetDataTableFromRFCTable(fmTableErp.GetTable("ET_DATA"));

                    double netTotal = 0;
                    foreach (DataRow rowErp in reportErp.Rows)
                    {
                        string netValue = rowErp["LINE"].ToString().Replace(".", ""); //.Split(new char[] { '|' })[0].Trim();
                        netTotal += double.Parse(netValue, CultureInfo.GetCultureInfo("fr-FR"));
                    }

                    netTotalValue += netTotal;
                }
                discountAmount = netTotalValue.ToString("N2", System.Globalization.CultureInfo.InvariantCulture);
                bool up = crud.Update($"UPDATE document_system SET discountAmount = '{discountAmount}' WHERE id = '{row["id"]}'", "document_system", syste);
                
            }
            catch (Exception ex)
            {

                console.WriteLine(ex.Message);
                resultAmount.result = false;
                resultAmount.discountAmount = discountAmount;
                return resultAmount;
            }

            resultAmount.result = true;
            resultAmount.discountAmount = discountAmount;
            return resultAmount;
        }
       
        private AmountResult getDMCisco(DataRow row)
        {
            AmountResult resultAmount = new AmountResult();
            string discountAmount = "";
            try
            {
                string poStringTradings = row["poTradings"].ToString();
                string[] potradings = poStringTradings.Split(',');
                double netTotalValue = 0;
                foreach (string potrading in potradings)
                {
                    RfcDestination destErp = sap.GetDestRFC("ERP", mand);

                    IRfcFunction fmTableErp = destErp.Repository.CreateFunction("RFC_READ_TABLE");

                    fmTableErp.SetValue("QUERY_TABLE", "EKPO");
                    fmTableErp.SetValue("USE_ET_DATA_4_RETURN", "X");
                    //fmTableErp.SetValue("DELIMITER", "|");


                    IRfcTable fieldsErp = fmTableErp.GetTable("FIELDS");
                    fieldsErp.Append();
                    fieldsErp.SetValue("FIELDNAME", "NETWR");


                    IRfcTable optionTableErp = fmTableErp.GetTable("OPTIONS");
                    optionTableErp.Append();
                    optionTableErp.SetValue("TEXT", $"EBELN EQ '{potrading}'");


                    fmTableErp.Invoke(destErp);

                    DataTable reportErp = sap.GetDataTableFromRFCTable(fmTableErp.GetTable("ET_DATA"));

                    double netTotal = 0;
                    foreach (DataRow rowErp in reportErp.Rows)
                    {
                        string netValue = rowErp["LINE"].ToString().Replace(".", ""); //.Split(new char[] { '|' })[0].Trim();
                        netTotal += double.Parse(netValue, CultureInfo.GetCultureInfo("fr-FR"));
                    }

                    netTotalValue += netTotal;
                }
                //discountAmount = netTotalValue.ToString("F2").Replace(",", ".");
                discountAmount = netTotalValue.ToString("N2", System.Globalization.CultureInfo.InvariantCulture);
                bool up = crud.Update($"UPDATE document_system SET discountAmount = '{netTotalValue}' WHERE id = '{row["id"]}'", "document_system", syste);
            }
            catch (Exception ex)
            {

                console.WriteLine(ex.Message);
                resultAmount.result = false;
                resultAmount.discountAmount = ex.Message;
                return resultAmount;
            }

            resultAmount.result = true;
            resultAmount.discountAmount = discountAmount;
            return resultAmount;

        }
    }
}
public class AmountResult
{
    public bool result { get; set; }
    public string discountAmount { get; set; }
}