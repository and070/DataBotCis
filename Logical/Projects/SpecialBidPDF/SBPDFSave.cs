using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using DataBotV5.Data;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Logical.Projects.BusinessSystem;

namespace DataBotV5.Logical.Projects.SpecialBidPDF
{
    /// <summary>
    /// Clase Logical encargada de guardar PDF en SB.
    /// </summary>
    class SBPDFSave
    {
        ProcessInteraction proc = new ProcessInteraction();
        public string SavePdfSb(string sbId, string solicitante, string subject)
        {
            Rooting roots = new Rooting();
            //SpecialBid_Form sb = new SpecialBid_Form();
            Credentials cred = new Credentials();
            MailInteraction mail = new MailInteraction();
            WebInteraction sel = new WebInteraction();
            ConsoleFormat console = new ConsoleFormat();
            string ibm_user = "";
            string ibm_pass = "";
            string respuesta = "";

            ibm_user = "vaarrieta@gbm.net";
            ibm_pass = cred.password_pdf_sb;

            IWebDriver chrome = sel.NewSeleniumChromeDriver(roots.FilesDownloadPath);
            //IWebDriver chrome = sel.NewSeleniumFirefoxDriver(roots.FilesDownloadPath);
            try
            {



                #region Ingreso al website
                string url_link = "https://www-112.ibm.com/software/howtobuy/passportadvantage/paoreseller/guidedselling/quotePartner.wss?jadeAction=DISPLAY_FIND_QUOTE_BY_NUM";
                chrome.Navigate().GoToUrl(url_link);
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

                #region Doble autentificacion
                BusinessSystemLogical bsLogical = new BusinessSystemLogical();
                bool auth = bsLogical.doubleAuth(chrome, sbId);
                if (!auth)
                {
                    chrome.Close();
                    proc.KillProcess("chromedriver", true);
                    return ("Error: al autentificar en el portal");
                }
                #endregion

                #region buscar ID

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
                jse.ExecuteScript("arguments[0].value='" + sbId + "';", sbnumber);

                //buscar
                try
                {
                    chrome.FindElement(By.XPath("//*[@id='findByNumForm']/div/div[7]/input")).Click();
                }
                catch (Exception)
                {
                }
                System.Threading.Thread.Sleep(2000);

                //view  status details 
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 300)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ibm-footer-module-links']/div[1]/ul/li[1]/a"))); }
                catch { }

                try { new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='ibm-footer-module-links']/div[1]/ul/li[1]/a"))).Perform(); }
                catch { }

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='findResultByNumForm']/div/table[2]/tbody/tr[7]/td/a"))); }
                catch { }
                try
                {

                    chrome.FindElement(By.XPath("//*[@id='findResultByNumForm']/div/table[2]/tbody/tr[7]/td/a")).Click();
                }
                catch (Exception ex)
                {
                }

                System.Threading.Thread.Sleep(2000);

                //download now

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 300)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ibm-content-main']/div[2]/div[1]/div[4]/div/h2"))); }
                catch { }

                try { new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='ibm-content-main']/div[2]/div[1]/div[4]/div/h2"))).Perform(); }
                catch { }

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ibm-content-main']/div[2]/div[1]/div[3]/div/table/tbody/tr[2]/td[3]"))); }
                catch { }
                string exist = chrome.FindElement(By.XPath("//*[@id='ibm-content-main']/div[2]/div[1]/div[3]/div/table/tbody/tr[2]/td[3]")).Text;
                if (exist == "Pending processing")
                {
                    return "No existe PDF";
                }
                //*[@id='ibm-content-main']/div[2]/div[1]/div[3]/div/table/tbody/tr[2]/td[3]

                IWebElement href_download = chrome.FindElement(By.XPath("//*[@id='ibm-content-main']/div[2]/div[1]/div[3]/div/table/tbody/tr[2]/td[3]/a"));

                if (href_download.GetAttribute("href").Length > 0)
                {
                    string statico = "www-112.ibm.com/software/howtobuy/passportadvantage/paoreseller/guidedselling/";

                    var descargar_link = chrome.FindElement(By.XPath("//*[@id='ibm-content-main']/div[2]/div[1]/div[3]/div/table/tbody/tr[2]/td[3]/a")).GetAttribute("href");

                    chrome.Navigate().GoToUrl(descargar_link);

                    for (var x = 0; x < 40; x++)
                    {
                        if (File.Exists(roots.FilesDownloadPath + "\\" + "Channel Bid Notification PDF " + sbId + ".pdf"))
                        { break; }
                        if (File.Exists(roots.FilesDownloadPath + "\\" + "customer quote PDF " + sbId + ".pdf"))
                        { break; }
                        System.Threading.Thread.Sleep(1000);
                    }

                }
                else
                {
                    respuesta = "error al descargar el PDF";
                }

                #endregion

                #region log off
                chrome.Navigate().GoToUrl("https://myibm.ibm.com/pkmslogout?filename=accountRedir.html");
                #endregion

                System.Threading.Thread.Sleep(1000);
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                return "ok";

            }
            catch (Exception)
            {
                try
                {
                    string pathErrors = roots.FilesDownloadPath + @"\";
                    Screenshot TakeScreenshot = ((ITakesScreenshot)chrome).GetScreenshot();
                    TakeScreenshot.SaveAsFile(pathErrors + $"errorDownloadSpecialBidPdf.png");
                }
                catch (Exception i)
                {
                }
                return "error";
            }
        }

        ~SBPDFSave()
        {

        }
    }
}
