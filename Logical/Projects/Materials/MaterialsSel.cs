using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.App.Global;

namespace DataBotV5.Logical.Projects.Materials
{
    /// <summary>
    /// Clase Logical encargada de Materiales Sel.
    /// </summary>
    class MaterialsSel 
    {
        ProcessInteraction Proc = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        public string ExecuteBaw(string idBaw)
        {
            string resp = "";
            try
            {
                WebInteraction sel = new WebInteraction();
                IWebDriver chrome = sel.NewSeleniumChromeDriver(@"C:\Users\jearaya\Desktop"); //LA RUTA NO SE USA

                #region Ingreso al website

                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://prod-ihs-03.gbm.net/ProcessPortal/login.jsp"); //PRD
                #endregion

                chrome.FindElement(By.XPath("//*[@id='username']")).SendKeys(cred.username_SAPPRD.ToUpper()); //USER RPA
                chrome.FindElement(By.XPath("//*[@id='password']")).SendKeys(cred.password_baw);
                chrome.FindElement(By.XPath("/html/body/div/div[2]/div/div/div[1]/form/a")).Click();

                IWebElement Buscador = chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input"));
                IJavaScriptExecutor jse = (IJavaScriptExecutor)chrome;

                jse.ExecuteScript("arguments[0].click();", Buscador);
                Buscador = chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input"));
                try
                {
                    Buscador.SendKeys("*" + idBaw + "*"); //buscar
                    chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button")).Click();
                    System.Threading.Thread.Sleep(3000);
                }
                catch (Exception)
                {

                }

                try
                {
                    chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                    //wait para esperar ese elemento
                    //*[@id="button-button-okbutton"]
                    IWebElement topFrame = chrome.FindElement(By.XPath("//*[@id='div_1_2_1_3']/div/div[4]/iframe"));
                    chrome.SwitchTo().Frame(topFrame);

                    IWebElement Frame1 = chrome.FindElement(By.XPath("//*[@id='coach_frame38793499.cf1']"));
                    chrome.SwitchTo().Frame(Frame1);

                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='button-button-okbutton']"))); }
                    catch { }

                    System.Threading.Thread.Sleep(1000);
                    IWebElement botonfinal = chrome.FindElement(By.XPath("//*[@id='button-button-okbutton']"));
                    chrome.FindElement(By.XPath("//*[@id='button-button-okbutton']")).Click();
                }
                catch (Exception)
                {
                    // posible error
                }
                chrome.FindElement(By.XPath("//*[@id='div_1_1_1_1']/div/div/div[2]/div[2]/div[2]/a")).Click();

                chrome.Close();
                Proc.KillProcess("chromedriver",true);
                resp = "true";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }

            return resp;
        }
       
    }
}
