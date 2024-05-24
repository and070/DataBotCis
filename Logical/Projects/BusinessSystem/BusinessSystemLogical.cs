using DataBotV5.App.ConsoleApp;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Mail;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Logical.Projects.BusinessSystem
{
    public class BusinessSystemLogical
    {
        public bool doubleAuth(IWebDriver chrome, string documentId)
        {
            MailInteraction mail = new MailInteraction();
            ConsoleFormat console = new ConsoleFormat();
            Rooting roots = new Rooting();
            #region Doble autentificacion
            try
            {
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div[4]/div/ul"))); }
                catch
                {
                    return true; //significa que no pidio la doble auth
                }
                IWebElement emailCodeList = chrome.FindElement(By.XPath("/html/body/div/div[4]/div/ul"));
                IList<IWebElement> LiCollection = emailCodeList.FindElements(By.TagName("li"));
                //por cada fila de la tabla
                bool find = false;
                foreach (IWebElement element in LiCollection)
                {
                    //IList<IWebElement> tdCollection;

                    IWebElement name = element.FindElement(By.TagName("div"));
                    string emailMand = (Start.enviroment == "QAS") ? "Email dat******@gbm.net" : "Email dat****@gbm.net";
                    if (name.Text == emailMand)
                    {
                        IWebElement link = element.FindElement(By.TagName("a"));
                        link.Click();
                        find = true;
                        break;
                    }

                }
                if (!find)
                {
                    mail.SendHTMLMail("Error: no se pudo encontrar el correo del databot en la lista de verificadores en el portal de IBM", new string[] { "dmeza@gbm.net" }, $"Error en la doble autentificación de IBM Portal del documento {documentId}", new string[] { "dmeza@gbm.net" });
                    return false;
                }
                int co = 0;
                string code = "";
                bool done = false;
                while (!done)
                {
                    mail.GetAttachmentEmail("IBM Verifications", "Procesados", "Procesados IBM Verifications");
                    if (roots.Email_Body != "")
                    {

                        string body = roots.Email_Body;
                        if (body.Contains("Please use the following verification code:"))
                        {
                            code = body.Split(new string[] { "Please use the following verification code:\r\n\r\n" }, StringSplitOptions.None)[1].Split(new string[] { "\r" }, StringSplitOptions.None)[0].Split(new string[] { "-" }, StringSplitOptions.None)[1];
                        }
                        else
                        {

                            code = body.Split(new string[] { "Utilice el siguiente código de verificación:\r\n\r\n" }, StringSplitOptions.None)[1];
                        }
                        done = true;
                        break;
                    }
                    System.Threading.Thread.Sleep(1000);
                    co++;
                    if (co == 300)
                    {
                        break;
                    }
                }
                if (code == "")
                {
                    //nunca llego el email de codigo
                    mail.SendHTMLMail("Error: no se pudo obtener el email de validación por parte de IBM", new string[] { "dmeza@gbm.net" }, $"Error en la doble autentificación de IBM Portal del documento {documentId}", new string[] { "dmeza@gbm.net" });
                    return false;


                }
                else
                {
                    //agregar correo y seguir
                    chrome.FindElement(By.XPath("//*[@id='otp']")).SendKeys(code);
                    chrome.FindElement(By.XPath("/html/body/div[2]/div/form[1]/p/button")).Click();
                    System.Threading.Thread.Sleep(2000);
                    //valida si el codigo dio error:
                    string ValError = "";
                    try
                    {
                        new WebDriverWait(chrome, new TimeSpan(0, 0, 20)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[2]/div/form[1]/div/div[2]")));
                        //dio error 
                        //enviar email de error
                        ValError = chrome.FindElement(By.XPath("/html/body/div/div/div[2]/div[2]/span")).Text;
                    }
                    catch { }
                    try
                    {
                        new WebDriverWait(chrome, new TimeSpan(0, 0, 20)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]/div[2]/span")));
                        //dio error 
                        //enviar email de error
                        ValError = chrome.FindElement(By.XPath("/html/body/div/div/div[2]/div[2]/span")).Text;
                    }
                    catch { }
                    if (ValError != "")
                    {
                        mail.SendHTMLMail($"Error: el código {code} no funcionó para la doble autentificación de IBM", new string[] { "dmeza@gbm.net" }, $"Error en la doble autentificación de IBM Portal del documento {documentId}", new string[] { "dmeza@gbm.net" });
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                return false;
            }

            return true;

            #endregion
        }

    }
}
