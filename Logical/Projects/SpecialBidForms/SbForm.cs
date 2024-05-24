using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.App.Global;

namespace DataBotV5.Logical.Projects.SpecialBidForms
{
    /// <summary>
    /// Clase Logical encargada de SbForm.
    /// </summary>
    class SbForm
    {
        ProcessInteraction proc = new ProcessInteraction();
        //[Obsolete]
        public string SbCreateFormWeb(string pais_ibm, string project, string useopp, string oppo, string usespecial,
                                       string prevbid, string priceupdate, string priceupjusti, string customer,
                                       string brand, string sedojusti, string soleprocur, string bpjusti,
                                             string swma, string renew, string totalprice, string customerprice)
        {

            Rooting roots = new Rooting();
            Credentials cred = new Credentials();
            WebInteraction sel = new WebInteraction();

            string ibm_user = "";
            string alerta_text = "";
            string ibm_pass = "";
            string contacto = "";
            string telf = "";
            string ibm_fecha = "";
            int id_pais_ibm = 0;
            string MES = "";
            string DIA = "";
            string sb_id = "";
            object fso = new object();
            string File_n = "";


            #region info general

            if (roots.BDUserCreatedBy.ToLower() == "vcamacho@gbm.net")
            {
                ibm_user = "vcamacho@gbm.net";
                ibm_pass = "IBMpass2019";
                contacto = "VICTOR CAMACHO";
                telf = "50625044500";
            }

            else if (roots.BDUserCreatedBy.ToLower() == "wocampo@gbm.net")
            {
                ibm_user = "wocampo@gbm.net";
                ibm_pass = "Romanos10:9";
                contacto = "WARREN OCAMPO";
                telf = "50625044500";
            }

            else if (roots.BDUserCreatedBy.ToLower() == "jmora@gbm.net")
            {
                ibm_user = "jmora@gbm.net";
                ibm_pass = "dominogbm6$";
                contacto = "JUAN FEDERICO MORA";
                telf = "50625044500";
            }

            else
            {
                ibm_user = "lfernandez@gbm.net";
                ibm_pass = "password";
                contacto = "LUIS FERNANDEZ";
                telf = "50625044582";
            }

            ibm_fecha = DateTime.Now.ToString("d");
            MES = DateTime.Now.Month.ToString();
            if (MES.Length == 1)
            {
                MES = "0" + MES;
            }


            DIA = DateTime.Now.Day.ToString();
            if (DIA.Length == 1)
            {
                DIA = "0" + DIA.Substring(0, 1);
            }

            //formato fecha DDMMYYYY
            ibm_fecha = DIA + "/" + MES + "/" + DateTime.Now.Year.ToString();

            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(roots.FilesDownloadPath);

            #region Ingreso al website
            chrome.Navigate().GoToUrl("https://bpms.podc.sl.edst.ibm.com/bpms/");
            //https://extbasicbpmsprd.podc.sl.edst.ibm.com/bpms/
            //https://bpms.podc.sl.edst.ibm.com/bpms/
            //https://extbasicbpmsprd.podc.sl.edst.ibm.com/bpms/
            System.Threading.Thread.Sleep(5000);
            //js executor para subir al inicio de pagina
            IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
            #endregion
            #region Formulario

            #region Login In

            
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("username"))); }
            catch { }
            IWebElement User = chrome.FindElement(By.Id("username"));
            if (User.Displayed)
            {
                User.SendKeys(ibm_user);
                System.Threading.Thread.Sleep(1000);
                IWebElement botoncontinuar = chrome.FindElement(By.XPath("//*[@id='continue-button']"));

                botoncontinuar.Submit();
                System.Threading.Thread.Sleep(1000);

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("password"))); }
                catch { }
                IWebElement Pass = chrome.FindElement(By.Id("password"));
                Pass.SendKeys(ibm_pass);
                System.Threading.Thread.Sleep(1000);
                IWebElement signinbutton = chrome.FindElement(By.XPath("//*[@id='signinbutton']"));
                signinbutton.Submit();
            }

            #endregion
            //System.Threading.Thread.Sleep(15000);
            System.Threading.Thread.Sleep(5000);

            #region seleccionar pais
            if (ibm_user == "vcamacho@gbm.net")
            {
                if (pais_ibm == "CR") { id_pais_ibm = 0; }
                else if (pais_ibm == "DO") { id_pais_ibm = 6; }
                else if (pais_ibm == "GT") { id_pais_ibm = 4; }
                else if (pais_ibm == "HN") { id_pais_ibm = 1; }
                else if (pais_ibm == "NI") { id_pais_ibm = 2; }
                else if (pais_ibm == "PA") { id_pais_ibm = 3; }
                else if (pais_ibm == "SV") { id_pais_ibm = 5; }
            }
            else
            {
                if (pais_ibm == "CR") { id_pais_ibm = 0; }
                else if (pais_ibm == "DO") { id_pais_ibm = 1; }
                else if (pais_ibm == "GT") { id_pais_ibm = 2; }
                else if (pais_ibm == "HN") { id_pais_ibm = 3; }
                else if (pais_ibm == "NI") { id_pais_ibm = 4; }
                else if (pais_ibm == "PA") { id_pais_ibm = 5; }
                else if (pais_ibm == "SV") { id_pais_ibm = 6; }
            }



            if (id_pais_ibm != 0)
            {
                try
                {
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 40)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(@"//*[@id=""ibm-universal-nav""]/div[3]/ul/li[1]/button"))); }
                    catch { }
                    IWebElement menupais = chrome.FindElement(By.XPath(@"//*[@id=""ibm-universal-nav""]/div[3]/ul/li[1]/button"));
                    //ibm-universal-nav"]/div[3]/ul/li[1]/button
                    //*[@id="ibm-universal-nav"]/div[3]/ul/li[1]/button
                    menupais.Click();
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(3000);
                    IWebElement menupais = chrome.FindElement(By.XPath(@"//*[@id=""ibm-universal-nav""]/div[3]/ul/li[1]/button"));
                    menupais.Click();
                }
                try
                {
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(@"//*[@id=""ibm-signin-minimenu-container""]/li[" + id_pais_ibm + @"]/a"))); }
                    catch { }
                    IWebElement menupais3 = chrome.FindElement(By.XPath(@"//*[@id=""ibm-signin-minimenu-container""]/li[" + id_pais_ibm + @"]/a"));
                    menupais3.Click();
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(2000);
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(@"//*[@id=""ibm-signin-minimenu-container""]/li[" + id_pais_ibm + @"]/a"))); }
                    catch { }
                    IWebElement menupais3 = chrome.FindElement(By.XPath(@"//*[@id=""ibm-signin-minimenu-container""]/li[" + id_pais_ibm + @"]/a"));
                    menupais3.Click();
                }

            }
            #endregion
            System.Threading.Thread.Sleep(1500);

            #region Crear SB
            try
            {
                chrome.FindElement(By.XPath("//*[@id='ibm-leadspace-body']/div[2]/p/a[2]")).Click();
            }
            catch (Exception)
            {
                IWebElement createSB = chrome.FindElement(By.XPath("//*[@id='ibm-content-wrapper']/nav/div/div[2]/ul/li[2]"));
                createSB.Click();
                System.Threading.Thread.Sleep(2000);
                IWebElement createSB2 = chrome.FindElement(By.XPath("//*[@id='ibm-content-wrapper']/nav/div/div[2]/ul/li[2]/ul/li[1]/a"));
                createSB2.Click();
            }
            System.Threading.Thread.Sleep(1000);
            IWebElement createSB3 = chrome.FindElement(By.XPath("//*[@id='btnAgreeCreateNewBid']"));
            createSB3.Click();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ibm-overlaywidget-17-content']/div/div[1]/div/p[3]/a[2]"))); }
            catch { }
            try
            {
                //sometimes pop ups a notifications
                chrome.FindElement(By.XPath("//*[@id='ibm-overlaywidget-17-content']/div/div[1]/div/p[3]/a[2]")).Submit();
            }
            catch (Exception)
            { }



            #endregion

            #region MAIN
            //*[@id='PropDescription']
            IWebElement PropDescription = chrome.FindElement(By.XPath("//input[@id='PropDescription']"));
            IJavaScriptExecutor jse = (IJavaScriptExecutor)chrome;
            jse.ExecuteScript("arguments[0].value='" + project + "';", PropDescription);
            System.Threading.Thread.Sleep(1000);
            if (useopp == "Yes")
            {
                IWebElement useOpportunity = chrome.FindElement(By.XPath("//*[@name='useOpportunity' and @value='Y']")); // and @value='Y'
                useOpportunity.Click();
                IWebElement useoppor = chrome.FindElement(By.XPath("/html/body/div[1]/main/div/div/div[2]/div/div[1]/div/div[1]/form/div[1]/div[5]/div[1]/p/span/label[1]/div/input"));
                try
                {
                    ((IJavaScriptExecutor)chrome).ExecuteScript("arguments[0].click();", useoppor);
                    useoppor.Click();
                }
                catch (Exception ex)
                {

                   
                }
          
                IWebElement useoppor2 = chrome.FindElement(By.Name("useOpportunity"));
                IWebElement OppCoce = chrome.FindElement(By.XPath("//*[@id='OppCoce']"));
                IJavaScriptExecutor jse2 = (IJavaScriptExecutor)chrome;
                jse2.ExecuteScript("arguments[0].value='" + oppo + "';", OppCoce);

                chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[5]/div[2]/p/span/a")).Click();


                try
                {
                    new WebDriverWait(chrome, new TimeSpan(0, 0, 15)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='infoFormStep1']/div[5]/div[3]/div/p/span/span/span[1]/span/span[2]")));
                    IWebElement ExcepCoce = chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[5]/div[3]/div/p/span/span/span[1]/span/span[2]"));
                    if (ExcepCoce.Displayed)
                    {
                        ExcepCoce.Click();
                        chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[2]")).Click();
                    }
                }
                catch { }

            }
            else
            {
                chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[5]/div[1]/p/span/label[2]/div")).Click();
                try
                {
                    new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='infoFormStep1']/div[5]/div[3]/div/p/span/span/span[1]/span/span[2]")));
                    IWebElement ExcepCoce = chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[5]/div[3]/div/p/span/span/span[1]/span/span[2]"));
                    if (ExcepCoce.Displayed)
                    {
                        ExcepCoce.Click();
                        chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[2]")).Click();
                    }
                }
                catch { }


            }


            if (usespecial == "Yes")
            {
                chrome.FindElement(By.XPath("//*[@id='REBD_Panel0']/div[1]/div[1]/p/span/label[1]/div")).Click();
                IWebElement REBD_BidNumber = chrome.FindElement(By.XPath("//*[@id='REBD_BidNumber']"));
                IJavaScriptExecutor jse3 = (IJavaScriptExecutor)chrome;
                jse3.ExecuteScript("arguments[0].value='" + prevbid + "';", REBD_BidNumber);
                System.Threading.Thread.Sleep(1000);

                chrome.FindElement(By.XPath("//*[@id='REBD_Panel1']/div/p/span/a")).Click();
                System.Threading.Thread.Sleep(1000);
                try
                {
                    IAlert alerta = chrome.SwitchTo().Alert();
                    alerta.Accept();
                }
                catch (Exception)
                { }

                if (priceupdate == "Yes")
                {
                    chrome.FindElement(By.XPath("//*[@id='REBD_Panel2']/div/div[1]/p/span/label[1]/div")).Click();
                    //chrome.FindElement(By.XPath("//*[@id='REBD_Justif']")).SendKeys(priceupjusti);
                    IWebElement REBD_Justif = chrome.FindElement(By.XPath("//*[@id='REBD_Justif']"));
                    IJavaScriptExecutor jse4 = (IJavaScriptExecutor)chrome;
                    jse4.ExecuteScript("arguments[0].value='" + priceupjusti + "';", REBD_Justif);

                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElementValue(REBD_Justif, priceupjusti)); }
                    catch { }
                    //System.Threading.Thread.Sleep(2000);
                }
                else { chrome.FindElement(By.XPath("//*[@id='REBD_Panel2']/div/div[1]/p/span/label[2]/div")).Click(); }
            }
            else { chrome.FindElement(By.XPath("//*[@id='REBD_Panel0']/div[1]/div[1]/p/span/label[2]/div")).Click(); }

            //select customer------------------------------------------------------------------------------------------------------------------


            chrome.FindElement(By.XPath("//*[@id='searchClientPanel']/div/div/p/span/span/input[2]")).Click();

            Actions escribir_cliente = new Actions(chrome);
            IAction tecleado = escribir_cliente.SendKeys(customer).Build();
            tecleado.Perform();
            try
            {
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='searchClientPanel']/div/div/p/span/span/div/div/li[1]"))); }
                catch { }
                try
                {
                    chrome.FindElement(By.XPath("//*[@id='searchClientPanel']/div/div/p/span/span/div/div/li[1]")).Click();
                }
                catch (Exception ex)
                {
                    return "Cliente no existe";
                }

            }
            catch (Exception)
            {
                chrome.FindElement(By.XPath("//*[@id='searchClientPanel']/div/div/p/span/span/input[2]")).Click();
                customer = "0" + customer;

                IWebElement ClientField = chrome.FindElement(By.XPath("//*[@id='SearchClientField']"));
                IJavaScriptExecutor jse5 = (IJavaScriptExecutor)chrome;
                jse5.ExecuteScript("arguments[0].value='" + "" + "';", ClientField);

                Actions escribir_cliente2 = new Actions(chrome);
                IAction tecleado2 = escribir_cliente2.SendKeys(customer).Build();
                tecleado2.Perform();
                try
                {
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='searchClientPanel']/div/div/p/span/span/div/div/li[1]"))); }
                    catch { }
                    chrome.FindElement(By.XPath("//*[@id='searchClientPanel']/div/div/p/span/span/div/div/li[1]")).Click();
                }
                catch (Exception)
                {
                    return "Cliente no existe";
                }
            }


            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='clientDetailHeader']/div[2]/div[1]/p/span/a"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='clientDetailHeader']/div[2]/div[1]/p/span/a")).Click(); //CLICK EN VALIDATE CMR

            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='clientDetailHeader']/div[2]/div[1]/p/span/span/div/div/li[1]"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='clientDetailHeader']/div[2]/div[1]/p/span/span/div/div/li[1]")).Click();


            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='clientDetail']/div[2]/div/p/span/label[1]/div"))); }
            catch { }

            chrome.FindElement(By.XPath("//*[@id='clientDetail']/div[2]/div/p/span/label[1]/div")).Click();
            chrome.FindElement(By.XPath("//*[@id='clientDetail']/div[3]/div[1]/p/span/label[2]/div")).Click();
            chrome.FindElement(By.XPath("//*[@id='clientDetail']/div[3]/div[2]/p/span/label[2]/div")).Click();


            IWebElement BPField = chrome.FindElement(By.XPath("//*[@id='SearchBPField']"));
            BPField.Click();
            Actions escribir_bp = new Actions(chrome);
            IAction tecla = escribir_bp.SendKeys("GBM").Build();
            tecla.Perform();


            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='searchBPPanel']/div/div/p/span/span/div/div/li[1]"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='searchBPPanel']/div/div/p/span/span/div/div/li[1]")).Click();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 120)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent()); }
            catch { }



            try
            {
                IAlert alerta = chrome.SwitchTo().Alert();
                alerta.Accept();
                IAlert alerta2 = chrome.SwitchTo().Alert();
                alerta2.Accept();
                IAlert alerta3 = chrome.SwitchTo().Alert();
                alerta3.Accept();
            }
            catch (Exception)
            { }

            //chrome.FindElement(By.XPath("//*[@id='BPDetails']/div[2]/div/p/span/a")).Click();

            //BP Type
            chrome.FindElement(By.XPath("//*[@id='t1Content']/div/div[1]/p[1]/span/span/span[1]/span/span[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[2]")).Click();

            //BP State
            chrome.FindElement(By.XPath("//*[@id='t1Content']/div/div[2]/p[1]/span/span/span[1]/span/span[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[2]")).Click();

            //chrome.FindElement(By.XPath("//*[@id='Contact']")).SendKeys(contacto);
            IWebElement Contact = chrome.FindElement(By.XPath("//*[@id='Contact']"));
            IJavaScriptExecutor jse7 = (IJavaScriptExecutor)chrome;
            jse7.ExecuteScript("arguments[0].value='" + contacto + "';", Contact);


            //chrome.FindElement(By.XPath("//*[@id='Phone_Number']")).SendKeys(telf);
            IWebElement Phone_Number = chrome.FindElement(By.XPath("//*[@id='Phone_Number']"));
            IJavaScriptExecutor jse8 = (IJavaScriptExecutor)chrome;
            jse8.ExecuteScript("arguments[0].value='" + telf + "';", Phone_Number);

            chrome.FindElement(By.XPath("//*[@id='t1Content']/div/div[1]/p[4]/span/div[2]")).Click();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='processTaRow']/span/div[2]"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='processTaRow']/span/div[2]")).Click();

            //chrome.FindElement(By.XPath("//*[@id='TransactionAgreement']")).SendKeys("N/A");
            IWebElement TransactionAgreement = chrome.FindElement(By.XPath("//*[@id='TransactionAgreement']"));
            IJavaScriptExecutor jse9 = (IJavaScriptExecutor)chrome;
            jse9.ExecuteScript("arguments[0].value='" + "N/A" + "';", TransactionAgreement);


            chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[14]/div[1]/p/span/div[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[14]/div[2]/p/span/span[2]/div")).Click();

            chrome.FindElement(By.XPath("//*[@id='infoFormStep1']/div[14]/div[3]/p/span/span/span[1]/span/span[2]")).Click();

            int ibrand = 1;

            if (brand == "INDUSTRY") { ibrand = 2; }
            else if (brand == "MAINFRAME") { ibrand = 3; }
            else if (brand == "POWER") { ibrand = 4; }
            else if (brand == "SOFTWARE") { ibrand = 5; }
            else if (brand == "STORAGE") { ibrand = 6; }
            else if (brand == "REMARKETING") { ibrand = 7; }
            else if (brand == "STORAGE_POWER") { ibrand = 8; }
            else if (brand == "POWER_STORAGE") { ibrand = 9; }
            else if (brand == "STORAGE_TIVOLI") { ibrand = 10; }
            else if (brand == "TIVOLI_STORAGE") { ibrand = 11; }
            else { ibrand = 0; }

            chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[" + ibrand + "]")).Click();
            chrome.FindElement(By.XPath("//*[@id='addlBrandOrSBDiv']/div[1]/p/span/div[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='labServices']/span/div[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='addlPricingQuestions']/div[2]/div[2]/p/span/div[2]")).Click();
            //save draft button

            IWebElement saveb = chrome.FindElement(By.XPath("//*[@id='btnSaveBid']"));
            chrome.FindElement(By.XPath("//*[@id='btnSaveBid']")).Click();

            try
            {
                IAlert alertasave = chrome.SwitchTo().Alert();
                alerta_text = alertasave.Text.ToString();
                if (alerta_text == "Error 500 - Internal Server Error")
                {
                    proc.KillProcess("chromedriver",true);
                    proc.KillProcess("chrome",true);
                    return alerta_text;
                }
                alertasave.Accept();
            }
            catch (Exception)
            { }
            //chrome.FindElement(By.XPath("//*[@id='proposalInfoForm']/p/a[4]")).Click();
            System.Threading.Thread.Sleep(5000);
            //try { new WebDriverWait(chrome, new TimeSpan(0, 0, 120)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElement(saveb, "Save as Draft")); }
            //catch { }
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='bidNumber']"))); }
            catch { }
            roots.id_special_bid = chrome.FindElement(By.XPath("//*[@id='bidNumber']")).Text;
            //sb_id = chrome.FindElement(By.XPath("//*[@id='ibm-leadspace-body']/div[2]/div[2]/div[1]/div/em[1]/span")).Text;


            #endregion

            #region transaction details
            new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='bidNumber']"))).Perform();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='step-selector-2']"))); }
            catch { }

            chrome.FindElement(By.XPath("//*[@id='step-selector-2']")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep2']/div[2]/div[1]/p/span[1]/div[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep2']/div[2]/div[2]/p[2]/span[1]/div[2]")).Click();

            //chrome.FindElement(By.XPath("//*[@id='RemunJustification']")).SendKeys(sedojusti);
            IWebElement RemunJustification = chrome.FindElement(By.XPath("//*[@id='RemunJustification']"));
            IJavaScriptExecutor jse10 = (IJavaScriptExecutor)chrome;
            jse10.ExecuteScript("arguments[0].value='" + sedojusti + "';", RemunJustification);

            //chrome.FindElement(By.XPath("//*[@id='SBRequest_Date']")).SendKeys(ibm_fecha);
            IWebElement SBRequest_Date = chrome.FindElement(By.XPath("//*[@id='SBRequest_Date']"));
            IJavaScriptExecutor jse11 = (IJavaScriptExecutor)chrome;
            jse11.ExecuteScript("arguments[0].value='" + ibm_fecha + "';", SBRequest_Date);

            chrome.FindElement(By.XPath("//*[@id='infoFormStep2']/div[8]/div[1]/p/span[1]/div[1]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep2']/div[8]/div[2]/p/span[1]/div[2]")).Click();

            #endregion

            #region additional questions
            new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='bidNumber']"))).Perform();
            try
            {
                new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='step-selector-3']")));
                IWebElement tab3 = chrome.FindElement(By.XPath("//*[@id='step-selector-3']"));
                if (tab3.Displayed)
                {
                    tab3.Click();
                    chrome.FindElement(By.XPath("//*[@id='divAddQuestions']/div[1]/div/p/span/span[1]/span/span[2]")).Click();
                    int isoleprocur = 1;

                    if (soleprocur == "Yes")
                    { isoleprocur = 2; }
                    else
                    { isoleprocur = 3; }

                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[" + isoleprocur + "]")).Click();

                    chrome.FindElement(By.XPath("//*[@id='divAddQuestions']/div[2]/div/p/span/span[1]/span/span[2]")).Click();
                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[3]")).Click();

                    chrome.FindElement(By.XPath("//*[@id='divAddQuestions']/div[3]/div/p/span/span[1]/span/span[2]")).Click();
                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[3]")).Click();

                    chrome.FindElement(By.XPath(" //*[@id='divAddQuestions']/div[4]/div/p/span/span[1]/span/span[2]")).Click();
                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[3]")).Click();

                    chrome.FindElement(By.XPath(" //*[@id='divAddQuestions']/div[5]/div/p/span/span[1]/span/span[2]")).Click();
                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[3]")).Click();

                    chrome.FindElement(By.XPath(" //*[@id='divAddQuestions']/div[6]/div/p/span/span[1]/span/span[2]")).Click();
                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[2]")).Click();

                    chrome.FindElement(By.XPath(" //*[@id='divAddQuestions']/div[7]/div/p/span/span[1]/span/span[2]")).Click();
                    chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[2]")).Click();

                    //chrome.FindElement(By.XPath("//*[@id='RA_Explanation']")).SendKeys("N/A");
                    IWebElement RA_Explanation = chrome.FindElement(By.XPath("//*[@id='RA_Explanation']"));
                    IJavaScriptExecutor jse12 = (IJavaScriptExecutor)chrome;
                    jse12.ExecuteScript("arguments[0].value='" + "N/A" + "';", RA_Explanation);


                }
            }
            catch { }

            #endregion

            #region business details

            new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='bidNumber']"))).Perform();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='step-selector-4']"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='step-selector-4']")).Click();

            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[2]/div[3]/p/span/div[2]")).Click();

            IWebElement BP_Justification = chrome.FindElement(By.XPath("//*[@id='BP_Justification']"));
            IJavaScriptExecutor jse13 = (IJavaScriptExecutor)chrome;
            jse13.ExecuteScript("arguments[0].value='" + bpjusti + "';", BP_Justification);
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 120)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElementValue(BP_Justification, bpjusti)); }
            catch { }

            IWebElement BPAC = chrome.FindElement(By.XPath("//*[@id='BPACInfo']"));
            IJavaScriptExecutor jse20 = (IJavaScriptExecutor)chrome;
            jse20.ExecuteScript("arguments[0].value='" + "" + "';", BPAC);

            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[8]/div[2]/p/span/span/span[1]/span/span[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='ibm-com']/span/span/span[2]/ul/li[4]")).Click();

            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[12]/div/p/span/div[1]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[14]/div/p/span/div[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[16]/div/p/span/div[2]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[18]/div/p/span/div[1]")).Click();
            chrome.FindElement(By.XPath("//*[@id='infoFormStep4']/div[20]/div/p/span/div[2]")).Click();


            #endregion

            #region review and submit

            new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='bidNumber']"))).Perform();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='step-selector-5']"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='step-selector-5']")).Click();

            chrome.FindElement(By.XPath("//*[@id='infoFormStep5']/p[2]/span[1]/label[1]/div")).Click();

            try
            {
                chrome.FindElement(By.XPath("//*[@id='btnSaveBid']")).Click();
                IWebElement botonsave = chrome.FindElement(By.XPath("//*[@id='btnSaveBid']"));
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 120)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElement(botonsave, "Save as Draft")); }
                catch { }
            }
            catch (Exception)
            { }

            #endregion

            #endregion
            #region Upload Products
            try
            {
                new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='ibm-leadspace-body']/div[1]/h1"))).Perform(); //subir arriba a la pagina
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='bpms-products-icon']"))); } //esperar que se vea el boton de Products
                catch { }
                chrome.FindElement(By.XPath("//*[@id='bpms-products-icon']")).Click(); //click en boton Products *[@id='ibm-primary-tabs']/ul/li[2]
                chrome.FindElement(By.XPath("//*[@id='productsButtons']/div/div/a[5]")).Click();
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='loadCFR']/p[2]/span/label[2]/div"))); }
                catch { }
                chrome.FindElement(By.XPath("//*[@id='loadCFR']/p[2]/span/label[2]/div")).Click();

                for (int i = 0; i <= roots.cfr_list.Length - 1; i++)
                {

                    IWebElement upload = chrome.FindElement(By.XPath("//*[@id='cfrInputFile']"));
                    chrome.FindElement(By.XPath("//*[@id='cfrInputFile']")).SendKeys(roots.cfr_list[i]);

                    chrome.FindElement(By.XPath("//*[@id='loadCFR']/p[5]/a")).Click();
                    try
                    {
                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 300)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent()); }
                        catch { }
                        IAlert alertasave = chrome.SwitchTo().Alert();
                        alerta_text = alerta_text + alertasave.Text.ToString() + "<br>";
                        alertasave.Accept();
                    }
                    catch (Exception)
                    { }
                    //System.Threading.Thread.Sleep(15000);
                }
                chrome.FindElement(By.XPath("//*[@id='loadCFR']/p[1]/a")).Click();
            }
            catch (Exception ex)
            {
                new ConsoleFormat().WriteLine(ex.ToString());

            }



            #endregion

            #region final step

            try
            { //a veces no sale el swma, por eso dentro del try
                new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='TotalRequestedPrice']"))).Perform();
                if (swma == "Yes")
                {
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='swma']/span/div[1]"))); }
                    catch { }
                    chrome.FindElement(By.XPath("//*[@id='swma']/span/div[1]")).Click();
                    if (renew == "Yes")
                    {
                        chrome.FindElement(By.XPath("//*[@id='renew']/span/div[1]")).Click();
                    }
                    else
                    {
                        chrome.FindElement(By.XPath("//*[@id='renew']/span/div[2]")).Click();
                    }
                }
                else
                {
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='swma']/span/div[2]"))); }
                    catch { }
                    chrome.FindElement(By.XPath("//*[@id='swma']/span/div[2]")).Click();

                }
            }
            catch (Exception)
            { }

            System.Threading.Thread.Sleep(10000);
            new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='TotalRecalculatedRequestedPrice']"))).Perform();
            
            IWebElement TotalRequestedPrice = chrome.FindElement(By.XPath("//*[@id='TotalRequestedPrice']"));
            IJavaScriptExecutor jse22 = (IJavaScriptExecutor)chrome;
            TotalRequestedPrice.Clear();
            jse22.ExecuteScript("arguments[0].value='" + totalprice + "';", TotalRequestedPrice);

            IWebElement EndCustomerPrice = chrome.FindElement(By.XPath("//*[@id='EndCustomerPrice']"));
            IJavaScriptExecutor jse21 = (IJavaScriptExecutor)chrome;
            EndCustomerPrice.Clear();
            jse21.ExecuteScript("arguments[0].value='" + customerprice + "';", EndCustomerPrice);

            #region ir a Review y guardar
            try
            {
                new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='bidNumber']"))).Perform(); //ir arriba
                chrome.FindElement(By.XPath("//*[@id='proposalInfoTab']")).Click(); //tab de Info

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='step-selector-5']"))); }
                catch { }
                chrome.FindElement(By.XPath("//*[@id='step-selector-5']")).Click();
                
                chrome.FindElement(By.XPath("//*[@id='btnSaveBid']")).Click();
                IWebElement botonsave2 = chrome.FindElement(By.XPath("//*[@id='btnSaveBid']"));
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 120)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElement(botonsave2, "Save as Draft")); }
                catch { }
            }
            catch (Exception)
            { }
            #endregion
            //System.Threading.Thread.Sleep(10000);
            #endregion
            #region log out
            chrome.Navigate().GoToUrl("https://myibm.ibm.com/pkmslogout?filename=accountRedir.html");
            //System.Threading.Thread.Sleep(5000);
            #endregion

            System.Threading.Thread.Sleep(1000);
            chrome.Close();
            proc.KillProcess("chromedriver",true);
            return alerta_text;
        }
        ~SbForm()
        {

        }

    }
}