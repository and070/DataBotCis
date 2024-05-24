using OpenQA.Selenium.Interactions;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Web;
using OpenQA.Selenium;
using System;
using System.Data;

namespace DataBotV5.Logical.Projects.ControlDesk
{
    /// <summary>
    /// Clase Logical encargada de interacciones de Selenium en Control Desk.
    /// </summary>
    class ControlDeskAprovalContract
    {
        ProcessInteraction proc = new ProcessInteraction();
        Credentials cred = new Credentials();

        private IWebDriver chrome;
        private void LoginCD(string usernameCD, string passwordCD)
        {
            try
            {
                chrome.FindElement(By.Id("j_username")).SendKeys(usernameCD);
                chrome.FindElement(By.Id("j_password")).SendKeys(passwordCD);
                chrome.FindElement(By.Id("loginbutton")).Click();
            }
            catch (Exception)
            {
                chrome.FindElement(By.Id("loginbutton")).Submit();
            }
        }
        private void WaitSystemMessage(int timeout)   //ojo que no sirve en la versión vieja de CD
        {
            System.Threading.Thread.Sleep(500);
            for (int i = 0; i < (5 * timeout); i++)
            {
                try
                {
                    //significa que ya se cerro la ventana de espera
                    if (chrome.FindElement(By.Id("mb_msg")).Text != "")
                    {
                        return;
                    }
                }
                catch (NoSuchElementException) { }

                System.Threading.Thread.Sleep(200);
                if (i == (5 * timeout))
                {
                    throw new System.TimeoutException("Se espero el tiempo máximo");
                }
            }
        }
        private void WaitLongLoading(int timeout)
        {
            System.Threading.Thread.Sleep(500);
            for (int i = 0; i < (5 * timeout); i++)
            {
                try
                {
                    //significa que ya se cerro la ventana de espera
                    string long1;
                    try
                    {
                        long1 = chrome.FindElement(By.Id("longopwait-lb")).Text;
                    }
                    catch (NoSuchElementException)
                    {
                        long1 = "";
                    }
                    if (!long1.Contains("Please wait"))
                    {
                        return;
                    }
                }
                catch (NoSuchElementException) { }

                System.Threading.Thread.Sleep(200);
                if (i == (5 * timeout))
                {
                    throw new System.TimeoutException("Se espero el tiempo máximo");
                }
            }
        }
        private void WaitCursorToClose(int timeout)
        {
            System.Threading.Thread.Sleep(500);
            for (int i = 0; i < (5 * timeout); i++)
            {
                string eval = chrome.FindElement(By.Id("wait")).GetAttribute("style"); //sacar la propiedad style del elemento

                if (eval.Contains("display: none"))
                {
                    //significa que ya se cerro la ventana de espera
                    return;
                }
                System.Threading.Thread.Sleep(200);
                if (i == (5 * timeout))
                {
                    throw new System.TimeoutException("Se espero el tiempo máximo");
                }
            }
        }
        private void CdIni(string link, bool automatic = true)
        {
            WebInteraction sel = new WebInteraction();
            chrome = sel.NewSeleniumChromeDriver();
            chrome.Navigate().GoToUrl(link);
        }
        private void CloseCD()
        {
            System.Threading.Thread.Sleep(1000);
            try { chrome.Close(); } catch (Exception) { }
            chrome.Quit();
            proc.KillProcess("chromedriver", true);
        }
        private void SearchRegister(string RegisterId)
        {
            string sys_msg = "";
            try { chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); } catch (Exception) { }
            WaitCursorToClose(30);

            chrome.FindElement(By.Id("quicksearch")).Click();
            chrome.FindElement(By.Id("quicksearch")).SendKeys(RegisterId);
            if (chrome.FindElement(By.Id("quicksearch")).GetAttribute("prekeyvalue") == "")
            {
                chrome.FindElement(By.Id("quicksearch")).SendKeys(RegisterId);
            }
            chrome.FindElement(By.Id("quicksearchQSImage")).Click();

            WaitCursorToClose(30);

            try { sys_msg += chrome.FindElement(By.Id("mb_msg")).Text; } catch (NoSuchElementException) { }

            if (sys_msg.Contains("No records were found"))
            {
                try { chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); } catch (Exception) { }
                WaitCursorToClose(30);
            }
        }

        //Métodos Públicos

        public string ChangeStatusCustomerAgreements(List<string> contracts, string link, string action)
        {
            string sysMsg = "";
            string query = "status = 'noexiste'";
            if (action == "appr")
            {
                action = "menu0_APPR_OPTION_a";
                query = "STATUS in ('DRAFT', 'PNDREV') and agreement in ('" + String.Join("','", contracts.ToArray()) + "')";
            }
            else if (action == "close")
            {
                action = "menu0_CLOSE_OPTION_a";
                query = "STATUS in ('APPR','DRAFT') and agreement in ('" + String.Join("','", contracts.ToArray()) + "')";
            }

            link = link.Replace("meaweb/os/", "") + "customers/ui/?event=loadapp&value=pluspagree";

            if (contracts.Count < 201)
            {
                CdIni(link);
                LoginCD(cred.username_CD, cred.password_CD);

                chrome.FindElement(By.Id("quicksearchQSMenuImage")).Click();
                WaitCursorToClose(30);
                chrome.FindElement(By.Id("menu0_SEARCHWHER_OPTION")).Click();
                WaitCursorToClose(30);
                chrome.FindElement(By.Id("m8366b731-ta")).Clear();
                chrome.FindElement(By.Id("m8366b731-ta")).SendKeys(query);
                chrome.FindElement(By.Id("m81200968-pb")).Click(); //find
                WaitCursorToClose(30);

                try { sysMsg = chrome.FindElement(By.Id("mb_msg")).Text; } catch (NoSuchElementException) { }

                if (!sysMsg.Contains("No records were found"))
                {
                    if (chrome.PageSource.Contains("Billing Schedule"))//significa solo apareció un contrato
                        chrome.FindElement(By.Id("m397b0593-tabs_middle")).Click();

                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id("md86fe08f_ns_menu_STATUS_OPTION_a")).Click(); //change status
                    WaitCursorToClose(30);


                    chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); //system msg que va a afectar todo

                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id("mba64a966-tb")).Click(); //drop down
                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id(action)).Click(); //opcion Approved
                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id("mb0118d33-pb")).Click(); //ok

                    WaitCursorToClose(30);
                    WaitSystemMessage(30);
                    sysMsg = chrome.FindElement(By.Id("mb_msg")).Text;
                    try { chrome.FindElement(By.Id("m15f1c9f0-pb")).Click(); } catch (Exception) { }//close   

                }
                CloseCD();
            }
            else
                throw new FormatException("La lista es mayor a 200 contratos");
            return sysMsg;
        }
        public void RunEscalationCD(string query, string escalation, string link)
        {
            IJavaScriptExecutor jse = (IJavaScriptExecutor)chrome;

            link = link.Replace("meaweb/os/", "") + "customers/ui/?event=loadapp&value=escalation";

            CdIni(link);
            LoginCD(cred.username_CD, cred.password_CD);
            #region Ingreso al website
            chrome.Navigate().GoToUrl(link);
            #endregion
            #region entrar al escalation
            chrome.FindElement(By.XPath("//*[@id='m6a7dfd2f_tfrow_[C:1]_txt-tb']")).SendKeys(escalation);
            chrome.FindElement(By.XPath("//*[@id='m6a7dfd2f-hb_header_5']")).Click();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='m6a7dfd2f_tdrow_[C:1]-c[R:0]']"))); } catch { }
            chrome.FindElement(By.XPath("//*[@id='m6a7dfd2f_tdrow_[C:1]-c[R:0]']")).Click();
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='m2066da58-ta']"))); } catch { }
            #endregion
            #region Poner Datos
            try
            {
                chrome.FindElement(By.XPath("//*[@id='m2066da58-ta']")).Click();
                chrome.FindElement(By.XPath("//*[@id='m2066da58-ta']")).Clear();
            }
            catch (Exception)
            {
                chrome.FindElement(By.XPath("//*[@id='m74daaf83_ns_menu_ADESC_OPTION_a_tnode']")).Click();
                System.Threading.Thread.Sleep(1000);
                chrome.FindElement(By.XPath("//*[@id='m2066da58-ta']")).Click();
                chrome.FindElement(By.XPath("//*[@id='m2066da58-ta']")).Clear();
            }


            IWebElement aprobacion = chrome.FindElement(By.XPath("//*[@id='m2066da58-ta']"));
            aprobacion.SendKeys(query);
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 120)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElementValue(aprobacion, query)); }
            catch { }
            #endregion
            #region Activar y esperar
            string fecha1 = chrome.FindElement(By.XPath("//*[@id='m3a5e2556-tb']")).GetAttribute("title").ToString();
            chrome.FindElement(By.XPath("//*[@id='m74daaf83_ns_menu_ADESC_OPTION_a_tnode']")).Click();
            try
            {
                System.Threading.Thread.Sleep(120000);
                chrome.FindElement(By.XPath("//*[@id='m74daaf83_ns_menu_ADESC_OPTION_a_tnode']")).Click();
            }
            catch (Exception)
            { chrome.FindElement(By.XPath("//*[@id='m74daaf83_ns_menu_ADESC_OPTION_a_tnode']")).Click(); }

            string fecha2 = chrome.FindElement(By.XPath("//*[@id='m3a5e2556-tb']")).GetAttribute("title").ToString();
            if (fecha1 != "" && fecha2 != "")
            {
                if (fecha1 == fecha2)
                {
                    chrome.FindElement(By.XPath("//*[@id='m74daaf83_ns_menu_ADESC_OPTION_a_tnode']")).Click();
                    System.Threading.Thread.Sleep(120000);
                    chrome.FindElement(By.XPath("//*[@id='m74daaf83_ns_menu_ADESC_OPTION_a_tnode']")).Click();
                }
            }

            #endregion
            #region log off
            Actions logout = new Actions(chrome);
            logout.KeyDown(Keys.Alt).SendKeys("S").Perform();
            #endregion
            CloseCD();
        }
        public string ChangeStatusResponsePlans(List<string> responsePlans, string link, string action)
        {
            bool singleRecord = false;
            string sysMsg = "";
            string query = "status = 'noexiste'";

            if (action.ToLower() == "active")
            {
                action = "menu0_ACTIVE_OPTION";
                query = "sanum in ('" + String.Join("','", responsePlans.ToArray()) + "')";
            }
            else if (action.ToLower() == "inactive")
            {
                action = "menu0_INACTIVE_OPTION";
                query = "sanum in ('" + String.Join("','", responsePlans.ToArray()) + "')";

            }

            link = link.Replace("meaweb/os/", "") + "customers/ui/?event=loadapp&value=pluspresp";

            CdIni(link);
            LoginCD(cred.username_CD, cred.password_CD);

            chrome.FindElement(By.Id("quicksearchQSMenuImage")).Click();
            WaitCursorToClose(30);
            chrome.FindElement(By.Id("menu0_SEARCHWHER_OPTION")).Click();
            WaitCursorToClose(30);
            chrome.FindElement(By.Id("m8366b731-ta")).Clear();
            chrome.FindElement(By.Id("m8366b731-ta")).SendKeys(query);
            chrome.FindElement(By.Id("m81200968-pb")).Click(); //find
            WaitCursorToClose(30);

            try { sysMsg = chrome.FindElement(By.Id("mb_msg")).Text; } catch (NoSuchElementException) { }

            if (!sysMsg.Contains("No records were found"))
            {
                if (chrome.PageSource.Contains("Response Actions"))//significa solo apareció un contrato
                {
                    chrome.FindElement(By.Id("m397b0593-tabs_middle")).Click();
                    singleRecord = true;
                }

                WaitCursorToClose(30);
                chrome.FindElement(By.Id("md86fe08f_ns_menu_STATUS_OPTION_a")).Click(); //change status
                WaitCursorToClose(30);


                chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); //system msg que va a afectar todo

                WaitCursorToClose(30);
                chrome.FindElement(By.Id("m632d0b90-tb")).Click(); //drop down
                WaitCursorToClose(30);
                chrome.FindElement(By.Id(action)).Click(); //opcion Active
                WaitCursorToClose(30);
                chrome.FindElement(By.Id("mb0118d33-pb")).Click(); //ok

                WaitCursorToClose(30);
                WaitSystemMessage(30);
                if (singleRecord)
                    sysMsg = chrome.FindElement(By.Id("titlebar_error")).GetAttribute("innerText");
                else
                    sysMsg = chrome.FindElement(By.Id("mb_msg")).Text;

                try { chrome.FindElement(By.Id("m15f1c9f0-pb")).Click(); } catch (Exception) { }//close   

            }
            CloseCD();

            return sysMsg;
        }
        public string CreateSlaEscalation(CdSlaData sla, string sch, string link)
        {
            ControlDeskInteraction cdi = new ControlDeskInteraction();
            string sysMsg = "ERROR";

            try
            {
                DataTable escalations = new DataTable();
                escalations.Columns.Add("ElapsedTimeAttribute");
                escalations.Columns.Add("ElapsedTimeInterval");
                escalations.Columns.Add("EscalationPointCondition");
                escalations.Columns.Add("CommTemplate");
                link += "/customers/ui/?event=loadapp&value=sla";

                #region Ir al registro
                CdIni(link);
                LoginCD(cred.username_CD, cred.password_CD);
                SearchRegister(sla.Sanum);
                chrome.FindElement(By.Id("m8ee1358-tb")).Click();
                #endregion

                #region Llenar schedule
                Actions scrollDown = new Actions(chrome);
                scrollDown.KeyDown(Keys.Control).SendKeys(Keys.End).Perform();
                scrollDown.KeyDown(Keys.Control).SendKeys(Keys.End).Perform();
                scrollDown.KeyDown(Keys.Control).SendKeys(Keys.End).Perform();

                chrome.FindElement(By.Id("meb931f6d_tdrow_[C:6]_hyperlink-lb[R:0]_image")).Click();
                WaitCursorToClose(30);

                sla.Escalation = chrome.FindElement(By.Id("m99b3cb0c-tb")).GetAttribute("value");

                //save
                chrome.FindElement(By.Id("toolactions_SAVE-tbb_image")).Click();
                WaitCursorToClose(30);
                try { sysMsg = chrome.FindElement(By.Id("titlebar_error")).GetAttribute("innerText").Trim(); } catch (NoSuchElementException) { }
                
                if (sysMsg.Contains("Record has been saved"))
                {
                    if (cdi.CreateSlaEscalation(sla, sch) == "OK")
                    {
                        #region Activar SLA
                        chrome.FindElement(By.Id("md86fe08f_ns_menu_STATUS_OPTION_a_tnode")).Click(); //change status
                        WaitCursorToClose(30);

                        try { sysMsg += chrome.FindElement(By.Id("mb_msg")).Text; } catch (NoSuchElementException) { }

                        if (sysMsg != "")
                        {
                            try { chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); } catch (Exception) { }
                            WaitCursorToClose(30);
                        }

                        chrome.FindElement(By.Id("m24f0575e-tb")).Click();
                        WaitCursorToClose(30);
                        chrome.FindElement(By.Id("menu0_ACTIVE_OPTION")).Click();
                        WaitCursorToClose(30);
                        chrome.FindElement(By.Id("mea5c4b9c-pb")).Click(); //ACTIVAR
                        WaitCursorToClose(30);
                        sysMsg = chrome.FindElement(By.Id("titlebar_error")).GetAttribute("innerText");
                        CloseCD();
                        #endregion
                    }
                    else
                        sysMsg = "ERROR: " + sysMsg;
                }
                else
                    sysMsg = "ERROR: " + sysMsg;
                #endregion

                CloseCD();
            }
            catch (Exception ex)
            {
                sysMsg = "ERROR: " + ex.Message;
                CloseCD();
            }
            return sysMsg;
        }

        internal string ActivateSla(CdSlaData sla, string urlCd)
        {
            string sysMsg = "ERROR";
            urlCd += "/customers/ui/?event=loadapp&value=sla";

            #region Ir al registro
            CdIni(urlCd);
            LoginCD(cred.username_CD, cred.password_CD);
            SearchRegister(sla.Sanum);
            chrome.FindElement(By.Id("m8ee1358-tb")).Click();
            #endregion

            #region Activar SLA
            chrome.FindElement(By.Id("md86fe08f_ns_menu_STATUS_OPTION_a_tnode")).Click(); //change status
            WaitCursorToClose(30);

            try { sysMsg += chrome.FindElement(By.Id("mb_msg")).Text; } catch (NoSuchElementException) { }

            if (sysMsg != "")
            {
                try { chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); } catch (Exception) { }
                WaitCursorToClose(30);
            }

            chrome.FindElement(By.Id("m24f0575e-tb")).Click();
            WaitCursorToClose(30);
            chrome.FindElement(By.Id("menu0_ACTIVE_OPTION")).Click();
            WaitCursorToClose(30);
            chrome.FindElement(By.Id("mea5c4b9c-pb")).Click(); //ACTIVAR
            WaitCursorToClose(30);
            sysMsg = chrome.FindElement(By.Id("titlebar_error")).GetAttribute("innerText");
            CloseCD();
            #endregion

            return sysMsg;
        }

        public string ActivateCommunicationTemplates(List<string> templates, string link)
        {
            string sysMsg = "ERROR";
            string action = "menu0_ACTIVE_OPTION";

            link += "/customers/ui/?event=loadapp&value=commtmplt";

            string query = "templateid in ('" + String.Join("','", templates.ToArray()) + "')";

            try
            {
                CdIni(link);
                LoginCD(cred.username_CD, cred.password_CD);

                chrome.FindElement(By.Id("quicksearchQSMenuImage")).Click();
                WaitCursorToClose(30);
                chrome.FindElement(By.Id("menu0_SEARCHWHER_OPTION")).Click();
                WaitCursorToClose(30);
                chrome.FindElement(By.Id("m8366b731-ta")).Clear();
                chrome.FindElement(By.Id("m8366b731-ta")).SendKeys(query);
                chrome.FindElement(By.Id("m81200968-pb")).Click(); //find
                WaitCursorToClose(30);

                try { sysMsg = chrome.FindElement(By.Id("mb_msg")).Text; } catch (NoSuchElementException) { }

                if (!sysMsg.Contains("No records were found"))
                {
                    if (chrome.PageSource.Contains("Track Failed Messages"))//significa solo apareció un contrato
                        chrome.FindElement(By.Id("m397b0593-tabs_middle")).Click();

                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id("md86fe08f_ns_menu_STATUS_OPTION_a")).Click(); //change status
                    WaitCursorToClose(30);

                    chrome.FindElement(By.Id("m88dbf6ce-pb")).Click(); //system msg que va a afectar todo

                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id("m59225159-tb")).Click(); //drop down
                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id(action)).Click(); //opcion Active
                    WaitCursorToClose(30);
                    chrome.FindElement(By.Id("mdb5530cf-pb")).Click(); //ok

                    WaitCursorToClose(30);
                    WaitLongLoading(30);
                }
            }
            catch (Exception ex)
            {
                sysMsg = "ERROR: " + ex;
            }

            CloseCD();

            return sysMsg;
        }
    }
}
