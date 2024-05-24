using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Exception = System.Exception;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using DataBotV5.App.ConsoleApp;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Webex;
using OpenQA.Selenium.Chrome;
using DataBotV5.Logical.Web;
using DataBotV5.Data.Stats;
using System.Globalization;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using OpenQA.Selenium;
using System.Linq;
using System.Data;
using System;

namespace DataBotV5.Automation.WEB.Autopp
{
    /// <summary>
    /// Através de Portal web Fábrica de Propuestas en Smart&Simple los vendedores regionales pueden ingresar los 
    /// diferentes casos de tipo BAW GTL asociados a la oportunidad a crear, este robot toma los casos ingresados 
    /// en el portal, y posterior abre la página de BAW GTL para ingresar los casos indicados por el vendedor, 
    /// de esta forma el robot logra la integración de las dos páginas haciendo más fluido el ciclo de venta.
    ///    
    /// Coded by: Eduardo Piedra Sanabria - Application Management Analyst
    /// </summary>
    class AutoppBAW
    {

        #region Variables locales 
        Logical.Projects.AutoppSS.AutoppLogical logical = new Logical.Projects.AutoppSS.AutoppLogical();
        ProcessInteraction process = new ProcessInteraction();
        string routeSS = "https://smartsimple.gbm.net/";
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        string enviroment = Start.enviroment;
        WebexTeams webex = new WebexTeams();
        DataRow employeeResponsibleData;
        Settings sett = new Settings();
        AutoppInformation oppGestion;
        string LDROrBOMDocument = "";
        Rooting root = new Rooting();
        DataRow employeeCreatorData;
        String notificationsConfig;
        string functionalUser = "";
        bool executeStats = false;
        DataTable configuration;
        string userAdmin = "";
        CRUD crud = new CRUD();
        string respFinal = "";
        DataTable salesTInfo;
        Log log = new Log();
        DataRow client;
        int idOpp;




        #endregion

        public void Main()
        {

            console.WriteLine("Consultando nuevas solicitudes...");
            ProcessBaw();

            if (executeStats == true)
            {
                root.requestDetails = respFinal;

                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

            console.WriteLine("Fin del proceso.");

        }

        /// <summary>
        /// Método principal que invoca los primeros 3 pasos del proceso de Autopp.
        /// </summary>
        public void ProcessBaw()
        {

            #region Status- "Creando casos en BAW" 
            idOpp = 0;
            DataTable newOppRequests1 = GetReqsForStatus("14");
            if (newOppRequests1.Rows.Count > 0)
            {

                executeStats = true;
                int indexReqs1 = 1;

                GetAutoppConfiguration();


                foreach (DataRow oppReq in newOppRequests1.Rows)
                {

                    console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes de creación casos BAW.");

                    idOpp = (int)oppReq.ItemArray[0];
                    LDROrBOMDocument = "";

                    #region General Data
                    GeneralData generalData = new GeneralData();

                    generalData.typeOpportunity = oppReq["typeOpportunity"].ToString();
                    generalData.typeOpportunityName = oppReq["typeOpportunityName"].ToString();
                    generalData.description = oppReq["description"].ToString();
                    generalData.initialDate = DateTime.Parse(oppReq["initialDate"].ToString()).ToString("yyyy-MM-dd");
                    generalData.finalDate = DateTime.Parse(oppReq["finalDate"].ToString()).ToString("yyyy-MM-dd");
                    generalData.cycle = oppReq["cycle"].ToString();
                    generalData.sourceOpportunity = oppReq["sourceOpportunity"].ToString();
                    generalData.salesType = oppReq["salesType"].ToString();
                    generalData.outsourcing = oppReq["outsourcing"].ToString();

                    generalData.typeOpportunity = oppReq["typeOpportunity"].ToString();

                    #endregion

                    #region OrganizationAndClientData
                    DataTable orgInfo = oppInfo(idOpp, "organizationAndClientData");
                    OrganizationAndClientData organizationAndClientData = new OrganizationAndClientData();
                    organizationAndClientData.client =/* "00" +*/ orgInfo.Rows[0]["idClient"].ToString().PadLeft(10, '0');
                    organizationAndClientData.contact = /*"00" +*/ orgInfo.Rows[0]["contact"].ToString().PadLeft(10, '0');
                    organizationAndClientData.salesOrganization = orgInfo.Rows[0]["salesOrgId"].ToString();
                    organizationAndClientData.servicesOrganization = orgInfo.Rows[0]["servOrgId"].ToString().Replace(" ", ""); //0000178579 

                    #endregion

                    #region SalesTeams
                    salesTInfo = oppInfo(idOpp, "salesTeam");
                    List<SalesTeams> salesTList = new List<SalesTeams>();
                    foreach (DataRow salesTItem in salesTInfo.Rows)
                    {
                        SalesTeams item = new SalesTeams();
                        item.role = salesTItem["code"].ToString();
                        item.employee = "AA" + salesTItem["UserID"].ToString().PadLeft(8, '0');

                        salesTList.Add(item);
                    }
                    #endregion

                    #region BAW
                    DataTable BAWInfo = oppInfo(idOpp, "BAW");
                    List<DataBAW> BAWList = new List<DataBAW>();
                    foreach (DataRow bawItem in BAWInfo.Rows)
                    {
                        DataBAW item = new DataBAW();

                        item.id = bawItem["id"].ToString();
                        item.oppId = bawItem["oppId"].ToString();
                        item.vendor = bawItem["vendor"].ToString();
                        item.product = bawItem["productName"].ToString();
                        item.requirementType = bawItem["requirementType"].ToString();
                        item.quantity = bawItem["quantity"].ToString();
                        item.integration = bawItem["isIntegration"].ToString();
                        item.comments = bawItem["comments"].ToString();

                        BAWList.Add(item);
                    }
                    #endregion


                    #region Objeto principal donde se une toda la información
                    //AutoppInformation oppGestion = new AutoppInformation();
                    oppGestion = new AutoppInformation();

                    oppGestion.id = oppReq["id"].ToString();
                    oppGestion.status = oppReq["status"].ToString();
                    oppGestion.employee = oppReq["createdBy"].ToString();
                    oppGestion.opp = oppReq["opp"].ToString();
                    oppGestion.generalData = generalData;
                    oppGestion.organizationAndClientData = organizationAndClientData;
                    oppGestion.salesTeams = salesTList;
                    oppGestion.BAW = BAWList;

                    #endregion


                    #region Empleado que creó la oportunidad.
                    //employeeCreatorData = new CCEmployee(oppGestion.employee);
                    //oppGestion.employee = employeeCreatorData.IdEmpleado;
                    string sqlEmployeeCreatorData = $"select * from MIS.digital_sign where user= '{oppGestion.employee}'";
                    employeeCreatorData = crud.Select(sqlEmployeeCreatorData, "MIS", enviroment).Rows[0];



                    #endregion

                    #region Empleado con rol de empleado responsable en la oportunidad.

                    string sqlEmployeeResponsible = $"select * from MIS.digital_sign where id=(SELECT employee FROM autopp2_db.SalesTeam where role= 41 and oppId= {oppGestion.id})";
                    //DataRow employeeResponsibleDt = crud.Select( sqlEmployeeResponsible, "databot_db", enviroment).Rows[0];
                    employeeResponsibleData = crud.Select(sqlEmployeeResponsible, "databot_db", enviroment).Rows[0];

                    //employeeResponsibleData = new CCEmployee(employeeResponsibleDt["user"].ToString());


                    #endregion

                    #region Extraer el nombre del cliente
                    //Extraer el nombre del cliente.
                    string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {organizationAndClientData.client}";
                    client = crud.Select(sqlClient, "databot_db", enviroment).Rows[0];

                    #endregion

                    console.WriteLine("");
                    console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

                    #region Paso - En proceso crear requisitos BAW
                        Step3CreateBAWRequirements(); 
                  
                    #endregion




                    console.WriteLine("");
                    indexReqs1++;

                }

                //Usuario funcional
                root.BDUserCreatedBy = functionalUser;

            }

            #endregion


        }


        #region Métodos con cada uno de los pasos del proceso Autopp

        /// <summary>
        /// Crea los requerimientos en BAW y notifica a las personas correspondientes.
        /// </summary>
        /// <returns>Retorna true si todo salió sin ningún error.</returns>
        public void Step3CreateBAWRequirements()
        {

            #region Paso 3- "En proceso crear requisitos BAW" 
            bool loginOpened3 = false;

            console.WriteLine("");
            console.WriteLine("******************************");
            console.WriteLine("*Fase: Crear requisitos BAW*");
            console.WriteLine("******************************");
            console.WriteLine("");

            IWebDriver chrome = null;
            try
            {
                //Establecer conexión en BAW
                chrome = SelConn(GetUrlBaw());
                //console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");


                #endregion

                #region Agregar requerimientos en BAW 


                #region Login en BAW.
                //console.WriteLine("Iniciando sesión en baw...");
                if (!loginOpened3)
                {

                    chrome.FindElement(By.Id("username")).SendKeys(cred.user_baw);
                    chrome.FindElement(By.Id("password")).SendKeys(cred.password_baw);

                    chrome.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/a[1]")).Submit();
                    //System.Threading.Thread.Sleep(7000);
                    loginOpened3 = true;
                    console.WriteLine("Inicio de sesión exitoso.");
                }
                #endregion

                #region Buscar la oportunidad 
                //console.WriteLine($"Buscando la oportunidad ...");
                //Ingresar la oportunidad a buscar
                try
                {
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input"))); } catch { System.Threading.Thread.Sleep(5000); }
                    //System.Threading.Thread.Sleep(5000);
                    chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(oppGestion.opp);
                }
                catch (Exception)
                {
                    try
                    {
                        System.Threading.Thread.Sleep(5000);
                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(oppGestion.opp);
                    }
                    catch (Exception)
                    {
                        System.Threading.Thread.Sleep(10000);
                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(oppGestion.opp);
                    }


                }
                //Click buscar               

                chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();

                System.Threading.Thread.Sleep(5000);

                //Primer registro
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a"))); } catch { }
                chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();

                console.WriteLine("Oportunidad encontrada.");

                #endregion

                #region Requerimientos de BAW
                //console.WriteLine("Ingresando los requerimientos de BAW de la oportunidad.");


                //Cambio a IFrame
                try
                {
                    System.Threading.Thread.Sleep(5000);
                    IWebElement iframe = chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/div[4]//following-sibling::iframe[1]"));
                    chrome.SwitchTo().Frame(iframe);
                }
                catch (Exception e)
                {
                    System.Threading.Thread.Sleep(3000);
                    IWebElement iframe = chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/div[4]//following-sibling::iframe[1]"));
                    chrome.SwitchTo().Frame(iframe);
                }



                int quantityRowsPerPAge = 5;

                int page = 0;

                string updateQueryBAW = "";

                for (int i = 1; i <= oppGestion.BAW.Count; i++)
                {
                    //Seleccionar nuevo en el dropdown de acción
                    System.Threading.Thread.Sleep(500);
                    try
                    {
                        string xpathActionTd = $"/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[3]/table/tbody/tr[{i - (page * 5)}]/td[2]/div/div/div/div/select";
                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(xpathActionTd))); } catch { }

                        SelectElement actionReq = new SelectElement(chrome.FindElement(By.XPath(xpathActionTd)));
                        System.Threading.Thread.Sleep(500);
                        actionReq.SelectByValue("nuevo");
                        System.Threading.Thread.Sleep(500);
                    }
                    catch (Exception k)
                    {
                        System.Threading.Thread.Sleep(5000);
                        string xpathActionTd = $"/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[3]/table/tbody/tr[{i - (page * 5)}]/td[2]/div/div/div/div/select";
                        SelectElement actionReq = new SelectElement(chrome.FindElement(By.XPath(xpathActionTd)));
                        System.Threading.Thread.Sleep(1000);
                        actionReq.SelectByValue("nuevo");
                        System.Threading.Thread.Sleep(1000);
                    }

                    //Obtener el número de caso
                    string caseNumber = chrome.FindElement(By.XPath($"/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[3]/table/tbody/tr[{i - (page * 5)}]/td[1]/div/div/div/div/input")).GetAttribute("value");
                    console.WriteLine($"Case number: {caseNumber} generado.");
                    updateQueryBAW += $"UPDATE `BAW` SET `caseNumber`= '{caseNumber}' WHERE id= {oppGestion.BAW[i - 1].id}; ";

                    //Entrar al requerimiento (botón azul)
                    string xpathBlueBtnTd = $"/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[3]/table/tbody/tr[{i - (page * 5)}]/td[3]/div/div/button";
                    chrome.FindElement(By.XPath(xpathBlueBtnTd)).Click();

                    SelectElement aux;

                    //Proveedor
                    try
                    { //El primer intento dura más de lo normal
                        try
                        {
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='combo_div_3']"))); } catch { }
                            //System.Threading.Thread.Sleep(2500);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_3']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].vendor.Trim());
                        }
                        catch (Exception)
                        {
                            System.Threading.Thread.Sleep(2000);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_3']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].vendor.Trim());
                        }
                    }
                    catch (Exception)
                    {
                        System.Threading.Thread.Sleep(15000);
                        aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_3']")));
                        aux.SelectByValue(oppGestion.BAW[i - 1].vendor.Trim());
                    };


                    //Nombre de Producto
                    try
                    {
                        System.Threading.Thread.Sleep(3500);
                        aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_4']")));
                        aux.SelectByValue(oppGestion.BAW[i - 1].product.Trim());
                    }
                    catch (Exception)
                    {
                        System.Threading.Thread.Sleep(8000);
                        aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_4']")));
                        aux.SelectByValue(oppGestion.BAW[i - 1].product.Trim());
                        try
                        {
                            System.Threading.Thread.Sleep(8000);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_4']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].product.Trim());
                        }
                        catch (Exception)
                        {

                        }
                    }


                    //Tipo de Requerimiento

                    try
                    {
                        System.Threading.Thread.Sleep(1400);
                        aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_5']")));
                        aux.SelectByValue(oppGestion.BAW[i - 1].requirementType.Trim());
                    }
                    catch (Exception e)
                    {

                        try
                        {
                            System.Threading.Thread.Sleep(1000);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_5']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].requirementType.Trim());

                        }
                        catch (Exception)
                        {
                            System.Threading.Thread.Sleep(7000);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_5']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].requirementType.Trim());

                        }
                    }

                    //Cantidad
                    try
                    {
                        System.Threading.Thread.Sleep(200);
                        chrome.FindElement(By.Id("input_div_2_1_2_1_1")).SendKeys(oppGestion.BAW[i - 1].quantity);
                        //Checkbox Documentos adjuntados.
                        chrome.FindElement(By.XPath("//*[@id='div_7']/div/div[2]/label/input")).Click();
                    }
                    catch (Exception)
                    {
                        System.Threading.Thread.Sleep(450);
                        chrome.FindElement(By.Id("input_div_2_1_2_1_1")).SendKeys(oppGestion.BAW[i - 1].quantity);
                        //Checkbox Documentos adjuntados.
                        chrome.FindElement(By.XPath("//*[@id='div_7']/div/div[2]/label/input")).Click();
                    }




                    if (oppGestion.BAW[i - 1].integration == "Si")
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(50);
                            //Checkbox Parte de una integración
                            chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();

                        }
                        catch (Exception)
                        {
                            System.Threading.Thread.Sleep(650);
                            //Checkbox Parte de una integración
                            chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();
                        }

                    }

                    if (oppGestion.BAW[i - 1].comments.ToString() != "undefined")
                    {
                        //Comentarios 
                        chrome.FindElement((By.XPath("//*[@id='textArea_div_9']"))).SendKeys(Regex.Replace(oppGestion.BAW[i - 1].comments, @"[^\w\.@-]", " ",
                                                   RegexOptions.None, TimeSpan.FromSeconds(1.5)));
                    }

                    System.Threading.Thread.Sleep(300);
                    //Botón de aceptar
                    chrome.FindElement(By.XPath("//*[@id='div_12_1_4_1_1_1_1']/div/button")).Click();



                    if (i != oppGestion.BAW.Count) //Descartar tocar el botón + en el último registro
                    {
                        try
                        {
                            //System.Threading.Thread.Sleep(1600);
                            //Boton de + en requerimiento
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[4]/button"))); } catch { }
                            chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[4]/button")).Click();

                        }
                        catch (Exception)
                        {

                            System.Threading.Thread.Sleep(30000);
                            //Boton de + en requerimiento

                            chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[4]/button")).Click();

                        }
                    }
                    #endregion

                    //Significa que la iteración actual es divisible y entera
                    if (i % quantityRowsPerPAge == 0)
                    {
                        page += 1;
                    }

                    //Se actualiza el status del requerimiento en la tabla BAW.
                    updateQueryBAW += $" UPDATE `BAW` SET `statusBAW`= 2 WHERE id={oppGestion.BAW[i - 1].id}; ";

                }

                //Botón final verde de continuar y finalizar 1 etapa de gestión en espera de respuesta del especialista.
                try
                {
                    new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_7_1_4_1_1_1_1']/div/button")));
                    chrome.FindElement(By.XPath("//*[@id='div_7_1_4_1_1_1_1']/div/button")).Click();
                }
                catch
                {
                    System.Threading.Thread.Sleep(2000);
                    chrome.FindElement(By.XPath("//*[@id='div_7_1_4_1_1_1_1']/div/button")).Click();
                }


                console.WriteLine($"Requisitos de la oportunidad {oppGestion.opp} creados en BAW.");
                console.WriteLine($"");

                System.Threading.Thread.Sleep(1500);

                ////Salir del iframe
                //chrome.SwitchTo().DefaultContent();
                //System.Threading.Thread.Sleep(500);

                ////Limpiar el buscador 
                //chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/div/a/i")).Click();


                #region Actualizar status y log de cambios   
                crud.Update(updateQueryBAW, "autopp2_db", enviroment);
                //log.LogDeCambios("Creación", "Autopp", oppGestion.employee, "", oppGestion.opp, "Creación de BAW");
                #endregion

                string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=4 WHERE id= {oppGestion.id}; ";
                crud.Update(updateQuery, "autopp2_db", enviroment);

                //Notificación de exito
                NotifySuccessOrErrors("Success", 3);

                //Notificar a los SalesTeams que han sido agregados a la opp
                NotifySuccessOrErrors("Success", 9);

                log.LogDeCambios("Creación", root.BDProcess, oppGestion.employee, "Creación de BAW", "Creación de requerimientos de BAW del cliente: " + oppGestion.organizationAndClientData.client, oppGestion.employee);
                respFinal = respFinal + "\\n" + "Creación de requerimientos de BAW del cliente: " + oppGestion.organizationAndClientData.client;





                #endregion
            }
            catch (Exception e)
            {
                string updateQuery = $"UPDATE OppRequests SET status = 8, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                crud.Update(updateQuery, "autopp2_db", enviroment);


                //Notificar error al usuario y application management.
                NotifySuccessOrErrors("Fail", 3, null, e, chrome);

                //Cerrar chrome
                process.KillProcess("chromedriver", true);
                process.KillProcess("chrome", true);


                loginOpened3 = false;

            }




            System.Threading.Thread.Sleep(2000);

            #region Cerrar Chrome
            process.KillProcess("chromedriver", true);
            process.KillProcess("chrome", true);
            #endregion




        }


        #endregion



        #region Métodos útiles para la gestión de cada uno de los pasos de AutoppProcess.



        /// <summary>
        /// Método para extraer las solicitudes en la tabla OppRequests según su status actual.
        /// </summary>
        /// <param name="status"></param>
        /// <returns></returns>
        /// 
        private DataTable GetReqsForStatus(string status)
        {
            string sql =
            $@"SELECT 
            opp.id, 
            opp.opp, 
            TypeOportunity.code typeOpportunity,  
            TypeOportunity.typeOportunity typeOpportunityName, 
            opp.description, 
            opp.initialDate, 
            opp.finalDate, 
            SalesCycle.code cycle, 
            SourceOportunity.code sourceOpportunity, 
            SalesType.code salesType, 
            ApplyOutsourcing.code outsourcing, 
            opp.status, 
            opp.createdBy 

            FROM OppRequests opp 
            LEFT JOIN TypeOportunity ON TypeOportunity.id = opp.typeOpportunity 
            LEFT JOIN SalesCycle ON SalesCycle.id = opp.cycle 
            LEFT JOIN SourceOportunity ON SourceOportunity.id = opp.sourceOpportunity 
            LEFT JOIN SalesType ON SalesType.id = opp.salesType 
            LEFT JOIN ApplyOutsourcing ON ApplyOutsourcing.id = opp.outsourcing 

            WHERE opp.active = 1 
            AND opp.status = {status} ";

            DataTable reqTable = crud.Select(sql, "autopp2_db", enviroment);

            return reqTable;
        }

        /// <summary>
        ///Extraer información en la DB relacionada a una oportunidad, como el equipo de ventas, cliente, LDRS, BAW.
        /// </summary>
        /// <param name="status"></param>
        /// <returns></returns>
        private DataTable oppInfo([Optional] int oppId, string typeInformation)
        {
            DataTable mytable = new DataTable();
            string sql = "";

            switch (typeInformation)
            {
                case "organizationAndClientData":
                    sql =
                    "SELECT " +
                    "databot_db.clients.idClient , " +
                    "contact, " +
                    "databot_db.salesOrganizations.salesOrgId , " +
                    "databot_db.serviceOrganizations.servOrgId " +

                    "FROM `OrganizationAndClientData` org " +
                    "LEFT JOIN databot_db.salesOrganizations ON databot_db.salesOrganizations.id = org.salesOrganization " +
                    "LEFT JOIN databot_db.serviceOrganizations ON databot_db.serviceOrganizations.id = org.servicesOrganization " +
                    "LEFT JOIN databot_db.clients ON databot_db.clients.id = org.client" +

                    $" WHERE org.oppId = {oppId} " +
                    "AND org.active = 1; ";
                    break;

                case "salesTeam":
                    sql =
                    "SELECT salesT.oppId, " +
                    "EmployeeRole.code, " +
                    "EmployeeRole.employeeRole, " +
                    "MIS.digital_sign.user, " +
                    "MIS.digital_sign.UserID, " +
                    "EmployeeRole.id " +



                    "FROM `SalesTeam` salesT " +

                    "LEFT JOIN EmployeeRole ON EmployeeRole.id = salesT.role " +
                    "LEFT JOIN MIS.digital_sign ON MIS.digital_sign.id = salesT.employee " +

                    $"WHERE salesT.oppId = {oppId} " +
                    "AND salesT.active = 1";
                    break;

                case "BAW":
                    sql =
                    "SELECT " +
                    "baw.id, " +
                    "baw.oppId, " +
                    "Vendor.vendor, " +
                    "ProductName.productName, " +
                    "RequirementType.requirementType, " +
                    "baw.quantity, " +
                    "IsIntegration.isIntegration, " +
                    "baw.comments, " +
                    "baw.statusBAW, " +
                    "baw.caseNumber " +


                    "FROM `BAW` baw " +

                    "LEFT JOIN Vendor ON Vendor.id = baw.vendor " +
                    "LEFT JOIN ProductName ON ProductName.id = baw.product " +
                    "LEFT JOIN RequirementType ON RequirementType.id = baw.requirementType " +
                    "LEFT JOIN IsIntegration ON IsIntegration.id = baw.integration "; //+

                    //$"WHERE baw.oppId = {oppId} " +
                    //"AND baw.active = 1 ";

                    //En caso que no tenga opp significa que lo esta solicitando Step5ForwardOpp, y solo se quiere saber las que ya estan listas para reenviar al especialista.
                    if (oppId == null || oppId == 0)
                    {
                        sql += " WHERE baw.statusBAW=6 " +
                            "AND baw.active = 1 ";
                    }
                    else
                    { //Se quiere extraer los requerimientos BAW por oppID
                        sql += $" WHERE baw.oppId = {oppId} " +
                        " AND baw.active = 1 ";
                    }
                    break;

            }

            mytable = crud.Select(sql, "autopp2_db", enviroment);

            return mytable;
        }

        /// <summary>
        ///Método para notificar errores o proceso exitoso através de Webex Teams y vía correo electrónico según sea el caso
        ///En successOrFailMode indicar "Success" para éxito, ó "Fail" para errores.
        /// </summary>
        /// <param name="status"></param>
        /// <returns></returns>
        private void NotifySuccessOrErrors(string successOrFailMode, int fase, [Optional] List<string> listErrorsFase2, [Optional] Exception exception, [Optional] IWebDriver chrome)
        {
            string titleWebex = "";
            string employeeName = "";

            try //Seleccionar el nombre y primer letra en mayúscula.
            {
                employeeName = BuildFirstName(employeeCreatorData["name"].ToString());
            }
            catch (Exception e) { }

            #region Éxito ó notificaciones de información
            if (successOrFailMode == "Success") //Éxito
            {

                switch (fase)
                {
                    case 3:
                        #region Notificación de éxito de creación de la opp, LDR, BAW.

                        if (oppGestion.opp != "" || true)
                        {

                            Dictionary<string, string> toReplace = new Dictionary<string, string>(){
                                    {"TITLENOTIFICATION", "Creación éxitosa de oportunidad"},
                                    {"USER", employeeCreatorData["user"].ToString()},
                                    {"EMPLOYEERESPONSIBLE", employeeResponsibleData["user"].ToString()},
                                    {"TYPEOPPORTUNITY", oppGestion.generalData.typeOpportunityName},
                                    {"OPP", oppGestion.opp},
                                    {"CLIENT", client["name"].ToString()}
                                };

                            logical.AutoppNotifications("successNotification", setUser(employeeCreatorData["user"].ToString()), toReplace);

                            if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                            {
                                logical.AutoppNotifications("successNotification", setUser(employeeResponsibleData["user"].ToString()), toReplace);
                            }

                            if (LDROrBOMDocument == "LDR")
                            {
                            }
                            else { }

                            //Respaldo para el admin.
                            logical.AutoppNotifications("successNotification", userAdmin, toReplace);


                        }
                        #endregion
                        break;

                    case 9:
                        #region Notificación al equipo de SalesTeam que ha sido agregado a la oportunidad.

                        if (oppGestion.opp != "" || true)
                        {
                            foreach (DataRow employee in salesTInfo.Rows)
                            {
                                if (employee["id"].ToString() != "41" /*Empleado Responsable*/ ||
                                    (employeeResponsibleData["user"].ToString() != employeeCreatorData["user"].ToString()) //Si el empleado que lo creó es el mismo empleado responsable, que no se notifique ha sido agregado al Sales Teams.
                                    )
                                {
                                    Dictionary<string, string> toReplace = new Dictionary<string, string>(){
                                    {"TITLENOTIFICATION", "Agregado en equipo de ventas"},
                                    {"USER", employee["user"].ToString()},
                                    {"EMPLOYEERESPONSIBLE", employeeResponsibleData["user"].ToString()},
                                    {"TYPEOPPORTUNITY", oppGestion.generalData.typeOpportunityName},
                                    {"TYPEROLE", employee["employeeRole"].ToString()},
                                    {"OPP", oppGestion.opp},
                                    {"CLIENT", client["name"].ToString()}
                                };

                                    logical.AutoppNotifications("salesTeamsNotification", setUser(employee["user"].ToString()), toReplace);

                                }

                            }

                        }
                        #endregion
                        break;

                }


            }
            #endregion

            #region Fallo

            else if (successOrFailMode == "Fail") //Fallo
            {

                switch (fase)
                {
                    case 3:
                        #region Error al crear el requerimiento.

                        if (oppGestion.opp != "")
                        {
                            #region Notificación por Webex Teams al usuario
                            titleWebex = $"Error al crear requerimiento en BAW - {oppGestion.opp} - {oppGestion.id}";

                            string msgWebex3 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                            $"del cliente: " + client["name"].ToString();
                            msgWebex3 += $", ha tenido un problema al crear los requerimientos en la página de BAW, puede revisar el estado de cada uno en la página de <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                            webex.SendCCNotification( setUser(employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear requerimientos en BAW", oppGestion.opp, msgWebex3);

                            //En caso que sea diferente, notifique al usuario responsable también
                            if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                            {

                                //Seleccionar el nombre y primer letra en mayúscula.
                                try { employeeName = BuildFirstName(employeeResponsibleData["name"].ToString()); } catch (Exception e) { employeeName = ""; }


                                msgWebex3 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                                    $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                                msgWebex3 += $", ha tenido un problema al crear los requerimientos en la página de BAW, puede revisar el estado de cada uno en la página de <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                                webex.SendCCNotification( setUser( employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear requerimientos en BAW", oppGestion.opp, msgWebex3);
                                //webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al crear requerimientos en BAW", oppGestion.opp, msgWebex3);
                            }

                            //Respaldo admin
                            webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al crear requerimientos en BAW", oppGestion.opp, msgWebex3);

                            #endregion



                            #region Notificar a Application Management


                            string msgApp = "Este error está en el try catch de la Fase 3: Crear requisitos BAW  - Autopp. Revisar el screenshot del error en la carpeta en el desktop/Databot/Autopp/ErrorsScreenshots.  Sugerencia: ingrese a BAW con las credenciales del rpauser, encuentre la opp, y limpie los requerimientos que quedaron en borrador, dejando solo la primera fila creada y en blanco, posteriormente corra la solicitud en su computadora local y finalice el proceso de creación de requerimientos.";
                            sett.SendError(this.GetType(), $"Error al crear requisitos en BAW del id #{idOpp}", msgApp, exception);


                            try
                            {
                                string pathErrors = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\databot\Autopp\ErrorsScreenshots\";
                                Screenshot TakeScreenshot = ((ITakesScreenshot)chrome).GetScreenshot();
                                TakeScreenshot.SaveAsFile(pathErrors + $"{idOpp} - {oppGestion.opp} - crear requisitos BAW.png");
                            }
                            catch (Exception i)
                            {
                                console.WriteLine("No se pudo guardar el screenshot del error. El error al guardar fue:");
                                console.WriteLine("");
                                console.WriteLine(i.ToString());
                            }


                            #endregion



                        }

                        #endregion
                        break;

                }

            }
            #endregion

            else
            {
                sett.SendError(this.GetType(), $"Problema al notificar a usuarios - Autopp",
                    $"El parámetro successOrFailMode del método NotifyErrorsOrSuccess es: {successOrFailMode}, por tanto " +
                    $"no está notificando a nadie si hay errores o éxito en sus gestiones. Por favor revisar las instancias " +
                    $"que envíen el parámetro successOrFailMode de manera correcta (sucess or fail).");
            }
        }

        /// <summary>
        /// Método para establecer la conexión a Chrome.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public IWebDriver SelConn(string url)
        {
            #region eliminar cache and cookies chrome
            try
            {
                process.KillProcess("chromedriver", true);
                process.KillProcess("chrome", true);
            }
            catch (Exception)
            { }
            #endregion

            ChromeOptions options = new ChromeOptions();
            WebInteraction sel = new WebInteraction();

            //options = sel.ChromeOptions();
            options.AddArguments("--ignore-ssl-errors=yes");
            options.AddArguments("--ignore-certificate-errors");
            // driver = webdriver.Chrome(options = options);

            //options.AddArguments("user-data-dir=" + optionsfolder);
            //options.AddUserProfilePreference("download.default_directory", downloadfolder);

            IWebDriver chrome = sel.NewSeleniumChromeDriver(options);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl(url);
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl(url); }

            chrome.Manage().Cookies.DeleteAllCookies();
            #endregion

            return chrome;
        }

        /// <summary>
        /// Método para obtener el URL de BAW según el enviroment establecido (PRD o QAS).
        /// </summary>
        /// <returns></returns>
        private string GetUrlBaw()
        {
            if (enviroment == "PRD")
                return "https://prod-ihs-03.gbm.net/ProcessPortal/login.jsp";
            else if (enviroment == "QAS")
                return "https://test-ihsbaw-01.gbm.net/ProcessPortal/login.jsp";
            else
                return "";
        }


        /// <summary>
        /// Método para en caso si viene un nombre como: "PIEDRA SANABRIA, EDUARDO ANTONIO", solo devuelva el primer nombre: Eduardo.
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Devuelve un string con el nombre.</returns>
        public string BuildFirstName(string name)
        {
            //string name = "MEZA CASTRO, DIEGO";
            string[] last = name.Split(' ');

            int index = 0;
            for (int i = 0; i < last.Count(); i++)
            {
                if (last[i].Contains(',') == true)
                {
                    index = i;
                }
            }

            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(last[index + 1].ToLower());

        }

        /// <summary>
        /// Método para extraer configuración de Autopp
        /// </summary>
        /// <param name="name"></param>
        /// <returns>void</returns>
        public void GetAutoppConfiguration()
        {
            #region Extraer la configuración de Autopp.
            string sqlConfiguration = $"select * from Configuration";
            configuration = crud.Select(sqlConfiguration, "autopp2_db", enviroment);

            //Si es admin o si es para todos los usuarios.
            notificationsConfig = configuration.Select($"typeConfiguration = 'notifications'")[0]["configuration"].ToString().ToLower();
            userAdmin = configuration.Select($"typeConfiguration = 'userAdmin'")[0]["configuration"].ToString().ToLower();
            functionalUser = configuration.Select($"typeConfiguration = 'funcionalUser'")[0]["configuration"].ToString().ToLower();

            #endregion
        }

        /// <summary>
        /// Método para definir si el usuario debe ser el admin o el los parámetros para las notificaciones.
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Devuelve un string con el usuario.</returns>
        public string setUser(string user)
        {
            //Si se lo envía a admin o a los usuarios.
            string userToSend = notificationsConfig == "admin" ? userAdmin : user;
            return userToSend;

        }

        #endregion




























    }
}
















