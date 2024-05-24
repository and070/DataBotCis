using DataBotV5.Logical.Projects.Modals.Single;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.Projects.EasyLDR;
using System.Runtime.InteropServices;
using DataBotV5.Data.Projects.Autopp;
using OpenQA.Selenium.Interactions;
using Exception = System.Exception;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Webex;
using OpenQA.Selenium.Chrome;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Web;
using DataBotV5.Data.Stats;
using System.Globalization;
using Newtonsoft.Json.Linq;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using OpenQA.Selenium;
using System.Linq;
using System.Data;
using System.IO;
using System;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Firefox;
using DataBotV5.Logical.Projects.AutoppSS;

namespace DataBotV5.Automation.QAS.Autopp
{
    class AutoppSS_QAS
    {

        #region Variables locales 
        ProcessInteraction process = new ProcessInteraction();
        string routeSS = "https://smartsimple.gbm.net/";
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        AutoppSQL autoppSQL = new AutoppSQL();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        WebexTeams webex = new WebexTeams();
        //CCEmployee employeeResponsibleData;
        DataRow employeeResponsibleData;
        Settings sett = new Settings();
        AutoppInformation oppGestion;
        string LDROrBOMDocument = "";
        Rooting root = new Rooting();
        DataRow employeeCreatorData;
        string enviroment = "QAS";
        bool executeStats = false;
        string userAdmin = "";
        string caseNumber = "";
        CRUD crud = new CRUD();
        string respFinal = "";
        DataTable salesTInfo;
        Log log = new Log();
        string sapSystem = "CRM";
        int mandante = 460;
        DataRow client;
        AutoppSS logical = new AutoppSS();
        int idOpp;
        DataTable configuration;
        String notificationsConfig;



        #endregion

        public void Main()
        {

            console.WriteLine("Consultando nuevas solicitudes...");
            ProcessAutopp();
            Step456EmailsActionsBAW();
            Step7ForwardBAWRequirement();

            if (executeStats == true)
            {
                //root.requestDetails = respFinal;

                //console.WriteLine("Creando estadísticas...");
                //using (Stats stats = new Stats())
                //{
                //    stats.CreateStat();
                //}
            }

            console.WriteLine("Fin del proceso.");

        }

        /// <summary>
        /// Método principal que invoca los primeros 3 pasos del proceso de Autopp.
        /// </summary>
        public void ProcessAutopp()
        {

            #region Status 1- "En Proceso" 
            idOpp = 0;
            DataTable newOppRequests1 = GetReqsForStatus("1");
            if (newOppRequests1.Rows.Count > 0)
            {

                executeStats = true;
                int indexReqs1 = 1;

                GetAutoppConfiguration();


                foreach (DataRow oppReq in newOppRequests1.Rows)
                {

                    console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes.");

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

                    #region LDRS deserializar
                    List<LDRSAutopp> listLDRFather = new List<LDRSAutopp>();

                    DataTable LDRSInfo = oppInfo(idOpp, "LDRS");

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

                    oppGestion.LDRS = LDRSInfo;
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

                    //Notificación de éxito
                    //NotifySuccessOrErrors("Success", 3);

                    //console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes.");
                    console.WriteLine("");
                    console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

                    #region Paso 1 - Crear Oportunidad CRM
                    bool resultStep1 = Step1CreateOpp();
                    #endregion

                    #region Paso 2 - En proceso crear LDRS
                    bool resultStep2 = false;

                    if (resultStep1

                        && (oppGestion.generalData.cycle == "Y3A" || oppGestion.generalData.cycle == "Y3")
                        )
                        resultStep2 = Step2CreateLDR();
                    #endregion

                    #region Paso 3 - En proceso crear requisitos BAW
                    if (resultStep1 && oppGestion.generalData.cycle == "Y3A" /*Solo para GTL Quotation*/)
                    { Step3CreateBAWRequirements(); }

                    if (resultStep1 && oppGestion.generalData.cycle != "Y3A") //En caso que no sea GTL Quotation mande una notificación 
                    {
                        //Notificación de éxito
                        NotifySuccessOrErrors("Success", 8);

                        //Notificar a los SalesTeams que han sido agregados a la opp
                        NotifySuccessOrErrors("Success", 9);

                        //Finalizar el proceso actualizando estado.
                        string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=5 WHERE id= {oppGestion.id}; ";
                        crud.Update(updateQuery, "autopp2_db", enviroment);
                    }
                    #endregion




                    console.WriteLine("");
                    indexReqs1++;

                }

                //Usuario funcional
                root.BDUserCreatedBy = "DGARCIA";

            }

            #endregion


        }


        #region Métodos con cada uno de los pasos del proceso Autopp
        /// <summary>
        /// Primer paso, creación de oportunidad en SAP.
        /// </summary>
        /// <returns>Retorna true si todo salió sin ningún error.</returns>

        public bool Step1CreateOpp()
        {

            console.WriteLine("");
            console.WriteLine("**********************************");
            console.WriteLine("*Fase 1: Crear Oportunidad en SAP*");
            console.WriteLine("**********************************");
            console.WriteLine("");


            //console.WriteLine("");
            //console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

            try
            {

                #region Crear la oportunidad y notificación de éxito o fallo.
                //oppGestion.opp = CreateOppCRM(oppGestion, employeeCreatorData.Usuario);
                oppGestion.opp = CreateOppCRM(oppGestion, employeeCreatorData["user"].ToString());
                //idOpp = 0;

                if (oppGestion.opp == "" || oppGestion.opp == null)
                {//Fallo 
                    NotifySuccessOrErrors("Fail", 1);

                    string updateQuery = $"UPDATE OppRequests SET status = 6, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                    crud.Update(updateQuery, "autopp2_db", enviroment);
                    return false;
                }

                //log.LogDeCambios("Creación", root.BDProcess, oppGestion.employee, "Nueva oportunidad: " + oppGestion.opp, "Se generó la oportunidad: " + oppGestion.opp + " del cliente: " + oppGestion.organizationAndClientData.client, oppGestion.employee);
                //respFinal = respFinal + "\\n" + "Se generó una nueva oportunidad: " + oppGestion.opp + " del cliente: " + oppGestion.organizationAndClientData.client;

                string updateQueryOpp = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {oppGestion.id}";
                crud.Update(updateQueryOpp, "autopp2_db", enviroment);

                #endregion

            }
            catch (Exception e)
            {
                //Notificar error al usuario y Application Management
                NotifySuccessOrErrors("Fail", 1);

                string updateQuery = $"UPDATE OppRequests SET status = 6, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                crud.Update(updateQuery, "autopp2_db", enviroment);

                return false;
            }

            return true;



        }


        /// <summary>
        /// Crea el LDR (levantamiento de requerimiento), y lo sube a SAP y FTP de Smart And Simple.
        /// </summary>
        /// <returns>Retorna true si todo salió sin ningún error.</returns>
        public bool Step2CreateLDR()
        {

            #region Paso 2- "En proceso crear LDRS" 


            console.WriteLine("");
            console.WriteLine("********************");
            console.WriteLine("*Fase 2: Crear LDRS*");
            console.WriteLine("********************");
            console.WriteLine("");


            //En caso que al subir el LDR diera error lo agrega a la siguiente lista.
            List<string> errorsList = new List<string>();

            try
            {

                //console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

                //Aquí se crea el LDR
                Dictionary<string, string> filesRoutes = CreateLDR(oppGestion);

                //En caso que el usuario haya almacenados archivos BOM, lo descarga y almacena en filesRoutes 
                DownloadsBOMFiles(filesRoutes, oppGestion.id);

                if (filesRoutes.Count > 0) //Verifica si se debe subir algún LDRS.
                {

                    #region Subir archivo al FTP          

                    bool resultFTP = true;
                    foreach (KeyValuePair<string, string> fileLDR in filesRoutes)
                    {
                        if (fileLDR.Key.Contains("LDR"))
                        {
                            resultFTP = autoppSQL.InsertFileAutopp(oppGestion.id, fileLDR.Value + fileLDR.Key, enviroment);
                            LDROrBOMDocument = "LDR";
                            console.WriteLine($"LDR subido al FTP con éxito.");
                        }
                    }

                    if (!resultFTP)
                    {
                        console.WriteLine($"Ocurrió un error no se pudo subir al FTP.");
                        errorsList.Add("Error al subir al FTP");
                    }


                    #endregion

                    #region Subir archivo a SAP

                    bool blockByAutoppLDR = false;

                    if (!sap.CheckLogin(sapSystem, mandante))
                    {
                        //Bloquear RPA User
                        sap.BlockUser(sapSystem, 1,  mandante);
                    }
                    else
                    {
                        while (sap.CheckLogin(sapSystem, mandante) && blockByAutoppLDR == false)
                        {
                            console.WriteLine($"Mandante de SAP {sapSystem + " " + mandante} bloqueado, esperando su desbloqueo...");
                            System.Threading.Thread.Sleep(30000);

                            //Por fin lo encontró desbloqueado, bloqueelo.
                            if (!sap.CheckLogin(sapSystem, mandante))
                            {
                                //Bloquear RPA User
                                sap.BlockUser(sapSystem, 1, mandante);
                                blockByAutoppLDR = true;

                            }
                        }
                    }

                    process.KillProcess("saplogon", false);
                    sap.LogSAP(sapSystem, mandante);

                    EasyLDR con = new EasyLDR();
                    bool resultLDR = con.ConnectSAP(filesRoutes, oppGestion.opp);

                    if (resultLDR)
                    {
                        console.WriteLine($"Archivos subidos a SAP con éxito.");
                    }
                    else
                    {
                        console.WriteLine($"Ocurrió un error no se pudo subir el archivo a SAP.");
                        errorsList.Add("Error al subir a SAP");
                    }

                    sap.KillSAP();

                    //Desbloquear RPA User
                    sap.BlockUser(sapSystem, 0, mandante);
                    //Cerrar proceso Excel
                    process.KillProcess("EXCEL", true);
                    process.KillProcess("saplogon", true);


                    #endregion

                    #region Notificar errores de carga.                    

                    if (errorsList.Count > 0)
                    {

                        //Notificar error al usuario
                        NotifySuccessOrErrors("Fail", 2, errorsList);

                        //Se pone en status error.
                        string updateQuery = $"UPDATE OppRequests SET status = 7, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                        crud.Update(updateQuery, "autopp2_db", enviroment);
                        return false;

                    }
                    else // Si no hay errores elimina los archivos de forma local
                    {
                        foreach (KeyValuePair<string, string> file in filesRoutes)
                        {
                            File.Delete(file.Value + file.Key);
                        }

                        //log.LogDeCambios("Creación", root.BDProcess, oppGestion.employee, "Creación de LDR", "Creación de LDR o subir archivos a SAP y FTP del cliente: " + oppGestion.organizationAndClientData.client, oppGestion.employee);
                        //respFinal = respFinal + "\\n" + "Creación de LDR o subir archivos a SAP y FTP del cliente: " + oppGestion.organizationAndClientData.client;

                    }

                    #endregion

                }


            }
            catch (Exception e)
            {
                //Notificar error al usuario.
                NotifySuccessOrErrors("Fail", 2, errorsList);

                string msg = "Este error está en el try catch de la Fase 2: Crear LDRS - Autopp.";
                sett.SendError(this.GetType(), $"Error al crear LDR id #{idOpp}", msg, e);

                string updateQuery = $"UPDATE OppRequests SET status = 7, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                crud.Update(updateQuery, "autopp2_db", enviroment);

                process.KillProcess("EXCEL", true);
                //Desbloquear RPA User
                sap.BlockUser(sapSystem, 0, mandante);

                return false;
            }

            return true;

            #endregion
        }


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
            console.WriteLine("*Fase 3: Crear requisitos BAW*");
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

                //log.LogDeCambios("Creación", root.BDProcess, oppGestion.employee, "Creación de BAW", "Creación de requerimientos de BAW del cliente: " + oppGestion.organizationAndClientData.client, oppGestion.employee);
                //respFinal = respFinal + "\\n" + "Creación de requerimientos de BAW del cliente: " + oppGestion.organizationAndClientData.client;





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


        /// <summary>
        /// En este paso, verifica la carpeta de BAW Notificaciones y en base al subject, realiza 3 tipos de acciones:
        /// 4- Recibir y verificar la solicitud: Da por finalizado el número de caso por el usuario.
        /// 5- Rechazado: pone el caso como rechazado y finaliza el caso.
        /// 6- Devolución: En este caso se pone el status del número de caso en devolución, para que el vendedor lo corrija en S&S y mande nuevamente la solicitud al especialista.
        /// </summary>
        /// <returns>No retorna ningún valor.</returns>
        public void Step456EmailsActionsBAW()
        {
            bool displayedTitle = false;

            bool getConfiguration = true;

            //Recorre cada una de las notificaciones BAW en cola.
            while (mail.GetAttachmentEmail("BAW Notificaciones", "Procesados", "Procesados BAW"))
            {
                if (getConfiguration)
                {
                    getConfiguration = false;
                    GetAutoppConfiguration();
                }

                //Subject de ejemplo:"Recibir y verificar solicitud GTL. Número de oportunidad:0000178752. Número de caso: 5042-7. Proveedor: Lenovo"
                idOpp = 0;
                string subject = root.Subject.Replace(" ", "");
                string body = root.Email_Body;

                #region Fase 4- Aprobación final BAW
                //Si es ese subject, realice el proceso.
                if (subject.Contains("ha sido finalizada".Replace(" ", "")))
                {
                    executeStats = true;

                    //Para mostrarlo una sola vez.
                    if (!displayedTitle)
                    {
                        console.WriteLine("");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("*Fase 4-5-6: Leyendo solicitudes de BAW Notificaciones*");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("");

                        displayedTitle = true;
                    }

                    console.WriteLine("");
                    console.WriteLine("Fase 4: Aprobación final BAW");
                    console.WriteLine("");

                    string oppNumber = "";
                    caseNumber = "";

                    //Desestructura para averiguar el número de oportunidad y número de caso.
                    try
                    {
                        //oppNumber = subject.Split('.')[1].Split(':')[1];

                        int init = subject.LastIndexOf("caso#") + 5;
                        int finish = subject.IndexOf(")hasido");
                        caseNumber = subject.Substring(init, finish - init);

                        //caseNumber = subject.Split('.')[2].Split(':')[1];
                    }
                    catch (Exception e)
                    {
                        string msg = $"Error al extraer el número de oportunidad en el subject, en aprobación final BAW. El subject es: {subject}";
                        sett.SendError(this.GetType(), $"Error al extraer el número de oportunidad", msg, e);
                        return; //Para que se salga de esta iteración
                    }



                    #region Obtener el registro según el número de opp.

                    DataRow oppRow;
                    //Extraer el oppRequest
                    try
                    {
                        string sqlOpp = $"select * from OppRequests WHERE id= (select oppId from BAW where caseNumber='{caseNumber}'and active =1) and active=1 ";
                        oppRow = crud.Select(sqlOpp, "autopp2_db", enviroment).Rows[0];
                        oppNumber = oppRow["opp"].ToString();
                        console.WriteLine($"Se va a proceder con la aprobación final de la oportunidad {oppNumber}");
                    }
                    catch (Exception e)
                    {
                        console.WriteLine($"La oportunidad {oppNumber} - relacionada al case number {caseNumber} no existe en los registros internos del robot.");
                        return;
                    }
                    idOpp = (int)oppRow["id"];
                    #endregion

                    #region OrganizationAndClientData
                    DataTable orgInfo = oppInfo(idOpp, "organizationAndClientData");
                    OrganizationAndClientData organizationAndClientData = new OrganizationAndClientData();
                    organizationAndClientData.client =/* "00" +*/ orgInfo.Rows[0]["idClient"].ToString().PadLeft(10, '0');
                    #endregion

                    #region Objeto principal donde se une toda la información
                    oppGestion = new AutoppInformation();
                    oppGestion.id = oppRow["id"].ToString();
                    oppGestion.employee = oppRow["createdBy"].ToString();
                    oppGestion.opp = oppRow["opp"].ToString();
                    oppGestion.organizationAndClientData = organizationAndClientData;
                    #endregion

                    #region Empleado que creó la oportunidad.
                    //employeeCreatorData = new CCEmployee(oppGestion.employee);
                    //oppGestion.employee = employeeCreatorData.IdEmpleado;

                    string sqlEmployeeCreatorData = $"select * from MIS.digital_sign where user='{oppGestion.employee}'";
                    employeeCreatorData = crud.Select(sqlEmployeeCreatorData, "MIS", enviroment).Rows[0];

                    #endregion

                    #region Empleado con rol de empleado responsable en la oportunidad.

                    string sqlEmployeeResponsible4 = $"select * from MIS.digital_sign where id=(SELECT employee FROM autopp2_db.SalesTeam where role= 41 and oppId= {oppGestion.id})";
                    //DataRow employeeResponsibleDt = crud.Select( sqlEmployeeResponsible4, "databot_db", enviroment).Rows[0];
                    employeeResponsibleData = crud.Select(sqlEmployeeResponsible4, "databot_db", enviroment).Rows[0];

                    //employeeResponsibleData = new CCEmployee(employeeResponsibleDt["user"].ToString());

                    #endregion

                    #region Extraer el nombre del cliente.
                    string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {oppGestion.organizationAndClientData.client}";
                    client = crud.Select(sqlClient, "databot_db", enviroment).Rows[0];
                    #endregion




                    if (oppRow != null)
                    {
                        /*#region Actualizar los registros de status 2 a 3(rechazado), debido a que no se sabe cuando el especialista rechaza, entonces si recibe un correo de aprobación se cambia.
                        string sqlUpdate = $"UPDATE `BAW` SET `statusBAW`=3  WHERE oppId= (SELECT oppR.id FROM OppRequests oppR WHERE oppR.opp= '{oppNumber}') and statusBAW=2";
                        crud.Update(sqlUpdate, "autopp2_db", enviroment);
                        #endregion*/

                        console.WriteLine("");
                        console.WriteLine($"Procesando solicitud # {oppRow["id"]} - {oppNumber}  Case Number # {caseNumber}");

                        IWebDriver chrome4 = null;
                        try
                        {
                            chrome4 = SelConn(GetUrlBaw());

                            #region Aprobación final en Baw 

                            #region Login en Baw
                            //console.WriteLine("Iniciando sesión en baw...");

                            chrome4.FindElement(By.Id("username")).SendKeys(cred.user_baw);
                            chrome4.FindElement(By.Id("password")).SendKeys(cred.password_baw);

                            chrome4.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/a[1]")).Submit();
                            System.Threading.Thread.Sleep(6000);

                            console.WriteLine("Inicio de sesión exitoso.");
                            #endregion

                            #region Buscar la oportunidad 
                            console.WriteLine($"Buscando la oportunidad...");
                            //Ingresar la oportunidad a buscar
                            try
                            {
                                try { new WebDriverWait(chrome4, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input"))); } catch { }
                                chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber.Trim());
                            }
                            catch (Exception b)
                            {
                                System.Threading.Thread.Sleep(35000);
                                chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber.Trim());
                            }
                            //Click buscar 
                            chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();

                            System.Threading.Thread.Sleep(1500);

                            bool continueAprobation = false;

                            bool exceptionError = false;
                            //Primer registro
                            try
                            {
                                //Intentar tocar el primer registro.
                                try { new WebDriverWait(chrome4, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a"))); } catch { }
                                chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                                //System.Threading.Thread.Sleep(1000);
                                System.Threading.Thread.Sleep(1000);
                                console.WriteLine("Encontrada.");
                                continueAprobation = true;
                            }
                            catch (Exception e)
                            {
                                try
                                {
                                    //Segundo intento de tocar el primer registro
                                    chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();
                                    System.Threading.Thread.Sleep(4000);
                                    chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                                    System.Threading.Thread.Sleep(1000);
                                    console.WriteLine("Encontrada.");
                                    continueAprobation = true;



                                }
                                catch (Exception k)
                                {
                                    #region Notificación de que el robot no encontró el registro del case number, en otras palabras por alguna razón BAW lo cerró de forma automática y no necesitó al robot para recibir y verificar requerimientos, este caso está en estudio por parte de los Databot-Developers.

                                    string pathError = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                                        @"\databot\Autopp\ErrorsScreenshots\FinalApprovalsNotFounds\" + $"{oppRow["id"]} - {oppRow["opp"]} - aprobación no encontrada.png";

                                    Screenshot TakeScreenshot = ((ITakesScreenshot)chrome4).GetScreenshot();
                                    TakeScreenshot.SaveAsFile(pathError);


                                    string[] cc = { "epiedra@gbm.net"/*, oppRow["createdBy"] + "@gbm.net",employeeResponsibleData["email"].ToString() */};
                                    string[] att = { pathError };
                                    string bodyEr = process.greeting() + $"\r\nEl robot no logró encontrar el requerimiento número {caseNumber} de la oportunidad {oppRow["opp"]} - {oppRow["id"]}" +
                                        $"para realizar la aprobación final en BAW. Por favor revisar si se cerró el caso adecuadamente. ";

                                    mail.SendHTMLMail(bodyEr, new string[] { "epiedra@gbm.net" }, $"No se encontró el requerimiento de aprobación final - {oppRow["opp"]} - Autopp", cc, att);

                                    #endregion

                                    //Noticar al usuario por Webex Teams
                                    //NotifySuccessOrErrors("Fail", 4);

                                    //string updateQueryBaw =
                                    // $"UPDATE `BAW` SET `statusBAW`=8 WHERE caseNumber= '{caseNumber.Trim()}'; " +
                                    // $"UPDATE OppRequests SET status = 9, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {oppRow["id"]} ;";
                                    //crud.Update(updateQueryBaw, "autopp2_db", enviroment);

                                    console.WriteLine($"BAW no encontró la aprobación del case number: {caseNumber}, de la oportunidad {oppRow["id"]} - {oppNumber}, revisar si el requerimiento cerró adecuadamente.");

                                    #region Cerrar Chrome
                                    process.KillProcess("chromedriver", true);
                                    process.KillProcess("chrome", true);
                                    #endregion

                                    exceptionError = true;
                                    //return;


                                }
                            }

                            #endregion

                            #region Ejecutar la aprobación final BAW
                            if (continueAprobation == true && exceptionError == false) //Si encontró el requerimiento en el buscador.
                            {
                                console.WriteLine("Realizando la aprobación final en BAW.");

                                try
                                {

                                    System.Threading.Thread.Sleep(500);
                                    //Cambio a IFrame
                                    IWebElement iframe = chrome4.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/div[4]//following-sibling::iframe[1]"));
                                    chrome4.SwitchTo().Frame(iframe);

                                }
                                catch (Exception)
                                {
                                    System.Threading.Thread.Sleep(1000);
                                    //Cambio a IFrame
                                    IWebElement iframe = chrome4.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/div[4]//following-sibling::iframe[1]"));
                                    chrome4.SwitchTo().Frame(iframe);

                                }

                                //System.Threading.Thread.Sleep(500);
                                //Seleccionar aceptado
                                try { new WebDriverWait(chrome4, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='combo_div_3']"))); } catch { }
                                SelectElement aux1 = new SelectElement(chrome4.FindElement(By.XPath("//*[@id='combo_div_3']")));
                                aux1.SelectByValue("aceptado");

                                //Botón final verde de continuar y finalizar 1 etapa de gestión en espera de respuesta del especialista.
                                try
                                {
                                    //System.Threading.Thread.Sleep(500);
                                    try { new WebDriverWait(chrome4, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_12_1_4_1_1_1_1']/div/button"))); } catch { }
                                    chrome4.FindElement(By.XPath("//*[@id='div_12_1_4_1_1_1_1']/div/button")).Click();

                                }
                                catch (Exception)
                                {
                                    System.Threading.Thread.Sleep(1500);
                                    chrome4.FindElement(By.XPath("//*[@id='div_12_1_4_1_1_1_1']/div/button")).Click();
                                }

                                System.Threading.Thread.Sleep(2500);
                                //console.WriteLine($"Aprobación de oportunidad {oppNumber} finalizada con éxito.");

                            }
                            #region Cerrar Chrome
                            process.KillProcess("chromedriver", true);
                            process.KillProcess("chrome", true);
                            #endregion

                            #endregion



                            #endregion



                            if (continueAprobation)
                                console.WriteLine($"Case number {caseNumber} aprobado con éxito.");

                            #region Actualizar status y log de cambios 

                            string updateQuery =
                                $"UPDATE `BAW` SET `statusBAW`=5 WHERE oppId= (SELECT oppR.id FROM OppRequests oppR WHERE oppR.opp= '{oppNumber.Trim()}') and caseNumber= '{caseNumber.Trim()}'; ";
                            crud.Update(updateQuery, "autopp2_db", enviroment);

                            #region Actualizar el nombre del especialista 
                            string bodyE = root.Email_Body;
                            int init = bodyE.LastIndexOf("Especialista GTL") + 16;
                            int finish = bodyE.IndexOf("Fecha de ingreso");
                            string specialist = bodyE.Substring(init, finish - init);
                            string sqlUpd = $"UPDATE BAW SET specialist= '{specialist}' WHERE caseNumber='{caseNumber}';";
                            crud.Update(sqlUpd, "autopp2_db", enviroment);
                            #endregion

                            //Consulta cuantos requerimientos BAW están pendientes aún de proceso completado
                            DataTable countBAWReqPendingsTable = new DataTable();
                            string sql1 = $"SELECT COUNT(*) countPendings FROM `BAW` WHERE oppId={oppGestion.id} and statusBAW!=5 ";
                            countBAWReqPendingsTable = crud.Select(sql1, "autopp2_db", enviroment);

                            string countPendings = countBAWReqPendingsTable.Rows[0]["countPendings"].ToString();


                            if (countPendings == "0") //No hay ninguno pendiente.
                            {
                                updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=5 WHERE id= {oppGestion.id}; ";
                                crud.Update(updateQuery, "autopp2_db", enviroment);
                            }

                            //Notificación de aprobación final de Requerimiento BAW
                            NotifySuccessOrErrors("Success", 4);


                            // log.LogDeCambios("Verficación final", root.BDProcess, oppGestion.employee, "Recibir y verificar solicitud", "Realizar la verificación final para finalizar el proceso del client y opp: " + oppGestion.organizationAndClientData.client + " -  " + oppGestion.opp, oppGestion.employee);
                            //respFinal = respFinal + "\\n" + "Realizar la verificación final para finalizar el proceso del client y opp: " + oppGestion.organizationAndClientData.client + " -  " + oppGestion.opp;
                            #endregion


                        }
                        catch (Exception e)
                        {
                            string msg = $"Este error está en el try catch de la Fase 4: Aprobación final BAW - Autopp. Revisar el screenshot del error en la carpeta en el desktop/Databot/Autopp/ErrorsScreenshots" +
                            $"La oppNumber es: {oppNumber} y  su Case number es: {caseNumber} ";

                            sett.SendError(this.GetType(), $"Error al realizar la aprobación final BAW #{oppRow["id"]}", msg, e);

                            string updateQuery =
                                $"UPDATE `BAW` SET `statusBAW`=8 WHERE caseNumber= '{caseNumber.Trim()}'; " +
                                $"UPDATE OppRequests SET status = 9, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {oppRow["id"]} ;";
                            crud.Update(updateQuery, "autopp2_db", enviroment);

                            try
                            {
                                string pathErrors = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\databot\Autopp\ErrorsScreenshots\";
                                Screenshot TakeScreenshot = ((ITakesScreenshot)chrome4).GetScreenshot();
                                TakeScreenshot.SaveAsFile(pathErrors + $"{oppRow["id"]} - {oppRow["opp"]} - error en aprobación final.png");
                            }
                            catch (Exception i)
                            {
                                console.WriteLine("No se pudo guardar el screenshot del error. El error al guardar fue:");
                                console.WriteLine("");
                                console.WriteLine(i.ToString());
                            }
                        }


                    }

                    #region Cerrar Chrome
                    process.KillProcess("chromedriver", true);
                    process.KillProcess("chrome", true);
                    #endregion



                }
                #endregion

                #region Fase 5- Devolución requerimiento de BAW
                else if (subject.Contains("debe ser corregido".Replace(" ", "")))
                {
                    executeStats = true;

                    //Para mostrarlo una sola vez.
                    if (!displayedTitle)
                    {
                        console.WriteLine("");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("*Fase 4-5-6: Leyendo solicitudes de BAW Notificaciones*");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("");

                        displayedTitle = true;
                    }

                    console.WriteLine("");
                    console.WriteLine("Fase 5: Devolución de caso BAW");
                    console.WriteLine("");

                    try
                    {
                        int init = subject.LastIndexOf("GTL#") + 4;
                        int finish = subject.IndexOf("debeser");
                        caseNumber = subject.Substring(init, finish - init);


                        string sqlMain = $@"

                        SELECT
                        opp.id as gestion, opp.opp as oppNumber, cli.name as client,  baw.createdBy, mis.user as employeeResponsible

                        FROM BAW baw, OrganizationAndClientData org, databot_db.clients cli, SalesTeam st, MIS.digital_sign mis, OppRequests opp

                        WHERE
                        baw.caseNumber = '{caseNumber}'
                        AND org.oppId = baw.oppId
                        AND org.client = cli.id
                        AND st.oppId = baw.oppId
                        AND st.role = 41
                        AND mis.id = st.employee
                        AND baw.oppId = opp.id
                     ";



                        DataTable dtMain = crud.Select(sqlMain, "autopp2_db", enviroment);

                        if (dtMain.Rows.Count > 0)
                        {
                            DataRow dr = dtMain.Rows[0];
                            string title = "Su caso ha sido devuelto";

                            #region Moldeo de mensaje
                            string msgWebex = "";
                            msgWebex = root.Email_Body.Replace("Se necesita corregir un requerimiento del GTL", $"{process.greeting()}, el caso #{caseNumber} ha sido devuelto y debe corregir un requerimiento del GTL.");
                            msgWebex = msgWebex.Replace("Oportunidad", "Oportunidad:").Replace("Cliente", "Cliente:").Replace("Proveedor", "Proveedor:").Replace("Producto", "Producto:");

                            msgWebex += $"\r\n\t\r\nPara corregir siga los siguientes pasos:" +
                             "\r\n 1. Ingresar a <a href='{routeSS}'>Smart And Simple</a> de GBM, seguidamente en Ventas - Autopp - Inicio - Mis Solicitudes." +
                             $"\r\n2. Encontrar la opp {dr["oppNumber"].ToString()}, tocar el botón BAW, hallar el número de caso {caseNumber} y tocar el botón Editar" +
                             "\r\n3. Por último, editar los datos del requerimiento BAW en base a su criterio y volverlo enviar al especialista.";
                            #endregion

                            //Enviar mensaje al creador del caso.
                            webex.SendNotification(dr["createdBy"].ToString() + "@GBM.NET", title, msgWebex);

                            //webex.SendCCNotification(dr["createdBy"].ToString() + "@GBM.NET", title, "", "", msgWebex);


                            if (dr["createdBy"].ToString() != dr["employeeResponsible"].ToString())
                            {
                                //Enviar mensaje al empleado responsable del caso.
                                //webex.SendNotification(dr["employeeResponsible"].ToString() + "@GBM.NET", title, msgWebex);
                            }

                            webex.SendNotification("EPIEDRA@GBM.NET", title, msgWebex);

                            #region Actualizar el status del caseNumber a devuelto por especialista

                            string sqlUpdateCaseNumberAndStatus = $"UPDATE `BAW` SET `statusBAW`= 4 WHERE caseNumber= '{caseNumber}';  ";
                            sqlUpdateCaseNumberAndStatus += $"UPDATE OppRequests SET opp ='{dr["oppNumber"].ToString()}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=2 WHERE id= {dr["gestion"].ToString()}; ";
                            crud.Update(sqlUpdateCaseNumberAndStatus, "autopp2_db", enviroment);

                            //log.LogDeCambios("Verficación final", root.BDProcess, oppGestion.employee, "Se estableció en devuelto el caso BAW:" + caseNumber, "Se estableció en devuelto el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString(), oppGestion.employee);
                            //respFinal = respFinal + "\\n" + "Se estableció en devuelto el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString();

                            #endregion
                        }
                        else
                        {
                            console.WriteLine($"El caso {caseNumber} no existe en los registros internos del databot.");
                            return;
                        }



                    }
                    catch (Exception e)
                    {
                        console.WriteLine(e.ToString());

                        return; //Para que se salga de esta iteración
                    }


                }

                #endregion


                #region Fase 6 - Rechazar solicitud
                else if (subject.Contains("ha sido rechazado por el siguiente motivo".Replace(" ", "")))
                {
                    executeStats = true;

                    //Para mostrarlo una sola vez.
                    if (!displayedTitle)
                    {
                        console.WriteLine("");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("*Fase 4-5-6: Leyendo solicitudes de BAW Notificaciones*");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("");

                        displayedTitle = true;
                    }


                    console.WriteLine("");
                    console.WriteLine("Fase: Notificación que el especialista rechazó un caso.");
                    console.WriteLine("");

                    try
                    {
                        int init = subject.LastIndexOf("estecaso") + 8;
                        int finish = subject.IndexOf("hasido");
                        caseNumber = subject.Substring(init, finish - init);


                        string sqlMain = $@"

                        SELECT
                        opp.id as gestion, opp.opp as oppNumber, cli.name as client,  baw.createdBy, mis.user as employeeResponsible

                        FROM BAW baw, OrganizationAndClientData org, databot_db.clients cli, SalesTeam st, MIS.digital_sign mis, OppRequests opp

                        WHERE
                        baw.caseNumber = '{caseNumber}'
                        AND org.oppId = baw.oppId
                        AND org.client = cli.id
                        AND st.oppId = baw.oppId
                        AND st.role = 41
                        AND mis.id = st.employee
                        AND baw.oppId = opp.id
                     ";



                        DataTable dtMain = crud.Select(sqlMain, "autopp2_db", enviroment);

                        if (dtMain.Rows.Count > 0)
                        {
                            DataRow dr = dtMain.Rows[0];
                            string title = "Su caso ha sido rechazado";

                            #region Moldeo de mensaje
                            string msgWebex = root.Email_Body.Replace("Caso Rechazado", $"{process.greeting()}, el caso #{caseNumber} ha sido rechazado");
                            msgWebex = msgWebex.Replace("oportunidad", "oportunidad:").Replace("Cliente", "Cliente:").Replace("Proveedor", "Proveedor:").Replace("Requerimiento", "Requerimiento:").Replace("Producto", "Producto:").Replace("rechazo", "rechazo:").Replace("Comentarios", "Comentarios:");

                            #endregion

                            //Enviar mensaje al creador del caso.
                            webex.SendNotification(dr["createdBy"].ToString() + "@GBM.NET", title, msgWebex);

                            //webex.SendCCNotification(dr["createdBy"].ToString() + "@GBM.NET", title, "", "", msgWebex);


                            if (dr["createdBy"].ToString() != dr["employeeResponsible"].ToString())
                            {
                                //Enviar mensaje al empleado responsable del caso.
                                //webex.SendNotification(dr["employeeResponsible"].ToString() + "@GBM.NET", title, msgWebex);
                            }

                            webex.SendNotification("EPIEDRA@GBM.NET", title, msgWebex);

                            #region Actualizar el status del caseNumber a rechazado por especialista

                            string sqlUpdateCaseNumberAndStatus = $"UPDATE `BAW` SET `statusBAW`= 3 WHERE caseNumber= '{caseNumber}'; ";
                            sqlUpdateCaseNumberAndStatus += $"UPDATE OppRequests SET opp ='{dr["oppNumber"].ToString()}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=3 WHERE id= {dr["gestion"].ToString()}; ";
                            crud.Update(sqlUpdateCaseNumberAndStatus, "autopp2_db", enviroment);

                            //log.LogDeCambios("Rechazo de caso BAW", root.BDProcess, oppGestion.employee, "Se estableció en rechazado el caso BAW:" + caseNumber, "Se estableció en rechazado el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString(), oppGestion.employee);
                            //respFinal = respFinal + "\\n" + "Se estableció en rechazado el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString();

                            #endregion

                        }
                        else
                        {
                            console.WriteLine($"El caso {caseNumber} no existe en los registros internos del databot.");
                            return;
                        }



                    }
                    catch (Exception e)
                    {
                        console.WriteLine(e.ToString());
                    }


                }

                #endregion


                #region Notificación de cual especialista tomó el caso 
                else if (subject.Contains("ha sido asignado al especialista".Replace(" ", "")))
                {
                    executeStats = true;

                    //Para mostrarlo una sola vez.
                    if (!displayedTitle)
                    {
                        console.WriteLine("");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("*Fase 4-5-6: Leyendo solicitudes de BAW Notificaciones*");
                        console.WriteLine("*******************************************************");
                        console.WriteLine("");

                        displayedTitle = true;
                    }


                    console.WriteLine("");
                    console.WriteLine("Fase: Notificación de cual especialista tiene el caso.");
                    console.WriteLine("");

                    try
                    {
                        int init = subject.LastIndexOf("Sucaso") + 6;
                        int finish = subject.IndexOf("hasido");
                        caseNumber = subject.Substring(init, finish - init);


                        string sqlMain = $@"

                        SELECT
                        opp.id as gestion, opp.opp as oppNumber, cli.name as client,  baw.createdBy, mis.user as employeeResponsible

                        FROM BAW baw, OrganizationAndClientData org, databot_db.clients cli, SalesTeam st, MIS.digital_sign mis, OppRequests opp

                        WHERE
                        baw.caseNumber = '{caseNumber}'
                        AND org.oppId = baw.oppId
                        AND org.client = cli.id
                        AND st.oppId = baw.oppId
                        AND st.role = 41
                        AND mis.id = st.employee
                        AND baw.oppId = opp.id
                     ";



                        DataTable dtMain = crud.Select(sqlMain, "autopp2_db", enviroment);

                        if (dtMain.Rows.Count > 0)
                        {
                            DataRow dr = dtMain.Rows[0];
                            string title = "Su caso ha sido asignado al siguiente especialista";

                            #region Moldeo de mensaje
                            string msgWebex = root.Email_Body.Replace("Caso asignado", $"{process.greeting()}, el caso #{caseNumber} ha sido asignado");
                            msgWebex.Replace("oportunidad", "oportunidad:").Replace("Cliente", "Cliente:").Replace("Proveedor", "Proveedor:").Replace("Tipo de Requerimiento", "Tipo de Requerimiento:").Replace("Cantidad", "Cantidad:").Replace("integración", "integración:").Replace("Comentarios", "Comentarios:");

                            #endregion

                            //Enviar mensaje al creador del caso.
                            webex.SendNotification(dr["createdBy"].ToString() + "@GBM.NET", title, msgWebex);

                            //webex.SendCCNotification(dr["createdBy"].ToString() + "@GBM.NET", title, "", "", msgWebex);


                            if (dr["createdBy"].ToString() != dr["employeeResponsible"].ToString())
                            {
                                //Enviar mensaje al empleado responsable del caso.
                                //webex.SendNotification(dr["employeeResponsible"].ToString() + "@GBM.NET", title, msgWebex);
                            }

                            webex.SendNotification("EPIEDRA@GBM.NET", title, msgWebex);

                            #region Actualizar el nombre del especialista 
                            string bodyE = root.Email_Body;
                            init = bodyE.LastIndexOf("por el especialista ") + 20;
                            finish = bodyE.IndexOf(", el tiempo");
                            string specialist = bodyE.Substring(init, finish - init);
                            string sqlUpd = $"UPDATE BAW SET specialist= '{specialist}' WHERE caseNumber='{caseNumber}';";
                            crud.Update(sqlUpd, "autopp2_db", enviroment);
                            #endregion

                        }
                        else
                        {
                            console.WriteLine($"El caso {caseNumber} no existe en los registros internos del databot.");
                            return;
                        }



                    }
                    catch (Exception e)
                    {
                        console.WriteLine(e.ToString());
                    }

                }



                #endregion



            }



        }

        /// <summary>
        /// Este paso es para un requerimiento BAW que fue editado en S&S por un vendedor, el motivo de la edición es que
        /// el requerimiento fue devuelto por el especialista, y el vendedor debe editarlo y enviarlo nuevamente para su aprobación.
        /// </summary>
        public void Step7ForwardBAWRequirement()
        {
            //Consulta las solicitudes BAW pendientes a reenviar al especialista.
            idOpp = 0;
            DataTable reqsForwardBaw = oppInfo(0, "BAW");
            bool getConfiguration = true;

            if (reqsForwardBaw.Rows.Count > 0)
            {
                executeStats = true;

                if (getConfiguration)
                {
                    getConfiguration = false;
                    GetAutoppConfiguration();
                }

                int indexReqsForward1 = 1;

                bool loginOpened7 = false;
                IWebDriver chrome = null;
                //Establecer conexión en BAW
                chrome = SelConn(GetUrlBaw());
                //console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

                console.WriteLine("");
                console.WriteLine("*******************************************************");
                console.WriteLine("*Fase 7: Reenviar requerimiento de BAW al especialista*");
                console.WriteLine("*******************************************************");
                console.WriteLine("");

                foreach (DataRow reqForwardBaw in reqsForwardBaw.Rows)
                {

                    console.WriteLine($"Procesando solicitud {indexReqsForward1} de {reqsForwardBaw.Rows.Count} solicitudes.");

                    try
                    {
                        #region Obtener el registro según el número de opp.

                        DataRow oppRow;
                        caseNumber = reqForwardBaw["caseNumber"].ToString();
                        //Extraer el oppRequest
                        try
                        {
                            string sqlOpp = $"SELECT * FROM `OppRequests` WHERE id='{reqForwardBaw["oppId"]}' and active=1";
                            oppRow = crud.Select(sqlOpp, "autopp2_db", enviroment).Rows[0];
                        }
                        catch (Exception e)
                        {
                            console.WriteLine($"Sucedió un problema al extraer la oportunidad de BAW." + e.ToString());
                            return;
                        }
                        idOpp = (int)oppRow["id"];
                        #endregion

                        #region OrganizationAndClientData
                        DataTable orgInfo = oppInfo(idOpp, "organizationAndClientData");
                        OrganizationAndClientData organizationAndClientData = new OrganizationAndClientData();
                        organizationAndClientData.client =/* "00" +*/ orgInfo.Rows[0]["idClient"].ToString().PadLeft(10, '0');
                        #endregion

                        #region Objeto principal donde se une toda la información
                        oppGestion = new AutoppInformation();
                        oppGestion.id = oppRow["id"].ToString();
                        oppGestion.employee = oppRow["createdBy"].ToString();
                        oppGestion.opp = oppRow["opp"].ToString();
                        oppGestion.organizationAndClientData = organizationAndClientData;
                        #endregion

                        #region Empleado que creó la oportunidad.
                        string sqlEmployeeCreatorData = $"select * from MIS.digital_sign where user='{oppGestion.employee}'";
                        employeeCreatorData = crud.Select(sqlEmployeeCreatorData, "MIS", enviroment).Rows[0];
                        #endregion

                        #region Empleado con rol de empleado responsable en la oportunidad.
                        string sqlEmployeeResponsible4 = $"select * from MIS.digital_sign where id=(SELECT employee FROM autopp2_db.SalesTeam where role= 41 and oppId= {oppGestion.id})";
                        employeeResponsibleData = crud.Select(sqlEmployeeResponsible4, "databot_db", enviroment).Rows[0];
                        #endregion

                        #region Extraer el nombre del cliente.
                        string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {oppGestion.organizationAndClientData.client}";
                        client = crud.Select(sqlClient, "databot_db", enviroment).Rows[0];
                        #endregion

                        console.WriteLine("");
                        console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

                        #region Agregar requerimientos en BAW 


                        #region Login en BAW.
                        //console.WriteLine("Iniciando sesión en baw...");
                        if (!loginOpened7)
                        {

                            chrome.FindElement(By.Id("username")).SendKeys(cred.user_baw);
                            chrome.FindElement(By.Id("password")).SendKeys(cred.password_baw);

                            chrome.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/a[1]")).Submit();
                            //System.Threading.Thread.Sleep(7000);
                            loginOpened7 = true;
                            console.WriteLine("Inicio de sesión exitoso.");
                        }
                        #endregion

                        #region Buscar la oportunidad 
                        //console.WriteLine($"Buscando la oportunidad ...");
                        //Ingresar la oportunidad a buscar
                        try
                        {
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input"))); } catch { }
                            //System.Threading.Thread.Sleep(5000);
                            chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                System.Threading.Thread.Sleep(5000);
                                chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber);
                            }
                            catch (Exception)
                            {
                                System.Threading.Thread.Sleep(10000);
                                chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber);
                            }


                        }
                        //Click buscar               

                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();

                        System.Threading.Thread.Sleep(6000);

                        //Primer registro
                        try
                        {
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a"))); } catch { }

                            chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                        }
                        catch (Exception e)
                        {
                            #region Notificación de que el robot no encontró el registro del case number para realizar el reenvío de requerimiento despues de devolucion de especialista, en otras palabras por alguna razón BAW lo cancelo de forma automática y no necesitó al robot para reenviar el requerimiento, este caso está en estudio por parte de los Databot-Developers.

                            string pathError = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                                @"\databot\Autopp\ErrorsScreenshots\ForwardCaseNumberNotFound\" + $"{oppRow["id"]} - {oppRow["opp"]} - requerimiento de reenvio no encontrado.png";

                            Screenshot TakeScreenshot = ((ITakesScreenshot)chrome).GetScreenshot();
                            TakeScreenshot.SaveAsFile(pathError);


                            //string[] cc = { "dmeza@gbm.net", new string[] { "epiedra@gbm.net" }, oppRow["createdBy"] + "@gbm.net", /*employeeResponsibleData.Correo*/ employeeCreatorData["email"].ToString(), employeeCreatorData["email"].ToString() };
                            //string[] att = { pathError };
                            //string body = process.greeting() + $"\r\nEl robot no logró encontrar el requerimiento número {caseNumber} de la oportunidad {oppRow["opp"]} - {oppRow["id"]}" +
                            //    $"para realizar la corrección pedida por el especialista en BAW después de la devolución. Por favor revisar si se cerró el caso adecuadamente. ";

                            //mail.SendHTMLMail(body, new string[] {"appmanagement@gbm.net"}, $"No se encontró el requerimiento para reenviar CaseNumber: #{caseNumber} de - {oppRow["opp"]} - Autopp", cc, att);

                            string[] cc = { "epiedra@gbm.net"/*, oppRow["createdBy"] + "@gbm.net", employeeCreatorData["email"].ToString()*/ };
                            string[] att = { pathError };
                            string body = $"\r\nEl robot no logró encontrar el requerimiento número {caseNumber} de la oportunidad {oppRow["opp"]} - {oppRow["id"]}" +
                                $"para realizar la corrección pedida por el especialista en BAW después de la devolución. Por favor revisar el caso adecuadamente. ";

                            mail.SendHTMLMail(body, new string[] { "epiedra@gbm.net" }, $"No se encontró el requerimiento para reenviar CaseNumber: #{caseNumber} de - {oppRow["opp"]} - Autopp", cc, att);

                            #endregion

                            //Noticar al usuario por Webex Teams
                            NotifySuccessOrErrors("Fail", 7);

                            string updateQueryBaw =
                             $"UPDATE `BAW` SET `statusBAW`=8 WHERE caseNumber= '{caseNumber.Trim()}'; " +
                             $"UPDATE OppRequests SET status = 12, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {oppRow["id"]} ;";
                            crud.Update(updateQueryBaw, "autopp2_db", enviroment);

                            console.WriteLine($"BAW no encontró el requerimiento para realizar la corrección despues de devolución del especialista del case number: {caseNumber}, de la oportunidad {oppRow["id"]} - {oppGestion.opp}, revisar el requerimiento adecuadamente.");

                            #region Cerrar Chrome
                            process.KillProcess("chromedriver", true);
                            process.KillProcess("chrome", true);
                            #endregion

                            return;

                        }

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



                        SelectElement aux;

                        //Proveedor
                        try
                        { //El primer intento dura más de lo normal
                            try
                            {
                                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 45)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='combo_div_7']"))); } catch { }
                                //System.Threading.Thread.Sleep(2500);
                                aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_7']")));
                                aux.SelectByValue(reqForwardBaw["vendor"].ToString().Trim());
                            }
                            catch (Exception)
                            {
                                System.Threading.Thread.Sleep(2000);
                                aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_7']")));
                                aux.SelectByValue(reqForwardBaw["vendor"].ToString().Trim());
                            }
                        }
                        catch (Exception)
                        {
                            System.Threading.Thread.Sleep(15000);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_7']")));
                            aux.SelectByValue(reqForwardBaw["vendor"].ToString().Trim());
                        };


                        //Nombre de Producto
                        try
                        {
                            System.Threading.Thread.Sleep(1500);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_8']")));
                            aux.SelectByValue(reqForwardBaw["productName"].ToString().Trim());
                        }
                        catch (Exception)
                        {
                            System.Threading.Thread.Sleep(3500);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_8']")));
                            aux.SelectByValue(reqForwardBaw["productName"].ToString().Trim());
                        }


                        //Tipo de Requerimiento

                        try
                        {
                            System.Threading.Thread.Sleep(200);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_9']"))); //*[@id="combo_div_9"]
                            aux.SelectByValue(reqForwardBaw["requirementType"].ToString().Trim());
                        }
                        catch (Exception e)
                        {

                            try
                            {
                                System.Threading.Thread.Sleep(1000);
                                aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_9']")));
                                aux.SelectByValue(reqForwardBaw["requirementType"].ToString().Trim());

                            }
                            catch (Exception)
                            {
                                System.Threading.Thread.Sleep(7000);
                                aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_9']")));
                                aux.SelectByValue(reqForwardBaw["requirementType"].ToString().Trim());

                            }
                        }

                        //Cantidad
                        try
                        {
                            System.Threading.Thread.Sleep(200);
                            chrome.FindElement(By.Id("input_div_6_1_3_1_1")).Clear();//*[@id="input_div_6_1_3_1_1"]
                            System.Threading.Thread.Sleep(100);
                            chrome.FindElement(By.Id("input_div_6_1_3_1_1")).SendKeys(reqForwardBaw["quantity"].ToString());//*[@id="input_div_6_1_3_1_1"]
                                                                                                                            //Checkbox Documentos adjuntados.
                                                                                                                            //chrome.FindElement(By.XPath("//*[@id='div_7']/div/div[2]/label/input")).Click();
                        }
                        catch (Exception)
                        {
                            System.Threading.Thread.Sleep(450);
                            chrome.FindElement(By.Id("input_div_6_1_3_1_1")).Clear();//*[@id="input_div_6_1_3_1_1"]
                            System.Threading.Thread.Sleep(100);
                            chrome.FindElement(By.Id("input_div_6_1_3_1_1")).SendKeys(reqForwardBaw["quantity"].ToString());//*[@id="input_div_6_1_3_1_1"]
                                                                                                                            //Checkbox Documentos adjuntados.
                                                                                                                            //chrome.FindElement(By.XPath("//*[@id='div_7']/div/div[2]/label/input")).Click();
                        }


                        //Averiguar si checkbox ¿Es parte de una integración? esta seleccionado
                        string integrationSelected = chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/label/input")).GetAttribute("checked");



                        //SI Es parte de una integración
                        if (reqForwardBaw["isIntegration"].ToString() == "Si")
                        {
                            //Anteriormente el checkbox estaba apagado
                            if (integrationSelected == null)
                            {
                                try
                                {
                                    System.Threading.Thread.Sleep(50);
                                    //Checkbox Parte de una integración
                                    //chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();
                                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/label/input")).Click();

                                }
                                catch (Exception)
                                {
                                    System.Threading.Thread.Sleep(650);
                                    //Checkbox Parte de una integración
                                    //chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();
                                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/label/input")).Click();

                                }
                            }
                        }
                        //NO Es parte de una integración
                        else
                        {
                            //Anteriormente el checkbox estaba encendido
                            if (integrationSelected == "true")
                            {
                                try
                                {
                                    System.Threading.Thread.Sleep(50);
                                    //Checkbox Parte de una integración
                                    //chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();
                                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/label/input")).Click();

                                }
                                catch (Exception)
                                {
                                    System.Threading.Thread.Sleep(650);
                                    //Checkbox Parte de una integración
                                    //chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();
                                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/label/input")).Click();

                                }

                            }
                        }
                        //Comentarios 
                        chrome.FindElement(By.XPath("//*[@id='textArea_div_11']")).Clear();
                        System.Threading.Thread.Sleep(100);
                        chrome.FindElement(By.XPath("//*[@id='textArea_div_11']")).SendKeys(reqForwardBaw["comments"].ToString());

                        System.Threading.Thread.Sleep(300);
                        //Botón de Continuar
                        chrome.FindElement(By.XPath("//*[@id='div_14_1_4_1_1_1_1']/div/button")).Click();


                        //Botón final verde de continuar y finalizar 1 etapa de gestión en espera de respuesta del especialista.
                        System.Threading.Thread.Sleep(5000);
                        console.WriteLine($"Case number {caseNumber} devuelto al especialista con éxito.");
                        console.WriteLine($"");

                        System.Threading.Thread.Sleep(1500);

                        //Salir del iframe
                        chrome.SwitchTo().DefaultContent();
                        System.Threading.Thread.Sleep(1500);

                        //Limpiar el buscador 
                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/div/a/i")).Click();


                        #region Actualizar status y log de cambios   
                        //crud.Update(updateQueryBAW, "autopp2_db", enviroment);
                        //log.LogDeCambios("Creación", "Autopp", oppGestion.employee, "", oppGestion.opp, "Creación de BAW");
                        #endregion

                        //Actualizar status en aprobación por especialista
                        string updateQuery7 = $"UPDATE BAW SET statusBAW=2 WHERE caseNumber= '{caseNumber}'";
                        crud.Update(updateQuery7, "autopp2_db", enviroment);

                        //Notificación de exito
                        NotifySuccessOrErrors("Success", 7);

                        //log.LogDeCambios("Devuelto al especialista el caso BAW", root.BDProcess, oppGestion.employee, "Después de la devolución se devolvió al especialista el caso BAW:" + caseNumber, "Se devolvió al especialista el  caso BAW: " + caseNumber + " - " + oppGestion.organizationAndClientData.client + " -  " + oppGestion.opp, oppGestion.employee);
                        //respFinal = respFinal + "\\n" + "Después de la devolución se devolvió al especialista el caso BAW:" + caseNumber;


                        //Consulta cuantos requerimientos BAW están devueltos o rechazados o con error
                        DataTable countBAWReqPendingsTable = new DataTable();
                        string sql1 = $"SELECT COUNT(*) countPendings, (SELECT COUNT(*) countRejected FROM `BAW` WHERE oppId={oppGestion.id} and statusBAW=3)countRejected FROM `BAW` WHERE oppId={oppGestion.id} and (statusBAW=3 or statusBAW=4 or statusBAW=8)  ";
                        countBAWReqPendingsTable = crud.Select(sql1, "autopp2_db", enviroment);

                        string countPendings = countBAWReqPendingsTable.Rows[0]["countPendings"].ToString();
                        string countRejected = countBAWReqPendingsTable.Rows[0]["countRejected"].ToString();



                        if (countPendings == "0") //No hay pendientes.
                        {
                            updateQuery7 = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=4 WHERE id= {oppGestion.id}; ";
                            crud.Update(updateQuery7, "autopp2_db", enviroment);
                        }
                        else if (countRejected != "0") //Si habia un rechazado, que ponga el status de la opp en rechazado
                        {
                            updateQuery7 = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=3 WHERE id= {oppGestion.id}; ";
                            crud.Update(updateQuery7, "autopp2_db", enviroment);

                        }



                        #endregion

                        #endregion



                    }
                    catch (Exception e)
                    {
                        //Cerrar chrome
                        process.KillProcess("chromedriver", true);
                        process.KillProcess("chrome", true);

                        NotifySuccessOrErrors("Fail", 7, null, e);

                        #region Actualizar el status del caseNumber a error
                        string updateQuery =
                            $"UPDATE `BAW` SET `statusBAW`=8 WHERE caseNumber= '{caseNumber.Trim()}'; " +
                            $"UPDATE OppRequests SET status = 12, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {oppGestion.id} ;";
                        crud.Update(updateQuery, "autopp2_db", enviroment);

                        #endregion

                        chrome = SelConn(GetUrlBaw());
                    }



                    indexReqsForward1++;

                }


                process.KillProcess("chromedriver", true);
                process.KillProcess("chrome", true);
            }


        }
        #endregion



        #region Métodos útiles para la gestión de cada uno de los pasos de AutoppProcess.


        /// <summary>
        /// Método para crear una oportunidad en CRM, incluyendo información general, salesTeams, colaboradores, etc.
        /// Devuelve el # de la oportunidad generada en SAP CRM, y si da error devuelve un "".
        /// </summary>
        /// <param name="oppInformation"></param>
        /// <param name="user"></param>
        /// <returns></returns>
        private string CreateOppCRM(AutoppInformation oppInformation, string user)
        {

            string idopp = "";
            RfcDestination destination = new SapVariants().GetDestRFC(sapSystem, mandante);

            console.WriteLine("Conectado con SAP CRM - " + sapSystem);

            RfcRepository repo = destination.Repository;
            IRfcFunction func = repo.CreateFunction("ZOPP_VENTAS");
            IRfcTable general = func.GetTable("GENERAL");
            IRfcTable partners = func.GetTable("PARTNERS");
            //console.WriteLine("Llenando información general de oportunidad");
            func.SetValue("USER", oppInformation.employee);
            //func.SetValue("USER", "DMEZA");
            general.Append();
            general.SetValue("TIPO", oppInformation.generalData.typeOpportunity);
            general.SetValue("DESCRIPCION", oppInformation.generalData.description.ToUpper());
            general.SetValue("FECHA_INICIO", oppInformation.generalData.initialDate);
            general.SetValue("FECHA_FIN", oppInformation.generalData.finalDate);
            general.SetValue("FASE_VENTAS", oppInformation.generalData.cycle);
            general.SetValue("OUTSOURCING", oppInformation.generalData.outsourcing);
            general.SetValue("SALES_TYPE", oppInformation.generalData.salesType);
            general.SetValue("PORCENTAJE", "100");
            general.SetValue("REVENUE", "");
            general.SetValue("MONEDA", "USD");
            general.SetValue("GRUPO_OPP", "0001");
            general.SetValue("ORIGEN", oppInformation.generalData.sourceOpportunity);
            general.SetValue("PRIORIDAD", "4");
            //console.WriteLine("Llenando información de cliente y el equipo de ventas");
            partners.Append();
            partners.SetValue("PARTNER", oppInformation.organizationAndClientData.client);
            partners.SetValue("FUNCTION", "00000021");
            partners.Append();
            partners.SetValue("PARTNER", oppInformation.organizationAndClientData.contact);
            partners.SetValue("FUNCTION", "00000015");
            /*partners.Append();
            partners.SetValue("PARTNER", oppInformation.employee);
            partners.SetValue("FUNCTION", "00000014");*/

            if (oppInformation.salesTeams != null)
            {
                for (int i = 0; i < oppInformation.salesTeams.Count; i++)
                {
                    partners.Append();
                    partners.SetValue("PARTNER", oppInformation.salesTeams[i].employee);
                    partners.SetValue("FUNCTION", oppInformation.salesTeams[i].role);
                }
            }

            //console.WriteLine("Llenando Organización de Servicios y Ventas");
            func.SetValue("SALES_ORG", oppInformation.organizationAndClientData.salesOrganization);
            func.SetValue("SRV_ORG", oppInformation.organizationAndClientData.servicesOrganization);

            if (oppInformation.generalData.cycle == "Y3" /*Quotation*/)
            {
                //Ciclo Quotation
                //func.SetValue("USER", "DMEZA");
                func.SetValue("USER", user); //Agregado
                //func.SetValue("USER", "RPAUSER");

            }
            else
            {
                //Los demás ciclos.
                func.SetValue("USER", "RPAUSER");
            }
            //console.WriteLine("Creando la oportunidad en SAP CRM - " + mandante);
            func.Invoke(destination);

            //Éxito
            if (func.GetValue("OPP_ID").ToString() != "")
            {
                //console.WriteLine($"La oportunidad id: {oppInformation.id} ha sido creada satisfactoriamente");
                //console.WriteLine("Id de la oportunidad creada: " + func.GetValue("OPP_ID").ToString());
                console.WriteLine($"Oportunidad creada con éxito: {oppInformation.id} - {func.GetValue("OPP_ID").ToString()}");


                idopp = func.GetValue("OPP_ID").ToString();
            }
            else
            { //Fallo

                IRfcTable validate = func.GetTable("VALIDATE");

                string errorList = $"No se pudo crear la solicitud de creación de oportunidad de la gestión #{oppInformation.id}, del usuario: {user}.<br>" +
                    "A continuación se detallan los errores generados: <br>";
                for (int i = 0; i < validate.Count; i++)
                {
                    errorList = "#" + (i + 1) + " " + validate[i].GetValue("MENSAJE") + "<br>";
                }

                //El sendError imprime en consola los errores a la vez
                sett.SendError(this.GetType(), $"Error al crear la opp id #{oppInformation.id} - {user}", errorList);

                idopp = "";
            }


            #region Evaluar si se necesita en un futuro o eliminar.
            /* if (func.GetValue("RESPONSE").ToString() != "")
             {
             }
             console.WriteLine("");
             //console.WriteLine(" End of processing");

             RfcRepository repo1 = destination.Repository;
             IRfcFunction func1 = repo.CreateFunction("ZOPP_BPM");
             func1.SetValue("ID", idopp);
             func1.SetValue("OPP_TYPE", oppInformation.generalData.typeOpportunity);
             func1.Invoke(destination);
             console.WriteLine("");
             console.WriteLine(" Replicando a BPM");*/
            #endregion

            return idopp;
        }

        /// <summary>
        /// Método para crear los LDRS en base a los criterios seleccionados en S&S
        /// </summary>
        /// <param name="Management"></param>
        /// /// <returns>Returna un diccionario donde el key es el path donde esta el LDR y value el nombre del LDR</returns>
        private Dictionary<string, string> CreateLDR(AutoppInformation oppInformation)
        {
            Dictionary<string, string> fileLDRRoute = new Dictionary<string, string>();

            if (oppInformation.LDRS.Rows.Count > 0)
            {
                console.WriteLine("Creando LDR...");

                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                //Rutas 
                string pathTemplate = desktop + @"\databot\Autopp\LDR Templates" + @"\LDR Template.xlsx";
                string pathSaveLDRS = desktop + @"\databot\Autopp\FilesToUpload";

                #region Inicializar Excel
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Sheets worksheets;

                xlApp = new Excel.Application();
                xlApp.Visible = true;
                xlWorkBook = xlApp.Workbooks.Open(pathTemplate);

                #endregion


                //Localizar los nombres de los forms
                DataTable formsDT = oppInformation.LDRS.DefaultView.ToTable(true, "idName", "ldrName", "ldrBrand");


                #region Crear hojas de LDR.
                foreach (DataRow row in formsDT.Rows)
                {
                    console.WriteLine("Editando hoja del LDR: " + Regex.Replace(row["ldrName"].ToString(), @"[^\wáéíóúÁÉÍÓÚ¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5)));

                    int rowCount = 7;

                    // Obtener la hoja de trabajo (worksheet) de la plantilla
                    Excel.Worksheet xlTemplateWorksheet = xlWorkBook.Sheets["template"];

                    // Duplicar la hoja de trabajo
                    xlTemplateWorksheet.Copy(Type.Missing, xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);

                    // Establecerse en la hoja de copia
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[xlWorkBook.Sheets.Count];

                    // Asignar nombre a la hoja
                    string nameSheet = Regex.Replace(row["ldrName"].ToString(), @"[^\wáéíóúÁÉÍÓÚ¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));

                    if (nameSheet.Length > 30) //Para que no exceda 30
                    {
                        nameSheet = nameSheet.Substring(0, 30);
                    }
                    xlWorkSheet.Name = nameSheet;

                    // Asignar título ldr a la hoja
                    xlWorkSheet.Cells[4, 1].value = "LDR: " + Regex.Replace(row["ldrName"].ToString(), @"[^\wáéíóúÁÉÍÓÚ¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));

                    // Asignar título brand a la hoja
                    xlWorkSheet.Cells[5, 1].value = Regex.Replace(row["ldrBrand"].ToString(), @"[^\wáéíóúÁÉÍÓÚ¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));

                    // En la tabla del LDR filtrar los campos del LDR actual 
                    DataRow[] ldrFields = oppInformation.LDRS.Select($"idName = '{row["idName"]}'");

                    foreach (DataRow ldrField in ldrFields)
                    {
                        if (ldrField["type"].ToString() != "mainTitle" && ldrField["type"].ToString() != "mainSubtitle")
                        {
                            //Títulos y subtítulos
                            if (ldrField["type"].ToString() == "title" || ldrField["type"].ToString() == "subtitle")
                            {
                                xlWorkSheet.Cells[rowCount, 1].value = Regex.Replace(ldrField["label"].ToString(), @"[^\wáéíóúÁÉÍÓÚ/¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));

                                // Obtener las celdas que se van a agrupar
                                Excel.Range range = xlWorkSheet.Range[$"A{rowCount}:B{rowCount}"];

                                // Agrupar las celdas
                                range.Merge();

                                // Aplicar formato de letra en negrita a las celdas agrupadas
                                range.Font.Bold = true;

                                // Aplicar color de fondo a las celdas agrupadas
                                if (ldrField["type"].ToString() == "title") { range.Interior.Color = System.Drawing.Color.PaleTurquoise; }

                                // Centrar el contenido de las celdas agrupadas
                                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            }
                            else //Los demás campos 
                            {
                                xlWorkSheet.Cells[rowCount, 1].value = Regex.Replace(ldrField["label"].ToString(), @"[^\wáéíóúÁÉÍÓÚ/¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));
                                xlWorkSheet.Cells[rowCount, 2].value = Regex.Replace(ldrField["value"].ToString(), @"[^\wáéíóúÁÉÍÓÚ/¿?@,\-()]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));

                                // Acomodar text a la izquierda
                                xlWorkSheet.Cells[rowCount, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                                // Habilitar el ajuste de texto automático
                                xlWorkSheet.Cells[rowCount, 1].WrapText = true;
                                xlWorkSheet.Cells[rowCount, 2].WrapText = true;

                            }

                            rowCount++;
                        }
                    }

                }
                #endregion

                console.WriteLine("Guardando cambios...");

                #region Eliminar la hoja template

                // Establecerse en la hoja de template
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["template"];

                xlApp.DisplayAlerts = false;
                xlWorkSheet.Delete();
                xlApp.DisplayAlerts = true;

                #endregion

                #region Guardar excel
                string nameFileExcel = @"\LDR - " + oppInformation.opp + " - " + oppInformation.id + ".xlsx";
                string locationLDRToSave = pathSaveLDRS + nameFileExcel;

                fileLDRRoute.Add(nameFileExcel, pathSaveLDRS);
                //fileLDRRoute.Add(nameFileExcel);

                xlApp.DisplayAlerts = false;
                xlWorkBook.SaveAs(locationLDRToSave);
                xlWorkBook.Close();
                xlApp.Quit();
                process.KillProcess("EXCEL", true);
                #endregion

                console.WriteLine("LDR de la oportunidad " + oppInformation.opp + " - " + oppInformation.id + " creado con éxito.");



            }
            return fileLDRRoute;

        }
        /// <summary>
        /// Método para descargar los BOM subidos al FTP por el usuario cuando creo la oportunidad
        /// </summary>
        /// <param name="Management"></param>
        /// /// <returns>Returna un diccionario donde el key es el path donde esta el BOM y value el nombre del BOM</returns>
        public void DownloadsBOMFiles(Dictionary<string, string> filesBOMRoutes, string oppId)
        {

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //Rutas 
            string pathDownloads = desktop + @"\databot\downloads";

            //user!=databot, porque significa que si no lo subio el databot, lo subio s&s, por consiguiente es un BOM.
            string sqlBom = $"SELECT * FROM UploadsFiles WHERE oppId={oppId} and user!='Databot'";
            DataTable bomTable = crud.Select(sqlBom, "autopp2_db", enviroment);

            foreach (DataRow bomFile in bomTable.Rows)
            {
                //Descargar el archivo del FTP
                bool result = autoppSQL.DownloadFile(bomFile["path"].ToString(), enviroment);
                console.WriteLine("Descargando el archivo BOM: " + bomFile["name"].ToString());

                if (!result)//Hubo un error con la descarga
                {
                    sett.SendError(this.GetType(), "No existe el archivo BOM en el server de S&S", "No se encontro el archivo " + bomFile["name"].ToString() + " de la oppId: " + oppId + " por tanto no se subió a SAP");

                }
                else if (result) //No hubo problema con la descarga.
                {
                    filesBOMRoutes.Add(@"\" + bomFile["name"].ToString(), pathDownloads);
                }
            }

        }

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
            AND opp.opp = '' 
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

                case "LDRS":
                    sql =
                    $@"

                        SELECT  LDRRequestsData.id, LDRName.idName as idName, LDRName.name as ldrName, LDRBrand.name as ldrBrand, LDRTypeField.name as type, LDRFields.label, LDRRequestsData.value 
                        
                        FROM LDRRequestsData 

                        INNER JOIN LDRFields ON LDRRequestsData.LDRFieldId = LDRFields.id 
                        INNER JOIN LDRTypeField ON LDRFields.typeField = LDRTypeField.id 
                        INNER JOIN LDRRequests ON LDRRequestsData.LDRRequestId = LDRRequests.id 
                        INNER JOIN LDRName ON LDRRequests.LDRNameId = LDRName.id 
                        INNER JOIN LDRBrand ON LDRName.LDRBrand = LDRBrand.id


                        WHERE LDRRequests.oppId={oppId} 
                        AND LDRRequests.active=1
                    ";
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

                            logical.AutoppNotifications("successNotification", employeeCreatorData["user"].ToString(), toReplace, notificationsConfig);

                            if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                            {
                                logical.AutoppNotifications("successNotification", employeeResponsibleData["user"].ToString(), toReplace, notificationsConfig);
                            }

                            if (LDROrBOMDocument == "LDR")
                            {
                            }
                            else { }

                            //Respaldo para el admin.
                            logical.AutoppNotifications("successNotification", employeeResponsibleData["user"].ToString(), toReplace, "admin");


                        }
                        #endregion
                        break;

                    case 4:
                        #region Notificación de aprobación final de requerimiento.
                        if (oppGestion.opp != "")
                        {
                            titleWebex = $"Creación de requerimiento exitoso - {oppGestion.opp}";

                            string greeting = process.greeting() + $" estimado(a) {employeeName}.";


                            #region Moldeo de mensaje
                            string msgWebex = root.Email_Body.Replace("Caso Finalizado", $" El caso #{caseNumber} ha finalizado.");
                            //Poner:
                            msgWebex = msgWebex.Replace("oportunidad", "oportunidad:").Replace("Caso", "Caso:").Replace("Solicitante", "Solicitante:").Replace("GTL", "GTL:").Replace("requerimiento", "requerimiento:").Replace("Cliente", "Cliente:").Replace("País", "País:").Replace("Proveedor", "Proveedor:").Replace("Producto", "Producto:");

                            //Reemplazar el mensaje de despedida.
                            msgWebex = msgWebex.Replace("Gracias por utilizar los servicios del GTL:, su caso ha sido finalizado, apreciamos que en las próximas 8 horas hábiles realice la revisión en BPM para dar por completado el caso",
                                $".\r\n \r\n Puede consultar el estado de la gestión en el botón de solicitudes de la Fábrica de Propuestas en <a href='{routeSS}'>Smart And Simple</a> de GBM");

                            #endregion

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - {oppGestion.opp}", oppGestion.opp, greeting + msgWebex);

                            //Respaldo al admin
                            webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, greeting + msgWebex);

                        }
                        #endregion
                        break;

                    case 5:
                        #region El especialista ha devuelto un requerimiento.
                        if (oppGestion.opp != "")
                        {
                            titleWebex = $"El especialista ha devuelto un requerimiento BAW - {oppGestion.opp}";
                            string msgWebex = "";

                            string greetingWebex = process.greeting() + $" estimado(a) EMPLOYEE.\r\n \r\nLe informo que el requerimiento Número de caso: {caseNumber} , de la gestión  #{oppGestion.id} - #{oppGestion.opp} " +
                            $"del cliente: " + client["name"].ToString();

                            string bodyWebex = ", fue devuelto por el especialista de BAW."
                                + ".\r\n \r\n En base a su criterio, puede editar el requerimiento, y volverlo a enviar al especialista. Para ello debe seguir los siguientes pasos:"
                                + $"\r\n 1. Ingresar a <a href='{routeSS}'>Smart And Simple</a> de GBM, seguidamente en Ventas - Autopp - Inicio - Mis Solicitudes."
                            + $"\r\n 2. Encontrar la opp {oppGestion.opp}, tocar el botón BAW, hallar el número de caso {caseNumber} y tocar el botón Editar"
                                + "\r\n3. Por último, editar los datos del requerimiento BAW en base a su criterio y volverlo enviar al especialista."
                                + "\r\n \r\nSaludos cordiales.";

                            msgWebex = greetingWebex.Replace("EMPLOYEE", employeeName) + bodyWebex;

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - {oppGestion.opp}", oppGestion.opp, msgWebex);

                            if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                            {
                                try
                                {
                                    employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());
                                }
                                catch (Exception e) { employeeName = ""; }

                                msgWebex = greetingWebex.Replace("EMPLOYEE", employeeName) + bodyWebex;

                                webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - {oppGestion.opp}", oppGestion.opp, msgWebex);

                            }

                            //Respaldo admin
                            webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex);

                        }
                        #endregion
                        break;

                    case 6:
                        #region El especialista ha rechazado un requerimiento.
                        if (oppGestion.opp != "")
                        {
                            titleWebex = $"El especialista ha rechazado un requerimiento BAW - {oppGestion.opp}";
                            string msgWebex = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el requerimiento Número de caso: {caseNumber} , de la gestión  #{oppGestion.id} - #{oppGestion.opp} " +
                            $"del cliente: " + client["name"].ToString();
                            msgWebex += ", fue rechazado por el especialista de BAW."
                                + "\r\n \r\nSaludos cordiales.";

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, $"El especialista ha rechazado un requerimiento BAW - {oppGestion.opp}", oppGestion.opp, msgWebex);

                            if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                            {
                                try
                                {
                                    employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());
                                }
                                catch (Exception e) { employeeName = ""; }

                                msgWebex = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el requerimiento Número de caso: {caseNumber} , de la gestión  #{oppGestion.id} - #{oppGestion.opp} " +
                                          $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                                msgWebex += ", fue rechazado por el especialista de BAW."
                                 + "\r\n \r\nSaludos cordiales.";


                                //webex.SendCCNotification(/*employeeResponsibleData.Usuario*/userAdmin + "@GBM.NET", titleWebex, $"El especialista ha rechazado un requerimiento BAW - { oppGestion.opp}", oppGestion.opp, msgWebex);
                                //webex.SendCCNotification(employeeResponsibleData["user"].ToString() + "@GBM.NET", titleWebex, $"El especialista ha rechazado un requerimiento BAW - {oppGestion.opp}", oppGestion.opp, msgWebex);
                                webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, $"El especialista ha rechazado un requerimiento BAW - {oppGestion.opp}", oppGestion.opp, msgWebex);



                            }

                            //Respaldo admin
                            webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "El especialista ha rechazado un requerimiento BAW", oppGestion.opp, msgWebex);

                        }
                        #endregion
                        break;

                    case 7:
                        #region Se ha devuelto un requerimiento al especialista
                        if (oppGestion.opp != "")
                        {
                            titleWebex = $"Requerimiento devuelto al especialista con éxito - {oppGestion.opp}";
                            string msgWebex = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el requerimiento Número de caso: {caseNumber} , de la gestión  #{oppGestion.id} - #{oppGestion.opp} " +
                            $"del cliente: " + client["name"].ToString();
                            msgWebex += ", fue devuelto al especialista de BAW con los nuevos requerimientos con éxito. "
                                + $"\r\nPuede consultar el estado del mismo en <a href='{routeSS}'>Smart And Simple</a> de GBM, seguidamente en Ventas - Autopp - Inicio - Mis Solicitudes."
                                + "\r\n \r\nSaludos cordiales.";

                            //webex.SendCCNotification(/*employeeCreatorData.Usuario*/userAdmin + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - { oppGestion.opp}", oppGestion.opp, msgWebex);
                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - {oppGestion.opp}", oppGestion.opp, msgWebex);

                            if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                            {
                                try
                                {
                                    employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());
                                }
                                catch (Exception e) { employeeName = ""; }

                                msgWebex = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el requerimiento Número de caso: {caseNumber} , de la gestión  #{oppGestion.id} - #{oppGestion.opp} " +
                                $"del cliente: " + client["name"].ToString();
                                msgWebex += " el cual usted fue seleccionado como empleado responsable, fue devuelto al especialista de BAW con los nuevos requerimientos con éxito. "
                                    + $"\r\n Puede consultar el estado del mismo en <a href='{routeSS}'>Smart And Simple</a> de GBM, seguidamente en Ventas - Autopp - Inicio - Mis Solicitudes."
                                    + "\r\n \r\nSaludos cordiales.";


                                //webex.SendCCNotification(/*employeeResponsibleData.Usuario*/userAdmin + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - { oppGestion.opp}", oppGestion.opp, msgWebex);
                                webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, $"Creación de requerimiento exitoso - {oppGestion.opp}", oppGestion.opp, msgWebex);



                            }

                            //Respaldo admin
                            webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex);

                        }
                        #endregion
                        break;
                        
                    case 8:
                        #region Notificación de éxito de creación de la opp, LDR pero sin BAW.

                        if (oppGestion.opp != "")
                        {

                            Dictionary<string, string> toReplace = new Dictionary<string, string>(){
                                    {"TITLENOTIFICATION", "Creación éxitosa de oportunidad"},
                                    {"USER", employeeCreatorData["user"].ToString()},
                                    {"EMPLOYEERESPONSIBLE", employeeResponsibleData["user"].ToString()},
                                    {"TYPEOPPORTUNITY", oppGestion.generalData.typeOpportunityName},
                                    {"OPP", oppGestion.opp},
                                    {"CLIENT", client["name"].ToString()}
                                };

                            logical.AutoppNotifications("successNotification", employeeCreatorData["user"].ToString(), toReplace, notificationsConfig);

                            if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                            {
                                logical.AutoppNotifications("successNotification", employeeResponsibleData["user"].ToString(), toReplace, notificationsConfig);
                            }

                            if (LDROrBOMDocument == "LDR")
                            {
                            }
                            else { }

                            //Respaldo para el admin.
                            logical.AutoppNotifications("successNotification", employeeResponsibleData["user"].ToString(), toReplace, "admin");


                        }
                        #endregion
                        break;

                    case 9:
                        #region Notificación al equipo de SalesTeam que ha sido agregado a la oportunidad.

                        if (oppGestion.opp != "" || true)
                        {
                            foreach (DataRow employee in salesTInfo.Rows)
                            {
                                if (employee["id"].ToString() != "41" ||
                                    (employeeResponsibleData["user"].ToString() != employeeCreatorData["user"].ToString()) //No notifique al empleado responsable de la oportunidad.
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

                                    logical.AutoppNotifications("salesTeamsNotification", employee["user"].ToString(), toReplace, notificationsConfig);

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
                    case 1:
                        #region Error al crear la opp
                        //Notificar al usuario
                        titleWebex = $"Error al crear la oportunidad - {oppGestion.id}";

                        string msgWebex1 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                        $" del cliente: " + client["name"].ToString();
                        msgWebex1 += ", no fue creada debido a un error inesperado.\r\nPor favor contáctese con Application Management y Support.";

                        webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);

                        //En caso que sea diferente, notifique al usuario responsable también
                        if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                        {

                            try //Seleccionar el nombre y primer letra en mayúscula.
                            {
                                employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());

                            }
                            catch (Exception e) { employeeName = ""; }

                            msgWebex1 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                                $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                            msgWebex1 += ", no fue creada debido a un error inesperado.\r\nPor favor contáctese con Application Management y Support.";


                            //webex.SendCCNotification(/*employeeResponsibleData["user"].ToString()*/userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);
                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);
                        }


                        //Respaldo admin
                        webex.SendCCNotification(/*employeeData.Usuario*/userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);


                        //Notificar a Application Management
                        string msg = "Este error está en el try catch de la fase 1 Crear Oportunidades en SAP - Autopp";
                        sett.SendError(this.GetType(), $"Error al crear la opp id #{idOpp}", msg, exception);

                        #endregion
                        break;

                    case 2:
                        #region Error al crear el LDR

                        if (oppGestion.opp != "")
                        {
                            #region Notificación por Webex Teams al usuario
                            titleWebex = $"Error al crear LDR - {oppGestion.opp} - {oppGestion.id}";

                            string msgWebex2 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                            $"del cliente: " + client["name"].ToString();
                            msgWebex2 += ", ha tenido un problema al crear y subir el archivo LDR a SAP y al servidor.\r\nPor favor contáctese con Application Management y Support.";

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear LDR", oppGestion.opp, msgWebex2);


                            //En caso que sea diferente, notifique al usuario responsable también
                            if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                            {

                                //Seleccionar el nombre y primer letra en mayúscula.
                                try { employeeName = BuildFirstName(employeeResponsibleData["name"].ToString()); } catch (Exception e) { employeeName = ""; }


                                msgWebex2 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                                    $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                                msgWebex2 += ", ha tenido un problema al crear y subir el archivo LDR a SAP y al servidor.\r\nPor favor contáctese con Application Management y Support.";


                                //webex.SendCCNotification(/*employeeResponsibleData.Usuario*/userAdmin + "@GBM.NET", titleWebex, "Error al crear LDR", oppGestion.opp, msgWebex2);
                                webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear LDR", oppGestion.opp, msgWebex2);
                            }


                            //Respaldo admin
                            webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex2);

                            #endregion

                            #region Notificación por correo electrónico resposables en Application Management
                            string errorList = $"No se pudo subir el archivo LDR de la gestión #{oppGestion.id}, de la oportunidad #{oppGestion.opp}, del usuario: {employeeCreatorData["user"].ToString()}.<br><br>" +
                            "A continuación se detallan los errores generados: <br><br>";
                            for (int i = 0; i < listErrorsFase2.Count; i++)
                            {
                                errorList += "#" + (i + 1) + " " + listErrorsFase2[i] + "<br>";
                            }

                            errorList += $"<br>Sugerencia de solución: <br>Contáctese con el usuario {employeeCreatorData["user"].ToString()} para verificar-informar lo que sucedió, subir el archivo LDR a SAP y el FTP, y finalmente poner" +
                                $" correr el robot localmente y crear requerimientos en BAW";

                            //El sendError imprime en consola los errores a la vez
                            sett.SendError(this.GetType(), $"Error al crear LDR id #{oppGestion.id} - opp #{oppGestion.opp} - {employeeCreatorData["user"].ToString()}", errorList, exception);
                            #endregion


                        }

                        #endregion
                        break;

                    case 3:
                        #region Error al crear el requerimiento.

                        if (oppGestion.opp != "")
                        {
                            #region Notificación por Webex Teams al usuario
                            titleWebex = $"Error al crear requerimiento en BAW - {oppGestion.opp} - {oppGestion.id}";

                            string msgWebex3 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                            $"del cliente: " + client["name"].ToString();
                            msgWebex3 += $", ha tenido un problema al crear los requerimientos en la página de BAW, puede revisar el estado de cada uno en la página de <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear requerimientos en BAW", oppGestion.opp, msgWebex3);

                            //En caso que sea diferente, notifique al usuario responsable también
                            if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                            {

                                //Seleccionar el nombre y primer letra en mayúscula.
                                try { employeeName = BuildFirstName(employeeResponsibleData["name"].ToString()); } catch (Exception e) { employeeName = ""; }


                                msgWebex3 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                                    $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                                msgWebex3 += $", ha tenido un problema al crear los requerimientos en la página de BAW, puede revisar el estado de cada uno en la página de <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                                webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear requerimientos en BAW", oppGestion.opp, msgWebex3);
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

                    case 4:
                        #region Error al realizar la aprobación final en BAW.

                        #region Notificación por Webex Teams al usuario
                        titleWebex = $"Error al realizar la aprobación final en BAW - {oppGestion.opp} - {oppGestion.id}";

                        string msgWebex4 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                        $"del cliente: " + client["name"].ToString();
                        msgWebex4 += ", ha tenido un problema al realizar la aprobación final en BAW, de los requerimientos que ya han sido aprobados por los especialistas.\r\nPor favor contáctese con Application Management y Support.";

                        //webex.SendCCNotification(employeeCreatorData["user"].ToString()/*userAdmin*/ + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex4);

                        //ELIMINAR CUANDO SE PASE A PRD
                        webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex4);


                        //En caso que sea diferente, notifique al usuario responsable también
                        if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                        {

                            try //Seleccionar el nombre y primer letra en mayúscula.
                            { employeeName = BuildFirstName(employeeResponsibleData["name"].ToString()); }
                            catch (Exception e) { employeeName = ""; }


                            msgWebex4 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                            $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                            msgWebex4 += ", ha tenido un problema al realizar la aprobación final en BAW, de los requerimientos que ya han sido aprobados por los especialistas.\r\nPor favor contáctese con Application Management y Support.";

                            //webex.SendCCNotification(employeeResponsibleData["user"].ToString()/*userAdmin*/ + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex4);
                            //webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex4);

                        }



                        #endregion

                        #region Notificar a Application Management


                        string msgApp4 = "Este error está en el try catch de la Fase 3: Crear requisitos BAW  - Autopp. Revisar el screenshot del error en la carpeta en el desktop/Databot/Autopp/ErrorsScreenshots.";
                        sett.SendError(this.GetType(), $"Error al crear requisitos en BAW del id #{idOpp}", msgApp4, exception);


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

                        #endregion
                        break;

                    case 5:
                        #region Error al establecer el caso que es devuelto por el especialista.

                        #region Notificación por Webex Teams al usuario
                        titleWebex = $"Error al establecer el número de caso BAW en devuelto por especialista - {oppGestion.opp} - {caseNumber}";

                        string msgWebex5 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el número de caso {caseNumber} de la opp  #{oppGestion.opp} " +
                        $"del cliente: " + client["name"].ToString();
                        msgWebex5 += $", ha sido devuelto por el especialista, pero el robot no ha podido establecerlo en <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                        webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex5);

                        //ELIMINAR CUANDO SE PASE A PRD
                        webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex5);


                        //En caso que sea diferente, notifique al usuario responsable también
                        if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                        {

                            try //Seleccionar el nombre y primer letra en mayúscula.
                            {
                                employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());
                            }
                            catch (Exception e) { employeeName = ""; }


                            msgWebex5 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el número de caso {caseNumber} de la opp  #{oppGestion.opp}" +
                                          $"del cliente: " + client["name"].ToString() + ", el cual se ha establecido a usted como empleado resposable";
                            msgWebex5 += $", ha sido devuelto por el especialista, pero el robot no ha podido establecerlo en <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex5);
                            //webex.SendCCNotification(/*employeeResponsibleData.Usuario*/userAdmin + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex5);

                        }



                        #endregion

                        #region Notificar a Application Management


                        string msgApp5 = $"No se pudo establecer como devolución por parte del especialista el Case Number de BAW {caseNumber}, es muy importante darle seguimiento y notificar al usuario";
                        sett.SendError(this.GetType(), $"Error al establecer como devolución por parte del especialista el Requerimiento BAW de la opp #{oppGestion.opp}", msgApp5, exception);


                        #endregion

                        #endregion
                        break;

                    case 6:
                        #region Error al establecer como rechazado el requerimiento.

                        #region Notificación por Webex Teams al usuario
                        titleWebex = $"Error al establecer como rechazado por especialista BAW - {oppGestion.opp} - {oppGestion.id}";
                        string msgWebex6 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                        $"del cliente: " + client["name"].ToString();
                        msgWebex6 += $", ha sido rechazado por el especialista, pero ha ocurrido un error en el robot y no se ha podido reflejar en <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                        webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al establecer como rechazado por especialista BAW ", oppGestion.opp, msgWebex6);

                        //ELIMINAR CUANDO SE PASE A PRD
                        webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex6);


                        //En caso que sea diferente, notifique al usuario responsable también
                        if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                        {

                            try //Seleccionar el nombre y primer letra en mayúscula.
                            {
                                employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());
                            }
                            catch (Exception e) { employeeName = ""; }

                            msgWebex6 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                            $"del cliente: " + client["name"].ToString() + "el cual usted es empleado responsable";
                            msgWebex6 += $", ha sido rechazado por el especialista, pero ha ocurrido un error en el robot y no se ha podido reflejar en <a href='{routeSS}'>Smart And Simple</a> de GBM.\r\nPor favor contáctese con Application Management y Support.";

                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al establecer como rechazado por especialista BAW ", oppGestion.opp, msgWebex6);
                            //webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al establecer como rechazado por especialista BAW ", oppGestion.opp, msgWebex6);

                        }


                        #endregion

                        #region Notificar a Application Management


                        string msgApp6 = $"No se pudo establecer como rechazado por parte del especialista el Case Number de BAW {caseNumber}, es muy importante darle seguimiento y notificar al usuario";
                        sett.SendError(this.GetType(), $"Error al establecer como rechazado por parte del especialista el Requerimiento BAW de la opp#{oppGestion.opp}", msgApp6, exception);

                        #endregion


                        #endregion
                        break;

                    case 7:
                        #region Error al establecer el caso que es devuelto al especialista.

                        #region Notificación por Webex Teams al usuario
                        titleWebex = $"Error al reenviar el número de caso BAW al especialista - {oppGestion.opp} - {caseNumber}";

                        string msgWebex7 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el número de caso {caseNumber} de la opp  #{oppGestion.opp} " +
                        $"del cliente: " + client["name"].ToString();
                        msgWebex7 += ", se ha intentado devolver al especialista con la nueva información, pero ocurrió un error inesperado.\r\nPor favor contáctese con Application Management y Support.";

                        webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al reenviar el número de caso BAW al especialista", oppGestion.opp, msgWebex7);



                        //En caso que sea diferente, notifique al usuario responsable también
                        if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                        {

                            try //Seleccionar el nombre y primer letra en mayúscula.
                            { employeeName = BuildFirstName(employeeResponsibleData["name"].ToString()); }
                            catch (Exception e) { employeeName = ""; }


                            msgWebex7 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que el número de caso {caseNumber} de la opp  #{oppGestion.opp}" +
                                $"del cliente: " + client["name"].ToString() + ", el cual se ha establecido a usted como empleado responsable"; ;
                            msgWebex7 += ", se ha intentado devolver nuevamente al especialista con la nueva información, pero ocurrió un error inesperado.\r\nPor favor contáctese con Application Management y Support.";


                            webex.SendCCNotification((notificationsConfig == "admin" ? userAdmin : employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al reenviar el número de caso BAW al especialista", oppGestion.opp, msgWebex7);
                            //webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al reenviar el número de caso BAW al especialista", oppGestion.opp, msgWebex7);

                        }

                        //Respaldo admin 
                        webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Error al realizar aprobación final en BAW", oppGestion.opp, msgWebex7);



                        #endregion

                        #region Notificar a Application Management

                        string msgApp7 = $"No se pudo devolver la nueva información al especialista del requerimiento Case Number de BAW {caseNumber}, es muy importante darle seguimiento y notificar al usuario";
                        sett.SendError(this.GetType(), $"Error al devolver nuevamente al especialista con la nueva información del Requerimiento BAW de la opp #{oppGestion.opp}", msgApp7, exception);


                        #endregion

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

            #endregion
        }



        #endregion

























        /// <summary>
        /// Método backup que ya no se utiliza, y es similar a ProcessAutopp, pero gestiona cada solicitud de manera paralela, entonces si hay 4 solicitudes 
        /// las 4 terminan al mismo tiempo, algo parecido a un hilo. No se borró por efectos de si un futuro es útil.
        /// </summary>
        public void ProcessAutoppBackUp()
        {

            #region Status 1- "En proceso crear opp CRM" 
            DataTable newOppRequests1 = GetReqsForStatus("1");
            if (newOppRequests1.Rows.Count > 0)
            {

                console.WriteLine("");
                console.WriteLine("************************************");
                console.WriteLine("*Fase 1: Crear Oportunidades en SAP*");
                console.WriteLine("************************************");
                console.WriteLine("");


                int indexReqs1 = 1;
                foreach (DataRow oppReq in newOppRequests1.Rows)
                {

                    int idOpp = (int)oppReq.ItemArray[0];
                    try
                    {
                        #region General Data
                        GeneralData generalData = new GeneralData();

                        generalData.typeOpportunity = oppReq["typeOpportunity"].ToString();
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
                        DataTable salesTInfo = oppInfo(idOpp, "salesTeam");
                        List<SalesTeams> salesTList = new List<SalesTeams>();
                        foreach (DataRow salesTItem in salesTInfo.Rows)
                        {
                            SalesTeams item = new SalesTeams();
                            item.role = salesTItem["code"].ToString();
                            item.employee = "AA" + salesTItem["UserID"].ToString().PadLeft(8, '0');

                            salesTList.Add(item);
                        }
                        #endregion

                        #region Objeto principal donde se une toda la información
                        AutoppInformation oppGestion = new AutoppInformation();
                        oppGestion.id = oppReq["id"].ToString();
                        oppGestion.status = oppReq["status"].ToString();
                        oppGestion.employee = oppReq["createdBy"].ToString();
                        oppGestion.opp = oppReq["opp"].ToString();
                        oppGestion.generalData = generalData;
                        oppGestion.organizationAndClientData = organizationAndClientData;
                        oppGestion.salesTeams = salesTList;
                        CCEmployee employeeData = new CCEmployee(oppGestion.employee);
                        oppGestion.employee = employeeData.IdEmpleado;

                        //Extraer el nombre del cliente.
                        string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {organizationAndClientData.client}";
                        DataTable dtClient = crud.Select(sqlClient, "databot_db", enviroment);

                        console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes.");
                        console.WriteLine("");
                        console.WriteLine($"Solicitud id {oppGestion.id} - {dtClient.Rows[0]["name"]}");

                        #endregion

                        #region Crear la oportunidad y notificación de éxito o fallo.
                        oppGestion.opp = CreateOppCRM(oppGestion, employeeData.Usuario);


                        string setStatus = "2";

                        //Es diferente a GTL Quotation, por tanto no necesita realizar las demás etapas (crear LDR, BAW, etc).
                        if (generalData.cycle != "Y3A")
                        {
                            setStatus = "5";
                        }

                        if (oppGestion.opp == "" || oppGestion.opp == null)
                        {//Fallo 
                            NotifySuccessOrErrors("Fail", 1);
                            setStatus = "6";
                        }
                        else
                        {//Éxito           
                            NotifySuccessOrErrors("Fail", 3);
                        }

                        string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status={setStatus} WHERE id= {oppGestion.id}";
                        crud.Update(updateQuery, "autopp2_db", enviroment);
                        //log.LogDeCambios("Creación", "Autopp", oppGestion.employee, "", oppGestion.opp, "Creación de oportunidad");

                        #endregion

                    }
                    catch (Exception e)
                    {
                        string msg = "Este error está en el try catch de la fase 1 Crear Oportunidades en SAP - Autopp";
                        sett.SendError(this.GetType(), $"Error al crear la opp id #{idOpp}", msg, e);

                        string updateQuery = $"UPDATE OppRequests SET status = 6, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                        crud.Update(updateQuery, "autopp2_db", enviroment);
                    }

                    console.WriteLine("");
                    indexReqs1++;


                }
            }
            #endregion

            #region Status 2- "En proceso crear LDRS" 
            DataTable newOppRequests2 = GetReqsForStatus("2");
            if (newOppRequests2.Rows.Count > 0 && !sap.CheckLogin(sapSystem, mandante))
            {

                console.WriteLine("");
                console.WriteLine("********************");
                console.WriteLine("*Fase 2: Crear LDRS*");
                console.WriteLine("********************");
                console.WriteLine("");

                //Bloquear RPA User
                sap.BlockUser(sapSystem, 1, mandante);


                int indexReqs2 = 1;
                foreach (DataRow oppReq2 in newOppRequests2.Rows)
                {
                    int idOpp = (int)oppReq2["id"];

                    try
                    {

                        #region LDRS deserializar
                        List<LDRSAutopp> listLDRFather = new List<LDRSAutopp>();

                        DataTable LDRSInfo = oppInfo(idOpp, "LDRS");

                        JObject ldrsJSON;
                        try { ldrsJSON = JObject.Parse(LDRSInfo.Rows[0][2].ToString()); } catch (Exception e) { ldrsJSON = new JObject(); };
                        Dictionary<string, object> DictLdrFather = ldrsJSON.ToObject<Dictionary<string, object>>();

                        //Recorre y crea el LDRSAutopp padre (technology, list<ItemLDRAutopp>)
                        foreach (KeyValuePair<string, object> dictLdrFather in DictLdrFather)
                        {
                            LDRSAutopp ldrFather = new LDRSAutopp();
                            List<ItemLDRAutopp> listldrsChilds = new List<ItemLDRAutopp>();

                            JObject ldrDeserialize = JObject.Parse(dictLdrFather.Value.ToString());
                            Dictionary<string, object> dictLdrChild = ldrDeserialize.ToObject<Dictionary<string, object>>();

                            //Recorre y crea el ItemLDRAutopp hijo (id, value)
                            foreach (KeyValuePair<string, object> ldrItemChild in dictLdrChild)
                            {
                                ItemLDRAutopp itemLDRChild = new ItemLDRAutopp();
                                itemLDRChild.id = ldrItemChild.Key;
                                itemLDRChild.value = ldrItemChild.Value.ToString();

                                listldrsChilds.Add(itemLDRChild);
                            }

                            ldrFather.technology = dictLdrFather.Key;
                            ldrFather.LDR = listldrsChilds;

                            //Agrega a la lista principal de LDRSAutopp
                            listLDRFather.Add(ldrFather);

                        }
                        #endregion

                        #region OrganizationAndClientData
                        DataTable orgInfo = oppInfo(idOpp, "organizationAndClientData");
                        OrganizationAndClientData organizationAndClientData = new OrganizationAndClientData();
                        organizationAndClientData.client = "00" + orgInfo.Rows[0]["idClient"].ToString();
                        #endregion

                        #region Objeto principal donde se une toda la información
                        AutoppInformation oppGestion = new AutoppInformation();

                        oppGestion.id = oppReq2["id"].ToString();
                        oppGestion.status = oppReq2["status"].ToString();
                        oppGestion.employee = oppReq2["createdBy"].ToString();
                        oppGestion.opp = oppReq2["opp"].ToString();


                        //oppGestion.LDRS = listLDRFather;
                        CCEmployee employeeData = new CCEmployee(oppGestion.employee);
                        oppGestion.employee = employeeData.IdEmpleado;

                        #endregion


                        //Extraer el nombre del cliente.
                        string sqlClient2 = $"SELECT name FROM `clients` WHERE `idClient` = {organizationAndClientData.client}";
                        DataTable dtClient2 = crud.Select(sqlClient2, "databot_db", enviroment);

                        console.WriteLine($"Procesando solicitud {indexReqs2} de {newOppRequests2.Rows.Count} solicitudes.");
                        console.WriteLine("");
                        console.WriteLine($"Solicitud id {oppGestion.id} - {dtClient2.Rows[0]["name"]}");

                        //Aquí se crea el LDR
                        List<string> fileLDRRoute = null; //CreateLDR(oppGestion);

                        int status = 0;

                        if (fileLDRRoute.Count > 0) //Verifica si se debe subir algún LDRS.
                        {

                            #region Subir archivo al FTP
                            //En caso que al subir el LDR diera error lo agrega a la siguiente lista.
                            List<string> uploadsFailed = new List<string>();


                            //En la [0] es ruta de la carpeta y en la [1] el nombre del Excel
                            bool resultFTP = autoppSQL.InsertFileAutopp(oppGestion.id, fileLDRRoute[0] + fileLDRRoute[1], enviroment);
                            if (resultFTP)
                            {
                                console.WriteLine($"LDR subido al FTP con éxito.");
                            }
                            else
                            {
                                console.WriteLine($"Ocurrió un error no se pudo subir al FTP.");
                                uploadsFailed.Add("Error al subir al FTP");
                            }


                            #endregion

                            #region Subir archivo a SAP


                            if (fileLDRRoute[0] != "" && fileLDRRoute[0] != null)
                            {
                                //console.WriteLine($"Subiendo el LDR a SAP - {mandante}...");
                                process.KillProcess("saplogon", false);
                                sap.LogSAP(sapSystem, mandante);

                                EasyLDR con = new EasyLDR();
                                bool resultLDR = true;//con.ConnectSAP(fileLDRRoute, oppGestion.opp);

                                if (resultLDR)
                                {
                                    console.WriteLine($"LDR subido a SAP con éxito.");
                                }
                                else
                                {
                                    console.WriteLine($"Ocurrió un error no se pudo subir el archivo a SAP.");
                                    uploadsFailed.Add("Error al subir a SAP");
                                }

                                sap.KillSAP();

                            }

                            #endregion

                            #region Notificar errores de carga.                    

                            if (uploadsFailed.Count > 0)
                            {
                                //Extraer el nombre del cliente.
                                string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {organizationAndClientData.client}";
                                DataTable dtClient = crud.Select(sqlClient, "databot_db", enviroment);

                                NotifySuccessOrErrors("Fail", 2, uploadsFailed);
                                status = 7; //El status se pone en error de LDRS.

                            }
                            else // Si no hay errores elimina el LDR de local.
                            {
                                File.Delete(fileLDRRoute[0] + fileLDRRoute[1]);
                            }

                            #endregion

                        }

                        #region Actualizar status y log de cambios 

                        string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status={status} WHERE id= {oppGestion.id}";
                        crud.Update(updateQuery, "autopp2_db", enviroment);
                        //log.LogDeCambios("Creación", "Autopp", oppGestion.employee, "", oppGestion.opp, "Creación de LDRS");

                    }
                    catch (Exception e)
                    {
                        string msg = "Este error está en el try catch de la Fase 2: Crear LDRS - Autopp.";
                        sett.SendError(this.GetType(), $"Error al crear LDR id #{idOpp}", msg, e);

                        string updateQuery = $"UPDATE OppRequests SET status = 7, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                        crud.Update(updateQuery, "autopp2_db", enviroment);

                        process.KillProcess("EXCEL", true);
                        sap.BlockUser(sapSystem, 0, mandante);
                    }

                    #endregion
                    indexReqs2++;

                }


                //Desbloquear RPA User
                sap.BlockUser(sapSystem, 0, mandante);
            }

            #endregion

            #region Status 3- "En proceso crear requisitos BAW" 
            DataTable newOppRequests3 = GetReqsForStatus("3");
            bool loginOpened3 = false;

            if (newOppRequests3.Rows.Count > 0)
            {
                console.WriteLine("");
                console.WriteLine("******************************");
                console.WriteLine("*Fase 3: Crear requisitos BAW*");
                console.WriteLine("******************************");
                console.WriteLine("");
                //console.WriteLine("\nExisten " + newOppRequests3.Rows.Count + " solicitud(es) de creación de requisitos en BAW...");

                //Establecer conexión en BAW
                IWebDriver chrome = SelConn(GetUrlBaw());

                int indexReqs3 = 1;
                foreach (DataRow oppReq3 in newOppRequests3.Rows)
                {
                    int idOpp = (int)oppReq3.ItemArray[0];

                    try
                    {

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

                        #region OrganizationAndClientData
                        DataTable orgInfo = oppInfo(idOpp, "organizationAndClientData");
                        OrganizationAndClientData organizationAndClientData = new OrganizationAndClientData();
                        organizationAndClientData.client = "00" + orgInfo.Rows[0]["idClient"].ToString();
                        #endregion

                        #region Objeto principal donde se une toda la información
                        AutoppInformation oppGestion = new AutoppInformation();
                        oppGestion.id = oppReq3["id"].ToString();
                        oppGestion.status = oppReq3["status"].ToString();
                        oppGestion.employee = oppReq3["createdBy"].ToString();
                        oppGestion.opp = oppReq3["opp"].ToString();
                        oppGestion.organizationAndClientData = organizationAndClientData;

                        oppGestion.BAW = BAWList;

                        CCEmployee employeeData = new CCEmployee(oppGestion.employee);
                        oppGestion.employee = employeeData.IdEmpleado;

                        //Extraer el nombre del cliente.
                        string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {organizationAndClientData.client}";
                        DataTable dtClient = crud.Select(sqlClient, "databot_db", enviroment);

                        console.WriteLine("");
                        console.WriteLine($"Procesando solicitud {indexReqs3} de {newOppRequests3.Rows.Count} solicitudes.");
                        console.WriteLine("");
                        console.WriteLine($"Solicitud id {oppGestion.id} - {dtClient.Rows[0]["name"]}");


                        #endregion

                        #region Agregar requerimientos en BAW 


                        #region Login en BAW.
                        //console.WriteLine("Iniciando sesión en baw...");
                        if (!loginOpened3)
                        {

                            chrome.FindElement(By.Id("username")).SendKeys(cred.user_baw);
                            chrome.FindElement(By.Id("password")).SendKeys(cred.password_baw);

                            chrome.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/a[1]")).Submit();
                            System.Threading.Thread.Sleep(5000);
                            loginOpened3 = true;
                            console.WriteLine("Inicio de sesión exitoso.");
                        }
                        #endregion

                        #region Buscar la oportunidad 
                        //console.WriteLine($"Buscando la oportunidad ...");
                        System.Threading.Thread.Sleep(3000);
                        //Ingresar la oportunidad a buscar
                        try
                        {
                            chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(oppGestion.opp);
                        }
                        catch (Exception p)
                        {
                            System.Threading.Thread.Sleep(10000);
                            chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(oppGestion.opp);

                        }
                        //Click buscar               

                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();

                        System.Threading.Thread.Sleep(2000);

                        //Primer registro
                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                        System.Threading.Thread.Sleep(8000);
                        console.WriteLine("Oportunidad encontrada.");

                        #endregion

                        #region Requerimientos de BAW
                        //console.WriteLine("Ingresando los requerimientos de BAW de la oportunidad.");


                        //Cambio a IFrame
                        IWebElement iframe = chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/div[4]//following-sibling::iframe[1]"));
                        chrome.SwitchTo().Frame(iframe);

                        System.Threading.Thread.Sleep(1000);

                        int quantityRowsPerPAge = 5;

                        int page = 0;

                        string updateQueryBAW = "";

                        for (int i = 1; i <= oppGestion.BAW.Count; i++)
                        {
                            //Seleccionar nuevo en el dropdown de acción
                            System.Threading.Thread.Sleep(1000);
                            try
                            {
                                string xpathActionTd = $"/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[3]/table/tbody/tr[{i - (page * 5)}]/td[2]/div/div/div/div/select";
                                SelectElement actionReq = new SelectElement(chrome.FindElement(By.XPath(xpathActionTd)));
                                System.Threading.Thread.Sleep(1000);
                                actionReq.SelectByValue("nuevo");
                                System.Threading.Thread.Sleep(1000);
                            }
                            catch (Exception k)
                            {
                                System.Threading.Thread.Sleep(3000);
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

                                System.Threading.Thread.Sleep(5000);
                                aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_3']")));
                                aux.SelectByValue(oppGestion.BAW[i - 1].vendor.Trim());
                            }
                            catch (Exception)
                            {
                                System.Threading.Thread.Sleep(20000);
                                aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_3']")));
                                aux.SelectByValue(oppGestion.BAW[i - 1].vendor.Trim());
                            };


                            //Nombre de Producto
                            System.Threading.Thread.Sleep(2500);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_4']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].product.Trim());

                            //Tipo de Requerimiento
                            System.Threading.Thread.Sleep(1000);
                            aux = new SelectElement(chrome.FindElement(By.XPath("//*[@id='combo_div_5']")));
                            aux.SelectByValue(oppGestion.BAW[i - 1].requirementType.Trim());

                            //Cantidad
                            System.Threading.Thread.Sleep(1000);
                            chrome.FindElement(By.Id("input_div_2_1_2_1_1")).SendKeys(oppGestion.BAW[i - 1].quantity);
                            //Checkbox Documentos adjuntados.
                            chrome.FindElement(By.XPath("//*[@id='div_7']/div/div[2]/label/input")).Click();

                            if (oppGestion.BAW[i - 1].integration == "Si")
                            {
                                System.Threading.Thread.Sleep(1000);
                                //Checkbox Parte de una integración
                                chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/label/input")).Click();
                            }

                            //Comentarios 
                            chrome.FindElement((By.XPath("//*[@id='textArea_div_9']"))).SendKeys(oppGestion.BAW[i - 1].comments);

                            System.Threading.Thread.Sleep(1000);
                            //Botón de aceptar
                            chrome.FindElement(By.XPath("//*[@id='div_12_1_4_1_1_1_1']/div/button")).Click();



                            if (i != oppGestion.BAW.Count) //Descartar tocar el botón + en el último registro
                            {
                                System.Threading.Thread.Sleep(2500);
                                //Boton de + en requerimiento
                                try
                                {
                                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[4]/button")).Click();
                                }
                                catch (Exception e)
                                {
                                    System.Threading.Thread.Sleep(30000);
                                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[4]/button")).Click();

                                }
                            }
                            #endregion

                            //Significa que la iteración actual es divisible y entera
                            if (i % quantityRowsPerPAge == 0)
                            {
                                page += 1;
                            }

                            //Se actualiza el status en la tabla BAW.
                            updateQueryBAW += $" UPDATE `BAW` SET `statusBAW`= 2 WHERE id={oppGestion.BAW[i - 1].id}; ";

                        }

                        //Botón final verde de continuar y finalizar 1 etapa de gestión en espera de respuesta del especialista.
                        System.Threading.Thread.Sleep(2000);
                        chrome.FindElement(By.XPath("//*[@id='div_7_1_4_1_1_1_1']/div/button")).Click();
                        console.WriteLine($"Requisitos de la oportunidad {oppGestion.opp} creados en BAW.");
                        console.WriteLine($"");

                        System.Threading.Thread.Sleep(6000);

                        //Salir del iframe
                        chrome.SwitchTo().DefaultContent();
                        System.Threading.Thread.Sleep(1000);

                        //Limpiar el buscador 
                        chrome.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/div/a/i")).Click();
                        System.Threading.Thread.Sleep(2000);



                        #endregion

                        #region Actualizar status y log de cambios 

                        ///string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=4 WHERE id= {oppGestion.id}";
                        updateQueryBAW += $" UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=4 WHERE id= {oppGestion.id}; ";
                        crud.Update(updateQueryBAW, "autopp2_db", enviroment);

                        //log.LogDeCambios("Creación", "Autopp", oppGestion.employee, "", oppGestion.opp, "Creación de BAW");
                        #endregion
                    }
                    catch (Exception e)
                    {
                        string msg = "Este error está en el try catch de la Fase 3: Crear requisitos BAW  - Autopp. Revisar el screenshot del error en la carpeta en el desktop/Databot/Autopp/ErrorsScreenshots";
                        sett.SendError(this.GetType(), $"Error al crear requisitos en BAW del id #{idOpp}", msg, e);

                        string updateQuery = $"UPDATE OppRequests SET status = 8, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                        crud.Update(updateQuery, "autopp2_db", enviroment);


                        try
                        {
                            string pathErrors = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\databot\Autopp\ErrorsScreenshots\";
                            Screenshot TakeScreenshot = ((ITakesScreenshot)chrome).GetScreenshot();
                            TakeScreenshot.SaveAsFile(pathErrors + $"{idOpp} - {oppReq3["opp"]} - crear requisitos BAW.png");
                        }
                        catch (Exception i)
                        {
                            console.WriteLine("No se pudo guardar el screenshot del error. El error al guardar fue:");
                            console.WriteLine("");
                            console.WriteLine(i.ToString());
                        }

                        //Cerrar chrome
                        process.KillProcess("chromedriver", true);
                        process.KillProcess("chrome", true);

                        //Establecer nueva conexión 
                        chrome = SelConn(GetUrlBaw());
                        loginOpened3 = false;

                    }

                    indexReqs3++;
                }

                System.Threading.Thread.Sleep(5000);

                #region Cerrar Chrome
                //process.KillProcess("chromedriver", true);
                //process.KillProcess("chrome", true);
                #endregion
            }
            #endregion

            #region Status 4- "En proceso aprobación final BAW" 
            bool displayedTitle = false;

            //Recorre cada una de las notificaciones BAW en cola.
            while (mail.GetAttachmentEmail("BAW Notificaciones", "Procesados", "Procesados BAW"))
            {
                //Subject de ejemplo:"Recibir y verificar solicitud GTL. Número de oportunidad:0000178752. Número de caso: 5042-7. Proveedor: Lenovo"
                string subject = root.Subject.Replace(" ", "");

                //Si es ese subject, realice el proceso.
                if (subject.Contains("Recibir y verificar solicitud GTL. Número de oportunidad".Replace(" ", "")))
                {

                    //Para mostrarlo una sola vez.
                    if (!displayedTitle)
                    {
                        console.WriteLine("");
                        console.WriteLine("*****************************");
                        console.WriteLine("*Fase 4: Aprobación final BAW*");
                        console.WriteLine("******************************");
                        console.WriteLine("");

                        displayedTitle = true;
                    }
                    string oppNumber = "";
                    string caseNumber = "";

                    //Desestructura para averiguar el número de oportunidad y número de caso.
                    try
                    {
                        oppNumber = subject.Split('.')[1].Split(':')[1];
                        caseNumber = subject.Split('.')[2].Split(':')[1];
                    }
                    catch (Exception e)
                    {
                        string msg = $"Error al extraer el número de oportunidad en el subject, en aprobación final BAW. El subject es: {subject}";
                        sett.SendError(this.GetType(), $"Error al extraer el número de oportunidad", msg, e);
                        return; //Para que se salga de esta iteración
                    }

                    //console.WriteLine($"Se va a proceder con la aprobación final de la oportunidad {oppNumber}");

                    #region Obtener el registro según el número de opp.

                    DataRow oppRow;
                    //Extraer el oppRequest
                    string sqlOpp = $"SELECT * FROM `OppRequests` WHERE opp='{oppNumber}' and active=1";
                    try
                    {
                        oppRow = crud.Select(sqlOpp, "autopp2_db", enviroment).Rows[0];
                    }
                    catch (Exception e)
                    {
                        console.WriteLine($"La oportunidad {oppNumber} - relacionada al case number {caseNumber} no existe en los registros internos del robot.");
                        return;
                    }
                    #endregion

                    if (oppRow != null)
                    {
                        #region Actualizar los registros de status 2 a 3(rechazado), debido a que no se sabe cuando el especialista rechaza, entonces si recibe un correo de arpobacion se cambia.
                        string sqlUpdate = $"UPDATE `BAW` SET `statusBAW`=3  WHERE oppId= (SELECT oppR.id FROM OppRequests oppR WHERE oppR.opp= '{oppNumber}') and statusBAW=2";
                        crud.Update(sqlUpdate, "autopp2_db", enviroment);
                        #endregion

                        console.WriteLine("");
                        console.WriteLine($"Procesando solicitud # {oppRow["id"]} - {oppNumber}  Case Number # {caseNumber}");


                        IWebDriver chrome4 = SelConn(GetUrlBaw());
                        try
                        {
                            #region Aprobación final en Baw 

                            #region Login en Baw
                            //console.WriteLine("Iniciando sesión en baw...");

                            chrome4.FindElement(By.Id("username")).SendKeys(cred.user_baw);
                            chrome4.FindElement(By.Id("password")).SendKeys(cred.password_baw);

                            chrome4.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/a[1]")).Submit();
                            System.Threading.Thread.Sleep(6000);

                            console.WriteLine("Inicio de sesión exitoso.");
                            #endregion

                            #region Buscar la oportunidad 
                            console.WriteLine($"Buscando la oportunidad...");
                            //Ingresar la oportunidad a buscar
                            try
                            {
                                chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber.Trim());
                            }
                            catch (Exception b)
                            {
                                System.Threading.Thread.Sleep(35000);
                                chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[1]/div[1]/input")).SendKeys(caseNumber.Trim());
                            }
                            //Click buscar 
                            chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();

                            System.Threading.Thread.Sleep(1500);

                            bool continueAprobation = false;
                            //Primer registro
                            try
                            {
                                //Intentar tocar el primer registro.
                                chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                                System.Threading.Thread.Sleep(1000);
                                continueAprobation = true; System.Threading.Thread.Sleep(1000);
                                console.WriteLine("Encontrada.");
                            }
                            catch (Exception e)
                            {
                                try
                                {
                                    //Segundo intento de tocar el primer registro
                                    chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_1']/div/div[1]/div[2]/button/i")).Click();
                                    System.Threading.Thread.Sleep(4000);
                                    chrome4.FindElement(By.XPath("//*[@id='div_1_2_1_2_2']/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[1]/a")).Click();
                                    continueAprobation = true;
                                    System.Threading.Thread.Sleep(1000);
                                    console.WriteLine("Encontrada.");


                                }
                                catch (Exception k)
                                {
                                    //Notificación de que el robot no encontró el primer registro, en otras palabras BAW lo cerró sólo.
                                    string pathError = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                                        @"\databot\Autopp\ErrorsScreenshots\FinalApprovalsNotFounds\" + $"{oppRow["id"]} - {oppRow["opp"]} - aprobación no encontrada.png";

                                    Screenshot TakeScreenshot = ((ITakesScreenshot)chrome4).GetScreenshot();
                                    TakeScreenshot.SaveAsFile(pathError);


                                    string[] cc = { "epiedra@gbm.net" };
                                    string[] att = { pathError };
                                    string body = process.greeting() + $"\r\nEl robot no logró encontrar el requerimiento número {caseNumber} de la oportunidad {oppRow["opp"]} - {oppRow["id"]}" +
                                        $"para realizar la aprobación final en BAW. Por favor revisar si se cerró el caso adecuadamente. ";

                                    mail.SendHTMLMail(body, new string[] { "epiedra@gbm.net" }, $"No se encontró el requerimiento de aprobación final - {oppRow["opp"]} - Autopp", cc, att);

                                    console.WriteLine($"BAW no encontró la aprobación del case number: {caseNumber}, de la oportunidad {oppRow["id"]} - {oppNumber}, revisar si el requerimiento cerró adecuadamente.");

                                }
                            }

                            #endregion

                            #region Ejecutar la aprobación final BAW
                            if (continueAprobation == true) //Si encontró el requerimiento en el buscador.
                            {
                                console.WriteLine("Realizando la aprobación final en BAW.");

                                System.Threading.Thread.Sleep(2000);

                                //Cambio a IFrame
                                IWebElement iframe = chrome4.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/div[4]//following-sibling::iframe[1]"));
                                chrome4.SwitchTo().Frame(iframe);

                                //Seleccionar aceptado
                                System.Threading.Thread.Sleep(1000);
                                SelectElement aux1 = new SelectElement(chrome4.FindElement(By.XPath("//*[@id='combo_div_3']")));
                                aux1.SelectByValue("aceptado");

                                //Botón final verde de continuar y finalizar 1 etapa de gestión en espera de respuesta del especialista.
                                System.Threading.Thread.Sleep(1500);
                                chrome4.FindElement(By.XPath("//*[@id='div_12_1_4_1_1_1_1']/div/button")).Click();
                                System.Threading.Thread.Sleep(1000);
                                //console.WriteLine($"Aprobación de oportunidad {oppNumber} finalizada con éxito.");

                            }
                            #region Cerrar Chrome
                            process.KillProcess("chromedriver", true);
                            process.KillProcess("chrome", true);
                            #endregion

                            #endregion



                            #endregion

                            #region Objeto principal donde se une toda la información
                            AutoppInformation oppGestion = new AutoppInformation();
                            oppGestion.id = oppRow["id"].ToString();
                            oppGestion.employee = oppRow["createdBy"].ToString();
                            oppGestion.opp = oppRow["opp"].ToString();

                            CCEmployee employeeData = new CCEmployee(oppGestion.employee);
                            oppGestion.employee = employeeData.IdEmpleado;
                            #endregion

                            if (continueAprobation)
                                console.WriteLine($"Case number {caseNumber} aprobado con éxito.");

                            #region Actualizar status y log de cambios 

                            string updateQuery =
                                $"UPDATE `BAW` SET `statusBAW`=4 WHERE oppId= (SELECT oppR.id FROM OppRequests oppR WHERE oppR.opp= '{oppNumber.Trim()}') and caseNumber= '{caseNumber.Trim()}'; ";
                            crud.Update(updateQuery, "autopp2_db", enviroment);

                            //Consulta cuantos archivos de BAW están con error
                            DataTable countErrorsTable = new DataTable();
                            string sql1 = $"SELECT COUNT(*) countErrors FROM `BAW`WHERE oppId = '{oppGestion.id}' AND statusBAW = 8 ";
                            countErrorsTable = crud.Select(sql1, "autopp2_db", enviroment);

                            string countErrors = countErrorsTable.Rows[0]["countErrors"].ToString();


                            if (countErrors == "0") //No hay errores BAW
                            {
                                updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=5 WHERE id= {oppGestion.id}; ";
                                crud.Update(updateQuery, "autopp2_db", enviroment);
                            }


                            //log.LogDeCambios("Creación", "Autopp", oppGestion.employee, "", oppGestion.opp, "Creación de BAW");
                            #endregion


                        }
                        catch (Exception e)
                        {
                            string msg = $"Este error está en el try catch de la Fase 4: Aprobación final BAW - Autopp. Revisar el screenshot del error en la carpeta en el desktop/Databot/Autopp/ErrorsScreenshots" +
                            $"La oppNumber es: {oppNumber} y  su Case number es: {caseNumber} ";

                            sett.SendError(this.GetType(), $"Error al realizar la aprobación final BAW #{oppRow["id"]}", msg, e);

                            string updateQuery =
                                $"UPDATE `BAW` SET `statusBAW`=8 WHERE caseNumber= '{caseNumber.Trim()}'; " +
                                $"UPDATE OppRequests SET status = 9, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {oppRow["id"]} ;";
                            crud.Update(updateQuery, "autopp2_db", enviroment);

                            try
                            {
                                string pathErrors = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\databot\Autopp\ErrorsScreenshots\";
                                Screenshot TakeScreenshot = ((ITakesScreenshot)chrome4).GetScreenshot();
                                TakeScreenshot.SaveAsFile(pathErrors + $"{oppRow["id"]} - {oppRow["opp"]} - error en aprobación final.png");
                            }
                            catch (Exception i)
                            {
                                console.WriteLine("No se pudo guardar el screenshot del error. El error al guardar fue:");
                                console.WriteLine("");
                                console.WriteLine(i.ToString());
                            }
                        }


                    }
                }



            }

            #region Cerrar Chrome
            process.KillProcess("chromedriver", true);
            process.KillProcess("chrome", true);
            #endregion

            #endregion

        }






    }
}







#region JSON para almacenar información de la oportunidad.
namespace DataBotV5.Automation.QAS.Autopp
{

    public class AutoppInformation
    {
        public string id { get; set; }
        public string identificator { get; set; }
        public GeneralData generalData { get; set; }
        public OrganizationAndClientData organizationAndClientData { get; set; }
        public List<SalesTeams> salesTeams { get; set; }
        public DataTable LDRS { get; set; }
        public List<DataBAW> BAW { get; set; }
        public string status { get; set; }
        public string employee { get; set; }
        public List<string> files { get; set; }
        public string opp { get; set; }
    }

    public class GeneralData
    {
        public string typeOpportunity { get; set; }
        public string typeOpportunityName { get; set; }
        public string description { get; set; }
        public string initialDate { get; set; }
        public string finalDate { get; set; }
        public string cycle { get; set; }
        public string sourceOpportunity { get; set; }
        public string salesType { get; set; }
        public string outsourcing { get; set; }
    }

    public class OrganizationAndClientData
    {
        public string client { get; set; }
        public string contact { get; set; }
        public string salesOrganization { get; set; }
        public string servicesOrganization { get; set; }
    }

    public class SalesTeams
    {
        public string role { get; set; }
        public string employee { get; set; }
    }

    public class LDRSAutopp
    {
        public string technology { get; set; }
        public List<ItemLDRAutopp> LDR { get; set; }
    }

    public class ItemLDRAutopp
    {
        public string id { get; set; }
        public string value { get; set; }
    }

    public class DataBAW
    {
        public string id { get; set; }
        public string oppId { get; set; }
        public string vendor { get; set; }
        public string product { get; set; }
        public string requirementType { get; set; }
        public string quantity { get; set; }
        public string integration { get; set; }
        public string comments { get; set; }
        /*public string statusBAW { get; set; }
        public string caseNumber { get; set; }*/
    }


}
#endregion











