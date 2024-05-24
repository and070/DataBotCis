using System.Runtime.InteropServices;
using Exception = System.Exception;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using DataBotV5.App.ConsoleApp;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Webex;
using OpenQA.Selenium.Chrome;
using DataBotV5.Logical.Mail;
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
    /// Cuando un caso ha sido devuelto por un especialista a un vendedor, este procede a editarlo en 
    /// el Portal de Fábrica de Propuestas de Smart & Simple y volverlo a enviar al especialista, 
    /// este robot se encarga de realizar la gestión de tomar el caso editado por el vendedor, 
    /// e ingresar a BAW GTL para enviar la corrección al especialista, actuando como un 
    /// integrador de las dos plataformas.
    /// 
    /// Coded by: Eduardo Piedra Sanabria - Application Management Analyst
    /// </summary>
    class AutoppForwardBAW
    {

        #region Variables locales 
        Logical.Projects.AutoppSS.AutoppLogical logical = new Logical.Projects.AutoppSS.AutoppLogical();
        ProcessInteraction process = new ProcessInteraction();
        string routeSS = "https://smartsimple.gbm.net/";
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        WebexTeams webex = new WebexTeams();
        string enviroment = Start.enviroment;
        DataRow employeeResponsibleData;
        Settings sett = new Settings();
        AutoppInformation oppGestion;
        Rooting root = new Rooting();
        string functionalUser = "";
        DataRow employeeCreatorData;
        String notificationsConfig;
        bool executeStats = false;
        DataTable configuration;
        string userAdmin = "";
        string caseNumber = "";
        CRUD crud = new CRUD();
        string respFinal = "";
        Log log = new Log();
        DataRow client;
        int idOpp;




        #endregion

        public void Main()
        {

            console.WriteLine("Consultando nuevas solicitudes...");

            Step7ForwardBAWRequirement();

            if (executeStats == true)
            {
                root.BDUserCreatedBy = functionalUser;
                root.requestDetails = respFinal;

                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

            console.WriteLine("Fin del proceso.");

        }



        #region Métodos con cada uno de los pasos del proceso Autopp


  
  
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


                            //string[] cc = { "dmeza@gbm.net", "epiedra@gbm.net", oppRow["createdBy"] + "@gbm.net", /*employeeResponsibleData.Correo*/ employeeCreatorData["email"].ToString(), employeeCreatorData["email"].ToString() };
                            //string[] att = { pathError };
                            //string body = process.greeting() + $"\r\nEl robot no logró encontrar el requerimiento número {caseNumber} de la oportunidad {oppRow["opp"]} - {oppRow["id"]}" +
                            //    $"para realizar la corrección pedida por el especialista en BAW después de la devolución. Por favor revisar si se cerró el caso adecuadamente. ";

                            //mail.SendMail(body, "appmanagement@gbm.net", $"No se encontró el requerimiento para reenviar CaseNumber: #{caseNumber} de - {oppRow["opp"]} - Autopp", 2, cc, att);

                            string[] cc = { "epiedra@gbm.net"/*, oppRow["createdBy"] + "@gbm.net", employeeCreatorData["email"].ToString()*/ };
                            string[] att = { pathError };
                            string body = $"\r\nEl robot no logró encontrar el requerimiento número {caseNumber} de la oportunidad {oppRow["opp"]} - {oppRow["id"]}" +
                                $"para realizar la corrección pedida por el especialista en BAW después de la devolución. Por favor revisar el caso adecuadamente. ";

                            mail.SendMail(body, "epiedra@gbm.net", $"No se encontró el requerimiento para reenviar CaseNumber: #{caseNumber} de - {oppRow["opp"]} - Autopp", 2, cc, att);

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

                        #region Log de Cambios
                        log.LogDeCambios("Verficación final", root.BDProcess, oppGestion.employee, $"Se estableció en devuelto al especialista la solicitud id {oppGestion.id} - {client["name"]}", $"Se estableció en devuelto al especialista la solicitud id {oppGestion.id} - {client["name"]}", oppGestion.employee);
                        respFinal = respFinal + "\\n" + $"Se estableció en devuelto al especialista la solicitud id {oppGestion.id} - {client["name"]}";
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

                }


            }
            #endregion

            #region Fallo

            else if (successOrFailMode == "Fail") //Fallo
            {

                switch (fase)
                {

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
            functionalUser = configuration.Select($"typeConfiguration = 'funcionalUser'")[0]["configuration"].ToString().ToLower();

            #endregion
        }




        #endregion


    }
}










