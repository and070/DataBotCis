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
    /// Este robot gestiona todas las notificaciones provenientes del portal de BAW GTL, para asimismo notificar 
    /// al usuario el status de un caso, y también para modificar los mismos en el portal de Fábrica de Propuestas de Smart & Simple.
    ///    
    /// Coded by: Eduardo Piedra Sanabria - Application Management Analyst
    /// </summary>
    class AutoppEmails
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
        DataRow employeeCreatorData;
        bool executeStats = false;
        String notificationsConfig;
        string functionalUser = "";
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
            Step456EmailsActionsBAW();

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
            while (mail.GetBodyAndOutlookConnection("BAW Notificaciones", "Procesados", "Procesados BAW"))
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

                                    mail.SendMail(bodyEr, "epiedra@gbm.net", $"No se encontró el requerimiento de aprobación final - {oppRow["opp"]} - Autopp", 2, cc, att);

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


                            log.LogDeCambios("Verficación final", root.BDProcess, oppGestion.employee, "Recibir y verificar solicitud", "Realizar la verificación final para finalizar el proceso del client y opp: " + oppGestion.organizationAndClientData.client + " -  " + oppGestion.opp, oppGestion.employee);
                            respFinal = respFinal + "\\n" + "Realizar la verificación final para finalizar el proceso del client y opp: " + oppGestion.organizationAndClientData.client + " -  " + oppGestion.opp;
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
                             $"\r\n 1. Ingresar a <a href='{routeSS}'>Smart And Simple</a> de GBM, seguidamente en Ventas - Autopp - Inicio - Mis Solicitudes." +
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

                            log.LogDeCambios("Verficación final", root.BDProcess, oppGestion.employee, "Se estableció en devuelto el caso BAW:" + caseNumber, "Se estableció en devuelto el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString(), oppGestion.employee);
                            respFinal = respFinal + "\\n" + "Se estableció en devuelto el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString();

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

                            log.LogDeCambios("Rechazo de caso BAW", root.BDProcess, oppGestion.employee, "Se estableció en rechazado el caso BAW:" + caseNumber, "Se estableció en rechazado el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString(), oppGestion.employee);
                            respFinal = respFinal + "\\n" + "Se estableció en rechazado el  caso BAW: " + caseNumber + " - " + dr["oppNumber"].ToString();

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

                }


            }
            #endregion

            #region Fallo

            else if (successOrFailMode == "Fail") //Fallo
            {
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



















