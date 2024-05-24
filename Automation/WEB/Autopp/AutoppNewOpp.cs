using System.Runtime.InteropServices;
using Exception = System.Exception;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Webex;
using DataBotV5.App.ConsoleApp;
using DataBotV5.Data.Stats;
using System.Globalization;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using OpenQA.Selenium;
using System.Linq;
using System.Data;
using System;

namespace DataBotV5.Automation.WEB.Autopp
{
    /// <summary>
    /// Automatiza la creación de una oportunidad en SAP, partiendo de información general como tipo de venta,
    /// origen de la oportunidad, ciclo de venta, el cliente, y su respectivo contacto. Este robot es perteneciente 
    /// al portal web en Smart&Simple llamado Fábrica de Propuestas, diseñado para que los vendedores regionales soliciten
    /// al robot de MIS la creación de oportunidades de venta en SAP CRM, todo esto con el propósito de agilizar el inicio 
    /// de proceso de ventas de GBM.
    ///   
    /// Coded by: Eduardo Piedra Sanabria - Application Management Analyst
    /// </summary>
    class AutoppNewOpp
    {

        #region Variables locales 
        Logical.Projects.AutoppSS.AutoppLogical logical = new Logical.Projects.AutoppSS.AutoppLogical();
        ProcessInteraction process = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        string enviroment = Start.enviroment;
        WebexTeams webex = new WebexTeams();
        SapVariants sap = new SapVariants();
        DataRow employeeResponsibleData;
        Settings sett = new Settings();
        AutoppInformation oppGestion;
        Rooting root = new Rooting();
        string LDROrBOMDocument = "";
        DataRow employeeCreatorData;
        String notificationsConfig;
        string functionalUser = "";
        bool executeStats = false;
        string sapSystem = "CRM";
        DataTable configuration;
        string userAdmin = "";
        CRUD crud = new CRUD();
        string respFinal = "";
        DataTable salesTInfo;
        Log log = new Log();
        //int mandante = 460;
        int mandante = 0;
        DataRow client;
        int idOpp;



        #endregion

        public void Main()
        {

            console.WriteLine("Consultando nuevas solicitudes...");

            ProcessOpp();

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
        /// Método principal que realiza la función de crear la oportunidad.
        /// </summary>
        public void ProcessOpp()
        {

            #region Status 1- "En Proceso" 
            idOpp = 0;
            DataTable newOppRequests1 = GetReqsForStatus("1");
            if (newOppRequests1.Rows.Count > 0)
            {

                executeStats = true;
                int indexReqs1 = 1;

                GetAutoppConfiguration();

                //Establece el mandante de SAP según el entorno.
                mandante = sap.checkDefault(sapSystem, 0);

                foreach (DataRow oppReq in newOppRequests1.Rows)
                {

                    console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes de creación de oportunidad.");

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

                    #endregion


                    #region Empleado que creó la oportunidad.
                    //employeeCreatorData = new CCEmployee(oppGestion.employee);
                    //oppGestion.employee = employeeCreatorData.IdEmpleado;
                    string sqlEmployeeCreatorData = $"select * from MIS.digital_sign where user= '{oppGestion.employee}'";
                    employeeCreatorData = crud.Select(sqlEmployeeCreatorData, "MIS", enviroment).Rows[0];



                    #endregion

                    #region Empleado con rol de empleado responsable en la oportunidad.

                    string sqlEmployeeResponsible = $"select * from MIS.digital_sign where id=(SELECT employee FROM autopp2_db.SalesTeam where role= 41 and oppId= {oppGestion.id})";
                    employeeResponsibleData = crud.Select(sqlEmployeeResponsible, "databot_db", enviroment).Rows[0];
                    #endregion

                    #region Extraer el nombre del cliente
                    //Extraer el nombre del cliente.
                    string sqlClient = $"SELECT name FROM `clients` WHERE `idClient` = {organizationAndClientData.client}";
                    client = crud.Select(sqlClient, "databot_db", enviroment).Rows[0];

                    #endregion

                    console.WriteLine("");
                    console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");

                    #region Paso 1 - Crear Oportunidad CRM
                    bool resultStep1 = Step1CreateOpp();
                    #endregion

                    if (resultStep1)
                    {
                        //Si es GTL o Quotation, que mande a crear el LDR
                        if (oppGestion.generalData.cycle == "Y3A" || oppGestion.generalData.cycle == "Y3")
                        {
                            //Envío a creación de LDR
                            string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=13 WHERE id= {oppGestion.id}; ";
                            crud.Update(updateQuery, "autopp2_db", enviroment);

                            console.WriteLine($"Se cambia el estado para creación de LDR.");

                        }
                        else
                        {
                            //Notificación de éxito
                            NotifySuccessOrErrors("Success", 8);

                            //Notificar a los SalesTeams que han sido agregados a la opp
                            NotifySuccessOrErrors("Success", 9);

                            //Finalizar el proceso actualizando estado.
                            string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=5 WHERE id= {oppGestion.id}; ";
                            crud.Update(updateQuery, "autopp2_db", enviroment);

                            console.WriteLine($"Se finaliza el proceso éxitosamente y se notifica a los usuarios.");

                        }

                    }

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
        /// Primer paso, creación de oportunidad en SAP.
        /// </summary>
        /// <returns>Retorna true si todo salió sin ningún error.</returns>

        public bool Step1CreateOpp()
        {

            console.WriteLine("");
            console.WriteLine("**********************************");
            console.WriteLine("*Crear Oportunidad en SAP*");
            console.WriteLine("**********************************");
            console.WriteLine("");



            try
            {

                #region Crear la oportunidad y notificación de éxito o fallo.
                oppGestion.opp = CreateOppCRM(oppGestion, employeeCreatorData["user"].ToString());

                if (oppGestion.opp == "" || oppGestion.opp == null)
                {//Fallo 
                    NotifySuccessOrErrors("Fail", 1);

                    string updateQuery = $"UPDATE OppRequests SET status = 6, updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot' WHERE id= {idOpp}";
                    crud.Update(updateQuery, "autopp2_db", enviroment);
                    return false;
                }

                log.LogDeCambios("Creación", root.BDProcess, oppGestion.employee, "Nueva oportunidad: " + oppGestion.opp, "Se generó la oportunidad: " + oppGestion.opp + " del cliente: " + oppGestion.organizationAndClientData.client, oppGestion.employee);
                respFinal = respFinal + "\\n" + "Se generó una nueva oportunidad: " + oppGestion.opp + " del cliente: " + oppGestion.organizationAndClientData.client;

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

            console.WriteLine("Conectado con SAP CRM - " + mandante);

            RfcRepository repo = destination.Repository;
            IRfcFunction func = repo.CreateFunction("ZOPP_VENTAS");
            IRfcTable general = func.GetTable("GENERAL");
            IRfcTable partners = func.GetTable("PARTNERS");
            //console.WriteLine("Llenando información general de oportunidad");
            func.SetValue("USER", oppInformation.employee);
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
                func.SetValue("USER", user); //Agregado
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

            return idopp;
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
                    case 8:
                        #region Notificación de éxito de creación de la opp.

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

                            logical.AutoppNotifications("successNotification", setUser(employeeCreatorData["user"].ToString()), toReplace);

                            if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                            {
                                logical.AutoppNotifications("successNotification", setUser(employeeResponsibleData["user"].ToString()), toReplace);
                            }

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
                    case 1:
                        #region Error al crear la opp


                        //Notificar al usuario
                        titleWebex = $"Error al crear la oportunidad - {oppGestion.id}";

                        string msgWebex1 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                        $" del cliente: " + client["name"].ToString();
                        msgWebex1 += ", no fue creada debido a un error inesperado.\r\nPor favor contáctese con Application Management y Support.";

                        webex.SendCCNotification(setUser(employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);

                        //En caso que sea diferente, notifique al usuario responsable también
                        if (employeeCreatorData["user"].ToString() != employeeResponsibleData["user"].ToString())
                        {

                            try //Seleccionar el nombre y primer letra en mayúscula.
                            {
                                employeeName = BuildFirstName(employeeResponsibleData["name"].ToString());

                            }
                            catch (Exception e) { employeeName = ""; }

                            msgWebex1 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                                $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                            msgWebex1 += ", no fue creada debido a un error inesperado.\r\nPor favor contáctese con Application Management y Support.";


                            webex.SendCCNotification(setUser(employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);
                        }


                        //Respaldo admin
                        webex.SendCCNotification(userAdmin + "@GBM.NET", titleWebex, "Crear Opp en CRM", oppGestion.opp, msgWebex1);


                        //Notificar a Application Management
                        string msg = "Este error está en el try catch de la fase 1 Crear Oportunidades en SAP - Autopp";
                        sett.SendError(this.GetType(), $"Error al crear la opp id #{idOpp}", msg, exception);

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
























#region JSON para almacenar información de la oportunidad.
namespace DataBotV5.Automation.WEB.Autopp
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

        //Nuevos Información de Cliente
        public string requestExecutive { get; set; }
        public string positionExecutive { get; set; }
        public string emailExecutive { get; set; }
        public string phoneExecutive { get; set; }
        public string deliveryAddress { get; set; }
        public string openingHours { get; set; }
        public string clientWebSide { get; set; }
        public string clientProblem { get; set; }
        public string basicNecesity { get; set; }
        public string expectationDate { get; set; }
        public string haveAnySolution { get; set; }
        public string anothersNotes { get; set; }
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
    }


}
#endregion


