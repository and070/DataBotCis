using DataBotV5.App.ConsoleApp;
using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.Autopp;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.EasyLDR;
using DataBotV5.Logical.Webex;
using OpenQA.Selenium;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;

namespace DataBotV5.Automation.WEB.Autopp
{
    /// <summary>
    /// Automatiza la creación de un documento de Excel donde se plasma el LDR establecido en el Portal web 
    /// Fábrica de Propuesta en Smart&Simple enviado por vendedores regionales, todo esto con el propósito de 
    /// agilizar el inicio de proceso de ventas de GBM, disminuyendo el error humano y subir el documento a la 
    /// herramienta de SAP.
    ///    
    /// Coded by: Eduardo Piedra Sanabria - Application Management Analyst
    /// </summary>
    class AutoppLDR
    {

        #region Variables locales 
        Logical.Projects.AutoppSS.AutoppLogical logical = new Logical.Projects.AutoppSS.AutoppLogical();
        ProcessInteraction process = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        AutoppSQL autoppSQL = new AutoppSQL();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        WebexTeams webex = new WebexTeams();
        string enviroment = Start.enviroment;
        DataRow employeeResponsibleData;
        Settings sett = new Settings();
        AutoppInformation oppGestion;
        string LDROrBOMDocument = "";
        Rooting root = new Rooting();
        DataRow employeeCreatorData;
        bool executeStats = false;
        string functionalUser = "";
        String notificationsConfig;
        string sapSystem = "CRM";
        DataTable configuration;
        string userAdmin = "";
        string caseNumber = "";
        CRUD crud = new CRUD();
        string respFinal = "";
        DataTable salesTInfo;
        Log log = new Log();
        int mandante = 0;
        DataRow client;
        int idOpp;




        #endregion

        public void Main()
        {

            console.WriteLine("Consultando nuevas solicitudes...");
            ProcessLDR();

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
        public void ProcessLDR()
        {

            #region Status - "Creando documento LDR" 
            idOpp = 0;
            DataTable newOppRequests1 = GetReqsForStatus("13");
            if (newOppRequests1.Rows.Count > 0)
            {

                executeStats = true;
                int indexReqs1 = 1;

                GetAutoppConfiguration();

                //Establece el mandante de SAP según el entorno.
                mandante = sap.checkDefault(sapSystem, 0);


                foreach (DataRow oppReq in newOppRequests1.Rows)
                {

                    console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes de creación de LDR.");

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

                    //Nuevos Información de cliente
                    organizationAndClientData.requestExecutive = orgInfo.Rows[0]["requestExecutive"].ToString();
                    organizationAndClientData.positionExecutive = orgInfo.Rows[0]["positionExecutive"].ToString();
                    organizationAndClientData.emailExecutive = orgInfo.Rows[0]["emailExecutive"].ToString();
                    organizationAndClientData.phoneExecutive = orgInfo.Rows[0]["phoneExecutive"].ToString();
                    organizationAndClientData.deliveryAddress = orgInfo.Rows[0]["deliveryAddress"].ToString();
                    organizationAndClientData.openingHours = orgInfo.Rows[0]["openingHours"].ToString();
                    organizationAndClientData.clientWebSide = orgInfo.Rows[0]["clientWebSide"].ToString();
                    organizationAndClientData.clientProblem = orgInfo.Rows[0]["clientProblem"].ToString();
                    organizationAndClientData.basicNecesity = orgInfo.Rows[0]["basicNecesity"].ToString();
                    organizationAndClientData.expectationDate = orgInfo.Rows[0]["expectationDate"].ToString();
                    organizationAndClientData.haveAnySolution = orgInfo.Rows[0]["haveAnySolution"].ToString();
                    organizationAndClientData.anothersNotes = orgInfo.Rows[0]["anothersNotes"].ToString();

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


                    //console.WriteLine($"Procesando solicitud {indexReqs1} de {newOppRequests1.Rows.Count} solicitudes.");
                    console.WriteLine("");
                    console.WriteLine($"Solicitud id {oppGestion.id} - {client["name"]}");


                    #region Paso - En proceso crear LDRS
                    bool resultStep2 = false;

                    resultStep2 = Step2CreateLDR();
                    #endregion

                    if (resultStep2)
                    {
                        //Si es GTL Quotation, que mande a crear casos en BAW
                        if (oppGestion.generalData.cycle == "Y3A")
                        {
                            //Envío a creación de BAW
                            string updateQuery = $"UPDATE OppRequests SET opp ='{oppGestion.opp}', updatedAt= CURRENT_TIMESTAMP, updatedBy= 'Databot', status=14 WHERE id= {oppGestion.id}; ";
                            crud.Update(updateQuery, "autopp2_db", enviroment);
                            console.WriteLine($"Se cambia el estado para creación de BAW");

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
        /// Crea el LDR (levantamiento de requerimiento), y lo sube a SAP y FTP de Smart And Simple.
        /// </summary>
        /// <returns>Retorna true si todo salió sin ningún error.</returns>
        public bool Step2CreateLDR()
        {

            #region Paso 2- "En proceso crear LDRS" 


            console.WriteLine("");
            console.WriteLine("********************");
            console.WriteLine("*Fase: Crear LDRS*");
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
                        sap.BlockUser(sapSystem, 1, mandante);
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
                    //sap.LogSAP(sapSystem, mandante);
                    sap.LogSAP(sapSystem);

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

                        log.LogDeCambios("Creación", root.BDProcess, oppGestion.employee, "Creación de LDR", "Creación de LDR o subir archivos a SAP y FTP del cliente: " + oppGestion.organizationAndClientData.client, oppGestion.employee);
                        respFinal = respFinal + "\\n" + "Creación de LDR o subir archivos a SAP y FTP del cliente: " + oppGestion.organizationAndClientData.client;

                    }

                    #endregion

                }


            }
            catch (Exception e)
            {
                //Notificar error al usuario.
                NotifySuccessOrErrors("Fail", 2, errorsList);

                string msg = "Este error está en el try catch de la Fase 2: Crear LDRS - Autopp. " + e.StackTrace;
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

        #endregion


        #region Métodos útiles para la gestión de cada uno de los pasos de AutoppProcess.

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


                #region Llenando la información general
                console.WriteLine("Llenando la información general del LDR");

                // Obtener la hoja de trabajo información general.
                Excel.Worksheet xlInformationWorksheet = xlWorkBook.Sheets["generalInformation"];

                // Establecerse en la hoja localizada
                xlWorkSheet = xlInformationWorksheet;

                //Nombre
                xlWorkSheet.Name = "Información General";

                // Cliente
                xlWorkSheet.Cells[7, 2].value =  client["name"];

                // Oportunidad
                xlWorkSheet.Cells[8, 2].value = oppInformation.opp;

                // Ejecutivo de la solicitud
                xlWorkSheet.Cells[10, 2].value = oppInformation.organizationAndClientData.requestExecutive;

                // Posición del Ejecutivo
                xlWorkSheet.Cells[11, 2].value = oppInformation.organizationAndClientData.positionExecutive;

                // Correo electrónico del ejecutivo
                xlWorkSheet.Cells[12, 2].value = oppInformation.organizationAndClientData.emailExecutive;

                // Número de Teléfono
                xlWorkSheet.Cells[13, 2].value = oppInformation.organizationAndClientData.phoneExecutive;

                // Dirección de entrega del servicio
                xlWorkSheet.Cells[14, 2].value = oppInformation.organizationAndClientData.deliveryAddress;

                // Horario de Atención
                xlWorkSheet.Cells[15, 2].value = oppInformation.organizationAndClientData.openingHours;

                // Dirección de la Página web del cliente
                xlWorkSheet.Cells[16, 2].value = oppInformation.organizationAndClientData.clientWebSide;

                // ¿Existe un problema o dolor asociado?
                xlWorkSheet.Cells[18, 2].value = oppInformation.organizationAndClientData.clientProblem;

                // ¿Cuál es la necesidad básica?
                xlWorkSheet.Cells[19, 2].value = oppInformation.organizationAndClientData.basicNecesity;

                // ¿Tiene alguna expectativa para cuándo requiere la solución operativa?
                xlWorkSheet.Cells[20, 2].value = oppInformation.organizationAndClientData.expectationDate;

                // ¿Se cuenta actualmente con una solución que cumpla (parcial/total) estos requerimientos?
                xlWorkSheet.Cells[21, 2].value = oppInformation.organizationAndClientData.haveAnySolution;

                // Notas Adicional
                xlWorkSheet.Cells[22, 2].value = oppInformation.organizationAndClientData.anothersNotes;

                //Lista de campos a realizar wraptext
                ArrayList fields = new ArrayList() {14,16,21,22 };

                foreach(int field in fields)
                {
                    // Accede a la celda específica
                    Excel.Range cell = xlWorkSheet.Cells[field, 2];

                    // Activa el ajuste de texto (wrap text) para la celda
                    cell.WrapText = true;
                }



                #endregion

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
                    $@"SELECT 
                    databot_db.clients.idClient , 
                    contact, 
                    databot_db.salesOrganizations.salesOrgId , 
                    databot_db.serviceOrganizations.servOrgId,
                    
                    org.*


                    FROM `OrganizationAndClientData` org 
                    LEFT JOIN databot_db.salesOrganizations ON databot_db.salesOrganizations.id = org.salesOrganization 
                    LEFT JOIN databot_db.serviceOrganizations ON databot_db.serviceOrganizations.id = org.servicesOrganization 
                    LEFT JOIN databot_db.clients ON databot_db.clients.id = org.client

                     WHERE org.oppId = {oppId} 
                    AND org.active = 1; ";
                    break;

                case "salesTeam":
                    sql =
                    $@"SELECT salesT.oppId, 
                    EmployeeRole.code, 
                    EmployeeRole.employeeRole, 
                    MIS.digital_sign.user, 
                    MIS.digital_sign.UserID, 
                    EmployeeRole.id 



                    FROM `SalesTeam` salesT 

                    LEFT JOIN EmployeeRole ON EmployeeRole.id = salesT.role 
                    LEFT JOIN MIS.digital_sign ON MIS.digital_sign.id = salesT.employee 

                    WHERE salesT.oppId = {oppId} 
                    AND salesT.active = 1";
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
                    case 2:
                        #region Error al crear el LDR

                        if (oppGestion.opp != "")
                        {
                            #region Notificación por Webex Teams al usuario
                            titleWebex = $"Error al crear LDR - {oppGestion.opp} - {oppGestion.id}";

                            string msgWebex2 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                            $"del cliente: " + client["name"].ToString();
                            msgWebex2 += ", ha tenido un problema al crear y subir el archivo LDR a SAP y al servidor.\r\nPor favor contáctese con Application Management y Support.";

                            webex.SendCCNotification(setUser(employeeCreatorData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear LDR", oppGestion.opp, msgWebex2);


                            //En caso que sea diferente, notifique al usuario responsable también
                            if (employeeCreatorData["user"].ToString() != /*employeeResponsibleData.Usuario*/employeeResponsibleData["user"].ToString())
                            {

                                //Seleccionar el nombre y primer letra en mayúscula.
                                try { employeeName = BuildFirstName(employeeResponsibleData["name"].ToString()); } catch (Exception e) { employeeName = ""; }


                                msgWebex2 = process.greeting() + $" estimado(a) {employeeName}.\r\n \r\nLe informo que la oportunidad de la gestión #{oppGestion.id} " +
                                    $"del cliente: {client["name"].ToString()} el cual lo asignaron a usted como empleado(a) responsable";
                                msgWebex2 += ", ha tenido un problema al crear y subir el archivo LDR a SAP y al servidor.\r\nPor favor contáctese con Application Management y Support.";


                                //webex.SendCCNotification(/*employeeResponsibleData.Usuario*/userAdmin + "@GBM.NET", titleWebex, "Error al crear LDR", oppGestion.opp, msgWebex2);
                                webex.SendCCNotification(setUser(employeeResponsibleData["user"].ToString()) + "@GBM.NET", titleWebex, "Error al crear LDR", oppGestion.opp, msgWebex2);
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























