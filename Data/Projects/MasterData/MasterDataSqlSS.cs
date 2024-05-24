using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Database;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using Newtonsoft.Json.Linq;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using System.Linq;
using System.Data;
using System.IO;
using WinSCP;
using System;
using DataBotV5.Logical.Projects.MasterData;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using DataBotV5.Logical.Webex;

namespace DataBotV5.Data.Projects.MasterData
{
    /// <summary>
    /// Clase Data encargada de manejo de datos maestros.
    /// </summary>
    class MasterDataSqlSS
    {
        Credentials.Credentials cred = new Credentials.Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();
        Database.Database db2 = new Database.Database();
        ProcessInteraction proc = new ProcessInteraction();
        MsExcel excel = new MsExcel();
        MasterDataLogical msl = new MasterDataLogical();
        WebexTeams wt = new WebexTeams();
        string ssMandante = "PRD";


        /// <summary>
        /// Método para obtener la gestión.
        /// </summary>
        /// <param name="type">Ver tabla motherTables de la DB masterData de Smart&Simple, columna Id</param>
        /// 1	Materiales
        /// 2	Clientes
        /// 3	Contactos
        /// 4	Equipos
        /// 5	Ibase
        /// 6	Materiales de Servicio
        /// 7	Servicios
        /// 8	Repuestos
        /// 9	Proveedores
        /// 10	Garantías
        /// <returns></returns>
        public string GetManagement(string type, string status = "2")
        {
            root.requestDetails = "";
            string respuesta = "";
            try
            {
                #region Connection DB
                //sql_select = sqlQueryGetData(type);
                string sql_select = $@"SET SESSION group_concat_max_len = 1000000; SELECT
                                masterDataRequests.*, 
                                factors.factor as factorText,
                                motherTables.generalTable,
                                motherTables.formTable,
                                motherTables.motherTable as formulario,
                                motherTables.factor as factorType,
                                motherTables.templateVersion,
                                typeOfManagement.description as typeOFManagementText,
                                (SELECT GROUP_CONCAT(JSON_OBJECT('name', uploadFiles.name, 'path', uploadFiles.path, 'type', uploadFiles.documentId)) as filesInfoApprovals FROM uploadFiles WHERE uploadFiles.requestId = masterDataRequests.id AND uploadFiles.documentId = 2) as filesApprovals,
                                (SELECT GROUP_CONCAT(JSON_OBJECT('name', uploadFiles.name, 'path', uploadFiles.path, 'type', uploadFiles.documentId)) as filesInfoMass FROM uploadFiles WHERE uploadFiles.requestId = masterDataRequests.id AND (uploadFiles.documentId = 1 OR uploadFiles.documentId = 3)) as filesMass
                                FROM masterDataRequests 
                                INNER JOIN factors ON masterDataRequests.factor = factors.id
                                INNER JOIN motherTables ON masterDataRequests.dataType = motherTables.id
                                INNER JOIN typeOfManagement ON masterDataRequests.typeOfManagement = typeOfManagement.id
                                WHERE masterDataRequests.dataType = {type} AND masterDataRequests.status = {status} AND masterDataRequests.active = 1
                                GROUP BY masterDataRequests.id;";

                DataTable masterDataRequest = crud.Select(sql_select, "masterData");

                #endregion

                if (masterDataRequest.Rows.Count > 0)
                {
                    console.WriteLine("Extraer datos...");
                    #region Establecer información básica de la solicitud
                    root.IdGestionDM = masterDataRequest.Rows[0]["id"].ToString();
                    root.BDUserCreatedBy = masterDataRequest.Rows[0]["createdBy"].ToString().ToLower() + "@gbm.net";
                    root.aprobadorDM = masterDataRequest.Rows[0]["requestApprovers"].ToString();
                    root.factorDM = masterDataRequest.Rows[0]["factorText"].ToString();
                    root.factorType = masterDataRequest.Rows[0]["factorType"].ToString();
                    root.fechaDM = masterDataRequest.Rows[0]["createdAt"].ToString();
                    root.tipo_gestion = masterDataRequest.Rows[0]["typeOfManagement"].ToString();
                    root.typeOfManagementText = masterDataRequest.Rows[0]["typeOFManagementText"].ToString();
                    root.metodoDM = masterDataRequest.Rows[0]["method"].ToString();
                    root.formDm = masterDataRequest.Rows[0]["formulario"].ToString();
                    root.Subject = "Formulario Creación de " + masterDataRequest.Rows[0]["formulario"].ToString() + " - Notificación de Finalización de Gestión - #" + root.IdGestionDM;
                    #endregion
                    #region Extraer los datos generales
                    string sqlGeneralData = sqlQueryDataGeneral(masterDataRequest.Rows[0]["generalTable"].ToString(), root.IdGestionDM);
                    DataTable generalData = crud.Select(sqlGeneralData, "masterData");
                    root.datagDM = Newtonsoft.Json.JsonConvert.SerializeObject(generalData);
                    #endregion


                    #region Verifica y extrae los documentos de aprobador
                    if (!string.IsNullOrEmpty(masterDataRequest.Rows[0]["filesApprovals"].ToString()))
                    {
                        root.doc_aprob = JArray.Parse("[" + masterDataRequest.Rows[0]["filesApprovals"].ToString() + "]");

                    }
                    else
                    {
                        //root.doc_aprob.Clear();
                    }
                    #endregion

                    if (root.metodoDM == "2") //Masivo
                    {
                        #region extraer la plantilla de los archivos de la solicitud
                        console.WriteLine("Buscando el archivo correcto...");
                        try
                        {
                            JArray gestiones = new JArray();
                            string massFiles = masterDataRequest.Rows[0]["filesMass"].ToString();
                            if (massFiles != "Null")
                            {
                                gestiones = JArray.Parse("[" + massFiles + "]");
                            }
                            else
                            {
                                //cambiar solicitud a status error y enviar comunicado de que es masivo pero no se encontró el archivo
                                ChangeStateDM(masterDataRequest.Rows[0]["id"].ToString(), "Error: no se encontró la plantilla", "4");
                                wt.SendNotification(root.BDUserCreatedBy, "No se encontró la plantilla de excel", "La solicitud indicada se encuentra en estado de ERROR debido a que no se encontró la plantilla de la solicitud, por favor comuniquese con Internal Customer Services de MIS");
                                //wt.SendNotification(masterDataRequest.Rows[0]["id"].ToString(),
                                //    masterDataRequest.Rows[0]["createdBy"].ToString(),
                                //    "4",
                                //    "La solicitud indicada se encuentra en estado de ERROR debido a que no se encontró la plantilla de la solicitud, por favor comuniquese con Internal Customer Services de MIS",
                                //    masterDataRequest.Rows[0]["formulario"].ToString(),
                                //    masterDataRequest.Rows[0]["typeOFManagementText"].ToString(),
                                //    masterDataRequest.Rows[0]["factorType"].ToString(),
                                //    masterDataRequest.Rows[0]["factorText"].ToString());
                                return "ERROR";
                            }

                            //verifica si tiene correcion o no
                            bool correctionFile = false;
                            for (int i = 0; i < gestiones.Count; i++)
                            {
                                JObject fila = JObject.Parse(gestiones[i].ToString());
                                string fileType = fila["type"].Value<string>();
                                if (fileType == "3")
                                {
                                    correctionFile = true;
                                }
                            }

                            for (int i = 0; i < gestiones.Count; i++)
                            {
                                JObject fila = JObject.Parse(gestiones[i].ToString());
                                string adjunto = fila["name"].Value<string>();
                                string fileType = fila["type"].Value<string>();
                                if (fileType == "1" && correctionFile)
                                {
                                    continue;
                                }
                                if (Path.GetExtension(adjunto).Substring(0, 4) == ".xls")
                                {
                                    string local_ruta = root.FilesDownloadPath + "\\" + adjunto;
                                    int index = 1;

                                    #region descargar archivo del SFTP   

                                    bool result = DownloadFile(fila["path"].Value<string>());

                                    if (!result)
                                    {
                                        //si da error es porque no se pudo descargar por lo que se deberia de cambiar el estado a error
                                        ChangeStateDM(masterDataRequest.Rows[0]["id"].ToString(), "Error: al descarga la plantilla del SFTP de SS", "4");
                                        wt.SendNotification(root.BDUserCreatedBy, "No se encontró la plantilla de excel", "No se pudó descargar la plantilla de excel de la solicitud indicada, por favor comuniquese con Internal Customer Services de MIS");
                                        return "ERROR";
                                    }
                                    #endregion

                                    #region Abre el excel y verifica que sea la plantilla
                                    bool valFile = false;
                                    DataTable plantilla = excel.GetExcel(local_ruta);
                                    foreach (DataColumn columnName in plantilla.Columns)
                                    {
                                        //verifica si la columna esta en el excel
                                        if (columnName.ColumnName.ToUpper().Contains("X"))
                                        {
                                            valFile = true;
                                            break;
                                        }
                                    }
                                    //bool valVersion = CheckVersion(type, local_ruta, masterDataRequest.Rows[0]["templateVersion"].ToString());
                                    if (valFile)
                                    {

                                        root.ExcelFile = adjunto;
                                    }
                                    else
                                    {
                                        //CAMBIAR EL STATUS A ERROR
                                        ChangeStateDM(masterDataRequest.Rows[0]["id"].ToString(), "Plantilla de excel incorrecta descargue la ultima versión en el portal de Datos Maestros", "4");
                                        wt.SendNotification(root.BDUserCreatedBy, "Plantilla de excel incorrecta", "No se adjuntó la plantilla de excel en su última versión, descargue la misma en el portal de Datos Maestros de Smart&Simple, por favor comuniquese con Internal Customer Services de MIS");
                                        return "ERROR";
                                    }

                                    #endregion
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            console.WriteLine(ex.Message);
                            //cambiar el status a error
                            ChangeStateDM(masterDataRequest.Rows[0]["id"].ToString(), ex.Message, "4");
                            wt.SendNotification(root.BDUserCreatedBy, "Error al leer su solicitud", "Error al leer su solicitud de Datos Maestros, por favor comuniquese con Internal Customer Services de MIS");
                            return "ERROR";
                        }
                        #endregion
                    }
                    else
                    {
                        //Lineal
                        #region Extraer la infomación del formulario de la solicitud
                        string formTable = masterDataRequest.Rows[0]["formTable"].ToString();
                        if (type == "5")
                        {
                            //ibase: determinar si es creacion o modificacion
                            string[] formTablas = formTable.Split(',');

                            formTable = (root.tipo_gestion == "1") ? formTablas[0].Trim() : formTablas[1].Trim();

                        }
                        string sqlFormData = sqlQueryFormData(formTable, root.IdGestionDM);
                        DataTable formData = crud.Select(sqlFormData, "masterData");

                        #region Fix error en el country en clientes
                        if (type == "2")
                        {
                            foreach (DataRow formDataRow in formData.Rows)
                            {
                                string sqlctr = "SELECT country.code FROM formClients LEFT JOIN masterData.country ON masterData.formClients.country = masterData.country.id WHERE requestID = " + root.IdGestionDM;
                                DataTable countryT = crud.Select(sqlctr, "masterData");
                                if (countryT.Rows.Count > 0)
                                {
                                    string country = countryT.Rows[0][0].ToString();
                                    formDataRow["countryCode"] = country;
                                }
                            }
                        }
                        #endregion

                        if (formData.Rows.Count == 0)
                        {
                            //error porque es lineal y no tiene formulario
                            ChangeStateDM(masterDataRequest.Rows[0]["id"].ToString(), "Error en la solicitud ya que no se encontró la información de las lineas de la solicitud", "4");
                            return "ERROR";
                        }
                        else
                        {
                            root.requestDetails = Newtonsoft.Json.JsonConvert.SerializeObject(formData);

                        }
                        #endregion

                    }

                    respuesta = "OK";
                }


            }
            catch (Exception ex)
            {
                //cambiar el status a error
                ChangeStateDM(root.IdGestionDM, ex.Message.ToString().Replace("'", ""), "4");
                wt.SendNotification(root.BDUserCreatedBy, "Error al leer su solicitud", "Error al leer su solicitud de Datos Maestros, por favor comuniquese con Internal Customer Services de MIS");
                return "ERROR";
            }
            return respuesta;
        }

        /// <summary>
        /// Método para cambiar estado en Datos Maestros.
        /// </summary>
        /// <param name="idGestion">id de la solicitud de la tabla masterDataRequest</param>
        /// <param name="response">la respuesta del bot o la razon del cambio de estado</param>
        /// <param name="status">ver tabla status de la DB masterData de Smart&Simple</param>
        /// 1	APROBACION
        /// 2	EN PROCESO
        /// 3	FINALIZADO
        /// 4	ERROR
        /// 5	RECHAZADO
        /// 6	APROBACION CONTADORES
        /// 7	APROBACION GESTORES
        /// 8	APROBACION FACTURACION
        /// 9	APROBACION CONTROLLER
        /// 10	APROBACION PRICE LIST
        /// 11	APROBACION SALES ADMIN
        /// 12	EN REVISION
        /// 13	APROBACION GERENTE VENTAS
        /// 14  PENDIENTE
        /// <returns>true or false</returns>
        public bool ChangeStateDM(string idGestion, string response, string status)
        {
            response = response.Replace("<br>", "\n");
            string fechaf = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            try
            {

                string sqlUpdate = $"UPDATE masterDataRequests SET response = '{response}', status = '{status}', finalizationAt = '{fechaf}' WHERE id = '{idGestion}'";
                crud.Update(sqlUpdate, "masterData");

                string sqlInsert = $"INSERT INTO logRequest (`requestId`, `status`, `botResponse`, `createdBy`) VALUES ('{idGestion}','{status}','{response}','RPAUSER')";
                crud.Insert(sqlInsert, "masterData");

                return true;

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                return false;
            }

        }

        /// <summary>
        /// Metodo para extraer las solicitudes en REVISION de proveedores (YA NO SE USA, se sustituye por GetManagement con status 12)
        /// </summary>
        /// <returns></returns>
        public string GetRequestRevision()
        {
            string response = "";
            root.requestDetails = "";

            try
            {
                //DataTable mytable = crud.Select("Databot", "select * from gestiones_dm where ESTADO = 'EN REVISION' and TIPO_DATO = 'PROVEEDORES'", "automation");

                //if (mytable.Rows.Count > 0)
                //{
                //    console.WriteLine(DateTime.Now + " > > > " + "Extraer datos...");
                //    root.IdGestionDM = mytable.Rows[0][1].ToString(); //ID GESTION
                //    root.BDUserCreatedBy = mytable.Rows[0][2].ToString().ToLower() + "@gbm.net"; //EMPLEADO
                //    root.aprobadorDM = mytable.Rows[0][3].ToString(); //APROBADOR
                //    root.factorDM = mytable.Rows[0][4].ToString(); //FACTOR
                //    root.datagDM = mytable.Rows[0][5].ToString(); //DG
                //    root.fechaDM = mytable.Rows[0][10].ToString(); //TS_CREACION
                //    root.tipo_gestion = mytable.Rows[0][14].ToString(); //TIPO_GESTION
                //    root.metodoDM = mytable.Rows[0][16].ToString();  //METODO
                //    root.Subject = "Formulario Creación de Proveedores - Notificación de Finalización de Gestión - #" + root.IdGestionDM;
                //    root.requestDetails = mytable.Rows[0][9].ToString();  //GESTION
                //    string docAprob = mytable.Rows[0][17].ToString(); //DOC_APROB  root.dm_files_list
                //    if (!String.IsNullOrEmpty(docAprob) && docAprob != "[]")
                //    {
                //        root.doc_aprob = JArray.Parse(docAprob);
                //    }
                //    string mass_aprob = mytable.Rows[0][18].ToString();  //MAS_APROB

                //    response = "OK";
                //}

            }
            catch (Exception)
            {
                response = "ERROR";
            }
            return response;
        }

        /// <summary>Método para agregar un proveedor.</summary>
        public bool AddVendor(string idVendor, string applicant, string idManagement)
        {
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                //string select = "select * from proveedores_dm where id_prov = " + idVendor;
                //mytable = crud.Select("Databot", select, "auditoria_bs");
                #endregion

                //if (mytable.Rows.Count <= 0)
                //{
                //string insert = "INSERT INTO `proveedores_dm`(`id_prov`, `solicitante`, `id_gestion`) VALUES (" + idVendor + ",'" + applicant + "'," + idManagement + ")";
                //crud.Insert("Databot", insert, "auditoria_bs");

                string insertSS = $"INSERT INTO `vendorsBsAudit`(`requestId`, `vendorSapId`, `requester`) VALUES ('{idManagement}','{idVendor}','{applicant}')";
                crud.Insert(insertSS, "masterData");
                //}
            }
            catch (Exception)
            {

            }
            return false;
        }

        /// <summary>Método para obtener la descripción de un proveedor.</summary>
        public string GetVendorDescription(string VendorCat)
        {
            string vendor_descrip = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                string select = $"SELECT codeDescription FROM `vendorCategory` WHERE `code` = '{VendorCat}'";
                mytable = crud.Select(select, "masterData");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    vendor_descrip = mytable.Rows[0]["codeDescription"].ToString(); //ID GESTION
                }
                else
                {
                    vendor_descrip = "D_Mix";
                }
            }
            catch (Exception)
            {
                vendor_descrip = "D_Mix";
            }
            return vendor_descrip;

        }

        /// <summary>Método para verificar la versión.</summary>
        public bool CheckVersion(string template, string atachment, string ver_db)
        {
            //leer la tabla, tomar la version
            string ver_file = "";
            //DataTable xx = crud.Select("Databot", "SELECT * FROM `plantilla_versiones` WHERE `Name` LIKE '" + template + "'", "automation");
            //string ver_db = xx.Rows[0]["version"].ToString();

            //leer la version del file
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder objFolder = shell.NameSpace(Path.GetDirectoryName(atachment));

            foreach (Shell32.FolderItem2 item in objFolder.Items())
            {
                if (item.Name == Path.GetFileName(atachment))
                {
                    ver_file = objFolder.GetDetailsOf(item, 18); //18 es etiqueta
                }
            }

            if (ver_db == ver_file)
                return true;
            else
                return false;
        }

        private string sqlQueryGetData(string masterDataType)
        {
            return $@"
                select
                generalTable.Gestion,
                generalTable.Formulario,
                generalTable.Fecha,
                generalTable.factorId,
                generalTable.Factor,
                generalTable.Estado,
                generalTable.statusType,
                generalTable.Aprobadores,
                generalTable.Respuesta,
                generalTable.createdAt,
                generalTable.createdBy,
                generalTable.methodId,
                generalTable.method,
                generalTable.comment,
                generalTable.commentApproval,
                generalTable.typeOfManagement,
                generalTable.typeOfManagementId,
                
                ClientsTable.ID IdClient,
                ClientsTable.requestId requestIdClient,
                ClientsTable.country countryClient,
                ClientsTable.countryClientId countryClientId,
                ClientsTable.valueTeam valueTeam,
                ClientsTable.valueTeamId valueTeamId,
                ClientsTable.channelName channel,
                ClientsTable.channelId channelId,
                ClientsTable.subjectVat subjectVat,
                ClientsTable.subjectVatId subjectVatId,


                ContactTable.ID IDContact,
                ContactTable.requestId requestIdContact,
                ContactTable.country countryContact,
                ContactTable.countryContactId countryContactId,


                EquipmentTable.ID IDEquipment,
                EquipmentTable.requestId requestIdEquipment,
                EquipmentTable.country countryEquipment,
                EquipmentTable.countryEquipmentId countryEquipmentId,


                IbaseTable.ID IDIbase,
                IbaseTable.requestId requestIdIbase,
                IbaseTable.country countryIbase,
                IbaseTable.countryIbaseId countryIbaseId,





                MaterialsTable.ID IDMaterials,
                MaterialsTable.requestId requestIdMaterials,

                MaterialsTable.materialGroup materialGroupMaterials,
                MaterialsTable.materialGroupId materialGroupMaterialsId,

                MaterialsTable.baw baw,
                MaterialsTable.bawId bawId,

                MaterialsTable.bawManagement bawManagement,





                ServiceMaterialsTable.ID IDServiceMaterials,
                ServiceMaterialsTable.requestId requestIdServiceMaterials,
                ServiceMaterialsTable.materialGroup materialGroupServiceMaterials,
                ServiceMaterialsTable.materialGroupId materialGroupServiceMaterialsId,


                ServicesTable.ID IDServices,
                ServicesTable.requestId requestIdServices,
                ServicesTable.materialGroup materialGroupServices,
                ServicesTable.materialGroupId materialGroupServicesId,



                SpareParts.ID IDSpareParts,
                SpareParts.requestId requestIdSpareParts,
                SpareParts.materialGroupSpartParts materialGroupSpartParts,
                SpareParts.materialGroupId materialGroupSpartPartsId,

                VendorsTable.ID idVendors,
                VendorsTable.requestId requestIdVendors,
                VendorsTable.supplierCompany companyCode,
                VendorsTable.supplierCompanyId companyCodeId,
                VendorsTable.vendorGroup vendorGroup,
                VendorsTable.vendorGroupId vendorGroupId,



                WarrantiesTable.ID IDWarranties,
                WarrantiesTable.requestId requestIdWarranties,
                WarrantiesTable.sendingCountry countryWarranties,
                WarrantiesTable.sendingCountryId countryWarrantiesId



            FROM
                    (select
                        md.ID Gestion,
                                    motherTables.motherTable Formulario,
                                    date_format(md.createdAt, '%d-%m-%Y') AS Fecha,
                                    md.factor factorId,
                                    factors.factor Factor,
                                    status.status Estado,
                                    status.statusType,
                                    md.requestApprovers Aprobadores,
                                    md.response Respuesta,
                                    md.createdAt,
                                    md.createdBy,
                                    md.method methodId,
                                    method.method,
                                    md.comment,
                                    md.commentApproval,
                                    CONCAT(typeOfManagement.code, ' - ', typeOfManagement.description) typeOfManagement,
                                    typeOfManagement.description typeOfManagementSingle,
                                    md.typeOfManagement typeOfManagementId



                                    FROM masterDataRequests md


                                     LEFT JOIN motherTables ON motherTables.ID = md.dataType
                                     LEFT JOIN status ON status.ID = md.status
                                     LEFT JOIN method ON method.ID = md.method
                                     LEFT JOIN typeOfManagement ON typeOfManagement.ID = md.typeOfManagement
                                     LEFT JOIN formIbaseCreation ON formIbaseCreation.requestId = md.ID
                                     LEFT JOIN formClients ON formClients.requestId = md.ID
                                     LEFT JOIN generalDataClients ON generalDataClients.requestId = md.ID
                                     LEFT JOIN factors ON factors.id = md.factor




                                    WHERE md.active = 1 AND md.status = 2 AND md.dataType = {masterDataType}

            ) as generalTable



            LEFT join

                 (SELECT
                            gen.ID,
                            gen.requestId,
                            CONCAT(country.code, ' - ', country.description) country,
                            country.id countryClientId,
                            CONCAT(clientGroup.key, ' - ', clientGroup.valueTeam) valueTeam,
                            clientGroup.id valueTeamId,
                            channel.description channelName,
                            channel.id channelId,
                            CONCAT(subjectVat.code, ' - ', subjectVat.description) subjectVat,
                            subjectVat.id subjectVatId



                            FROM generalDataClients gen, sendingCountry country, databot_db.valueTeam clientGroup, channel, subjectVat


                            WHERE gen.sendingCountry = country.ID
                            AND gen.valueTeam = clientGroup.id
                            AND gen.channel = channel.ID
                            AND gen.subjectVat = subjectVat.ID
                            AND gen.active = 1) 


            as ClientsTable


            on generalTable.Gestion = ClientsTable.requestId



            LEFT join

                  (SELECT
                            gen.ID,
                            gen.requestId,
                            CONCAT(country.code, ' - ', country.description) country,
                            country.id countryContactId



                            FROM generalDataContact gen, sendingCountry country


                            WHERE gen.sendingCountry = country.ID
                            AND gen.active = 1) 


            as ContactTable


            on generalTable.Gestion = ContactTable.requestId



            LEFT join

                (SELECT
                            gen.ID,
                            gen.requestId,
                            CONCAT(country.code, ' - ', country.description) country,
                            country.id countryEquipmentId


                            FROM generalDataEquipment gen, sendingCountry country


                            WHERE gen.sendingCountry = country.ID
                            AND gen.active = 1) 


            as EquipmentTable


            on generalTable.Gestion = EquipmentTable.requestId



            LEFT join

                (SELECT
                            gen.ID, gen.requestId,
                            CONCAT(country.code, ' - ', country.description) country,
                            country.id countryIbaseId


                            FROM generalDataIbase gen, sendingCountry country


                            WHERE gen.sendingCountry = country.ID
                            AND gen.active = 1) 


            as IbaseTable


            on generalTable.Gestion = IbaseTable.requestId


            LEFT join

                (SELECT
                          gen.ID, gen.requestId,

                          CONCAT(mg.code, ' - ', mg.description) materialGroup, 
                          mg.id materialGroupId,

                          baw.description baw,
                          baw.id bawId,

                          gen.bawManagement,
                          gen.id bawManagementId



                          FROM generalDataMaterials gen, materialGroup mg, baw


                          WHERE gen.materialGroup = mg.ID
                          AND gen.baw = baw.ID
                          AND gen.active = 1 ) 


            as MaterialsTable


            on generalTable.Gestion = MaterialsTable.requestId



            LEFT join

                (SELECT
                          gen.ID,
                          gen.requestId,
                          CONCAT(mg.code, ' - ', mg.description) materialGroup,
                          mg.id materialGroupId


                          FROM generalDataServiceMaterials gen, materialGroup mg
                          WHERE gen.materialGroup = mg.ID
                          AND gen.active = 1 ) 


            as ServiceMaterialsTable


            on generalTable.Gestion = ServiceMaterialsTable.requestId



            LEFT join

                        (SELECT gen.ID, gen.requestId,
                        CONCAT(mg.code, ' - ', mg.description) materialGroup,
                        mg.id materialGroupId


                        FROM generalDataServices gen, materialGroup mg


                        WHERE gen.materialGroup = mg.ID
                        AND gen.active = 1) 


            as ServicesTable


            on generalTable.Gestion = ServicesTable.requestId



            LEFT join

                        (
                        SELECT gen.ID, gen.requestId,
                        CONCAT(mg.code, ' - ', mg.description) materialGroupSpartParts,
                        mg.id materialGroupId



                        FROM generalDataSpareParts gen, materialGroupSpartParts mg


                        WHERE gen.materialGroupSpartParts = mg.ID
                        AND gen.active = 1) 


            as SpareParts


            on generalTable.Gestion = SpareParts.requestId


            LEFT join

                        (
                        SELECT gen.ID, gen.requestId,
                        CONCAT(sp.code, ' - ', sp.name) supplierCompany,
                        sp.id supplierCompanyId,

                        CONCAT(vd.code, ' - ', vd.description) vendorGroup,
                        vd.id vendorGroupId



                        FROM generalDataVendors gen,  databot_db.companyCode sp, vendorGroup vd

                        WHERE gen.companyCode = sp.id
                        AND gen.vendorGroup = vd.ID
                        AND gen.active = 1) 


            as VendorsTable


            on generalTable.Gestion = VendorsTable.requestId


            LEFT join

                        (
                        SELECT gen.ID, gen.requestId,
                        CONCAT(sendingCountry.code, ' - ', sendingCountry.description) sendingCountry,
                        sendingCountry.id sendingCountryId




                        FROM generalDataWarranties gen, sendingCountry

                        WHERE gen.sendingCountry = sendingCountry.ID
                        AND gen.active = 1
            ) 


            as WarrantiesTable


            on generalTable.Gestion = WarrantiesTable.requestId
                ";
        }

        /// <summary>
        /// Metodo para extraer la información general de una solicitud y los codigos de sus llaves foreaneas
        /// </summary>
        /// <param name="generalTable">tabla a cual extraer</param>
        /// <param name="id">el requestId de la solicitud en masterDataRequest</param>
        /// <returns>query SELECT para ejecutar</returns>
        private string sqlQueryDataGeneral(string generalTable, string id)
        {
            DataTable formTableInfo = crud.Select($@"SELECT k.TABLE_SCHEMA, k.COLUMN_NAME, k.REFERENCED_TABLE_NAME, k.REFERENCED_COLUMN_NAME, k.REFERENCED_TABLE_SCHEMA
                                                                        FROM information_schema.TABLE_CONSTRAINTS i 
                                                                        LEFT JOIN information_schema.KEY_COLUMN_USAGE k ON i.CONSTRAINT_NAME = k.CONSTRAINT_NAME 
                                                                        WHERE i.CONSTRAINT_TYPE = 'FOREIGN KEY' 
                                                                        AND i.TABLE_SCHEMA = 'masterData'
                                                                        AND i.TABLE_NAME = '{generalTable}'  
                                                                        ORDER BY `k`.`REFERENCED_COLUMN_NAME` ASC", "masterData");


            DataTable formTableColumns = crud.Select($@"SHOW FULL COLUMNS FROM {generalTable}", "masterData");
            string joins = "";
            string sql = $"SELECT ";

            foreach (DataRow fRow in formTableColumns.Rows)
            {
                sql += $"IFNULL({generalTable}.{fRow["Field"]}, '') as {fRow["Field"]}, ";
            }

            sql = sql.Substring(0, sql.Length - 2); //qutar ultima coma

            foreach (DataRow fRow in formTableInfo.Rows)
            {
                if (fRow["REFERENCED_TABLE_NAME"].ToString() != "masterDataRequests")
                {
                    sql += $", {fRow["REFERENCED_TABLE_SCHEMA"]}.{fRow["REFERENCED_TABLE_NAME"]}.code as {fRow["COLUMN_NAME"]}Code ";
                    joins += $"LEFT JOIN {fRow["REFERENCED_TABLE_SCHEMA"]}.{fRow["REFERENCED_TABLE_NAME"]} ON {fRow["TABLE_SCHEMA"]}.{generalTable}.{fRow["COLUMN_NAME"]} = {fRow["REFERENCED_TABLE_SCHEMA"]}.{fRow["REFERENCED_TABLE_NAME"]}.{fRow["REFERENCED_COLUMN_NAME"]} ";
                }
            }

            sql += $"FROM {generalTable} " + joins;
            sql += $" WHERE requestId = {id}";
            return sql;
        }
        /// <summary>
        /// metodo para extraer la data del formulario de una solicitud lineal
        /// </summary>
        /// <param name="formTable">la tabla a buscar</param>
        /// <param name="id">el requestId de la solicitud en masterDataRequest</param>
        /// <returns>query SELECT para ejecutar</returns>
        private string sqlQueryFormData(string formTable, string id)
        {
            DataTable formTableInfo = crud.Select($@"SELECT  k.TABLE_SCHEMA, k.COLUMN_NAME, k.REFERENCED_TABLE_NAME, k.REFERENCED_COLUMN_NAME, k.REFERENCED_TABLE_SCHEMA
                                                                        FROM information_schema.TABLE_CONSTRAINTS i 
                                                                        LEFT JOIN information_schema.KEY_COLUMN_USAGE k ON i.CONSTRAINT_NAME = k.CONSTRAINT_NAME 
                                                                        WHERE i.CONSTRAINT_TYPE = 'FOREIGN KEY' 
                                                                        AND i.TABLE_SCHEMA = 'masterData'
                                                                        AND i.TABLE_NAME = '{formTable}'  
                                                                        ORDER BY `k`.`REFERENCED_COLUMN_NAME` ASC", "masterData");

            DataTable formTableColumns = crud.Select($@"SHOW FULL COLUMNS FROM {formTable}", "masterData");
            string joins = "";
            string sql = $"SELECT ";

            foreach (DataRow fRow in formTableColumns.Rows)
            {
                sql += $"IFNULL({formTable}.{fRow["Field"]}, '') as {fRow["Field"]}, ";
            }

            sql = sql.Substring(0, sql.Length - 2); //qutar ultima coma

            foreach (DataRow fRow in formTableInfo.Rows)
            {
                if (fRow["REFERENCED_TABLE_NAME"].ToString() != "masterDataRequests" && !string.IsNullOrWhiteSpace(fRow["REFERENCED_TABLE_NAME"].ToString()))
                {
                    sql += $", IFNULL({fRow["REFERENCED_TABLE_SCHEMA"]}.{fRow["REFERENCED_TABLE_NAME"]}.{((fRow["REFERENCED_TABLE_NAME"].ToString() == "digital_sign") ? "UserID" : "code")}, '') as {fRow["COLUMN_NAME"]}Code ";
                    if (!joins.Contains(fRow["REFERENCED_TABLE_NAME"].ToString()))
                    {

                        joins += $"LEFT JOIN {fRow["REFERENCED_TABLE_SCHEMA"]}.{fRow["REFERENCED_TABLE_NAME"]} ON {fRow["TABLE_SCHEMA"]}.{formTable}.{fRow["COLUMN_NAME"]} = {fRow["REFERENCED_TABLE_SCHEMA"]}.{fRow["REFERENCED_TABLE_NAME"]}.{fRow["REFERENCED_COLUMN_NAME"]} ";
                    }
                }
            }

            sql += $" FROM {formTable} " + joins;
            sql += $" WHERE requestId = {id}";
            return sql;
        }

        /// <summary>
        /// Método para descargar archivos a el SFTP de SmartAndSimple
        /// </summary>
        /// <param name="filePathName"></param>
        /// <param name="enviroment"> "PRD" o "DEV"</param>
        /// <returns></returns>
        public bool DownloadFile(string filePathName)
        {
            try
            {
                string fileName = Path.GetFileName(filePathName);
                string pathfile = filePathName;
                return db2.DownloadFileSftp(filePathName, root.FilesDownloadPath + "\\" + fileName);


            }
            catch (Exception ex)
            {
                return false;
            }
        }

        //public void getDataBaseRequest()
        //{
        //    DataTable mat = new DataTable();
        //    DataTable rep = new DataTable();
        //    DataTable serv = new DataTable();
        //    DataTable matServ = new DataTable();
        //    DataTable equi = new DataTable();
        //    DataTable ven = new DataTable();
        //    DataTable cust = new DataTable();
        //    DataTable cont = new DataTable();
        //    //DataTable warr = new DataTable();
        //    DataTable ibase = new DataTable();

        //    //DataTable dt = crud.Select("Databot", "SELECT * FROM `gestiones_dm` WHERE ESTADO = 'APROBACION'", "automation");

        //    //AGREGAR LAS COLUMNAS DEL EXCEL
        //    #region columnas generales

        //    mat.Columns.Add("ID_GESTION");
        //    mat.Columns.Add("EMPLEADO");
        //    mat.Columns.Add("ESTADO");
        //    mat.Columns.Add("COMENTARIOS");
        //    mat.Columns.Add("COMENTARIOS_APROBADOR");
        //    mat.Columns.Add("RESPUESTA");
        //    mat.Columns.Add("TIPO_GESTION");
        //    mat.Columns.Add("TIPO_DATO");
        //    mat.Columns.Add("METODO");

        //    rep.Columns.Add("ID_GESTION");
        //    rep.Columns.Add("EMPLEADO");
        //    rep.Columns.Add("ESTADO");
        //    rep.Columns.Add("COMENTARIOS");
        //    rep.Columns.Add("COMENTARIOS_APROBADOR");
        //    rep.Columns.Add("RESPUESTA");
        //    rep.Columns.Add("TIPO_GESTION");
        //    rep.Columns.Add("TIPO_DATO");
        //    rep.Columns.Add("METODO");

        //    serv.Columns.Add("ID_GESTION");
        //    serv.Columns.Add("EMPLEADO");
        //    serv.Columns.Add("ESTADO");
        //    serv.Columns.Add("COMENTARIOS");
        //    serv.Columns.Add("COMENTARIOS_APROBADOR");
        //    serv.Columns.Add("RESPUESTA");
        //    serv.Columns.Add("TIPO_GESTION");
        //    serv.Columns.Add("TIPO_DATO");
        //    serv.Columns.Add("METODO");

        //    matServ.Columns.Add("ID_GESTION");
        //    matServ.Columns.Add("EMPLEADO");
        //    matServ.Columns.Add("ESTADO");
        //    matServ.Columns.Add("COMENTARIOS");
        //    matServ.Columns.Add("COMENTARIOS_APROBADOR");
        //    matServ.Columns.Add("RESPUESTA");
        //    matServ.Columns.Add("TIPO_GESTION");
        //    matServ.Columns.Add("TIPO_DATO");
        //    matServ.Columns.Add("METODO");

        //    equi.Columns.Add("ID_GESTION");
        //    equi.Columns.Add("EMPLEADO");
        //    equi.Columns.Add("ESTADO");
        //    equi.Columns.Add("COMENTARIOS");
        //    equi.Columns.Add("COMENTARIOS_APROBADOR");
        //    equi.Columns.Add("RESPUESTA");
        //    equi.Columns.Add("TIPO_GESTION");
        //    equi.Columns.Add("TIPO_DATO");
        //    equi.Columns.Add("METODO");

        //    ven.Columns.Add("ID_GESTION");
        //    ven.Columns.Add("EMPLEADO");
        //    ven.Columns.Add("ESTADO");
        //    ven.Columns.Add("COMENTARIOS");
        //    ven.Columns.Add("COMENTARIOS_APROBADOR");
        //    ven.Columns.Add("RESPUESTA");
        //    ven.Columns.Add("TIPO_GESTION");
        //    ven.Columns.Add("TIPO_DATO");
        //    ven.Columns.Add("METODO");

        //    cust.Columns.Add("ID_GESTION");
        //    cust.Columns.Add("EMPLEADO");
        //    cust.Columns.Add("ESTADO");
        //    cust.Columns.Add("COMENTARIOS");
        //    cust.Columns.Add("COMENTARIOS_APROBADOR");
        //    cust.Columns.Add("RESPUESTA");
        //    cust.Columns.Add("TIPO_GESTION");
        //    cust.Columns.Add("TIPO_DATO");
        //    cust.Columns.Add("METODO");

        //    cont.Columns.Add("ID_GESTION");
        //    cont.Columns.Add("EMPLEADO");
        //    cont.Columns.Add("ESTADO");
        //    cont.Columns.Add("COMENTARIOS");
        //    cont.Columns.Add("COMENTARIOS_APROBADOR");
        //    cont.Columns.Add("RESPUESTA");
        //    cont.Columns.Add("TIPO_GESTION");
        //    cont.Columns.Add("TIPO_DATO");
        //    cont.Columns.Add("METODO");

        //    ibase.Columns.Add("ID_GESTION");
        //    ibase.Columns.Add("EMPLEADO");
        //    ibase.Columns.Add("ESTADO");
        //    ibase.Columns.Add("COMENTARIOS");
        //    ibase.Columns.Add("COMENTARIOS_APROBADOR");
        //    ibase.Columns.Add("RESPUESTA");
        //    ibase.Columns.Add("TIPO_GESTION");
        //    ibase.Columns.Add("TIPO_DATO");
        //    ibase.Columns.Add("METODO");

        //    #endregion

        //    DataRow[] rowsMat = dt.Select("TIPO_DATO = 'MATERIALES'");
        //    DataRow[] rowsRep = dt.Select("TIPO_DATO = 'REPUESTOS'");
        //    DataRow[] rowsserv = dt.Select("TIPO_DATO = 'SERVICIOS'");
        //    DataRow[] rowsmatServ = dt.Select("TIPO_DATO = 'TERCEROS'");
        //    DataRow[] rowsequi = dt.Select("TIPO_DATO = 'EQUIPOS'");
        //    DataRow[] rowsven = dt.Select("TIPO_DATO = 'PROVEEDORES'");
        //    DataRow[] rowscust = dt.Select("TIPO_DATO = 'CLIENTES'");
        //    DataRow[] rowscont = dt.Select("TIPO_DATO = 'CONTACTOS'");
        //    DataRow[] rowsibase = dt.Select("TIPO_DATO = 'IBASE'");


        //    //JArray dgArrayCol = JArray.Parse(rowsMat[0]["DG"].ToString());
        //    //JObject dgCol = JObject.Parse(dgArrayCol[0].ToString());
        //    //Dictionary<string, string> dgJsonCol = dgCol.ToObject<Dictionary<string, string>>();

        //    //foreach (KeyValuePair<string, string> campo in dgJsonCol)
        //    //{
        //    //    string key = campo.Key;
        //    //    mat.Columns.Add(key);
        //    //}

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsMat[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        mat.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsRep[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        rep.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsserv[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        serv.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsmatServ[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        matServ.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsven[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        ven.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowscust[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        cust.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowscont[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        cont.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsibase[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        ibase.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsequi[0]["DG"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        equi.Columns.Add(key);
        //    }




        //    //JArray gestionesCol = JArray.Parse(dt.Rows[0]["GESTION"].ToString());
        //    //JObject filaCol = JObject.Parse(gestionesCol[0].ToString());
        //    //Dictionary<string, string> filaJsonCol = filaCol.ToObject<Dictionary<string, string>>();

        //    //foreach (KeyValuePair<string, string> campo in filaJsonCol)
        //    //{
        //    //    string key = campo.Key;
        //    //    mat.Columns.Add(key);
        //    //    rep.Columns.Add(key);
        //    //    serv.Columns.Add(key);
        //    //    matServ.Columns.Add(key);
        //    //    equi.Columns.Add(key);
        //    //    ven.Columns.Add(key);
        //    //    cust.Columns.Add(key);
        //    //    cont.Columns.Add(key);
        //    //    ibase.Columns.Add(key);
        //    //}


        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsMat[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        mat.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsRep[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        rep.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsserv[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        serv.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsmatServ[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        matServ.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsven[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        ven.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowscust[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        cust.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowscont[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        cont.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsibase[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        ibase.Columns.Add(key);
        //    }

        //    foreach (KeyValuePair<string, string> campo in JObject.Parse(JArray.Parse(rowsequi[0]["GESTION"].ToString())[0].ToString()).ToObject<Dictionary<string, string>>())
        //    {
        //        string key = campo.Key;
        //        equi.Columns.Add(key);
        //    }

        //    //AGREGAR LOS VALORES DEL EXCEL
        //    foreach (DataRow row in dt.Rows)
        //    {

        //        try
        //        {


        //            string ID_GESTION = row["ID_GESTION"].ToString();
        //            string EMPLEADO = row["EMPLEADO"].ToString();
        //            string FACTOR = row["FACTOR"].ToString();
        //            string ESTADO = row["ESTADO"].ToString();
        //            string COMENTARIOS = row["COMENTARIOS"].ToString();
        //            string COMENTARIOS_APROBADOR = row["COMENTARIOS_APROBADOR"].ToString();
        //            string RESPUESTA = row["RESPUESTA"].ToString();
        //            string TIPO_GESTION = row["TIPO_GESTION"].ToString();
        //            string TIPO_DATO = row["TIPO_DATO"].ToString();
        //            string METODO = row["METODO"].ToString();



        //            string generalData = row["DG"].ToString();
        //            string lineas = row["GESTION"].ToString();

        //            //[
        //            //"{\"FACTOR\":\"201010102\",
        //            //\"BAW\":\"\",
        //            //\"GESTION_BAW\":\"\"}"
        //            //]
        //            JArray dgArray = JArray.Parse(generalData);
        //            JObject dg = JObject.Parse(dgArray[0].ToString());
        //            Dictionary<string, string> dgJson = dg.ToObject<Dictionary<string, string>>();


        //            JArray gestiones = JArray.Parse(lineas);
        //            for (int i = 0; i < gestiones.Count; i++) //FOR POR CADA LINEA QUE TENGA LA SOLICITUD
        //            {
        //                DataRow rRow = null;

        //                switch (TIPO_DATO)
        //                {
        //                    case "MATERIALES":
        //                        rRow = mat.Rows.Add();
        //                        break;
        //                    case "PROVEEDORES":
        //                        rRow = ven.Rows.Add();
        //                        break;
        //                    case "REPUESTOS":
        //                        rRow = rep.Rows.Add();
        //                        break;
        //                    case "CLIENTES":
        //                        rRow = cust.Rows.Add();
        //                        break;
        //                    case "CONTACTOS":
        //                        rRow = cont.Rows.Add();
        //                        break;
        //                    case "SERVICIOS":
        //                        rRow = serv.Rows.Add();
        //                        break;
        //                    case "EQUIPOS":
        //                        rRow = equi.Rows.Add();
        //                        break;
        //                    case "IBASE":
        //                        rRow = ibase.Rows.Add();
        //                        break;
        //                    case "TERCEROS":
        //                        rRow = matServ.Rows.Add();
        //                        break;

        //                    default:
        //                        break;
        //                }


        //                rRow["ID_GESTION"] = ID_GESTION;
        //                rRow["EMPLEADO"] = EMPLEADO;
        //                rRow["FACTOR"] = FACTOR;
        //                rRow["ESTADO"] = ESTADO;
        //                rRow["COMENTARIOS"] = COMENTARIOS;
        //                rRow["COMENTARIOS_APROBADOR"] = COMENTARIOS_APROBADOR;
        //                rRow["RESPUESTA"] = RESPUESTA;
        //                rRow["TIPO_GESTION"] = TIPO_GESTION;
        //                rRow["TIPO_DATO"] = TIPO_DATO;
        //                rRow["METODO"] = METODO;

        //                foreach (KeyValuePair<string, string> campo in dgJson)
        //                {
        //                    string key = campo.Key;
        //                    string val = campo.Value;
        //                    rRow[key] = val;
        //                }

        //                JObject fila = JObject.Parse(gestiones[i].ToString());
        //                Dictionary<string, string> filaJson = fila.ToObject<Dictionary<string, string>>();

        //                foreach (KeyValuePair<string, string> campo in filaJson)
        //                {
        //                    string key = campo.Key;
        //                    string val = campo.Value;
        //                    rRow[key] = val;
        //                }

        //                //mat.AcceptChanges();

        //                switch (TIPO_DATO)
        //                {
        //                    case "MATERIALES":
        //                        mat.AcceptChanges();
        //                        break;
        //                    case "PROVEEDORES":
        //                        ven.AcceptChanges();
        //                        break;
        //                    case "REPUESTOS":
        //                        rep.AcceptChanges();
        //                        break;
        //                    case "CLIENTES":
        //                        cust.AcceptChanges();
        //                        break;
        //                    case "CONTACTOS":
        //                        cont.AcceptChanges();
        //                        break;
        //                    case "SERVICIOS":
        //                        serv.AcceptChanges();
        //                        break;
        //                    case "EQUIPOS":
        //                        equi.AcceptChanges();
        //                        break;
        //                    case "IBASE":
        //                        ibase.AcceptChanges();
        //                        break;
        //                    case "TERCEROS":
        //                        matServ.AcceptChanges();
        //                        break;

        //                    default:
        //                        break;
        //                }
        //            }

        //            //  [
        //            //  "{\"TIPO_MATERIAL\":\"ZHRW\",
        //            //  \"ID_MATERIAL\":\"21DJ00G1US\",
        //            //  \"SERIALIZABLE\":\"SI\",
        //            //  \"MODELO\":\"Lenovo ThinkBook 15 G4 IAP 21DJ00G1US 15.6\\\" Notebook - Full \",
        //            //  \"GM1\":\"01\",
        //            //  \"GM2\":\"\",
        //            //  \"GRUPO_POSICION\":\"NORM\",
        //            //  \"DESCRIPCION\":\"Lenovo ThinkBook 15 G4 IAP 21DJ00G1US 15.6\\\" Notebook - Full HD - 1920 x 1080 - Intel Core i5 12th Gen i5-1235U Deca-core (10 Core) 1.30 GHz - 8 GB Total RAM - 8 GB On-board Memory - 256 GB SSD - Miner\",
        //            //  \"PRECIO\":\"999999\",
        //            //  \"PROVEEDOR\":\"LENOVO\",
        //            //  \"GARANTIA\":\"WAR-01Y-G/5X9/24\"
        //            //  ,\"COMENTARIOSG\":\"Lenovo ThinkBook 15 G4 IAP 21DJ00G1US 15.6\\\" Notebook - Full HD - 1920 x 1080 - Intel Core i5 12th Gen i5-1235U Deca-core (10 Core) 1.30 GHz - 8 GB Total RAM - 8 GB On-board Memory - 256 GB SSD - Mineral Gray - Intel Chip - Windows 11 - Intel Iris Xe Graph\",
        //            //  \"LLAVE\":\"21DJ00G1US\"}
        //            //  "]
        //        }
        //        catch (Exception ex)
        //        {
        //            console.WriteLine(ex.Message);
        //        }
        //    }
        //    MsExcel ms = new MsExcel();


        //    string attachmentmat = "gestionesMateriales_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentrep = "gestionesRepuestos_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentserv = "gestionesServicios_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentmatServ = "gestionesMaterialesDeServicio_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentequi = "gestionesEquipos_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentven = "gestionesProveedores_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentcust = "gestionesClientes_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentcont = "gestionesContactos_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //    string attachmentibase = "gestionesIbase_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";

        //    ms.CreateExcel(mat, "Datos", root.FilesDownloadPath + "\\" + attachmentmat, true);
        //    ms.CreateExcel(rep, "Datos", root.FilesDownloadPath + "\\" + attachmentrep, true);
        //    ms.CreateExcel(serv, "Datos", root.FilesDownloadPath + "\\" + attachmentserv, true);
        //    ms.CreateExcel(matServ, "Datos", root.FilesDownloadPath + "\\" + attachmentmatServ, true);
        //    ms.CreateExcel(equi, "Datos", root.FilesDownloadPath + "\\" + attachmentequi, true);
        //    ms.CreateExcel(ven, "Datos", root.FilesDownloadPath + "\\" + attachmentven, true);
        //    ms.CreateExcel(cust, "Datos", root.FilesDownloadPath + "\\" + attachmentcust, true);
        //    ms.CreateExcel(cont, "Datos", root.FilesDownloadPath + "\\" + attachmentcont, true);
        //    ms.CreateExcel(ibase, "Datos", root.FilesDownloadPath + "\\" + attachmentibase, true);
        //}
    
    }
}
