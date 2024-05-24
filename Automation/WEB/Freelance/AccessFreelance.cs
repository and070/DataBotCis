using DataBotV5.Data.Root;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBotV5.Data.Database;
using DataBotV5.Security;
using DataBotV5.Data.Projects.Freelance;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.Freelance;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;
using System.IO;
using DataBotV5.Logical.MicrosoftTools;
using System.Security;
using System.Net;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using System.Data;
using SAP.Middleware.Connector;
using System.Security.Cryptography;
using System.Globalization;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Text.RegularExpressions;

namespace DataBotV5.Automation.WEB.Freelance
{
    public class AccessFreelance
    {
        Rooting root = new Rooting();
        ProcessInteraction process = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        CRUD crud = new CRUD();
        FreelanceSqlSS freelanceSql = new FreelanceSqlSS();
        FreelanceFiSS freelanceFi = new FreelanceFiSS();
        string sapSystem = "ERP";
        string ssMandante = "QAS";
        int mandante = 260;
        Credentials cred = new Credentials();
        SecureAccess sec = new SecureAccess();
        Database db = new Database();
        MailInteraction mail = new MailInteraction();
        WebexTeams webex = new WebexTeams();
        MsExcel excel = new MsExcel();
        SharePoint sharepoint = new SharePoint();
        public void Main()
        {

            string query = $@"SELECT 
                  alf.*,
                  CONCAT(alf.firstName, ' ', alf.lastName) as fullName,
                  als.statusText,
                  als.statusType,
                  alt.altaType as altaTypeText,
                  cc.code as companyCodeCode,
                  cc.name as companyCodeName,
                  cc.nrRangeNr as nrRangeNr,
                  eesg.description as eeSubGroupText,
                  eesg.code as eeSubGroupCode,
                  sapc.countryName as countryName,
                  sapc.countryCode,
                  nat.nationalityName,
                  suba.subAreaName,
                  suba.subAreaCode,
                  idt.code as idTypeCode
                   FROM freelance_db.altasFreelance as alf
                   INNER JOIN freelance_db.altasStatus as als ON als.id = alf.status
                   INNER JOIN freelance_db.altasType as alt ON alt.id = alf.type
                   INNER JOIN databot_db.companyCode as cc ON cc.id = alf.companyCode
                   INNER JOIN hcm_hiring_db.eeSubGroup as eesg ON eesg.id = alf.eeSubGroup
                   INNER JOIN databot_db.sapCountries as sapc ON sapc.id = alf.country
                   INNER JOIN hcm_hiring_db.Nationality as nat ON nat.nationalityCode = alf.nationality
                   INNER JOIN hcm_hiring_db.SubArea as suba ON suba.id = alf.subArea
                   INNER JOIN hcm_hiring_db.idType as idt ON idt.id = alf.idType
                    WHERE alf.status = 3
                   ORDER BY alf.id DESC";

            DataTable dt = crud.Select(query, "freelance_db");
            if (dt.Rows.Count > 0)
            {
                altaProcess(dt);
            }

            #region Verifica si ya alguna solicitud se le completo email y usuario de SAP
            string queryFinal = $@"SELECT 
                  alf.*,
                  CONCAT(alf.lastName, ', ', alf.firstName) as fullName,
                  als.statusText,
                  als.statusType,
                  alt.altaType as altaTypeText,
                  cc.code as companyCodeCode,
                  cc.name as companyCodeName,
                  cc.nrRangeNr as nrRangeNr,
                  eesg.description as eeSubGroupText,
                  eesg.code as eeSubGroupCode,
                  sapc.countryName as countryName,
                  sapc.countryCode,
                  nat.nationalityName,
                  suba.subAreaName,
                  suba.subAreaCode,
                  idt.code as idTypeCode,
                  itr.sapUser as sapUserBpm,
                  itr.userId as emailExtGbm,
                  itr.isFinished as isFinishedBpm,
                  itr.password as tempPass
                   FROM freelance_db.altasFreelance as alf
                   INNER JOIN freelance_db.altasStatus as als ON als.id = alf.status
                   INNER JOIN freelance_db.altasType as alt ON alt.id = alf.type
                   INNER JOIN databot_db.companyCode as cc ON cc.id = alf.companyCode
                   INNER JOIN hcm_hiring_db.eeSubGroup as eesg ON eesg.id = alf.eeSubGroup
                   INNER JOIN databot_db.sapCountries as sapc ON sapc.id = alf.country
                   INNER JOIN hcm_hiring_db.Nationality as nat ON nat.nationalityCode = alf.nationality
                   INNER JOIN hcm_hiring_db.SubArea as suba ON suba.id = alf.subArea
                   INNER JOIN hcm_hiring_db.idType as idt ON idt.id = alf.idType
                   LEFT JOIN it_request.requests as itr ON itr.id = alf.idBpm
                    WHERE alf.status = 4 AND itr.isFinished = 1
                   ORDER BY alf.id DESC";

            DataTable dtFinal = crud.Select(queryFinal, "freelance_db");
            if (dtFinal.Rows.Count > 0)
            {
                finalizarAlta(dtFinal);
            }
            #endregion

            #region Verifica si al usuario le agregaron el supplier
            string queryVendor = $@"SELECT 
                  alf.*,
                  CONCAT(alf.firstName, ' ', alf.lastName) as fullName,
                  als.statusText,
                  als.statusType,
                  alt.altaType as altaTypeText,
                  cc.code as companyCodeCode,
                  cc.nrRangeNr as nrRangeNr,
                  eesg.description as eeSubGroupText,
                  eesg.code as eeSubGroupCode,
                  sapc.countryName as countryName,
                  sapc.countryCode,
                  nat.nationalityName,
                  suba.subAreaName,
                  suba.subAreaCode,
                  idt.code as idTypeCode,
                  itr.sapUser as sapUserBpm,
                  itr.userId as emailExtGbm,
                  itr.isFinished as isFinishedBpm
                   FROM freelance_db.altasFreelance as alf
                   INNER JOIN freelance_db.altasStatus as als ON als.id = alf.status
                   INNER JOIN freelance_db.altasType as alt ON alt.id = alf.type
                   INNER JOIN databot_db.companyCode as cc ON cc.id = alf.companyCode
                   INNER JOIN hcm_hiring_db.eeSubGroup as eesg ON eesg.id = alf.eeSubGroup
                   INNER JOIN databot_db.sapCountries as sapc ON sapc.id = alf.country
                   INNER JOIN hcm_hiring_db.Nationality as nat ON nat.nationalityCode = alf.nationality
                   INNER JOIN hcm_hiring_db.SubArea as suba ON suba.id = alf.subArea
                   INNER JOIN hcm_hiring_db.idType as idt ON idt.id = alf.idType
                   LEFT JOIN it_request.requests as itr ON itr.id = alf.idBpm
                    WHERE alf.addSupplier = 1 AND alf.status != 7
                   ORDER BY alf.id DESC";

            DataTable dtVendor = crud.Select(queryVendor, "freelance_db");
            if (dtVendor.Rows.Count > 0)
            {
                addSuppierToSap(dtVendor);
            }

            #endregion
        }
        private void altaProcess(DataTable info)
        {
            string altaCoordinatorQuery = @"SELECT
                                                ds.email
                                                FROM SS_Access_Permissions.UserAccess as ua
                                                INNER JOIN MIS.digital_sign as ds ON ds.id = ua.fk_SignID
                                                INNER JOIN SS_Access_Permissions.Permissions as pe ON pe.id = ua.fk_Permissions
                                                WHERE pe.name = 'Freelance Coordinator Altas'";
            DataTable dtAlta = crud.Select(altaCoordinatorQuery, "SS_Access_Permissions");
            string[] cc = dtAlta.AsEnumerable().Select(row => row.Field<string>("email")).ToArray();
            cc = cc.Concat(new[] { "dmeza@gbm.net" }).ToArray();
            string htmlpage = Properties.Resources.emailtemplate1;
            foreach (DataRow item in info.Rows)
            {

                string fullName = item["fullName"].ToString().ToUpper();
                string solicitante = item["createdBy"].ToString() + "@gbm.net";
                try
                {

                    #region extraer información
                    string hiredDate = DateTime.Parse(item["hireDate"].ToString()).ToString("dd.MM.yyyy");
                    string birthDate = DateTime.Parse(item["birthDate"].ToString()).ToString("yyyy-MM-dd");
                    string hireDate = DateTime.Parse(item["hireDate"].ToString()).ToString("yyyy-MM-dd");
                    string contractEndDate = DateTime.Parse(item["contractEndDate"].ToString()).ToString("yyyy-MM-dd");
                    string organizationalUnit = "70009902"; //CONSULTING SERVICES
                    string companyCode = item["companyCodeCode"].ToString();
                    string subArea = item["subAreaCode"].ToString();
                    string eGroup = "E";
                    string eeSubGroupCode = item["eeSubGroupCode"].ToString();
                    string pais = "";
                    string sociedad = "";
                    if (companyCode.Substring(0, 2) == "GB" || companyCode.Substring(0, 2) == "LC")
                    {
                        pais = companyCode.Substring(2, 2);
                        sociedad = companyCode.Substring(2, 2);
                    }
                    else
                    {
                        pais = companyCode.Substring(0, 2);
                        sociedad = companyCode.Substring(0, 2);
                    }
                    if (sociedad == "MD")
                    {
                        sociedad = "MI";
                    }
                    string personalArea = sociedad + "CO"; //Consulting
                    #endregion
                    string positionId = "";
                    string respuesta = "";
                    string employeeId = "";
                    DataTable itReturn = new DataTable();
                    #region Crear Posición del Freelance

                    console.WriteLine("Crear Posición del Freelance");
                    Dictionary<string, string> parameters = new Dictionary<string, string>();
                    parameters["NOMBRE_DE_LA_POSICION"] = "CONSULTOR FREELANCE";
                    parameters["FECHA_DE_INGRESO"] = hiredDate;
                    parameters["ORGANIZATIONAL_UNIT"] = organizationalUnit;
                    parameters["IGNORE_ORGUNIT_Z"] = "X";
                    //parameters["NAME_OF_MANAGER"] = ""; //TIENE QUE VENIR VACIO
                    //parameters["COST_CENTER"] = ""; //TIENE QUE VENIR VACIO
                    //parameters["JOB"] = ""; //TIENE QUE VENIR VACIO
                    parameters["COMPANY_CODE"] = companyCode;
                    parameters["PERSONNEL_AREA"] = personalArea;
                    parameters["PERSONNEL_SUBAREA"] = subArea;
                    //parameters["ADMIN_PRODUCT"] = ""; //002 (servicios) TIENE QUE VENIR VACIO adminProduct;
                    //parameters["DIRECCION"] = ""; //007 (Services) TIENE QUE VENIR VACIO direction; 
                    //parameters["EPM"] = ""; //002 (NO) TIENE QUE VENIR VACIO
                    //parameters["FIJO"] = ""; //002 (COSTO) TIENE QUE VENIR VACIO
                    //parameters["GERENCIA"] = "";//002 (NO) TIENE QUE VENIR VACIO;
                    //parameters["HEADCOUNT"] = "";//001 (presupuestado)  TIENE QUE VENIR VACIO;
                    //parameters["LINEA_DE_NEGOCIO"] = ""; // 001 (consulting) TIENE QUE VENIR VACIO
                    //parameters["LOCAL_REGIONAL_PLA"] = ""; //002 (regional) TIENE QUE VENIR VACIO
                    //parameters["LOCAL_REGIONAL_SALARIAL"] = ""; //002 (regional) TIENE QUE VENIR VACIO
                    //parameters["PAGO_FIJO"] = ""; //020 (100%) TIENE QUE VENIR VACIO
                    //parameters["PAGO_VARIABLE"] = ""; //000 (0%) TIENE QUE VENIR VACIO
                    //parameters["PRODUCTIVIDAD"] = ""; //002 NO TIENE QUE VENIR VACIO
                    //parameters["PROTECCION"] = ""; //002 NO TIENE QUE VENIR VACIO
                    //parameters["RECURSO_DE_INVERSION"] = ""; //002 NO TIENE QUE VENIR VACIO
                    //parameters["VARIABLE"] = ""; //002 (COSTO) TIENE QUE VENIR VACIO
                    parameters["EMPLOYEE_GROUP"] = eGroup;
                    parameters["EE_SUBGROUP"] = eeSubGroupCode;
                    //parameters["PUESTO_CCSS"] = ""; //TIENE QUE VENIR VACIO
                    //parameters["PUESTO_INS"] = ""; //TIENE QUE VENIR VACIO
                    parameters["IGNORE_INFO_GT"] = "X";
                    parameters["VACANTY"] = "2"; //FILLED

                    IRfcFunction func = sap.ExecuteRFC(sapSystem, "ZHR_POSITION_CREATE", parameters, mandante);
                    respuesta = func.GetValue("RESPUESTA").ToString();
                    if (respuesta == "NA JOB")
                    {
                        respuesta = "No se encontro el Job seleccionado";
                    }
                    else if (respuesta == "NA Unidad Z")
                    { respuesta = "No se encontro la Unidad Z homologa de la unidad: " + organizationalUnit; }
                    else if (respuesta.Contains("Error:"))
                    { respuesta = func.GetValue("RESPUESTA").ToString(); }
                    else if (respuesta == "OK")
                    {
                        respuesta = "Posicion creada con exito";
                        positionId = func.GetValue("ID_POSICION").ToString();
                    }
                    else
                    {
                        respuesta = "Error insesperado, contacte a Application Management";
                        positionId = func.GetValue("ID_POSICION").ToString();
                    }

                    if (positionId == "")
                    {
                        //error al crear posicion 
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al crear la posición del Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al crear la posición del Freelance {fullName}", new string[] {"dmeza@gbm.net"}, null);
                        //error status
                        crud.Update($"UPDATE altasFreelance SET status = 7, botResponse = '{respuesta}' WHERE id = {item["id"]}", "freelance_db");

                        continue;
                    }
                    string upQuery = $"UPDATE altasFreelance SET positionId = '{positionId}' WHERE id = {item["id"]}";
                    bool upPosition = crud.Update(upQuery, "freelance_db");
                    if (!upPosition)
                    {
                        //error al actualizar posicion 
                        //notificar al solicitante y coordinadores de altas
                        //el proceso continua porque igual puede crear el empleado
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar la posición {positionId} del Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar la posición {positionId} del Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);
                        //error status
                        crud.Update($"UPDATE altasFreelance SET status = 7, botResponse = '{respuesta}' WHERE id = {item["id"]}", "freelance_db");

                    }
                    #endregion

                    #region Crear número de colaborador
                    RfcDestination destination = sap.GetDestRFC(sapSystem, mandante);
                    RfcRepository repo = destination.Repository;
                    IRfcFunction func2 = repo.CreateFunction("ZHR_EMPLOYEE_CREATE");
                    IRfcTable general = func2.GetTable("EMPLOYEE_INFO");


                    general.Append();

                    general.SetValue("ZPOSITION", positionId);
                    general.SetValue("ACTION_TYPE", "ZA");
                    general.SetValue("REASON_ACTION", "01");
                    general.SetValue("COMPANY_CODE", companyCode);
                    general.SetValue("NR_RANGE", item["nrRangeNr"]);
                    general.SetValue("PERSONNEL_AREA", personalArea);
                    general.SetValue("SUB_AREA", subArea);
                    general.SetValue("EMPLOYEE_GROUP", eGroup);
                    general.SetValue("EMPLOYEE_SUBGROUP", eeSubGroupCode);
                    general.SetValue("ORG_UNIT", organizationalUnit);
                    general.SetValue("PAYROLL_AREA", "99");
                    general.SetValue("PAYMENT_METHOD", "C");
                    //general.SetValue("BANK_KEY", item[""]);
                    //general.SetValue("BANK_ACCOUNT", item[""]);
                    general.SetValue("WORK_SCHEDULE_RULE", "PAG DIR");
                    general.SetValue("FIRST_NAME", item["firstName"].ToString().ToUpper());
                    general.SetValue("MIDDLE_NAME", item["middleName"].ToString() == "" || item["middleName"].ToString() == "NULL" ? "" : item["middleName"].ToString().ToUpper());
                    general.SetValue("LAST_NAME", item["lastName"].ToString().ToUpper());
                    general.SetValue("BIRTH_DATE", birthDate);
                    general.SetValue("LANGUAGE", "EN");
                    general.SetValue("ADDRESS", item["address"].ToString().ToUpper());
                    general.SetValue("POSTAL_CODE", "99999");
                    general.SetValue("COUNTRY", item["countryCode"]);
                    //general.SetValue("REGION", item[""]);
                    general.SetValue("NATIONALITY", item["nationality"]);
                    general.SetValue("FORM_OF_ADDRESS", "1");
                    general.SetValue("GENDER", "1");
                    general.SetValue("ID_TYPE", item["idTypeCode"]);
                    general.SetValue("IDENTIFICATION", item["identification"]);
                    //general.SetValue("EMAIL", item[""]);
                    //general.SetValue("TELEPHONE", item[""]);
                    general.SetValue("HIRED_DATE", hireDate);
                    general.SetValue("CONTRACT_END_DATE", contractEndDate);
                    general.SetValue("SUPPLIER", item["supplier"].ToString() == "" || item["supplier"].ToString() == "NULL" ? "" : item["supplier"].ToString().PadLeft(10, '0'));
                    //general.SetValue("SAPUSERNAME", item[""]);
                    //general.SetValue("GBMEMAIL", item[""]);
                    general.SetValue("ACTIVITY_NUMBER", "SAP_CONSULTORIA");
                    //general.SetValue("CECO", item[""]);
                    //general.SetValue("ACTIVITY_TYPE", item[""]);

                    //func.SetValue("PROFILE", profile);

                    func2.Invoke(destination);


                    respuesta = func2.GetValue("RESPONSE").ToString();
                    employeeId = func2.GetValue("EMPLOYEE_ID").ToString();
                    itReturn = sap.GetDataTableFromRFCTable(func2.GetTable("RETURN"));

                    if (itReturn.Rows.Count > 0)
                    {
                        StringBuilder botResponseBuilder = new StringBuilder();
                        if (respuesta != "OK")
                        {
                            botResponseBuilder.Append(respuesta);
                        }

                        foreach (DataRow row in itReturn.Rows)
                        {
                            botResponseBuilder.Append($", {row["MESSAGE"]}");
                        }

                        string botResponse = botResponseBuilder.ToString();
                        crud.Update($"UPDATE altasFreelance SET botResponse = '{botResponse}' WHERE id = {item["id"]}", "freelance_db");
                    }

                    if (respuesta != "OK")
                    {
                        //error al crear empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al crear el id de empleado del Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al crear el id de empleado del Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                        crud.Update($"UPDATE altasFreelance SET status = 7 WHERE id = {item["id"]}", "freelance_db");
                        continue;
                    }

                    string upQueryEmployee = $"UPDATE altasFreelance SET idSap = '{employeeId}' WHERE id = {item["id"]}";
                    bool upEmployee = crud.Update(upQueryEmployee, "freelance_db");
                    if (!upEmployee)
                    {
                        //error al actualizar empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar el id {employeeId} del colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar el id {employeeId} del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);
                        //error status
                        crud.Update($"UPDATE altasFreelance SET status = 7 WHERE id = {item["id"]}", "freelance_db");

                    }

                    #endregion

                    #region Enviar la nueva Alta de Freelance en S&S
                    string subAreaName = item["subAreaName"].ToString();
                    string normalizedString = subAreaName.Normalize(NormalizationForm.FormD);
                    string subAreaNameFinal = new string(normalizedString
                        .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                        .ToArray());

                    //buscar country y location para IT_REQUEST
                    string companyCodeName = item["companyCodeName"].ToString();
                    DataTable dataTable = crud.Select($"SELECT id FROM countries WHERE country = '{companyCodeName}';", "it_request");
                    string countryIdTi = "";
                    if (dataTable.Rows.Count > 0)
                    {
                        countryIdTi = dataTable.Rows[0]["id"].ToString();
                    }
                    else
                    {
                        switch (companyCodeName)
                        {
                            case "Lat. Cap. Venezuela":
                                countryIdTi = "8"; //venezuela
                                break;
                            default:
                                countryIdTi = "14"; //otro
                                break;
                        }
                    }

                    DataTable dataTable2 = crud.Select($"SELECT id FROM locations WHERE location = '{subAreaNameFinal}' AND fkIdCountry = '{countryIdTi}';", "it_request");
                    string locationIdTi = "";
                    if (dataTable2.Rows.Count > 0)
                    {
                        locationIdTi = dataTable2.Rows[0]["id"].ToString();
                    }
                    else
                    {
                        locationIdTi = "20"; //desconocido
                    }



                    string insertQueryBpm = $@"INSERT INTO
                                            `requests` (`firstName`,
                                                        `lastName`, 
                                                        `personalArea`,
                                                        `country`, 
                                                        `location`, 
                                                        `jobTypeFreelance`,     
                                                        `employeeId`, 
                                                        `phoneType`, 
                                                        `genDataComments`,
                                                        `hcApprover`,
                                                        `initialDate`,
                                                        `requestState`,
                                                        `lastStep`, 
                                                        `isNewColab`,
                                                        `isFinished`, 
                                                        `isFreelance`,
                                                        `pendingMail`, 
                                                        `endDate`,
                                                        `createdBy`)          
                                            VALUES ('{item["firstName"]}', 
                                                    '{item["lastName"]}', 
                                                    '1',
                                                    '{countryIdTi}',
                                                    '{locationIdTi}',
                                                    '1',
                                                    '{employeeId}',
                                                    'NO',
                                                    'Por favor crear: NUEVO
                                                    Correo corporativo
                                                    Usuario SAP
                                                    Accesos Microsoft teams
                                                    enviar notificación a {item["createdBy"]}@gbm.net',
                                                    '{item["createdBy"]}',
                                                    '{DateTime.Now:yyyy-MM-dd}',
                                                    'Alta',
                                                    '15',
                                                    '0',
                                                    '0',
                                                    '1',
                                                    '1',
                                                    '{contractEndDate}',
                                                    '{item["createdBy"]}')";

                    long idBpm = crud.NonQueryAndGetId(insertQueryBpm, "it_request");
                    if (idBpm == 0)
                    {
                        //error al enviar solicitud a BPM
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al envíar la solicitud de alta del colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", insertQueryBpm);
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al envíar la solicitud de alta del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);
                        //error status
                        crud.Update($"UPDATE altasFreelance SET status = 7 WHERE id = {item["id"]}", "freelance_db");

                        continue;
                    }
                    //actualizar el id en la solicitud
                    string upQueryIdBpm = $"UPDATE altasFreelance SET idBpm = {idBpm} WHERE id = {item["id"]}";
                    bool upIdBpm = crud.Update(upQueryIdBpm, "freelance_db");
                    if (!upIdBpm)
                    {
                        //error al actualizar empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar el id de Altas MIS {idBpm} del colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar el id de Altas MIS {idBpm} del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    }
                    #endregion

                    #region Modificar el ID de la solicitud a 4 - En Proceso BPM
                    string upQueryStatus = $"UPDATE altasFreelance SET status = 4 WHERE id = {item["id"]}";
                    bool upStatus = crud.Update(upQueryStatus, "freelance_db");
                    if (!upStatus)
                    {
                        //error al actualizar empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar el status En Proceso Altas MIS del colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar el status En Proceso Altas MIS del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    }
                    #endregion

                    #region Notificar al solicitante y HCM
                    //estimado se le notifica que ya se creo la posicion y el id del colaborador freelance, al mismo tiempo, se levanto la solicitud de Alta 
                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Proceso en Espera Alta MIS del colaborador Freelance {fullName}");
                    htmlpage = htmlpage.Replace("{cuerpo}", $"Estimado(a) se le notifica que ya se creo la posición {positionId} del colaborador Freelance {fullName} con el id del colaborador {employeeId}, se encuentra en espera que se haga el Alta por parte del equipo de MIS.");
                    htmlpage = htmlpage.Replace("{contenido}", "Se le notificará la finalización del proceso de alta.");
                    mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Proceso en Espera de Altas MIS de Alta del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    #endregion

                    #region notificar a los colaboradores de MIS sobre el alta
                    //lo hace el robot de TI Solicitudes
                    #endregion


                }
                catch (Exception ex)
                {
                    //error inesperado
                    //notificar al solicitante y coordinadores de altas
                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Error insesperado del colaborador Freelance {fullName}");
                    htmlpage = htmlpage.Replace("{cuerpo}", ex.Message);
                    htmlpage = htmlpage.Replace("{contenido}", ex.ToString());
                    mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error insesperado del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    crud.Update($"UPDATE altasFreelance SET status = 7, botResponse = '{fullName + ": " + ex.Message}' WHERE id = {item["id"]}", "freelance_db");

                }
            }
        }

        private void finalizarAlta(DataTable info)
        {
            string altaCoordinatorQuery = @"SELECT
                                                ds.email
                                                FROM SS_Access_Permissions.UserAccess as ua
                                                INNER JOIN MIS.digital_sign as ds ON ds.id = ua.fk_SignID
                                                INNER JOIN SS_Access_Permissions.Permissions as pe ON pe.id = ua.fk_Permissions
                                                WHERE pe.name = 'Freelance Coordinator Altas'";
            DataTable dtAlta = crud.Select(altaCoordinatorQuery, "SS_Access_Permissions");
            string[] cc = dtAlta.AsEnumerable().Select(row => row.Field<string>("email")).ToArray();
            cc = cc.Concat(new[] { "dmeza@gbm.net" }).ToArray();

            string htmlpage = Properties.Resources.emailtemplate1;

            foreach (DataRow row in info.Rows)
            {
                string fullName = row["fullName"].ToString().ToUpper();
                string solicitante = row["createdBy"].ToString() + "@gbm.net";
                string employeeId = row["idSap"].ToString();
                try
                {
                    #region Buscar la información de id de solicitud de BPM SS
                    bool isFinishedBpm = bool.Parse(row["isFinishedBpm"].ToString());
                    if (!isFinishedBpm)
                    {
                        //no hace nada xq todavia no ha finalizado
                        continue;
                    }
                    string sapUserBpm = row["sapUserBpm"].ToString().ToString().ToUpper();
                    string emailExtGbm = row["emailExtGbm"].ToString().ToUpper();
                    if (!Regex.IsMatch(emailExtGbm, @"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"))
                    {
                        //emailExtGbm No es email
                        if (emailExtGbm.Contains("@"))
                        {
                            //formato "example@ex"
                            emailExtGbm = emailExtGbm.Split('@')[0] + "@EXT.GBM.NET";
                        }
                        else
                        {
                            //formato "example"
                            emailExtGbm += "@EXT.GBM.NET";
                        }
                        
                    }
                    string personalEmail = row["personalEmail"].ToString().ToUpper();
                    string eeSubGroupCode = row["eeSubGroupCode"].ToString();
                    string tempPass = row["tempPass"].ToString();
                    string positionId = row["positionId"].ToString();
                    string hiredDate = DateTime.Parse(row["hireDate"].ToString()).ToString("dd.MM.yyyy");
                    string birthDate = DateTime.Parse(row["birthDate"].ToString()).ToString("yyyy-MM-dd");
                    string hireDate = DateTime.Parse(row["hireDate"].ToString()).ToString("yyyy-MM-dd");
                    string contractEndDate = DateTime.Parse(row["contractEndDate"].ToString()).ToString("yyyy-MM-dd");
                    string organizationalUnit = "70009902"; //CONSULTING SERVICES
                    string companyCode = row["companyCodeCode"].ToString();
                    string subArea = row["subAreaCode"].ToString();
                    string eGroup = "E";
                    string pais = "";
                    string sociedad = "";
                    if (companyCode.Substring(0, 2) == "GB" || companyCode.Substring(0, 2) == "LC")
                    {
                        pais = companyCode.Substring(2, 2);
                        sociedad = companyCode.Substring(2, 2);
                    }
                    else
                    {
                        pais = companyCode.Substring(0, 2);
                        sociedad = companyCode.Substring(0, 2);
                    }
                    if (sociedad == "MD")
                    {
                        sociedad = "MI";
                    }
                    string personalArea = sociedad + "CO"; //Consulting
                    #endregion

                    #region ejecutar la FM para actualizar el email y usuario en 0105
                    RfcDestination destination = sap.GetDestRFC(sapSystem, mandante);
                    RfcRepository repo = destination.Repository;
                    IRfcFunction func2 = repo.CreateFunction("ZHR_EMPLOYEE_CREATE");
                    IRfcTable general = func2.GetTable("EMPLOYEE_INFO");
                    func2.SetValue("EMPLOYEE_NUMBER", employeeId);

                    general.Append();

                    general.SetValue("ZPOSITION", positionId);
                    general.SetValue("ACTION_TYPE", "ZA");
                    general.SetValue("REASON_ACTION", "01");
                    general.SetValue("COMPANY_CODE", companyCode);
                    general.SetValue("NR_RANGE", row["nrRangeNr"]);
                    general.SetValue("PERSONNEL_AREA", personalArea);
                    general.SetValue("SUB_AREA", subArea);
                    general.SetValue("EMPLOYEE_GROUP", eGroup);
                    general.SetValue("EMPLOYEE_SUBGROUP", eeSubGroupCode);
                    general.SetValue("ORG_UNIT", organizationalUnit);
                    general.SetValue("PAYROLL_AREA", "99");
                    general.SetValue("PAYMENT_METHOD", "C");
                    //general.SetValue("BANK_KEY", item[""]);
                    //general.SetValue("BANK_ACCOUNT", item[""]);
                    general.SetValue("WORK_SCHEDULE_RULE", "PAG DIR");
                    general.SetValue("FIRST_NAME", row["firstName"].ToString().ToUpper());
                    general.SetValue("MIDDLE_NAME", row["middleName"].ToString() == "" || row["middleName"].ToString() == "NULL" ? "" : row["middleName"].ToString().ToUpper());
                    general.SetValue("LAST_NAME", row["lastName"].ToString().ToUpper());
                    general.SetValue("BIRTH_DATE", birthDate);
                    general.SetValue("LANGUAGE", "EN");
                    general.SetValue("ADDRESS", row["address"].ToString().ToUpper());
                    general.SetValue("POSTAL_CODE", "99999");
                    general.SetValue("COUNTRY", row["countryCode"]);
                    //general.SetValue("REGION", item[""]);
                    general.SetValue("NATIONALITY", row["nationality"]);
                    general.SetValue("FORM_OF_ADDRESS", "1");
                    general.SetValue("GENDER", "1");
                    general.SetValue("ID_TYPE", row["idTypeCode"]);
                    general.SetValue("IDENTIFICATION", row["identification"]);
                    //general.SetValue("EMAIL", item[""]);
                    //general.SetValue("TELEPHONE", item[""]);
                    general.SetValue("HIRED_DATE", hireDate);
                    general.SetValue("CONTRACT_END_DATE", contractEndDate);
                    general.SetValue("SUPPLIER", row["supplier"].ToString() == "" || row["supplier"].ToString() == "NULL" ? "" : row["supplier"].ToString().PadLeft(10, '0'));
                    //*************************************************************************************
                    general.SetValue("SAPUSERNAME", sapUserBpm);
                    general.SetValue("GBMEMAIL", emailExtGbm);
                    //*************************************************************************************
                    general.SetValue("ACTIVITY_NUMBER", "SAP_CONSULTORIA");
                    //general.SetValue("CECO", item[""]);
                    //general.SetValue("ACTIVITY_TYPE", item[""]);

                    //func.SetValue("PROFILE", profile);

                    func2.Invoke(destination);


                    string respuesta = func2.GetValue("RESPONSE").ToString();
                    employeeId = func2.GetValue("EMPLOYEE_ID").ToString();
                    DataTable itReturn = new DataTable();
                    itReturn = sap.GetDataTableFromRFCTable(func2.GetTable("RETURN"));
                    string botResponse = "";
                    if (itReturn.Rows.Count > 0)
                    {
                        StringBuilder botResponseBuilder = new StringBuilder();
                        if (respuesta != "OK")
                        {
                            botResponseBuilder.Append(respuesta);
                        }

                        foreach (DataRow rowReturn in itReturn.Rows)
                        {
                            botResponseBuilder.Append($", {rowReturn["MESSAGE"]}");
                        }

                        botResponse = botResponseBuilder.ToString();
                        crud.Update($"UPDATE altasFreelance SET botResponse = '{botResponse}' WHERE id = {row["id"]}", "freelance_db");
                    }

                    if (respuesta != "OK" || botResponse.Contains("P0105"))
                    {
                        //error al crear empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar el usuario y email del Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", botResponse);
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar el usuario y email del Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                        crud.Update($"UPDATE altasFreelance SET status = 7 WHERE id = {row["id"]}", "freelance_db");
                        continue;
                    }

                    #endregion

                    #region actualizar los emails y usuario en la DB //cambiar status
                    string upQuery = $"UPDATE altasFreelance SET gbmEmail = '{emailExtGbm}', sapUserName = '{sapUserBpm}', status = 5 WHERE id = {row["id"]}";
                    bool up = crud.Update(upQuery, "freelance_db");
                    if (!up)
                    {
                        //error al actualizar empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar la información de usuario y email del colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", upQuery);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar la información de usuario y email del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);
                    }

                    #endregion

                    #region brindarle acceso al portal externo de Freelance
                    string pass = GenerateRandomPassword(18);
                    string encodePass = ConvertToBase64(pass);
                    string externQuery = $@"INSERT INTO freelance_db.accessFreelance (`type`, `vendor`, `email`, `user`, `pass`,`active`,`createdBy`)
                                                                               VALUE ('{(eeSubGroupCode == "ZR" ? 1 : 3)}','{fullName}','{emailExtGbm}','{sapUserBpm}','{encodePass}','1','{solicitante}')";
                    bool ins = crud.Insert(externQuery, "freelance_db");
                    if (!ins)
                    {
                        //error al brindarle acceso al portal externo
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al crear el acceso en el portal externo del colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", upQuery);
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al crear el acceso en el portal externo del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);
                    }
                    #endregion

                    #region enviar notificación al freelance
                    string body = "<table class='myCustomTable' width='100 %'>";
                    body += "<thead><tr><th>User</th><th>Pass Portal</th><th>Pass Email Temporal</th></tr></thead>";
                    body += "<tbody>";
                    body += $"<tr><td>{emailExtGbm}</td><td>{pass}</td><td>{tempPass}</td></tr>"; //+ "<br><br>"
                    body += "</tbody>";
                    body += "</table>";
                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Acceso al portal de proveedores de GBM");
                    htmlpage = htmlpage.Replace("{cuerpo}", $"Estimado(a) {fullName}, se le notifica que ya posee acceso al <a href=\"https://proveedores.gbm.net\" target=\"_blank\">portal de proveedores de GBM</a> haga click en el enlace anterior e ingrese con sus credenciales: ");
                    htmlpage = htmlpage.Replace("{contenido}", body);
                    string[] cc2 = new string[] { solicitante, "dmeza@gbm.net" };
                    List<string> adj = new List<string>();
                    string[] adjuntos = null;
                    string fileName1 = "Ingreso correo GBM Externo 1.pdf";
                    string fileName2 = "Consulting - Reporte de Horas proveedores Consulting.pdf";
                    string adj1 = root.FilesDownloadPath + "\\" + fileName1;
                    string adj2 = root.FilesDownloadPath + "\\" + fileName2;
                    bool validate = sharepoint.DownloadFileFromSharePoint("https://gbmcorp.sharepoint.com/sites/FreelancePortal/", "Documentos", fileName1);
                    if (validate)
                    {
                        adj.Add(adj1);
                    }
                    bool validate2 = sharepoint.DownloadFileFromSharePoint("https://gbmcorp.sharepoint.com/sites/FreelancePortal/", "Documentos", fileName2);
                    if (validate2)
                    {
                        adj.Add(adj2);
                    }
                    if (adj.Count > 0)
                    {
                        adjuntos = adj.ToArray();
                    }

                    mail.SendHTMLMail(htmlpage, new string[] { personalEmail }, $"Acceso al portal de proveedores de GBM", cc2, adjuntos);
                    #endregion

                    #region contestar la solicitud al solicitante y HCM

                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Colaborador Externo Freelance {fullName} Creado En SAP con Éxito");
                    htmlpage = htmlpage.Replace("{cuerpo}", $"Estimado(a) se le notifica que ya se realizó el alta del colaborador Freelance {fullName} con el id del colaborador {employeeId}, el mismo ya tiene acceso al portal de proveedores https://proveedores.gbm.net");
                    htmlpage = htmlpage.Replace("{contenido}", "");
                    mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Finalización de alta del colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net", "hogonzalez@gbm.net" }, null);

                    #endregion
                }
                catch (Exception ex)
                {
                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Error insesperado al finalizar el colaborador Freelance {fullName}");
                    htmlpage = htmlpage.Replace("{cuerpo}", ex.Message);
                    htmlpage = htmlpage.Replace("{contenido}", ex.ToString());
                    mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error insesperado al finalizar el colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    crud.Update($"UPDATE altasFreelance SET status = 7, botResponse = '{fullName + ": " + ex.Message}' WHERE id = {row["id"]}", "freelance_db");

                }
            }
        }

        private void addSuppierToSap(DataTable info)
        {
            string altaCoordinatorQuery = @"SELECT
                                                ds.email
                                                FROM SS_Access_Permissions.UserAccess as ua
                                                INNER JOIN MIS.digital_sign as ds ON ds.id = ua.fk_SignID
                                                INNER JOIN SS_Access_Permissions.Permissions as pe ON pe.id = ua.fk_Permissions
                                                WHERE pe.name = 'Freelance Coordinator Altas'";
            DataTable dtAlta = crud.Select(altaCoordinatorQuery, "SS_Access_Permissions");
            string[] cc = dtAlta.AsEnumerable().Select(row => row.Field<string>("email")).ToArray();
            cc = cc.Concat(new[] { "dmeza@gbm.net" }).ToArray();
            string htmlpage = Properties.Resources.emailtemplate1;

            foreach (DataRow row in info.Rows)
            {
                string employeeId = row["idSap"].ToString();
                string fullName = row["fullName"].ToString().ToUpper();
                string solicitante = row["createdBy"].ToString() + "@gbm.net";
                try
                {
                    string supplier = row["supplier"].ToString().PadLeft(10, '0');
                    string eeSubGroupCode = row["eeSubGroupCode"].ToString();
                    string positionId = row["positionId"].ToString();
                    string hiredDate = DateTime.Parse(row["hireDate"].ToString()).ToString("dd.MM.yyyy");
                    string birthDate = DateTime.Parse(row["birthDate"].ToString()).ToString("yyyy-MM-dd");
                    string hireDate = DateTime.Parse(row["hireDate"].ToString()).ToString("yyyy-MM-dd");
                    string contractEndDate = DateTime.Parse(row["contractEndDate"].ToString()).ToString("yyyy-MM-dd");
                    string organizationalUnit = "70009902"; //CONSULTING SERVICES
                    string companyCode = row["companyCodeCode"].ToString();
                    string subArea = row["subAreaCode"].ToString();
                    string sapUserBpm = row["sapUserName"].ToString().ToUpper();
                    string emailExtGbm = row["gbmEmail"].ToString().ToUpper();
                    string eGroup = "E";
                    string pais = "";
                    string sociedad = "";
                    if (companyCode.Substring(0, 2) == "GB" || companyCode.Substring(0, 2) == "LC")
                    {
                        pais = companyCode.Substring(2, 2);
                        sociedad = companyCode.Substring(2, 2);
                    }
                    else
                    {
                        pais = companyCode.Substring(0, 2);
                        sociedad = companyCode.Substring(0, 2);
                    }
                    if (sociedad == "MD")
                    {
                        sociedad = "MI";
                    }
                    string personalArea = sociedad + "CO"; //Consulting

                    #region verificar que el proveedor exista en SAP (utilizar FM para leer tabla LFA1)
                    bool existVendor = true;
                    RfcDestination destErp = new SapVariants().GetDestRFC("ERP");
                    IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
                    fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
                    fmMg.SetValue("QUERY_TABLE", "LFA1");
                    fmMg.SetValue("DELIMITER", "");

                    IRfcTable fields = fmMg.GetTable("FIELDS");
                    fields.Append();
                    fields.SetValue("FIELDNAME", "LIFNR");

                    IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
                    fmOptions.Append();
                    fmOptions.SetValue("TEXT", $"LIFNR = '{supplier}'");

                    fmMg.Invoke(destErp);

                    DataTable tableSap = sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA"));
                    if (tableSap.Rows.Count == 0)
                    {
                        existVendor = false;
                        //notificar al solicitante
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error el proveedor no existe para el colaborador Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", $"Estimado(a), se le notifica que el proveedor {supplier} no existe en SAP, por favor verifique e ingrese al portal de Proveedores para agregar el ID correcto.");
                        htmlpage = htmlpage.Replace("{contenido}", "");
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error el proveedor no existe para el colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);
                        string upQuerys = $"UPDATE altasFreelance SET addSupplier = 0 WHERE id = {row["id"]}";
                        bool upS = crud.Update(upQuerys, "freelance_db");
                        continue;
                    }


                    #endregion
                    //actualizar el proveedor en SAP
                    #region ejecutar la FM para actualizar el email y usuario en 0105
                    RfcDestination destination = sap.GetDestRFC(sapSystem, mandante);
                    RfcRepository repo = destination.Repository;
                    IRfcFunction func2 = repo.CreateFunction("ZHR_EMPLOYEE_CREATE");
                    IRfcTable general = func2.GetTable("EMPLOYEE_INFO");
                    func2.SetValue("EMPLOYEE_NUMBER", employeeId);

                    general.Append();

                    general.SetValue("ZPOSITION", positionId);
                    general.SetValue("ACTION_TYPE", "ZA");
                    general.SetValue("REASON_ACTION", "01");
                    general.SetValue("COMPANY_CODE", companyCode);
                    general.SetValue("NR_RANGE", row["nrRangeNr"]);
                    general.SetValue("PERSONNEL_AREA", personalArea);
                    general.SetValue("SUB_AREA", subArea);
                    general.SetValue("EMPLOYEE_GROUP", eGroup);
                    general.SetValue("EMPLOYEE_SUBGROUP", eeSubGroupCode);
                    general.SetValue("ORG_UNIT", organizationalUnit);
                    general.SetValue("PAYROLL_AREA", "99");
                    general.SetValue("PAYMENT_METHOD", "C");
                    //general.SetValue("BANK_KEY", item[""]);
                    //general.SetValue("BANK_ACCOUNT", item[""]);
                    general.SetValue("WORK_SCHEDULE_RULE", "PAG DIR");
                    general.SetValue("FIRST_NAME", row["firstName"].ToString().ToUpper());
                    general.SetValue("MIDDLE_NAME", row["middleName"].ToString() == "" || row["middleName"].ToString() == "NULL" ? "" : row["middleName"].ToString().ToUpper());
                    general.SetValue("LAST_NAME", row["lastName"].ToString().ToUpper());
                    general.SetValue("BIRTH_DATE", birthDate);
                    general.SetValue("LANGUAGE", "EN");
                    general.SetValue("ADDRESS", row["address"].ToString().ToUpper());
                    general.SetValue("POSTAL_CODE", "99999");
                    general.SetValue("COUNTRY", row["countryCode"]);
                    //general.SetValue("REGION", item[""]);
                    general.SetValue("NATIONALITY", row["nationality"]);
                    general.SetValue("FORM_OF_ADDRESS", "1");
                    general.SetValue("GENDER", "1");
                    general.SetValue("ID_TYPE", row["idTypeCode"]);
                    general.SetValue("IDENTIFICATION", row["identification"]);
                    //general.SetValue("EMAIL", item[""]);
                    //general.SetValue("TELEPHONE", item[""]);
                    general.SetValue("HIRED_DATE", hireDate);
                    general.SetValue("CONTRACT_END_DATE", contractEndDate);

                    //*************************************************************************************
                    general.SetValue("SUPPLIER", supplier == "" || supplier == "NULL" ? "" : supplier);
                    //*************************************************************************************
                    if (sapUserBpm != "")
                    {
                        general.SetValue("SAPUSERNAME", sapUserBpm);
                    }
                    if (emailExtGbm != "")
                    {
                        general.SetValue("GBMEMAIL", emailExtGbm);
                    }
                    general.SetValue("ACTIVITY_NUMBER", "SAP_CONSULTORIA");
                    //general.SetValue("CECO", item[""]);
                    //general.SetValue("ACTIVITY_TYPE", item[""]);

                    //func.SetValue("PROFILE", profile);

                    func2.Invoke(destination);


                    string respuesta = func2.GetValue("RESPONSE").ToString();
                    employeeId = func2.GetValue("EMPLOYEE_ID").ToString();
                    DataTable itReturn = new DataTable();
                    itReturn = sap.GetDataTableFromRFCTable(func2.GetTable("RETURN"));
                    string botResponse = "";
                    if (itReturn.Rows.Count > 0)
                    {
                        StringBuilder botResponseBuilder = new StringBuilder();
                        if (respuesta != "OK")
                        {
                            botResponseBuilder.Append(respuesta);
                        }

                        foreach (DataRow rowReturn in itReturn.Rows)
                        {
                            botResponseBuilder.Append($", {rowReturn["MESSAGE"]}");
                        }

                        botResponse = botResponseBuilder.ToString();
                        crud.Update($"UPDATE altasFreelance SET botResponse = '{botResponse}' WHERE id = {row["id"]}", "freelance_db");
                    }

                    if (respuesta != "OK" || botResponse.Contains("P0315"))
                    {
                        //error al crear empleado
                        //notificar al solicitante y coordinadores de altas
                        htmlpage = Properties.Resources.emailtemplate1;
                        htmlpage = htmlpage.Replace("{subject}", $"Error al actualizar el usuario y email del Freelance {fullName}");
                        htmlpage = htmlpage.Replace("{cuerpo}", respuesta);
                        htmlpage = htmlpage.Replace("{contenido}", botResponse);
                        mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error al actualizar el usuario y email del Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                        crud.Update($"UPDATE altasFreelance SET status = 7 WHERE id = {row["id"]}", "freelance_db");
                        continue;
                    }

                    #endregion



                    //actualizar el campo addSupplier a 0
                    string upQueryStatus = $"UPDATE altasFreelance SET addSupplier = 0 WHERE id = {row["id"]}";
                    bool upStatus = crud.Update(upQueryStatus, "freelance_db");
                    if (!upStatus)
                    {
                        //error al actualizar empleado
                        //notificar al solicitante y coordinadores de altas
                        continue;
                    }

                    #region contestar la solicitud al solicitante y HCM

                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Proveedor agregado con éxito al Freelance {fullName}");
                    htmlpage = htmlpage.Replace("{cuerpo}", $"Estimado(a) se le notifica que ya se agregó el proveedor {supplier} al Freelance {fullName} con el id del colaborador {employeeId}");
                    htmlpage = htmlpage.Replace("{contenido}", "");
                    mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Proveedor Agregado con éxito al colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    #endregion


                }
                catch (Exception ex)
                {
                    htmlpage = Properties.Resources.emailtemplate1;
                    htmlpage = htmlpage.Replace("{subject}", $"Error insesperado al finalizar el colaborador Freelance {fullName}");
                    htmlpage = htmlpage.Replace("{cuerpo}", ex.Message);
                    htmlpage = htmlpage.Replace("{contenido}", ex.ToString());
                    mail.SendHTMLMail(htmlpage, new string[] { solicitante }, $"Error insesperado al finalizar el colaborador Freelance {fullName}", new string[] { "dmeza@gbm.net" }, null);

                    crud.Update($"UPDATE altasFreelance SET status = 7, botResponse = '{fullName + ": " + ex.Message}' WHERE id = {row["id"]}", "freelance_db");

                }

            }
        }

        static string GenerateRandomPassword(int length)
        {
            const string chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz$%#";
            byte[] data = new byte[length];

            using (RNGCryptoServiceProvider crypto = new RNGCryptoServiceProvider())
            {
                crypto.GetBytes(data);
            }

            StringBuilder result = new StringBuilder(length);
            foreach (byte b in data)
            {
                result.Append(chars[b % chars.Length]);
            }

            return result.ToString();
        }
        static string ConvertToBase64(string input)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(bytes);
        }
    }
}
