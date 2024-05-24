using System;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Projects.HcmHiring;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using System.Data;
using DataBotV5.Logical.MicrosoftTools;
using System.Web;
using ClosedXML.Excel;
using System.IO;

namespace DataBotV5.Automation.WEB.HumanCapital
{
    /// <summary>
    /// Clase WEB Automation encargada de la gestión de PB10 de Human Capital.
    /// </summary>
    class PB10WebSmartAndSimple
    {
        public string response = "";
        public bool failure = false;
        public string response_failure = "";
        Stats estadisticas = new Stats();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        SqlPb10 sqlpb10 = new SqlPb10();
        WebexTeams wt = new WebexTeams();
        
        #region Variables Globales	

        string filas;
        string carpeta1;
        string carpeta2;
        string email;
        string fecha;
        string mes = "", ano = "", dia = "";
        string MES;
        string DIA;
        string ASP;
        string address;
        string insti;
        string name_company;
        string contact_company;
        string country_company;
        string first_name;
        string Tipo_Puesto;
        string second_name;
        string first_lastname;
        string e_mail;
        string second_lastname;
        string validacion;
        string nacionalidad;
        string pais_vacante;
        string estado_civil;
        string Person_Area;
        string segundo_idioma;
        string subarea;
        string nativo;
        string nombre;
        string apellido;
        string segundo_nombre;
        string nombre_completo;
        int largo;
        string tip_vacante;
        string numposition;
        string genero;
        string cedula;
        string PersonArea;
        string tip_plaza;
        string respuesta = "";
        string respuesta2 = "";
        string TelfCasa;
        string TelfCelu;
        string DispViajar;
        string DispReubi;
        string DispIngreso;
        string Formacion;
        string InstiFormacion;
        string GradoAcad;
        string fecha_nacimiento;
        string fecha_educacion;
        string TituloObt;
        string SegIdiomaDominio;
        string FechaStartExp;
        string FechaEndExp;
        string NameCompany;
        string PuestoCompany;
        string ContactTelf;
        string DebugFlag;
        DateTime fecha_start_exp;
        DateTime fecha_end_exp;
        int dif_fecha;
        string fecha_vuelta;

        string mandante = "ERP";

        string respFinal = "";

        #endregion

        public void Main()
        {
            Dictionary<string, string> sol_info = new Dictionary<string, string>();
            sol_info = sqlpb10.newRequest();
            if (sol_info.Count > 0)
            {
                console.WriteLine(DateTime.Now + " > > > " + "Procesando...");
                PB10Processing(sol_info);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void PB10Processing(Dictionary<string, string> info)
        {
            string id = info["id"].ToString();
            try
            {
                string jobTitle = val.RemoveSpecialChars(info["jobTitle"].ToString().Trim().ToUpper(), 1);
                string subject_cand = "Solicitud de Contratación en GBM - " + jobTitle;
                string createdBy = info["createdBy"].ToString().ToLower() + "@gbm.net";
                string emailCandidate = val.RemoveSpecialChars(info["emailCandidate"].ToString(), 1);
                string status = val.RemoveSpecialChars(info["status"].ToString(), 1);
                string esubject = "Error en la solicitud de PB10: " + id;
                string respuesta = "";


                if (emailCandidate == "")
                {
                    //notificacion a HCA de error
                    //modificar el status a rechazado
                    sqlpb10.ChangeStatePb10SmartAndSimple(id, "Email vacio", "ER", DateTime.Now);

                    wt.SendNotification(createdBy, "Solicitud rechazada: " + id, "**Notificacion de Rechazado de PB10:** Estimado(a) se le notifica que la solicitud **#" + id + "** ha sido rechazada");
                    return;
                }


                console.WriteLine(DateTime.Now + " > > > " + "Estado En Proceso, procesando...");
                #region Variables Privadas
                int rows;
                string mensaje_devolucion = "";
                string validar_strc;
                bool validar_lineas = true;
                var valor = "";

                string id_candidato = "";
                #endregion

                #region Extraer info de json
                DataTable personalData = sqlpb10.getCandidateInfo("CandidatePersonalData", id);

                //Extrae información de la tabla hcm_hiring_db.CandidatePersonalData en DB Smart And Simple
                string identification = val.RemoveSpecialChars(personalData.Rows[0]["identification"].ToString().Trim(), 1);
                string firstName = val.RemoveSpecialChars(personalData.Rows[0]["firstName"].ToString().Trim().ToUpper(), 1);
                string secondName = val.RemoveSpecialChars(personalData.Rows[0]["secondName"].ToString().Trim().ToUpper(), 1);
                string firstSurname = val.RemoveSpecialChars(personalData.Rows[0]["firstSurname"].ToString().Trim().ToUpper(), 1);
                string secondSurname = val.RemoveSpecialChars(personalData.Rows[0]["secondSurname"].ToString().Trim().ToUpper(), 1);
                string civilStatus = personalData.Rows[0]["civilStatus"].ToString().Trim();
                string gender = personalData.Rows[0]["gender"].ToString().Trim();
                string nationality = personalData.Rows[0]["nationality"].ToString().Trim();
                string birthDate = personalData.Rows[0]["birthDate"].ToString().Trim();
                string address = val.RemoveSpecialChars(personalData.Rows[0]["address"].ToString().Trim().ToUpper(), 1);
                string telephone = personalData.Rows[0]["telephone"].ToString().Trim();
                string cellPhone = personalData.Rows[0]["cellPhone"].ToString().Trim();
                string dispTravel = personalData.Rows[0]["dispTravel"].ToString().Trim();
                string dispRelocation = personalData.Rows[0]["dispRelocation"].ToString().Trim();
                string dispEntry = personalData.Rows[0]["dispEntry"].ToString().Trim();
                string wageAspiration = personalData.Rows[0]["wageAspiration"].ToString().Trim();


                //Extrae información de la tabla hcm_hiring_db.CandidateEducation en DB Smart And Simple

                DataTable candidateDataEducation = sqlpb10.getCandidateInfo("CandidateEducation", id);

                string academicTraining = candidateDataEducation.Rows[0]["academicTraining"].ToString().Trim();
                string educationalInstitution = candidateDataEducation.Rows[0]["educationalInstitution"].ToString().Trim();
                string academicDegree = candidateDataEducation.Rows[0]["academicDegreeCode"].ToString().Trim();
                string institutionName = val.RemoveSpecialChars(candidateDataEducation.Rows[0]["institutionName"].ToString().Trim().ToUpper(), 1);
                string degree = val.RemoveSpecialChars(candidateDataEducation.Rows[0]["degree"].ToString().Trim().ToUpper(), 1);
                string graduationDate = candidateDataEducation.Rows[0]["graduationDate"].ToString().Trim();
                string nativeLang = candidateDataEducation.Rows[0]["nativeLang"].ToString().Trim();
                string secondLang = candidateDataEducation.Rows[0]["secondLang"].ToString().Trim();
                string domainLevel = candidateDataEducation.Rows[0]["domainLevel"].ToString().Trim();

                //Extrae información de la tabla hcm_hiring_db.CandidateExperience en DB Smart And Simple

                DataTable CandidateDataExperience = sqlpb10.getCandidateInfo("CandidateExperience", id);
                string experienceStartDate = "";
                string experienceFinDate = "";
                string companyName = "";
                string companyCountry = "";
                string companyJob = "";
                string companyContact = "";
                string contactPhone = "";
                try
                {
                    experienceStartDate = CandidateDataExperience.Rows[0]["experienceStartDate"].ToString().Trim();
                    experienceFinDate = CandidateDataExperience.Rows[0]["experienceFinDate"].ToString().Trim();
                    companyName = val.RemoveSpecialChars(CandidateDataExperience.Rows[0]["companyName"].ToString().Trim().ToUpper(), 1);
                    companyCountry = CandidateDataExperience.Rows[0]["companyCountry"].ToString().Trim();
                    companyJob = CandidateDataExperience.Rows[0]["companyJob"].ToString().Trim();
                    companyContact = val.RemoveSpecialChars(CandidateDataExperience.Rows[0]["companyContact"].ToString().Trim().ToUpper(), 1);
                    contactPhone = CandidateDataExperience.Rows[0]["contactPhone"].ToString().Trim();
                }
                catch (Exception)
                {
                }


                #endregion

                #region extraer datos generales //Vienen de la tabla PB10Request
                string position = val.RemoveSpecialChars(info["position"].ToString().Trim(), 1);

                string positionType = info["positionType"].ToString();
                string vacancyType = info["vacancyType"].ToString();
                string plazaType = info["plazaType"].ToString();
                string country = info["country"].ToString();
                string personalArea = info["personalArea"].ToString();
                string subArea = info["subArea"].ToString();
                #endregion

                #region validacion de datos
                country = (country == "DR") ? "DO" : country;
                //country = (country == "HQ") ? "CR" : country;
                _ = identification.Replace("´", "");
                string nombre_completo = $"{firstName} {firstSurname}";
                degree = (degree.Length > 100) ? degree.Substring(0, 100) : degree;
                institutionName = (institutionName.Length > 80) ? institutionName.Substring(0, 80) : institutionName;
                companyName = (companyName.Length > 60) ? companyName.Substring(0, 60) : companyName;
                companyContact = (companyContact.Length > 40) ? companyContact.Substring(0, 40) : companyContact;

                DateTime idate = DateTime.MinValue;
                if (experienceStartDate == "")
                {
                    experienceStartDate = DateTime.Now.ToString("yyyy-MM-dd");
                    idate = DateTime.Parse(DateTime.Now.ToString("dd/MM/yyyy"));
                }
                else
                {
                    idate = DateTime.Parse(experienceStartDate);
                }

                DateTime fdate = DateTime.MinValue;
                if (experienceFinDate == "")
                {
                    experienceFinDate = DateTime.Now.ToString("yyyy-MM-dd");
                    fdate = DateTime.Parse(DateTime.Now.ToString("dd/MM/yyyy"));
                }
                else
                {
                    fdate = DateTime.Parse(experienceFinDate);
                }

                dif_fecha = DateTime.Compare(idate, fdate);

                if (dif_fecha > 0)
                {
                    string tempdate;
                    tempdate = experienceStartDate;
                    experienceStartDate = experienceFinDate;
                    experienceFinDate = tempdate;
                }

                address = val.RemoveSpecialChars(address, 1);
                if (address.Length > 60)
                { address = address.Substring(0, 60).ToUpper(); }
                address = address.Replace("\n", " ");
                address = address.Replace("\r", " ");
                if (telephone != "")
                {
                    if (telephone.Length > 14)
                    { telephone = telephone.Substring(0, 14).Trim(); }
                }

                if (cellPhone != "")
                {
                    if (cellPhone.Length > 20)
                    { cellPhone = cellPhone.Substring(0, 20).Trim(); ; }
                }
                institutionName = val.RemoveSpecialChars(institutionName, 1);
                companyName = (companyName != "") ? val.RemoveSpecialChars(companyName, 1) : companyName;
                companyContact = (companyContact != "") ? val.RemoveSpecialChars(companyContact, 1) : "NA";
                birthDate = DateTime.Parse(birthDate).ToString("yyyy-MM-dd");
                graduationDate = DateTime.Parse(graduationDate).ToString("yyyy-MM-dd");
                experienceStartDate = DateTime.Parse(experienceStartDate).ToString("yyyy-MM-dd");
                experienceFinDate = DateTime.Parse(experienceFinDate).ToString("yyyy-MM-dd");
                string debugflag = "";
                #endregion
                #region SAP
                console.WriteLine(DateTime.Now + " > > > " + "Corriendo RFC de SAP: " + root.BDProcess);
                try
                {
                    #region Parametros de SAP
                    Dictionary<string, string> parametros = new Dictionary<string, string>
                    {
                        ["NUMERO_POSITION"] = position,
                        ["TIPO_VACANTE"] = vacancyType,
                        ["TIPO_PLAZA"] = plazaType,
                        ["TIPO_PUESTO"] = positionType,
                        ["COUNTRY_APPLICANT"] = country,
                        ["PERSON_AREA"] = personalArea,
                        ["SUB_AREA"] = subArea,
                        ["EMAIL"] = emailCandidate,
                        ["CEDULA"] = identification,
                        ["FIRSTNAME"] = firstName,
                        ["SECONDNAME"] = secondName,
                        ["LASTNAME"] = firstSurname,
                        ["LASTNAME_2"] = secondSurname,
                        ["ESTADO_CIVIL"] = civilStatus,
                        ["GENERO"] = gender,
                        ["COUNTRY1"] = nationality,
                        ["BIRTHDATE"] = birthDate,
                        ["ADDRESS"] = address,
                        ["TELF_CASA"] = telephone,
                        ["TELF_CELU"] = cellPhone,
                        ["DISP_VIAJAR"] = dispTravel,
                        ["DISP_REUBI"] = dispRelocation,
                        ["DISP_INGRESO"] = dispEntry,
                        ["ASPIRACION"] = wageAspiration,
                        ["FORMACION"] = academicTraining,
                        ["INSTI_FORMACION"] = educationalInstitution,
                        ["GRADO_ACAD"] = academicDegree,
                        ["INSTITUTO"] = institutionName,
                        ["TITULO_OBT"] = degree,
                        ["EDUCACION_STARTDATE"] = graduationDate,
                        ["IDIOMA_NATIVO"] = nativeLang,
                        ["SEGUNDO_IDIOMA"] = secondLang,
                        ["SEG_IDIOMA_DOMINIO"] = domainLevel,
                        ["FECHA_START_EXP"] = experienceStartDate,
                        ["FECHA_END_EXP"] = experienceFinDate,
                        ["NAME_COMPANY"] = companyName,
                        ["COUNTRY_COMPANY"] = companyCountry,
                        ["PUESTO_COMPANY"] = companyJob,
                        ["CONTACT_COMPANY"] = companyContact,
                        ["CONTACT_TELF"] = contactPhone,
                        ["DEBUG_FLAG"] = debugflag
                    };


                    IRfcFunction func = sap.ExecuteRFC(mandante, "ZHR_PB10", parametros);
                    #endregion

                    #region Procesar Salidas del FM
                    id_candidato = func.GetValue("NUMBER").ToString();
                    respuesta = func.GetValue("RESULTADO").ToString();
                    console.WriteLine(DateTime.Now + " > > > " + nombre_completo + " " + func.GetValue("NUMBER").ToString() + "<br>" + func.GetValue("RESULTADO").ToString());

                    if (respuesta.Contains("Error"))
                    { validar_lineas = false; }

                    #endregion
                }
                catch (Exception ex)
                {
                    response_failure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 1);
                    console.WriteLine(DateTime.Now + " > > > " + " Finishing process " + response_failure);
                    respuesta = respuesta + nombre_completo + ": " + ex.ToString() + ex.StackTrace;
                    validar_lineas = false;

                }

                #endregion
                log.LogDeCambios("Creacion", root.BDProcess, createdBy, "Crear Candidato PB10", respuesta, id);
                respFinal = respFinal + "\\n" + respuesta;

                console.WriteLine(DateTime.Now + " > > > " + "Respondiendo solicitud");
                if (validar_lineas == false)
                {
                    //enviar email de repuesta de error
                    string[] cc = { "dmeza@gbm.net", "appmanagement@gbm.net" };
                    sqlpb10.ChangeStatePb10SmartAndSimple(id, respuesta, "ER", DateTime.Now);
                    mail.SendHTMLMail(respuesta, new string[] { createdBy }, esubject, cc);
                }
                else
                {
                    //modificar el estado a finalizado
                    sqlpb10.ChangeStatePb10SmartAndSimple(id, respuesta, "FI", DateTime.Now);
                    //enviar email de repuesta de exito
                    wt.SendNotification(createdBy, "Solicitud finalizada: " + id, "**Notificacion de finalización de PB10:** Estimado(a) se le notifica que la solicitud **#" + id + "** ha finalizado, con el siguiente resultado: <br><br> " + respuesta);

                }
                root.BDUserCreatedBy = createdBy;
                root.requestDetails = respFinal;
            }
            catch (Exception ex)
            {
                //modificar el status a error
                //enviar error a datos maestros
                sqlpb10.ChangeStatePb10SmartAndSimple(id, ex.Message, "ER", DateTime.Now);
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Gestion: " + id + "<br>" + ex.Message, new string[] {"appmanagement@gbm.net"}, "Error PB10 en la solicitud: " + id, cc);
            }

        }


        #region robotLocal
        public void migrarDB()
        {
            bool valLines = true;
            string ruta = @"C:\Users\dmeza\Desktop\Mis Documentos\RFC\Nuevo Sistema PB10\DB data.xlsx";
            DataSet excelBook = new MsExcel().GetExcelBook(ruta);
            DataTable excel = excelBook.Tables["Sheet1"];



            foreach (DataRow row in excel.Rows)
            {
                string id = "";
                string candidateEmail = row["candidateEmail"].ToString();
                string candidateTime = row["candidateTime"].ToString();
                string user = row["user"].ToString();
                try
                {
                    id = row["id"].ToString();
                    if (id != "")
                    {
                        string CandidatePersonalData = row["datos_personales"].ToString();
                        JObject jPData = JObject.Parse(CandidatePersonalData);
                        jPData["id"] = id;
                        jPData["createdBy"] = candidateEmail;
                        jPData["createdAt"] = candidateTime;
                        string CandidatePersonalDataQuery = queryInsert(jPData, "CandidatePersonalData");

                        string CandidateEducation = row["educacion"].ToString();
                        JObject jEData = JObject.Parse(CandidateEducation);
                        jEData["id"] = id;
                        jEData["createdBy"] = candidateEmail;
                        jEData["createdAt"] = candidateTime;
                        string CandidateEducationQuery = queryInsert(jEData, "CandidateEducation");

                        string CandidateExperience = row["experiencia"].ToString();
                        JObject jXData = JObject.Parse(CandidateExperience);
                        jXData["id"] = id;
                        jXData["createdBy"] = candidateEmail;
                        jXData["createdAt"] = candidateTime;
                        string CandidateExperienceQuery = queryInsert(jXData, "CandidateExperience");


                        string adjuntos = row["adjuntos"].ToString();
                        string[] aadj = adjuntos.Split(',');
                        string adjuntosQuery = "INSERT INTO `UploadFiles` (`name`, `candidateId`, `user`, `codification`, `type`, `path`, `active`, `createdAt`, `createdBy`, `updatedAt`, `updatedBy`) VALUES ";
                        string aQuery = "";

                        foreach (string item in aadj)
                        {
                            string fileName = item.Replace("[\"", "");
                            fileName = fileName.Replace("\"]", "");
                            fileName = fileName.Replace("\"", "");
                            //fileName = fileName.Replace("[", "");
                            //fileName = fileName.Replace("]", "");
                            string mimeType = MimeMapping.GetMimeMapping(fileName);
                            aQuery = aQuery + $"('{fileName}', {id}, '{user}', '7bit', '{mimeType}', '/home/appmanager/projects/recruitment-form/src/files/{id}/{fileName}', 1, '{candidateTime}', '{user}', NULL, NULL),";
                        }
                        adjuntosQuery = adjuntosQuery + aQuery;
                        adjuntosQuery = adjuntosQuery.Remove(adjuntosQuery.Length - 1);
                        adjuntosQuery = adjuntosQuery + ";";
                        row["datos_personales_query"] = CandidatePersonalDataQuery;
                        row["educacion_query"] = CandidateEducationQuery;
                        row["experiencia_query"] = CandidateExperienceQuery;
                        row["adjuntos_query"] = adjuntosQuery;

                        excel.AcceptChanges();
                    }


                }
                catch (Exception ex)
                {
                    valLines = false;
                    response = ex.Message;
                }

            }

            console.WriteLine(DateTime.Now + " > > > " + "Save Excel...");
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(excel, "Resultados");
            ruta = root.FilesDownloadPath + $"\\Resultados migracion DB {DateTime.Now.ToString("yyyyMMddHHmmssffff")}.xlsx";
            if (File.Exists(ruta))
            {
                File.Delete(ruta);
            }
            wb.SaveAs(ruta);

        }
        public string queryInsert(JObject values, string table)
        {
            string query = $"INSERT INTO `{table}` (";
            string query2 = " VALUES (";
            foreach (JProperty key in (JToken)values)
            {
                string keyName = key.Name;
                string keyValue = key.Value.ToString();
                if (query == $"INSERT INTO `{table}` (")
                {
                    query = query + " `" + keyName + "`";
                }
                else
                {
                    query = query + ", `" + keyName + "`";
                }

                if (query2 == " VALUES (")
                {
                    query2 = query2 + "'" + keyValue + "'";
                }
                else
                {
                    query2 = query2 + ", '" + keyValue + "'";
                }

            }
            query = query + ")";
            query2 = query2 + ")";
            query = query + query2;
            return query;
        }
        #endregion



    }
}
