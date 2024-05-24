using DataBotV5.Logical.Projects.ControlDesk;
using System.Text.RegularExpressions;
using DataBotV5.Logical.Processes;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using DataBotV5.Data.Credentials;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Web;
using Newtonsoft.Json.Linq;
using System.Globalization;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using OpenQA.Selenium;
using Newtonsoft.Json;
using ExcelDataReader;
using System.Threading;
using System.Data;
using System.Linq;
using System.IO;
using System;

namespace DataBotV5.Logical.Projects.TIRequest
{
    internal class TiFunctions
    {
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly SapVariants sap = new SapVariants();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();

        string respFinal = "";

        /// <summary>
        /// Nota importante: Dejar en el json solo el mandante especificado.
        /// </summary>
        /// <param name="json"></param>
        /// <param name="mandante"></param>
        /// <returns></returns>
        private string[] SelectMandLines(string[] json, string mandante)
        {
            string[] json2 = (string[])json.Clone();
            for (int i = 0; i < json2.Length; i++)
            {
                DataTable itemTable = JsonConvert.DeserializeObject<DataTable>(json2[i]);
                try
                {
                    DataTable fsd = itemTable.Select("MANDANTE = '" + mandante + "'").CopyToDataTable();
                    json2[i] = JsonConvert.SerializeObject(fsd);
                }
                catch (Exception)
                {
                    json2[i] = "[]";
                }
            }

            return json2;
        }

        /// <summary>
        /// Crea el usuario con la información del json en el sistema especificado
        /// </summary>
        /// <param name="json">los datos</param>
        /// <param name="system">SAP, PORTAL o CD</param>
        /// <returns></returns>
        private DataTable CreateUser(string[] json, string system = "SAP")
        {
            if (json != null)
            {
                #region Tomar lista de los mandantes a Procesar
                DataTable mands = new DataTable();
                mands.Columns.Add("MANDANTE");

                foreach (string item in json)
                {
                    if (item != "[]")
                    {
                        DataTable itemTable = JsonConvert.DeserializeObject<DataTable>(item).DefaultView.ToTable(true, "MANDANTE"); //quita los duplicados
                        mands.Merge(itemTable);
                    }
                }
                mands = mands.DefaultView.ToTable(true, "MANDANTE");
                #endregion

                #region Procesar cada mandante de la lista

                mands.Columns.Add("RESPUESTA", typeof(string[]));

                foreach (DataRow item in mands.Rows)
                {
                    string mand = (string)item["MANDANTE"];
                    string[] response;

                    if (system.ToLower() == "sap")
                    {
                        response = CreateSapAll(mand, SelectMandLines(json, mand.ToString())); //procesar todo lo de sap
                    }
                    else if (system.ToLower() == "portal")
                    {
                        if (mand == "300")
                        {
                            response = CreateUserPortal(mand, SelectMandLines(json, mand.ToString()));
                        }
                        else
                        {
                            response = new string[] { "", "" };
                        }
                    }
                    else//CD
                    {
                        response = CreateUserCd(mand, SelectMandLines(json, mand.ToString()));
                    }

                    item["RESPUESTA"] = response;

                }

                return mands;
                #endregion
            }
            else
                return null;
        }
        /// <summary>
        /// Dar de baja el usuario en SAP
        /// </summary>
        /// <param name="mand"></param>
        /// <param name="InactiveUser">id de SAP del usuario</param>
        internal string DeleteUserSap(string sapUserName, int sapMand)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["USERNAME"] = sapUserName
            };

            IRfcFunction func = sap.ExecuteRFC("", "BAPI_USER_LOCK", parameters, sapMand);  // FM para bloquear el usuario
            string ret = func.GetTable("RETURN").GetValue("MESSAGE").ToString();
            if (ret.ToLower().Contains("locked"))
            {
                RfcDestination destErp = sap.GetDestRFC("", sapMand);

                IRfcFunction fmTableErp = destErp.Repository.CreateFunction("BAPI_USER_CHANGE");   // FM para la fecha de validez
                fmTableErp.SetValue("USERNAME", sapUserName);

                IRfcStructure fieldsLock = fmTableErp.GetStructure("LOGONDATA");
                fieldsLock.SetValue("GLTGB", DateTime.Today.ToString("yyyyMMdd"));

                IRfcStructure fieldsLockX = fmTableErp.GetStructure("LOGONDATAX");
                fieldsLockX.SetValue("GLTGB", "X");

                IRfcStructure licTypeX = fmTableErp.GetStructure("UCLASSX");
                licTypeX.SetValue("UCLASS", "X");

                fmTableErp.Invoke(destErp);

                ret = fmTableErp.GetTable("RETURN").GetValue("MESSAGE").ToString();

                if (ret.ToLower().Contains("changes"))
                    ret = "OK";
            }
            return ret;

        }

        /// <summary>
        /// Envía correo de notificación en caso de que la solicitud tenga roles ZERP
        /// </summary>
        /// <param name="request">el cuerpo del correo de la solicitud</param>
        private void ZerpNotification(string request)
        {
            string body = "";

            #region Leer info correo

            string pos = GetValFromBPM("Posición", request);
            string email = GetValFromBPM("Correo del usuario", request);
            string userId = email.Split(new string[] { "@" }, StringSplitOptions.None)[0];

            #endregion

            DataTable roles = crud.Select("SELECT name, idPos, rolId, client FROM newRolesPosicion WHERE idPos = " + pos, "ti_requests_db");

            foreach (DataRow rol in roles.Rows)
            {
                if (rol["rolId"].ToString().ToLower().Contains("zerp"))
                    body += rol["rolId"].ToString() + "<br>";
            }

            if (body != "")
            {
                body = "Se creo el siguiente usuario: " + userId + " posición: " + roles.Rows[0]["name"].ToString() + "<br>Con la asignación de los siguientes roles de su parte, por favor agregarlos, muchas gracias<br><br>" + body;
                mail.SendHTMLMail(body, new string[] { "rofernandez@gbm.net", "internalcustomersrvs@gbm.net" }, "Asignación de roles de Human Capital");
            }
        }

        /// <summary>
        /// Crea el usuario en SAP en el mandante especificado
        /// </summary>
        /// <param name="mandante">"300","500" o "400"</param>
        /// <param name="json">la solicitud en formato json</param>
        /// <returns>la respuesta del proceso en formato json y html</returns>
        private string[] CreateSapAll(string mandanteS, string[] json)
        {
            int mandante = int.Parse(mandanteS);
            string[] response = new string[6];

            RfcDestination destination = sap.GetDestRFC("", mandante);//RfcDestinationManager.GetDestination(cred.parametros);

            //USERS
            DataTable tempDt = CreateUserSap(json[0], destination);
            response[0] = JsonConvert.SerializeObject(tempDt);//JSON
            response[1] = val.ConvertDataTableToHTML(tempDt);//HTML

            //ROLES
            tempDt = CreateRolesSap(json[1], destination);
            response[2] = JsonConvert.SerializeObject(tempDt);//JSON
            response[3] = val.ConvertDataTableToHTML(tempDt);//HTML

            //PARAMETROS
            tempDt = CreateParametersSap(json[2], destination);
            response[4] = JsonConvert.SerializeObject(tempDt);//JSON
            response[5] = val.ConvertDataTableToHTML(tempDt);//HTML

            return response;
        }

        /// <summary>
        /// Crear el usuario en SAP(SU01)
        /// </summary>
        /// <param name="request">Json del usuario parte de datos generales</param>
        /// <param name="destination">destination de SAP</param>
        /// <returns>Datatable con el resultado del proceso</returns>
        private DataTable CreateUserSap(string request, RfcDestination destination)
        {
            DataTable resDt = new DataTable();

            if (request != "[]")
            {
                IRfcFunction zCreateUsers = destination.Repository.CreateFunction("Z_CREATE_USERS");
                IRfcTable inputTable = zCreateUsers.GetTable("INPUT");

                JArray jArray = JArray.Parse(request);

                foreach (JToken jsonUsers in jArray)
                {
                    inputTable.Append(); //linea del empleado

                    string user = jsonUsers["USUARIO"].ToString().Trim();
                    string name = jsonUsers["NOMBRE"].ToString().Trim();
                    string lastName = jsonUsers["APELLIDO"].ToString().Trim();
                    string email = jsonUsers["EMAIL"].ToString().Trim();
                    //Fix correo temporal del correo
                    email = email.Replace("@GBM.NET", "@gbmcorp.onmicrosoft.com");
                    string pass = jsonUsers["PASS"].ToString().Trim();
                    string type = jsonUsers["TIPO"].ToString().Trim();
                    string date = jsonUsers["FECHA VALIDEZ"].ToString().Trim();
                    string lic = jsonUsers["LICENCIA"].ToString().Trim();

                    inputTable.SetValue("USUARIO", user);
                    inputTable.SetValue("NOMBRE", name);
                    inputTable.SetValue("APELLIDO", lastName);
                    inputTable.SetValue("PASS", pass);
                    inputTable.SetValue("TIPO", type);
                    inputTable.SetValue("EMAIL", email);
                    inputTable.SetValue("LIC_TYPE", lic);
                    inputTable.SetValue("FECHA_VALIDEZ", date);
                }

                zCreateUsers.Invoke(destination);
                resDt = sap.GetDataTableFromRFCTable(zCreateUsers.GetTable("OUTPUT"));

                try { resDt.Columns.Remove("NOMBRE"); } catch (Exception) { }
                try { resDt.Columns.Remove("APELLIDO"); } catch (Exception) { }
                try { resDt.Columns.Remove("EMAIL"); } catch (Exception) { }
                try { resDt.Columns.Remove("PASS"); } catch (Exception) { }
                try { resDt.Columns.Remove("TIPO"); } catch (Exception) { }
                try { resDt.Columns.Remove("LIC_TYPE"); } catch (Exception) { }
                try { resDt.Columns.Remove("FECHA_VALIDEZ"); } catch (Exception) { }

            }
            return resDt;
        }

        /// <summary>
        /// Crear el usuario en SAP(Portal)
        /// </summary>
        /// <param name="mand">mandante de SAP actualmente solo "300"</param>
        /// <param name="json"></param>
        /// <returns></returns>
        private string[] CreateUserPortal(string mand, string[] json)
        {
            int retries = 0;
            string user = cred.username_SAPPRD;
            string pass = cred.password_rpauser_dominio;
            string[] response = new string[2];

            DataTable outputDatatable = new DataTable();
            outputDatatable.Columns.Add("STATUS");
            outputDatatable.Columns.Add("DISPLAY_NAME");
            outputDatatable.Columns.Add("ID");
            outputDatatable.Columns.Add("RESULT");

            IWebDriver chrome = null;
            string link = "http://ep-prod-app.gbm.net:50100/useradmin";

            try
            {
                WebInteraction sel = new WebInteraction();
                chrome = sel.NewSeleniumChromeDriver();
            }
            catch (Exception ex)
            {
                response[0] = ""; //en json
                response[1] = "<b>" + ex.Message; //en HTML
                return response;
            }

            try { ResetPage(); }
            catch (Exception) { ResetPage(); }

            if (json[0] != "[]")
            {
                JArray jArray = JArray.Parse(json[0]);

                for (int i = 0; i < jArray.Count;)
                {
                    JToken jsonRoles = jArray[i];

                    string createUser = jsonRoles["USUARIO"].ToString().Trim();
                    string name = jsonRoles["NOMBRE"].ToString().Trim();
                    string lastName = jsonRoles["APELLIDO"].ToString().Trim();
                    string email = jsonRoles["EMAIL"].ToString().Trim();
                    string portal = jsonRoles["PORTAL"].ToString().Trim();
                    string password = jsonRoles["PASS"].ToString().Trim();

                    try
                    {
                        #region Crear usuario
                        if (portal.ToUpper() == "SI" && mand == "300")
                        {
                            string batch = "";
                            if (email == "")
                            {
                                //es freelance
                                batch = "[User]\n" +
                                        "uid=" + createUser + "\n" +
                                        "Password=Inicio01\n" +
                                        "last_name=" + lastName + "\n" +
                                        "first_name=" + name + "\n" +
                                        "language=en_US\n" +
                                        "group=Authenticated Users; SW_KM_Usuario_Final; Cuenta_gastos_tiempos; Everyone;";
                            }
                            else
                            {
                                //el regular
                                batch = "[User]\nuid=" + createUser + "\nlanguage=en_US\ngroup=Cuenta_gastos_tiempos;\n";
                            }


                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJHHFA.BatchImportCompView.TextEdit"))); } catch { }
                            chrome.FindElement(By.Id("INMJHHFA.BatchImportCompView.TextEdit")).SendKeys(batch); //llenar el text field


                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJHHFA.BatchImportCompView.CheckBox0-img"))); } catch { }
                            chrome.FindElement(By.Id("INMJHHFA.BatchImportCompView.CheckBox0-img")).Click(); //click en overwrite


                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJHHFA.BatchImportCompView.Button0"))); } catch { }
                            chrome.FindElement(By.Id("INMJHHFA.BatchImportCompView.Button0")).Click(); //click en upload

                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='INMJHHFA.BatchImportCompProtocolView.Table-contentTBody']/tr"))); } catch { }
                            IReadOnlyCollection<IWebElement> webResponseTable = chrome.FindElements(By.XPath("//*[@id='INMJHHFA.BatchImportCompProtocolView.Table-contentTBody']/tr"));

                            foreach (IWebElement item in webResponseTable)
                            {
                                if (item.GetAttribute("outerHTML").Contains("udat="))
                                {
                                    string[] temp = item.GetAttribute("innerText").Split(new char[] { '\t' });
                                    DataRow row = outputDatatable.NewRow();
                                    for (int j = 1; j < temp.Length; j++)
                                    {
                                        row[j - 1] = temp[j];
                                    }
                                    outputDatatable.Rows.Add(row);
                                }
                            }


                            //return to import, SE PUEDEN HACER VARIOS AL MISMO PERO POR SIMPLICIDAD DE CODIGO MAS FACIL DARLE ATRAS
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJHHFA.BatchImportCompProtocolView.Back"))); } catch { }
                            chrome.FindElement(By.Id("INMJHHFA.BatchImportCompProtocolView.Back")).Click(); //click en return to import

                            retries = 0;
                        }

                        i++;
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        retries++;
                        if (retries > 5 && retries != 0)
                        {
                            DataRow row = outputDatatable.NewRow();
                            row["DISPLAY_NAME"] = createUser;
                            row["RESULT"] = "Error de Selenium(" + ex.Message + "), revisar en Portal";
                            outputDatatable.Rows.Add(row);
                            retries = 0;
                            i++;
                            ResetPage();
                        }
                        else
                            ResetPage();
                    }
                }

                try { chrome.FindElement(By.Id("INMJHHFA.BatchImportCompProtocolView.Back")).Click(); } catch (Exception) { } //click en return to import 

                chrome.Quit();
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);

                //respuestas
                response[0] = JsonConvert.SerializeObject(outputDatatable); //en json
                response[1] = val.ConvertDataTableToHTML(outputDatatable); //en HTML

            }

            void ResetPage()
            {
                try
                {
                    chrome.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(1000);
                    chrome.Navigate().GoToUrl(link);
                    Thread.Sleep(1000);
                    chrome.FindElement(By.Id("logonuidfield")).SendKeys(user);
                    chrome.FindElement(By.Id("logonpassfield")).SendKeys(pass);
                    chrome.FindElement(By.Name("uidPasswordLogon")).Click();
                    chrome.FindElement(By.Id("INMJ.UmeAdminCompView.Button2")).Click(); //click en import
                }
                catch (Exception) { }
            }

            return response;


        }

        /// <summary>
        /// Crear el usuario en Control Desk
        /// </summary>
        /// <param name="mandante"></param>
        /// <param name="json"></param>
        /// <returns></returns>
        private string[] CreateUserCd(string mandante, string[] json)
        {
            cred.SelectCdMand(mandante);

            DataTable outputDatatable = new DataTable();
            string[] columns = { "USUARIO", "RESPUESTA" };
            string[] response = new string[2];

            foreach (string column in columns)
                outputDatatable.Columns.Add(column);

            if (json[0] != "[]")
            {
                JArray jArray = JArray.Parse(json[0]);

                for (int i = 0; i < jArray.Count; i++)
                {
                    JToken jsonRoles = jArray[i];

                    string user = jsonRoles["USUARIO"].ToString().Trim();
                    string[] roles = jsonRoles["ROLES"].ToString().Split(',');

                    CdUserData cdUser = new CdUserData
                    {
                        User = user,
                        Roles = roles
                    };

                    string resultCd = cdi.CreateUser(cdUser);

                    DataRow dtrow = outputDatatable.NewRow();
                    dtrow[columns[0]] = user;
                    dtrow[columns[1]] = resultCd;
                    outputDatatable.Rows.Add(dtrow);
                }

                //respuestas
                response[0] = JsonConvert.SerializeObject(outputDatatable); //en json
                response[1] = val.ConvertDataTableToHTML(outputDatatable); //en HTML
            }

            return response;
        }

        /// <summary>
        /// Agregar los parámetros al Usuario (SU01)
        /// </summary>
        /// <param name="json"></param>
        /// <param name="destination"></param>
        /// <returns></returns>
        private DataTable CreateParametersSap(string json, RfcDestination destination)
        {
            DataTable outputDatatable = new DataTable();

            if (json != "[]")
            {
                //hacer roles
                IRfcFunction func = destination.Repository.CreateFunction("Z_ADD_USER_PARAMETERS");
                IRfcTable inputTable = func.GetTable("PARAMETROS");

                JArray jArray = JArray.Parse(json);

                DataTable parametersDt = JsonConvert.DeserializeObject<DataTable>(jArray.ToString());
                DataTable uniqueUsers = parametersDt.DefaultView.ToTable(true, "USUARIO");

                for (int i = 0; i < uniqueUsers.Rows.Count; i++)
                {
                    inputTable.Append(); //linea del empleado
                    IRfcTable parametrosTable = inputTable[i].GetTable("PARAMETROS");
                    string user = uniqueUsers.Rows[i]["USUARIO"].ToString();
                    inputTable.SetValue("USUARIO", user);
                    DataRow[] uniqueUserParameters = parametersDt.Select("USUARIO = '" + user + "'");
                    foreach (DataRow item in uniqueUserParameters)
                    {
                        parametrosTable.Append();
                        parametrosTable.SetValue("PARID", item["PARAMETRO"].ToString().Trim());
                        parametrosTable.SetValue("PARVA", item["VALOR"].ToString().Trim());
                    }
                }

                func.Invoke(destination);
                outputDatatable = sap.GetDataTableFromRFCTable(func.GetTable("OUT_PARAMETROS"));
                outputDatatable.Columns.Remove("PARAMETROS");
            }

            return outputDatatable;
        }

        /// <summary>
        /// Agregar los roles al Usuario (SU01)
        /// </summary>
        /// <param name="json"></param>
        /// <param name="destination"></param>
        /// <returns></returns>
        private DataTable CreateRolesSap(string json, RfcDestination destination)
        {
            DataTable outputDatatable = new DataTable();

            if (json != "[]")
            {
                //hacer roles
                IRfcFunction zAddUsersRoles = destination.Repository.CreateFunction("Z_ADD_USERS_ROLES");
                IRfcFunction prgnJ2eeUserGetRoles = destination.Repository.CreateFunction("PRGN_J2EE_USER_GET_ROLES");
                IRfcTable inputTable = zAddUsersRoles.GetTable("ROLES");

                JArray jArray = JArray.Parse(json);

                for (int i = 0; i < jArray.Count; i++)
                {
                    JToken jsonRoles = jArray[i];
                    inputTable.Append(); //linea del empleado

                    IRfcTable rolesTable = inputTable[i].GetTable("ROLES");
                    string userId = jsonRoles["USUARIO"].ToString().Trim();
                    string[] roles = jsonRoles["ROLES"].ToString().Split(',');

                    inputTable.SetValue("USUARIO", userId);

                    //Antes de agregar roles hay que leer los actuales ya que la FM solo sobrescribe los que se le envie

                    prgnJ2eeUserGetRoles.SetValue("IF_UNAME", userId.ToUpper());
                    prgnJ2eeUserGetRoles.Invoke(destination);
                    DataTable currentRoles = sap.GetDataTableFromRFCTable(prgnJ2eeUserGetRoles.GetTable("ET_AGR"));

                    //llenar roles_table

                    foreach (string rol in roles)
                    {
                        rolesTable.Append();
                        rolesTable.SetValue("AGR_NAME", rol.Trim());
                    }

                    if (currentRoles.Rows.Count > 0)
                    {
                        //agregar los roles que tiene actualmente
                        foreach (DataRow fila in currentRoles.Rows)
                        {
                            rolesTable.Append();
                            rolesTable.SetValue("AGR_NAME", fila["AGR_NAME"]);
                        }
                    }
                }

                zAddUsersRoles.Invoke(destination);
                outputDatatable = sap.GetDataTableFromRFCTable(zAddUsersRoles.GetTable("OUT_ROLES"));
                outputDatatable.Columns.Remove("ROLES");
            }

            return outputDatatable;
        }

        //Metodos públicos

        /// <summary>
        /// Convertir el excel de la plantilla al formato json, para el sistema especificado
        /// </summary>
        /// <param name="path">ruta del archivo del excel</param>
        /// <param name="system">SAP o CD</param>
        /// <returns></returns>
        internal string[] ExcelToJson(string path, string system = "SAP")
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = null;
            string[] response;

            DataTable usersTable = new DataTable();
            DataTable rolesTable = new DataTable();
            DataTable parametersTable = new DataTable();
            DataTable rolesCdTable = new DataTable();

            JsonSerializerSettings settings = new JsonSerializerSettings { Converters = { new FormatNumbersAsTextConverter() } };

            try
            {
                result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false,
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                });

                stream.Close();

                usersTable = result.Tables["Usuarios"];
                rolesTable = result.Tables["Roles"];
                parametersTable = result.Tables["Parametros"];
                rolesCdTable = result.Tables["CD"];

            }
            catch (Exception)
            {
                stream.Close();
                mail.SendHTMLMail("Error al leer la plantilla", new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
            }

            string validation = "";
            try { validation = usersTable.Columns[11].ColumnName; } catch (Exception) { }

            if (validation.Substring(0, 1) != "x")
            {
                response = new string[1];
                response[0] = "Utilizar la plantilla oficial";
                mail.SendHTMLMail(response[0], new string[] { root.BDUserCreatedBy }, root.Subject);
            }
            else
            {
                usersTable.Columns.Remove("x");
                //convertir a Json
                if (system.ToLower() == "sap")
                {
                    response = new string[3];
                    response[0] = JsonConvert.SerializeObject(usersTable, settings);
                    response[1] = JsonConvert.SerializeObject(rolesTable, settings);
                    response[2] = JsonConvert.SerializeObject(parametersTable, settings);
                }
                else if (system.ToLower() == "105")
                {
                    response = new string[1];
                    DataTable table105 = new DataTable();
                    table105.Columns.Add("colabId");
                    table105.Columns.Add("email");

                    foreach (DataRow userRow in usersTable.Rows)
                    {
                        DataRow row105 = table105.NewRow();
                        row105["colabId"] = userRow["ID SAP"];
                        row105["email"] = userRow["EMAIL"];
                        table105.Rows.Add(row105);
                    }
                    response[0] = JsonConvert.SerializeObject(table105, settings);
                }
                else //por ahora pues seria CD
                {
                    response = new string[1];
                    response[0] = JsonConvert.SerializeObject(rolesCdTable, settings);
                }
            }

            return response;
        }

        /// <summary>
        /// Verifica que la posición esta en la BD con sus roles y mandante
        /// </summary>
        /// <param name="position">id de la posición de BPM(correo)</param>
        /// <returns>true: Si existe, false: No existe</returns>
        internal bool IsValidPosition(string position)
        {
            if (crud.Select("SELECT name FROM newRolesPosicion WHERE idPos ='" + position + "'", "ti_requests_db").Rows.Count > 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Verifica el tipo de la solicitud para saber si es una solicitud valida, en caso de que sea un tipo valido con un correo erróneo se notifica por correo.
        /// </summary>
        /// <param name="request">Cuerpo de la solicitud de BPM</param>
        /// <returns>
        /// "INVALID" si no es valido, "NUEVO" si es alta, "BAJA" si es BAJA
        /// </returns>
        internal string GetRequestType(string request)
        {
            string type = "INVALID";

            //Lista con los correos que son excepciones, o correos inválidos
            DataTable temp = crud.Select("SELECT `invalid` FROM `invalidMails`", "ti_requests_db");

            List<string> invalidMails = new List<string>();
            foreach (DataRow row in temp.Rows)
                invalidMails.Add(row[0].ToString());

            //Validar solicitud
            string requestType = GetValFromBPM("Tipo Solicitud:", request);
            if (requestType == "NUEVO")
            {
                if (invalidMails.Contains(GetValFromBPM("Correo del usuario:", request))) //si el correo esta en la lista InvalidMails esta malo 
                    mail.SendHTMLMail(root.Email_Body.Replace("\n", "<br>"), new string[] { "ajrojas@gbm.net", "internalcustomersrvs@gbm.net" }, root.Subject + "** Email Inválido **");
                else
                    type = requestType;
            }
            else if (requestType == "BAJA")
                type = requestType;
            else if (requestType == "BAJAICS")
                type = requestType;

            return type;
        }

        /// <summary>
        /// Tomar el campo deseado del correo de BPM
        /// </summary>
        /// <param name="field">el campo a buscar</param>
        /// <param name="body">el formato del cuerpo del correo que envía BPM</param>
        /// <returns>El valor del campo deseado</returns>
        internal string GetValFromBPM(string field, string body)
        {
            Regex htmlFix = new Regex("[*'\"_&+^><]");
            Regex alphanum = new Regex(@"[^\p{L}0-9 -.@]");
            string[] separator = new string[] { field };

            body = htmlFix.Replace(body, string.Empty);

            string[] bodySplit = body.Split(separator, StringSplitOptions.None);
            bodySplit[1] = bodySplit[1].Replace('\r', ' ');
            bodySplit = bodySplit[1].Split('\n');
            string val = alphanum.Replace(bodySplit[0], "").Trim().ToUpper();

            if (field == "País")
            {
                switch (val.ToLower())
                {
                    case "1":
                    case "costa rica":
                        val = "CR";
                        break;
                    case "2":
                    case "república dominicana":
                        val = "DR";
                        break;
                    case "3":
                    case "guatemala":
                        val = "GT";
                        break;
                    case "4":
                    case "honduras":
                        val = "HN";
                        break;
                    case "5":
                    case "el salvador":
                        val = "SV";
                        break;
                    case "6":
                    case "panamá":
                        val = "PA";
                        break;
                    case "7":
                    case "gbmhq":
                        val = "HQ";
                        break;
                    case "8":
                    case "venezuela":
                        val = "HQ";
                        break;
                    case "9":
                    case "nicaragua":
                        val = "NI";
                        break;
                    case "11":
                    case "bvi":
                        val = "BVI";
                        break;
                    case "12":
                    case "florida":
                        val = "MD";
                        break;
                    case "miami":
                        val = "MD";
                        break;
                    case "13":
                    case "colombia":
                        val = "CO";
                        break;
                    default:
                        val = "";
                        break;
                }
            }

            if (field == "Correo del usuario:" || field == "Correo del usuario")
            {
                if (val.ToUpper().Contains("MAILTO"))
                    val = val.Split(new string[] { "MAILTO" }, StringSplitOptions.None)[0].Trim();
                val = val.Replace("@GBM.NET@GBM.NET", "@GBM.NET");
                val = val.Replace("@GBM@GBM.NET", "@GBM.NET");
            }

            return val;
        }

        /// <summary>
        /// Toma el cuerpo del correo de BPM y lo convierte a Json
        /// </summary>
        /// <param name="request">el cuerpo del correo de BPM</param>
        /// <param name="system">"SAP", "CD" o "105"</param>
        /// <returns>El Json adecuado para el sistema especificado</returns>
        internal string[] BpmToJson(string request, string system = "SAP")
        {
            //3 becado, 1 regular, 2 temporal, 4 practicante

            Dictionary<string, List<string>> mandRoles = new Dictionary<string, List<string>>();
            Dictionary<string, string> mandTypes = new Dictionary<string, string>();
            Dictionary<string, string> mandLicense = new Dictionary<string, string>();

            List<string> userRolesErp = new List<string>();
            List<string> userRolesCrm = new List<string>();
            List<string> userRolesFiori = new List<string>();

            mandRoles.Add("300", userRolesErp);
            mandRoles.Add("500", userRolesCrm);
            mandRoles.Add("400", userRolesFiori);

            #region Leer info correo

            string name = GetValFromBPM("Nombre Colaborador", request);
            string lastname = GetValFromBPM("Apellido del Colaborador", request);
            string country = GetValFromBPM("País", request);
            string pos = GetValFromBPM("Posición", request);
            string date = GetValFromBPM("Fecha de fin de contrato", request);
            string email = GetValFromBPM("Correo del usuario", request);
            string location = GetValFromBPM("Localidad", request); //5 lindora uno y 6 lindora dos
            string colabId = GetValFromBPM("Número de Colaborador", request);

            //string comments = GetValFromBPM("Comentarios", solicitud);
            //string solType = GetValFromBPM("Tipo Solicitud", solicitud);
            //string plaType = GetValFromBPM("Tipo Plaza", solicitud);

            #endregion

            #region Tomar el UserId

            string[] userId = email.Split(new string[] { "@" }, StringSplitOptions.None);

            #endregion

            #region Tomar la fecha

            if (date == "1-1-1970" || date == "01-01-1000")
                date = "";
            else
            {
                string[] splitDate = date.Split('-');
                string dd = splitDate[0];
                string mm = splitDate[1];
                string yyyy = splitDate[2];
                date = DateTime.Parse(yyyy + "-" + mm + "-" + dd).ToString("yyyy-MM-dd");
            }

            #endregion

            #region Leer Roles

            DataTable roles = crud.Select("SELECT name, idPos, rolId, client FROM newRolesPosicion WHERE idPos IN ('0','" + pos + "')", "ti_requests_db");  //0 es default

            //eliminar los roles ZERP
            bool dialogTypeException = false;
            for (int i = roles.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = roles.Rows[i];
                if (dr["rolId"].ToString().ToLower().Contains("zerp"))
                {
                    dr.Delete();
                    dialogTypeException = true; //Flag para saltar la validación del tipo y licencia
                }
            }
            roles.AcceptChanges();

            foreach (DataRow rol in roles.Rows)
            {
                foreach (string mand in rol["client"].ToString().Split(','))
                {
                    if (mand.Trim() == "300")
                        userRolesErp.Add(rol["rolId"].ToString());
                    else if (mand.Trim() == "400")
                        userRolesFiori.Add(rol["rolId"].ToString());
                    else if (mand.Trim() == "500")
                        if (roles.Select("client = '500' and idPos <> '0'").Count() > 0) //si hay mas roles aparte que los default
                            userRolesCrm.Add(rol["rolId"].ToString());
                }
            }
            #endregion

            #region Arreglar Roles XX
            foreach (KeyValuePair<string, List<string>> item in mandRoles)
            {
                for (int i = 0; i < mandRoles[item.Key].Count; i++)
                {
                    //Casos que no se reemplaza por el país
                    if (country == "HQ" & (mandRoles[item.Key][i] == "ZSAP_EMPLOYEE_ERP_XX" || mandRoles[item.Key][i] == "Z_PERFIL_GENERAL_ENTRADA_XX"))
                    {
                        if (mandRoles[item.Key][i] == "ZSAP_EMPLOYEE_ERP_XX")
                        {
                            mandRoles[item.Key][i] = mandRoles[item.Key][i].Replace("XX", "CORP");
                        }

                        if ((location == "5" || location == "6") && mandRoles[item.Key][i] == "Z_PERFIL_GENERAL_ENTRADA_XX")//es de Lindora
                            mandRoles[item.Key][i] = mandRoles[item.Key][i].Replace("XX", "LIND");
                        else
                            mandRoles[item.Key][i] = mandRoles[item.Key][i].Replace("XX", "CORP");
                    }
                    else
                        mandRoles[item.Key][i] = mandRoles[item.Key][i].Replace("XX", country);


                    string posName = "";
                    try { posName = roles.Select("idPos = " + pos)[0]["name"].ToString().ToUpper(); } catch (Exception) { }


                    if (posName.Contains("HUMAN CAPITAL"))
                    {
                        if (mandRoles[item.Key][i].Contains("ZSAP_EMPLOYEE_ERP"))
                            mandRoles[item.Key][i] = "ZSAP_EMPLOYEE_ERP_HCM";
                    }

                }
            }
            #endregion

            #region Colocar el tipo
            //400 -> siempre es A
            mandTypes.Add("400", "A");
            //500 -> siempre es A
            mandTypes.Add("500", "A");
            //300 -> si tiene mas roles que los default en 300 es A sino es B
            if (roles.Select("client = '300' and idPos <> '0'").Count() > 0)
                mandTypes.Add("300", "A");
            else
            {
                if (dialogTypeException)
                    mandTypes.Add("300", "A");
                else
                    mandTypes.Add("300", "B");
            }
            #endregion

            #region Colocar la licencia
            //400 -> no lleva lic
            mandLicense.Add("400", "null");

            //300 -> depende
            if (mandTypes["300"] == "A" && mandTypes["500"] == "A") //si el tipo es A en 300 y 500 = CB
                mandLicense.Add("300", "CB");
            else if ((mandTypes["300"] == "A" && mandTypes["500"] == "B") || (mandTypes["300"] == "B" && mandTypes["500"] == "A")) //A en uno y B en otro = CC,
            {
                if (roles.Select("client = '500' and idPos <> '0'").Count() > 0 /*&& tipos_mandante["300"] == "A"*/) //si no tiene roles de 500
                    mandLicense.Add("300", "CC");
                else
                {
                    if (dialogTypeException)
                        mandLicense.Add("300", "CB");
                    else
                        mandLicense.Add("300", "CE");
                }
            }
            else
                mandLicense.Add("300", "");

            //500 -> es la misma de 300
            mandLicense.Add("500", mandLicense["300"]);
            #endregion

            #region Crear los json
            date = date == "" ? "null" : "\"" + date + "\"";

            string json1Crm = "";
            if (roles.Select("client = '500' and idPos <> '0'").Count() > 0) //tiene roles de 500
            {
                json1Crm = "{\"MANDANTE\":\"500\"," +
                            "\"USUARIO\":\"" + userId[0] + "\"," +
                            "\"NOMBRE\":\"" + name + "\"," +
                            "\"APELLIDO\":\"" + lastname + "\"," +
                            "\"EMAIL\":\"" + email + "\"," +
                            "\"PASS\":\"Inicio01\"," +
                            "\"TIPO\":\"" + mandTypes["500"] + "\"," +
                            "\"LICENCIA\":\"" + mandLicense["500"] + "\"," +
                            "\"FECHA VALIDEZ\":" + date + "," +
                            "\"PORTAL\":null},";
            }

            string json1 = "[" +

                json1Crm +

                "{\"MANDANTE\":\"300\"," +
                "\"USUARIO\":\"" + userId[0] + "\"," +
                "\"NOMBRE\":\"" + name + "\"," +
                "\"APELLIDO\":\"" + lastname + "\"," +
                "\"EMAIL\":\"" + email + "\"," +
                "\"PASS\":\"Inicio01\"," +
                "\"TIPO\":\"" + mandTypes["300"] + "\"," +
                "\"LICENCIA\":\"" + mandLicense["300"] + "\"," +
                "\"FECHA VALIDEZ\":" + date + "," +
                "\"PORTAL\":\"SI\"}," +

                "{\"MANDANTE\":\"400\"," +
                "\"USUARIO\":\"" + userId[0] + "\"," +
                "\"NOMBRE\":\"" + name + "\"," +
                "\"APELLIDO\":\"" + lastname + "\"," +
                "\"EMAIL\":\"" + email + "\"," +
                "\"PASS\":\"Inicio01\"," +
                "\"TIPO\":\"" + mandTypes["400"] + "\"," +
                "\"LICENCIA\":null," +
                "\"FECHA VALIDEZ\":" + date + "," +
                "\"PORTAL\":null}" +

                "]";


            string json2Crm = "";
            if (roles.Select("client = '500' and idPos <> '0'").Count() > 0) //tiene roles de 500
                json2Crm = "{\"MANDANTE\":\"500\",\"USUARIO\":\"" + userId[0] + "\",\"ROLES\":\"" + String.Join(",", mandRoles["500"].ToArray()) + "\"}";


            string json2 = "[" +

                "{\"MANDANTE\":\"300\",\"USUARIO\":\"" + userId[0] + "\",\"ROLES\":\"" + String.Join(",", mandRoles["300"].ToArray()) + "\"}," +
                "{\"MANDANTE\":\"400\",\"USUARIO\":\"" + userId[0] + "\",\"ROLES\":\"" + String.Join(",", mandRoles["400"].ToArray()) + "\"}," + json2Crm +

                "]";

            string json3 = "[" +

                "{\"MANDANTE\":\"300\",\"USUARIO\":\"" + userId[0] + "\",\"PARAMETRO\":\"CAC\",\"VALOR\":\"co01\"}," +
                "{\"MANDANTE\":\"300\",\"USUARIO\":\"" + userId[0] + "\",\"PARAMETRO\":\"EFB\",\"VALOR\":\"gb\"}," +
                "{\"MANDANTE\":\"300\",\"USUARIO\":\"" + userId[0] + "\",\"PARAMETRO\":\"EVO\",\"VALOR\":\"01\"}," +
                "{\"MANDANTE\":\"300\",\"USUARIO\":\"" + userId[0] + "\",\"PARAMETRO\":\"CVR\",\"VALOR\":\"zgbm\"}" +

                "]";

            string json4 = "[{\"colabId\":\"" + colabId + "\",\"email\":\"" + email + "\"}]";

            string jsonCd = "[{\"MANDANTE\":\"PRD\",\"USUARIO\":\"" + email + "\",\"ROLES\":\"SDASELFSERV,IBMINC\"}]";


            #endregion

            switch (system)
            {
                case "CD":
                    return new string[] { jsonCd };
                case "105":
                    return new string[] { json4 };
                default:
                    return new string[] { json1, json2, json3 };
            }
        }

        /// <summary>
        /// Procesar ambos Json
        /// </summary>
        /// <param name="jsonSap">Json de SAP</param>
        /// <param name="jsonCd">Json de CD</param>
        /// <param name="json105">Json creación infotipo 105</param>
        internal void ProcessAllSystems(string[] jsonSap, string[] jsonCd, string[] json105 = null)
        {
            console.WriteLine(" > > > " + "Creando usuario de SAP");

            #region Crear usuarios de SAP
            DataTable resultSap = CreateUser(jsonSap, "SAP");
            #endregion

            console.WriteLine(" > > > " + "Creando usuario de CD");

            #region Crear Usuarios de CD
            DataTable resultCd = CreateUser(jsonCd, "CD");
            #endregion

            console.WriteLine(" > > > " + "Creando usuario en Portal");

            #region Crear Usuarios del Portal
            DataTable resultPortal = CreateUser(jsonSap, "portal");
            #endregion

            console.WriteLine(" > > > " + "Creando infotipo 105");

            #region Añadir Email al infotipo 105
            DataTable result105 = AddEmailto105(json105);
            #endregion

            console.WriteLine(" > > > " + "Enviar notificaciones en caso de Roles ZERP");

            #region Notificar si el Usuario tenia Roles ZERP
            if (!string.IsNullOrWhiteSpace(root.Email_Body) && root.BDUserCreatedBy.ToLower() == "bpm@mailgbm.com")
                ZerpNotification(root.Email_Body);
            #endregion

            console.WriteLine(" > > > " + "Finalizando solicitud");

            #region Construir Notificación
            string body = "";

            foreach (DataRow row in resultSap.Rows)
            {
                string[] result = (string[])row["RESPUESTA"];

                body = body + "RESULTADOS MANDANTE: " + row["MANDANTE"].ToString() + "<br><br>";
                body = body + "Creación de usuario en SAP:<br>" + result[1] + "<br>Asignar roles al usuario:<br>" + result[3] + "<br>Asignar parámetros al usuario:<br>" + result[5] + "<br><hr>";
                log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, row["MANDANTE"].ToString(), result[0] + result[2] + result[4], root.Subject);
                respFinal = respFinal + "\\n" + "Creación de usuario en SAP:<br>" + result[1] + " Asignar roles al usuario: " + result[3] + " Asignar parámetros al usuario: " + result[5];
            }

            foreach (DataRow row in resultCd.Rows)
            {
                string[] result = (string[])row["RESPUESTA"];

                body = body + "RESULTADOS Control Desk " + row["MANDANTE"].ToString() + ":<br><br>";
                body = body + "Asignación de roles en Control Desk:<br>" + result[1] + "<br><hr>";
                log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Control Desk " + row["MANDANTE"].ToString(), result[0], root.Subject);
                respFinal = respFinal + "\\n" + "Asignación de roles en Control Desk:<br>" + result[1];
            }

            foreach (DataRow row in resultPortal.Rows)
            {
                string[] result = (string[])row["RESPUESTA"];
                if (result[0] != "" || result[1] != "")
                {
                    body += "RESULTADOS Portal :<br><br>";
                    body = body + "Creación de usuario en Portal:<br>" + result[1] + "<br><hr>";
                    log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Portal " + row["MANDANTE"].ToString(), result[0], root.Subject);
                    respFinal = respFinal + "\\n" + "Creación de usuario en Portal:<br>" + result[1];
                }
            }

            #region Resultado 105
            if (result105.Rows.Count > 0)
            {
                body += "RESULTADOS infotipo 105 :<br><br>";
                body = body + val.ConvertDataTableToHTML(result105) + "<br><hr>";
                log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Infotipo 105 en TiRequests ", val.ConvertDataTableToHTML(result105), root.Subject);
                respFinal = respFinal + "\\n" + "infotipo 105:<br>" + result105;
            }
            #endregion

            root.BDUserCreatedBy = root.BDUserCreatedBy == "bpm@mailgbm.com" ? "internalcustomersrvs@gbm.net" : root.BDUserCreatedBy;
            root.requestDetails = respFinal;
            mail.SendHTMLMail(body, new string[] { root.BDUserCreatedBy }, root.Subject);
            #endregion
        }

        /// <summary>
        /// Agrega el Email al infotipo 105 (este metodo reemplazaba el boton "105" de BPM)
        /// </summary>
        /// <param name="jsonSap">Json con la info necesaria</param>
        /// <returns></returns>
        private DataTable AddEmailto105(string[] json)
        {

            DataTable outputDt = new DataTable();
            if (json != null)
            {
                outputDt.Columns.Add("colabId");
                outputDt.Columns.Add("user");
                outputDt.Columns.Add("result");

                string json105 = json[0];

                if (!string.IsNullOrEmpty(json105) && json105 != "[]")
                {
                    JArray jUsers = JArray.Parse(json105);
                    DataTable users = JsonConvert.DeserializeObject<DataTable>(jUsers.ToString());

                    foreach (DataRow user in users.Rows)
                    {
                        DataRow outputDtRow = outputDt.NewRow();
                        try
                        {
                            string colabId = user["colabId"].ToString();
                            string email = user["email"].ToString();
                            string userId = email.Split('@')[0].Trim().ToUpper();

                            Dictionary<string, string> fmParams = new Dictionary<string, string>
                            {
                                ["ID"] = colabId.ToUpper(),
                                ["USER"] = userId,
                                ["BEGDATE"] = DateTime.Now.ToString("yyyyMMdd"),
                                ["ENDDATE"] = "99991231",
                                ["SUBTY"] = "0001"
                            };

                            IRfcFunction zChangePa105 = sap.ExecuteRFC("ERP", "ZCHANGE_PA105", fmParams);
                            string response = zChangePa105.GetValue("RESPONSE").ToString();

                            outputDtRow["colabId"] = colabId;
                            outputDtRow["user"] = userId;
                            outputDtRow["result"] = response;

                        }
                        catch (Exception ex)
                        {
                            outputDtRow["result"] = ex.Message;
                        }

                        outputDt.Rows.Add(outputDtRow);
                    }
                }
            }
            return outputDt;
        }

        /// <summary>
        /// Metodo que obtiene el username de SAP a partir de su ID de colaborador 
        /// </summary>
        /// <param name="sapUserId">Id de colaborador</param>
        internal string GetSapUserName(string sapUserId)
        {
            try
            {
                RfcDestination destErp = sap.GetDestRFC("ERP");
                IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
                fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
                fmMg.SetValue("QUERY_TABLE", "PA0105");
                fmMg.SetValue("DELIMITER", "|");

                IRfcTable fields = fmMg.GetTable("FIELDS");
                fields.Append();
                fields.SetValue("FIELDNAME", "ENDDA");
                fields.Append();
                fields.SetValue("FIELDNAME", "USRID");

                IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
                fmOptions.Append();
                fmOptions.SetValue("TEXT", "PERNR = '" + sapUserId + "' and SUBTY = '0001'");

                fmMg.Invoke(destErp);

                DataTable parsedResponse = new DataTable();
                parsedResponse.Columns.Add("ENDDA");
                parsedResponse.Columns.Add("USRID");

                foreach (DataRow row in sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA")).Rows)
                {
                    DataRow parsedResponseRow = parsedResponse.NewRow();
                    string date = row["LINE"].ToString().Split(new char[] { '|' })[0].Trim();
                    parsedResponseRow["ENDDA"] = DateTime.ParseExact(date, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                    parsedResponseRow["USRID"] = row["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                    parsedResponse.Rows.Add(parsedResponseRow);
                }

                string maxDate = parsedResponse.AsEnumerable().Max(r => DateTime.Parse(r.Field<string>("ENDDA"))).ToString("yyyy-MM-dd");

                string usrId = parsedResponse.Select("ENDDA = '" + maxDate + "'")[0]["USRID"].ToString();

                return usrId;
            }
            catch (Exception ex)
            {
                return "ERROR: " + ex.Message;
            }
        }

        /// <summary>
        /// Toma las solicitudes del correo y las guarda en la BD
        /// </summary>
        internal void ProcessEmailRequest(string crudDb = "PRD")
        {
            mail.GetAttachmentEmail("Solicitudes TI BPM", "Procesados", "Procesados Solicitudes TI");
            if (!string.IsNullOrWhiteSpace(root.Email_Body) && (root.BDUserCreatedBy.ToLower() == "bpm@mailgbm.com" || root.BDUserCreatedBy.ToLower().Contains("databot")))
            {
                console.WriteLine(" > > > " + "Nueva Solicitud de TI por CORREO de BPM");

                string requestType = GetRequestType(root.Email_Body); //Baja, Alta o Invalid

                if (requestType != "INVALID")
                    crud.NonQueryAndGetId("INSERT INTO `pending` (`emailBody`,`date`) VALUES ('" + root.Email_Body + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "');", "ti_requests_db");
            }
        }
        /// <summary>
        /// Dar de baja el usuario en SAP(Portal)
        /// </summary>
        /// <param name="mand"></param>
        /// <param name="InactiveUser">id de SAP del usuario</param>
        public string DeleteUserPortal(string InactiveUser)
        {
            WebInteraction sel = new WebInteraction();
            try
            {
                IWebDriver chrome = sel.NewSeleniumChromeDriver();
                string link = "http://ep-prod-app.gbm.net:50100/useradmin";
                string user = cred.username_SAPPRD;
                string pass = cred.password_rpauser_dominio;

                try { ResetPage(); }
                catch (Exception) { ResetPage(); }

                chrome.FindElement(By.Id("INMJJKNE.BasicSearchView.PrincipalIdIF_NL")).SendKeys(InactiveUser);
                chrome.FindElement(By.Id("INMJJKNE.BasicSearchView.SearchButton1")).Click();

                try
                {
                    new WebDriverWait(chrome, new TimeSpan(0, 0, 5)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("INMJ.UmeAdminCompView.MessageArea1")));

                    string errMsg = "ERROR";
                    try
                    {
                        errMsg = chrome.FindElement(By.Id("INMJ.UmeAdminCompView.MessageArea1")).Text;
                    }
                    catch (Exception) { }

                    return errMsg;
                }
                catch { }

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJJKNEPINJ.UserSearchResultView.userResultTable:1.0"))); } catch { }
                chrome.FindElement(By.Id("INMJJKNEPINJ.UserSearchResultView.userResultTable:1.0")).Click();

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJJKNEPINJ.DisplayUserView.edit"))); } catch { }
                chrome.FindElement(By.Id("INMJJKNEPINJ.DisplayUserView.edit")).Click();
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJJKNEPINJ.ModifyUserView.accountInformation-focus"))); } catch { }
                chrome.FindElement(By.Id("INMJJKNEPINJ.ModifyUserView.accountInformation-focus")).Click();
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJJKNEPINJ.ModifyUserView.isaccountlocked1-img"))); } catch { }
                chrome.FindElement(By.Id("INMJJKNEPINJ.ModifyUserView.isaccountlocked1-img")).Click();
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJJKNEPINJ.ModifyUserView.validTo"))); } catch { }
                chrome.FindElement(By.Id("INMJJKNEPINJ.ModifyUserView.validTo")).Clear();
                chrome.FindElement(By.Id("INMJJKNEPINJ.ModifyUserView.validTo")).SendKeys("6/9/2022");

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("INMJJKNEPINJ.ModifyUserView.save-r"))); } catch { }
                chrome.FindElement(By.Id("INMJJKNEPINJ.ModifyUserView.save-r")).Click();

                void ResetPage()
                {
                    chrome.Manage().Cookies.DeleteAllCookies();
                    Thread.Sleep(1000);
                    chrome.Navigate().GoToUrl(link);
                    Thread.Sleep(1000);
                    chrome.FindElement(By.Id("logonuidfield")).SendKeys(user);
                    chrome.FindElement(By.Id("logonpassfield")).SendKeys(pass);
                    chrome.FindElement(By.Name("uidPasswordLogon")).Click();
                }

                chrome.Quit();
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);

                return "OK";
            }
            catch (Exception ex)
            {
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);

                return ex.Message;
            }
        }


        private sealed class FormatNumbersAsTextConverter : JsonConverter
        {
            public override bool CanRead => false;
            public override bool CanWrite => true;
            public override bool CanConvert(Type type) => type == typeof(double);
            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            {
                double number = (double)value;
                writer.WriteValue(number.ToString(CultureInfo.InvariantCulture));
            }
            public override object ReadJson(JsonReader reader, Type type, object existingValue, JsonSerializer serializer)
            {
                throw new NotSupportedException();
            }
        }
    }
}
