using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Projects.MasterData;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.Contacts;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;
using DataBotV5.Logical.MicrosoftTools;
using System.Data;

namespace DataBotV5.Automation.DM.Contacts
{
    /// <summary><c>ContactsCreation:</c> 
    /// Clase DM Automation encargada de creación de contactos en datos maestros.</summary>
    class ContactsCreationSS
    {
        #region Global Variables
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        WebexTeams wt = new WebexTeams();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ContactSapSS cSap = new ContactSapSS();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        WebInteraction web = new WebInteraction();
        Stats estadisticas = new Stats();
        MsExcel ms = new MsExcel();

        public string resFailure = "";
        bool validateData = true;

        string customer = "";
        string contactTitle = "";
        string contactName = "";
        string contactLastNames = "";
        string contactCountry = "";
        string contactAdress = "";
        string contactEmail = "";
        string contactPhone = "";
        string contactJob = "";
        string department = "";
        string language = "";
        string validate = "";
        string retMsg = "";
        string fmRep = "";
        string res1 = "", res2 = "";

        int erpMand = 260, rows = 0, start_row = 0;

        string respFinal = "";

        #endregion

        public void Main()
        {
            string respuesta = DM.GetManagement("3"); //Contactos en la tabla MotherTable de la DB masterData
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessContacts();
                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }

        /// <summary>Metodo para procesar contactos.</summary>
        public void ProcessContacts()
        {
            //no se extraen los Datos Generales ya que no se necesita el pais que envia la solicitud
            try
            {
                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                string countryRequest = root.factorDM;
                //Por cada adjunto de la solicitud
                if (root.metodoDM == "1") //lineal
                {
                    JArray gestiones = JArray.Parse(root.requestDetails);
                    for (int i = 0; i < gestiones.Count; i++)
                    {
                        JObject fila = JObject.Parse(gestiones[i].ToString());
                        ContactInfoSS contactInfo = new ContactInfoSS
                        {
                            cliente = fila["clientId"].Value<string>(),
                            tratamiento = fila["contactTreatmentCode"].Value<string>(),
                            nombre = fila["firstName"].Value<string>().Trim().ToUpper(),
                            apellido = fila["lastName"].Value<string>().Trim().ToUpper(),
                            pais = fila["countryCode"].Value<string>().Trim().ToUpper(),
                            direccion = fila["address"].Value<string>().Trim().ToUpper(),
                            email = fila["email"].Value<string>().Trim().ToLower()
                        };

                        List<phonesSS> phones = new List<phonesSS>();
                        phonesSS phone = new phonesSS
                        {
                            TELEPHONE = fila["phone"].Value<string>().Trim().ToLower(),
                            MOBILE = "",
                            EXT = ""
                        };
                        phones.Add(phone);
                        contactInfo.telefonos = phones;

                        contactInfo.idioma = fila["languageCode"].Value<string>().Trim().ToUpper();
                        contactInfo.puesto = fila["positionCode"].Value<string>().Trim();
                        try
                        { contactInfo.departamento = fila["DEPARTAMENTO"].Value<string>().Trim(); }
                        catch (Exception)
                        { contactInfo.departamento = ""; }

                        string resp = cSap.CreateContactSAP(contactInfo);
                        if (resp.Contains("Error"))
                            validateData = false;

                        res1 = res1 + resp + "<br>";
                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Contacto", resp, root.Subject);
                        respFinal = respFinal + "\\n" + "Crear Contacto: " + resp;

                    }
                }
                else //MASIVO
                {
                    string attach = root.ExcelFile; //ya viene 
                    if (!String.IsNullOrEmpty(attach))
                    {
                        #region abrir excel
                        console.WriteLine("Abriendo excel y validando");

                        DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + attach);
                        #endregion


                        foreach (DataRow row in xlWorkSheet.Rows)
                        {
                            ContactInfoSS contactInfo = new ContactInfoSS();
                            customer = row["xID Cliente"].ToString().Trim();
                            contactInfo.cliente = customer;
                            if (customer != "")
                            {
                                contactInfo.tratamiento = row["Tratamiento"].ToString().Trim();
                                contactInfo.nombre = row["Nombre"].ToString().Trim().ToUpper();
                                contactInfo.apellido = row["Apellido"].ToString().Trim().ToUpper();
                                contactInfo.pais = row["Pais"].ToString().Trim().ToUpper();
                                contactInfo.direccion = row["Direccion"].ToString().Trim().ToUpper();
                                contactInfo.email = row["Correo"].ToString().Trim().ToLower();
                                List<phonesSS> phones = new List<phonesSS>();
                                phonesSS phone = new phonesSS
                                {
                                    TELEPHONE = row["Telefono"].ToString().Trim().ToLower(),
                                    MOBILE = "",
                                    EXT = ""
                                };
                                phones.Add(phone);
                                contactInfo.telefonos = phones;
                                contactInfo.idioma = row["Idioma"].ToString().Trim().ToUpper();
                                contactInfo.puesto = row["Puesto"].ToString().Trim();
                                contactInfo.departamento = "";
                           
                                
                                string resp = cSap.CreateContactSAP(contactInfo);
                                if (resp.Contains("Error"))
                                    validateData = false;

                                res1 = res1 + resp + "<br>";
                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Contacto", resp, root.Subject);
                                respFinal = respFinal + "\\n" + "Crear Contacto: " + resp;
                            }

                        } //for de cada fila del excel

                       
                    }
                }

                console.WriteLine("Finalizando solicitud");
                if (validateData == false)
                {
                    //enviar email de repuesta de error a datos maestros
                    DM.ChangeStateDM(root.IdGestionDM, res1 + "<br>" + res2, "4");
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + res1 + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
                }
                else
                {
                    //finalizar solicitud
                    DM.ChangeStateDM(root.IdGestionDM, res1, "3");
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Contactos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + res1);
                }

                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                string[] cc = { "dmeza@gbm.net" };
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4");
                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }
    }
}
