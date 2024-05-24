using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using System.Data;
using System;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DataBotV5.Automation.ICS.TIRequest
{
    internal class ItRequestsNotification
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly Log log = new Log();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();

        readonly string crudDb = "QAS";
        bool executeStats = false;

        public void Main()
        {

            DataTable requests = crud.Select("SELECT * FROM `requests` WHERE `pendingMail` = 1", "it_request", crudDb);

            foreach (DataRow request in requests.Rows)
            {
                DateTime endDate = Convert.ToDateTime(request["endDate"].ToString());
                DateTime actualDate = DateTime.Now;
                string idRequest = request["id"].ToString();

                List<EmailDataItRequest> emailsData = GetEmailData(request);

                foreach (EmailDataItRequest emailData in emailsData)
                {
                    if (emailData.Body != null)
                    {
                        string htmlToPlain = Regex.Replace(emailData.Body, @"<br\s*/?>", Environment.NewLine);
                        htmlToPlain = Regex.Replace(htmlToPlain, @"<[^>]*>", "");
                        console.WriteLine("Enviando correo: " + Environment.NewLine + htmlToPlain);

                        mail.SendHTMLMail(emailData.Body, emailData.Sender, emailData.Subject,  emailData.Cc);
                        log.LogDeCambios("Creacion", root.BDProcess, "databot@gbm.net", "Nueva Notificación de Solicitud de TI", htmlToPlain, emailData.Subject);
                    }

                    executeStats = true;
                }

                crud.Update($"UPDATE `requests` SET `pendingMail` = 0  WHERE `id` = '{idRequest}'", "it_request", crudDb);

                if (executeStats)
                {
                    root.requestDetails = "Nueva Notificación de Solicitud de TI";
                    root.BDUserCreatedBy = "INTERNALCOSTUMERSRVS";

                    console.WriteLine("Creando estadísticas... ");
                    using (Stats stats = new Stats())
                        stats.CreateStat();
                }
            }
        }
        private List<EmailDataItRequest> GetEmailData(DataRow request)
        {
            List<EmailDataItRequest> emailsData = new List<EmailDataItRequest>();

            try
            {
                #region requestData
                string country = request["country"].ToString();
                string createdBy = request["managerApprover"].ToString(); //managerApprover
                string employeeId = request["employeeId"].ToString();
                string idRequest = request["id"].ToString();
                string firstName = request["firstName"].ToString();
                string manager = request["manager"].ToString();
                string jobType = request["jobType"].ToString();
                string lastStep = request["lastStep"].ToString();
                string lastName = request["lastName"].ToString();
                string location = request["location"].ToString();
                string position = request["jobPosition"].ToString();
                string requestState = request["requestState"].ToString();
                string userID = request["userID"].ToString();
                string requestId = request["id"].ToString();
                DateTime endDate = Convert.ToDateTime(request["endDate"].ToString());
                DateTime ssDate = Convert.ToDateTime(request["ssDate"].ToString());
                DateTime asDate = Convert.ToDateTime(request["asDate"].ToString());
                DateTime csDate = Convert.ToDateTime(request["csDate"].ToString());

                DataTable SystemsData = crud.Select($"SELECT * FROM `countries` WHERE `id` = {country}", "it_request", crudDb);
                string countryName = SystemsData.Rows[0]["country"].ToString();

                #region actualStepSenders
                string[] GBCO = new string[] { "ATREJOS@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET", "ROMUNOZ@GBM.NET", "MDELEON@GBM.NET" };
                string[] GBCR = new string[] { "ATREJOS@GBM.NET", "GCARRION@GBM.NET", "POCAMPO@GBM.NET", "GVILLALOBOS@GBM.NET" };
                string[] GBDR = new string[] { "ATREJOS@GBM.NET", "GCARRION@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBGT = new string[] { "ATREJOS@GBM.NET", "GCARRION@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBHN = new string[] { "ATREJOS@GBM.NET", "GCARRION@GBM.NET", "GVILLALOBOS@GBM.NET", "JAHERNANDEZ@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBNI = new string[] { "ATREJOS@GBM.NET", "GCARRION@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBPA = new string[] { "ATREJOS@GBM.NET", "MDELEON@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBSV = new string[] { "ATREJOS@GBM.NET", "GCARRION@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET", "SSALINAS@GBM.NET" };
                string[] GBHQ = new string[] { "ATREJOS@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBVE = new string[] { "ATREJOS@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] GBMI = new string[] { "ATREJOS@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] BVI = new string[] { "ATREJOS@GBM.NET", "GVILLALOBOS@GBM.NET", "POCAMPO@GBM.NET" };
                string[] applicationSupport = new string[] { "INTERNALCUSTOMERSRVS@GBM.NET" };
                string[] serverSupport = new string[] { "MISINFRASTRUCTURE@GBM.NET", "jfalvarado@gbm.net", "davcastillo@gbm.net", "dvillalobos@gbm.net" };
                string[] communicationSupport = new string[] { "UMORA@GBM.NET" };

                if (lastStep == "6" || lastStep == "10" || lastStep == "18" || lastStep == "21" || lastStep == "15")
                    createdBy = request["createdBy"].ToString();
                #endregion

                #endregion

                #region emailStartRequest
                if (lastStep == "1" || lastStep == "6" || lastStep == "10" || lastStep == "15" || lastStep == "18" || lastStep == "21")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = new string[] { createdBy + "@gbm.net" };

                    // ***** Subject *****
                    emailData.Subject = "Inicio de Proceso de TI: " + requestState + " - " + firstName + " " + lastName;

                    // ***** Body *****
                    emailData.Body += "<b> Detalle de Solicitud:</b> El proceso se ha iniciado de forma satisfactoria.<br><br>";
                    emailData.Body += "<b> Nombre del Colaborador: </b>" + firstName + " " + lastName + "<br>";
                    emailData.Body += "<b> Tipo de Solicitud: </b>" + requestState + "<br>";
                    emailData.Body += "<b> País: </b>" + countryName + "<br><br>";
                    if (employeeId != "")
                        emailData.Body += "<b> Id de Colaborador: </b>" + employeeId + "<br><br>";

                    emailsData.Add(emailData);

                }
                #endregion

                #region emailPendingTaskHC
                if (lastStep == "1")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    switch (country)
                    {
                        case "1":
                            emailData.Sender = GBCR;
                            break;
                        case "2":
                            emailData.Sender = GBDR;
                            break;
                        case "3":
                            emailData.Sender = GBGT;
                            break;
                        case "4":
                            emailData.Sender = GBHN;
                            break;
                        case "5":
                            emailData.Sender = GBSV;
                            break;
                        case "6":
                            emailData.Sender = GBPA;
                            break;
                        case "7":
                            emailData.Sender = GBHQ;
                            break;
                        case "8":
                            emailData.Sender = GBVE;
                            break;
                        case "9":
                            emailData.Sender = GBNI;
                            break;
                        case "11":
                            emailData.Sender = BVI;
                            break;
                        case "12":
                            emailData.Sender = GBDR;
                            break;
                        case "13":
                            emailData.Sender = GBCO;
                            break;
                        default:
                            emailData.Sender = new string[] { "internalcustomersrvs@gbm.net" };
                            emailData.Body = $"<b> Error en Notificar Payroll, País: {country}</b>";
                            break;
                    }

                    // ***** Subject *****
                    emailData.Subject = "Inicio de Proceso de TI: Alta - " + firstName + " " + lastName;

                    // ***** Body *****
                    emailData.Body += "<b> La Solicitud del colaborador: </b>" + firstName + " " + lastName + ", ha sido asignada a su unidad.<br><br>";
                    emailData.Body += "<b> Tipo de Solicitud: </b>" + requestState + "<br>";
                    emailData.Body += "<b> País: </b>" + countryName + "<br><br>";
                    if (employeeId != "")
                        emailData.Body += "<b> Id de Colaborador: </b>" + employeeId + "<br><br>";

                    emailsData.Add(emailData);

                }
                #endregion

                #region emailSystemAccesses
                //if (lastStep == "1")
                //{
                //    DataTable systemAccesses = crud.Select($"SELECT * FROM `systemsAccessesrequests` WHERE `requestID` = {requestId}", "it_request", crudDb);

                //    if (systemAccesses.Rows.Count > 0)
                //    {

                //        foreach (DataRow systemAccess in systemAccesses.Rows)
                //        {
                //            EmailDataItRequest emailData = new EmailDataItRequest();
                //            string systemId = systemAccess["systemAccessId"].ToString();
                //            DataTable SystemsData = crud.Select($"SELECT * FROM `systemsAccesses` WHERE `id` = {systemId}", "it_request", crudDb);
                //            string systemType = SystemsData.Rows[0]["accessType"].ToString();

                //            // ***** Subject *****
                //            emailData.Subject = "Notificación de Solicitud de Software Adicional";

                //            // ***** Body *****
                //            emailData.Body += "<div style = 'text-align: center; font-size: 18px;' ><b>Notificación de Solicitud de Software Adicional</b><br><br><br></div>";
                //            emailData.Body += "<b> Colaborador: </b>" + firstName + " " + lastName + "<br><br>";
                //            emailData.Body += "<b> Solicitud de los siguientes softwares adicionales: </b><br>";


                //            switch (SystemsData.Rows[0]["systemAccess"].ToString())
                //            {
                //                case "Power Reserved":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Power Reserved<br>";
                //                    break;
                //                case "BARCODE BOX":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "BARCODE BOX<br>";
                //                    break;
                //                case "IBM INVOICES":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "IBM INVOICES<br>";
                //                    break;
                //                case "Documento ITC/WTC":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Documento ITC/WTC<br>";
                //                    break;
                //                case "PIMMS":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "PIMMS<br>";
                //                    break;
                //                case "Documento MD":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Documento MD<br>";
                //                    break;
                //                case "Sistema de Reclutamiento":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Sistema de Reclutamiento<br>";
                //                    break;
                //                case "Sist. Solicitudes de usuario":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Sist. Solicitudes de usuario<br>";
                //                    break;
                //                case "BPM":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "BPM<br>";
                //                    break;
                //                case "Sistemas Quejas":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Sistemas Quejas<br>";
                //                    break;
                //                case "Sistemas Encuestas":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Sistemas Encuestas<br>";
                //                    break;
                //                case "COGNOS":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "COGNOS<br>";
                //                    break;
                //                case "Línea Teléfono":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Línea Teléfono<br>";
                //                    break;
                //                case "Cisco Online Connections":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Cisco Online Connections<br>";
                //                    break;
                //                case "Extensión":
                //                    emailData.Sender = new string[] {"asd@gbm.net"};
                //                    emailData.Body += "Extensión<br>";
                //                    break;
                //                case "ERP":
                //                    emailData.Sender = new string[] {"internalcustomersrvs@gbm.net"};
                //                    emailData.Body += "ERP<br>";
                //                    break;
                //                case "CRM":
                //                    emailData.Sender = new string[] {"internalcustomersrvs@gbm.net"};
                //                    emailData.Body += "CRM<br>";
                //                    break;
                //                case "PORTAL":
                //                    emailData.Sender = new string[] {"internalcustomersrvs@gbm.net"};
                //                    emailData.Body += "PORTAL<br>";
                //                    break;
                //                default:
                //                    break;
                //            }

                //            emailData.Body += "<br><br>";

                //            emailsData.Add(emailData);
                //        }
                //    }
                //}
                #endregion

                #region emailSyncCollab
                if (lastStep == "2")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = new string[] { "databot@gbm.net" };

                    // ***** Subject *****
                    emailData.Subject = "Sincronización Colaborador";

                    // ***** Body *****
                    emailData.Body += "<b> Nombre Colaborador: </b>" + firstName + " " + lastName + "<br>";
                    emailData.Body += "<b> Número Colaborador: </b>" + employeeId + "<br>";
                    emailsData.Add(emailData);

                }
                #endregion

                #region emailPendingTaskServerSupport
                if (lastStep == "2" || lastStep == "6" || lastStep == "10" || lastStep == "15" || lastStep == "18" || lastStep == "21")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = serverSupport;

                    // ***** Subject *****
                    emailData.Subject = "Notificación de Asignación de Tarea en Proceso de TI";

                    // ***** Body *****
                    emailData.Body += "<b> La Solicitud del colaborador: </b>" + firstName + " " + lastName + ", ha sido asignada a su unidad.<br><br>";
                    emailData.Body += "<b> Tipo de Solicitud: </b>" + requestState + "<br>";
                    emailData.Body += "<b> País: </b>" + countryName + "<br>";

                    emailsData.Add(emailData);

                }
                #endregion

                #region emailRpaNew
                if (lastStep == "3")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = new string[] { "databot@gbm.net" };

                    // ***** Subject *****
                    emailData.Subject = "Nueva Solicitud de TI Notificación RPA";

                    // ***** Body *****
                    emailData.Body = "Nueva Solicitud de TI <br><br>";
                    emailData.Body += "<div style = 'text-align: center; font-size: 18px;' ><b>Información Solicitud</b><br><br><br></div>";
                    emailData.Body += "<b> Nombre Colaborador: </b>" + firstName + "<br>";
                    emailData.Body += "<b> Apellido del Colaborador: </b>" + lastName + "<br>";
                    emailData.Body += "<b> País: </b>" + countryName + "<br>";
                    emailData.Body += "<b> Localidad: </b>" + location + "<br>";
                    emailData.Body += "<b> Posición: </b>" + position + "<br>";
                    emailData.Body += "<b> Fecha de fin de contrato: </b>" + endDate.ToString("dd-MM-yyyy") + "<br>";
                    emailData.Body += "<b> Tipo Solicitud:</b> nuevo<br>";
                    emailData.Body += "<b> Tipo Plaza: </b>" + jobType + "<br>";
                    emailData.Body += "<b> Correo del usuario: </b>" + userID + "@gbm.net<br>";
                    emailData.Body += "<b> Número de Colaborador: </b>" + employeeId + "<br>";
                    emailsData.Add(emailData);

                }
                #endregion

                #region emailPendingTaskApplicationSupport
                if (lastStep == "3" || lastStep == "7" || lastStep == "11" || lastStep == "16" || lastStep == "19" || lastStep == "22")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = applicationSupport;

                    // ***** Subject *****
                    emailData.Subject = "Notificación de Asignación de Tarea en Proceso de TI";

                    // ***** Body *****
                    emailData.Body += "<b> La Solicitud del colaborador: </b>" + firstName + " " + lastName + ", ha sido asignada a su unidad.<br><br>";
                    emailData.Body += "<b> Tipo de Solicitud: </b>" + requestState + "<br>";

                    emailsData.Add(emailData);

                }
                #endregion

                #region emailPendingTaskCommunicationSupport
                if (lastStep == "5" || lastStep == "9" || lastStep == "13")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = communicationSupport;

                    // ***** Subject *****
                    emailData.Subject = "Notificación de Asignación de Tarea en Proceso de TI";

                    // ***** Body *****
                    emailData.Body += "<b> La Solicitud del colaborador: </b>" + firstName + " " + lastName + ", ha sido asignada a su unidad.<br><br>";
                    emailData.Body += "<b> Tipo de Solicitud: </b>" + requestState + "<br>";

                    emailsData.Add(emailData);

                }
                #endregion

                #region emailSystemAccesses
                if (lastStep == "5")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    DataTable requestsSystemAcc = crud.Select($"SELECT * FROM `applicationAccessesRequests` WHERE `requestId` = '{idRequest}'", "it_request", crudDb);

                    if (requestsSystemAcc.Rows.Count > 0)
                    {
                        // ***** Sender *****
                        emailData.Sender = new string[] { userID + "@gbm.net" };

                        // ***** cc *****
                        emailData.Cc = new string[] { createdBy + "@gbm.net" };

                        // ***** Subject *****
                        emailData.Subject = "Notificación para accesos a los Sistemas: ID " + employeeId + " " + firstName + " " + lastName;

                        // ***** Body *****
                        emailData.Body += "<b><u>Acceso(s) a Sistema(s):</u></b><br><br>";
                        emailData.Body += "Estos son los accesos de los sistemas a los cuales puede acceder:<br><br>";
                        emailData.Body += "<b>User ID: </b>" + userID + "<br><br>"; // usar sapUser

                        foreach (DataRow requestsSystemAccRow in requestsSystemAcc.Rows)
                        {
                            string systemId = requestsSystemAccRow["systemApplicationId"].ToString();

                            switch (systemId)
                            {
                                case "1":
                                    emailData.Body += "<b> ERP:</b><br> Contactar al soporte local para configurar su acceso a ERP," +
                                            " después de esto no será necesario escribir su contraseña.<br><br>";
                                    break;
                                case "2":
                                    emailData.Body += "<b> CRM:</b><br> Contactar al soporte local para configurar su acceso a CRM," +
                                                  " después de esto no será necesario escribir su contraseña.<br><br>";
                                    break;
                                case "3":
                                    emailData.Body += "<b> Portal:</b><br> <a>http://ep-prod-app.gbm.net:50100/irj/portal</a><br>" +
                                                  "Utilizar usuario y contraseña de dominio (acceso a Windows).<br><br>";
                                    break;
                                case "4":
                                    emailData.Body += "<b> Control Desk:</b><br> <a>https://controldesk.gbm.net/customers</a><br>" +
                                                 "Utilizar dirección de correo como usuario (incluyendo @gbm.net)  y contraseña de dominio (acceso a Windows).<br><br>";
                                    break;
                                default:
                                    break;
                            }
                        }
                        emailData.Body += "<br>Gracias por su atención<br>";
                        emailsData.Add(emailData);
                    }

                }
                #endregion

                #region emailEndRequest
                if (lastStep == "4")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = new string[] { createdBy + "@gbm.net" };

                    // ***** Subject *****
                    emailData.Subject = $"Finalización del Proceso de TI: Alta - {firstName} {lastName}";

                    // ***** Body *****
                    emailData.Body += "<b>Detalle Solicitud:</b> Se informa que la solicitud de TI ha sido procesada y se ha completado con satisfacción.<br>";
                    emailData.Body += "<b> Nombre: </b>" + firstName + "<br>";
                    emailData.Body += "<b> Apellidos: </b>" + lastName + "<br>";
                    emailData.Body += "<b> Id Colaborador SAP: </b>" + employeeId + "<br>";
                    emailData.Body += "<b> Tipo de Solicitud: </b>" + requestState + "<br>";
                    emailData.Body += "<b> País: </b>" + countryName + "<br><br>";

                    emailsData.Add(emailData);
                }
                #endregion

                #region notif de baja
                if (lastStep == "8")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();
                    // ***** Sender *****
                    emailData.Sender = new string[] { "activospa@gbm.net" };

                    // ***** Subject *****
                    emailData.Subject = $"Notificación para baja del Colaborador: {firstName} {lastName}";

                    // ***** Body *****
                    emailData.Body += "<b>Detalle baja de colaborador <br>";
                    emailData.Body += "<b> Nombre del Colaborador: </b>" + firstName + "<br>";
                    emailData.Body += "<b> Apellidos del Colaborador: </b>" + lastName + "<br>";
                    emailData.Body += "<b> País del Colaborador: </b>" + countryName + "<br><br>";
                    emailData.Body += "<b> Manager del colaborador: </b>" + manager + "<br>";
                    emailData.Body += "Se dio de baja al colaborador. ";

                    emailsData.Add(emailData);
                }
                #endregion

                #region emailTerminationCollab
                if (lastStep == "6" || lastStep == "7")
                {
                    EmailDataItRequest emailData = new EmailDataItRequest();

                    // ***** Sender *****
                    emailData.Sender = new string[] { "databot@gbm.net" };
                    // ***** cc *****
                    emailData.Cc = new string[] { "internalcustomersrvs@gbm.net" };

                    // ***** Subject *****
                    emailData.Subject = "Nueva Solicitud de TI Notificación RPA";

                    // ***** Body *****
                    emailData.Body = "Nueva Solicitud de TI <br><br>";
                    emailData.Body += "<div style = 'text-align: center; font-size: 18px;' ><b>Información Solicitud</b><br><br><br></div>";
                    emailData.Body += "<b> Nombre del Colaborador </b>: " + firstName + "<br>";
                    emailData.Body += "<b> Apellido del Colaborador </b>: " + lastName + "<br>";
                    emailData.Body += "<b> País</b>: " + country + "<br>";
                    emailData.Body += "<b> Localidad: </b>" + location + "<br>";
                    emailData.Body += "<b> Posición: </b>" + position + "<br>";
                    emailData.Body += "<b> Fecha de fin de contrato: </b>" + endDate.ToString("dd-MM-yyyy") + "<br>";
                    if (lastStep == "6") { emailData.Body += "<b> Tipo Solicitud</b>: Baja<br>"; }
                    else if (lastStep == "7") { emailData.Body += "<b> Tipo Solicitud</b>: BajaICS<br>"; }
                    emailData.Body += "<b> Tipo Plaza: </b>" + jobType + "<br>";
                    emailData.Body += "<b> Correo del usuario: </b>" + userID + "@gbm.net<br>";
                    emailData.Body += "<b> Número de Colaborador: </b>" + employeeId + "<br>";

                    emailsData.Add(emailData);
                }
                #endregion

            }
            catch (Exception)
            {
                #region emailError
                EmailDataItRequest emailData = new EmailDataItRequest();

                emailData.Sender = new string[] { "internalcustomersrvs@gbm.net" };
                emailData.Subject = "Error al enviar notificación de Solicitud de TI";
                emailData.Body = $"No se pudo entregar la notificación del paso {request["lastStep"]} de la gestión {request["Id"]}";
                emailData.Cc = new string[] { "smarin@gbm.net" };
                emailsData.Add(emailData);
                #endregion
                throw;
            }

            return emailsData;
        }
        class EmailDataItRequest
        {
            public string Subject { get; set; }
            public string Body { get; set; }
            public string[] Sender { get; set; }
            public string[] Cc { get; set; }
        }
    }
}
