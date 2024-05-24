using System.Text.RegularExpressions;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.TIRequest
{
    internal class OldRequestsNotification
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();

        readonly string[] threeDaysNotifs = { "smarin@gbm.net", "kasanchez@gbm.net" };
        readonly string[] misEscalationNotifs = { "groyo@gbm.net", "ahernandez@gbm.net" };
        readonly string[] errorNotifs = { "smarin@gbm.net" };

        public void Main()
        {
            #region Notif. mas de 3 dias
            NotifLastThreeDaysPendingRequests();
            #endregion

            #region Notif. Escalación MIS
            NotifEscalationMisPendingRequests();
            #endregion
        }
        private void NotifEscalationMisPendingRequests()
        {
            string respFinal = "";

            DataTable requestsMis = crud.Select("SELECT * FROM `requests` WHERE `isFinished` = 0 AND DATEDIFF(NOW(), `initialDate`) >= 12", "it_request"); //trae solicitudes con 12 días naturales

            foreach (DataRow requestMis in requestsMis.Rows)
            {
                EmailDataItRequest emailData = GetEmailData(requestMis);
                if (emailData.Body != null)
                {
                    mail.SendHTMLMail(emailData.Body, emailData.Sender, emailData.Subject);
                    string htmlToPlain = Regex.Replace(emailData.Body, @"<br\s*/?>", Environment.NewLine);
                    htmlToPlain = Regex.Replace(htmlToPlain, @"<[^>]*>", "");
                    log.LogDeCambios("Creacion", root.BDProcess, "databot@gbm.net", "Escalación de Solicitud de TI", htmlToPlain, emailData.Subject);

                    respFinal = respFinal + "\\n" + "Nueva Escalación de Solicitud de TI";
                    CreateStat(respFinal);
                }
            }
        }
        private void NotifLastThreeDaysPendingRequests()
        {
            string respFinal = "";
            DataTable requests = crud.Select("SELECT * FROM `pending`", "ti_requests_db");

            foreach (DataRow request in requests.Rows)
            {
                DateTime actualDate = DateTime.Now;
                DateTime reqDate = Convert.ToDateTime(request["date"].ToString());

                if ((actualDate - reqDate).Days > 2)// si la fecha es mayor de 3 días, notificar
                {
                    string body = "La siguiente solicitud tiene mas de 3 días, por favor revisar<br><br>";
                    body += request["emailBody"].ToString().Replace("\n", "<br>");
                    mail.SendHTMLMail(body, new string[] { "internalcustomersrvs@gbm.net" }, "Nueva Solicitud de TI sin Procesar Notificación RPA", threeDaysNotifs);

                    body = Regex.Replace(body, @"<br\s*/?>", Environment.NewLine);
                    body = Regex.Replace(body, @"<[^>]*>", "");

                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Nueva Solicitud de TI sin Procesar Notificación RPA", body, root.Subject);

                    respFinal = respFinal + "\\n" + "Nueva Solicitud de TI sin Procesar Notificación RPA: " + body;

                    CreateStat(respFinal);
                }
            }
        }
        private EmailDataItRequest GetEmailData(DataRow request)
        {
            EmailDataItRequest emailDataItRequest = new EmailDataItRequest();

            try
            {
                #region requestData
                DateTime initialDate = Convert.ToDateTime(request["initialDate"].ToString());
                TimeSpan timeSpan = DateTime.Now - initialDate;
                int totalDays = (int)Math.Ceiling(timeSpan.TotalDays);
                int workingDays = 0;
                DateTime ssDate = Convert.ToDateTime(request["ssDate"].ToString());
                DateTime asDate = Convert.ToDateTime(request["asDate"].ToString());
                DateTime csDate = Convert.ToDateTime(request["csDate"].ToString());
                string firstName = request["firstName"].ToString();
                string lastName = request["lastName"].ToString();
                string requestState = request["requestState"].ToString();
                #endregion

                #region workingDays
                for (int i = 0; i < totalDays; i++)
                {
                    DateTime currentDate = initialDate.AddDays(i);
                    if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
                        workingDays++; // suma sólo de Lunes a Viernes
                }
                #endregion

                #region emailData
                if (workingDays >= 12)
                {
                    // ***** Sender *****
                    emailDataItRequest.Sender = misEscalationNotifs;

                    // ***** Subject *****
                    emailDataItRequest.Subject = $"Notificación de Escalación {workingDays} días hábiles";

                    // ***** Body *****
                    emailDataItRequest.Body = "<div style='text-align: center; font-size: 18px;'><b>Según el tiempo estimado para esta solicitud en el proceso de Active Directory, ICS </b><br>" +
                    "<b>y Communication Support tiene más de 12 días de iniciada, por favor tomar las medidas del caso</b><br><br><br></div>";

                    emailDataItRequest.Body += "<b> Solicitud de: </b>" + requestState + "<br>";
                    emailDataItRequest.Body += "<b> Nombre Colaborador: </b>" + firstName + "<br>";
                    emailDataItRequest.Body += "<b> Apellidos: </b>" + lastName + "<br><br>";

                    if (ssDate.ToString("dd-MM-yyyy") == "01-01-1000")
                        emailDataItRequest.Body += "<b> Accesos Active Directory: </b>Incompleto<br>";
                    else if (ssDate.ToString("dd-MM-yyyy") != "01-01-1000")
                        emailDataItRequest.Body += "<b> Accesos Active Directoy: </b>Finalizado<br>";

                    if (asDate.ToString("dd-MM-yyyy") == "01-01-1000")
                        emailDataItRequest.Body += "<b> Accesos ICS: </b>Incompleto<br>";
                    else if (asDate.ToString("dd-MM-yyyy") != "01-01-1000")
                        emailDataItRequest.Body += "<b> Accesos ICS: </b>Finalizado<br>";

                    if (csDate.ToString("dd-MM-yyyy") == "01-01-1000")
                        emailDataItRequest.Body += "<b> Accesos Communication Support: </b>Incompleto<br>";
                    else if (csDate.ToString("dd-MM-yyyy") != "01-01-1000")
                        emailDataItRequest.Body += "<b> Accesos Communication Support: </b>Finalizado<br>";

                }
                #endregion

            }
            catch (Exception)
            {
                #region emailError
                emailDataItRequest.Sender = new string[] { "internalcustomersrvs@gbm.net" };
                emailDataItRequest.Subject = "Error al enviar notificación de Solicitud de TI";
                emailDataItRequest.Body = $"No se pudo entregar la notificación de escalación de la gestión {request["Id"]}";
                emailDataItRequest.Cc = errorNotifs;
                #endregion
                throw;
            }
            return emailDataItRequest;
        }
        private void CreateStat(string respFinal)
        {
            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "INTERNALCOSTUMERSRVS";

            console.WriteLine("Creando estadísticas... ");
            using (Stats stats = new Stats())
                stats.CreateStat();
        }
        private class EmailDataItRequest
        {
            public string Subject { get; set; }
            public string Body { get; set; }
            public string[] Sender { get; set; }
            public string[] Cc { get; set; }
        }
    }
}
