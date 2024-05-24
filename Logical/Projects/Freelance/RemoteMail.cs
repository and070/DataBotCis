using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using DataBotV5.Logical.Mail;

namespace DataBotV5.Logical.Projects.Freelance
{
    /// <summary>
    /// Clase Logical encargada de correos freelance.
    /// </summary>
    class RemoteMail
    {
        public void SendNotificationFreelancePayment(List<string> email, string subject, string content, string link, [Optional] string table, [Optional] string[] attachments, [Optional] List<string> copy)
        {
            MailInteraction mail = new MailInteraction();
            //mail.SendNotificationPaymentAccountants(email, copy, subject, content, table, attachments, link);

        }
        public void SendNotificationFreelanceEx(string email, string subject, string title, string content, string link, [Optional]string table, [Optional]string[] attachments)
        {
            try
            {
                MailInteraction mail = new MailInteraction();
                //mail.SendNotificationPortalFreelance(email, subject, content, link);

            }
            catch (Exception)
            {


            }
        }
        public void SendNotificationFreelance(string email, string subject, string title, string content, string link, [Optional] string table, [Optional] string[] attachments)
        {
            try
            {
                MailInteraction mail = new MailInteraction();
                //mail.SendNotificationPortalFreelance(email, subject, content, link);

            }
            catch (Exception)
            {


            }
        }
    }
}
