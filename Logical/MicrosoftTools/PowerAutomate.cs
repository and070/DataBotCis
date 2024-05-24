using DataBotV5.Data;
using System.IO;
using System;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;

namespace DataBotV5.Logical.MicrosoftTools
{
    /// <summary>
    /// Clase Logical encargada de Power Automate.
    /// </summary>
    internal class PowerAutomate
    {
        Credentials cred = new Credentials();
        SharePoint sp = new SharePoint();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();

        /// <summary>
        /// Método que inicia una aprobación en Power Automate
        /// </summary>
        /// <param name="apprTitle">Titulo de la aprobación</param>
        /// <param name="approver">Correo del aprobador</param>
        /// <param name="message">El mensaje de la aprobación</param>
        /// <param name="reportedBy">Correo que envía la solicitud</param>
        /// <param name="specificDataJson">Json con algún dato que se quiera guardar en la aprobación</param>
        /// <param name="resSubject">Subject del correo que llegara de respuesta cuando termine el flujo(normalmente seria el proceso)</param>
        /// <param name="attachment">Ruta de un archivo que se quiera adjuntar en la aprobación</param>
        public void SendApproval(string apprTitle, string approver, string message, string reportedBy, string specificDataJson, string resSubject, string attachment = "")
        {
            //contruir el JSON

            string id = DateTime.Now.Ticks.ToString();

            string json = "{";
            json += "\"appr_title\":\"" + apprTitle + "\",";
            json += "\"approver\":\"" + approver + "\",";
            json += "\"message\":\"" + message + "\",";
            json += "\"reported_by\":\"" + reportedBy + "\",";
            json += "\"attachment\":\"" + Path.GetFileName(attachment) + "\",";
            json += "\"subject\":\"appr_request_" + resSubject + "\",";
            json += "\"json_name\":\"" + id + "\",";
            json += "\"specific_data\":" + specificDataJson;
            json += "}";

            string jsonFile = CreateTxt(json, id);
        
            if (attachment != "")
                sp.UploadFileToSharePointV2("https://gbmcorp.sharepoint.com/sites/flowDatabot", "Documentos/Approvals/Pending/" + id + "/", attachment);
            sp.UploadFileToSharePointV2("https://gbmcorp.sharepoint.com/sites/flowDatabot", "Documentos/Approvals/Pending/", jsonFile);
          
            //el proceso sigue en:
            ////https://us.flow.microsoft.com/manage/environments/Default-95028a14-d2f9-43c0-bc73-42a4be7c0492/flows/5939f797-a208-49d8-af1e-05192db9b001/details
        }
        private string CreateTxt(string content, string id)
        {
            string path = root.FilesDownloadPath + "\\" + id + ".json";

            if (!File.Exists(path))
            {
                File.Create(path).Dispose();

                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.WriteLine(content);
                }

            }
            else if (File.Exists(path))
            {
                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.WriteLine(content);
                }
            }

            return path;
        }

        public Dictionary<string, string> GetApprovalRequests(string process)
        {
            return mail.GetApprovalRequests(process);
        }
    }
}
