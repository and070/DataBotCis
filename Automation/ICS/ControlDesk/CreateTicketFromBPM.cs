using DataBotV5.Logical.Projects.ControlDesk;
using System.Text.RegularExpressions;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using System;

namespace DataBotV5.Automation.ICS.ControlDesk
{
    internal class CreateTicketFromBPM
    {
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly Credentials cred = new Credentials();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        const string mandCd = "PRD";

        string respFinal = "";

        public void Main()
        {
            mail.GetAttachmentEmail("Solicitudes creacion de SR BPM", "Procesados", "Procesados creacion de SR BPM");
            if (!string.IsNullOrWhiteSpace(root.Email_Body) /*&& root.BDUserCreatedBy.ToLower() == "bpm@mailgbm.com"*/)
            {
                CdTicketData sr = ParseEmail(root.Email_Body);

                cred.SelectCdMand(mandCd);
                string[] srResult = cdi.CreateTicket(sr); //{SR id, SR uid, error}

                if (srResult[2] != "")
                {
                    mail.SendHTMLMail("Error al crear Service Request de BPM<br><br>" + srResult[2].Replace("\\n", "<br>"), new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                }
                else
                {
                    string[] senders = { "calas@gbm.net", "kalonzo@gbm.net", "jrodriguez@gbm.net", "jstevens@gbm.net", "luhernandez@gbm.net", "jmercado@gbm.net" };

                    mail.SendHTMLMail("Se ha generado una nueva tarea en el BPM Validación de diseño para el grupo SSG.\r\nLa tarea quedo radicada bajo el número de SR: " + srResult[0], senders, "SR BPM Validación de diseño: " + srResult[0]);
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Service Request de BPM", srResult[0], root.Subject);
                    respFinal = respFinal + "\\n" + srResult[0];
                    root.requestDetails = respFinal;
                }
                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }
        private CdTicketData ParseEmail(string body)
        {
            string reportedemail, country, classstructureid, commoditygroup, commodity, pluspcustomer, description, longDesc, impact, urgency;

            string GetVal(string field, string text)
            {
                Regex htmlFix = new Regex("[*'\"_&+^><]");
                Regex alphanum = new Regex(@"[^\p{L}0-9 -.@]");
                string[] Separator = new string[] { field };

                text = htmlFix.Replace(text, string.Empty);

                string[] bodySplit = text.Split(Separator, StringSplitOptions.None);
                bodySplit[1] = bodySplit[1].Replace('\r', ' ');
                bodySplit = bodySplit[1].Split('\n');
                string val = alphanum.Replace(bodySplit[0], "").Trim().ToUpper();

                return val;
            }

            #region Leer info correo

            //Cliente:	Claro S.A
            //Número de Opp:	6780976
            //País:	GBMGT
            //Correo del resposable:	atrigueros@gbm.net
            //Service Group:	Infraestructure
            //Service:	TSG
            //Clasificación:	30
            //Impacto y Urgencia:	3
            //El Due Date de esta tarea es de:	48 horas

            string title = body.Split(new string[] { "Información Solicitud" }, 2, StringSplitOptions.None)[0];
            body = body.Split(new string[] { "Información Solicitud" }, 2, StringSplitOptions.None)[1];

            //string bpmCustomer = GetVal("Cliente", body);
            string bpmOpp = GetVal("Número de Opp", body);
            string bpmCountry = GetVal("País", body);
            string bpmEmail = GetVal("Correo del resposable", body);
            string bpmServiceGroup = GetVal("Service Group", body);
            string bpmService = GetVal("Service:", body);
            string bpmClasif = GetVal("Clasificación", body);
            string bpmImpactUrgency = GetVal("Impacto y Urgencia", body);
            string bpmDueDate = GetVal("El Due Date de esta tarea es de", body);

            #endregion

            #region Convertir Valores
            reportedemail = bpmEmail;
            impact = urgency = bpmImpactUrgency;
            commodity = bpmService;

            if (bpmCountry.Length == 5)
            {
                if (bpmCountry.Substring(0, 3) == "GBM")
                    country = bpmCountry.Substring(3, 2);
                else
                    country = "";
            }
            else
                country = "";


            if (bpmClasif == "30")
                classstructureid = "6084"; //QAS 5553
            else
                classstructureid = bpmClasif;

            if (bpmServiceGroup.ToLower() == "infraestructure")//support, 1050802,1051002
                commoditygroup = "Infraestru";
            else
                commoditygroup = bpmServiceGroup;

            switch (country)
            {
                case "PA":
                    pluspcustomer = "0010000811";
                    break;
                case "CR":
                    pluspcustomer = "0010000663";
                    break;
                case "NI":
                    pluspcustomer = "0010000799";
                    break;
                case "HN":
                    pluspcustomer = "0010000735";
                    break;
                case "GT":
                    pluspcustomer = "0010000731";
                    break;
                case "SV":
                    pluspcustomer = "0010000829";
                    break;
                case "DR":
                    pluspcustomer = "0010000681";
                    break;
                default:
                    pluspcustomer = "";
                    break;
            }

            description = title;
            title += "Número de Opp: " + bpmOpp + "\n";
            title += "El Due Date de esta tarea es de: " + bpmDueDate + "\n";
            longDesc = title;

            #endregion

            CdTicketData sr = new CdTicketData
            {
                ReportedEmail = reportedemail,
                Country = country,
                ClassStructureId = classstructureid,
                CommodityGroup = commoditygroup,
                Commodity = commodity,
                PluspCustomer = pluspcustomer,
                Description = description,
                LongDescription = longDesc,
                Impact = impact,
                Urgency = urgency,
                ExternalSystem = "EMAIL",
                TicketType = "SR"
            };

            return sr;
        }
    }

}
