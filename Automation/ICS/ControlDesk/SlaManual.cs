using DataBotV5.Logical.Projects.ControlDesk;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using System.Data;
using System;
using System.Security.AccessControl;
using Microsoft.Graph;
using static DataBotV5.Automation.ICS.ControlDesk.ResponsePlans;
using DataBotV5.Data.Process;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SharePoint.Client;
using System.Linq;

namespace DataBotV5.Automation.ICS.ControlDesk
{
    internal class SlaManual
    {
        readonly ControlDeskAprovalContract cdSelenium = new ControlDeskAprovalContract();
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();
        readonly ResponsePlans rp = new ResponsePlans();

        public void Main()
        {
            mail.GetAttachmentEmail("Solicitudes SLA", "Procesados", "Procesados SLA");
            if (root.filesList != null)
            {
                if (root.filesList.Length > 0)
                {
                    foreach (string excelFile in root.filesList)
                    {
                        string filePath = root.FilesDownloadPath + "\\" + excelFile;
                        DataSet excelDts = excel.GetExcelBook(filePath);
                        if (excelDts != null)
                        {
                            ProcessTemplate2(excelDts, filePath);
                        }
                    }
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }
        private void ProcessTemplate2(DataSet excelDts, string filePath)
        {
            //posiciones en la plantilla
            const int mandRow = 1;
            const int objectNameRow = 3;
            const int saNumRow = 5;
            const int descriptionRow = 6;
            const int rankingRow = 7;
            const int intPriorityEvalRow = 8;
            const int intPriorityValueRow = 9;
            const int calcCalendarRow = 10;
            const int conditionRow = 11;
            const int responseRow = 19;
            const int solutionRow = 20;

            int[] servicesRows = { 12, 18 }, responseRows = { 28, 37 }, resolutionRows = { 44, 53 }, genDataCols = { 3, 6 }, escalationCols = { 3, 7 };

            string ticketEmailRes = "";
            DataTable infoDt = excelDts.Tables[0];
            List<string> templates = new List<string>();

            ResponseRP responseRP = new ResponseRP();

            DataTable emailRes = new DataTable();
            emailRes.Columns.Add("SLA");
            emailRes.Columns.Add("Respuesta");

            string agreementForTicket = infoDt.Rows[1][5].ToString();
            string customerForTicket = infoDt.Rows[2][5].ToString();
            string commodityForTicket = infoDt.Rows[3][5].ToString();

            string mandCd = infoDt.Rows[mandRow][genDataCols[0]].ToString();

            int cont = 0;

            bool sendIcs = false;

            for (int slaColCont = genDataCols[0]; slaColCont <= genDataCols[1]; slaColCont++) //for de las Prioridades
            {
                bool hasResponse = false, hasResolution = false;

                CdSlaData sla = new CdSlaData();
                CdSlaCommitments responseCommitment = new CdSlaCommitments();
                CdSlaCommitments resolutionCommitment = new CdSlaCommitments();

                sla.PluspApplServCommodity = new List<string>();
                sla.CdSlaEscalations = new List<CdSlaEscalation>();
                sla.CdSlaCommitments = new List<CdSlaCommitments>();

                string saNum = infoDt.Rows[saNumRow][slaColCont].ToString();

                //crear los SLAs
                string objectName = infoDt.Rows[objectNameRow][genDataCols[0]].ToString();
                string description = infoDt.Rows[descriptionRow][slaColCont].ToString();
                string CalcCalendar = infoDt.Rows[calcCalendarRow][slaColCont].ToString();

                if (saNum != "")
                {
                    #region SLA DATOS GENERALES

                    sla.Sanum = saNum;
                    sla.ObjectName = objectName;
                    sla.Description = description;

                    sla.Ranking = infoDt.Rows[rankingRow][slaColCont].ToString();
                    sla.Condition = infoDt.Rows[conditionRow][slaColCont].ToString();

                    sla.IntPriorityEval = infoDt.Rows[intPriorityEvalRow][slaColCont].ToString();
                    sla.IntPriorityValue = infoDt.Rows[intPriorityValueRow][slaColCont].ToString();
                    sla.CalcCalendar = CalcCalendar;
                    if (CalcCalendar != "24x7")
                    {
                        sla.CalcOrgId = "GBM";
                        if (CalcCalendar == "GBMDO20")
                            sla.CalcShift = "DO8X5";
                        else if (CalcCalendar == "GBMPA20")
                            sla.CalcShift = "PA8X5";
                        else if (CalcCalendar == "GBMUS20")
                            sla.CalcShift = "MIA8X5";
                        else
                            sla.CalcShift = "CA8X5";
                    }
                    else
                    {
                        sla.CalcCalendar = "";
                        sla.CalcShift = "";
                        sla.CalcOrgId = "";
                    }

                    #endregion

                    #region SERVICIOS
                    for (int i = servicesRows[0]; i < servicesRows[1]; i++)
                        if (infoDt.Rows[i][genDataCols[0]].ToString() != "")
                            sla.PluspApplServCommodity.Add(infoDt.Rows[i][genDataCols[0]].ToString());
                    #endregion

                    #region COMMITMENTS
                    if (infoDt.Rows[solutionRow][escalationCols[0] + cont].ToString() != "")
                    {
                        hasResolution = true;
                        resolutionCommitment.Description = "Tiempo de Solucion";
                        resolutionCommitment.Type = "RESOLUTION";
                        resolutionCommitment.Value = infoDt.Rows[solutionRow][escalationCols[0] + cont].ToString();
                        resolutionCommitment.UnitOfMeasure = "HOURS";
                        sla.CdSlaCommitments.Add(resolutionCommitment);
                    }

                    if (infoDt.Rows[responseRow][escalationCols[0] + cont].ToString() != "")
                    {
                        hasResponse = true;
                        responseCommitment.Description = "Tiempo de Respuesta";
                        responseCommitment.Type = "RESPONSE";
                        responseCommitment.Value = infoDt.Rows[responseRow][escalationCols[0] + cont].ToString();
                        responseCommitment.UnitOfMeasure = "HOURS";
                        sla.CdSlaCommitments.Add(responseCommitment);
                    }

                    #endregion

                    #region ESCALATIONS

                    float[] intervals = { 0.25f, 0.5f, 0.75f, 1, 3 };

                    //response
                    if (hasResponse)
                    {
                        int intervalCount = 0;
                        for (int escalationColumnCount = escalationCols[0]; escalationColumnCount <= escalationCols[1]; escalationColumnCount++) //columnas 
                        {
                            CdSlaEscalation slaEscalation = new CdSlaEscalation { Notifications = new List<CdCommTemplate>() };
                            CdCommTemplate commTemplate = new CdCommTemplate { CommTmpltSendToValue = new List<string>() };

                            if (infoDt.Rows[responseRows[0]][escalationColumnCount].ToString() != "")
                            {
                                string timeInterval = ((float.Parse(responseCommitment.Value) * intervals[intervalCount]) - float.Parse(responseCommitment.Value)).ToString();
                                intervalCount++;

                                slaEscalation.TimeAttribute = "ADJUSTEDTARGETRESPONSETIME";
                                slaEscalation.TimeInterval = timeInterval;
                                slaEscalation.IntervalUnit = "HOURS";
                                slaEscalation.Condition = "ACTUALSTART is null and status in ('NEW','QUEUED')";

                                commTemplate.ObjectName = objectName;
                                commTemplate.Description = GetEscPercentage(timeInterval, float.Parse(responseCommitment.Value)) + " R " + description;

                                for (int i = responseRows[0]; i < responseRows[1]; i++)
                                    if (infoDt.Rows[i][escalationColumnCount].ToString().Trim() != "")
                                        commTemplate.CommTmpltSendToValue.Add(infoDt.Rows[i][escalationColumnCount].ToString().Trim());

                                slaEscalation.Notifications.Add(commTemplate);
                                sla.CdSlaEscalations.Add(slaEscalation);
                            }
                        }
                    }

                    //resolution
                    if (hasResolution)
                    {
                        int intervalCount = 0;
                        for (int escalationColumnCount = escalationCols[0]; escalationColumnCount <= escalationCols[1]; escalationColumnCount++) //columnas 
                        {
                            CdSlaEscalation slaEscalation = new CdSlaEscalation { Notifications = new List<CdCommTemplate>() };
                            CdCommTemplate commTemplate = new CdCommTemplate { CommTmpltSendToValue = new List<string>() };

                            if (infoDt.Rows[resolutionRows[0]][escalationColumnCount].ToString() != "")
                            {
                                string timeInterval = ((float.Parse(resolutionCommitment.Value) * intervals[intervalCount]) - float.Parse(resolutionCommitment.Value)).ToString();
                                intervalCount++;

                                slaEscalation.TimeAttribute = "ADJUSTEDTARGETRESOLUTIONTIME";
                                slaEscalation.TimeInterval = timeInterval;
                                slaEscalation.IntervalUnit = "HOURS";
                                slaEscalation.Condition = "ACTUALFINISH is null and status not in ('RESOLVED') and status not in (select value from synonymdomain where maxvalue='SLAHOLD')";

                                commTemplate.ObjectName = objectName;
                                commTemplate.Description = GetEscPercentage(timeInterval, float.Parse(resolutionCommitment.Value)) + " S " + description;

                                for (int i = resolutionRows[0]; i < resolutionRows[1]; i++)
                                    if (infoDt.Rows[i][escalationColumnCount].ToString().Trim() != "")
                                        commTemplate.CommTmpltSendToValue.Add(infoDt.Rows[i][escalationColumnCount].ToString().Trim());

                                slaEscalation.Notifications.Add(commTemplate);
                                sla.CdSlaEscalations.Add(slaEscalation);
                            }
                        }
                    }

                    #endregion

                    string[] slaRes;

                    if (cont == 0)//SLA creando templates
                    {
                        slaRes = ProcessSlas(sla, mandCd);
                        //tomar id de los comm templ creados
                        foreach (CdSlaEscalation escalation in sla.CdSlaEscalations)
                            foreach (CdCommTemplate noti in escalation.Notifications)
                                templates.Add(noti.TemplateId);

                    }
                    else//SLA sin crear templates
                    {
                        //colocar los comm templ creados anteriormente.
                        int i = 0;
                        foreach (CdSlaEscalation escalation in sla.CdSlaEscalations)
                            foreach (CdCommTemplate noti in escalation.Notifications)
                            {
                                try { noti.TemplateId = templates[i]; }
                                catch (Exception) { }
                                i++;
                            }
                        slaRes = ProcessSlas(sla, mandCd, false);

                    }
                    if (slaRes[1].ToUpper().Contains("ERROR"))
                        sendIcs = true;

                    DataRow emailResRow = emailRes.NewRow();
                    emailResRow[0] = slaRes[0];
                    emailResRow[1] = slaRes[1];
                    emailRes.Rows.Add(emailResRow);
                }

                #region Crear Response Plan
                if (sla.Sanum != null)
                {
                    DataTable excelRpsDt = excelDts.Tables["Response Plans"];
                    excelRpsDt.Columns.Add("Mandante");
                    excelRpsDt.Columns.Add("Aplicación");
                    excelRpsDt.Columns.Add("Acción");
                    excelRpsDt.Columns.Add("ID Response Plan");
                    excelRpsDt.Columns.Add("Descripción");
                    excelRpsDt.Columns.Add("Cliente");
                    excelRpsDt.Columns.Add("Condition");
                    excelRpsDt.Columns.Add("Ranking");

                    for (int i = 0; i < sla.PluspApplServCommodity.Count; i++)
                        excelRpsDt.Columns.Add("Servicio" + i);

                    excelRpsDt.Columns.Add("Clasificación");
                    excelRpsDt.Columns.Add("NOC 4.0");
                    excelRpsDt.Columns.Add("Configuration Item (CI)");



                    foreach (DataRow excelRpRow in excelRpsDt.Rows)
                    {
                        string auto = excelRpRow["Auto Asignación"].ToString();
                        string schedule = excelRpRow["Horario"].ToString();

                        for (int i = 0; i < sla.PluspApplServCommodity.Count; i++)
                            excelRpRow["Servicio" + i] = sla.PluspApplServCommodity[i].ToString();


                        string groupRes = excelRpRow["Grupo resolutor"].ToString();
                        string classStructureId = excelRpRow["Class Structure Id"].ToString();

                        excelRpRow["Mandante"] = mandCd;
                        excelRpRow["Aplicación"] = sla.ObjectName;
                        excelRpRow["Acción"] = "Crear";
                        excelRpRow["Descripción"] = GetResponsePlanDescription(agreementForTicket, customerForTicket, objectName, auto, schedule);
                        excelRpRow["Cliente"] = customerForTicket;
                        excelRpRow["Condition"] = sla.Condition;
                        excelRpRow["Ranking"] = excelRpRow["Horario"].ToString() == "No Hábil" ? "2" : "1";
                        excelRpRow["Configuration Item (CI)"] = "";
                        excelRpRow["Clasificación"] = "";
                        excelRpRow["NOC 4.0"] = "No";
                    }

                    responseRP = rp.ProcessResponsePlanICS(excelRpsDt);

                }



                #endregion
                //crear los Tickets

                if (mandCd == "DEV" && sla.Sanum != null)
                {
                    CdTicketData ticket = new CdTicketData
                    {
                        TicketType = sla.ObjectName,
                        ReportedEmail = "ROFERNANDEZ@GBM.NET",//root.BDUserCreatedBy,
                        PluspCustomer = customerForTicket,
                        Description = customerForTicket + " " + agreementForTicket + " " + sla.PluspApplServCommodity[0],
                        LongDescription = customerForTicket + " " + agreementForTicket + " " + sla.PluspApplServCommodity[0],
                        GbmPluspAgreement = agreementForTicket,
                        ExternalSystem = "SAM",
                        Country = "CR"
                    };

                    if (sla.PluspApplServCommodity.Count > 0)
                        ticket.Commodity = sla.PluspApplServCommodity[0];
                    else
                        ticket.Commodity = commodityForTicket;

                    if (sla.ObjectName == "SR")
                        ticket.ClassStructureId = "6084";
                    else
                        ticket.ClassStructureId = "4792";

                    switch (sla.IntPriorityValue)
                    {
                        case "1":
                            ticket.Impact = "1";
                            ticket.Urgency = "1";
                            break;
                        case "2":
                            ticket.Impact = "3";
                            ticket.Urgency = "1";
                            break;
                        case "3":
                            ticket.Impact = "3";
                            ticket.Urgency = "3";
                            break;
                        case "4":
                            ticket.Impact = "3";
                            ticket.Urgency = "5";
                            break;
                    }

                    string[] ticketRes = cdi.CreateTicket(ticket);  //{SR id, SR uid, error}

                    ticketEmailRes = "<b>SLA:</b> " + sla.Sanum + " <b>SR:</b> " + ticketRes[0] + " <b>ERROR:</b> " + ticketRes[2] + "<br>";
                }

                cont++;
            }


            //notificaciones
            console.WriteLine("Enviando correos");
            if (sendIcs)
                mail.SendHTMLMail(val.ConvertDataTableToHTML(emailRes), new string[] { "internalcustomersrvs@gbm.net" }, root.BDProcess, attachments: new string[] { filePath });
            else
                mail.SendHTMLMail(val.ConvertDataTableToHTML(emailRes) + "<br><br>" + ticketEmailRes + "<br><br>Resultado creación de Response Plans<br>" + val.ConvertDataTableToHTML(responseRP.ResponseDt), new string[] { root.BDUserCreatedBy }, root.Subject, attachments: new string[] { filePath });
        }
        private string GetResponsePlanDescription(string agreementForTicket, string customerForTicket, string objectName, string auto, string schedule)
        {
            string customerName = cdi.GetCustomerName(customerForTicket);

            string description = "SLA " + string.Concat(objectName.Take(2)) + " - " + customerName + agreementForTicket + " - Horario: " + schedule;

            if (auto == "Sí")
                description += " - Auto";

            return description;
        }        
        private string[] ProcessSlas(CdSlaData sla, string mandCd, bool createCommTemp = true)
        {
            string commtmplRes;
            string slaRes = "";
            string[] ret = { sla.Sanum, "" };

            cred.SelectCdMand(mandCd);

            if (createCommTemp)
                commtmplRes = CreateAndActivateCommTemplates(sla);
            else
                commtmplRes = "OK";

            if (!commtmplRes.Contains("ERROR"))
            {
                slaRes = CreateAndActivateSla(sla);
                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Creacion de SLAs", sla.Sanum + ": " + slaRes, root.Subject);
            }
            else
                ret[1] = "Error creación de SLAs en Control Desk: <b>" + sla.Sanum + "</b><br>Mensaje Communication template: <b>" + commtmplRes + "</b><br>Mensaje SLAs: <b>" + slaRes + "</b>";

            return ret;
        }
        private string CreateAndActivateCommTemplates(CdSlaData sla)
        {
            string activateMsg = "OK";

            console.WriteLine("Creando Communication Templates MX");
            sla = CreateStandardCommTemplates(sla);

            foreach (CdSlaEscalation escalation in sla.CdSlaEscalations)
                foreach (CdCommTemplate ct in escalation.Notifications)
                    if (ct.TemplateId.Contains("Error"))
                        return "ERROR: " + ct.TemplateId;

            List<string> templates = new List<string>();
            foreach (CdSlaEscalation escalation in sla.CdSlaEscalations)
                foreach (CdCommTemplate noti in escalation.Notifications)
                    templates.Add(noti.TemplateId);

            console.WriteLine("Creando Communication Templates  (" + string.Join(", ", templates) + ")");

            return ChangeCommunicationTemplatesStatus(templates, "ACTIVE");


        }
        private string ChangeCommunicationTemplatesStatus(List<string> templates, string status)
        {
            List<string> commTemplatesIds = new List<string>();

            Dictionary<string, string> commTemplatecommTemplatesIds = cdi.GetCommTemplatesStatus(templates);

            foreach (KeyValuePair<string, string> commTemplate in commTemplatecommTemplatesIds)
                commTemplatesIds.Add(commTemplate.Key);

            return cdi.ChangeCommunicationTemplatesStatus(commTemplatesIds, "ACTIVE");

        }
        private string CreateAndActivateSla(CdSlaData sla)
        {
            string sch = GetSlaSchedule(sla);

            console.WriteLine("Creando Sla MX");
            string ret = cdi.CreateSla(sla);

            console.WriteLine("Creando Sla Selenium");
            if (ret == "OK")
            {
                ret = cdSelenium.CreateSlaEscalation(sla, sch, root.UrlCd);
                if (ret.Contains("ERROR"))//reintentar 
                    ret = cdSelenium.CreateSlaEscalation(sla, sch, root.UrlCd);
            }
            else
                ret = "ERROR: " + ret;

            return ret;
        }
        private CdSlaData CreateStandardCommTemplates(CdSlaData sla)
        {
            string ticketType = "";
            if (sla.ObjectName == "INCIDENT")
                ticketType = "Incidente";
            else if (sla.ObjectName == "SR")
                ticketType = "Service Request";

            foreach (CdSlaEscalation escalation in sla.CdSlaEscalations)
            {
                string attrib = "", commitmentType = "";
                float commitmentTime = 0;
                if (escalation.TimeAttribute == "ADJUSTEDTARGETRESPONSETIME")
                {
                    attrib = "Respuesta";
                    commitmentType = "RESPONSE";
                }

                else if (escalation.TimeAttribute == "ADJUSTEDTARGETRESOLUTIONTIME")
                {
                    attrib = "Resolución";
                    commitmentType = "RESOLUTION";
                }

                //Buscar el value del attrib
                foreach (CdSlaCommitments commitment in sla.CdSlaCommitments)
                    if (commitment.Type == commitmentType)
                        commitmentTime = float.Parse(commitment.Value.ToString());

                //calcular porcentaje
                string percentage = GetEscPercentage(escalation.TimeInterval, commitmentTime);

                foreach (CdCommTemplate ct in escalation.Notifications)
                {

                    ct.Message = "&lt;div&gt;El ticket número :ticketid continua en estado :status. Por favor tomar las acciones correspondientes para evitar atrasos en la entrega del servicio." +
                           "&lt;/div&gt;&lt;div&gt;&lt;br /&gt;&lt;/div&gt;&lt;div&gt;&lt;br /&gt;&lt;/div&gt;&lt;div&gt;" +
                           "Contacto: :reportedbyname&lt;/div&gt;&lt;div&gt;" +
                           "Servicio: :COMMODITIES.description&lt;/div&gt;&lt;div&gt;" +
                           "Cliente: :PLUSPCUSTOMER.name&lt;/div&gt;&lt;div&gt;" +
                           "Descripción: :description&lt;/div&gt;&lt;div&gt;" +
                           "Hora del reporte: :reportdate&lt;/div&gt;&lt;div&gt;" +
                           "Clasificación: :CLASSSTRUCTURE.classificationdesc&lt;/div&gt;&lt;div&gt;" +
                           "Fecha esperada de finalización: :targetfinish&lt;/div&gt;&lt;!-- RICH TEXT --&gt;";
                    ct.Subject = "Notificación " + percentage + " de SLA " + attrib + " por vencer del " + ticketType + " :ticketid";


                    ct.TemplateId = cdi.CreateCommunicationTemplates(ct);
                }
            }
            return sla;
        }
        private string GetEscPercentage(string escTime, float commitmentTime)
        {
            float percentageF = 0;
            try
            {
                float timeInterval = float.Parse(escTime);
                if (timeInterval <= 0)
                    percentageF = (timeInterval + commitmentTime) / commitmentTime * 100;
                else
                    percentageF = timeInterval / commitmentTime * 100;
            }
            catch (Exception) { }
            string percentage = Math.Round(percentageF, 0) + "%";

            return percentage;
        }
        private string GetSlaSchedule(CdSlaData sla)
        {
            List<float> respoHours = new List<float>();
            List<float> resoHours = new List<float>();

            string sched;
            int schedule = 6;

            //tomar las listas iguales
            foreach (var slaEscalation in sla.CdSlaEscalations)
            {
                if (slaEscalation.TimeAttribute == "ADJUSTEDTARGETRESPONSETIME")
                {
                    if (float.TryParse(slaEscalation.TimeInterval.ToString(), out float respoF))
                        respoHours.Add(respoF);
                }
                else if (slaEscalation.TimeAttribute == "ADJUSTEDTARGETRESOLUTIONTIME")
                {
                    if (float.TryParse(slaEscalation.TimeInterval.ToString(), out float resoF))
                        resoHours.Add(resoF);
                }
            }
            resoHours.Sort();
            respoHours.Sort();

            try
            {
                int scheduleS = 0;
                try { scheduleS = (int)Math.Abs(Math.Round((resoHours[0] - resoHours[1]) * 6, 0)); } catch (Exception) { }
                int scheduleR = 0;
                try { scheduleR = (int)Math.Abs(Math.Round((respoHours[0] - respoHours[1]) * 6, 0)); } catch (Exception) { }

                if (scheduleR != 0 && scheduleS != 0)
                    schedule = Math.Min(scheduleR, scheduleS);
                else if (scheduleR != 0)
                    schedule = scheduleR;
                else
                    schedule = scheduleS;

            }
            catch (Exception) { }

            sched = schedule + "m,*,0,*,*,*,*,*,*,*";

            return sched;
        }

    }
}
