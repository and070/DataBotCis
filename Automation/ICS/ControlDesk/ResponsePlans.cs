using DataBotV5.Logical.Projects.ControlDesk;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using System.Data;
using System.Linq;
using System;
using System.IO;

namespace DataBotV5.Automation.ICS.ControlDesk
{
    internal class ResponsePlans
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

        string respFinal = "";
        const string rpMonMandCd = "DEV";

        public void Main()
        {
            cred.SelectCdMand("DEV");


            List<string> cisList = new List<string>();
            cisList.Add("CRI400A");
            cisList.Add("CRI400B_PRIMA");

            CdCollectionData collection = new CdCollectionData();
            collection.CollectionNum = "EVCMI001";
            collection.Cis = cisList;

            string XXX = cdi.AddCollectionCis(collection);

            

            //y crear la clase de CIS

 
        }

        /// <summary> 
        /// Verifica la plantilla del excel y clasifica el RP en ICS, MON o SLA
        /// </summary>
        /// <param name="excelDt"></param>
        /// <param name="filePath"></param>
        private void ProcessResponsePlans(DataSet excelFile, string filePath)
        {
            ResponseRP responseRP;

            string excelVersion = GetExcelVersion(filePath);

            if (excelVersion == "ICS")
            {
                DataTable excelDt = excelFile.Tables[0];
                responseRP = ProcessResponsePlanICS(excelDt);
                SendMails(responseRP, filePath);
            }
            else if (excelVersion == "MON")
            {
                DataTable excelDt = excelFile.Tables["Response Plan"];
                responseRP = ProcessResponsePlanMon(excelDt);
                SendMails(responseRP, filePath);
            }
            else
            {
                mail.SendHTMLMail("Por favor utilizar la plantilla oficial de creación de Response Plan", new string[] { root.BDUserCreatedBy }, root.Subject);
            }

            root.requestDetails = respFinal;
        }
        /// <summary>
        /// Procesar RP de monitoreo
        /// </summary>
        /// <param name="excelDt"></param>
        /// <returns></returns>
        private ResponseRP ProcessResponsePlanMon(DataTable excelDt)
        {
            ResponseRP ret = new ResponseRP();
            bool valExcel = false;

            if (excelDt == null)
                valExcel = false;

            try
            {
                if (excelDt.Columns[12].ColumnName == "x")
                    valExcel = true;
            }
            catch (Exception) { }


            if (valExcel)
            {
                string actionCol = "Acción";

                cred.SelectCdMand(rpMonMandCd);
                ResponsePlanProcessResult rpResAll = new ResponsePlanProcessResult();
                rpResAll.ResponseDt = new DataTable();

                rpResAll.ResponseDt.Columns.Add("Id del Response Plan");
                rpResAll.ResponseDt.Columns.Add("Clasificación");
                rpResAll.ResponseDt.Columns.Add("Acción");
                rpResAll.ResponseDt.Columns.Add("Respuesta");

                rpResAll.SendIcs = false;
                rpResAll.SendUser = false;

                foreach (DataRow responsePlansRow in excelDt.Rows)
                {
                    string rpAction = responsePlansRow[actionCol].ToString().Trim();

                    if (rpAction == "Crear")
                    {
                        ResponsePlanProcessResult rpResRow = UploadResponsePlansMon(responsePlansRow);
                        rpResAll.ResponseDt.Merge(rpResRow.ResponseDt);
                        rpResAll.SendIcs = rpResAll.SendIcs || rpResRow.SendIcs;
                        rpResAll.SendUser = rpResAll.SendUser || rpResRow.SendUser;
                    }
                    else if (rpAction.Contains("Actualizar"))
                    {
                        //inactivar y después activar
                        ResponsePlanProcessResult rpResRow = UploadResponsePlansMon(responsePlansRow, true);
                        rpResAll.ResponseDt.Merge(rpResRow.ResponseDt);
                        rpResAll.SendIcs = rpResAll.SendIcs || rpResRow.SendIcs;
                        rpResAll.SendUser = rpResAll.SendUser || rpResRow.SendUser;

                    }
                    else if (rpAction.Contains("Inactivar") || rpAction.Contains("Activar"))
                    {
                        ResponsePlanProcessResult rpResRow = ChangeResponsePlanStatus(responsePlansRow);
                        rpResAll.ResponseDt.Merge(rpResRow.ResponseDt);
                        rpResAll.SendIcs = rpResAll.SendIcs || rpResRow.SendIcs;
                        rpResAll.SendUser = rpResAll.SendUser || rpResRow.SendUser;
                    }
                }
                ret.ResponseDt = rpResAll.ResponseDt;
                ret.SendICS = rpResAll.SendIcs;
                ret.SendUser = rpResAll.SendUser;
            }
            else
            {
                DataTable retDt = new DataTable();
                retDt.Columns.Add("Respuesta");
                DataRow retDtRow = retDt.NewRow();
                retDtRow["Respuesta"] = "Por favor utilizar la plantilla oficial de creación de Response Plan";
                retDt.Rows.Add(retDtRow);

                ret.ResponseDt = retDt;
                ret.SendUser = true;

            }
            root.requestDetails = respFinal;
            return ret;
        }
        private ResponsePlanProcessResult ChangeResponsePlanStatus(DataRow responsePlansRow)
        {
            ResponsePlanProcessResult ret = new ResponsePlanProcessResult();

            DataTable responseDt = new DataTable();
            responseDt.Columns.Add("Id del Response Plan");
            responseDt.Columns.Add("Clasificación");
            responseDt.Columns.Add("Acción");
            responseDt.Columns.Add("Respuesta");

            string rpAction = responsePlansRow["Acción"].ToString().Trim();
            string rpId = responsePlansRow["ID Response Plan"].ToString().Trim();

            try
            {
                string statusRes = "";
                if (rpAction == "Activar")
                    statusRes = cdi.ChangeResponsePlanStatus(rpId, "ACTIVE");
                else if (rpAction == "Inactivar")
                    statusRes = cdi.ChangeResponsePlanStatus(rpId, "INACTIVE");

                DataRow responseRow = responseDt.NewRow();
                responseRow["Id del Response Plan"] = rpId;
                responseRow["Acción"] = rpAction;
                responseRow["Respuesta"] = statusRes;
                responseDt.Rows.Add(responseRow);

                ret.ResponseDt = responseDt;

                if (statusRes != "OK")
                {
                    ret.SendUser = true;
                    ret.SendIcs = true;
                }
                else
                {
                    ret.SendUser = false;
                    ret.SendIcs = false;
                }
            }
            catch (Exception ex)
            {
                DataRow errorRow = responseDt.NewRow();
                errorRow["Respuesta"] = ex.Message;
                responseDt.Rows.Add(errorRow);
                ret.ResponseDt = responseDt;
                ret.SendUser = false;
                ret.SendIcs = true;
            }

            return ret;
        }
        private ResponsePlanProcessResult UploadResponsePlansMon(DataRow responsePlansRow, bool isChange = false)
        {
            ResponsePlanProcessResult ret = new ResponsePlanProcessResult();
            string deleteCommodityRes = "OK";
            string deleteCiRes = "OK";

            DataTable responseDt = new DataTable();
            responseDt.Columns.Add("Id del Response Plan");
            responseDt.Columns.Add("Clasificación");
            responseDt.Columns.Add("Acción");
            responseDt.Columns.Add("Respuesta");

            DataRow responseRow = responseDt.NewRow();
            DataTable internalClassificationId = new DataTable();

            string rpAction = responsePlansRow["Acción"].ToString().Trim();
            string rpId = responsePlansRow["ID Response Plan"].ToString().Trim();
            string rpClassification = responsePlansRow["Class Structure ID"].ToString().Trim();
            string rpCountry = responsePlansRow["País"].ToString().Trim();
            string rpPersonGroup = responsePlansRow["Grupo resolutor"].ToString().Trim();
            string rpCustomer = responsePlansRow["Cliente"].ToString().Trim();
            string rpSchedule = responsePlansRow["Horario"].ToString().Trim();
            string rpNoc = responsePlansRow["NOC 4.0"].ToString().Trim();
            string rpAuto = responsePlansRow["Auto Asignación"].ToString().Trim();
            string rpCommodity = responsePlansRow["Servicio"].ToString().Trim();
            string ranking = "1";
            string shift = "CA8X5";
            string actionGroup = "";

            try
            {
                CdResponsePlanData rpData = cdi.GetResponsePlanData(rpId);

                if (isChange)
                {
                    // Inactivar RP para poder modificarlo
                    cdi.ChangeResponsePlanStatus(rpId, "INACTIVE");

                    foreach (CdServicesData service in rpData.Services)
                    {
                        deleteCommodityRes = cdi.DeleteRpCommodity(rpId, service.PluspApplServId);

                        if (deleteCommodityRes != "OK")
                        {
                            responseRow["Id del Response Plan"] = deleteCommodityRes;
                            ret.SendIcs = true;
                            ret.SendUser = false;
                        }
                    }
                }

                if (deleteCommodityRes == "OK")
                {
                    if (rpSchedule == "8x7")
                    {
                        shift = "LV8X7";
                        if (rpCountry == "PA")
                            shift = "LVPA8X7";
                        if (rpCountry == "DO")
                            shift = "LVDO8X7";
                        if (rpCountry == "CO")
                            shift = "LVCO8X7";
                    }
                    else if (rpSchedule == "7x10")
                    {
                        shift = "LV7-10";
                        if (rpCountry == "PA")
                            shift = "LVPA7-10";
                        if (rpCountry == "DO")
                            shift = "LVDO7-10";
                        if (rpCountry == "CO")
                            shift = "LVCO7-10";
                    }
                    else if (rpSchedule == "7x7")
                    {
                        shift = "LV7X7";
                        if (rpCountry == "PA")
                            shift = "LVPA7X7";
                        if (rpCountry == "DO")
                            shift = "LVDO7X7";
                        if (rpCountry == "CO")
                            shift = "LVCO7X7";
                    }
                    else
                    {
                        if (rpCountry == "PA")
                            shift = "PA8X5";
                        if (rpCountry == "DO")
                            shift = "DO8X5";
                        if (rpCountry == "CO")
                            shift = "CO8X5";
                    }

                    string calendar = "GBM" + rpCountry;
                    if (rpCountry != "GT")
                        calendar += "20";
                    string calendarOrgId = "GBM";

                    List<string> rpRowCisList = new List<string>();

                    for (int i = 11; i < responsePlansRow.ItemArray.Length; i++) //11 es la columna "Configuration Item (CI)"
                        if (responsePlansRow[i].ToString() != "")
                            rpRowCisList.Add(responsePlansRow[i].ToString());

                    rpRowCisList = rpRowCisList.Distinct().ToList();

                    if (isChange)
                    {
                        //Procesar los CIS
                        List<string> currentRpCisList = new List<string>();

                        foreach (string configurationItem in rpData.ConfigurationItems)
                            currentRpCisList.Add(configurationItem);

                        if (rpAction.Contains("Agregar CIs"))
                        {
                            rpRowCisList.AddRange(currentRpCisList);//Suma Cis del excel a los actuales de CD
                            rpRowCisList = rpRowCisList.Distinct().ToList();
                        }
                        else if (rpAction.Contains("Eliminar Cis"))
                        {
                            foreach (string rpRowCi in rpRowCisList)
                            {
                                currentRpCisList.Remove(rpRowCi);
                                deleteCiRes = cdi.DeleteRpCis(rpId, rpRowCi);

                                if (deleteCiRes != "OK")
                                    break;
                            }
                            rpRowCisList = currentRpCisList.Distinct().ToList();
                        }

                    }

                    #region Validaciones
                    if (deleteCiRes == "OK")
                    {
                        #region Validar rpClassification
                        try
                        {
                            internalClassificationId = cdi.GetInternalClassificationId(rpClassification);
                            responseRow["Clasificación"] = isChange ? "" : internalClassificationId.Rows[0]["DESCRIPTION"].ToString();
                        }
                        catch (Exception ex)
                        {
                            responseRow["Respuesta"] = ex.Message;
                            ret.SendIcs = true;
                        }
                        #endregion


                        if (internalClassificationId.Rows.Count > 0)
                        {
                            if (rpCustomer.Length < 10)
                                rpCustomer = rpCustomer.PadLeft(10, '0');

                            string rpApp = internalClassificationId.Rows[0]["APPLICATION"].ToString();
                            string sanum = rpId;

                            if (!isChange)
                                sanum = CreateSanum(rpPersonGroup, rpCountry, rpApp);


                            string classificationId = "";
                            foreach (DataRow classificationIds in internalClassificationId.Rows)
                            {
                                classificationId = string.Concat(classificationIds["CLASSIFICATIONID"].ToString().Take(2));
                                if (classificationId == "30" || classificationId == "20")
                                    break;
                            }

                            string customerName = cdi.GetCustomerName(rpCustomer);

                            if (customerName != "NE")
                            {
                                string description = "Response plan " + string.Concat(rpApp.Take(2)) + " Eventos " + rpPersonGroup + " - " + customerName;

                                if (rpNoc == "Sí")
                                    description += " - NOC 4.0";

                                if (rpAuto == "Sí")
                                    description += " - Auto";

                                if (rpSchedule == "No Hábil")
                                {
                                    ranking = "3";
                                    shift = "";
                                    calendar = "";
                                    calendarOrgId = "";
                                    description += " " + rpSchedule;
                                }


                                if (rpCommodity != "")//Validar ActionGroup
                                    actionGroup = rpCommodity + string.Concat(rpApp.Take(2));


                                if (rpAuto != "Sí" || rpSchedule != "No Hábil" || classificationId != "20")
                                {
                                    if (classificationId != "21" && classificationId != "31" && classificationId != "")
                                    {
                                        if (cdi.CheckPersonGroupExistence(rpPersonGroup) == "OK")
                                        {
                                            rpAuto = rpAuto == "Sí" ? "1" : "0";

                                            CdServicesData commodity = new CdServicesData { Commodity = rpCommodity };

                                            CdResponsePlanData rp = new CdResponsePlanData
                                            {
                                                Sanum = sanum,
                                                Description = description,
                                                GbmAutoAssignment = rpAuto,
                                                Ranking = ranking,
                                                ObjectName = rpApp,
                                                Services = new CdServicesData[] { commodity },
                                                CalendarOrgId = calendarOrgId,
                                                Calendar = calendar,
                                                Shift = shift,
                                                AssignOwnerGroup = rpPersonGroup,
                                                Status = "ACTIVE",
                                                ClassStructureId = internalClassificationId.Rows[0]["CLASSSTRUCTUREID"].ToString(),
                                                CustomerId = rpCustomer,
                                                ConfigurationItems = rpRowCisList,
                                                Condition = $":externalsystem = 'BOTOMNIBUSCONTROLDESK' AND :pluspcustomer = '{rpCustomer}'",
                                                Action = actionGroup
                                            };

                                            rp = cdi.CreateOrChangeResponsePlans(rp);

                                            if (rp.ResponseMessage == "OK")
                                            {
                                                string logMsg = isChange ? "Modificación de Response Plan" : "Creación de Response Plan";

                                                responseRow["Id del Response Plan"] = sanum;
                                                responseRow["Acción"] = rpAction;
                                                responseRow["Respuesta"] = "OK";
                                                log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, logMsg, rp.Sanum, root.Subject);
                                                respFinal = respFinal + "\\n" + logMsg + ": " + rp.Sanum + " " + root.Subject;

                                            }
                                            else
                                            {

                                                responseRow["Id del Response Plan"] = isChange ? sanum : "";
                                                responseRow["Acción"] = rpAction;
                                                responseRow["Respuesta"] = rp.ResponseMessage;
                                                ret.SendIcs = true;
                                                ret.SendUser = true;
                                            }
                                        }
                                        else
                                        {
                                            responseRow["Id del Response Plan"] = isChange ? sanum : "";
                                            responseRow["Acción"] = rpAction;
                                            responseRow["Respuesta"] = "El Person group: " + rpPersonGroup + " no existe.";
                                            ret.SendUser = true;
                                        }
                                    }
                                    else
                                    {
                                        responseRow["Id del Response Plan"] = isChange ? sanum : "";
                                        responseRow["Acción"] = rpAction;
                                        responseRow["Respuesta"] = "El Classification seleccionado no está habilitado o no existe.";
                                        ret.SendIcs = true;
                                        ret.SendUser = true;
                                    }
                                }
                                else
                                {
                                    responseRow["Id del Response Plan"] = isChange ? sanum : "";
                                    responseRow["Acción"] = rpAction;
                                    responseRow["Respuesta"] = "La opción de auto asignación no es válida para horario no hábil";
                                    ret.SendIcs = true;
                                    ret.SendUser = true;
                                }
                            }
                            else
                            {
                                responseRow["Id del Response Plan"] = isChange ? sanum : "";
                                responseRow["Acción"] = rpAction;
                                responseRow["Respuesta"] = "El cliente indicado no existe";
                                ret.SendUser = true;
                            }
                        }
                        else
                        {
                            responseRow["Acción"] = rpAction;
                            responseRow["Respuesta"] = "La Clasificación no existe";
                            ret.SendUser = true;
                        }
                    }
                    else
                    {
                        responseRow["Id del Response Plan"] = isChange ? rpId : "";
                        responseRow["Acción"] = rpAction;
                        responseRow["Respuesta"] = "Error al eliminar CI: " + deleteCiRes;
                        ret.SendUser = true;
                        ret.SendIcs = true;
                    }
                    #endregion
                }

                responseDt.Rows.Add(responseRow);
            }
            catch (Exception ex)
            {
                DataRow errorRow = responseDt.NewRow();
                errorRow["Respuesta"] = ex.Message;
                responseDt.Rows.Add(errorRow);
                ret.SendUser = false;
                ret.SendIcs = true;
            }

            ret.ResponseDt = responseDt;
            return ret;
        }
        /// <summary>
        /// Procesar RP custom
        /// </summary>
        /// <param name="excelDt"></param>
        /// <returns></returns>
        public ResponseRP ProcessResponsePlanICS(DataTable excelDt)
        {
            ResponseRP ret = new ResponseRP();

            DataTable response = new DataTable();
            response.Columns.Add("Clasificación");
            response.Columns.Add("Id del Response Plan");

            bool sendIcs = false, sendUser = false;

            List<string> rpCommodityColumns = new List<string>();

            foreach (DataColumn column in excelDt.Columns)
                if (column.ColumnName.Contains("Servicio"))
                    rpCommodityColumns.Add(column.ColumnName);

            foreach (DataRow responsePlansRow in excelDt.Rows)
            {
                DataRow responseRow = response.NewRow();
                DataTable internalClassificationId = new DataTable();

                string rpAction = responsePlansRow["Acción"].ToString().Trim();

                if (rpAction != "") // Valida y si viene en blanco acción sale de foreach
                {
                    List<string> rpCommodities = new List<string>();
                    List<string> rpCisList = new List<string>();

                    string mandCd = responsePlansRow["Mandante"].ToString().Trim();
                    string rpId = responsePlansRow["ID Response Plan"].ToString().Trim();
                    string rpClassification = responsePlansRow["Class Structure ID"].ToString().Trim();
                    string rpClassificationDesc = responsePlansRow["Clasificación"].ToString().Trim();
                    string rpPersonGroup = responsePlansRow["Grupo resolutor"].ToString().Trim();
                    string rpCustomer = responsePlansRow["Cliente"].ToString().Trim();
                    string rpCondition = responsePlansRow["Condition"].ToString().Trim();
                    string rpRanking = responsePlansRow["Ranking"].ToString().Trim();
                    string rpDescription = responsePlansRow["Descripción"].ToString().Trim();
                    string sanum = responsePlansRow["ID Response Plan"].ToString().Trim();
                    #region rpCommodities
                    foreach (string rpCommodityColumn in rpCommodityColumns)
                        rpCommodities.Add(responsePlansRow[rpCommodityColumn].ToString());
                    rpCommodities = rpCommodities.Distinct().ToList();
                    #endregion
                    string rpSchedule = responsePlansRow["Horario"].ToString().Trim();
                    string rpNoc = responsePlansRow["NOC 4.0"].ToString().Trim();
                    string rpAuto = responsePlansRow["Auto Asignación"].ToString().Trim();
                    string rpApp = responsePlansRow["Aplicación"].ToString().Trim();
                    string rpCountry = responsePlansRow["País"].ToString().Trim();
                    #region rpCIs
                    for (int i = excelDt.Columns["Configuration Item (CI)"].Ordinal; i < excelDt.Columns.Count; i++)
                        if (responsePlansRow[i].ToString() != "")
                            rpCisList.Add(responsePlansRow[i].ToString());

                    rpCisList = rpCisList.Distinct().ToList();
                    #endregion

                    string shift = "CA8X5";
                    string calendarOrgId = "GBM";

                    //Calendarios

                    if (rpSchedule == "8x7")
                    {
                        shift = "LV8X7";
                        if (rpCountry == "PA")
                            shift = "LVPA8X7";
                        if (rpCountry == "DO")
                            shift = "LVDO8X7";
                        if (rpCountry == "CO")
                            shift = "LVCO8X7";
                    }
                    else if (rpSchedule == "7x10")
                    {
                        shift = "LV7-10";
                        if (rpCountry == "PA")
                            shift = "LVPA7-10";
                        if (rpCountry == "DO")
                            shift = "LVDO7-10";
                        if (rpCountry == "CO")
                            shift = "LVCO7-10";
                    }
                    else
                    {
                        if (rpCountry == "PA")
                            shift = "PA8X5";
                        if (rpCountry == "DO")
                            shift = "DO8X5";
                        if (rpCountry == "CO")
                            shift = "CO8X5";
                    }

                    string calendar = "GBM" + rpCountry;
                    if (rpCountry != "GT")
                        calendar += "20";


                    #region Validaciones

                    cred.SelectCdMand(mandCd);

                    #region Validar rpClassification
                    if (rpClassification != "")
                    {
                        try
                        {
                            internalClassificationId = cdi.GetInternalClassificationId(rpClassification);
                            responseRow["Clasificación"] = internalClassificationId.Rows[0]["DESCRIPTION"].ToString();
                        }
                        catch (Exception ex)
                        {
                            responseRow["Id del Response Plan"] = ex.Message;
                            sendIcs = true;
                        }
                    }
                    #endregion

                    if (internalClassificationId.Rows.Count > 0 || rpClassification == "")
                    {
                        if (rpCustomer.Length < 10)
                            rpCustomer = rpCustomer.PadLeft(10, '0');

                        string classStructureId = "";
                        if (internalClassificationId.Rows.Count > 0)
                        {
                            rpApp = internalClassificationId.Rows[0]["APPLICATION"].ToString();
                            classStructureId = internalClassificationId.Rows[0]["CLASSSTRUCTUREID"].ToString();
                        }

                        string classificationId = "";
                        foreach (DataRow classificationIds in internalClassificationId.Rows)
                        {
                            classificationId = string.Concat(classificationIds["CLASSIFICATIONID"].ToString().Take(2));
                            if (classificationId == "30" || classificationId == "20")
                                break;
                        }

                        string customerName = cdi.GetCustomerName(rpCustomer);

                        if (customerName != "NE")
                        {
                            if (rpSchedule == "24X7" || rpSchedule == "No Hábil")
                            {
                                shift = "";
                                calendar = "";
                                calendarOrgId = "";
                            }

                            if (rpAuto != "Sí" || rpSchedule != "No Hábil" || classificationId != "20")
                            {
                                if (classificationId != "21" && classificationId != "31")
                                {
                                    if (cdi.CheckPersonGroupExistence(rpPersonGroup) == "OK")
                                    {
                                        List<CdServicesData> commoditiesList = new List<CdServicesData>();

                                        rpAuto = rpAuto == "Sí" ? "1" : "0";

                                        foreach (string rpCommodity in rpCommodities)
                                        {
                                            CdServicesData commodity = new CdServicesData { Commodity = rpCommodity };
                                            commoditiesList.Add(commodity);
                                        }

                                        CdResponsePlanData rp = new CdResponsePlanData
                                        {
                                            Sanum = sanum,
                                            Description = rpDescription,
                                            GbmAutoAssignment = rpAuto,
                                            Ranking = rpRanking,
                                            ObjectName = rpApp,
                                            CalendarOrgId = calendarOrgId,
                                            Calendar = calendar,
                                            Shift = shift,
                                            AssignOwnerGroup = rpPersonGroup,
                                            Status = "ACTIVE",
                                            ClassStructureId = classStructureId,
                                            CustomerId = rpCustomer,
                                            Condition = rpCondition,
                                            ConfigurationItems = rpCisList,
                                            Services = commoditiesList.ToArray()
                                        };

                                        rp = cdi.CreateOrChangeResponsePlans(rp);

                                        if (!rp.ResponseMessage.ToUpper().Contains("ERROR"))
                                        {
                                            responseRow["Id del Response Plan"] = rp.Sanum;
                                            log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Creación de Response Plan", rp.Sanum, root.Subject);
                                            respFinal = respFinal + "\\n" + "Creación de Response Plan: " + rp.Sanum + " " + root.Subject;
                                        }
                                        else
                                        {
                                            responseRow["Id del Response Plan"] = rp.ResponseMessage;
                                            sendIcs = true;
                                            sendUser = true;
                                        }
                                    }
                                    else
                                    {
                                        responseRow["Id del Response Plan"] = "El Person group: " + rpPersonGroup + " no existe.";
                                        sendUser = true;
                                    }
                                }
                                else
                                {
                                    responseRow["Id del Response Plan"] = "El Classification seleccionado no está habilitado o no existe.";
                                    sendIcs = true;
                                    sendUser = true;
                                }
                            }
                            else
                            {
                                responseRow["Id del Response Plan"] = "La opción de auto asignación no es válida para horario no hábil";
                                sendIcs = true;
                                sendUser = true;
                            }
                        }
                        else
                        {
                            responseRow["Id del Response Plan"] = "El cliente indicado no existe";
                            sendUser = true;
                        }
                    }
                    else
                    {
                        responseRow["Id del Response Plan"] = "La Clasificación no existe";
                        sendUser = true;
                    }

                    #endregion

                    response.Rows.Add(responseRow);
                }
            }

            ret.SendICS = sendIcs;
            ret.SendUser = sendUser;
            ret.ResponseDt = response;

            return ret;
        }
        /// <summary>
        /// Genera un id para un RP
        /// </summary>
        /// <param name="rpPersonGroup"></param>
        /// <param name="rpCountry"></param>
        /// <param name="rpApp"></param>
        /// <returns></returns>
        private string CreateSanum(string rpPersonGroup, string rpCountry, string rpApp)
        {
            string sanum = "";

            try
            {
                string sanumpart = "";
                sanum = "GBM" + string.Concat(rpApp.Take(2)) + "EV";
                #region PersonGroup part

                if (rpPersonGroup.Length > 3)
                {
                    string[] split = rpPersonGroup.Split('-', ' ', '_');
                    if (split.Length == 1)
                        sanumpart = string.Concat(split[0].Take(3));
                    else
                        sanumpart = string.Concat(split[0].Take(2)) + string.Concat(split[1].Take(1));
                }
                else
                    sanumpart = rpPersonGroup;

                #endregion
                sanum += sanumpart + rpCountry;
                sanumpart = "001";

                int i = 1;
                bool rpExists = false;
                do
                {
                    sanumpart = i.ToString().PadLeft(3, '0');
                    rpExists = cdi.CheckResponsePlansExistence(sanum + sanumpart);
                    i++;
                } while (i > 999 || rpExists);

                sanum += sanumpart;
            }
            catch (Exception) { }

            return sanum;
        }
        /// <summary>
        /// Verifica si la plantilla es correcta
        /// </summary>
        /// <param name="attachPath"></param>
        /// <returns></returns>
        public string GetExcelVersion(string attachPath)
        {
            string excelTag = "";
            DataSet excelFile = new DataSet();
            try
            {
                excelFile = excel.GetExcelBook(attachPath);
            }
            catch (Exception)
            {
                return "ERROR";
            }

            if (excelFile.Tables.Count > 1)
            {
                foreach (DataTable sheet in excelFile.Tables)
                {
                    if (sheet.TableName == "Response Plan")
                    {
                        return "MON";
                    }
                }
            }

            //leer la version del file
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder objFolder = shell.NameSpace(Path.GetDirectoryName(attachPath));

            foreach (Shell32.FolderItem2 item in objFolder.Items())
            {
                if (item.Name == Path.GetFileName(attachPath))
                {
                    excelTag = objFolder.GetDetailsOf(item, 18); //18 es etiqueta
                }
            }

            if (excelTag == "ICS")
                return excelTag;
            else if (excelTag == "MON")
                return excelTag;
            else if (excelTag == "SLA")
                return excelTag;
            else
                return "ERROR";
        }
        /// <summary>
        /// Envia las notificaciones
        /// </summary>
        /// <param name="responseRP"></param>
        /// <param name="filePath"></param>
        private void SendMails(ResponseRP responseRP, string filePath)
        {

            if (responseRP.SendICS && responseRP.SendUser)
                mail.SendHTMLMail("Falló Creación de Response Plan en Control Desk<br><br>" + val.ConvertDataTableToHTML(responseRP.ResponseDt), new string[] { root.BDUserCreatedBy }, root.BDProcess);
            else if (responseRP.SendUser)
                mail.SendHTMLMail("Falló Creación de Response Plan en Control Desk<br><br>" + val.ConvertDataTableToHTML(responseRP.ResponseDt), new string[] { root.BDUserCreatedBy }, root.BDProcess);
            else if (responseRP.SendICS)
                mail.SendHTMLMail("Falló Creación de Response Plan en Control Desk<br><br>" + val.ConvertDataTableToHTML(responseRP.ResponseDt), new string[] { "internalcustomersrvs@gbm.net" }, root.BDProcess);
            else
                mail.SendHTMLMail("Creación Satisfactoria de Response Plan en Control Desk<br><br>" + val.ConvertDataTableToHTML(responseRP.ResponseDt), new string[] { root.BDUserCreatedBy }, root.Subject);
        }
        public class ResponseRP
        {
            public DataTable ResponseDt { get; set; }
            public bool SendICS { get; set; }
            public bool SendUser { get; set; }
        }

        private class ResponsePlanProcessResult
        {
            public DataTable ResponseDt { get; set; }
            public bool SendIcs { get; set; }
            public bool SendUser { get; set; }
        }
    }
}
