using DataBotV5.Logical.Projects.ControlDesk;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Data.Credentials;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System.Linq;
using System.Xml;
using System;

namespace DataBotV5.Automation.ICS.ControlDesk
{
    /// <summary>
    /// Clase ICS Automation encargada de gestionar los contratos del control desk.
    /// </summary>
    class ContractsControlDesk
    {
        readonly ControlDeskAprovalContract cdSelenium = new ControlDeskAprovalContract();
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly SapVariants sap = new SapVariants();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();
        string respFinal = "";

        readonly List<string> contractsSelenium = new List<string>();

        public void Main()
        {
            string date = ManualOn();
            int mandCrm = Int32.Parse(GetDefaultMand("CRM"));
            int mandErp = Int32.Parse(GetDefaultMand("ERP"));
            string mandCd = GetDefaultMand("CD");

            if (date == "")
                ProcessContracts(mandCd, mandCrm, mandErp, DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd"));
            else
                ProcessContracts(mandCd, mandCrm, mandErp, date, date);

            console.WriteLine("Creando Estadísticas");
            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "internalcustomersrvs";

            using (Stats stats = new Stats()) { stats.CreateStat(); }
        }
        private void ProcessContracts(string mandCd, int mandCrm, int mandErp, string sDate, string eDate)
        {
            cred.SelectCdMand(mandCd);

            console.WriteLine("Tomar información de SAP: " + mandCrm);

            #region Tomar información de SAP

            DataTable newOnHoldCon = GetNewOnHoldContracts(mandCrm, mandErp, sDate, eDate); //tabla del los nuevos contratos OnHold
            DataTable newContracts = GetNewContracts(mandCrm, sDate, eDate);                //tabla de los nuevos contratos InProcess
            DataSet dsRenewals = GetRenewalContracts(mandCrm, sDate, eDate);                //tablas de los contratos renovados
            List<string> closedContracts = GetClosedContracts(mandCrm, sDate, eDate);       //lista con los contratos que se cerraron en el día
            List<CdContractData> closedConsWithCis = GetContractsWithCis(closedContracts);  //lista de los datos de los cons a cerrar que tienen CIS

            #endregion

            console.WriteLine("Creando nuevos contratos para CD: " + mandCd);

            #region Crear contratos nuevos

            DataTable resConCreation = UploadNewContracts(newContracts, mandErp, mandCrm); //Crear en CD los contratos nuevos

            foreach (DataRow item in resConCreation.Rows)
            {
                if (!(item[0].ToString().Contains("No se creo el Contrato") || item[0].ToString().Contains("Ya existe en Control Desk")))
                {
                    log.LogDeCambios("", root.BDProcess, "Datos Maestros", "Crear Contrato", item[0].ToString() + ": " + item[1].ToString(), "**" + mandCd + "** " + root.Subject);
                    respFinal = respFinal + "\\n" + "Crear Contrato: " + item[0].ToString() + ": " + item[1].ToString();
                    contractsSelenium.Add(item[0].ToString());
                }
            }

            #endregion

            console.WriteLine("Creando nuevos contratos ONHOLD para CD: " + mandCd);

            #region Crear contratos ONHOLD nuevos
            DataTable resOnHoldCreation = UploadNewContracts(newOnHoldCon, mandErp, mandCrm); //Crear en CD ciertos contratos ONHOLD

            foreach (DataRow item in resOnHoldCreation.Rows)
            {
                if (!(item[0].ToString().Contains("No se creo el Contrato") || item[0].ToString().Contains("Ya existe en Control Desk")))
                {
                    log.LogDeCambios("", root.BDProcess, "Datos Maestros", "Crear Contrato ONHOLD", item[0].ToString() + ": " + item[1].ToString(), "**" + mandCd + "** " + root.Subject);
                    respFinal = respFinal + "\\n" + "Crear Contrato ONHOLD " + item[0].ToString() + ": " + item[1].ToString();
                    contractsSelenium.Add(item[0].ToString());
                }
            }

            #endregion

            console.WriteLine("Renovando contratos en CD: " + mandCd);

            #region Actualizar contratos renovados

            DataTable updatedConDates = RenewContract(dsRenewals, mandErp); //renueva los contratos en CD, devuelve los contratos para activar

            foreach (DataRow item in updatedConDates.Rows)
            {
                if (item[1].ToString() == "OK")
                {
                    contractsSelenium.Add(item[0].ToString());
                    log.LogDeCambios("", root.BDProcess, "Datos Maestros", "Renovar Contrato", item[0].ToString() + ": " + item[1].ToString(), "**" + mandCd + "** " + root.Subject);
                    respFinal = respFinal + "\\n" + "Renovar Contrato " + item[0].ToString() + ": " + item[1].ToString();
                }
            }

            #endregion

            console.WriteLine("Actualizar ITEMS contratos en CD: " + mandCd);

            #region Actualizar items en contratos

            DataTable updateItems = RenewItems(mandErp, dsRenewals, contractsSelenium);//cambios recomendados(buscar los status masivos)

            foreach (DataRow item in updateItems.Rows)
            {
                if (item[1].ToString() == "OK")
                {
                    contractsSelenium.Add(item[0].ToString());
                    log.LogDeCambios("", root.BDProcess, "Datos Maestros", "Actualizar Items Contrato", item[0].ToString() + ": " + item[1].ToString(), "**" + mandCd + "** " + root.Subject);
                    respFinal = respFinal + "\\n" + "Actualizar Items Contrato " + item[0].ToString() + ": " + item[1].ToString();
                }
            }

            #endregion

            console.WriteLine("Aprobando contratos creados en CD: " + mandCd);

            #region Aprobar contratos creados

            string apprCon = ApprContracts(contractsSelenium);
            if (!apprCon.ToLower().Contains("error"))
            {
                log.LogDeCambios("", root.BDProcess, "Datos Maestros", "Aprobar contratos", apprCon, "**" + mandCd + "** " + root.Subject);
                respFinal = respFinal + "\\n" + "Aprobar contratos: " + apprCon;
            }
            #endregion

            console.WriteLine("Finalizando contratos en CD: " + mandCd);

            #region Cerrar contratos finalizados

            string closeCon = CloseContracts(closedContracts);
            if (!closeCon.ToLower().Contains("error"))
            {
                log.LogDeCambios("", root.BDProcess, "Datos Maestros", "Cerrar Contrato", closeCon, "**" + mandCd + "** " + root.Subject);
                respFinal = respFinal + "\\n" + "Cerrar Contrato: " + closeCon;
            }

            #endregion

            console.WriteLine("Creando Releases: " + mandCd);

            #region Crear Releases de los contratos ONHOLD

            DataTable relCreation = new DataTable();
            if (newOnHoldCon.Rows.Count > 0) // hay contratos validos
            {
                relCreation = CreateReleases(newOnHoldCon);//release, status, CON, DESC, Value, response
                foreach (DataRow item in relCreation.Rows)
                {
                    if (item[5].ToString() == "")
                    {
                        log.LogDeCambios("", root.BDProcess, "automatico", "Creación de releases Control desk", item[0].ToString(), item[1].ToString());
                        respFinal = respFinal + "\\n" + "Creación de releases Control desk: " + item[0].ToString() + " " + item[1].ToString();
                    }
                }
            }
            #endregion

            console.WriteLine("Actualizando Releases: " + mandCd);

            #region Actualizar fechas de Releases

            DataTable relUpdate = new DataTable();
            if (dsRenewals.Tables["ContractsNewDates"].Rows.Count > 0) // hay contratos validos
            {
                relUpdate = UpdateReleases(dsRenewals.Tables["ContractsNewDates"]);
                foreach (DataRow item in relUpdate.Rows)
                {
                    log.LogDeCambios("", root.BDProcess, "automatico", "Actualización de releases Control desk", string.Join(" | ", item.ItemArray.ToList()), "");
                    respFinal = respFinal + "\\n" + "Actualización de releases Control desk: " + string.Join(" | ", item.ItemArray.ToList());
                }
            }
            #endregion

            console.WriteLine("Enviando notificaciones: " + mandCd);

            #region Enviar Notificaciones

            string emailBody = "";

            foreach (DataRow fila in resConCreation.Rows)
            {
                try
                {
                    if (((DataTable)fila["RESPONSE_EQUI"]).Rows.Count > 0)
                        emailBody += val.ConvertDataTableToHTML(((DataTable)fila["RESPONSE_EQUI"]));
                }
                catch (Exception) { }
            }
            foreach (DataRow fila in resOnHoldCreation.Rows)
            {
                try
                {
                    if (((DataTable)fila["RESPONSE_EQUI"]).Rows.Count > 0)
                        emailBody += val.ConvertDataTableToHTML(((DataTable)fila["RESPONSE_EQUI"]));
                }
                catch (Exception) { }
            }
            if (emailBody != "")
                emailBody = "<b>LOG DE EQUIPOS:</b><br><br>" + emailBody + "<br>";

            try { resConCreation.Columns.Remove("RESPONSE_EQUI"); } catch (Exception) { }
            emailBody += "<b>LOG DE NUEVOS CONTRATOS:</b><br><br>" + val.ConvertDataTableToHTML(resConCreation) + "<br>";
            emailBody += "<b>LOG DE RENOVACIONES:</b><br><br>" + val.ConvertDataTableToHTML(updatedConDates) + "<br>";

            if (updateItems.Rows.Count > 0)
                emailBody += "<b>LOG DE ACTUALIZACIÓN ITEMS:</b><br><br>" + val.ConvertDataTableToHTML(updateItems) + "<br><br>";

            emailBody += "<b>LOG DE APROBACIONES:</b><br><br>" + apprCon.Replace("\n", "<br>") + "<br><br>";
            emailBody += "<b>LOG CONTRATOS FINALIZADOS:</b><br><br>" + closeCon.Replace("\n", "<br>") + "<br><br>";

            try { resConCreation.Columns.Remove("RESPONSE_EQUI"); } catch (Exception) { }
            if (resOnHoldCreation.Rows.Count > 0)
                emailBody += "<b>LOG DE NUEVOS CONTRATOS ONHOLD:</b><br><br>" + val.ConvertDataTableToHTML(resOnHoldCreation) + "<br><br>";

            if (relCreation.Rows.Count > 0)
            {
                emailBody += "<b>LOG DE CREACION DE RELEASES:</b><br><br>" + val.ConvertDataTableToHTML(relCreation) + "<br>";
                SendReleasesCreationMail(relCreation, mandCd, sDate);
            }
            if (relUpdate.Rows.Count > 0)
            {
                emailBody += "<b>LOG DE ACTUALIZACIÓN DE RELEASES:</b><br><br>" + val.ConvertDataTableToHTML(relUpdate) + "<br><br>";
                SendReleasesUpdateMail(relUpdate, mandCd, sDate);
            }


            if (closedConsWithCis.Count > 0)
            {
                string sendRes = SendCisMail(closedConsWithCis);
                if (sendRes != "OK")
                    emailBody += "<b>LOG DE NOTIFICACION DE CIS DE CONTRATOS CLOSED:</b><br><br>" + sendRes + "<br>";
            }

            mail.SendHTMLMail(emailBody, new string[] { "internalcustomersrvs@gbm.net" }, "**" + mandCd + "** Creación de Contratos en CD " + sDate, new string[] { "smarin@gbm.net" });

            #endregion
        }

        /// <summary>
        /// Envía las notificaciones de los Releases que cambiaron de fechas en CD
        /// </summary>
        /// <param name="relUpdate"></param>
        /// <param name="mandCd"></param>
        /// <param name="sDate"></param>
        private void SendReleasesUpdateMail(DataTable relUpdate, string mandCd, string sDate)
        {
            DataView view = new DataView(relUpdate);
            DataTable distinctValues = view.ToTable(true, "Owner");

            foreach (DataRow item in distinctValues.Rows)
            {
                string owner = item["Owner"].ToString();
                if (owner != "")
                {
                    DataTable ownerDt = relUpdate.Select("Owner = '" + owner + "'").CopyToDataTable();
                    mail.SendHTMLMail("<b>LOG DE ACTUALIZACION DE RELEASES:</b><br><br>" + val.ConvertDataTableToHTML(ownerDt) + "<br>", new string[] { owner }, "**" + mandCd + "** Actualización de Releases en CD " + sDate);
                }
            }

            string[] senders = { "grios@gbm.net", "larias@gbm.net" };
            mail.SendHTMLMail("<b>LOG DE ACTUALIZACION DE RELEASES:</b><br><br>" + val.ConvertDataTableToHTML(relUpdate) + "<br>", senders, "**" + mandCd + "** Actualización de Releases en CD " + sDate);
        }

        /// <summary>
        /// Actualiza los Releases cuando un contrato cambia de fecha de inicio
        /// </summary>
        /// <param name="dtRenewals"></param>
        /// <returns></returns>
        private DataTable UpdateReleases(DataTable dtRenewals)
        {
            string relCol = "Release";
            string conCol = "Contrato";
            string ownerCol = "Owner";
            string dateCol = "Nueva Fecha Final";
            string resCol = "Respuesta";
            DataTable ret = new DataTable();

            ret.Columns.Add(relCol);
            ret.Columns.Add(conCol);
            ret.Columns.Add(dateCol);
            ret.Columns.Add(ownerCol);
            ret.Columns.Add(resCol);

            DataRow[] selectRes = dtRenewals.Select("APPT_TYPE = 'CONTSTART'");
            DataTable newStartDateCons = selectRes.Count() > 0 ? selectRes.CopyToDataTable() : new DataTable();

            foreach (DataRow newStartDateCon in newStartDateCons.Rows)
            {
                string con = newStartDateCon["OBJECT_ID"].ToString();
                try
                {
                    string eDate = DateTime.ParseExact(newStartDateCon["VALID_TO"].ToString(), "yyyy-MM-dd", null).AddDays(-1).ToString("yyyy-MM-dd");  //un día antes del con start
                    List<CdReleaseData> releasesData = cdi.GetContractReleaseData(con);// 8030019755  8030018583  8030019568 8030019836  8030020581   ejemplo de PRD

                    if (releasesData.Count > 0)
                        foreach (CdReleaseData releaseData in releasesData)
                        {
                            DataRow retRow = ret.NewRow();
                            retRow[conCol] = con;
                            retRow[relCol] = releaseData.relId;
                            retRow[ownerCol] = releaseData.Owner;
                            retRow[dateCol] = eDate;
                            retRow[resCol] = cdi.UpdateReleaseTargCompDate(releaseData.relId, eDate);
                            ret.Rows.Add(retRow);
                        }
                }
                catch (Exception ex)
                {
                    DataRow retRow = ret.NewRow();
                    retRow[conCol] = con;
                    retRow[resCol] = ex.Message;
                    ret.Rows.Add(retRow);
                }
            }
            return ret;
        }

        /// <summary>
        /// Envía las notificaciones de los contratos cerrados que tenían CIs asignados
        /// </summary>
        /// <param name="closedConsWithCis">lista con los ids de los contratos</param>
        private string SendCisMail(List<CdContractData> closedConsWithCis)
        {
            string ret = "OK";
            try
            {
                List<string> allCis = new List<string>();
                List<string> allOwners = new List<string>();

                string emailBodyDefault = "Se le informa que los siguientes contratos finalizaron y tenían CIs asignados:<br><br>";

                DataColumn cons = new DataColumn("Contrato");
                DataColumn cis = new DataColumn("Configuration Item");
                DataColumn owners = new DataColumn("Responsable");

                DataTable dt = new DataTable();
                dt.Columns.Add(cons);
                dt.Columns.Add(cis);
                dt.Columns.Add(owners);

                #region obtener la info de todo los CIs en una sola consulta para ahorrar tiempo
                foreach (CdContractData closedCon in closedConsWithCis)
                    foreach (string ci in closedCon.CisArray)
                        allCis.Add(ci);

                string cisXml = cdi.GetCisXml(allCis);
                List<CdConfigurationItemData> cisData = cdi.ParseCiXml(cisXml);
                #endregion

                #region obtener los cons-cis-owners
                foreach (CdContractData closedCon in closedConsWithCis)
                {
                    foreach (string ci in closedCon.CisArray)
                    {
                        foreach (CdConfigurationItemData ciData in cisData)
                        {
                            if (ciData.CiNum == ci)
                            {
                                DataRow dr = dt.NewRow();
                                dr[cons.ColumnName] = closedCon.IdContract;
                                dr[cis.ColumnName] = ci;
                                dr[owners.ColumnName] = ciData.PersonId;
                                allOwners.Add(ciData.PersonId);
                                dt.Rows.Add(dr);
                            }
                        }
                    }
                }
                #endregion

                #region Enviar los correos
                allOwners = allOwners.Distinct().ToList();

                foreach (string owner in allOwners)
                {
                    DataRow[] rowRes = dt.Select(owners.ColumnName + " = '" + owner + "'");
                    string emailBody = emailBodyDefault + val.ConvertDataTableToHTML(rowRes.CopyToDataTable());
                    mail.SendHTMLMail(emailBody, new string[] { owner }, "Contratos Finalizados con CIs", new string[] { "smarin@gbm.net" });
                }
                #endregion
            }
            catch (Exception ex)
            {
                ret = ex.Message;
            }

            return ret;
        }

        /// <summary>
        /// Obtiene los contratos(con sus datos) que contienen CIs
        /// </summary>
        /// <param name="closedContracts">lista de ids de contratos</param>
        /// <returns></returns>
        private List<CdContractData> GetContractsWithCis(List<string> closedContracts)
        {
            List<CdContractData> ret = new List<CdContractData>();
            string allConsData = cdi.GetContractsXml(closedContracts);
            List<CdContractData> consData = cdi.ParseContractXml(allConsData);

            if (allConsData.Contains("PLUSPAPPLCI"))
            {
                foreach (CdContractData conData in consData)
                {
                    List<string> conCis = conData.CisArray.Distinct().ToList();

                    if (conCis.Count > 0)
                        ret.Add(conData);
                }
            }
            return ret;
        }

        /// <summary>
        /// Envía las notificaciones de los Releases creados
        /// </summary>
        /// <param name="relCreation">Cualquier DataTable con el resultado del proceso de creación de releases</param>
        /// <param name="mandCd">mandante de CD en el cual se aplicó el proceso, para colocarlo en el subject del correo</param>
        /// <param name="sDate">fecha en la cual se aplicó el proceso, para colocarlo en el subject del correo </param>
        private void SendReleasesCreationMail(DataTable relCreation, string mandCd, string sDate)
        {
            string[] cc = cdi.GetPersonGroupPeople("GBMBPTPMO");
            string[] senders = { "grios@gbm.net", "AArguedas@gbm.net", "LArias@gbm.net" };
            mail.SendHTMLMail("<b>LOG DE CREACION DE RELEASES:</b><br><br>" + val.ConvertDataTableToHTML(relCreation) + "<br>", senders, "**" + mandCd + "** Creación de Releases en CD " + sDate, cc);
        }

        /// <summary>
        /// Carga a Cd los equipos
        /// </summary>
        /// <param name="equipments">???????????</param>
        /// <param name="customerName">Nombre del Cliente</param>
        /// <param name="customerId">Id del cliente</param>
        /// <param name="matGroups">?????????</param>
        /// <returns></returns>
        private DataTable UploadNewEquips(DataTable equipments, string customerName, string customerId, DataTable matGroups)
        {
            string equipCol = "EQUIP";
            string responseCol = "RESPONSE";

            DataTable ret = new DataTable();
            ret.Columns.Add(equipCol);
            ret.Columns.Add(responseCol);

            if (equipments.Rows.Count != 0)
            {
                foreach (DataRow equipment in equipments.Select("EQ_SER <> ''"))
                {
                    DataRow retRow = ret.NewRow();

                    string materialGroup = matGroups.Select("ITEM_PROD = '" + equipment["ITEM_PROD"].ToString().Trim() + "'")[0]["MG"].ToString();

                    if (materialGroup == "30103")
                    {
                        console.WriteLine("Crea Location y Asset");

                        string assetNum = equipment["EQ_ID"].ToString();
                        if (ret.Select(equipCol + " = " + assetNum).Count() == 0)
                        {
                            CdAssetData asset = new CdAssetData
                            {
                                InstallDate = equipment["WARR_START"].ToString().Substring(0, 4) + "-" + equipment["WARR_START"].ToString().Substring(4, 2) + "-" + equipment["WARR_START"].ToString().Substring(6, 2),
                                EndDate = equipment["WARR_END"].ToString().Substring(0, 4) + "-" + equipment["WARR_END"].ToString().Substring(4, 2) + "-" + equipment["WARR_END"].ToString().Substring(6, 2),
                                WarrantyText = val.RemoveSpecialChars(equipment["WARR_DESC"].ToString(), 1).ToUpper(),
                                MaterialNum = equipment["EQ_MAT"].ToString(),
                                AssetText = equipment["EQ_DESC"].ToString(),
                                SerialNum = equipment["EQ_SER"].ToString(),
                                Warranty = equipment["WARR_ID"].ToString(),
                                MaterialGroup = materialGroup,
                                Location = customerId,
                                AssetNum = assetNum,
                                Placa = ""
                            };

                            CdLocationData loc = new CdLocationData
                            {
                                Location = customerId,
                                Description = customerName
                            };

                            if (cdi.CreateLocation(loc) != "OK")
                            {
                                retRow[equipCol] = assetNum;
                                retRow[responseCol] = "Error al crear Location";
                            }
                            else //crea los asset
                            {
                                if (cdi.CreateAsset(asset) == "OK")
                                {
                                    retRow[equipCol] = assetNum;
                                    retRow[responseCol] = "Equipo creado con éxito";
                                }
                                else
                                {
                                    retRow[equipCol] = assetNum;
                                    retRow[responseCol] = "Error equipo no creado";
                                }
                            }
                            ret.Rows.Add(retRow);
                        }
                    }
                }
            }
            return ret;
        } //terminar el summary

        /// <summary>
        /// Obtiene de SAP, los contratos que finalizaron en un rango de fechas
        /// </summary>
        /// <param name="mandCrm">mandante de SAP donde hacer la consulta</param>
        /// <param name="sDate">Fecha inicial del rango</param>
        /// <param name="eDate">Fecha final del rango</param>
        /// <returns></returns>
        private List<string> GetClosedContracts(int mandCrm, string sDate, string eDate)
        {
            List<string> closeCon = new List<string>();

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FECHA_INI"] = sDate,
                ["FECHA_FIN"] = eDate
            };

            IRfcFunction func = sap.ExecuteRFC("CRM", "ZDM_GET_CONTRACT_RENEWAL", parameters, mandCrm);


            IRfcTable response = func.GetTable("RESPONSE_STATUS");

            if (response.RowCount != 0)
            {
                //tabla con resultados, buscar 
                for (int i = 0; i < response.RowCount; i++)
                {
                    string con = response[i].GetValue("OBJECT_ID").ToString();
                    string stat = response[i].GetValue("STAT").ToString();

                    if ((stat == "I1005" || stat == "E0013") && (!(con.StartsWith("801") || con.StartsWith("805")))) //Dejar solo los completed or Canceled y que no sean 801 ni 805
                        closeCon.Add(con);
                }
            }

            return closeCon;
        }

        /// <summary>
        /// Obtiene de SAP los contratos que pasaron a onHold en un rango de fechas
        /// </summary>
        /// <param name="mandCrm">mandante de SAP CRM donde hacer la consulta</param>
        /// <param name="mandErp">mandante de SAP ERP donde hacer la consulta</param>
        /// <param name="startDate">Fecha inicial del rango</param>
        /// <param name="endDate">Fecha final del rango</param>
        /// <returns></returns>
        private DataTable GetNewOnHoldContracts(int mandCrm, int mandErp, string startDate, string endDate)
        {
            console.WriteLine("Leyendo nuevos contratos On Hold");

            DataTable result;
            List<string> onHold = new List<string>();

            RfcDestination destination = sap.GetDestRFC("CRM", mandCrm);
            RfcRepository repo = destination.Repository;

            #region Tomar los contratos que pasaron a status E0011 (OnHold)

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FECHA_INI"] = startDate,
                ["FECHA_FIN"] = endDate
            };

            IRfcFunction func = sap.ExecuteRFC("CRM", "ZDM_GET_CONTRACT_RENEWAL", parameters, mandCrm);


            DataTable fmResponse = sap.GetDataTableFromRFCTable(func.GetTable("RESPONSE_STATUS"));

            if (fmResponse.Rows.Count > 0)
            {
                foreach (DataRow con in fmResponse.Rows)
                {
                    if (con["STAT"].ToString() == "E0011")
                        onHold.Add(con["GUID"].ToString());
                }
            }
            onHold = onHold.Distinct().ToList();
            #endregion

            #region Tomar la mayoría de la info de esos contratos
            func = repo.CreateFunction("ZICS_GET_CONTRACT_DATA");
            IRfcTable zContract = func.GetTable("ZCONTRACT");

            foreach (string guid in onHold)
            {
                zContract.Append();
                zContract.SetValue("GUID", guid);
            }

            func.Invoke(destination);
            IRfcTable sapResult = func.GetTable("RESPONSE");
            result = sap.GetDataTableFromRFCTable(sapResult);
            #endregion

            #region Tomar la fecha del cambio a OnHold
            result.Columns.Add("ON_HOLD_DATE");
            foreach (DataRow row in result.Rows)
            {
                string onHoldDateRow = "";
                try { onHoldDateRow = fmResponse.Select("OBJECT_ID = " + row["CONTRACT"].ToString() + " AND STAT = 'E0011'")[0]["UDATE"].ToString(); } catch (Exception) { }
                row["ON_HOLD_DATE"] = onHoldDateRow;
            }
            #endregion

            #region Eliminar contratos con el MG no valido
            for (int i = result.Rows.Count - 1; i >= 0; i--)
            {
                DataTable dt = sap.GetDataTableFromRFCTable(sapResult[i]["ITEMS"].GetTable());
                dt = GetMaterialGroup(dt, mandErp);


                //if (dt.Select("MG IN ('301010702')").Length > 0)
                //{
                //el contrato tiene 301010702

                List<string> validMgsList = new List<string> {
                          "301010402" ,
                          "3010106"   ,
                          "301010702" ,
                          "30104"     ,
                          "3010101"   ,
                          "3010102"   ,
                          "3010103"   ,
                          "3010104"   ,
                          "301010402" ,
                          "301010701" ,
                          "301010705" ,
                          "3010105"   ,
                          "3010401"   ,
                          "3040603"   ,
                          "40301"     ,
                          "30405"     ,
                          "30406"
                    };

                foreach (DataRow row in dt.Rows)
                {
                    string mg = row["MG"].ToString();
                    if (!validMgsList.Contains(mg))
                    {
                        result.Rows[i].Delete();
                        break;
                    }
                }

                //}
                //else
                //    result.Rows[i].Delete();
            }
            result.AcceptChanges();

            #endregion

            #region Tomar el Sales Rep ya que no viene en la info de arriba y el Nombre del contacto porque lo que viene es el ID

            if (result.Rows.Count > 0)
            {
                result.Columns.Add("SALES_REP");
                result.Columns.Add("EXTERNAL_REFERENCE");
                result.Columns.Add("CONTACT_NAME");
                result.Columns.Add("STATUS");
                result.Columns.Add("NET_VALUE");

                foreach (DataRow row in result.Rows)
                {
                    ////esta FM dura una eternidad  cambiarla cuando regrese DEV y QAS//
                    //func = repo.CreateFunction("ZDM_CONTRACT_READ");                //
                    //func.SetValue("DOCUMENTO", row["CONTRACT"].ToString());         //
                    //func.Invoke(destination);                                       //
                    ////////////////////////////////////////////////////////////////////


                    ////esta FM dura una eternidad  cambiarla cuando regrese DEV y QAS//////////////////////////
                    func = repo.CreateFunction("CRM_SERVICE_CONTRACTS_SEARCH");                               //
                    func.SetValue("OBJECT_ID", row["CONTRACT"].ToString());                                   //
                    func.Invoke(destination);                                                                 //
                    DataTable conInfo = sap.GetDataTableFromRFCTable(func.GetTable("SERVICE_CONTRACT_LIST")); //
                    ////////////////////////////////////////////////////////////////////////////////////////////

                    row["EXTERNAL_REFERENCE"] = conInfo.Rows[0]["PO_NUMBER_SOLD"].ToString();
                    row["STATUS"] = conInfo.Rows[0]["CONCATSTATUSER"].ToString();
                    row["NET_VALUE"] = conInfo.Rows[0]["NET_VALUE"].ToString();

                    func = repo.CreateFunction("ZDM_READ_BP");
                    func.SetValue("BP", conInfo.Rows[0]["PERSON_RESP"].ToString());
                    func.Invoke(destination);
                    row["SALES_REP"] = func.GetValue("EMAIL").ToString();

                    func.SetValue("BP", row["CONTACT"].ToString());
                    func.Invoke(destination);
                    row["CONTACT_NAME"] = func.GetValue("FIRSTNAME").ToString() + " " + func.GetValue("LASTNAME").ToString();
                }
            }
            #endregion

            #region Eliminar los documentos (801,805) ya que no deben ir a CD(lo ideal seria hacerlo en la FM pero bueno)
            for (int i = result.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = result.Rows[i];
                string docNum = row["CONTRACT"].ToString();
                if (docNum.StartsWith("801") || docNum.StartsWith("805"))
                    result.Rows.Remove(row);
            }
            #endregion

            return result;
        } //OK

        /// <summary>
        /// Carga a CD los contratos que han tenido cambios en los items
        /// </summary>
        /// <param name="mandErp">mandante de SAP ERP donde hacer la consulta</param>
        /// <param name="dsRenewals">??????????????</param>
        /// <param name="contractsAlreadyProcessed">Lista de contratos a ignorar(porque ya se hicieron en la creación)</param>
        /// <returns></returns>
        private DataTable RenewItems(int mandErp, DataSet dsRenewals, List<string> contractsAlreadyProcessed)
        {
            try
            {
                List<string> conItems = new List<string>();
                string contractCol = "CONTRACT";
                string responseCol = "RESPONSE";

                DataTable ret = new DataTable();
                ret.Columns.Add(contractCol);
                ret.Columns.Add(responseCol);

                DataTable responseItems = dsRenewals.Tables["ContractsNewItems"];

                #region Dejar solo los contratos que no han sido procesados
                foreach (DataRow contract in responseItems.Rows)
                {
                    if (!contractsAlreadyProcessed.Contains(contract["OBJECT_ID"].ToString()))
                        conItems.Add(contract["OBJECT_ID"].ToString());
                }

                conItems = conItems.Distinct().ToList();
                #endregion

                if (conItems.Count != 0)
                {
                    foreach (string con in conItems)
                    {
                        DataRow tempRow = ret.NewRow();
                        Dictionary<string, string> serviceGroups = new Dictionary<string, string>();

                        string status = cdi.GetContractStatus(con);
                        if (!(status == "NE" || status == "REVISED"))
                        {
                            try
                            {
                                if (status == "APPR")
                                {
                                    #region 1- Si hay items vencidos eliminarlos de la lista, y si hay "ADD" agregarlo a la lista
                                    serviceGroups = GetItemsAddedToContractToday(responseItems, con);
                                    #endregion

                                    #region 2- Buscar el MG de los service groups
                                    serviceGroups = GetServiceGroupsMgs(serviceGroups, mandErp);
                                    #endregion

                                    #region 3- Leer el contrato actual de CD (contrato = con y status active)

                                    string infoXml = cdi.GetContractsXml(new List<string> { con });

                                    CdContractData conInfo = cdi.ParseContractXml(infoXml)[0];

                                    List<string> currentConMgs = conInfo.MaterialArray;
                                    List<string> currentConManualMgs = conInfo.ManualServiceArray;

                                    XmlDocument outXml = new XmlDocument();

                                    #region Crear XML de actualizar
                                    infoXml = infoXml.Replace("><", ">\r\n<");
                                    string[] xmlLines = infoXml.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                    string[] addChange = { "<PLUSPAGREEMENT>", "<MXCUSTAGREEMENTSet>", "<PLUSPPRICESCHED>", "<PLUSPAPPLSERV>", "<PLUSPAPPLASSET>" };
                                    string[] removes = { "RENEWALDATE", "STATUS>", "PLUSPAPPLSERVID>", "GBMGUID>", "SANUM>", "OWNERID>", "OWNERTABLE>" };

                                    for (int k = 0; k < xmlLines.Length; k++)
                                    {
                                        for (int j = 0; j < addChange.Length; j++)
                                        {
                                            if (xmlLines[k].Contains(addChange[j]))
                                                xmlLines[k] = xmlLines[k].Replace(addChange[j], addChange[j].Substring(0, addChange[j].Length - 1) + @" action=""AddChange"">");
                                        }

                                        for (int j = 0; j < removes.Length; j++)
                                        {
                                            if (xmlLines[k].Contains(removes[j]))
                                                xmlLines[k] = "";
                                        }

                                        if (xmlLines[k].Contains("<QueryMXCUSTAGREEMENTResponse"))
                                            xmlLines[k] = @"<SyncMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""	xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" >";
                                        if (xmlLines[k].Contains("</QueryMXCUSTAGREEMENTResponse>"))
                                            xmlLines[k] = "</SyncMXCUSTAGREEMENT>";
                                    }
                                    string outXmlS = string.Join(Environment.NewLine, xmlLines);
                                    #endregion

                                    outXml.LoadXml(outXmlS);
                                    #endregion

                                    #region 4- Construir el nuevo, sumando 1 a la version y colocando la nueva fecha
                                    XmlNodeList tempList = outXml.GetElementsByTagName("REVISIONNUM");
                                    int nextRev = int.Parse(tempList[tempList.Count - 1].InnerText) + 1;
                                    tempList[tempList.Count - 1].InnerText = nextRev.ToString();
                                    #endregion

                                    #region 5- Eliminar items vencidos
                                    List<string> sg = new List<string>();

                                    foreach (var item in serviceGroups)
                                        sg.Add(item.Value);

                                    XmlNodeList deleteNodes = outXml.GetElementsByTagName("PLUSPAPPLSERV");
                                    for (int j = deleteNodes.Count - 1; j >= 0; j--)
                                    {
                                        if (!currentConManualMgs.Contains(deleteNodes[j]["COMMODITY"].InnerText))//si no es commodity manual
                                        {
                                            if (!sg.Contains(deleteNodes[j]["COMMODITY"].InnerText))//si el commodity(CD) no esta en la lista SG(SAP), pues se borra
                                                deleteNodes[j].ParentNode.RemoveChild(deleteNodes[j]);
                                        }
                                    }
                                    #endregion

                                    #region 6- Insertar nuevos items
                                    List<string> cdServices = new List<string>();
                                    foreach (XmlNode item in outXml.GetElementsByTagName("COMMODITY"))
                                        cdServices.Add(item.InnerText);

                                    cdServices = cdServices.Distinct().ToList(); //servicios del contrato en CD
                                    foreach (string item in sg) //for del array de servicios de SAP
                                    {
                                        if (!cdServices.Contains(item)) // si el item NO esta en el array de CD agregarlo
                                        {
                                            XmlNodeList pluspPriceScheds = outXml.GetElementsByTagName("PLUSPPRICESCHED");
                                            foreach (XmlNode pluspPriceSched in pluspPriceScheds)
                                            {
                                                XmlElement pluspApplServ = outXml.CreateElement("PLUSPAPPLSERV", "http://www.ibm.com/maximo");
                                                pluspApplServ.SetAttribute("action", "AddChange");

                                                XmlElement commodity = outXml.CreateElement("COMMODITY", "http://www.ibm.com/maximo");
                                                XmlElement itemsetId = outXml.CreateElement("ITEMSETID", "http://www.ibm.com/maximo");
                                                XmlElement ownerTable = outXml.CreateElement("OWNERTABLE", "http://www.ibm.com/maximo");

                                                commodity.InnerText = item;
                                                itemsetId.InnerText = "ITEMSET1";
                                                ownerTable.InnerText = "PLUSPPRICESCHED";

                                                pluspApplServ.AppendChild(commodity);
                                                pluspApplServ.AppendChild(itemsetId);
                                                pluspApplServ.AppendChild(ownerTable);

                                                pluspPriceSched.AppendChild(pluspApplServ);
                                            }
                                        }
                                    }

                                    outXmlS = outXml.InnerXml;
                                    #endregion

                                    #region 7- Post to  CD

                                    List<string> UpdatedConMgs = cdi.ParseContractXml(outXmlS)[0].MaterialArray;

                                    if (!currentConMgs.OrderBy(x => x).SequenceEqual(UpdatedConMgs.OrderBy(x => x))) //verificar si listas tienen los mismos valores
                                    {
                                        console.WriteLine("Renovando ITEMS contrato: " + con);
                                        string resXml = cdi.PostCD(root.UrlCd, "MXCUSTAGREEMENT", outXmlS); //cons con nuevo items
                                        if (resXml.Contains("PLUSPAGREEMENTID"))
                                        {
                                            tempRow[contractCol] = con;
                                            tempRow[responseCol] = "OK: (" + String.Join(", ", currentConMgs.ToArray()) + ")->(" + String.Join(", ", UpdatedConMgs.ToArray()) + ")";
                                        }
                                        else
                                        {
                                            tempRow[contractCol] = con;
                                            tempRow[responseCol] = "Error actualizando items en el contrato: " + resXml;
                                        }
                                    }
                                    else
                                        continue;

                                    #endregion
                                }
                                else
                                {
                                    tempRow[contractCol] = con;
                                    tempRow[responseCol] = "No se actualizaron los items, el contrato " + status;
                                }
                            }
                            catch (Exception ex)
                            {
                                tempRow[contractCol] = con;
                                tempRow[responseCol] = "No se actualizaron los items, el contrato dio ERROR:" + ex.Message;
                            }
                        }
                        else
                        {
                            tempRow[contractCol] = con;
                            tempRow[responseCol] = "No se actualizaron los items, el contrato " + status;
                        }
                        ret.Rows.Add(tempRow);
                    }
                }
                else
                    ret.TableName = "No se actualizaron items hoy";

                if (ret.Rows.Count == 0)
                    ret.TableName = "Ningún contrato en CD";

                return ret;
            }
            catch (Exception ex)
            {
                DataTable errorDt = new DataTable();
                errorDt.Columns.Add("ERROR", typeof(string));
                errorDt.Rows.Add();
                errorDt.Rows[0][0] = ex.Message;
                return errorDt;
            }
        } //terminar el summary

        /// <summary>
        /// Obtiene de SAP el material group del los materiales dados
        /// </summary>
        /// <param name="serviceGroups">Diccionario con los materiales en los KEYS</param>
        /// <param name="mandErp">mandante de SAP ERP donde hacer la consulta</param>
        /// <returns>Un diccionario con los materiales(KEY) y material group(VALUE)</returns>
        private Dictionary<string, string> GetServiceGroupsMgs(Dictionary<string, string> serviceGroups, int mandErp)
        {
            DataTable itemsDt = new DataTable();
            itemsDt.Columns.Add("ITEM_PROD");

            foreach (KeyValuePair<string, string> itemAndServiceGroup in serviceGroups)
            {
                DataRow retRow = itemsDt.NewRow();
                retRow["ITEM_PROD"] = itemAndServiceGroup.Key;
                itemsDt.Rows.Add(retRow);
            }

            itemsDt = GetMaterialGroup(itemsDt, mandErp);

            serviceGroups = new Dictionary<string, string>();

            foreach (DataRow itemRow in itemsDt.Rows)
                serviceGroups.Add(itemRow[0].ToString(), itemRow[1].ToString());

            return serviceGroups;
        }

        /// <summary>
        /// Obtiene de SAP el material group del los materiales dados
        /// </summary>
        /// <param name="equipsAndItems">????????????????</param>
        /// <param name="mandErp">mandante de SAP ERP donde hacer la consulta</param>
        /// <returns></returns>
        private DataTable GetMaterialGroup(DataTable equipsAndItems, int mandErp)
        {
            DataTable ret = equipsAndItems.Copy();

            Dictionary<string, string> dict = new Dictionary<string, string>();
            ret.Columns.Add("MG");

            RfcDestination destErp = sap.GetDestRFC("ERP", mandErp);
            IRfcFunction fmMg = destErp.Repository.CreateFunction("RFC_READ_TABLE");
            fmMg.SetValue("USE_ET_DATA_4_RETURN", "X");
            fmMg.SetValue("QUERY_TABLE", "MARA");
            fmMg.SetValue("DELIMITER", "|");

            IRfcTable fields = fmMg.GetTable("FIELDS");
            fields.Append();
            fields.SetValue("FIELDNAME", "MATNR");
            fields.Append();
            fields.SetValue("FIELDNAME", "MATKL");

            IRfcTable fmOptions = fmMg.GetTable("OPTIONS");
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "MATNR IN ( ");

            foreach (DataRow item in ret.Rows)
            {
                string prod = item["ITEM_PROD"].ToString();

                fmOptions.Append();
                fmOptions.SetValue("TEXT", "'" + prod + "',");

            }
            fmOptions.Append();
            fmOptions.SetValue("TEXT", "'' )");
            fmMg.Invoke(destErp);

            //llenar diccionario
            foreach (DataRow fila in sap.GetDataTableFromRFCTable(fmMg.GetTable("ET_DATA")).Rows)
            {
                string prod = fila["LINE"].ToString().Split(new char[] { '|' })[0].Trim();
                string mg = fila["LINE"].ToString().Split(new char[] { '|' })[1].Trim();
                try { dict.Add(prod, mg); } catch (Exception) { }
            }

            //actualizar tabla
            foreach (DataRow fila in ret.Rows)
            {
                fila["MG"] = dict[fila["ITEM_PROD"].ToString()];
            }

            return ret;
        } //terminar el summary

        /// <summary>
        /// Obtiene los items que se añadieron el dia actual al contrato
        /// </summary>
        /// <param name="responseItems"></param>
        /// <param name="con"></param>
        /// <returns>Diccionario con los materiales(KEY) y el VALUE vacío</returns>
        private Dictionary<string, string> GetItemsAddedToContractToday(DataTable responseItems, string con)
        {
            Dictionary<string, string> newItems = new Dictionary<string, string>();

            foreach (DataRow item in responseItems.Select("OBJECT_ID = '" + con + "'"))
            {
                if (DateTime.Parse(item["END_DATE"].ToString()) > DateTime.Today)
                    try { newItems.Add(item["ITEM"].ToString(), ""); } catch (Exception) { }
            }

            return newItems;
        }

        /// <summary>
        /// Carga a CD los contratos nuevos
        /// </summary>
        /// <param name="newContracts">??????</param>
        /// <param name="mandErp">mandante de SAP ERP donde hacer la consulta</param>
        /// <param name="mandCrm">mandante de SAP CRM donde hacer la consulta</param>
        /// <returns></returns>
        private DataTable UploadNewContracts(DataTable newContracts, int mandErp, int mandCrm)
        {
            try
            {
                string contractCol = "CONTRACT";
                string responseCol = "RESPONSE";
                string responseEquiCol = "RESPONSE_EQUI";

                DataTable ret = new DataTable();
                ret.Columns.Add(contractCol);
                ret.Columns.Add(responseCol);
                ret.Columns.Add(responseEquiCol, typeof(DataTable));

                RfcDestination destCrm = sap.GetDestRFC("CRM", mandCrm);

                #region Procesar contratos nuevos

                foreach (DataRow newContract in newContracts.Rows)
                {
                    DataRow retRow = ret.NewRow();

                    string contractId = newContract[contractCol].ToString();

                    if (cdi.GetContractStatus(contractId) == "NE") //Verifica si el contrato ya existe en CD
                    {
                        #region Extraer los datos del contrato

                        string emailContact = "", contactName = "", contactLasName = "", contactTel = "", contactId = "", customerCountry;

                        List<string> eqArray = new List<string>();
                        List<string> matArray = new List<string>();

                        #region Tomar la info que viene de la tabla newContracts

                        string conDesc = newContract["CONTRACT_DESC"].ToString();
                        string customerId = newContract["CUSTOMER"].ToString();
                        string customerName = newContract["CUSTOMER_DESC"].ToString();
                        string sDate = newContract["CON_START"].ToString().Substring(0, 4) + "-" + newContract["CON_START"].ToString().Substring(4, 2) + "-" + newContract["CON_START"].ToString().Substring(6, 2);
                        string eDate = newContract["CON_END"].ToString().Substring(0, 4) + "-" + newContract["CON_END"].ToString().Substring(4, 2) + "-" + newContract["CON_END"].ToString().Substring(6, 2);
                        string idContact = newContract["CONTACT"].ToString();

                        #endregion

                        #region Tomar los equipos y material groups

                        DataTable equipments = new DataTable();
                        DataTable items = new DataTable();
                        try { equipments = (DataTable)newContract["EQUI"]; } catch (Exception) { }
                        try { items = (DataTable)newContract["ITEMS"]; } catch (Exception) { }

                        #endregion

                        #region Tomar los MG desde SAP (RFC_READ_TABLE)
                        DataTable itemsMgDt = GetMaterialGroup(items, mandErp);
                        #endregion

                        #region Tomar info de contacto desde SAP
                        if (idContact != "")
                        {
                            IRfcFunction fmReadContactData = destCrm.Repository.CreateFunction("ZDM_READ_BP");
                            fmReadContactData.SetValue("BP", idContact);
                            fmReadContactData.Invoke(destCrm);

                            emailContact = fmReadContactData.GetValue("EMAIL").ToString();
                            if (emailContact.Length > 30)
                                contactId = emailContact.Substring(0, 30).ToString().ToUpper();
                            else
                                contactId = emailContact.ToString().ToUpper();

                            contactName = fmReadContactData.GetValue("FIRSTNAME").ToString();
                            contactLasName = fmReadContactData.GetValue("LASTNAME").ToString();
                            contactTel = fmReadContactData.GetValue("PHONE").ToString();
                        }
                        #endregion

                        #region Tomar info de cliente desde SAP (país)
                        IRfcFunction fmGetCountry = destCrm.Repository.CreateFunction("ZDM_READ_BP");
                        fmGetCountry.SetValue("BP", customerId);
                        fmGetCountry.Invoke(destCrm);
                        customerCountry = fmGetCountry.GetValue("PAIS").ToString();
                        #endregion

                        #endregion

                        #region Crear el contrato
                        console.WriteLine("Crea el contacto, cliente y Contrato y sus equipos");

                        //create equipment
                        retRow[responseEquiCol] = UploadNewEquips(equipments, customerName, customerId, itemsMgDt); //Si tiene equipos, crea Location y Assets

                        //create contact
                        CdContactData contact = new CdContactData
                        {
                            PersonId = contactId,
                            FirstName = contactName,
                            LastName = contactLasName,
                            SapId = idContact,
                            Email = emailContact,
                            Telephone = contactTel
                        };
                        string responseContact = cdi.CreateContact(contact);

                        //create customer
                        CdCustomerData customer = new CdCustomerData
                        {
                            PersonId = contactId,
                            IdCustomer = customerId,
                            NameCustomer = customerName,
                            Country = customerCountry
                        };
                        string responseCustomer = cdi.CreateCustomer(customer);

                        //create contract
                        DataTable equipsResponseDt = (DataTable)retRow[responseEquiCol];
                        foreach (DataRow equipResponseRow in equipsResponseDt.Rows)
                        {
                            if (equipResponseRow[1].ToString() == "Equipo creado con éxito")
                            {
                                if (!eqArray.Contains(equipResponseRow[0].ToString()))
                                    eqArray.Add(equipResponseRow[0].ToString());
                            }
                        }

                        matArray = itemsMgDt.Rows.OfType<DataRow>().Select(dr => dr["MG"].ToString()).Distinct().ToList();

                        CdContractData con = new CdContractData
                        {
                            IdContract = contractId,
                            IdCustomer = customerId,
                            Description = conDesc,
                            StartDate = sDate,
                            EndDate = eDate,
                            MaterialArray = matArray,
                            EquipArray = eqArray
                        };

                        string responseContract = cdi.CreateContract(con);
                        #endregion

                        #region Procesar Respuesta

                        if (responseCustomer == "OK" && responseContact == "OK" && responseContract == "OK")
                        {
                            retRow[contractCol] = contractId;
                            retRow[responseCol] = "CONTRATO creado con éxito en Control Desk";
                        }
                        else if (responseCustomer == "OK" && responseContact != "OK" && responseContract == "OK")
                        {
                            retRow[contractCol] = contractId;
                            retRow[responseCol] = "CONTRATO creado con éxito en Control Desk, problemas con el CONTACTO: " + responseContact;
                        }
                        else if (responseCustomer != "OK" && responseContact == "OK" && responseContract == "OK")
                        {
                            retRow[contractCol] = contractId;
                            retRow[responseCol] = "CONTRATO creado con éxito en Control Desk, problemas con el CLIENTE: " + responseCustomer;
                        }
                        else if (responseCustomer != "OK" && responseContact != "OK" && responseContract == "OK")
                        {
                            retRow[contractCol] = contractId;
                            retRow[responseCol] = "CONTRATO creado con éxito en Control Desk, problemas con el CLIENTE: " + responseCustomer + " y problemas con el contacto: " + responseContact;
                        }
                        else
                        {
                            retRow[contractCol] = contractId;
                            retRow[responseCol] = "Problemas con el CONTRATO: " + responseContract + "problemas con el CLIENTE: " + responseCustomer + " y problemas con el CONTACTO: " + responseContact;
                        }

                        #endregion

                    }
                    else
                    {
                        retRow[contractCol] = contractId;
                        retRow[responseCol] = "CONTRATO ya existe en Control Desk";
                    }
                    ret.Rows.Add(retRow);
                }

                #endregion

                return ret;
            }
            catch (Exception ex)
            {
                DataTable errorDt = new DataTable();
                errorDt.Columns.Add("ERROR", typeof(string));
                errorDt.Rows.Add();
                errorDt.Rows[0][0] = ex.Message;
                return errorDt;
            }
        } // terminar el summary

        /// <summary>
        /// Obtiene de SAP los contratos nuevos en un rango de fechas
        /// </summary>
        /// <param name="mandCrm">mandante de SAP CRM donde hacer la consulta</param>
        /// <param name="startDate">Fecha inicial del rango</param>
        /// <param name="endDate">Fecha final del rango</param>
        /// <returns></returns>
        private DataTable GetNewContracts(int mandCrm, string startDate, string endDate)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FECHA_INI"] = startDate,
                ["FECHA_FIN"] = endDate
            };

            IRfcFunction func = sap.ExecuteRFC("CRM", "ZDM_GET_NEW_CONTRACT", parameters, mandCrm);
            IRfcTable response = func.GetTable("RESPONSE");

            DataTable responseTable = sap.GetDataTableFromRFCTable(response);

            for (int i = 0; i < responseTable.Rows.Count; i++)
            {
                responseTable.Rows[i]["EQUI"] = sap.GetDataTableFromRFCTable(response[i].GetTable("EQUI"));
                responseTable.Rows[i]["ITEMS"] = sap.GetDataTableFromRFCTable(response[i].GetTable("ITEMS"));
            }

            #region Eliminar los documentos (801,805) ya que no deben ir a CD(lo ideal seria hacerlo en la FM pero bueno)
            for (int i = responseTable.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = responseTable.Rows[i];
                string docNum = row["CONTRACT"].ToString();
                if (docNum.StartsWith("801") || docNum.StartsWith("805"))
                    responseTable.Rows.Remove(row);
            }
            #endregion

            return responseTable;
        }

        /// <summary>
        /// Obtiene de SAP los contratos que se renovaron en un rango de fechas
        /// </summary>
        /// <param name="mandCrm">mandante de SAP CRM donde hacer la consulta</param>
        /// <param name="sDate">Fecha inicial del rango</param>
        /// <param name="eDate">Fecha final del rango</param>
        /// <returns></returns>
        private DataSet GetRenewalContracts(int mandCrm, string sDate, string eDate)
        {
            DataSet retDs = new DataSet();

            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["FECHA_INI"] = sDate,
                ["FECHA_FIN"] = eDate
            };

            IRfcFunction func = sap.ExecuteRFC("CRM", "ZDM_GET_CONTRACT_RENEWAL", parameters, mandCrm);

            DataTable tempDt = sap.GetDataTableFromRFCTable(func.GetTable("RESPONSE_DATE"));
            tempDt.TableName = "ContractsNewDates";
            retDs.Tables.Add(tempDt.Copy());

            tempDt = sap.GetDataTableFromRFCTable(func.GetTable("RESPONSE_ITEMS"));
            tempDt.TableName = "ContractsNewItems";
            retDs.Tables.Add(tempDt.Copy());

            #region Eliminar los documentos (801,805) ya que no deben ir a CD(lo ideal seria hacerlo en la FM pero bueno)
            foreach (DataTable dt in retDs.Tables)
            {
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow row = dt.Rows[i];
                    string docNum = row["OBJECT_ID"].ToString();
                    if (docNum.StartsWith("801") || docNum.StartsWith("805"))
                        dt.Rows.Remove(row);
                }
            }
            #endregion

            return retDs;
        }

        /// <summary>
        /// Pasa a status close contratos en CD
        /// </summary>
        /// <param name="contractsToClose">lista de contratos a cerrar</param>
        /// <returns></returns>
        private string CloseContracts(List<string> contractsToClose)
        {
            string response;

            try
            {
                if (0 < contractsToClose.Count)
                {
                    #region cerrar contratos
                    try
                    {
                        #region 1. Cerrar contratos
                        try
                        {
                            response = cdSelenium.ChangeStatusCustomerAgreements(contractsToClose, root.UrlCd, "close");
                        }
                        catch (Exception)//si selenium diera error re-intentar otra vez
                        {
                            response = cdSelenium.ChangeStatusCustomerAgreements(contractsToClose, root.UrlCd, "close");
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        #region 2. Si ya no se pudiera esta es la respuesta
                        response = "Error cerrando contratos favor hacerlo manual ('" + String.Join("','", contractsToClose.ToArray()) + "') " + ex.Message;
                        #endregion
                    }

                    #endregion
                }
                else if (contractsToClose.Count == 0)
                    response = "No hay contratos para cerrar";
                else
                    response = "Error cerrando contratos favor hacerlo manual ('" + String.Join("','", contractsToClose.ToArray()) + "')";
            }
            catch (Exception)
            {
                response = "Error cerrando contratos favor hacerlo manual ('" + String.Join("','", contractsToClose.ToArray()) + "')";
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);
            }

            return response;
        }

        /// <summary>
        /// Pasa a status appr contratos en CD
        /// </summary>
        /// <param name="contractsToApprove">lista de contratos a aprobar</param>
        /// <returns></returns>
        private string ApprContracts(List<string> contractsToApprove)
        {
            string response;

            try
            {
                if (0 < contractsToApprove.Count)
                {
                    #region aprobación de contratos
                    try
                    {
                        #region 1. Aprobar contratos
                        try
                        {
                            response = cdSelenium.ChangeStatusCustomerAgreements(contractsToApprove, root.UrlCd, "appr");
                        }
                        catch (Exception)//si selenium diera error re-intentar otra vez
                        {
                            response = cdSelenium.ChangeStatusCustomerAgreements(contractsToApprove, root.UrlCd, "appr");
                        }
                        #endregion

                        #region 2. Detecta si algo quedo en DRAFT, en caso de que hay entonces corre de nuevo 

                        List<string> consStillDraft = new List<string>();
                        foreach (string con in contractsToApprove)
                        {
                            if (cdi.GetContractStatus(con) == "DRAFT")
                                consStillDraft.Add(con);
                        }
                        if (consStillDraft.Count > 0)
                        {
                            try
                            {
                                response = cdSelenium.ChangeStatusCustomerAgreements(consStillDraft, root.UrlCd, "appr");
                            }
                            catch (Exception)
                            {
                                response = cdSelenium.ChangeStatusCustomerAgreements(consStillDraft, root.UrlCd, "appr");
                            }
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        #region 3. Si ya no se pudiera esta es la respuesta
                        response = "Error aprobando contratos favor hacerlo manual ('" + String.Join("','", contractsToApprove.ToArray()) + "') " + ex.Message;
                        #endregion
                    }

                    #endregion
                }
                else if (contractsToApprove.Count == 0)
                    response = "No hay contratos que aprobar";
                else
                    response = "Error aprobando contratos favor hacerlo manual ('" + String.Join("','", contractsToApprove.ToArray()) + "')";
            }
            catch (Exception)
            {
                response = "Error aprobando contratos favor hacerlo manual ('" + String.Join("','", contractsToApprove.ToArray()) + "')";
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);
            }

            return response;
        }

        /// <summary>
        /// Carga a CD los contratos que han tenido cambios en su fecha de fin
        /// </summary>
        /// <param name="dsRenewals">????????????</param>
        /// <param name="mandErp">mandante de SAP ERP donde hacer la consulta</param>
        /// <returns></returns>
        private DataTable RenewContract(DataSet dsRenewals, int mandErp)
        {
            try
            {
                string info;
                string contractCol = "CONTRACT";
                string responseCol = "RESPONSE";

                DataTable ret = new DataTable();
                ret.Columns.Add(contractCol);
                ret.Columns.Add(responseCol);

                if (dsRenewals.Tables["ContractsNewDates"].Rows.Count > 0)
                {
                    DataTable responseEndDate = dsRenewals.Tables["ContractsNewDates"].Select("APPT_TYPE = 'CONTEND'").CopyToDataTable();
                    DataTable responseItems = dsRenewals.Tables["ContractsNewItems"]; //items de los contratos con cambios

                    if (responseEndDate.Rows.Count != 0)
                    {
                        foreach (DataRow row in responseEndDate.Rows)
                        {
                            ret.TableName = "OK";

                            Dictionary<string, string> itemsAndServiceGroups = new Dictionary<string, string>();
                            DataRow retRow = ret.NewRow();

                            string newEnd = row["VALID_TO"].ToString();
                            string con = row["OBJECT_ID"].ToString();

                            try
                            {
                                string status = cdi.GetContractStatus(con);

                                if (status == "APPR")
                                {
                                    XmlDocument outXml = new XmlDocument();

                                    #region 1 - Si hay items vencidos eliminarlos de la lista, y si hay "ADD" agregarlo a la lista
                                    itemsAndServiceGroups = GetItemsAddedToContractToday(responseItems, con);
                                    #endregion

                                    #region 2 - Buscar el MG de los service groups
                                    itemsAndServiceGroups = GetServiceGroupsMgs(itemsAndServiceGroups, mandErp);
                                    #endregion

                                    #region 3 - Leer el contrato actual de CD (contrato = con y status active)

                                    info = cdi.GetContractsXml(new List<string> { con });

                                    #region Crear XML de actualizar
                                    info = info.Replace("><", ">\r\n<");
                                    string[] xmlLines = info.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                    string[] addChange = { "<PLUSPAGREEMENT>", "<MXCUSTAGREEMENTSet>", "<PLUSPPRICESCHED>", "<PLUSPAPPLSERV>", "<PLUSPAPPLASSET>" };
                                    string[] removes = { "RENEWALDATE", "STATUS>", "PLUSPAPPLSERVID>", "GBMGUID>", "SANUM>", "OWNERID>", "OWNERTABLE>" };

                                    for (int k = 0; k < xmlLines.Length; k++)
                                    {
                                        for (int j = 0; j < addChange.Length; j++)
                                            if (xmlLines[k].Contains(addChange[j]))
                                                xmlLines[k] = xmlLines[k].Replace(addChange[j], addChange[j].Substring(0, addChange[j].Length - 1) + @" action=""AddChange"">");

                                        for (int j = 0; j < removes.Length; j++)
                                            if (xmlLines[k].Contains(removes[j]))
                                                xmlLines[k] = "";

                                        if (xmlLines[k].Contains("<QueryMXCUSTAGREEMENTResponse"))
                                            xmlLines[k] = @"<SyncMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""	xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" >";

                                        if (xmlLines[k].Contains("</QueryMXCUSTAGREEMENTResponse>"))
                                            xmlLines[k] = "</SyncMXCUSTAGREEMENT>";
                                    }
                                    string outXmlS = string.Join(Environment.NewLine, xmlLines);
                                    #endregion

                                    outXml.LoadXml(outXmlS);
                                    #endregion

                                    #region 4 - Construir el nuevo, sumando 1 a la version y colocando la nueva fecha
                                    XmlNodeList tempList = outXml.GetElementsByTagName("REVISIONNUM");
                                    int nextRev = int.Parse(tempList[tempList.Count - 1].InnerText) + 1;
                                    tempList[tempList.Count - 1].InnerText = nextRev.ToString();
                                    tempList = outXml.GetElementsByTagName("ENDDATE");
                                    tempList[tempList.Count - 1].InnerText = newEnd + "T00:00:00-05:00";
                                    #endregion

                                    #region 5 - Eliminar items vencidos
                                    List<string> sg = new List<string>();

                                    foreach (var item in itemsAndServiceGroups)
                                        sg.Add(item.Value);

                                    XmlNodeList deleteNodes = outXml.GetElementsByTagName("PLUSPAPPLSERV");
                                    for (int j = deleteNodes.Count - 1; j >= 0; j--)
                                        if (!sg.Contains(deleteNodes[j]["COMMODITY"].InnerText))//si el commodity(CD) no esta en la lista SG(SAP), pues se borra
                                            deleteNodes[j].ParentNode.RemoveChild(deleteNodes[j]);
                                    #endregion

                                    #region 6 - Insertar nuevos items
                                    List<string> cdServices = new List<string>();
                                    foreach (XmlNode item in outXml.GetElementsByTagName("COMMODITY"))
                                        cdServices.Add(item.InnerText);

                                    cdServices = cdServices.Distinct().ToList(); //servicios del contrato en CD
                                    foreach (string item in sg) //for del array de servicios de SAP
                                    {
                                        if (!cdServices.Contains(item)) // si el item NO esta en el array de CD agregarlo
                                        {
                                            XmlNodeList pluspPriceScheds = outXml.GetElementsByTagName("PLUSPPRICESCHED");
                                            foreach (XmlNode pluspPriceSched in pluspPriceScheds)
                                            {
                                                XmlElement pluspApplServ = outXml.CreateElement("PLUSPAPPLSERV", "http://www.ibm.com/maximo");
                                                pluspApplServ.SetAttribute("action", "AddChange");

                                                XmlElement commodity = outXml.CreateElement("COMMODITY", "http://www.ibm.com/maximo");
                                                XmlElement itemsetId = outXml.CreateElement("ITEMSETID", "http://www.ibm.com/maximo");
                                                XmlElement ownerTable = outXml.CreateElement("OWNERTABLE", "http://www.ibm.com/maximo");

                                                commodity.InnerText = item;
                                                itemsetId.InnerText = "ITEMSET1";
                                                ownerTable.InnerText = "PLUSPPRICESCHED";

                                                pluspApplServ.AppendChild(commodity);
                                                pluspApplServ.AppendChild(itemsetId);
                                                pluspApplServ.AppendChild(ownerTable);

                                                pluspPriceSched.AppendChild(pluspApplServ);
                                            }
                                        }
                                    }

                                    outXmlS = outXml.InnerXml;
                                    #endregion

                                    #region 7 - Post to  CD
                                    console.WriteLine("Renovando FECHAS contrato: " + con);
                                    string resXml = cdi.PostCD(root.UrlCd, "MXCUSTAGREEMENT", outXmlS); //cons con nuevas fechas

                                    if (resXml.Contains("PLUSPAGREEMENTID"))
                                    {
                                        retRow[contractCol] = con;
                                        retRow[responseCol] = "OK";
                                    }
                                    else
                                    {
                                        retRow[contractCol] = con;
                                        retRow[responseCol] = "Error renovando el contrato: " + resXml;
                                    }
                                    #endregion
                                }
                                else
                                {
                                    retRow[contractCol] = con;
                                    retRow[responseCol] = "Contrato status: " + status + ", no se renovó";
                                }
                            }
                            catch (Exception ex)
                            {
                                retRow[contractCol] = con;
                                retRow[responseCol] = "El contrato dio error: " + ex.Message;
                            }

                            ret.Rows.Add(retRow);
                        }
                    }
                    else
                        ret.TableName = "No se renovaron contratos hoy";

                    if (ret.Rows.Count == 0)
                        ret.TableName = "Ningún contrato en CD";
                }

                return ret;
            }
            catch (Exception ex)
            {
                DataTable errorDt = new DataTable();
                errorDt.Columns.Add("ERROR", typeof(string));
                errorDt.Rows.Add();
                errorDt.Rows[0][0] = ex.Message;
                return errorDt;
            }
        } //terminar el summary

        /// <summary>
        /// Crea releases cuando se encuentra contratos onHold
        /// </summary>
        /// <param name="onHoldCon"></param>
        /// <returns></returns>
        private DataTable CreateReleases(DataTable onHoldCon)
        {
            try
            {
                DataTable ret = new DataTable();

                ret.Columns.Add("ID Release");
                ret.Columns.Add("Status");
                ret.Columns.Add("ID Contract");
                ret.Columns.Add("Description");
                ret.Columns.Add("Net Value");
                ret.Columns.Add("Response");

                foreach (DataRow con in onHoldCon.Rows)
                {
                    List<string> releaseList = cdi.GetContractReleases(con["CONTRACT"].ToString());
                    DataRow tempRow = ret.NewRow();
                    if (releaseList.Count == 0) //No existen releases para el contrato
                    {
                        string targStartDate = con["ON_HOLD_DATE"].ToString(); //En CRM cambia de estado de Open a On Hold   ///HOY???
                        string targCompDate = DateTime.ParseExact(con["CON_START"].ToString(), "yyyyMMddHHmmss", null).AddDays(-1).ToString("yyyy-MM-dd");  //un día antes del con start
                        string netValue = con["NET_VALUE"].ToString();
                        string impact = "";

                        DateTime startDate = DateTime.Parse(targStartDate);
                        DateTime compDate = DateTime.Parse(targCompDate);

                        if (startDate > compDate)
                            targCompDate = DateTime.Now.ToString("yyyy-MM-dd");

                        decimal.TryParse(netValue, out decimal parsedNetValue);

                        if (parsedNetValue > 500000)      //mas de 500k
                            impact = "1. Critical";
                        else if (parsedNetValue > 200000) //200k a 500k
                            impact = "2. High";
                        else if (parsedNetValue > 50000)  //50k a 200k
                            impact = "3. Medium";
                        else                              //0 a 50k
                            impact = "4. Low";


                        CdReleaseData rel = new CdReleaseData
                        {
                            Description = con["CONTRACT_DESC"].ToString(),
                            PluspCustomer = con["CUSTOMER"].ToString(),
                            Contract = con["CONTRACT"].ToString(),
                            ExtRef = con["EXTERNAL_REFERENCE"].ToString(),
                            OwnerGroup = "GBMBPTPMO",
                            TargStartDate = targStartDate,
                            TargCompDate = targCompDate,
                            Employee = con["SALES_REP"].ToString(),
                            //Contact = string.Concat(con["CONTACT"].ToString().Take(20)), //AQUI NO DEBERIA IR EL ID 700 sino el correo del cliente, si viene de SAP bien, si no entonces sacarlo del mismo CD
                            ConRev = cdi.GetContractRevision(con["CONTRACT"].ToString()), //hay que sacarla del mismo CD
                            PmRelImpact = impact
                        };

                        #region Variables para los valores obligatorios en la aplicación (por si algún día los piden) y si lo pidieron 😅
                        //rel.Commodity = "4020505";
                        //rel.CommodityGroup = "CONSULTING";
                        //rel.Environment = "Calidad";
                        //rel.PmRelEmergency = "EMERGENCY FIX";
                        //rel.PmRelUrgency = "CRITICAL";
                        //rel.WoPriority = "1";
                        //rel.Classification = "8040";
                        #endregion

                        string[] releaseRes = cdi.CreateRelease(rel); //release, orden, error

                        tempRow[0] = releaseRes[0];
                        tempRow[1] = con["STATUS"].ToString();
                        tempRow[2] = con["CONTRACT"].ToString();
                        tempRow[3] = con["CONTRACT_DESC"].ToString();
                        tempRow[4] = netValue;
                        tempRow[5] = releaseRes[2];
                    }
                    else
                    {
                        tempRow[2] = con["CONTRACT"].ToString(); //contract
                        tempRow[5] = "Ya existe un release para este contrato"; //response
                    }
                    ret.Rows.Add(tempRow);
                }

                return ret;
            }
            catch (Exception ex)
            {
                DataTable errorDt = new DataTable();
                errorDt.Columns.Add("ERROR", typeof(string));
                errorDt.Rows.Add();
                errorDt.Rows[0][0] = ex.Message;
                return errorDt;
            }
        }

        /// <summary>
        /// Verifica si se debe ejecutar en una fecha especifica
        /// </summary>
        /// <returns>la fecha en formato "YYYY-MM-DD"</returns>
        private string ManualOn()
        {
            string date = "";
            try
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\CdOn.txt";
                string text = System.IO.File.ReadAllText(path);
                System.IO.File.WriteAllText(path, "");
                date = DateTime.Parse(text).ToString("yyyy-MM-dd");
            }
            catch (Exception) { }
            return date;
        }

        /// <summary>
        /// Devuelve el mandante adecuado dependiendo de donde se ejectue el robot
        /// </summary>
        /// <param name="mand">ERP, CRM o ERP</param>
        /// <returns>500,300,PRD... etc</returns>
        private string GetDefaultMand(string mand)
        {
            if (mand == "CD")
                return App.ConsoleApp.Start.enviroment;
            else
                return sap.checkDefault(mand, 0).ToString();
        }
    }
}
