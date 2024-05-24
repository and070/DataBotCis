using DataBotV5.Logical.Projects.ControlDesk;
using DataBotV5.Logical.Projects.TIRequest;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.TIRequest
{
    internal class InactiveIcsUserBPM
    {
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly TiFunctions tiReq = new TiFunctions();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();
        string respFinal = "";
        readonly string crudDb = "QAS";
        readonly int[] sapMands = new int[] { 300, 500, 460, 260, 420, 120, 410, 110, 100, 400, 4002 };
        readonly string[] cdMands = new string[] { "DEV", "QAS", "PRD" };

        public void Main()
        {
            tiReq.ProcessEmailRequest(crudDb);
            ProcessBpmRequest();
        }

        private void ProcessBpmRequest()
        {
            bool statsCreate = false;
            DataTable pendingRequests = crud.Select("SELECT * FROM `pending`", "ti_requests_db");
            DataTable resDt = new DataTable();
            resDt.Columns.Add("ID Usuario");
            resDt.Columns.Add("Sistema");
            resDt.Columns.Add("Resultado");

            if (pendingRequests.Rows.Count > 0)
            {
                foreach (DataRow request in pendingRequests.Rows)
                {
                    DataRow resRow;
                    string id = request["id"].ToString();
                    root.Email_Body = request["emailBody"].ToString();
                    root.BDUserCreatedBy = "bpm@mailgbm.com";
                    root.Subject = "Nueva Solicitud de TI Notificación RPA";

                    try
                    {
                        #region Nuevas solicitudes de BPM

                        string requestType = tiReq.GetRequestType(root.Email_Body); //si es baja
                        if (requestType == "BAJAICS")
                        {
                            string sapUserId = tiReq.GetValFromBPM("Número de Colaborador:", root.Email_Body);
                            string sapUserName = tiReq.GetSapUserName(sapUserId);
                            string sapUserEmail = tiReq.GetValFromBPM("Correo del usuario:", root.Email_Body);
                            //Portal
                            string lockMsgPortal = tiReq.DeleteUserPortal(sapUserName);
                            console.WriteLine("Desactivando en Portal: " + lockMsgPortal);

                            resRow = resDt.NewRow();
                            resRow[0] = sapUserName;
                            resRow[1] = "Portal";
                            resRow[2] = lockMsgPortal;
                            resDt.Rows.Add(resRow);
                            //log.LogDeCambios("Modificación", root.BDProcess, root.BDUserCreatedBy, "Baja " + sapUserName + " en Portal", lockMsgPortal, root.Subject + "**BAJA**");


                            //SAP
                            foreach (int sapMand in sapMands)
                            {
                                //desarollo 300?? solman?
                                string lockMsgSap;
                                try
                                {
                                    lockMsgSap = tiReq.DeleteUserSap(sapUserName, sapMand);
                                }
                                catch (Exception ex)
                                {
                                    lockMsgSap = ex.Message;
                                }
                                console.WriteLine("Desactivando en SAP " + sapMand + ": " + lockMsgSap);

                                resRow = resDt.NewRow();
                                resRow[0] = sapUserName;
                                resRow[1] = sapMand.ToString();
                                resRow[2] = lockMsgSap;
                                resDt.Rows.Add(resRow);
                                log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Baja " + sapUserName + " en SAP: " + sapMand, lockMsgSap, root.Subject + "**BAJA**");
                                respFinal = respFinal + "\\n" + "Baja " + sapUserName + " en SAP: " + sapMand;
                                statsCreate = true;
                            }

                            //CD
                            foreach (string cdMand in cdMands)
                            {
                                cred.SelectCdMand(cdMand);
                                string lockMsgCd;
                                try
                                {
                                    lockMsgCd = cdi.InactivateUser(sapUserName);
                                }
                                catch (Exception ex)
                                {
                                    lockMsgCd = ex.Message;
                                }
                                console.WriteLine("Desactivando en CD " + cdMand + ": " + lockMsgCd);
                                resRow = resDt.NewRow();
                                resRow[0] = sapUserName;
                                resRow[1] = cdMand;
                                resRow[2] = lockMsgCd;
                                resDt.Rows.Add(resRow);
                                log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Baja " + sapUserName + " en Control desk: " + cdMand, lockMsgCd, root.Subject + "**BAJA**");
                                respFinal = respFinal + "\\n" + "Baja " + sapUserName + " Control desk: " + cdMand;
                                statsCreate = true;
                            }

                            //universidad

                            //portal de DC
                            bool updateOnDc = terminateDcPortal(sapUserEmail.ToLower());
                            if (!updateOnDc)
                            {
                                mail.SendHTMLMail(root.Email_Body, new string[] { "joarojas@gbm.net" }, "Error al dar de baja en el portal de DataCenter", new string[] { "epiedra@gbm.net" }, null);
                            }

                            //borrar la solicitud
                            crud.NonQueryAndGetId("DELETE FROM `pending` WHERE `pending`.`id` = " + id, "ti_requests_db");

                            //enviar correos
                            mail.SendHTMLMail("Resultado de baja, usuario: " + sapUserName + " en los sistemas:<br><br>" + val.ConvertDataTableToHTML(resDt), new string[] { "internalcustomersrvs@gbm.net" }, root.Subject + "**BAJA**");
                        }

                        #endregion
                    }
                    catch (Exception ex)
                    {
                        mail.SendHTMLMail("Error al dar de baja al usuario<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                    }
                }

                if (statsCreate)
                {
                    root.requestDetails = respFinal;
                    root.BDUserCreatedBy = "internalcustomersrvs";
                    using (Stats stats = new Stats())
                        stats.CreateStat();
                }
            }
        }

        private bool terminateDcPortal(string sapUserEmail)
        {
            bool up = true;
            if (crud.Select($"SELECT email FROM user WHERE LOWER(email) = '{sapUserEmail}'", "gbmcloud").Rows.Count > 0)
            {
                up = crud.Update($"UPDATE user SET enabled = 0 WHERE LOWER(email) = '{sapUserEmail}'", "gbmcloud");
            }
            return up;
        }
    }
}