using DataBotV5.Logical.Projects.ControlDesk;
using DataBotV5.Logical.Projects.TIRequest;
using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.TIRequest
{
    internal class CreateNewUsersBPM
    {
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ActiveDirectory ad = new ActiveDirectory();
        readonly TiFunctions tiReq = new TiFunctions();
        readonly Credentials cred = new Credentials();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        string respFinal = "";

        Log log = new Log();


        const string mandCd = "PRD";
        const string crudDb = "PRD";

        public void Main()
        {
            tiReq.ProcessEmailRequest(crudDb);
            ProcessBpmRequest();
        }

        /// <summary>
        /// Toma las solicitudes de la BD y las procesa
        /// </summary>
        private void ProcessBpmRequest()
        {
            bool statsCreate = false;
            DataTable pendingRequests = crud.Select("SELECT * FROM `pending`", "ti_requests_db");

            if (pendingRequests.Rows.Count > 0)
            {
                foreach (DataRow request in pendingRequests.Rows)
                {
                    string id = request["id"].ToString();
                    root.Email_Body = request["emailBody"].ToString();
                    root.BDUserCreatedBy = "bpm@mailgbm.com";
                    root.Subject = "Nueva Solicitud de TI Notificación RPA";

                    try
                    {
                        #region Nuevas solicitudes de BPM

                        cred.SelectCdMand(mandCd);

                        string requestType = tiReq.GetRequestType(root.Email_Body); //verificar si es un email válido
                        if (requestType == "NUEVO")
                        {
                            try
                            {
                                if (tiReq.IsValidPosition(tiReq.GetValFromBPM("Posición", root.Email_Body)))//la posición existe y tiene roles
                                {
                                    if (ad.ExistAD(tiReq.GetValFromBPM("Correo del usuario:", root.Email_Body)) && cdi.CheckUserExistence(tiReq.GetValFromBPM("Correo del usuario:", root.Email_Body))) //existe en AD y en CD
                                    {
                                        crud.NonQueryAndGetId("DELETE FROM `pending` WHERE pending.id = " + id, "ti_requests_db");
                                        console.WriteLine(" > > > " + "Procesar Roles por CORREO de BPM");

                                        string[] jsonArray = tiReq.BpmToJson(root.Email_Body);
                                        string[] jsonCd = tiReq.BpmToJson(root.Email_Body, "CD");
                                        string[] json105 = tiReq.BpmToJson(root.Email_Body, "105");
                                        tiReq.ProcessAllSystems(jsonArray, jsonCd, json105);

                                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear usuario BPM", $"Se creó el usuario BPM {root.Email_Body} con éxito.", root.Subject);
                                        respFinal = respFinal + "\\n" + $"Se creó el usuario BPM {root.Email_Body} con éxito.";
                                        statsCreate = true;
                                    }
                                }
                                else
                                {
                                    //si no encuentra nada en la tabla pues que de error
                                    mail.SendHTMLMail("La posición id: " + tiReq.GetValFromBPM("Posición", root.Email_Body) + " no se encontró en la Base de Datos, por favor agregarla, o procesar la solicitud manualmente<br><br><hr><br><br>" + root.Email_Body.Replace("\n", "<br>"),
                                        new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);

                                    crud.NonQueryAndGetId("DELETE FROM `pending` WHERE pending.id = " + id, "ti_requests_db");
                                }
                            }
                            catch (Exception ex)
                            {
                                mail.SendHTMLMail("Error alta solicitudes de TI<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                                crud.NonQueryAndGetId("DELETE FROM `pending` WHERE pending.id = " + id, "it_request_db");
                            }
                        }




                        #endregion
                    }
                    catch (Exception ex)
                    {
                        mail.SendHTMLMail("Error al crear Usuarios SAP<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                    }
                }

                if (statsCreate)
                {

                    root.requestDetails = respFinal;
                    root.BDUserCreatedBy = "internalcustomersrvs";

                    using (Stats stats = new Stats()) stats.CreateStat();
                }
            }
        }
    }
}
