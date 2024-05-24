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
    internal class InactiveADUserBPM
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ActiveDirectory ad = new ActiveDirectory();
        readonly TiFunctions tiReq = new TiFunctions();
        readonly Credentials cred = new Credentials();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();

        readonly string crudDb = "QAS";
        readonly string[] errorNotifEmails = new string[] { "smarin@gbm.net", "internalcustomersrvs@gbm.net", "MISInfrastructure@gbm.net" };
        readonly string[] okNotifEmails = new string[] { "smarin@gbm.net", "MISInfrastructure@gbm.net" };

        string respFinal = "";


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
                        string requestType = tiReq.GetRequestType(root.Email_Body); //verificar si es un email válido

                        if (requestType == "BAJA")
                        {
                            try
                            {
                                string sapUserId = tiReq.GetValFromBPM("Número de Colaborador:", root.Email_Body);
                                string adUser = tiReq.GetSapUserName(sapUserId);

                                if (!adUser.ToUpper().Contains("ERROR"))
                                {
                                    bool inactiveResponse = ad.InactiveUser(adUser, /*"adrobot"*/cred.userAdminActiveDirectory, /*"itg4X3TQtCU$mq$"*/ cred.passAdminActiveDirectory);

                                    if (inactiveResponse)
                                    {
                                        string msg = "Se desactivó correctamente el usuario: " + adUser + " en Active Directory";
                                        console.WriteLine(" > > > " + msg);
                                        mail.SendHTMLMail(msg, okNotifEmails, root.Subject + "**BAJA**");
                                        log.LogDeCambios("Creación", root.BDProcess, root.BDUserCreatedBy, "Baja " + adUser + " en Active Directory", root.Subject + "**BAJA**", "");
                                        respFinal = respFinal + "\\n" + "Baja " + adUser + " en Active Directory";

                                    }
                                    else
                                    {
                                        string msg = "ERROR al desactivar usuario: " + adUser + " en Active Directory";
                                        console.WriteLine(" > > > " + msg);
                                        mail.SendHTMLMail(msg, errorNotifEmails, root.Subject + "**BAJA**");
                                    }

                                    crud.NonQueryAndGetId("DELETE FROM `pending` WHERE pending.id = " + id, "ti_requests_db");
                                }
                                else
                                {
                                    console.WriteLine(" > > > " + "No se pudo obtener el username del colaborador: " + sapUserId + " > > > " + adUser);
                                    mail.SendHTMLMail("No se pudo obtener el username del colaborador: " + sapUserId + "<br>" + adUser, errorNotifEmails, root.Subject);
                                }
                            }
                            catch (Exception ex)
                            {
                                string msg = "Error al dar de baja Usuarios Active Directory<br>";
                                console.WriteLine(" > > > " + msg);
                                mail.SendHTMLMail(msg + ex.Message, errorNotifEmails, root.Subject);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string msg = "Error al inactivar usuario en Active Directory<br>" + ex.Message;
                        console.WriteLine(" > > > " + msg);
                        mail.SendHTMLMail(msg, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                    }
                }


                root.requestDetails = respFinal;
                root.BDUserCreatedBy = "internalcustomersrvs";

                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                    stats.CreateStat();
            }
        }
    }
}
