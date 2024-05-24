using System;
using System.Text.RegularExpressions;
using DataBotV5.Data.Projects.MasterData;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.DM.Warranties
{
    /// <summary>
    /// Clase DM Automation encargada de finalizar garantías de datos maestros.
    /// </summary>
    class FinalizeWarrantySS
    {
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        Log log = new Log();

        string respFinal = "";



        public void Main()
        {
            string[] cc = { "smarin@gbm.net" };

            mail.GetAttachmentEmail("Solicitudes Garantias", "Procesados", "Procesados Garantias");
            try
            {
                if (root.BDUserCreatedBy != null)
                {
                    if (root.Email_Body.Contains("FINALIZAR_GARANTIA"))
                    {
                        Regex alphanum = new Regex(@"[^\p{L}0-9 ]");
                        string[] separator = new string[] { "Gestion_garantia:" };
                        string[] bodySplit = root.Email_Body.Split(separator, StringSplitOptions.None);
                        bodySplit[1] = bodySplit[1].Replace('\r', ' ');
                        bodySplit = bodySplit[1].Split('\n');
                        string idGestionDM = alphanum.Replace(bodySplit[0], "").Trim().ToUpper();
                        DM.ChangeStateDM(idGestionDM, "Finalizado Manualmente", "3"); //FINALIZADO
                        separator = new string[] { "Solicitante_garantia: " };
                        bodySplit = root.Email_Body.Split(separator, StringSplitOptions.None);
                        bodySplit[1] = bodySplit[1].Replace('\r', ' ');
                        bodySplit = bodySplit[1].Split('\n');
                        separator = new string[] { "<" };
                        bodySplit = bodySplit[0].Split(separator, StringSplitOptions.None);
                        string sender = bodySplit[0].Trim().ToUpper();
                        mail.SendHTMLMail("Se actualizaron las garantías", new string[] { sender }, "Formulario Garantías - Notificación de Finalización de Gestión - #" + idGestionDM);

                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Proveedor", "Se actualizaron las garantías Gestión - #" + idGestionDM, root.Subject);
                        respFinal = respFinal + "\\n" + "Se actualizaron las garantías Gestión - #" + idGestionDM;

                        root.requestDetails = respFinal;

                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                mail.SendHTMLMail("Error finalizando garantía<br>" + ex, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, cc);
            }
        }
    }
}

