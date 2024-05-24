using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using System;
using System.IO;

namespace DataBotV5.Automation.MASS.Backups
{
    /// <summary>
    /// Robot para respaldar todos los logs semanales de la carpeta BackupLogs por cada area, en el site BackupsLogs del 
    /// sharepoint del robot. Se realiza de forma semanal los días Domingos.
    /// </summary>
    class BackupLogs
    {
        #region Variables globales
        ConsoleFormat console = new ConsoleFormat();
        SharePoint sharepoint = new SharePoint();
        Rooting root = new Rooting();
        Credentials credentials = new Credentials();
        MailInteraction mail = new MailInteraction();
        Settings sett = new Settings();
        Log log = new Log();
        bool executeStats = false;

        string respFinal = "";


        #endregion
        public void Main()
        {
            console.WriteLine("Inicio de proceso de respaldo de logs alojados en la carpeta BackupLogs de esta semana a SharePoint.");
            console.WriteLine("Site de Sharepoint: BackupLogs, en la cuenta de office del Databot.");
            UploadLogsToSharepoint();

        }

        /// <summary>
        /// Este método sube a Sharepoint los logs en la carpeta BackupsLogs de cada area WEB, RPA, DM, MASS
        /// de todos los días de la semana, exceptuando el del Domingo que es cuando se ejecuta el robot.
        /// La cuenta de Sharepoint es con la cuenta office del Databot, en el site de BackupsLogs ubicado en el mismo. 
        /// </summary>
        private void UploadLogsToSharepoint()
        {

            string[,] userAreas = new string[7, 2]
            {
                { "WEB",  @"\\VM-DMAESTROS\Users\databot01\Desktop\databot\BackupLogs"},
                { "RPA",  @"\\VM-DMAESTROS\Users\databot02\Desktop\databot\BackupLogs"},
                { "RPA2", @"\\VM-DMAESTROS\Users\databot03\Desktop\databot\BackupLogs"},
                { "MASS", @"\\VM-DMAESTROS\Users\databot04\Desktop\databot\BackupLogs"},
                { "QAS", @"\\DATABOT05\Users\databotqa\Desktop\Databot\BackupLogs"},
                { "DM", @"\\DATABOT05\Users\databot05\Desktop\Databot\BackupLogs"},
                { "ICS", @"\\DATABOT05\Users\databot06\Desktop\Databot\BackupLogs"},
            };


            string linkGBMSharepoint = "https://gbmcorp.sharepoint.com/sites/DatabotDevelopers";

            string FileTxtToday = $"\\LogRobot{DateTime.Now.ToString("yyyyMMdd")}.txt";

            //Respaldar cada una de las userAreas.
            for (int i = 0; i < userAreas.GetLength(0); i++)
            {
                string area = userAreas[i, 0];
                string ruta = userAreas[i, 1];
                console.WriteLine($"Respaldando la área: {area}");
                string pathFileTxtToday = ruta + FileTxtToday;

                //Arreglo con las rutas de los logs de carpeta del area actual.
                string[] nameFiles = Directory.GetFiles(ruta);

                //Recorre los logs a respaldar del área actual.
                for (int e = 0; e < nameFiles.Length; e++)
                {
                    if (nameFiles[e] != pathFileTxtToday)//No respalde el log del día de hoy.
                    {
                        executeStats = true;
                        try
                        {
                            try
                            {
                                sharepoint.UploadFileToSharePointV2(linkGBMSharepoint, $"Documents/Respaldos/BackupLogs/{area}", nameFiles[e]);
                            }
                            catch (Exception exp)
                            {
                                sharepoint.UploadFileToSharePointV2(linkGBMSharepoint, $"Documents/Respaldos/BackupLogs/{area}", nameFiles[e]);
                            }

                            File.Delete(nameFiles[e]);
                            console.WriteLine($"Archivo respaldado: {nameFiles[e].Substring(nameFiles[0].Length - 20, 20)}.");

                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Respaldo de Logs", $"Archivo respaldado: { nameFiles[e].Substring(nameFiles[0].Length - 20, 20)}.", "");
                            respFinal = respFinal + "\\n" + $"Archivo respaldado: { nameFiles[e].Substring(nameFiles[0].Length - 20, 20)}.";


                        }
                        catch (Exception ex)
                        {
                            string message = $"No se pudo realizar el correcto backup a Sharepoint del siguiente archivo: lx{nameFiles[e]}";
                            sett.SendError(this.GetType(), "Error en enviar a Backup un log", message, ex);
                            //mail.SendHTMLMail(message + "  EXC: " + ex, new string[] { "epiedra@gbm.net" }, $"Error en enviar a Backup un log - " + nameFiles[e].Substring(nameFiles[0].Length - 20, 20), 2);
                        }
                    }
                }

            }


            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "appmanagement";

            if (executeStats == true)
            {
                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

            console.WriteLine("Proceso de respaldo de logs y eliminación en local finalizado exitosamente.");
        }

    }
}
