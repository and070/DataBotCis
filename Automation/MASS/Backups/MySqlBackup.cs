using System;
using MySql.Data.MySqlClient;
using System.Data;
using System.IO;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Data.Database;
using DataBotV5.App.Global;
using DataBotV5.App.ConsoleApp;

namespace DataBotV5.Automation.MASS.Backups

{
    /// <summary>
    /// Clase MASS Automation encargada de la creación de Backups SQL.
    /// </summary>
    class MySqlBackup
    {
        Credentials cred = new Credentials();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        SharePoint sharep = new SharePoint();
        ProcessAdmin padmin = new ProcessAdmin();
        Database db = new Database();
        public void Main()
        {
            #region variables privadas
            string sql_select = "SHOW SCHEMAS";
            DataTable mytable = new DataTable();
            ConsoleFormat console = new ConsoleFormat();
            string fecha = DateTime.Today.ToString();
            fecha = fecha.Replace("/", "_");
            fecha = fecha.Replace(" 00:00:00", "");
            bool error = false;
            string mensaje_error = "";
            string bd = "";
            string file_name = "";
            string respFinal = "";
            bool executeStats = false;

            #endregion


            try
            {
                MySqlConnection connexion = db.ConnSmartSimple("databot_db", Start.enviroment);

                connexion.Open();
                using (MySqlCommand execute = new MySqlCommand(sql_select, connexion))
                {
                    using (IDataReader dr = execute.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            MySqlConnection conn = db.ConnSmartSimple(dr[0].ToString(), Start.enviroment);

                            using (MySqlCommand cmd = new MySqlCommand())
                            {
                                cmd.CommandTimeout = 600;
                                using (MySql.Data.MySqlClient.MySqlBackup mb = new MySql.Data.MySqlClient.MySqlBackup(cmd))
                                {
                                    try
                                    {
                                        cmd.Connection = conn;
                                        conn.Open();
                                        file_name = dr[0].ToString() + " - " + fecha;
                                        string file = root.backup_root + "\\" + file_name + ".sql";

                                        if (File.Exists(file)) { File.Delete(file); }
                                        mb.ExportToFile(file);
                                        sharep.UploadFileToSharePointV2("https://gbmcorp.sharepoint.com/sites/DatabotDevelopers", $"Documents/Respaldos/MySqlBackup", file);
                                        executeStats = true;
                                        if (File.Exists(file)) { File.Delete(file); }
                                        conn.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        error = true;
                                        bd = bd + "<br>" + dr[0].ToString();
                                        mensaje_error = mensaje_error + "<br>" + ex.ToString();
                                    }

                                }
                            }


                            log.LogDeCambios("Creacion", root.BDProcess, "appmanagement@gbm.net", "Crear BackUp BD", file_name, mensaje_error);
                            respFinal = respFinal + "\\n" + "Crear BackUp BD: " + file_name;


                        }
                    }
                }

                connexion.Close();


            }
            catch (Exception ex)
            {
                error = true;
                mensaje_error = ex.ToString();
            }

            if (error == true)
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Nombre de la(s) base de dato(s): " + "<br>" + bd + "<br>" + "Error:" + "<br>" + mensaje_error, new string[] {"appmanagement@gbm.net"}, "BackUp de la BD - Databot", cc);
            }

            root.requestDetails = respFinal;
            root.requestDetails = "appmanagement";

            if (executeStats == true)
            {
                console.WriteLine("Creando estadísticas...");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }



        }

    }
}
