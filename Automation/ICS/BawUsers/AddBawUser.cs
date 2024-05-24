using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Projects.BAW;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.BawUsers
{
    /// <summary>
    /// Clase ICS Automation encargada de agregar usuarios BAW.
    /// </summary>
    class AddBawUser
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ActiveDirectory ad = new ActiveDirectory();
        readonly BawInteraction baw = new BawInteraction();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();

        string respFinal = "";

        public void Main()
        {
            List<string> validMails = new List<string>();
            DataTable validMailsDt = crud.Select("SELECT `mail` FROM `authMails`", "add_baw_user_db");
            foreach (DataRow fila in validMailsDt.Rows)
                validMails.Add(fila["mail"].ToString().ToUpper());

            if (mail.GetAttachmentEmail("Solicitudes Usuarios BAW", "Procesados", "Procesados Usuarios BAW"))
            {
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessUsers(excelDt);

                console.WriteLine("Creando estadísticas...");
                root.requestDetails = respFinal;

                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }
        private void ProcessUsers(DataTable excel)
        {
            DataTable emailResponse = new DataTable();

            emailResponse.Columns.Add("Usuario");
            emailResponse.Columns.Add("Respuesta");

            cred.SelectBawMand(App.ConsoleApp.Start.enviroment); //DEV,QAS o PRD
            baw.SetBawApiToken();

            foreach (DataRow row in excel.Rows)
            {
                DataRow tempRow = emailResponse.NewRow();
                GetProcessData processData = GetProcessData(row["Proceso en el cual agregar usuario"].ToString().Trim());
                string newUser = row["Usuario que se desea agregar"].ToString().Trim();
                string refUser = row["Usuario con mismos permisos "].ToString().Trim();

                tempRow["Usuario"] = newUser;

                if (newUser != "" && refUser != "" && processData.Container != "")
                {
                    string res;
                    if (ad.ExistAD(newUser))
                    {
                        if (ad.ExistAD(refUser))
                        {
                            List<string> newGroups = baw.AddUserToGroup(newUser, refUser, processData);
                            res = String.Join(", ", newGroups.ToArray());
                            console.WriteLine(newUser + " --> " + res);
                        }
                        else
                        {
                            //el ref user no existe
                            res = "El Usuario " + refUser + " no existe";
                        }
                    }
                    else
                    {
                        //el new user no existe
                        res = "El Usuario " + newUser + " no existe";
                    }

                    tempRow["Respuesta"] = res;
                    log.LogDeCambios("", root.BDProcess, root.BDUserCreatedBy, "Agregar Usuarios BAW", newUser, res);
                    respFinal = respFinal + "\\n" + "Agregar Usuarios BAW " + " " + newUser + " " + res;

                }
                emailResponse.Rows.Add(tempRow);
            }

            string msg = "En la siguiente tabla se muestra los grupos de BAW a los cuales se agregaron los usuarios solicitados<br><br>";
            msg += val.ConvertDataTableToHTML(emailResponse);

            mail.SendHTMLMail(msg, new string[] { root.BDUserCreatedBy }, root.Subject, new string[] { "atrigueros@gbm.net", "smarin@gbm.net" });
        }
        private GetProcessData GetProcessData(string process)
        {
            GetProcessData ret = new GetProcessData();
            DataTable dt = crud.Select("SELECT `container`,`version` FROM `processesNames` WHERE excelName = '" + process + "'", "add_baw_user_db");
            ret.Container = dt.Rows[0]["container"].ToString();
            ret.Version = dt.Rows[0]["version"].ToString();

            return ret;
        }
    }
    public class GetProcessData
    {
        internal string Container { get; set; }
        internal string Version { get; set; }
    }
}
