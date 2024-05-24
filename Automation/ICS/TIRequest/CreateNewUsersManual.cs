using DataBotV5.Logical.Projects.TIRequest;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;


namespace DataBotV5.Automation.ICS.TIRequest
{
    internal class CreateNewUsersManual
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly TiFunctions tiReq = new TiFunctions();
        readonly Rooting root = new Rooting();
        string respFinal = "";


        public void Main()
        {
            mail.GetAttachmentEmail("Solicitudes TI MANUAL", "Procesados", "Procesados Solicitudes TI");
            if (root.ExcelFile != null && root.ExcelFile != "")
            {
                console.WriteLine(" > > > " + "Nueva Solicitud de TI por CORREO de Usuario");

                string[] jsonSap = tiReq.ExcelToJson(root.FilesDownloadPath + "\\" + root.ExcelFile);
                string[] jsonCd = tiReq.ExcelToJson(root.FilesDownloadPath + "\\" + root.ExcelFile, "CD");
                string[] json105 = tiReq.ExcelToJson(root.FilesDownloadPath + "\\" + root.ExcelFile, "105");

                tiReq.ProcessAllSystems(jsonSap, jsonCd, json105);

                using (Stats stats = new Stats())
                    stats.CreateStat();
            }
        }
    }
}
