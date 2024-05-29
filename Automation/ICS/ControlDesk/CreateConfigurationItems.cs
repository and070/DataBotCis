using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.ControlDesk;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Automation.ICS.ControlDesk
{
    internal class CreateConfigurationItems
    {
        readonly ControlDeskInteraction cdi = new ControlDeskInteraction();
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly Credentials cred = new Credentials();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        const string mandCd = "DEV";

        public void Main()
        {
            //configurar el repositorio

            cred.SelectCdMand("DEV");


            //List<string> cisList = new List<string>();
            //cisList.Add("CRI400A");
            //cisList.Add("CRI400B_PRIMA");

            //CdCollectionData collection = new CdCollectionData();
            //collection.CollectionNum = "EVCMI001";
            //collection.Cis = cisList;

            //string XXX = cdi.AddCollectionCis(collection);

            CdCiSpecData ciSpec1 = new CdCiSpecData();
            ciSpec1.AssetAttrId = "IP";
            ciSpec1.CCiSumSpecValue = "172.19.251.250";

            CdCiSpecData ciSpec2 = new CdCiSpecData();
            ciSpec2.AssetAttrId = "HOSTNAME";
            ciSpec2.CCiSumSpecValue = "IBM_8960_P64_HO_SWA";

            List<CdCiSpecData> cdCiSpecDatas = new List<CdCiSpecData>();
            cdCiSpecDatas.Add(ciSpec1);
            cdCiSpecDatas.Add(ciSpec2);

            CdConfigurationItemData ci = new CdConfigurationItemData();
            ci.CiName = "IBM_8960_P64_HO_SW";
            ci.CiNum = "SSW40.PA.754754N";
            ci.PersonId = "jbobadilla@gbm.net";
            ci.Description = "IBM_8960_P64_HO_SWA";
            ci.Status = "PRODUCTION";
            ci.PmcCiImpact = "1";
            ci.CiLocation = "PA.CDAH.HW";
            ci.ClassStructureId = "1030";
            ci.ServiceGroup = "OUTSOURCIN";
            ci.Service = "30407";
            ci.PluspCustomer = "0010000531";
            ci.GbmAdministrator = "GBMOGS_COE_PLA_BACKUP/STORAGE";
            ci.CiSpecs = cdCiSpecDatas;
            string csadsadas = cdi.CreateCi(ci);



            //mail.GetAttachmentEmail("La carpeta de CIS", "Procesados", "Procesados de CIS");
            //if (root.ExcelFile != "")
            //{
            //    string filePath = root.FilesDownloadPath + "\\" + root.ExcelFile;
            //    DataTable excelDt = excel.GetExcelBook(filePath).Tables["Formulario Colecciones"];
            //    CreateCis(excelDt);
            //    using (Stats stats = new Stats()) { stats.CreateStat(); }
            //}
        }

        private void CreateCis(DataTable cisDt)
        {
            //crear el CI
            //Hola mundo//
            //Hello/
        }
    }
}
