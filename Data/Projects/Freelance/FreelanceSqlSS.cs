using DataBotV5.App.Global;
using DataBotV5.Automation.WEB.Freelance;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Projects.Freelance;
using Newtonsoft.Json;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using System.Runtime.InteropServices;

namespace DataBotV5.Data.Projects.Freelance
{
    class FreelanceSqlSS
    {
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        Database.Database db = new Database.Database();
        Credentials.Credentials cred = new Credentials.Credentials();
        Rooting root = new Rooting();
        Settings settings = new Settings();
        string ssMandante = "PRD";
        public List<ArchivoBinario> ExtractFilesBinaries(string gestion)
        {
            List<ArchivoBinario> listado = new List<ArchivoBinario>();
            DataTable mytable = new DataTable();
            string sql = $"SELECT * FROM freelance_archivos WHERE ID_GESTION = '{gestion}'";
            //mytable = crud.Select("Databot", sql, "automation");

            //if (mytable != null)
            //{
            //    if (mytable.Rows.Count > 0)
            //    {
            //        for (int i = 0; i < mytable.Rows.Count; i++)
            //        {
            //            ArchivoBinario ar = new ArchivoBinario
            //            {
            //                NombreArchivo = $"{mytable.Rows[i][2]}",
            //                Contenido = (byte[])mytable.Rows[i][3]
            //            };
            //            listado.Add(ar);
            //        }
            //    }
            //}
            
            return listado;
        }
        /// <summary>
        /// Metodo para extraer las facturas ya aprobadas por el L2
        /// </summary>
        /// <returns></returns>
        public DataTable GetBillingsL2()
        {
            string sql = $@"SELECT * FROM billsFreelance WHERE status = '8' order by approvedTsL2 asc"; //procesando L2
            DataTable mytable = crud.Select( sql, "freelance_db");
            return mytable;
        }
        /// <summary>
        /// metodo para extraer las HES (no creada) y PO asociadas a una factura
        /// </summary>
        /// <param name="id">el id de la tabla billsFreelance</param>
        /// <returns></returns>
        public DataTable getBillPoInfo(string id)
        {
            string sql = $@"SELECT 
              billsPoInfo.id as uniqueId,
              billsPoInfo.id,
              billsPoInfo.createdAt,
              billsPoInfo.billId,
              billsPoInfo.handleUnitId,
              purcharseOrderAssignation.consultant AS consultant,
              accessFreelance.vendor AS consultantName,
              accessFreelance.user AS consultantUser,
              accessFreelance.email AS consultantEmail,
              handleUnitMain.poId,
                purcharseOrderAssignation.purchaseOrder,
              purcharseOrderAssignation.byHito,
                purcharseOrderAssignation.mountHito,
              purcharseOrderAssignation.item,
              purcharseOrderAssignation.companyCode,
              purcharseOrderAssignation.description AS descriptions, 
              purcharseOrderAssignation.responsible AS responsible,
              assignationProjects.project AS projects,
              helpArea.type AS helpAreas,
              billsArea.areaCode AS areaCodes,
              SUM(hourReportFreelance.hours) AS hours,
              hourReportFreelance.catsId,
              handleUnitMain.id as hesId,
              handleUnitMain.handleUnitId AS hesNumber

              FROM billsPoInfo

              INNER JOIN (handleUnitMain           
              INNER JOIN (hourReportFreelance  
              INNER JOIN purcharseOrderAssignation
              ON purcharseOrderAssignation.id = hourReportFreelance.poId)
              ON hourReportFreelance.handleUnitId = handleUnitMain.id)
              ON handleUnitMain.id = billsPoInfo.handleUnitId
  
              LEFT JOIN assignationProjects ON purcharseOrderAssignation.project = assignationProjects.id
              INNER JOIN helpArea ON purcharseOrderAssignation.areaHelp = helpArea.id
              INNER JOIN billsArea ON purcharseOrderAssignation.area = billsArea.id
              INNER JOIN accessFreelance on purcharseOrderAssignation.consultant = accessFreelance.id
  
              WHERE billsPoInfo.active = 1 AND billsPoInfo.billId = {id}
              GROUP BY billsPoInfo.id,
              purcharseOrderAssignation.consultant,
              purcharseOrderAssignation.purchaseOrder,
              purcharseOrderAssignation.item,
              purcharseOrderAssignation.companyCode,
              purcharseOrderAssignation.description, 
              purcharseOrderAssignation.responsible,
              assignationProjects.project,
              helpArea.type,
              billsArea.areaCode ";


            string sqlHito = $@"SELECT 
              billsPoInfo.id as uniqueId,
              billsPoInfo.id,
              billsPoInfo.createdAt,
              billsPoInfo.billId,
              billsPoInfo.handleUnitId,
              purcharseOrderAssignation.consultant AS consultant,
              accessFreelance.vendor AS consultantName,
              accessFreelance.user AS consultantUser,
              accessFreelance.email AS consultantEmail,
              handleUnitMain.poId,
                purcharseOrderAssignation.purchaseOrder,
              purcharseOrderAssignation.byHito,
                purcharseOrderAssignation.mountHito,
              purcharseOrderAssignation.item,
              purcharseOrderAssignation.companyCode,
              purcharseOrderAssignation.description AS descriptions, 
              purcharseOrderAssignation.responsible AS responsible,
              assignationProjects.project AS projects,
              helpArea.type AS helpAreas,
              billsArea.areaCode AS areaCodes,
              0 AS hours,
              '' as catsId,
              handleUnitMain.id as hesId,
              handleUnitMain.handleUnitId AS hesNumber

              FROM billsPoInfo

              INNER JOIN (handleUnitMain           
              INNER JOIN purcharseOrderAssignation
              ON purcharseOrderAssignation.id = handleUnitMain.poId)
              ON handleUnitMain.id = billsPoInfo.handleUnitId
  
              LEFT JOIN assignationProjects ON purcharseOrderAssignation.project = assignationProjects.id
              INNER JOIN helpArea ON purcharseOrderAssignation.areaHelp = helpArea.id
              INNER JOIN billsArea ON purcharseOrderAssignation.area = billsArea.id
              INNER JOIN accessFreelance on purcharseOrderAssignation.consultant = accessFreelance.id
  
              WHERE billsPoInfo.active = 1 AND billsPoInfo.billId = {id} AND purcharseOrderAssignation.byHito = 1
              GROUP BY billsPoInfo.id,
              purcharseOrderAssignation.consultant,
              purcharseOrderAssignation.purchaseOrder,
              purcharseOrderAssignation.item,
              purcharseOrderAssignation.companyCode,
              purcharseOrderAssignation.description, 
              purcharseOrderAssignation.responsible,
              assignationProjects.project,
              helpArea.type,
              billsArea.areaCode ";

                string sqlString = "select * from (" + sql + ") AS a UNION ALL select * from " + "(" + sqlHito + ") AS b ORDER BY uniqueId DESC";

            DataTable mytable = crud.Select( sqlString, "freelance_db");
            return mytable;
        }
        /// <summary>
        /// metodo para extrae los senders y copies de facturacion
        /// </summary>
        /// <param name="country"></param>
        /// <returns></returns>
        public billEmails getBillEmails(string country)
        {
            billEmails billEmails = new billEmails();
            string sql = $@"SELECT GROUP_CONCAT(MIS.digital_sign.email) as emails,
                            freelance_db.financialRol.rol as rol
                            FROM freelance_db.financialCountryEmails
                            inner join MIS.digital_sign on MIS.digital_sign.id = freelance_db.financialCountryEmails.accountant
                            inner join freelance_db.financialRol on freelance_db.financialRol.id = freelance_db.financialCountryEmails.rol
                            where freelance_db.financialCountryEmails.country = '{country}' and freelance_db.financialCountryEmails.rol = '1'
                            union ALL
                            SELECT GROUP_CONCAT(MIS.digital_sign.email) as emails,
                            freelance_db.financialRol.rol as rol
                            FROM freelance_db.financialCountryEmails
                            inner join MIS.digital_sign on MIS.digital_sign.id = freelance_db.financialCountryEmails.accountant
                            inner join freelance_db.financialRol on freelance_db.financialRol.id = freelance_db.financialCountryEmails.rol
                            where freelance_db.financialCountryEmails.country = '{country}' and freelance_db.financialCountryEmails.rol = '2'";

            DataTable mytable = crud.Select( sql, "freelance_db");
            if (mytable.Rows.Count > 0)
            {
                billEmails.senders = mytable.Rows[0]["emails"].ToString().Split(',');
                billEmails.copies = mytable.Rows[1]["emails"].ToString().Split(',');
            }
            return billEmails;
        }
        /// <summary>
        /// Método para descargar archivos a el FTP de SmartAndSimple
        /// </summary>
        /// <param name="filePathName"></param>
        /// <param name="enviroment"> "PRD" o "DEV"</param>
        /// <returns></returns>
        public bool downloadFile(string filePathName)
        {
            try
            {
                string fileName = Path.GetFileName(filePathName);
                string pathfile = filePathName;
                //subir al FTP de S&S
                bool subir_files = db.DownloadFileSftp(filePathName, root.FilesDownloadPath + "\\" + fileName);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="file"></param>
        /// <param name="po"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool updateErrorHes(string id, byte[] file, string po, string item, string status, string consultantUser, string enviroment)
        {
            try
            {
                if (File.Exists(root.FilesDownloadPath + "\\" + "ErrorHes_" + id + ".jpg"))
                {
                    File.Delete(root.FilesDownloadPath + "\\" + "ErrorHes_" + id + ".jpg");
                }
                File.WriteAllBytes(root.FilesDownloadPath + "\\" + "ErrorHes_" + id + ".jpg", file);
                bool uploadFileError = db.uploadSftp(root.FilesDownloadPath + "\\" + "ErrorHes_" + id + ".jpg", "/home/appmanager/projects/smartsimple/gbm-hub-api/src/assets/files/Freelance", $"{id}-{po}-{item}");
                crud.Insert($"INSERT INTO `uploadFiles`(`name`, `idRequest`, `motherTable`, `user`, `codification`, `type`, `path`, `active`, `createdBy`) VALUES ('{"ErrorHes_" + id + ".jpg"}', '{id}', '1', '{consultantUser}', '7bit', 'image/jpeg' , '/home/appmanager/projects/smartsimple/gbm-hub-api/src/assets/files/Freelance/{id}-{po}-{item}/ErrorHes_{id}.jpg', 1, '{consultantUser}');", "freelance_db");
                crud.Update($"UPDATE handleUnitMain SET status = {status} WHERE id = {id}", "freelance_db");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="file"></param>
        /// <param name="po"></param>
        /// <param name="item"></param>
        /// <param name="hes"></param>
        /// <returns></returns>
        public bool updateHes(string id, byte[] file, string po, string item, string hes, string enviroment)
        {
            try
            {
                string filePath = root.FilesDownloadPath + "\\" + "freelance_hoja-" + id + ".jpg";
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                File.WriteAllBytes(filePath, file);
                bool uploadFileError = db.uploadSftp(filePath, "/home/appamanger/projects/smartsimple/gbm-hub-api/src/assets/files/Freelance", $"{id}-{po}-{item}");
               
            }
            catch (Exception)
            {
                return false;
            }

            crud.Update($"UPDATE handleUnitMain SET status = 1, handleUnitId = '{hes}' WHERE id = {id}", "freelance_db");
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetxSheet()
        {
            string sql = "SELECT * FROM handleUnitMain WHERE status = '5' AND active = '1';"; //EN PROCESO

            sql = $@"SELECT 
              purcharseOrderAssignation.consultant AS consultant,
              accessFreelance.vendor AS consultantName,
              accessFreelance.user AS consultantUser,
              accessFreelance.email AS consultantEmail,
              handleUnitMain.createdAt,
              handleUnitMain.poId,
              purcharseOrderAssignation.purchaseOrder,
              purcharseOrderAssignation.item,
              purcharseOrderAssignation.companyCode,
              purcharseOrderAssignation.description AS descriptions, 
              purcharseOrderAssignation.responsible AS responsible,
              billsArea.areaCode AS areaCodes,
              SUM(hourReportFreelance.hours) AS hours,
              hourReportFreelance.catsId,
              handleUnitMain.id as hesId,
              handleUnitMain.handleUnitId AS hesNumber

              FROM handleUnitMain
       
              INNER JOIN (hourReportFreelance  
              INNER JOIN purcharseOrderAssignation
              ON purcharseOrderAssignation.id = hourReportFreelance.poId)
              ON hourReportFreelance.handleUnitId = handleUnitMain.id

              INNER JOIN billsArea ON purcharseOrderAssignation.area = billsArea.id
              INNER JOIN accessFreelance on purcharseOrderAssignation.consultant = accessFreelance.id
  
              WHERE handleUnitMain.active = 1 AND handleUnitMain.status = '5'

              GROUP BY handleUnitMain.id,
              purcharseOrderAssignation.consultant,
              purcharseOrderAssignation.purchaseOrder,
              purcharseOrderAssignation.item,
              purcharseOrderAssignation.companyCode,
              purcharseOrderAssignation.description, 
              purcharseOrderAssignation.responsible,
              billsArea.areaCode,
              hourReportFreelance.catsId;";
            DataTable mytable = crud.Select( sql, "freelance_db");
            return mytable;

        }
        /// <summary>
        /// Metodo para extraer las horas aprobadas
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        public DataTable GetxState(string state)
        {
            string sql = $@"SELECT hourReportFreelance.*,
  purcharseOrderAssignation.consultant,
  purcharseOrderAssignation.id as PoId,
  purcharseOrderAssignation.purchaseOrder as purchaseOrder,
  purcharseOrderAssignation.item,
  purcharseOrderAssignation.description,
  purcharseOrderAssignation.responsible,
  (select vendor from accessFreelance WHERE accessFreelance.user = purcharseOrderAssignation.responsible) as responsibleName,
  purcharseOrderAssignation.project,
  purcharseOrderAssignation.areaHelp,
  purcharseOrderAssignation.area,
  assignationProjects.project as projectName,
  helpArea.type as helpAreaName,
  billsArea.area as areaName,
  billsArea.areaCode, 
  DATE_FORMAT(purcharseOrderAssignation.createdAt, '%d/%m/%Y %T') as Podate,
  DATE_FORMAT(hourReportFreelance.createdAt, '%d/%m/%Y %T') as Cdate,
  hourReportStatus.status as statusName,
  accessFreelance.vendor as consultantName,
  cancelationReasons.text as rejectedText
  from hourReportFreelance 
  INNER JOIN hourReportStatus on hourReportFreelance.status = hourReportStatus.id
  LEFT JOIN cancelationReasons on hourReportFreelance.rejectedReason = cancelationReasons.id
  INNER JOIN purcharseOrderAssignation on hourReportFreelance.poId = purcharseOrderAssignation.id
  LEFT JOIN assignationProjects on purcharseOrderAssignation.project = assignationProjects.id
  INNER JOIN helpArea on purcharseOrderAssignation.areaHelp = helpArea.id
  INNER JOIN billsArea on purcharseOrderAssignation.area = billsArea.id
  INNER JOIN accessFreelance on purcharseOrderAssignation.consultant = accessFreelance.id
  WHERE hourReportFreelance.active = 1 AND hourReportFreelance.status = '{state}'
  GROUP BY hourReportFreelance.id ORDER BY hourReportFreelance.id ASC;";

            DataTable mytable = crud.Select( sql, "freelance_db");
            return mytable;
        }
        public DataTable GetxCats(string po, string item)
        {
            DataTable mytable = new DataTable();
            string sql = "SELECT CATS,HOJA FROM freelance_g WHERE PO = '" + po + "' AND ITEM = '" + item + "'";
            //mytable = crud.Select("Databot", sql, "automation");
            return mytable;

        }
    }
}
