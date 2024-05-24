using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace DataBotV5.Automation.QAS.Autopp
{
    class AutoppSQL_QAS
    {
        CRUD crud = new CRUD();
        Credentials cred = new Credentials();
        Rooting root = new Rooting();
        Database db = new Database();

        /// <summary>
        /// Sube archivos a el FTP de SmartAndSimple, según el ambiente inserta a la tabla uploadFiles de Autopp el registro del nombre del nuevo archivo,
        /// y finalmente sube el archivo al servidor de SmartAndSimple através del FTP.
        /// </summary>
        /// <param name="bidNumber"></param>
        /// <param name="filePathName"></param>
        /// <param name="enviroment"> "PRD" o "DEV"</param>
        /// <returns></returns>
        public bool InsertFileAutopp(string oppId, string filePathName, string enviroment)
        {
            bool uploadFile = false;
            try
            {
                string user = "";
                string fileName = Path.GetFileName(filePathName);
                if (enviroment == "QAS")
                {
                    user = cred.QA_SS_APP_SERVER_USER;
                }
                else if (enviroment == "PRD")
                {
                    user = cred.PRD_SS_APP_SERVER_USER;
                }

                string pathfile = $"/home/{user}/projects/SS-QA/gbm-hub-api/src/assets/files/Autopp/{oppId}/{fileName}";
                string mimeType = MimeMapping.GetMimeMapping(fileName);
                string sql = "INSERT INTO `UploadsFiles` (`id`, `oppId`, `name`, `user`, `codification`, `type`, `path`, `active`, `createdAt`, `createdBy`) " +
                    $"VALUES (NULL, '{oppId}', '{fileName}', 'Databot', '7bit', '{mimeType}', '{pathfile}', '1', CURRENT_TIMESTAMP, 'Databot');";


                crud.Insert(sql, "autopp_db", enviroment);

                //subir al FTP de S&S
                //PRODUCTIVO uploadFile = db.uploadSftp(ipAdress, user, password, filePathName, $"/home/{user}/projects/smartsimple/gbm-hub-api/src/assets/files/Autopp", oppId, port);
                uploadFile = db.uploadSftp(filePathName, $"/home/gbmadmin/projects/SS-QA/gbm-hub-api/src/assets/files/Autopp", oppId, enviroment);

                return uploadFile;
            }
            catch (Exception ex)
            {
                return uploadFile;
            }
        }


        /// <summary>
        /// Método para descargar archivos a el FTP de SmartAndSimple
        /// </summary>
        /// <param name="filePathName"></param>
        /// <param name="enviroment"> "PRD" o "DEV"</param>
        /// <returns></returns>
        public bool DownloadFile(string filePathName, string enviroment)
        {
            try
            {
                string user = "";
                string fileName = Path.GetFileName(filePathName);
                if (enviroment == "QAS")
                {
                    user = cred.QA_SS_APP_SERVER_USER;
                }
                else if (enviroment == "PRD")
                {
                    user = cred.PRD_SS_APP_SERVER_USER;
                }

                string pathfile = filePathName;

                //subir al FTP de S&S
                return db.DownloadFileSftp(filePathName, root.FilesDownloadPath + "\\" + fileName, enviroment);


            }
            catch (Exception ex)
            {
                return false;
            }
        }


    }
}

