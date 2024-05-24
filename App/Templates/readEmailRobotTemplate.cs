using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.ActiveDirectory;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.App.Templates
{
    class readEmailRobotTemplate
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        CRUD crud = new CRUD();
        Log log = new Log();

        string system = "ERP";
        /// <summary>
        /// 
        /// </summary>
        public void Main()
        {
            //utilizar en caso de que el robot utilice SAP Logon GUI
            if (!sap.CheckLogin(system))
            {
                //Leer email y extraer adjunto
                if (mail.GetAttachmentEmail("Solicitudes xxxx", "Procesados", "Procesados xxxx"))
                {
                    //convertir el excel en un DataTable para ser procesada
                    DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    console.WriteLine("Processing...");
                    //utilizar en caso de que el robot utilice SAP Logon GUI
                    sap.BlockUser(system, 1);
                    ///Procesar------------------
                    Process(excelDt);
                    ///--------------------------
                    //utilizar en caso de que el robot utilice SAP Logon GUI
                    sap.BlockUser(system, 0);
                }
            }
        }
        /// <summary>
        ///
        /// </summary>
        /// <param name="ExcelFile">el excel que envía el usuario por email outlook</param>
        private void Process(DataTable ExcelFile)
        {
            #region private variables

            //columnas necesarias para cargar en SAP
            string[] readColumns = {
                "column1",
                "column2",
                "columnN"
            };
            //columnas del excel
            DataColumnCollection columns = ExcelFile.Columns;
            //contador para verificar el excel
            int contTrue = 0;
            //variable de respuesta en caso de que se necesite
            string response = "";
            //DataTable furuto excel de respuesta
            DataTable dtResponse = new DataTable();
            //variable nombre de la hoja de resultados
            string dtResponseSheetName = "Results";
            //variable nombre del libro de resultados + extension
            string dtResponseBookName = $"ResultsBook{DateTime.Now.ToString("yyyyMMdd")}" + root.ExcelFile;
            //ruta + nombre donde se guardará el excel de resultado
            string dtResponseRoute = root.FilesDownloadPath + "\\" + dtResponseBookName;
            //PLantilla en html para el envío de email
            string htmlEmail = Properties.Resources.emailtemplate1;
            //variable titulo del cuerpo del correo
            string htmlSubject = "Resultados";
            //variable contenido del correo: texto, cuadros, tablas, imagenes, etc
            string htmlContents = "";
            //se agregan los parametros anterior al html del cuerpo del email de respuesta
            htmlEmail = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", response).Replace("{contenido}", htmlContents);
            //variable remitente del email de respuesta
            string sender = root.BDUserCreatedBy;
            //variable copias del email de respuesta
            string[] cc = new string[] { "ejemplo1@gbm.net", "ejemplo2@gbm.net" };
            //variable ruta de adjunto
            string[] attachments = new string[] { dtResponseRoute };
            //variable cambio para log
            string logText = "";
            #endregion
            #region excel verification and Molded of the Results Excel
            console.WriteLine("Checking...");
            foreach (string columnName in readColumns)
            {
                //verifica si la columna esta en el excel
                if (columns.Contains(columnName))
                {
                    contTrue++;
                }
                //crea la columna en el excel de resultados
                dtResponse.Columns.Add(columnName);
            }
            //agrega la columna de resultado en el Excel
            dtResponse.Columns.Add("Response");
            //si es diferente a 4 significa que no encontro una de las columnas necesarias para cargar la reconocimiento
            if (contTrue != readColumns.Length)
            {
                response = "";
                htmlEmail = htmlEmail.Replace("{subject}", htmlSubject).Replace("{cuerpo}", response).Replace("{contenido}", htmlContents);
                mail.SendHTMLMail(htmlEmail, new string[] { sender }, "Error: " + root.Subject, cc, new string[] { root.FilesDownloadPath + "\\" + root.ExcelFile });
                return;
            }
            #endregion
            #region loop each excel row
            console.WriteLine("Foreach Excel row...");
            foreach (DataRow rRow in ExcelFile.Rows)
            {
                #region robot Process

                //log de cambios
                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, logText, "", "");
                #endregion
            }
            dtResponse.AcceptChanges();
            #endregion
            #region Create results Excel
            console.WriteLine("Save Excel...");
            excel.CreateExcel(dtResponse, dtResponseSheetName, dtResponseRoute);
            #endregion
            #region SendEmail
            console.WriteLine("Send Email...");
            mail.SendHTMLMail(htmlEmail, new string[] { sender }, root.Subject, cc, attachments);

            #endregion
        }
    
    }
}
