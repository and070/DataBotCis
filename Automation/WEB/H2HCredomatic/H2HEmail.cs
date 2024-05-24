using System;
using WinSCP;
using System.IO;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.WEB.H2HCredomatic
{
    /// <summary>
    /// Clase RPA Automation encargada de leer un correo, despúes lo carga en SAP cada uno de los archivos .txt que venga en el email.
    /// </summary>
    class H2HEmail
    {
        MASS.H2HCredomatic.H2HCredomatic h2hc = new MASS.H2HCredomatic.H2HCredomatic();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        SharePoint sharep = new SharePoint();
        Rooting root = new Rooting();
        ProcessInteraction proc = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        Log log = new Log();
        Stats estadisticas = new Stats();
        SapVariants sap = new SapVariants();
        ProcessAdmin padmin = new ProcessAdmin();
        Database db2 = new Database();
        ValidateData val = new ValidateData();
        internal Rooting Root { get => root; set => root = value; }
        string saldo_feba = "";
        string[] mt940_files;
        string[] link;
        string[] adjunto;
        string[] adjunto_F = new string[1];
        string[] adjunto_M = new string[1];
        string[] adjunto_P = new string[1];
        string file_name = "", file_path = "";
        string mandante = "ERP";
        string account_id = "";
        string statement_number = "";
        string company_code;
        string bank = "";
        string mensaje_sap = "";
        string mensaje_sap2 = "";
        string mensaje_sap3 = "";
        string respuesta = "";
        string respuesta2 = "";
        string respuesta_M = "";
        string respuesta_F = "";
        string respuesta_P = "";
        bool validar_lineas = true;
        int cantidad_files = 0;
        int contador = 0;
        int contador_F = 0;
        int contador_M = 0;
        int contador_P = 0;
        DateTime file_date = DateTime.MinValue;
        DateTime file_date_before = DateTime.MinValue;
        string sap_date = "";
        string dia = "";
        string mes = "";
        string ano = "";
        public string fldrpath = "";
        public string fldrpathDest = "";
        DateTime datenow = DateTime.Now.Date;

        string respFinal = "";


        public void Main()
        {
            if (!sap.CheckLogin(mandante))
            {
                if (mail.GetAttachmentEmail("Solicitudes H2H", "Procesados", "Procesados H2H"))
                {
                    sap.BlockUser(mandante, 1);
                    fldrpath = Root.h2hDownload + "\\";
                    fldrpathDest = Root.h2hDownloadArchive + "\\";
                    //buscar archivos ver si no hay attach.
                    bool ist = false;
                    for (int i = 0; i < root.filesList.Length; i++)
                    {
                        string pathArchivo = Root.FilesDownloadPath + "\\" + root.filesList[i].ToString();
                        string nameArchivo = root.filesList[i].ToString();
                        string extArchivo = Path.GetExtension(root.filesList[i].ToString());
                        if (extArchivo.ToLower().Contains("txt"))
                        {
                            if (nameArchivo.Substring(0, 5) == "MT940")
                            {
                                string destFile = fldrpath + nameArchivo;
                                try
                                {
                                    if (File.Exists(destFile))
                                    {
                                        File.Delete(destFile);
                                    }
                                    File.Copy(pathArchivo, destFile);
                                    ist = true;
                                }
                                catch (Exception)
                                {
                                    ist = false;
                                }
                            }
                        }
                    }
                    string email_gen = Properties.Resources.emailtemplate1;
                    if (ist)
                    {
                        TransferOperationResult transferResult = null;
                        h2hc.H2HProcess(transferResult);
                                      
                        root.requestDetails = "Se realizó el proceso de archivo vía email de H2H.";

                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }

                    }
                    else
                    {

                        email_gen = email_gen.Replace("{subject}", "Notificación Carga de Extractos Bancarios - BAC");
                        email_gen = email_gen.Replace("{cuerpo}", "No se encontró ningún archivo MT940 en el correo, por favor enviarlo nuevamente con los adjuntos correctos.");
                        email_gen = email_gen.Replace("{contenido}", "");
                        mail.SendHTMLMail(email_gen, new string[] { root.BDUserCreatedBy }, root.Subject, null);
                    }



                    sap.BlockUser(mandante, 0);

            
                }

            }
        }


    }
    public class accounts
    {
        public string account { get; set; }
        public string owner { get; set; }
        public string active { get; set; }
    }
}
