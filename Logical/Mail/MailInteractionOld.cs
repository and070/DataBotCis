//using Microsoft.Office.Interop.Outlook;
//using System;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.IO;
//using System.Text;
//using System.Runtime.InteropServices;
//using Microsoft.Exchange.WebServices.Data;
//using Outlook = Microsoft.Office.Interop.Outlook;
//using System.Net;
//using DataBotV5.Data.Credentials;
//using DataBotV5.Data.Process;
//using DataBotV5.Data.Projects.Autopp;
//using DataBotV5.Data.Root;
//using DataBotV5.Data.Stats;
//using DataBotV5.Logical.Encode;
//using DataBotV5.Automation.WEB.Freelance;
//using DataBotV5.App.Global;
//using System.Net.Mail;
//using System.Linq;
//using Microsoft.Graph;
//using Microsoft.Identity.Client;
//using Application = Microsoft.Office.Interop.Outlook.Application;

//namespace DataBotV5.Logical.Mail
//{
//    /// <summary>
//    /// Clase Logical encargada de gestionar todas las interacciones con correos electrónicos.
//    /// </summary>
//    class MailInteractionOld : IDisposable
//    {
//        string clientId = "Your_Client_ID";
//        string clientSecret = "Your_Client_Secret";
//        string tenantId = "Your_Tenant_ID";

//        Rooting root = new Rooting();
//        Credentials cred = new Credentials();
//        Stats esta = new Stats();
//        ConsoleFormat console = new ConsoleFormat();

//        public Items MailItems;
//        public Application App;
//        public NameSpace Mapeo;
//        public MAPIFolder Carpeteo;
//        public MAPIFolder Carpeteo_Body;
//        public Items correo;
//        public Recipients recipientes;
//        private bool disposedValue;


//        #region notiFreelance
//        /// <summary>Envía una notificación al portal de Freelance Analytics.</summary>
//        public void SendNotificationPortalFreelanceAnalytics(string email, string subject, string route)
//        {
//            Microsoft.Office.Interop.Outlook.MailItem mail;
//            Microsoft.Office.Interop.Outlook.Recipients mailRecipients;
//            Microsoft.Office.Interop.Outlook.Recipient mailrecipient;
//            App = new Microsoft.Office.Interop.Outlook.Application();
//            mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//            mail.Subject = subject;
//            string html = "";
//            using (WebClient client = new WebClient())
//            {
//                client.Encoding = UTF8Encoding.UTF8;
//                html = client.DownloadString("https://databot.ngrok.io/ext/assets/js/freelance/formatoReporte");
//            }
//            mail.HTMLBody = html;
//            mailRecipients = mail.Recipients;
//            mailrecipient = mailRecipients.Add(email);
//            mail.Attachments.Add(route);
//            Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//            mail.SendUsingAccount = desiredAccount;
//            mail.Send();
//        }

//        #endregion
//        /// <summary>
//        /// Metodo para conectarse a la aplicación de Outlook que este logeada en la máquina en un determinado folder
//        /// </summary>
//        /// <param name="email"></param>
//        /// <param name="folder"></param>
//        /// <param name="subfolder"></param>
//        /// <param name="subsubfolder"></param>
//        public void SetOutlookConnection(string email, string folder, [Optional] string subfolder, [Optional] string subsubfolder)
//        {
//            try
//            {
//                App = new Microsoft.Office.Interop.Outlook.Application();

//                Mapeo = App.GetNamespace("MAPI");
//                if (subfolder != null)
//                {
//                    if (subsubfolder != null)
//                    {
//                        Carpeteo = Mapeo.Folders[email.ToString()].Folders[folder.ToString()].Folders[subfolder].Folders[subsubfolder];
//                    }
//                    else
//                    {
//                        Carpeteo = Mapeo.Folders[email.ToString()].Folders[folder].Folders[subfolder];
//                    }
//                }
//                else
//                {
//                    // Carpeteo = Mapeo.Folders[email].Folders[folder];
//                    Carpeteo = Mapeo.Folders[email].Folders[folder];
//                }
//                // Carpeteo = Mapeo.Folders["databot01@gbm.net"].Folders[folder];
//                MailItems = Carpeteo.Items;
//                //correo = MailItems.Restrict("[Unread] = true");
//            }
//            catch (System.Exception ex)
//            {
//                try
//                {
//                    console.WriteLine(ex.ToString() + "" + ex.Message);
//                    console.WriteLine("Error encontrado, manejando la excepcion");
//                    console.WriteLine("Cerrando OutLook");
//                    App.Quit();
//                    System.Threading.Thread.Sleep(500);
//                    console.WriteLine("Iniciando nueva instancia de OutLook");
//                    System.Diagnostics.Process process = new System.Diagnostics.Process();
//                    process.StartInfo = new ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
//                    process.Start();
//                    System.Threading.Thread.Sleep(20000);
//                    console.WriteLine("OutLook reiniciado, probando leer el correo de nuevo");
//                    App = null;

//                    App = new Microsoft.Office.Interop.Outlook.Application();
//                    Mapeo = App.GetNamespace("MAPI");
//                    if (subfolder != null)
//                    {
//                        if (subsubfolder != null)
//                        {
//                            Carpeteo = Mapeo.Folders[email.ToString()].Folders[folder.ToString()].Folders[subfolder].Folders[subsubfolder];
//                        }
//                        else
//                        {
//                            Carpeteo = Mapeo.Folders[email.ToString()].Folders[folder].Folders[subfolder];
//                        }
//                    }
//                    else
//                    {
//                        // Carpeteo = Mapeo.Folders[email].Folders[folder];
//                        Carpeteo = Mapeo.Folders[email].Folders[folder];
//                    }
//                    // Carpeteo = Mapeo.Folders["databot01@gbm.net"].Folders[folder];
//                    MailItems = Carpeteo.Items;
//                    //correo = MailItems.Restrict("[Unread] = true");
//                }
//                catch (System.Exception)
//                {
//                    console.WriteLine("No se pudo leer el correo");
//                }
//            }

//        }

//        /// <summary>
//        /// Método diseñado para ahorrar 2 líneas de código (SetOutlookConnection, ProcessMailAttachment, para usarlo en el if), el cual establece la instancia con Outlook,
//        /// usando internamente el método SetOutlookConnection (el cual define las rutas de carpeta de los correos electrónicos pendientes), 
//        /// despúes descarga todos los adjuntos de los correos existentes sin leer en el folder y lo almacena en una carpeta del bot para su procesamiento, posteriormente los mueve a la carpeta procesados 
//        /// y por último los borra de la carpeta de los correos electrónicos pendientes.
//        /// </summary>
//        /// <returns>
//        /// Retorna una variable tipo bool para verificar si descargó algún adjunto de correo(verificando si root.NArchivo_Excel != null).
//        /// </returns>
//        /// <param name="email"></param>
//        /// <param name="folderRequestToProcess">Localiza la carpeta donde están los correos de solicitudes por procesar.</param>
//        /// <param name="folderProcesed">Localiza la carpeta donde se desea almacenar los correos procesados.</param>
//        /// <param name="subfolderProcesed"></param>
//        /// <param name="subsubfolderProcesed"></param>
//        public bool GetAttachmentEmail(string folderRequestToProcess, string folderProcesed = null, string subfolderProcesed = null, string subsubfolderProcesed = null)
//        {

//            Rooting root = new Rooting();
//            bool downloadAnyAttachment = false;
//            root.ExcelFile = "";
//            //Se establece la conexión a Outlook, a la vez extrae y almacena los correos pendientes en la variable local mailItems.
//            SetOutlookConnection(root.Direccion_email, folderRequestToProcess);

//            int copycount = 0;
//            int cfgcount = 0;
//            try
//            {
//                #region establece la variable "carpeteo" donde se va a mover
//                if (subfolderProcesed != null)
//                {
//                    if (subsubfolderProcesed != null)
//                        Carpeteo = Mapeo.Folders[root.Direccion_email.ToString()].Folders[folderProcesed.ToString()].Folders[subfolderProcesed].Folders[subsubfolderProcesed];
//                    else
//                        Carpeteo = Mapeo.Folders[root.Direccion_email.ToString()].Folders[folderProcesed].Folders[subfolderProcesed];
//                }
//                else if (folderProcesed != null)
//                    Carpeteo = Mapeo.Folders[root.Direccion_email].Folders[folderProcesed];
//                #endregion

//                correo = MailItems.Restrict("[Unread] = true");

//                foreach (MailItem mail in correo)
//                {
//                    try
//                    {
//                        if (mail.UnRead)
//                        {
//                            root.BDUserCreatedBy = mail.SenderEmailType;
//                            if (mail.SenderEmailType == "EX")
//                            {
//                                AddressEntry sender = mail.Sender;
//                                ExchangeUser user = sender.GetExchangeUser();
//                                root.BDUserCreatedBy = user.PrimarySmtpAddress;
//                                if (String.IsNullOrEmpty(root.BDUserCreatedBy))
//                                    root.BDUserCreatedBy = mail.SenderEmailAddress.ToString();
//                            }
//                            else
//                                root.BDUserCreatedBy = mail.SenderEmailAddress.ToString();


//                            root.Subject = mail.Subject.ToString();
//                            root.ReceivedTime = mail.ReceivedTime;
//                            root.Email = mail;
//                            recipientes = mail.Recipients;

//                            foreach (Outlook.Recipient recip in recipientes)
//                                if (recip.Type == (int)OlMailRecipientType.olCC)
//                                    copycount++;

//                            if (copycount != 0)
//                            {
//                                root.CopyCC = new string[copycount];
//                                int copycount2 = 0;

//                                foreach (Microsoft.Office.Interop.Outlook.Recipient recip in recipientes)
//                                {
//                                    if (recip.Type == (int)OlMailRecipientType.olCC)
//                                    {
//                                        root.CopyCC[copycount2] = recip.Address;
//                                        copycount2++;
//                                    }
//                                }
//                            }
//                            else
//                                root.CopyCC = null;

//                            //extrae el attachments
//                            root.Email_Body = mail.Body;
//                            for (int i = 1; i <= mail.Attachments.Count; i++)
//                            {
//                                string attachfileName;
//                                try
//                                {
//                                    attachfileName = mail.Attachments[i].FileName.ToString();
//                                }
//                                catch (System.Exception)
//                                {
//                                    attachfileName = "";
//                                }

//                                if (attachfileName != "")
//                                {
//                                    string fileExt = Path.GetExtension(attachfileName);
//                                    if (fileExt.ToLower().Substring(0, 4) == ".xls" || fileExt.ToLower() == ".pdf")
//                                    {
//                                        root.ExcelFile = attachfileName;
//                                        mail.Attachments[i].SaveAsFile(root.FilesDownloadPath + @"\" + attachfileName);
//                                    }
//                                    else if (fileExt.ToLower() == ".cfr")
//                                        cfgcount++;
//                                }

//                            }

//                            if (cfgcount != 0)
//                            {
//                                root.cfr_list = new string[cfgcount];
//                                int cfgcount2 = 0;

//                                for (int i = 1; i <= mail.Attachments.Count; i++)
//                                {
//                                    string attachfile;
//                                    attachfile = mail.Attachments[i].FileName.ToString();
//                                    string extArchivo = Path.GetExtension(attachfile);
//                                    if (extArchivo.ToLower() == ".cfr")
//                                    {
//                                        mail.Attachments[i].SaveAsFile(root.FilesDownloadPath + @"\" + attachfile);
//                                        root.cfr_list[cfgcount2] = root.FilesDownloadPath + @"\" + attachfile;
//                                        cfgcount2++;
//                                    }
//                                }
//                            }

//                            if (folderProcesed != null || subfolderProcesed != null || subsubfolderProcesed != null)
//                            {
//                                mail.UnRead = false;
//                                mail.Move(Carpeteo);

//                                try
//                                {
//                                    mail.Delete();
//                                    mail.Save();
//                                }
//                                catch (System.Exception) { }

//                                break;
//                            }
//                            else
//                                root.Email = mail;
//                        }
//                    }
//                    catch (System.Exception ex)
//                    {
//                        console.WriteLine(ex.ToString());
//                        break;
//                    }
//                }
//                try
//                {
//                    App = null;
//                    Mapeo = null;
//                    Carpeteo = null;
//                    MailItems = null;
//                    correo = null;
//                    Marshal.ReleaseComObject(App);
//                    Marshal.ReleaseComObject(Mapeo);
//                    Marshal.ReleaseComObject(MailItems);
//                    Marshal.ReleaseComObject(correo);
//                    GC.Collect();
//                    GC.WaitForPendingFinalizers();
//                    GC.Collect();
//                }
//                catch (System.Exception) { }
//            }
//            catch (System.Exception)
//            {
//                console.WriteLine("No se pudo leer el correo");
//            }

//            //Validación clave para indicar si se descargó algún adjunto de correo electrónico.
//            if (!string.IsNullOrWhiteSpace(root.ExcelFile))
//                downloadAnyAttachment = true;

//            return downloadAnyAttachment;
//        }

//        /// <summary>
//        /// Método diseñado para ahorrar dos líneas de código(SetOutlookConnection y GetBody usado en el if), para obtener el body del 
//        /// folder procesados.
//        /// </summary>
//        /// <returns>Retorna una variable tipo bool para verificar si descargó algún body del correo (root.Email_Body != null)</returns>
//        /// <param name="folderRequestToProcess"></param>
//        /// <param name="folderProcesed"></param>
//        /// <param name="subfolderProcesed"></param>
//        /// <param name="subsubfolderProcesed"></param>
//        public bool GetAttachmentEmail(string folderRequestToProcess, string folderProcesed, string subfolderProcesed, [Optional] string subsubfolderProcesed)
//        {
//            bool downloadAnyBody = false;
//            //Se establece la conexión a Outlook, a la vez extrae y almacena los correos pendientes en la variable local mailItems.
//            SetOutlookConnection(root.Direccion_email, folderRequestToProcess);
//            root.Email_Body = "";
//            try
//            {
//                #region carpeta procesados
//                if (subfolderProcesed != null)
//                {
//                    if (subsubfolderProcesed != null)
//                    {
//                        Carpeteo_Body = Mapeo.Folders[root.Direccion_email.ToString()].Folders[folderProcesed.ToString()].Folders[subfolderProcesed].Folders[subsubfolderProcesed];
//                    }
//                    else
//                    {
//                        Carpeteo_Body = Mapeo.Folders[root.Direccion_email.ToString()].Folders[folderProcesed].Folders[subfolderProcesed];
//                    }
//                }
//                else
//                {
//                    // Carpeteo = Mapeo.Folders[email].Folders[folder];
//                    Carpeteo_Body = Mapeo.Folders[root.Direccion_email].Folders[folderProcesed];
//                }
//                #endregion

//                int copycount = 0;

//                correo = MailItems.Restrict("[Unread] = true");
//                foreach (Microsoft.Office.Interop.Outlook.MailItem mail in correo)
//                {
//                    try
//                    {

//                        if (mail.UnRead)
//                        {
//                            //root.BDStartDate = DateTime.Now;
//                            root.BDUserCreatedBy = mail.SenderEmailAddress.ToString();
//                            root.Subject = mail.Subject.ToString();
//                            root.Email = mail;

//                            recipientes = mail.Recipients;

//                            foreach (Microsoft.Office.Interop.Outlook.Recipient recip in recipientes)
//                            {
//                                if (recip.Type == (int)OlMailRecipientType.olCC)
//                                { copycount++; }
//                            }

//                            if (copycount != 0)
//                            {
//                                root.CopyCC = new string[copycount];
//                                int copycount2 = 0;

//                                foreach (Microsoft.Office.Interop.Outlook.Recipient recip in recipientes)
//                                {
//                                    if (recip.Type == (int)OlMailRecipientType.olCC)
//                                    {
//                                        root.CopyCC[copycount2] = recip.Address;
//                                        copycount2++;
//                                    }
//                                }
//                            }
//                            root.Email_Body = mail.Body;
//                            root.ReceivedTime = mail.ReceivedTime; ;
//                            mail.UnRead = false;
//                            mail.Move(Carpeteo_Body);

//                            try
//                            {
//                                mail.Delete();
//                                mail.Save();
//                            }
//                            catch (System.Exception)
//                            { }
//                            break;
//                        }



//                    }
//                    catch (System.Exception)
//                    {
//                    }
//                }
//                try
//                {
//                    this.App = null;
//                    this.Mapeo = null;
//                    this.Carpeteo = null;
//                    this.MailItems = null;
//                    this.correo = null;
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Mapeo);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(MailItems);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(correo);
//                    GC.Collect();
//                    GC.WaitForPendingFinalizers();
//                    GC.Collect();
//                }
//                catch (System.Exception)
//                { }

//            }
//            catch (System.Exception)
//            {
//                console.WriteLine("No se pudo leer el correo");
//            }

//            //Validación clave para indicar si se descargó algún adjunto de correo electrónico.
//            if (!string.IsNullOrWhiteSpace(root.Email_Body))
//                downloadAnyBody = true;

//            return downloadAnyBody;

//        }

//        /// <summary>Función creada para ahorrar dos lineas de código (SetOutlookConnection y ProcessMailAttachmentAll),
//        /// el cual procesa todos los adjuntos en el correo, y también valida si descargó todos los adjuntos, la ventaja
//        /// es que se puede usar en el if de cada Main de robots.</summary>
//        /// <returns>Retorna una variable tipo entero para validar si descargó todos los adjuntos, comparando: root.filesList != null y root.filesList[0] != null.</returns>
//        public bool GetAttachmentEmail(string folderRequestToProcess, string folderProcesed, string subfolderProcesed, [Optional] string subsubfolderProcesed)
//        {

//            bool downloadAllAttachment = false;
//            //Se establece la conexión a Outlook, a la vez extrae y almacena los correos pendientes en la variable local mailItems.
//            SetOutlookConnection(new Rooting().Direccion_email, folderRequestToProcess);


//            Rooting root = new Rooting();
//            int copycount = 0;
//            int file_count = 0;
//            root.filesList = null;
//            try
//            {
//                #region establece la variable "carpeteo" donde se va a mover
//                if (subfolderProcesed != null)
//                {
//                    if (subsubfolderProcesed != null)
//                    {
//                        Carpeteo = Mapeo.Folders[root.Direccion_email.ToString()].Folders[folderProcesed.ToString()].Folders[subfolderProcesed].Folders[subsubfolderProcesed];
//                    }
//                    else
//                    {
//                        Carpeteo = Mapeo.Folders[root.Direccion_email.ToString()].Folders[folderProcesed].Folders[subfolderProcesed];
//                    }
//                }
//                else
//                {
//                    // Carpeteo = Mapeo.Folders[email].Folders[folder];
//                    Carpeteo = Mapeo.Folders[root.Direccion_email].Folders[folderProcesed];
//                }
//                #endregion

//                correo = MailItems.Restrict("[Unread] = true");

//                // DefinirConneccion(folder);


//                foreach (Microsoft.Office.Interop.Outlook.MailItem mail in correo)
//                {
//                    try
//                    {
//                        if (mail.UnRead && mail.Attachments.Count > 0)
//                        {
//                            //console.WriteLine(mail.Body);
//                            //console.WriteLine(mail.SenderEmailAddress);
//                            root.BDUserCreatedBy = mail.SenderEmailAddress.ToString();
//                            root.Subject = mail.Subject.ToString();
//                            root.Email = mail;
//                            recipientes = mail.Recipients;

//                            foreach (Microsoft.Office.Interop.Outlook.Recipient recip in recipientes)
//                            {
//                                if (recip.Type == (int)OlMailRecipientType.olCC)
//                                { copycount++; }
//                            }

//                            if (copycount != 0)
//                            {
//                                root.CopyCC = new string[copycount];
//                                int copycount2 = 0;

//                                foreach (Microsoft.Office.Interop.Outlook.Recipient recip in recipientes)
//                                {
//                                    if (recip.Type == (int)OlMailRecipientType.olCC)
//                                    {
//                                        root.CopyCC[copycount2] = recip.Address;
//                                        copycount2++;
//                                    }
//                                }
//                            }

//                            if (mail.Attachments.Count > 0)
//                            {
//                                root.filesList = new string[mail.Attachments.Count];
//                                int filecount2 = 0;

//                                for (int i = 1; i <= mail.Attachments.Count; i++)
//                                {
//                                    string attachfile;
//                                    attachfile = mail.Attachments[i].FileName.ToString();
//                                    string extArchivo = Path.GetExtension(attachfile);
//                                    int char_file = attachfile.Length - extArchivo.Length;
//                                    if (char_file > 80)
//                                    {
//                                        attachfile = attachfile.Substring(0, 80) + extArchivo;
//                                    }
//                                    mail.Attachments[i].SaveAsFile(root.FilesDownloadPath + @"\" + attachfile);
//                                    root.filesList[filecount2] = attachfile;
//                                    filecount2++;

//                                }
//                            }
//                            try
//                            {
//                                mail.UnRead = false;
//                                mail.Move(Carpeteo);

//                                mail.Delete();
//                                mail.Save();
//                            }
//                            catch (System.Exception)
//                            { }
//                            break;
//                        }
//                    }
//                    catch (System.Exception ex)
//                    {
//                        console.WriteLine(ex.ToString());
//                        break;
//                    }
//                }
//                try
//                {
//                    this.App = null;
//                    this.Mapeo = null;
//                    this.Carpeteo = null;
//                    this.MailItems = null;
//                    this.correo = null;
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Mapeo);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(MailItems);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(correo);
//                    GC.Collect();
//                    GC.WaitForPendingFinalizers();
//                    GC.Collect();
//                }
//                catch (System.Exception)
//                { }



//            }
//            catch (System.Exception)
//            {
//                console.WriteLine("No se pudo leer el correo");
//            }
//            //Validación clave para indicar si se descargó la lista de archivos.
//            if (root.filesList != null && root.filesList[0] != null)
//                downloadAllAttachment = true;

//            return downloadAllAttachment;

//        }

//        /// <summary>
//        /// Método para envío de correos electrónicos a un único usuario final.
//        /// </summary>
//        /// <param name="message"> Cuerpo de mensaje.</param>
//        /// <param name="sender"></param>
//        /// <param name="subject"></param>
//        /// <param name="type">0=nada, 1=éxito, 2=error.</param>
//        /// <param name="cc">CC</param>
//        /// <param name="attachment">Adjunto.</param>
//        /// <param name="responseType">0 = no agregó el parámetro: Re:, 1 para tipo respuesta Re:, 2 para mensaje nuevo.</param>
//        public void SendHTMLMail(string message, string sender, string subject, int type, [Optional] string[] cc, [Optional] string[] attachment, [Optional] int responseType)
//        {
//            string ccs = "";
//            string lineas = "";
//            bool formato = true;
//            Microsoft.Office.Interop.Outlook.MailItem mail;
//            Microsoft.Office.Interop.Outlook.Recipients mailRecipients;
//            Microsoft.Office.Interop.Outlook.Recipient mailrecipient;
//            App = new Microsoft.Office.Interop.Outlook.Application();
//            Rooting root = new Rooting();

//            string form_firma = ConvertSignature(root.Formato_Firma);
//            switch (type)
//            {
//                case 0:
//                    formato = false;
//                    break;
//                case 1:
//                    lineas = System.IO.File.ReadAllText(root.Formato_Listo);
//                    break;
//                case 2:
//                    lineas = System.IO.File.ReadAllText(root.Formato_Error);
//                    break;
//            }
//            try
//            {
//                mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

//                //MailItem mailItem = App.CreateItemFromTemplate("\\Content\\RLW.oft", OlItemType.olMailItem);

//                mail.Subject = (responseType == 2) ? subject : "Re: " + subject;

//                mail.HTMLBody = (formato == true) ? String.Format(lineas, message) + "<br>" + form_firma : message + "<br>" + form_firma;

//                mailRecipients = mail.Recipients;
//                mailrecipient = mailRecipients.Add(sender);
//                if (cc != null && cc[0] != null)
//                {
//                    if (cc.Length == 1)
//                    {
//                        ccs = cc[0].ToString();
//                    }
//                    else
//                    {
//                        for (int i = 0; i < cc.Length; i++)
//                        {
//                            if (i == 0)
//                            {
//                                ccs = cc[i].ToString();
//                            }
//                            else
//                            {
//                                ccs = ccs + ";" + cc[i].ToString();
//                            }
//                        }
//                    }
//                    mail.CC = ccs;
//                }

//                if (attachment != null && attachment[0] != null)
//                {
//                    for (int i = 0; i < attachment.Length; i++)
//                    {
//                        mail.Attachments.Add(attachment[i].ToString());
//                    }
//                }
//                Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//                mail.SendUsingAccount = desiredAccount;
//                SendHTMLMailAndCatchErrors(mail);
//            }
//            catch (System.Exception ex)
//            {
//                console.WriteLine(ex.ToString() + "" + ex.Message);
//                console.WriteLine("Error encontrado, manejando la excepcion");
//                console.WriteLine("Cerrando OutLook");
//                App.Quit();
//                System.Threading.Thread.Sleep(500);
//                console.WriteLine("Iniciando nueva instancia de OutLook");
//                System.Diagnostics.Process process = new System.Diagnostics.Process();
//                process.StartInfo = new ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
//                process.Start();
//                System.Threading.Thread.Sleep(10000);
//                console.WriteLine("OutLook reiniciado, probando enviar el correo de nuevo");
//                try
//                {
//                    mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//                    mail.Subject = "Re: " + subject;
//                    if (formato == true)
//                    {
//                        mail.HTMLBody = String.Format(lineas, message) + "<br>" + form_firma;
//                    }
//                    else
//                    {
//                        mail.HTMLBody = message + "<br>" + form_firma;
//                    }
//                    mailRecipients = mail.Recipients;
//                    mailrecipient = mailRecipients.Add(sender);
//                    if (cc != null && cc[0] != null)
//                    {
//                        if (cc.Length == 1)
//                        {
//                            ccs = cc[0].ToString();
//                        }
//                        else
//                        {
//                            for (int i = 0; i < cc.Length; i++)
//                            {
//                                if (i == 0)
//                                {
//                                    ccs = cc[i].ToString();
//                                }
//                                else
//                                {
//                                    ccs = ccs + ";" + cc[i].ToString();
//                                }
//                            }
//                        }
//                        mail.CC = ccs;
//                    }

//                    if (attachment != null && attachment[0] != null)
//                    {
//                        for (int i = 0; i < attachment.Length; i++)
//                        {
//                            mail.Attachments.Add(attachment[i].ToString());
//                        }
//                    }
//                    Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//                    mail.SendUsingAccount = desiredAccount;
//                    mail.Send();
//                }
//                catch (System.Exception)
//                {
//                    console.WriteLine("Error encontrado, fallo en el manejo de excepcion");
//                    console.WriteLine("No se pudo enviar el correo electronico");
//                    SendHTMLMailOutlook(subject, message, new string[] { sender }, cc, attachment);
//                }
//            }
//        }

//        /// <summary>
//        /// Método tipo sobrecarga para envío de correos electrónicos a multiples usuarios, además, eliminando el tipo "Re:" en el subject.
//        /// </summary>
//        public void SendHTMLMail(string message, string[] sender, string subject, int type, [Optional] string[] cc, [Optional] string[] attachment, [Optional] int responseType)
//        {
//            string ccs = "";
//            string lineas = "";
//            bool formato = true;
//            Microsoft.Office.Interop.Outlook.MailItem mail;
//            Microsoft.Office.Interop.Outlook.Recipients mailRecipients;
//            Microsoft.Office.Interop.Outlook.Recipient mailrecipient;
//            App = new Microsoft.Office.Interop.Outlook.Application();
//            Rooting root = new Rooting();

//            string form_firma = ConvertSignature(root.Formato_Firma);
//            switch (type)
//            {
//                case 0:
//                    formato = false;
//                    break;
//                case 1:
//                    lineas = System.IO.File.ReadAllText(root.Formato_Listo);
//                    break;
//                case 2:
//                    lineas = System.IO.File.ReadAllText(root.Formato_Error);
//                    break;
//            }
//            try
//            {
//                mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

//                //MailItem mailItem = App.CreateItemFromTemplate("\\Content\\RLW.oft", OlItemType.olMailItem);

//                mail.Subject = (responseType == 2) ? subject : subject;

//                mail.HTMLBody = (formato == true) ? String.Format(lineas, message) + "<br>" + form_firma : message + "<br>" + form_firma;

//                mailRecipients = mail.Recipients;
//                ////////////agregar varios correos///////////////
//                foreach (string correo in sender)              //
//                {                                              //
//                    mailrecipient = mailRecipients.Add(correo);//
//                }////////////////////////////////////////////////

//                if (cc != null && cc[0] != null)
//                {
//                    if (cc.Length == 1)
//                    {
//                        ccs = cc[0].ToString();
//                    }
//                    else
//                    {
//                        for (int i = 0; i < cc.Length; i++)
//                        {
//                            if (i == 0)
//                            {
//                                ccs = cc[i].ToString();
//                            }
//                            else
//                            {
//                                ccs = ccs + ";" + cc[i].ToString();
//                            }
//                        }
//                    }
//                    mail.CC = ccs;
//                }

//                if (attachment != null && attachment[0] != null)
//                {
//                    for (int i = 0; i < attachment.Length; i++)
//                    {
//                        mail.Attachments.Add(attachment[i].ToString());
//                    }
//                }
//                Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//                mail.SendUsingAccount = desiredAccount;
//                SendHTMLMailAndCatchErrors(mail);
//            }
//            catch (System.Exception ex)
//            {
//                console.WriteLine(ex.ToString() + "" + ex.Message);
//                console.WriteLine("Error encontrado, manejando la excepcion");
//                console.WriteLine("Cerrando OutLook");
//                App.Quit();
//                System.Threading.Thread.Sleep(500);
//                console.WriteLine("Iniciando nueva instancia de OutLook");
//                System.Diagnostics.Process process = new System.Diagnostics.Process();
//                process.StartInfo = new ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
//                process.Start();
//                System.Threading.Thread.Sleep(10000);
//                console.WriteLine("OutLook reiniciado, probando enviar el correo de nuevo");
//                try
//                {
//                    mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//                    mail.Subject = subject;
//                    if (formato == true)
//                    {
//                        mail.HTMLBody = String.Format(lineas, message) + "<br>" + form_firma;
//                    }
//                    else
//                    {
//                        mail.HTMLBody = message + "<br>" + form_firma;
//                    }
//                    mailRecipients = mail.Recipients;
//                    ////////////agregar varios correos///////////////
//                    foreach (string correo in sender)              //
//                    {                                              //
//                        mailrecipient = mailRecipients.Add(correo);//
//                    }////////////////////////////////////////////////
//                    if (cc != null && cc[0] != null)
//                    {
//                        if (cc.Length == 1)
//                        {
//                            ccs = cc[0].ToString();
//                        }
//                        else
//                        {
//                            for (int i = 0; i < cc.Length; i++)
//                            {
//                                if (i == 0)
//                                {
//                                    ccs = cc[i].ToString();
//                                }
//                                else
//                                {
//                                    ccs = ccs + ";" + cc[i].ToString();
//                                }
//                            }
//                        }
//                        mail.CC = ccs;
//                    }

//                    if (attachment != null && attachment[0] != null)
//                    {
//                        for (int i = 0; i < attachment.Length; i++)
//                        {
//                            mail.Attachments.Add(attachment[i].ToString());
//                        }
//                    }
//                    Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//                    mail.SendUsingAccount = desiredAccount;
//                    mail.Send();
//                }
//                catch (System.Exception)
//                {
//                    console.WriteLine("Error encontrado, fallo en el manejo de excepcion");
//                    console.WriteLine("No se pudo enviar el correo electronico");
//                    foreach (string correo in sender)
//                    {
//                        SendHTMLMailOutlook(subject, message, correo, cc, attachment);
//                    }
//                }
//            }
//        }

//        /// <summary>Enviar correo electrónico en formato HTML.</summary>
//        public void SendHTMLMail(string html, string sender, string subject, [Optional] string[] cc, [Optional] string[] attachment)
//        {

//            string ccs = "";
//            Microsoft.Office.Interop.Outlook.MailItem mail;
//            Microsoft.Office.Interop.Outlook.Recipients mailRecipients;
//            Microsoft.Office.Interop.Outlook.Recipient mailrecipient;
//            App = new Microsoft.Office.Interop.Outlook.Application();

//            try
//            {
//                mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//                mail.Subject = subject;

//                mail.HTMLBody = html;

//                mailRecipients = mail.Recipients;
//                mailrecipient = mailRecipients.Add(sender);

//                if (cc != null && cc[0] != null)
//                {
//                    if (cc.Length == 1)
//                    {
//                        ccs = cc[0].ToString();
//                    }
//                    else
//                    {
//                        for (int i = 0; i < cc.Length; i++)
//                        {
//                            if (i == 0)
//                            {
//                                ccs = cc[i].ToString();
//                            }
//                            else
//                            {
//                                ccs = ccs + ";" + cc[i].ToString();
//                            }
//                        }
//                    }
//                    mail.CC = ccs;
//                }

//                if (attachment != null && attachment[0] != null)
//                {
//                    for (int i = 0; i < attachment.Length; i++)
//                    {
//                        mail.Attachments.Add(attachment[i].ToString());
//                    }
//                }

//                Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//                mail.SendUsingAccount = desiredAccount;
//                SendHTMLMailAndCatchErrors(mail);
//            }
//            catch (System.Exception ex) //Se intenta manejar la excepción y enviar el correo nuevamente.
//            {
//                console.WriteLine(ex.ToString() + "" + ex.Message);
//                console.WriteLine("Error encontrado, manejando la excepcion");
//                console.WriteLine("Cerrando OutLook");
//                App.Quit();
//                System.Threading.Thread.Sleep(500);
//                console.WriteLine("Iniciando nueva instancia de OutLook");
//                System.Diagnostics.Process process = new System.Diagnostics.Process();
//                process.StartInfo = new ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE");
//                process.Start();
//                System.Threading.Thread.Sleep(10000);
//                console.WriteLine("OutLook reiniciado, probando enviar el correo de nuevo");
//                try
//                {
//                    string ccs2 = "";
//                    Microsoft.Office.Interop.Outlook.MailItem mail2;
//                    Microsoft.Office.Interop.Outlook.Recipients mailRecipients2;
//                    Microsoft.Office.Interop.Outlook.Recipient mailrecipient2;
//                    App = new Microsoft.Office.Interop.Outlook.Application();

//                    mail2 = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//                    mail2.Subject = subject;

//                    mail2.HTMLBody = html;

//                    mailRecipients2 = mail2.Recipients;
//                    mailrecipient2 = mailRecipients2.Add(sender);

//                    if (cc != null && cc[0] != null)
//                    {
//                        if (cc.Length == 1)
//                        {
//                            ccs = cc[0].ToString();
//                        }
//                        else
//                        {
//                            for (int i = 0; i < cc.Length; i++)
//                            {
//                                if (i == 0)
//                                {
//                                    ccs = cc[i].ToString();
//                                }
//                                else
//                                {
//                                    ccs = ccs + ";" + cc[i].ToString();
//                                }
//                            }
//                        }
//                        mail2.CC = ccs;
//                    }

//                    if (attachment != null && attachment[0] != null)
//                    {
//                        for (int i = 0; i < attachment.Length; i++)
//                        {
//                            mail2.Attachments.Add(attachment[i].ToString());
//                        }
//                    }
//                    mail2.Send();
//                }
//                catch (System.Exception e) //Notificación final de que no se pudo enviar el correo apesar de haber intentado manejar
//                                           //la excepción y enviarlo otra vez.
//                {
//                    console.WriteLine("Error encontrado, fallo en el manejo de excepcion");
//                    console.WriteLine("No se pudo enviar el correo electronico");

//                    string[] cc2 = new string[] { "dmeza@gbm.net", "epiedra@gbm.net" };
//                    string sender2 = "appmanagement@gbm.net";

//                    console.WriteLine("Tipo de Excepción: " + e.ToString());
//                    try { console.WriteLine("StackTrace: " + e.StackTrace.ToString()); } catch { };
//                    console.WriteLine("Sender: " + sender);
//                    console.WriteLine("Subject: " + subject);

//                    string stringCC = "";
//                    string msgEmail = "";
//                    msgEmail += "<br><br>Exception: <br>" + e.ToString() + "<br><br>";
//                    try { msgEmail = msgEmail + "StackTrace: <br>" + e.StackTrace.ToString(); } catch { }
//                    msgEmail += "<br><br>Datos que se iban a enviar vía correo electrónico y que dieron error:<br>";
//                    try
//                    {
//                        msgEmail += "Sender: " + sender + "</br>";
//                        msgEmail += "Subject: " + subject + "</br>";
//                        msgEmail += "</br>CC: ";

//                        for (int i = 0; i < cc.Length; i++)
//                        {
//                            msgEmail += cc[i].ToString() + ", ";
//                            stringCC += cc[i].ToString() + ", ";
//                        }
//                    }
//                    catch { }

//                    console.WriteLine("CC: " + stringCC);
//                    msgEmail += "</br></br>HTML: </br></br></br>" + html + "<br><br>Databot.";

//                    new MailInteraction().SendHTMLMail(msgEmail, new string[] { sender },2, "No he podido enviar el correo electrónico - Databot", 2, cc2);

//                }
//            }


//        }

//        /// <summary>Enviar correo electrónico en formato HTML.</summary>
//        public void SendHTMLMail(string html, string[] sender, string subject, [Optional] string[] cc, [Optional] string[] attachment)
//        {

//            string ccs = "";
//            Microsoft.Office.Interop.Outlook.MailItem mail;
//            Microsoft.Office.Interop.Outlook.Recipients mailRecipients;
//            Microsoft.Office.Interop.Outlook.Recipient mailrecipient;
//            App = new Microsoft.Office.Interop.Outlook.Application();

//            try
//            {
//                mail = App.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
//                mail.Subject = subject;

//                mail.HTMLBody = html;

//                mailRecipients = mail.Recipients;
//                ////////////agregar varios correos///////////////
//                foreach (string correo in sender)              //
//                {                                              //
//                    mailrecipient = mailRecipients.Add(correo);//
//                }////////////////////////////////////////////////

//                if (cc != null && cc[0] != null)
//                {
//                    if (cc.Length == 1)
//                    {
//                        ccs = cc[0].ToString();
//                    }
//                    else
//                    {
//                        for (int i = 0; i < cc.Length; i++)
//                        {
//                            if (i == 0)
//                            {
//                                ccs = cc[i].ToString();
//                            }
//                            else
//                            {
//                                ccs = ccs + ";" + cc[i].ToString();
//                            }
//                        }
//                    }
//                    mail.CC = ccs;
//                }

//                if (attachment != null && attachment[0] != null)
//                {
//                    for (int i = 0; i < attachment.Length; i++)
//                    {
//                        mail.Attachments.Add(attachment[i].ToString());
//                    }
//                }
//                Outlook.Account desiredAccount = App.Session.Accounts[root.Direccion_email];
//                mail.SendUsingAccount = desiredAccount;
//                SendHTMLMailAndCatchErrors(mail);
//            }
//            catch (System.Exception ex)
//            {
//                console.WriteLine(" Error enviando notificacion al usuario: " + sender);
//            }
//        }

//        /// <summary>Reenvíar correo electrónico.</summary>
//        public void ForwardEmail(string sender, [Optional] string[] cc, [Optional] string[] attachment)
//        {
//            try
//            {

//                string ccs = "";
//                var newItem = root.Email.Forward(); //ES UNA VARIABLE GLOBAL QUE GUARDA EL MAIL OBJECT para utilizarlo aca
//                newItem.Recipients.Add(sender);
//                if (cc != null && cc[0] != null)
//                {
//                    if (cc.Length == 1)
//                    {
//                        ccs = cc[0].ToString();
//                    }
//                    else
//                    {
//                        for (int i = 0; i < cc.Length; i++)
//                        {
//                            if (i == 0)
//                            {
//                                ccs = cc[i].ToString();
//                            }
//                            else
//                            {
//                                ccs = ccs + ";" + cc[i].ToString();
//                            }
//                        }
//                    }
//                    newItem.CC = ccs;
//                }

//                if (attachment != null && attachment[0] != null)
//                {
//                    for (int i = 0; i < attachment.Length; i++)
//                    {
//                        newItem.Attachments.Add(attachment[i].ToString());
//                    }
//                }
//                newItem.Send();

//            }
//            catch (System.Exception ex)
//            {
//                console.WriteLine(" Error enviando notificacion al usuario: " + sender);
//                SendHTMLMail("Error al reenviar el correo electronico", new string[] {"appmanagement@gbm.net"}, root.Subject, new string[] { "dmeza@gbm.net" }, null);
//            }
//        }

//        /// <summary>Método para convertir firma.</summary>
//        public string ConvertSignature(string sFile)
//        {
//            string ConvierteFirma = "";

//            ConvierteFirma = System.IO.File.ReadAllText(sFile, Encoding.UTF8);


//            return ConvierteFirma;

//        }

//        /// <summary>Enviar correo electrónico atraves de outllook.</summary>
//        public void SendHTMLMailOutlook(string subject, string msj, string sender, [Optional] string[] cc, [Optional] string[] attachment)
//        {

//            using (SmtpClient client = new SmtpClient()
//            {
//                Host = "smtp.office365.com",
//                Port = 587, //22
//                UseDefaultCredentials = false, // This require to be before setting Credentials property
//                DeliveryMethod = SmtpDeliveryMethod.Network,
//                Credentials = new NetworkCredential(root.Direccion_email, cred.passOutlook), // you must give a full email address for authentication 
//                TargetName = "STARTTLS/smtp.office365.com", // Set to avoid MustIssueStartTlsFirst exception
//                EnableSsl = true // Set to avoid secure connection exception
//            })
//            {

//                MailMessage message = new MailMessage()
//                {
//                    From = new MailAddress(root.Direccion_email), // sender must be a full email address
//                    Subject = subject,
//                    IsBodyHtml = true,
//                    Body = msj,
//                    BodyEncoding = System.Text.Encoding.UTF8,
//                    SubjectEncoding = System.Text.Encoding.UTF8,

//                };
//                message.To.Add(sender);
//                if (cc != null && cc[0] != null)
//                {
//                    foreach (string ccMail in cc)
//                    {
//                        message.CC.Add(ccMail);
//                    }
//                }



//                client.Send(message);

//            }

//        }

//        /// <summary>
//        /// Método para capturar posibles errores con los correos(objeto mail) y notificar con el correo adjunto
//        /// Error 1: correos invalidos en los Destinatarios
//        /// </summary>
//        /// <param name="mail">MailItem con la info del correo</param>
//        private void SendHTMLMailAndCatchErrors(MailItem mail)
//        {
//            try { mail.Send(); }
//            catch (COMException ex)
//            {
//                try
//                {
//                    MailItem mail2 = App.CreateItem(OlItemType.olMailItem);
//                    string failedMail = root.FilesDownloadPath + "\\" + "failedMail.msg";
//                    mail.SaveAs(failedMail);
//                    mail2.Attachments.Add(failedMail);

//                    if (root.BDArea == "DM" || root.BDArea == "ICS")
//                    {
//                        mail2.Recipients.Add("dmeza@gbm.net");
//                        mail2.Recipients.Add("smarin@gbm.net");
//                        mail2.Recipients.Add("internalcustomersrvs@gbm.net");
//                    }
//                    else
//                    {
//                        mail2.Recipients.Add("dmeza@gbm.net");
//                        mail2.Recipients.Add("epiedra@gbm.net");
//                        mail2.Recipients.Add("appmanagement@gbm.net");
//                    }

//                    //Error 1
//                    if (ex.Message.Contains("does not recognize one or more names"))
//                    {
//                        mail2.Subject = "No se pudo enviar el siguiente mail por error los destinatarios";
//                        mail2.Send();
//                    }
//                    //Cualquier otro error
//                    else
//                        mail.Send();
//                }
//                catch (System.Exception) { mail.Send(); }
//            }
//        }


//        #region Metodos de proyectos en especifico
//        /// <summary>Método para obtener solicitudes aprobadas.</summary>
//        public Dictionary<string, string> GetApprovalRequests(string process)
//        {
//            Application app = new Application();
//            NameSpace mapi = app.GetNamespace("MAPI");
//            MAPIFolder requestFolder = mapi.Folders[root.Direccion_email].Folders["Solicitudes Aprobaciones de Power automate"];
//            MAPIFolder processedFolder = mapi.Folders[root.Direccion_email].Folders["Procesados"].Folders["Procesados Aprobaciones de Power automate"];
//            Items mailItems = requestFolder.Items;
//            Items unreadMails = mailItems.Restrict("[Unread] = true");

//            Dictionary<string, string> infoRet = new Dictionary<string, string>();

//            foreach (MailItem mail in unreadMails)
//            {
//                if (mail.Subject.Replace("appr_request_", "") == process)
//                {
//                    infoRet.Add("ResponseJson", mail.Body);

//                    #region Tomar los adjuntos
//                    if (mail.Attachments.Count > 0)
//                    {
//                        foreach (Outlook.Attachment attachment in mail.Attachments)
//                        {
//                            string attachFile = attachment.FileName.ToString();
//                            string fileExtension = Path.GetExtension(attachFile);

//                            int charFile = attachFile.Length - fileExtension.Length;
//                            if (charFile > 80)
//                                attachFile = attachFile.Substring(0, 80) + fileExtension;

//                            string fullPath = root.FilesDownloadPath + @"\" + attachFile;
//                            attachment.SaveAsFile(fullPath);

//                            if (fileExtension == ".json")
//                                infoRet.Add("OriginalJson", String.Join("", System.IO.File.ReadAllLines(fullPath)));
//                            else
//                                infoRet.Add("AttachmentPath", fullPath);
//                        }
//                    }
//                    #endregion

//                    #region Mover el correo a procesados y terminar
//                    try
//                    {
//                        mail.UnRead = false;
//                        mail.Move(processedFolder);
//                        mail.Delete();
//                        mail.Save();
//                    }
//                    catch (System.Exception) { }
//                    break;
//                    #endregion
//                }
//            }
//            return infoRet;
//        }

//        /// <summary>Método para obtener el reporte Cognos de Tarifas mas reciente </summary>
//        public string GetLastPriceReportMail()
//        {
//            string fullPath = "";
//            Application app = new Application();
//            NameSpace mapi = app.GetNamespace("MAPI");
//            MAPIFolder requestFolder = mapi.Folders[root.Direccion_email].Folders["Reporte Tarifas"];
//            MailItem lastMail = requestFolder.Items.GetLast();

//            #region Tomar los adjuntos
//            if (lastMail.Attachments.Count > 0)
//            {
//                foreach (Outlook.Attachment attachment in lastMail.Attachments)
//                {
//                    string attachFile = attachment.FileName.ToString();

//                    if (attachFile.Contains("Tarifas HCM Distribucion"))
//                    {
//                        string fileExtension = Path.GetExtension(attachFile);

//                        int charFile = attachFile.Length - fileExtension.Length;
//                        if (charFile > 80)
//                            attachFile = attachFile.Substring(0, 80) + fileExtension;

//                        fullPath = root.FilesDownloadPath + @"\" + attachFile;
//                        attachment.SaveAsFile(fullPath);
//                    }
//                }
//            }
//            #endregion

//            return fullPath;
//        }

//        /// <summary>
//        /// Metodo para extraer en una lista de strings los usuarios que han enviado solicitudes en una carpeta en especifico
//        /// </summary>
//        /// <param name="folderRequestToProcess"></param>
//        /// <returns></returns>
//        public List<string> getSenderFolder(string folderRequestToProcess)
//        {
//            List<string> list = new List<string>();
//            bool downloadAllAttachment = false;
//            //Se establece la conexión a Outlook, a la vez extrae y almacena los correos pendientes en la variable local mailItems.
//            SetOutlookConnection(new Rooting().Direccion_email, "Procesados", folderRequestToProcess);


//            Rooting root = new Rooting();
//            int copycount = 0;
//            int file_count = 0;
//            root.filesList = null;
//            try
//            {

//                correo = MailItems;

//                // DefinirConneccion(folder);


//                foreach (Microsoft.Office.Interop.Outlook.MailItem mail in correo)
//                {
//                    try
//                    {

//                        //console.WriteLine(mail.Body);
//                        //console.WriteLine(mail.SenderEmailAddress);
//                        //list.Add(mail.Sender.Name.ToString());
//                        list.Add(mail.Sender.GetExchangeUser().PrimarySmtpAddress.ToString());


//                    }
//                    catch (System.Exception ex)
//                    {
//                        console.WriteLine(ex.ToString());
//                        break;
//                    }
//                }


//            }
//            catch (System.Exception)
//            {
//                console.WriteLine("No se pudo leer el correo");
//            }

//            return list;

//        }

//        #endregion

//        #region Metodos locales
//        /// <summary>Método para crear una regla.</summary>
//        //public bool CreateRule(string ruleName, [Optional] string[] subjectllave, [Optional] string[] sender, [Optional] string foldermove)
//        //{
//        //    console.WriteLine("procesando...");
//        //    //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
//        //    //Credenciales cred = new Credenciales();
//        //    //service.Credentials = new WebCredentials("databot@gbm.net", "Tomates20$");
//        //    //// service.TraceEnabled = true;
//        //    ////  service.TraceFlags = TraceFlags.All;
//        //    //service.AutodiscoverUrl("databot@gbm.net", RedirectionUrlValidationCallback);
//        //    //RuleCollection ruleCollection = service.GetInboxRules("databot@gbm.net");
//        //    //console.WriteLine("Collection count: " + ruleCollection.Count);

//        //    Microsoft.Office.Interop.Outlook.MailItem mail;
//        //    Microsoft.Office.Interop.Outlook.Recipients mailRecipients;
//        //    Microsoft.Office.Interop.Outlook.Recipient mailrecipient;
//        //    App = new Microsoft.Office.Interop.Outlook.Application();
//        //    Mapeo = App.GetNamespace("MAPI");
//        //    Outlook.Rules rules = null;
//        //    //Outlook.MAPIFolder OutlookInbox = null;
//        //    //Outlook.Application OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;


//        //    try
//        //    {
//        //        rules = App.Session.DefaultStore.GetRules(); //Gets list of outlook rules
//        //    }
//        //    catch (System.Exception ex)
//        //    {
//        //        console.WriteLine(ex.Message);
//        //        console.WriteLine("Could not obtain rules collection.");
//        //        return false;
//        //    }


//        //    Outlook.Rule rule = rules.Create(ruleName, Outlook.OlRuleType.olRuleReceive);  //Creates new rule in collection
//        //    rule.Name = ruleName;

//        //    //From condition
//        //    if (sender != null && !String.IsNullOrEmpty(sender[0]))
//        //    {
//        //        for (int i = 0; i <= sender.Length - 1; i++)
//        //        {
//        //            rule.Conditions.From.Recipients.Add(sender[i].ToString());
//        //            rule.Conditions.From.Recipients.ResolveAll();
//        //            rule.Conditions.From.Enabled = true;
//        //        }

//        //    }
//        //    //Subject condition
//        //    //if (!String.IsNullOrEmpty(subjectllave))
//        //    if (subjectllave != null && !String.IsNullOrEmpty(subjectllave[0]))
//        //    {
//        //        rule.Conditions.Subject.Text = subjectllave;
//        //        rule.Conditions.Subject.Enabled = true;
//        //    }
//        //    //Move action   
//        //    if (!String.IsNullOrEmpty(foldermove))
//        //    {

//        //        Carpeteo = Mapeo.Folders[root.Direccion_email].Folders[foldermove];
//        //        Outlook.MAPIFolder ruleFolder = Carpeteo;
//        //        rule.Actions.MoveToFolder.Folder = ruleFolder;
//        //        rule.Actions.MoveToFolder.Enabled = true;
//        //    }
//        //    try
//        //    {
//        //        Outlook.RuleAction stop = rule.Actions.Stop;
//        //        rule.Actions.Stop.Enabled = true;
//        //        rule.Exceptions.Subject.Text = new string[] { "RE:", "FW:" };
//        //        rule.Exceptions.Subject.Enabled = true;
//        //    }
//        //    catch (System.Exception)
//        //    {

//        //    }




//        //    rule.Enabled = true;

//        //    //Save rules
//        //    try
//        //    {
//        //        rules.Save(true);
//        //    }
//        //    catch (System.Exception ex)
//        //    {
//        //        console.WriteLine(ex.Message);
//        //        return false;
//        //    }
//        //    return true;

//        //}


//        //public void exportFolders()
//        //{
//        //    try
//        //    {

//        //        string sourceAccountEmail = "databot@gbm.net";
//        //        string destinationAccountEmail = "databotqa@gbm.net";

//        //        Outlook.Application outlookApp = new Outlook.Application();
//        //        Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");

//        //        // Log in to source and destination accounts
//        //        Outlook.Account sourceAccount = ns.Accounts.Cast<Outlook.Account>().FirstOrDefault(acc => acc.SmtpAddress.Equals(sourceAccountEmail, StringComparison.OrdinalIgnoreCase));
//        //        Outlook.Account destinationAccount = ns.Accounts.Cast<Outlook.Account>().FirstOrDefault(acc => acc.SmtpAddress.Equals(destinationAccountEmail, StringComparison.OrdinalIgnoreCase));

//        //        if (sourceAccount == null || destinationAccount == null)
//        //        {
//        //            Console.WriteLine("Source or destination account not found.");
//        //            return;
//        //        }

//        //        Outlook.Folders sourceFolders = ns.Folders[sourceAccountEmail].Folders["Procesados"].Folders;


//        //        // Recursively copy folder structure
//        //        CopyFolderStructure(sourceFolders, destinationAccount);


//        //    }
//        //    catch (System.Exception ex)
//        //    {
//        //        console.WriteLine(ex.Message);
//        //    }
//        //}

//        //public void CopyFolderStructure(Outlook.Folders sourceFolder, Outlook.Account destinationAccount)
//        //{
//        //    // Create folder with the same name in destination account
//        //    try
//        //    {

//        //        // Recursively copy subfolders
//        //        foreach (Outlook.Folder subfolder in sourceFolder)
//        //        {
//        //            try
//        //            {
//        //                console.WriteLine(subfolder.Name);
//        //                if (SearchFilesForString(@"C:\Users\dmeza\Documents\GitHub\Databotv5\", subfolder.Name))
//        //                {
//        //                    Outlook.Folder destinationFolder = destinationAccount.DeliveryStore.GetRootFolder().Folders["Procesados"] as Outlook.Folder;
//        //                    Outlook.Folder newFolder = destinationFolder.Folders.Add(subfolder.Name) as Outlook.Folder;
//        //                }
//        //            }
//        //            catch (System.Exception exs)
//        //            {
//        //                console.WriteLine(exs.Message);
//        //            }
//        //        }

//        //    }
//        //    catch (System.Exception exeses)
//        //    {
//        //        console.WriteLine(exeses.Message);
//        //    }
//        //}

//        //static bool SearchFilesForString(string directory, string searchString)
//        //{
//        //    try
//        //    {
//        //        foreach (string file in System.IO.Directory.GetFiles(directory, "*.cs", SearchOption.AllDirectories))
//        //        {
//        //            string[] lines = System.IO.File.ReadAllLines(file);
//        //            for (int lineNumber = 0; lineNumber < lines.Length; lineNumber++)
//        //            {
//        //                if (lines[lineNumber].Contains(searchString))
//        //                {
//        //                    Console.WriteLine($"Found in file: {file}, line: {lineNumber + 1}");
//        //                    Console.WriteLine(lines[lineNumber]);
//        //                    return true;
//        //                }
//        //            }
//        //        }
//        //    }
//        //    catch (System.Exception ex)
//        //    {
//        //        Console.WriteLine("An error occurred: " + ex.Message);

//        //    }
//        //    return false;
//        //}

//        //public Outlook.Folder GetRootFolderForAccount(Outlook.Account account, Outlook.Folders folders)
//        //{
//        //    foreach (Outlook.Folder folder in folders)
//        //    {
//        //        if (folder.FolderPath == account.DeliveryStore.GetRootFolder().FolderPath)
//        //        {
//        //            return folder;
//        //        }

//        //        var subFolder = GetRootFolderForAccount(account, folder.Folders);
//        //        if (subFolder != null)
//        //        {
//        //            return subFolder;
//        //        }
//        //    }

//        //    return null;
//        //}

//        #endregion
//        protected virtual void Dispose(bool disposing)
//        {
//            if (!disposedValue)
//            {
//                if (disposing)
//                {
//                    // TODO: dispose managed state (managed objects)
//                }

//                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
//                // TODO: set large fields to null
//                disposedValue = true;
//            }
//        }

//        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
//        // ~MailInteraction()
//        // {
//        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
//        //     Dispose(disposing: false);
//        // }
//        ~MailInteractionOld()
//        {
//            this.App = null;
//            this.Mapeo = null;
//            this.Carpeteo = null;
//            this.MailItems = null;
//            this.correo = null;

//        }
//        void IDisposable.Dispose()
//        {
//            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
//            Dispose(disposing: true);
//            GC.SuppressFinalize(this);
//        }



//    }
//}
