using System;
using System.Runtime.InteropServices;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using System.Net.Mail;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DataBotV5.Logical.Mail
{
    /// <summary>
    /// Clase Logical encargada de gestionar todas las interacciones con correos electrónicos.
    /// </summary>
    class MailInteraction : IDisposable
    {

        private bool disposedValue;

        Rooting root = new Rooting();
        Credentials cred = new Credentials();
        Stats esta = new Stats();
        ConsoleFormat console = new ConsoleFormat();

        /// <summary>
        /// Metodo para adquirir el token del app registrada en Azure mediante Exchange Services
        /// </summary>
        /// <param name="cca"></param>
        /// <param name="ewsScopes"></param>
        /// <returns></returns>
        public async Task<AuthenticationResult> atr(IConfidentialClientApplication cca, string[] ewsScopes)
        {
            return await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();
        }

        /// <summary>
        /// autentificarse a exchanges services para outlook con el app registrada en Azure
        /// </summary>
        /// <returns></returns>
        public ExchangeService exchangeAuth()
        {

            ExchangeService ewsClient = new ExchangeService();

            // Using Microsoft.Identity.Client 4.22.0
            var cca = ConfidentialClientApplicationBuilder
                .Create(cred.clientId)
                .WithClientSecret(cred.clientSecret)
                .WithTenantId(cred.tenantId)
                .Build();


            //The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            //Make the interactive token request
            Task<AuthenticationResult> authResult = atr(cca, ewsScopes);

            authResult.Wait();

            ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            ewsClient.Credentials = new OAuthCredentials(authResult.Result.AccessToken);

            ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, root.Direccion_email);
            return ewsClient;
        }

        /// <summary>
        /// Metodo ayuda para buscar el ID de Outlook del Folder dado por nombre
        /// </summary>
        /// <param name="service"></param>
        /// <param name="folderName"></param>
        /// <param name="subfolderName"></param>
        /// <param name="subsubfolderName"></param>
        /// <returns></returns>
        private FolderId FindOrCreateFolder(ExchangeService service, string folderName, string subfolderName = null, string subsubfolderName = null)
        {
            //El folder root de databot
            Folder rootFolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);
            //una colección de todos los folders del databot
            FindFoldersResults subfolders = rootFolder.FindFolders(new FolderView(int.MaxValue));
            //incializar variable FolderId
            FolderId parentFolderId = WellKnownFolderName.Inbox;
            //por cada folder de la colección se busca el folder indicado por el parametro y asigno el id del mismo
            foreach (Folder folder in subfolders)
            {
                if (folder.DisplayName.ToLower() == folderName.ToLower())
                {
                    parentFolderId = folder.Id;
                    break;
                }
            }

            //si el sub folder es diferente a blanco (usualmente es Procesados XXXX) 
            if (!string.IsNullOrEmpty(subfolderName))
            {
                FolderId subfolderId = FindOrCreateSubfolder(service, parentFolderId, subfolderName);
                parentFolderId = subfolderId;
            }
            //Si el sub sub folder de procesados es diferente a blanco
            if (!string.IsNullOrEmpty(subsubfolderName))
            {
                FolderId subsubfolderId = FindOrCreateSubfolder(service, parentFolderId, subsubfolderName);
                parentFolderId = subsubfolderId;
            }

            return parentFolderId;
        }

        /// <summary>
        /// Metodo para buscar un subFolder
        /// </summary>
        /// <param name="service"></param>
        /// <param name="parentFolderId"></param>
        /// <param name="subfolderName"></param>
        /// <returns></returns>
        private FolderId FindOrCreateSubfolder(ExchangeService service, FolderId parentFolderId, string subfolderName)
        {
            FolderId fid = WellKnownFolderName.Inbox;
            Folder parentFolder = Folder.Bind(service, parentFolderId);
            FindFoldersResults subfolders = parentFolder.FindFolders(new FolderView(int.MaxValue));

            foreach (Folder folder in subfolders)
            {
                if (folder.DisplayName.ToLower() == subfolderName.ToLower())
                {
                    parentFolderId = folder.Id;
                    break;
                }
            }
            return fid;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>
        /// Retorna una variable tipo bool para verificar si descargó algún adjunto de correo(verificando si root.NArchivo_Excel != null).
        /// </returns>
        /// <param name="folderRequestToProcess">Localiza la carpeta donde están los correos de solicitudes por procesar.</param>
        /// <param name="folderProcesed">Localiza la carpeta donde se desea almacenar los correos procesados.</param>
        /// <param name="subfolderProcesed"></param>
        /// <param name="subsubfolderProcesed"></param>
        public bool GetAttachmentEmail(string folderRequestToProcess, string folderProcessed = null, string subfolderProcessed = null, string subsubfolderProcessed = null)
        {
            try
            {
                ExchangeService ewsClient = exchangeAuth();
                //find folder id of the "Solicitudes" folder
                FolderId folderRequestToProcessId = FindOrCreateFolder(ewsClient, folderRequestToProcess);
                // Find unread emails in the specified folder
                FindItemsResults<Item> unreadEmails = ewsClient.FindItems(
                    folderRequestToProcessId,
                    new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                    new ItemView(1)
                    {
                        //PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments),
                        Traversal = ItemTraversal.Shallow,
                        //OrderBy = new SortOrder(ItemSchema.DateTimeReceived, SortDirection.Descending),
                    });


                if (unreadEmails.TotalCount == 0)
                {
                    console.WriteLine("No unread emails found in the specified folder.");
                    return false;
                }

                // Process the first unread email found
                EmailMessage unreadEmail = unreadEmails.Items[0] as EmailMessage;

                //take info
                unreadEmail.Load();
                root.BDUserCreatedBy = unreadEmail.Sender.Address;
                root.Subject = unreadEmail.Subject;
                root.ReceivedTime = unreadEmail.DateTimeReceived;
                root.EmailObject = unreadEmail;
                root.recipientes = unreadEmail.ToRecipients.Select(r => r.Address).ToArray();
                root.CopyCC = unreadEmail.CcRecipients.Select(r => r.Address).ToArray();
                string htmlContent = unreadEmail.Body.Text;
                string plainText = Regex.Replace(htmlContent, @"<br\s*/?>", "\r\n");
                plainText = Regex.Replace(plainText, "<.*?>", String.Empty);
                root.Email_Body = plainText;

                // Download and save attachments (if any)
                if (unreadEmail.HasAttachments)
                {

                    List<string> attachs = new List<string>();
                    foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in unreadEmail.Attachments)
                    {
                        if (attachment is FileAttachment fileAttachment && !attachment.IsInline)
                        {
                            root.ExcelFile = fileAttachment.Name;
                            attachs.Add(fileAttachment.Name);
                            // Download and save the attachment
                            string attachmentFilePath = Path.Combine(root.FilesDownloadPath, fileAttachment.Name);
                            fileAttachment.Load(attachmentFilePath);
                            //console.WriteLine($"Attachment '{fileAttachment.Name}' saved to '{attachmentFilePath}'.");
                        }
                    }
                    root.filesList = attachs.ToArray();
                }

                // Mark the email as read
                unreadEmail.IsRead = true;
                unreadEmail.Update(ConflictResolutionMode.AutoResolve);

                // Move the email to the processed folder
                if (!string.IsNullOrEmpty(folderProcessed))
                {
                    FolderId processedFolderId = FindOrCreateFolder(ewsClient, folderProcessed, subfolderProcessed, subsubfolderProcessed);
                    if (processedFolderId != null)
                    {
                        unreadEmail.Move(processedFolderId);
                    }
                    else
                    {
                        console.WriteLine("Failed to move the email to the processed folder.");
                    }
                }

                return true;
            }
            catch (System.Exception ex)
            {
                console.WriteAnnounce(ex.Message);
                return false;
            }
        }

        /// <summary>Enviar correo electrónico en formato HTML.</summary>
        public bool SendHTMLMail(string html, string[] to, string subject, [Optional] string[] cc, [Optional] string[] attachments)
        {
            try
            {
                string pattern = @"<[^>]+>";
                string htmlBody = html;
                //check if the html indeed had html tags
                if (!Regex.IsMatch(htmlBody, pattern))
                {
                    string htmlpage = Properties.Resources.emailtemplate1;

                    htmlBody = htmlpage.Replace("{subject}", subject).Replace("{cuerpo}", html).Replace("{contenido}", "");
                }

                ExchangeService ewsClient = exchangeAuth();

                EmailMessage email = new EmailMessage(ewsClient);

                foreach (string emailAddress in to)
                {
                    email.ToRecipients.Add(emailAddress);

                }

                email.Subject = subject;
                email.Body = new MessageBody(Microsoft.Exchange.WebServices.Data.BodyType.HTML, htmlBody);

                if (cc != null)
                {
                    foreach (string ccAddress in cc)
                    {
                        email.CcRecipients.Add(ccAddress);
                    }
                }

                if (attachments != null)
                {
                    foreach (string attachmentPath in attachments)
                    {
                        email.Attachments.AddFileAttachment(attachmentPath);
                    }
                }

                email.SendAndSaveCopy(WellKnownFolderName.SentItems);
                return true;
            }
            catch (System.Exception ex)
            {
                console.WriteAnnounce(ex.Message);
                SendHTMLMailSmtp(subject, html, to, cc, attachments);
                return false;
            }
        }

        /// <summary>Reenvíar correo electrónico.</summary>
        public bool ForwardEmail(EmailMessage forwardEmail, string to, [Optional] string[] cc, [Optional] string[] attachments)
        {
            try
            {
                //root.EmailObject //variable global que se asigna al leer un email 
                forwardEmail.Subject = "Fwd: " + forwardEmail.Subject;
                forwardEmail.Body = new MessageBody(forwardEmail.Body);
                forwardEmail.ToRecipients.Add(to);
                if (cc != null)
                {
                    foreach (string ccAddress in cc)
                    {
                        forwardEmail.CcRecipients.Add(ccAddress);
                    }
                }
                forwardEmail.SendAndSaveCopy(WellKnownFolderName.SentItems);
                return true;
            }
            catch (System.Exception ex)
            {
                console.WriteAnnounce(ex.Message);
                return false;
            }
        }

        /// <summary>Enviar correo electrónico atraves de SMTP outlook.</summary>
        public void SendHTMLMailSmtp(string subject, string msj, string[] to, [Optional] string[] cc, [Optional] string[] attachments)
        {

            using (SmtpClient client = new SmtpClient()
            {
                Host = "smtp.office365.com",
                Port = 587, //22
                UseDefaultCredentials = false, // This require to be before setting Credentials property
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials = new NetworkCredential(root.Direccion_email, cred.passOutlook), // you must give a full email address for authentication 
                TargetName = "STARTTLS/smtp.office365.com", // Set to avoid MustIssueStartTlsFirst exception
                EnableSsl = true // Set to avoid secure connection exception
            })
            {

                MailMessage message = new MailMessage()
                {
                    From = new MailAddress(root.Direccion_email), // sender must be a full email address
                    Subject = subject,
                    IsBodyHtml = true,
                    Body = msj,
                    BodyEncoding = System.Text.Encoding.UTF8,
                    SubjectEncoding = System.Text.Encoding.UTF8,

                };
                foreach (string item in to)
                {
                    message.To.Add(item);

                }
                if (cc != null && cc[0] != null)
                {
                    foreach (string ccMail in cc)
                    {
                        message.CC.Add(ccMail);
                    }
                }



                client.Send(message);

            }

        }

        #region Metodos de proyectos en especifico
        /// <summary>Método para obtener solicitudes aprobadas.</summary>
        public Dictionary<string, string> GetApprovalRequests(string process)
        {
            Dictionary<string, string> infoRet = new Dictionary<string, string>();

            ExchangeService ewsClient = exchangeAuth();
            //find folder id of the "Solicitudes" folder
            FolderId folderId = FindOrCreateFolder(ewsClient, "Solicitudes Aprobaciones de Power automate");
            // Find unread emails in the specified folder
            FindItemsResults<Item> unreadEmails = ewsClient.FindItems(
                folderId,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                new ItemView(1)
                {
                    //PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments),
                    Traversal = ItemTraversal.Shallow,
                    //OrderBy = new SortOrder(ItemSchema.DateTimeReceived, SortDirection.Descending),
                });


            if (unreadEmails.TotalCount == 0)
            {
                console.WriteLine("No unread emails found in the specified folder.");
                return infoRet;
            }

            foreach (EmailMessage urEmail in unreadEmails.Items)
            {
                urEmail.Load();
                if (urEmail.Subject.Replace("appr_request_", "") == process)
                {
                    infoRet.Add("ResponseJson", urEmail.Body);
                    // Download and save attachments (if any)
                    if (urEmail.HasAttachments)
                    {

                        foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in urEmail.Attachments)
                        {
                            if (attachment is FileAttachment fileAttachment)
                            {

                                string attachFile = fileAttachment.Name.ToString();
                                string fileExtension = Path.GetExtension(attachFile);

                                int charFile = attachFile.Length - fileExtension.Length;
                                if (charFile > 80)
                                    attachFile = attachFile.Substring(0, 80) + fileExtension;

                                string fullPath = root.FilesDownloadPath + @"\" + attachFile;
                                fileAttachment.Load(fullPath);

                                if (fileExtension == ".json")
                                    infoRet.Add("OriginalJson", String.Join("", System.IO.File.ReadAllLines(fullPath)));
                                else
                                    infoRet.Add("AttachmentPath", fullPath);

                            }
                        }
                    }

                    // Mark the email as read
                    urEmail.IsRead = true;
                    urEmail.Update(ConflictResolutionMode.AutoResolve);

                    // Mover el correo a procesados y terminar
                    try
                    {
                        FolderId processedFolderId = FindOrCreateFolder(ewsClient, "Procesados", "Procesados Aprobaciones de Power automate");
                        if (processedFolderId != null)
                        {
                            urEmail.Move(processedFolderId);
                        }
                        else
                        {
                            console.WriteLine("Failed to move the email to the processed folder.");
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }

            return infoRet;
        }

        /// <summary>Método para obtener el reporte Cognos de Tarifas mas reciente </summary>
        public string GetLastPriceReportMail()
        {
            string fullPath = "";

            ExchangeService ewsClient = exchangeAuth();
            //find folder id of the "Solicitudes" folder
            FolderId folderId = FindOrCreateFolder(ewsClient, "Reporte Tarifas");
            // Find unread emails in the specified folder
            FindItemsResults<Item> unreadEmails = ewsClient.FindItems(
                folderId,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                new ItemView(1)
                {
                    //PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments),
                    Traversal = ItemTraversal.Shallow,
                    //OrderBy = new SortOrder(ItemSchema.DateTimeReceived, SortDirection.Descending),
                });


            if (unreadEmails.TotalCount == 0)
            {
                console.WriteLine("No unread emails found in the specified folder.");
                return fullPath;
            }

            EmailMessage urEmail = unreadEmails.Items.LastOrDefault() as EmailMessage;

            if (urEmail != null)
            {
                urEmail.Load();
                // Download and save attachments (if any)
                if (urEmail.HasAttachments)
                {

                    foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in urEmail.Attachments)
                    {
                        if (attachment is FileAttachment fileAttachment)
                        {

                            string attachFile = fileAttachment.Name.ToString();


                            if (attachFile.Contains("Tarifas HCM Distribucion"))
                            {
                                string fileExtension = Path.GetExtension(attachFile);

                                int charFile = attachFile.Length - fileExtension.Length;
                                if (charFile > 80)
                                    attachFile = attachFile.Substring(0, 80) + fileExtension;

                                fullPath = root.FilesDownloadPath + @"\" + attachFile;
                                fileAttachment.Load(fullPath);
                            }


                        }
                    }
                }


            }

            return fullPath;
        }
        #endregion
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        ~MailInteraction()
        {

        }
        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }



    }
}
