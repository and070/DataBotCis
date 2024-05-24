using SP = Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Security;
using System.Linq;
using System.Text;
using System.IO;
using System;
using DataBotV5.Data.Root;
using DataBotV5.Data.Credentials;
using DataBotV5.App.Global;

namespace DataBotV5.Logical.MicrosoftTools
{
    /// <summary>
    /// Clase Logical encargada de SharePoint.
    /// </summary>
    class SharePoint
    {
        Rooting root = new Rooting();
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        /// <summary>
        /// Sube un archivo a Sharepoint(SP) en la carpeta Documentos
        /// </summary>
        /// <param name="siteUrl">link de SP</param>
        /// <param name="fileName">nombre del archivo</param>
        /// <param name="user">Usuario de SP</param>
        /// <param name="pass">Contraseña de SP</param>
        /// <param name="clientFolder">subfolder opcional</param>
        /// <param name="clientSubFolder">sub subfolder opcional</param>
        /// <returns></returns>
        public bool UploadFileToSharePoint(string siteUrl, string fileName, string clientFolder = "", string clientSubFolder = "")
        {
            string DocLibrary = "Documentos";
            string user = root.Direccion_email;
            string pass = cred.passOutlook;
            try
            {
                #region ConnectToSharePoint
                SecureString securePassword = new SecureString();
                foreach (char c in pass)
                    securePassword.AppendChar(c);

                SP.SharePointOnlineCredentials onlineCredentials = new SP.SharePointOnlineCredentials(user, securePassword);
                #endregion

                #region Insert the data
                using (SP.ClientContext CContext = new SP.ClientContext(siteUrl))
                {
                    CContext.Credentials = onlineCredentials;
                    SP.Web web = CContext.Web;
                    SP.FileCreationInformation newFile = new SP.FileCreationInformation();
                    byte[] FileContent = File.ReadAllBytes(fileName);
                    newFile.ContentStream = new MemoryStream(FileContent);
                    newFile.Url = Path.GetFileName(fileName);

                    SP.List DocumentLibrary = web.Lists.GetByTitle(DocLibrary); //documents core
                    SP.Folder folder;
                    SP.File uploadFile;

                    if (!string.IsNullOrEmpty(clientFolder))
                    {
                        folder = DocumentLibrary.RootFolder;
                        folder = folder.Folders.Add(clientFolder);
                        //folder.Update();
                        CContext.Load(folder);
                        CContext.ExecuteQuery(); //crea la carpeta dentro de "Documentos"

                        if (!string.IsNullOrEmpty(clientSubFolder))
                        {
                            folder = folder.Folders.Add(clientSubFolder); //la nueva carpeta
                            folder.Update();
                            CContext.Load(folder);
                            CContext.ExecuteQuery();
                            uploadFile = folder.Files.Add(newFile);
                        }
                        else
                            uploadFile = folder.Files.Add(newFile);
                    }
                    else
                        uploadFile = DocumentLibrary.RootFolder.Files.Add(newFile);

                    CContext.Load(DocumentLibrary);
                    CContext.Load(uploadFile);
                    CContext.ExecuteQuery();

                    return true;

                }
                #endregion
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Sube un archivo a Sharepoint(SP) en la ruta especificada
        /// </summary>
        /// <param name="siteUrl">link de SP</param>
        /// <param name="user">Usuario de SP</param>
        /// <param name="pass">Contraseña de SP</param>
        /// <param name="folderPath">ruta del archivo donde guardar</param>
        /// <param name="fileName">nombre del archivo</param>
        public void UploadFileToSharePointV2(string siteUrl, string folderPath, string fileName)
        {
            string user = root.Direccion_email;
            string pass = cred.passOutlook;
            List<string> folders = folderPath.Trim('/').Split('/').ToList<string>();

            #region ConnectToSharePoint
            SecureString securePassword = new SecureString();
            foreach (char c in pass)
                securePassword.AppendChar(c);
            SP.SharePointOnlineCredentials onlineCredentials = new SP.SharePointOnlineCredentials(user, securePassword);
            #endregion

            #region Insert the data
            using (SP.ClientContext CContext = new SP.ClientContext(siteUrl))
            {
                CContext.Credentials = onlineCredentials;
                SP.Web web = CContext.Web;
                SP.FileCreationInformation newFile = new SP.FileCreationInformation();
                byte[] FileContent = File.ReadAllBytes(fileName);
                newFile.ContentStream = new MemoryStream(FileContent);
                newFile.Url = Path.GetFileName(fileName);

                int subFoldersCount = folders.Count;
                SP.List DocumentLibrary = web.Lists.GetByTitle(folders[0]);
                SP.Folder folder = DocumentLibrary.RootFolder;  //documentos

                if (subFoldersCount > 1)
                {
                    for (int i = 1; i < subFoldersCount; i++)
                    {
                        folder = folder.Folders.Add(folders[i]); //el siguiente folder

                        //folder.Update();
                        CContext.Load(folder);
                        CContext.ExecuteQuery(); //cd del siguiente folder digamos o lo crea si no existe

                        if (i == subFoldersCount - 1) //si es la ultima carpeta, suba el archivo
                        {
                            SP.File uploadFile = folder.Files.Add(newFile);
                            CContext.RequestTimeout = 360000; // 6 minutos (6 * 60 * 1000)
                            CContext.Load(uploadFile);
                            try
                            {
                                CContext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                if (!ex.Message.ToLower().Contains("ya existe") && !ex.Message.ToLower().Contains("already exist"))
                                    throw ex;
                            }
                        }
                    }
                }
                else
                {
                    SP.File uploadFile = folder.Files.Add(newFile);
                    CContext.Load(uploadFile);
                    CContext.ExecuteQuery();
                }
            }
            #endregion

        }
        public static void CreateDocumentLink(SP.List list, string documentName, string documentUrl, SP.Folder Folder)
        {
            if (list is SP.List)
            {
                SP.List docLib = list;
                if (docLib.ContentTypesEnabled)
                {
                    SP.ContentType myCType = list.ContentTypes[0];
                    if (myCType != null)
                    {

                        //replace string template with values
                        string redirectAspx = RedirectAspxPage();
                        redirectAspx.Replace("{0}", documentUrl);

                        //should change the name of the .aspx file per item
                        SP.FileCreationInformation newFile = new SP.FileCreationInformation();
                        byte[] FileContent = System.IO.File.ReadAllBytes(documentName);
                        newFile.ContentStream = new MemoryStream(FileContent);
                        newFile.Url = Path.GetFileName(documentName);

                        SP.File file = Folder.Files.Add(newFile);

                        //set list item properties
                        SP.ListItem item = file.ListItemAllFields;


                        item["ContentTypeId"] = myCType.Id;
                        item.Update();

                        if (item["ContentType"].ToString() == "Link to a Document")
                        {
                            SP.FieldUrlValue fieldUrl = new SP.FieldUrlValue() { Description = documentName, Url = documentUrl };

                            item["URL"] = fieldUrl;
                            item.Update();
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Metodo para descargar un archivo de sharepoint
        /// </summary>
        /// <param name="siteUrl">Link o URL del site</param>
        /// <param name="libraryName">Nombre de la libreria, usualmente es "Documentos"</param>
        /// <param name="fileName">Nombre del archivo que se desea descargar</param>
        /// <param name="folderPath">Nombre del folder donde se encuentra el archivo a descargar (si el archivo no esta en un folder enviar null)</param>
        /// <returns></returns>
        public bool DownloadFileFromSharePoint(string siteUrl, string libraryName, string fileName, string folderPath = "")
        {
            try
            {
                string destinationFolderPath = root.FilesDownloadPath;
                string user = root.Direccion_email;
                string pass = cred.passOutlook;
                string downloadLocation = root.FilesDownloadPath;

                #region ConnectToSharePoint
                SecureString securePassword = new SecureString();
                foreach (char c in pass)
                    securePassword.AppendChar(c);
                Microsoft.SharePoint.Client.SharePointOnlineCredentials onlineCredentials = new SP.SharePointOnlineCredentials(user, securePassword);
                #endregion

                using (SP.ClientContext context = new SP.ClientContext(siteUrl))
                {
                    context.Credentials = onlineCredentials;

                    SP.Web web = context.Web;
                    SP.List library = web.Lists.GetByTitle(libraryName);
                    context.Load(library, l => l.RootFolder);
                    context.ExecuteQuery();

                    SP.Folder folder = null;
                    if (!string.IsNullOrEmpty(folderPath))
                    {
                        folder = library.RootFolder.Folders.GetByUrl(folderPath);
                        context.Load(folder);
                        context.ExecuteQuery();
                    }

                    SP.File file = (folder != null) ? folder.Files.GetByUrl(fileName) : library.RootFolder.Files.GetByUrl(fileName);
                    context.Load(file);
                    context.ExecuteQuery();

                    // Download the file and save it to the destination folder
                    string destinationFilePath = System.IO.Path.Combine(destinationFolderPath, fileName);
                    SP.FileInformation fileInfo = SP.File.OpenBinaryDirect(context, file.ServerRelativeUrl);
                    using (var fileStream = System.IO.File.Create(destinationFilePath)) { fileInfo.Stream.CopyTo(fileStream); }


                    return true; // File downloaded successfully
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while downloading the file: {ex.Message}");
                return false; // File download failed
            }
        }
        public static string RedirectAspxPage()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("<%@ Assembly Name='Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' %>");
            builder.Append("<%@ Register TagPrefix='SharePoint' Namespace='Microsoft.SharePoint.WebControls' Assembly='Microsoft.SharePoint' %>");
            builder.Append("<%@ Import Namespace='System.IO' %>");
            builder.Append("<%@ Import Namespace='Microsoft.SharePoint' %>");
            builder.Append("<%@ Import Namespace='Microsoft.SharePoint.Utilities' %>");
            builder.Append("<%@ Import Namespace='Microsoft.SharePoint.WebControls' %>");
            builder.Append("<html xmlns:mso=\"urn:schemas-microsoft-com:office:office\" xmlns:msdt=\"uuid:C2F41010-65B3-11d1-A29F-00AA00C14882\">");
            builder.Append("<head>");
            builder.Append("<meta name=\"WebPartPageExpansion\" content=\"full\" /> <meta name='progid' content='SharePoint.Link' /> ");
            builder.Append("<!--[if gte mso 9]><SharePoint:CTFieldRefs runat=server Prefix=\"mso:\" FieldList=\"FileLeafRef,URL\"><xml>");
            builder.Append("<mso:CustomDocumentProperties>");
            builder.Append("<mso:ContentTypeId msdt:dt=\"string\">0x01010A00DC3917D9FAD55147B56FF78B40FF3ABB</mso:ContentTypeId>");
            builder.Append("<mso:IconOverlay msdt:dt=\"string\">|docx|linkoverlay.gif</mso:IconOverlay>");
            builder.Append("<mso:URL msdt:dt=\"string\">{0}, {0}</mso:URL>");
            builder.Append("</mso:CustomDocumentProperties>");
            builder.Append("</xml></SharePoint:CTFieldRefs><![endif]-->");
            builder.Append("</head>");
            builder.Append("<body>");
            builder.Append("<form id='Form1' runat='server'>");
            builder.Append("<SharePoint:UrlRedirector id='Redirector1' runat='server' />");
            builder.Append("</form>");
            builder.Append("</body>");
            builder.Append("</html>");
            return builder.ToString();
        }
    }
}

