using MySql.Data.MySqlClient;
using System;
using System.Net;
using WinSCP;
using System.IO;
using DataBotV5.App.Global;
using Renci.SshNet;
using System.Runtime.InteropServices;
using DataBotV5.App.ConsoleApp;

namespace DataBotV5.Data.Database
{
    /// <summary>
    /// Clase Data encargada de la interacción con las bases de datos.
    /// </summary>
    public class Database : IDisposable
    {
        private bool disposedValue;

        Credentials.Credentials cred = new Credentials.Credentials();
        ConsoleFormat console = new ConsoleFormat();


        /// <summary>
        /// Conexión a Smart and Simple
        /// </summary>
        /// <param name="database">la base de datos o esquema</param>
        /// <param name="ambient"> PRD para productivo, QAS para QA</param>
        /// <returns></returns>
        public MySqlConnection ConnSmartSimple(string database, string ambient)
        {
            string conString = null;
            string server = null;
            string usuario = "";
            string password = cred.password_ss;
            string port = "3307";
            switch (ambient)
            {
                case "PRD":
                    usuario = cred.username_ss;
                    server = cred.PRD_SS_BASE_SERVER;
                    break;
                case "QAS":
                    usuario = cred.username_ss_qa;
                    server = cred.QA_SS_BASE_SERVER;
                    break;
                case "DEV":
                    server = cred.DEV_DATA_BASE_SERVER;
                    usuario = cred.DEV_DATA_BASE_SERVER_USER;
                    password = cred.DEV_DATA_BASE_SERVER_PASS;
                    port = "3306";
                    break;
            }
            conString = $"Server={server}; Port={port}; Database={database}; Uid={usuario}; Pwd={password}; Connection Timeout = 3960;";

            //conString = "datasource = " + server + "; User Id = " + usuario + "; Password = " + password + "; Database = " + database + "; Connection Timeout = 60";

            MySqlConnection myCon = new MySqlConnection(conString);
            return myCon;
        }
        /// <summary>
        /// Metodo para bloquear el SAP Gui Logon cuando un robot lo esta usando en un proceso, el metodo CheckSapLogin lo verifica
        /// </summary>
        /// <param name="mandante"> es el mandante que quiere bloquear o desbloquear: 120, 260, 300</param>
        /// <param name="block">0 para desbloquear, 1 para bloquear</param>
        /// <summary>
        /// 
        /// </summary>
        /// <param name="protocol">1 = Ftp, 2 = Sftp</param>
        /// <param name="hostName">link</param>
        /// <param name="port"></param>
        /// <param name="userName"></param>
        /// <param name="pass"></param>
        /// <param name="useHostkey">true = use, false = no usa</param>
        /// <returns></returns>
        public SessionOptions ConnectFTP(int protocol, string hostName, int port, string userName, string pass, bool useHostkey, string hostKey)
        {
            SessionOptions sessionOptions;

            
            if (protocol == 1)
            {
                if (useHostkey == true)
                {
                    sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        HostName = hostName,
                        PortNumber = port,
                        UserName = userName,
                        Password = pass,
                        SshHostKeyFingerprint = hostKey,
                    };
                }
                else
                {
                    sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        HostName = hostName,
                        PortNumber = port,
                        UserName = userName,
                        Password = pass,
                    };
                }

            }
            else
            {
                if (useHostkey == true)
                {
                    sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Sftp,
                        HostName = hostName,
                        PortNumber = port,
                        UserName = userName,
                        Password = pass,
                        SshHostKeyFingerprint = hostKey,
                    };
                    
                }
                else
                {
                    sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Sftp,
                        HostName = hostName,
                        PortNumber = port,
                        UserName = userName,
                        Password = pass,
                    };
                }

            }


            return sessionOptions;
        }
        public bool UploadFtp(string address, string ftpUsername, string ftpPassword, string localFile)
        {
            try
            {

                using (var client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(ftpUsername, ftpPassword);
                    string rutaserver = address + Path.GetFileName(localFile);
                    client.UploadFile(rutaserver, WebRequestMethods.Ftp.UploadFile, localFile);

                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }



        }
        public bool uploadSftp(string sourcefile, string destination, string finalFolder, [Optional] string enviroment)
        {
            if (enviroment == null)
            {
                enviroment = Start.enviroment;
            }
            string host = "";
            string username = "";
            int port = 0;
            string password = "";
            if (enviroment == "QAS")
            {
                host = cred.QA_SS_BASE_SERVER;
                username = cred.QA_SS_APP_SERVER_USER;
                password = cred.passSmartSimpleServerQA;
                port = 22;
            }
            else if (enviroment == "PRD")
            {
                host = cred.PRD_SS_APP_SERVER;
                username = cred.PRD_SS_APP_SERVER_USER;
                password = cred.passSmartSimpleServerPRD;
                port = 22;

            }
            bool resp = false;
            string finalDestination = destination + "/" + finalFolder;
            try
            {
                using (SftpClient client = new SftpClient(host, port, username, password))
                {
                    client.Connect();

                    if (!client.Exists(finalDestination))
                    {
                        client.CreateDirectory(finalDestination);
                    }

                    client.ChangeDirectory(finalDestination);
                    using (FileStream fs = new FileStream(sourcefile, FileMode.Open))
                    {
                        client.BufferSize = 4 * 1024;
                        client.UploadFile(fs, Path.GetFileName(sourcefile));
                    }
                }
                resp = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return resp;
        }
        public bool DeleteDirectoryFtp(string address, string ftpUsername, string ftpPassword)
        {
            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    HostName = "10.7.60.72",
                    UserName = ftpUsername,
                    Password = ftpPassword,
                };

                using (WinSCP.Session session = new WinSCP.Session())
                {
                    // Connect
                    session.Open(sessionOptions);

                    // Delete folder
                    session.RemoveFiles(address).Check();
                }
                //ftp://10.7.60.72/licitaciones_files/

                using (var client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

                    FtpWebRequest request1 = (FtpWebRequest)WebRequest.Create(address + "Cotizaciones Rapidas - ACOLINA.xlsx");
                    request1.Method = WebRequestMethods.Ftp.DeleteFile;
                    request1.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

                    using (FtpWebResponse response1 = (FtpWebResponse)request1.GetResponse())
                    {
                        string resp = response1.StatusDescription;
                    }


                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create(address);
                    request.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

                    request.Method = WebRequestMethods.Ftp.DeleteFile;
                    FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                    console.WriteLine("Delete status: {0}" + response.StatusDescription);
                    response.Close();

                    request.Method = WebRequestMethods.Ftp.RemoveDirectory;
                    request.GetResponse().Close();

                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }
        public string UploadFolder(string address, string ftpUsername, string ftpPassword, string directory)
        {
            string res = "OK";
            string folder = Path.GetFileName(directory);
            System.Net.NetworkCredential credenciales = new NetworkCredential(ftpUsername, ftpPassword);

            try
            {
                WebRequest request = WebRequest.Create(address + folder + "/");
                request.Method = WebRequestMethods.Ftp.MakeDirectory;
                request.Credentials = credenciales;
                FtpWebResponse resp = (FtpWebResponse)request.GetResponse();

                using (var client = new WebClient())
                {
                    string[] archivos = Directory.GetFiles(directory);
                    client.Credentials = credenciales;

                    foreach (string archivo in archivos)
                    {
                        string rutaserver = address + folder + "/" + Path.GetFileName(archivo);
                        var sada = client.UploadFile(rutaserver, WebRequestMethods.Ftp.UploadFile, archivo);

                    }
                }
                return res;
            }
            catch (Exception ex)
            {
                res = ex.Message;
                return res;
            }

        }

       /// <summary>
       /// metodo para descargar un archivo del SFTP
       /// </summary>
       /// <param name="host"></param>
       /// <param name="username"></param>
       /// <param name="password"></param>
       /// <param name="sourcefile">ruta del server</param>
       /// <param name="destination">ruta de la maquina virual del bot</param>
       /// <param name="port"></param>
       /// <returns></returns>
        public bool DownloadFileSftp(string sourcefile, string destination, [Optional] string enviroment)
        {
            try
            {
                if (enviroment == null)
                {
                    enviroment = Start.enviroment;
                }
                string host = "";
                string username = "";
                int port = 0;
                string password = "";
                if (enviroment == "QAS")
                {
                    host = cred.QA_SS_BASE_SERVER;
                    username = cred.QA_SS_APP_SERVER_USER;
                    password = cred.passSmartSimpleServerQA;
                    port = 22;
                }
                else if (enviroment == "PRD")
                {
                    host = cred.PRD_SS_APP_SERVER;
                    username = cred.PRD_SS_APP_SERVER_USER;
                    password = cred.passSmartSimpleServerPRD;
                    port = 22;

                }
                using (SftpClient client = new SftpClient(host, port, username, password))
                {
                    client.Connect();

                    if (!client.Exists(sourcefile))
                    {
                        console.WriteLine("El archivo no existe en el servidor de SS");
                        return false;
                    }
                    //Stream fileStream = File.Create(destination);
             
                    using (Stream fileStream = File.OpenWrite(destination))
                    {
                        client.DownloadFile(sourcefile, fileStream);
                    }

                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
        }
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

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~Database()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
