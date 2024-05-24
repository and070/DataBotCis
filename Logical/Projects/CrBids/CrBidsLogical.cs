using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections.ObjectModel;
using System.Web;
using SAP.Middleware.Connector;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using MySql.Data.MySqlClient;
using System.Globalization;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Support.UI;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.Web;
using DataBotV5.Automation.RPA2.CrBids;
using DataBotV5.App.Global;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Firefox;

namespace DataBotV5.Logical.Projects.CrBids
{
    /// <summary>
    /// Clase Logical con todos los procesos logicos de selenium o de apoyo para el bot de licitaciones de CR.
    /// </summary>
    class CrBidsLogical
    {
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        ProcessAdmin padmin = new ProcessAdmin();
        Rooting roots = new Rooting();
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        
        BidsGbCrSql liccr = new BidsGbCrSql();
        MailInteraction mail = new MailInteraction();
        Database db2 = new Database();
        Rooting root = new Rooting();
        WebexTeams wt = new WebexTeams();
        CRUD crud = new CRUD();
        string downloadfolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\LICCR\\DOWNLOADS";
        string optionsfolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\LICCR\\chrome-options";


        string mandanteSAPCRM = "CRM";
        string mandanteSAPERP = "ERP";

        /// <summary>
        /// Crea una conexión de selenium
        /// </summary>
        /// <param name="url">La url de la pagina de SICOP</param>
        /// <returns></returns>
        public IWebDriver SelConn(string url)
        {
            #region eliminar cache and cookies chrome
            try
            {
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);
            }
            catch (Exception)
            { }
            #endregion

            ChromeOptions options = new ChromeOptions();
            WebInteraction sel = new WebInteraction();

            options.AddArguments("user-data-dir=" + optionsfolder);
            //options.SetPreference("browser.download.folderList", 2); // 2 represents custom location
            //options.SetPreference("browser.download.dir", downloadfolder);
            //options.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream"); // Optional: Set the MIME type of files to automatically save
            options.AddUserProfilePreference("download.default_directory", downloadfolder);

            IWebDriver chrome = sel.NewSeleniumChromeDriver(options);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl(url);
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl(url); }

            chrome.Manage().Cookies.DeleteAllCookies();
            #endregion

            return chrome;
        }
        #region Métodos OnlyDownload y ModConcourse modificados para evitar conflictos en la migracion a SmartAndSimple, cuando se terminen de probar las migraciones se borran los antiguos.



        public bool OnlyDownload2(string concourse, IWebDriver chrome, bool chromeActive, string enviroment)
        {
            bool resp = false; //Se inicializa false para cuando el FTP inserte la resp sea true.
            try
            {
                #region Borrar las descargas anteriores
                try
                {
                    Directory.Delete(downloadfolder, true);
                    Directory.CreateDirectory(downloadfolder);
                }
                catch (Exception) { }
                #endregion
                DataTable archivosActuales =
                crud.Select( $"SELECT name FROM `uploadFiles` WHERE `bidNumber` LIKE '{concourse}'", "costa_rica_bids_db");//tomar los concursos actuales

                List<string> currentFiles = archivosActuales.AsEnumerable().Select(r => r.Field<string>("name")).ToList();
                HtmlAgilityPack.HtmlDocument pagina_actual = new HtmlAgilityPack.HtmlDocument();
                if (!chromeActive)
                {
                    IWebElement topFrame = chrome.FindElement(By.XPath("//*[@id='topFrame']"));
                    chrome.SwitchTo().Frame(topFrame);
                    System.Threading.Thread.Sleep(1000);



                    chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[4]/ul/li[2]/div[1]/a[3]")).Click(); //concursos tab



                    chrome.SwitchTo().ParentFrame();
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    chrome.SwitchTo().DefaultContent();



                    IWebElement mainFrame = chrome.FindElement(By.XPath("//*[@id='mainFrame']"));
                    chrome.SwitchTo().Frame(mainFrame);
                    System.Threading.Thread.Sleep(1000);



                    IWebElement frame = chrome.FindElement(By.XPath("//*[@id='rightFrame']"));



                    chrome.SwitchTo().Frame(frame);
                    System.Threading.Thread.Sleep(1000);



                    //Aqui van los filtros de busqueda
                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div[2]/form[2]/table/tbody/tr[2]/td/input")).SendKeys(concourse);
                    chrome.FindElement(By.XPath("/html/body/div[1]/div/div[2]/form[2]/table/tbody/tr[2]/td/span/a")).Click(); //consultar boton




                    pagina_actual.LoadHtml(chrome.PageSource);
                    HtmlAgilityPack.HtmlNodeCollection eptable = pagina_actual.DocumentNode.SelectNodes("//*[@class='eptable']/tbody/tr"); //tomar toda la tabla
                    string link = eptable[1].SelectSingleNode("*//a[contains(@href, 'js_cartelSearch')]").GetAttributeValue("href", null).Replace("javascript:", ""); //uno por que solo deberia haber una linea
                    IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                    js.ExecuteScript(link);
                }
                List<string> linkAttachments = new List<string>();
                pagina_actual.LoadHtml(chrome.PageSource);
                HtmlAgilityPack.HtmlNodeCollection adjuntos_nodes = pagina_actual.DocumentNode.SelectNodes("//a[contains(@href, 'js_downloadFile')]");
                if (adjuntos_nodes != null)
                {
                    foreach (HtmlAgilityPack.HtmlNode adjunto in adjuntos_nodes)
                    {
                        string href = adjunto.GetAttributeValue("href", string.Empty);
                        string docName = adjunto.InnerText.ToString();
                        if (archivosActuales.Rows.Count > 0)
                        {
                            //Valida si el archivo ya ha sido descargado, para no agregarlo a la lista currentFiles.
                            if (archivosActuales.Select("name = '" + docName + "'").Count() == 0)
                            {
                                linkAttachments.Add(href.Replace("javascript:", ""));
                                currentFiles.Add(docName);
                            }
                        }
                        else
                        {
                            linkAttachments.Add(href.Replace("javascript:", ""));
                            currentFiles.Add(docName);
                        }



                    }
                }



                List<string> newAttachmentsNames = DownloadAttachment(chrome, linkAttachments, concourse); //nombre de los nuevos adjuntos descargados
                byte[] zip = null;

                #region Subir Archivos al FTP
                newAttachmentsNames.ForEach(delegate (string name)
                {
                    //bool ins = liccr.InsertFile(concourse, root.downloadfolder + "\\" + concourse + "\\" + name);
                    resp = liccr.InsertFile(concourse, root.downloadfolder + "\\" + concourse + "\\" + name, enviroment);
                });



                #endregion




                #region Borrar las descargas anteriores
                try
                {
                    Directory.Delete(downloadfolder, true);
                    Directory.CreateDirectory(downloadfolder);
                }
                catch (Exception) { }
                #endregion



                //resp = true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                resp = false;
            }
            if (!chromeActive)
            {
                chrome.Close();
                proc.KillProcess("chromedriver", true);
            }
            return resp;



        }
        public ModifyConcourse ModConcourse2(IWebDriver chrome, string concourse, bool downloads, string enviroment)
        {
            ModifyConcourse modc = new ModifyConcourse();
            try
            {
                #region buscador avanzado
                IWebElement topFrame = chrome.FindElement(By.XPath("//*[@id='topFrame']"));
                chrome.SwitchTo().Frame(topFrame);
                System.Threading.Thread.Sleep(1000);



                chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[4]/ul/li[2]/div[1]/a[3]")).Click(); //concursos tab
                                                                                                               //chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[3]/a[2]")).Click(); //proveedores prueba



                chrome.SwitchTo().ParentFrame();
                chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                chrome.SwitchTo().DefaultContent();



                IWebElement mainFrame = chrome.FindElement(By.XPath("//*[@id='mainFrame']"));
                chrome.SwitchTo().Frame(mainFrame);
                System.Threading.Thread.Sleep(1000);



                //chrome.FindElement(By.XPath("//*[@id='overlay']/div/button[1]")).Click(); //Boton de firma digital (prueba)



                IWebElement frame = chrome.FindElement(By.XPath("//*[@id='rightFrame']"));
                chrome.SwitchTo().Frame(frame);
                System.Threading.Thread.Sleep(1000);



                //Aqui van los filtros de busqueda



                chrome.FindElement(By.Id("regDtFrom")).Clear();




                chrome.FindElement(By.Id("regDtTo")).Clear();



                chrome.FindElement(By.Name("cartelNoNm")).SendKeys(concourse);
                //chrome.FindElement(By.XPath("/html/body/div[1]/div/div[2]/p/span/a")).Click(); //consultar boton
                chrome.FindElement(By.XPath("/html/body/div[1]/div/div[2]/form[2]/table/tbody/tr[2]/td/span/a")).Click(); //consultar boton
                System.Threading.Thread.Sleep(3000);
                #endregion
                #region Busqueda de nueva info
                HtmlAgilityPack.HtmlDocument pagina_actual = new HtmlAgilityPack.HtmlDocument();
                pagina_actual.LoadHtml(chrome.PageSource);
                HtmlAgilityPack.HtmlNodeCollection eptable = pagina_actual.DocumentNode.SelectNodes("//*[@class='eptable']/tbody/tr"); //tomar toda la tabla
                                                                                                                                       //string fechaApertura = eptable[1].SelectSingleNode("td[4]").FirstChild.InnerText.Trim();
                string fechaApertura = eptable[1].SelectSingleNode("td[4]").InnerText.Trim();
                modc.OpenningDate = fechaApertura;
                #endregion
                #region nueva documentacion
                if (downloads)
                {
                    string link = eptable[1].SelectSingleNode("*//a[contains(@href, 'js_cartelSearch')]").GetAttributeValue("href", null).Replace("javascript:", ""); //uno por que solo deberia haber una linea
                    IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                    js.ExecuteScript(link);
                    System.Threading.Thread.Sleep(1000);
                    modc.NewDownloads = OnlyDownload2(concourse, chrome, true, enviroment);



                }
                #endregion
                #region Cerrar chrome
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                #endregion
                modc.Val = true;
            }
            catch (Exception)
            {
                modc.Val = false;
                modc.NewDownloads = false;
            }
            return modc;
        }
        #endregion

        /// <summary>
        /// Obtiene si un texto hace match con una lista de palabras, devuelve "SI" si encontro match y "NO" si no encontro match
        /// </summary>
        /// <param name="text">El texto a evaluar</param>
        /// <param name="words">Lista con las palabras a buscar</param>
        /// <returns></returns>
        public string KeyMatch(string text, List<string> words)
        {
            string interes_gbm = "NO";
            try
            {
                bool result2 = false;
                text = val.RemoveAccents(text);
                foreach (string word in words)
                {
                    result2 = Regex.IsMatch(text, $"\\b{word}\\b");
                    if (result2)
                    {
                        interes_gbm = "SI";
                        break;
                    }
                }
                //var result = words.Where(x => text.Contains(x)).ToList();
                //if (result.Count > 0)
                //{ interes_gbm = "SI"; }
            }
            catch (Exception)
            {
                interes_gbm = "SI";
            }

            return interes_gbm;
        }
        /// <summary>
        /// Descarga los achivos adjuntos de un concurso
        /// </summary>
        /// <param name="chrome">Webdriver que se esta usando</param>
        /// <param name="file_num">El numero de archivos adjuntos que tiene el concurso</param>
        /// <param name="folder">El nombre de la carpeta donde se guardaran los archivos descargados</param>
        /// <returns></returns>
        public List<string> DownloadAttachment(IWebDriver chrome, List<string> javascript, string folder)
        {
            List<string> respuesta = new List<string>();
            List<string> files = new List<string>();
            try
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                Directory.CreateDirectory(downloadfolder + "\\" + folder);
                string[] busqueda = { ".tmp", ".crdownload" };

                foreach (string item in javascript)
                {
                    try
                    {
                        js.ExecuteScript(item + ";");
                    }
                    catch (UnhandledAlertException)
                    {
                        System.Threading.Thread.Sleep(50);
                        js.ExecuteScript(item + ";");
                    }
                }
                files = Directory.GetFiles(downloadfolder, "*.*").ToList();
                List<String> result = files.Where(r => busqueda.Any(t => r.Contains(t))).ToList();

                Stopwatch sw = new Stopwatch();
                sw.Start();
                while (result.Count != 0)
                {
                    System.Threading.Thread.Sleep(50);
                    files = Directory.GetFiles(downloadfolder, "*.*").ToList();
                    result = files.Where(r => busqueda.Any(t => r.Contains(t))).ToList();
                    if (sw.ElapsedMilliseconds > 60000) throw new TimeoutException();
                }

                foreach (string file in files)
                {
                    FileInfo mFile = new FileInfo(file);
                    mFile.MoveTo(downloadfolder + "\\" + folder + "\\" + mFile.Name);
                    string tmp = Path.GetFileNameWithoutExtension(file) + Path.GetExtension(file);
                    respuesta.Add(tmp);
                }

            }
            catch (Exception ex)
            {
                files = Directory.GetFiles(downloadfolder, "*.*").ToList();
                foreach (string file in files)
                {
                    File.Delete(file);
                }
                //borrar los files
                respuesta.Clear();
                respuesta.Add(ex.Message);
            }

            return respuesta;
        }

        /// <summary>
        /// extrae el nombre de un cliente o contacto de CRM
        /// </summary>
        /// <param name="bp">el id de CRM del business partner (bp)</param>
        /// <param name="type">1 para cliente, 2 para contacto</param>
        /// <returns></returns>
        public string GetInfoBP(string bp, int type)
        {
            string name = "";
            try
            {
                Dictionary<string, string> parametros = new Dictionary<string, string>();
                parametros["BP"] = bp;

                IRfcFunction func = new SapVariants().ExecuteRFC(mandanteSAPCRM, "ZDM_READ_BP", parametros);
                if (type == 1)
                {
                    //cliente 
                    name = func.GetValue("NOMBRE").ToString();
                }
                else
                {
                    //eduardo piedra
                    name = func.GetValue("FIRSTNAME").ToString() + " " + func.GetValue("LASTNAME").ToString() ;
                }
            }
            catch (Exception)
            {
                name = "";
            }

            return name;
        }


    }
}
