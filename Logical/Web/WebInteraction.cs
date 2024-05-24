using WebDriverManager.DriverConfigs.Impl;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using WebDriverManager;
using DataBotV5.Logical.Mail;
using System.IO;
using System;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Data;
using System.Web;
using System.Linq;
using DataBotV5.Logical.Files;

namespace DataBotV5.Logical.Web
{ 
    /// <summary>
    /// Clase Logical encargada de la interacción web.
    /// </summary>
    class WebInteraction
    {
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        Rooting roots = new Rooting();


        /// <summary>
        /// Genera una conexión con Selenium, dándole los argumentos y las preferencias de descarga.
        /// </summary>
        /// <param name="downloadsRoute">Ruta donde guardar las descargas</param>
        /// <param name="automatic">true: usa la última versión del driver, false: toma el driver de "Desktop\Databot\chromedriver"</param>
        /// <returns></returns>
        public IWebDriver NewSeleniumChromeDriver(string downloadsRoute = "", bool automatic = true)
        {
            try
            {
                return SeleniumConnection(downloadsRoute, automatic);
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                return SeleniumConnection(downloadsRoute, automatic);
            }
        }
        public IWebDriver NewSeleniumChromeDriver(ChromeOptions options, bool automatic = true)
        {
            try
            {
                return SeleniumConnection(options, automatic);
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                return SeleniumConnection(options, automatic);
            }
        }
        public IWebDriver SeleniumConnection(string downloadsRoute = "", bool automatic = true)
        {
            using (DestroyProcess kill = new DestroyProcess())
            {
                kill.KillProcess("chromedriver", true);
                kill.KillProcess("EXCEL", true);
                //kill.KillProcess("chrome", true);
            }

            string driverPath = roots.Google_Driver;

            if (automatic)
                driverPath = Path.GetDirectoryName(new DriverManager().SetUpDriver(new ChromeConfig(), WebDriverManager.Helpers.VersionResolveStrategy.MatchingBrowser));

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("start-maximized");                                            //Maximiza la pantallla de Chrome
            options.AddArgument("no-sandbox");                                                 //Modo primer plano
            options.AddArgument("disable-infobars");                                           //Desactivar la barra de developer
            //options.AddArgument("--crash-dumps-dir=/tmp");
            options.AddUserProfilePreference("download.prompt_for_download", false);           //Descargas inmediatas 
            options.AddUserProfilePreference("disable-popup-blocking", "true");                //Desactiva los popups
            if (downloadsRoute != "")
                options.AddUserProfilePreference("download.default_directory", downloadsRoute); //Descarga en la ruta asignada
            //options.AddAdditionalOption("useAutomationExtension", false);
            ChromeDriverService service = ChromeDriverService.CreateDefaultService(driverPath);
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;
            service.LogPath = roots.chromeLog;

            IWebDriver chrome = new ChromeDriver(service, options, new TimeSpan(0, 0, 250));

            return chrome;
        }
        /// <summary>
        /// Genera una conexión con Selenium, dándole los argumentos y las preferencias de descarga.
        /// </summary>
        /// <param name="options">Recibe ChromeOptions especificas y las agrega a las predeterminadas</param>
        /// <param name="automatic">false: usa la última versión del driver, true: toma el driver de "Desktop\Databot\chromedriver"</param>
        /// <returns></returns>
        public IWebDriver SeleniumConnection(ChromeOptions options, bool automatic = true)
        {
            using (DestroyProcess kill = new DestroyProcess())
            {
                kill.KillProcess("chromedriver", true);
                kill.KillProcess("EXCEL", true);
                //kill.KillProcess("chrome", true);
            }

            string driverPath = roots.Google_Driver;

            if (automatic)
                driverPath = Path.GetDirectoryName(new DriverManager().SetUpDriver(new ChromeConfig(), WebDriverManager.Helpers.VersionResolveStrategy.MatchingBrowser));

            ChromeDriverService service = ChromeDriverService.CreateDefaultService(driverPath);
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;
            service.LogPath = roots.chromeLog;

            options.AddArgument("start-maximized");                                            //Maximiza la pantallla de Chrome
            options.AddArgument("no-sandbox");                                                 //Modo primer plano
            options.AddArgument("disable-infobars");                                           //Desactivar la barra de developer
            options.AddUserProfilePreference("download.prompt_for_download", false);           //Descargas inmediatas 
            options.AddUserProfilePreference("disable-popup-blocking", "true");                //Desactiva los popups

            IWebDriver chrome = new ChromeDriver(service, options, new TimeSpan(0, 0, 250));

            return chrome;
        }

        /// <summary>
        /// Genera una conexión con Selenium, dándole los argumentos y las preferencias de descarga.
        /// </summary>
        /// <param name="downloadsRoute">Ruta donde guardar las descargas.</param>
        /// <param name="automatic">false: usa la última versión del driver, true: toma el driver de "Desktop\Databot\firefoxdriver"</param>
        /// <returns></returns>
        public IWebDriver NewSeleniumFirefoxDriver(string downloadsRoute = "", bool automatic = true)
        {
            string driverPath = roots.Mozilla_Driver;

            if (automatic)
                driverPath = Path.GetDirectoryName(new DriverManager().SetUpDriver(new FirefoxConfig()));

            FirefoxOptions options = new FirefoxOptions();
            if (downloadsRoute != "")
            {
                FirefoxProfile profile = new FirefoxProfile();
                profile.SetPreference("browser.download.folderList", 2);
                profile.SetPreference("browser.download.dir", roots.FilesDownloadPath);
                profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/java-archive, application/x-msexcel,application/excel,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml,application/vnd.microsoft.portable-executable");
                options.Profile = profile;
            }

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(driverPath);
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;

            IWebDriver firefox = new FirefoxDriver(service, options, new TimeSpan(0, 0, 250));

            return firefox;

        }
        /// <summary>
        /// Genera una conexión con Selenium, dándole los argumentos y las preferencias de descarga. En Firefox
        /// </summary>
        /// <param name="options">Recibe ChromeOptions especificas y las agrega a las predeterminadas</param>
        /// <param name="automatic">true: usa la última versión del driver, false: toma el driver de "Desktop\Databot\chromedriver"</param>
        /// <returns></returns>
        public IWebDriver NewSeleniumFirefoxDriver(FirefoxOptions options, bool automatic = true)
        {
            string driverPath = roots.Mozilla_Driver;

            if (automatic)
                driverPath = Path.GetDirectoryName(new DriverManager().SetUpDriver(new FirefoxConfig()));

            //FirefoxOptions options = new FirefoxOptions();
            //if (downloadsRoute != "")
            //{
            //    FirefoxProfile profile = new FirefoxProfile();
            //    profile.SetPreference("browser.download.folderList", 2);
            //    profile.SetPreference("browser.download.dir", roots.FilesDownloadPath);
            //    profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/java-archive, application/x-msexcel,application/excel,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml,application/vnd.microsoft.portable-executable");
            //    options.Profile = profile;
            //}

            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(driverPath);
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;

            IWebDriver firefox = new FirefoxDriver(service, options, new TimeSpan(0, 0, 250));

            return firefox;

        }






        /// <summary>
        /// Convierte una tabla de Selenium en Datable
        /// </summary>
        /// <param name="seleniumTable"></param>
        /// <returns></returns>
        public DataTable TableToDatatable(IWebElement seleniumTable)
        {
            DataTable tabla = new DataTable();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

            doc.LoadHtml(seleniumTable.GetAttribute("innerHTML"));

            int columnas = doc.DocumentNode.SelectNodes("//*[@class='epthc']").Count;
            int filas = doc.DocumentNode.SelectNodes("//tr").Count;

            int.TryParse(doc.DocumentNode.SelectNodes(".//tbody/tr[2]/td").LastOrDefault().GetAttributeValue("colspan", null), out int colspan);

            #region añadir columnas
            for (int i = 1; i <= columnas; i++)
            {
                string col = HttpUtility.HtmlDecode(doc.DocumentNode.SelectNodes(".//tbody/tr[1]/th[" + i.ToString() + "]").LastOrDefault().InnerText).Trim();

                col = val.RemoveSpecialChars(col, 1);
                col = col.Replace(" ", "_");
                col = col.Replace("(%)", "");

                tabla.Columns.Add(col.Trim(), typeof(string));
            }

            #endregion

            //si la tabla tiene colspan = al numero de cols en la fila 2 significa que no hay datos
            // --------------------------------------------------
            // |           |           |             |          |
            // --------------------------------------------------
            // |                    ALGO ASI                    |
            // --------------------------------------------------

            if (columnas != colspan)
            {
                #region añadir filas


                string rowspan = doc.DocumentNode.SelectNodes("//tr")[1].SelectNodes("td")[0].GetAttributeValue("rowspan", null);
                string rs_col = HttpUtility.HtmlDecode(doc.DocumentNode.SelectNodes("//tr")[1].SelectNodes("td")[0].FirstChild.InnerText).Trim();

                for (int qq = 2; qq <= filas; qq++)
                {
                    rowspan = doc.DocumentNode.SelectNodes("//tr")[qq - 1].SelectNodes("td")[0].GetAttributeValue("rowspan", null);
                    rs_col = HttpUtility.HtmlDecode(doc.DocumentNode.SelectNodes("//tr")[qq - 1].SelectNodes("td")[0].FirstChild.InnerText).Trim();

                    if (rowspan != null && rowspan != "1")
                    {
                        int span_cont = 0;
                        int filas_rowspan = int.Parse(rowspan) + qq - 1;

                        for (int i = qq; i <= filas_rowspan; i++)//for de las lineas del rowspan
                        {
                            DataRow row = tabla.NewRow();
                            for (int j = 1; j <= columnas; j++)
                            {
                                string fil;
                                if (0 == span_cont)
                                {
                                    fil = HttpUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode(".//tbody/tr[" + (i).ToString() + "]/td[" + j.ToString() + "]").InnerText).Trim();
                                    fil = fil.Replace("Motivo de anulación", "");//Replace para la columna de las partidas.
                                    fil = fil.Replace(Convert.ToChar(160), ' '); //quitar caracter &nbsp;
                                    fil = fil.Replace(Convert.ToChar(8220), ' ').Replace(Convert.ToChar(8221), ' ').Replace(Convert.ToChar(34), ' ');
                                    row[j - 1] = fil.Trim();
                                }
                                else
                                {
                                    if (j == columnas)
                                    {
                                        row[0] = rs_col.Trim();
                                    }
                                    else
                                    {
                                        fil = HttpUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode(".//tbody/tr[" + i.ToString() + "]/td[" + j.ToString() + "]").InnerText).Trim();
                                        fil = fil.Replace("Motivo de anulación", "");//Replace para la columna de las partidas
                                        fil = fil.Replace(Convert.ToChar(160), ' '); //quitar caracter &nbsp;
                                        fil = fil.Replace(Convert.ToChar(8220), ' ').Replace(Convert.ToChar(8221), ' ').Replace(Convert.ToChar(34), ' ');
                                        row[j] = fil.Trim();
                                    }
                                }
                            }
                            tabla.Rows.Add(row);
                            span_cont++;
                        }
                        qq = filas_rowspan;
                    }
                    else
                    {
                        DataRow row = tabla.NewRow();
                        for (int j = 1; j <= columnas; j++)
                        {
                            string fil = HttpUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode(".//tbody/tr[" + qq.ToString() + "]/td[" + j.ToString() + "]").InnerText).Trim();
                            fil = fil.Replace("Motivo de anulación", "");//Replace para la columna de las partidas
                            fil = fil.Replace(Convert.ToChar(160), ' '); //quitar caracter &nbsp;
                            fil = fil.Replace(Convert.ToChar(8220), ' ').Replace(Convert.ToChar(8221), ' ').Replace(Convert.ToChar(34), ' ');
                            row[j - 1] = fil.Trim();
                        }
                        tabla.Rows.Add(row);
                    }
                }
                #endregion
            }

            return tabla;
        }


    }
}

