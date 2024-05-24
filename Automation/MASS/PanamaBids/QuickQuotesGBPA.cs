using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.PanamaBids;
using DataBotV5.Data.Projects.PanamaBids;
using DataBotV5.App.Global;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Data.Database;

namespace DataBotV5.Automation.MASS.PanamaBids
{
    /// <summary>
    /// Clase MASS Automation encargada de extraer la información de cotizaciones programadas y abiertas del portal del Gobierno de Panamá (3 veces al día).
    /// </summary>
    class QuickQuotesGBPA
    {
        #region variables_globales
        PanamaPurchase pa_compra = new PanamaPurchase();
        Stats estadisticas = new Stats();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        ProcessAdmin padmin = new ProcessAdmin();
        MsExcel MsExcel = new MsExcel();
        ValidateData val = new ValidateData();
        WebInteraction sel = new WebInteraction();
        WebexTeams wt = new WebexTeams();
        Database wb2 = new Database();
        Log log = new Log();
        string respuesta = "";
        string SSMandante = "QAS";

        string respFinal = "";

        #endregion


        public void Main() //cotizaciones_rapidas
        {
            string respuesta = GetGotiRapidas();
            using (Stats stats = new Stats())
            {
                stats.CreateStat();
            }
        }
        private string GetGotiRapidas()
        {
            #region variables privadas

            //crea una lista con todas las palabras claves de la base de datos.
            string singleOrderRecord = "";
            List<string> words = lpsql.KeyWords("");
            DataTable CotiAll = lpsql.AllQuickQuotes();
            DataTable entitiesInfo = lpsql.entitiesInfo();
            Dictionary<string, string> adj_names = new Dictionary<string, string>();
            Dictionary<string, DataTable> excelsAM = new Dictionary<string, DataTable>();
            string[] adjunto = new string[1];
            string[] AMarrays = new string[1];
            int cont_adj = 0;
            int cont_am = 0;
            int cont = 1;
            string resp_sql = "";
            string respMAMBURG = "";
            string respERUIZ = "";
            bool resp_add_sql = true;
            bool validar_lineas = true;
            #endregion
            #region eliminar cache and cookies chrome
            try
            {
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);
                padmin.DeleteFiles(@"C:\Users\" + Environment.UserName + @"\AppData\Local\Google\Chrome\User Data\Default\Cache\");
                string cookies = @"C:\Users\" + Environment.UserName + @"\AppData\Local\Google\Chrome\User Data\Default\Cookies";
                string cookiesj = @"C:\Users\" + Environment.UserName + @"\AppData\Local\Google\Chrome\User Data\Default\Cookies-journal";
                if (File.Exists(cookies))
                { File.Delete(cookies); }
                if (File.Exists(cookiesj))
                { File.Delete(cookiesj); }
            }
            catch (Exception)
            { }
            #endregion
            #region Abrir excel
            DataTable excelResults = MsExcel.GetExcel(root.quickQuoteReport);
            if (excelResults == null)
            {
                mail.SendHTMLMail("Error al leer la plantilla de Cotizaciones Rapidas", new string[] {"appmanagement@gbm.net"}, "Error al leer la plantilla de Cotizaciones Rapidas", new string[] { "dmeza@gbm.net" });
                return "";
            }
            #endregion
            #region Ingreso al website
            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#/busqueda-avanzada");
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#/busqueda-avanzada"); }
            //https://www.panamacompra.gob.pa/Inicio/#!/busquedaAvanzada?BusquedaTipos=True&IdTipoBusqueda=53&estado=51&title=Cotizaciones%20en%20L%C3%ADnea
            //js executor para subir al inicio de pagina
            IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
            chrome.Manage().Cookies.DeleteAllCookies();
            #endregion
            #region extraer info abierta
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='toTopBA']/h5/b"))); }
            catch { }
            string cant_filas = "";
            try
            { cant_filas = chrome.FindElement(By.XPath("//*[@id='toTopBA']/h5/b")).Text; }
            catch (Exception)
            { cant_filas = ""; }
            if (cant_filas == "Se encontraron 0 Cotizaciones Abiertas")
            {
                //no hay cotizaciones abiertas 
            }
            else
            {
                //si hay cotizaciones abiertas

            }
            #endregion
            #region extraer cotis programada

            #region buscar programadas
            System.Threading.Thread.Sleep(2000);
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[2]/label[4]"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[2]/label[4]")).Click(); //click en el boton Programadas
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/center/pre/small[2]"))); }
            catch { }

            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='toTopBA']/h5/b"))); }
            catch { }
            string cant_filas_prog = "";
            try
            { cant_filas_prog = chrome.FindElement(By.XPath("//*[@id='toTopBA']/h5/b")).Text; }
            catch (Exception)
            { cant_filas_prog = ""; }
            #endregion

            if (cant_filas_prog != "Se encontraron 0 Cotizaciones Programadas")
            {

                //si hay cotizaciones programadas
                #region extraer la cantidad de registros
                int mas = 0;
                if (cant_filas_prog.Contains("+"))
                {
                    do
                    {
                        IWebElement pagination_last = chrome.FindElement(By.ClassName("pagination-last"));
                        var pag_last = pagination_last.FindElement(By.TagName("a"));
                        pag_last.Click();
                        System.Threading.Thread.Sleep(1500);
                        cant_filas_prog = chrome.FindElement(By.XPath("//*[@id='toTopBA']/h5/b")).Text;
                        mas++;
                    } while (cant_filas_prog.Contains("+"));



                    if (mas > 0)
                    {
                        IWebElement pagination_first = chrome.FindElement(By.ClassName("pagination-first"));
                        var pag_first = pagination_first.FindElement(By.TagName("a"));
                        pag_first.Click();
                        System.Threading.Thread.Sleep(1000);
                        cant_filas_prog = chrome.FindElement(By.XPath("//*[@id='toTopBA']/h5/b")).Text;
                    }
                }
                cant_filas_prog = cant_filas_prog.Substring(15, 3).Trim();
                if (cant_filas_prog.Contains(" "))
                {
                    cant_filas_prog = cant_filas_prog.Substring(0, 2).Trim();
                }

                double num1 = double.Parse(cant_filas_prog);
                double num2 = 10;
                double filas = 0;
                filas = (num1 / num2);
                double pag_row2 = Math.Ceiling(filas);
                if (pag_row2 == 0)
                { pag_row2 = 1; }
                #endregion
                int e = 0;
                //FOR por cada página de la tabla principal
                for (int i = 1; i <= pag_row2; i++)
                {
                    int rows = excelResults.Rows.Count + 1;
                    try
                    {

                        #region next pagination
                        if (i != 1)
                        {
                            console.WriteLine("     Siguiente pagina");
                            IWebElement pagination_next = chrome.FindElement(By.ClassName("pagination-next"));
                            var pag_next = pagination_next.FindElement(By.TagName("a"));
                            pag_next.Click();
                        }
                        #endregion
                        IWebElement tableElement = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[4]/table"));
                        IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));

                        //LOOP por cada fila de la tabla principal de cotizaciones programadas
                        foreach (IWebElement element in trCollection.Skip(1))
                        {
                            try
                            {

                                IList<IWebElement> tdCollection;
                                tdCollection = element.FindElements(By.TagName("td"));
                                if (tdCollection.Count > 0)
                                {
                                    singleOrderRecord = tdCollection[1].Text; //Registro único de pedido
                                                                              //verifica si ya el registro esta en la base de datos
                                    if (!CotiAll.AsEnumerable().Any(row => singleOrderRecord == row.Field<String>("singleOrderRecord")))
                                    {
                                        #region get general info
                                        string id = tdCollection[0].Text;
                                        string descripcion = cleanText(tdCollection[2].Text);
                                        string entidad = cleanText(tdCollection[3].Text);
                                        entidad = entidad.Split(new char[] { '/' })[0];
                                        string fecha_publicacion = tdCollection[5].Text;
                                        #endregion
                                        #region go to quote url
                                        IWebElement link = tdCollection[1];
                                        var linkhref = link.FindElement(By.TagName("a"));
                                        string href = linkhref.GetAttribute("href");
                                        console.WriteLine("    Click en el Codigo Unico Pedido: #" + tdCollection[0].Text + " " + singleOrderRecord);

                                        IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                                        js.ExecuteScript("window.open('{0}', '_blank');");
                                        chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab
                                        chrome.Navigate().GoToUrl(href); //ir al link de la cotizacion
                                        System.Threading.Thread.Sleep(2000);
                                        #endregion
                                        #region get quote info                                                                                                                                   
                                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='elementToPDF']/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]"))); }
                                        catch { }

                                        string coti_url = chrome.Url;
                                        string dependencia = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]")).Text;
                                        string unidad_compra = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[2]")).Text;
                                        string fecha_presentacion = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[6]/td[2]")).Text;
                                        string precio = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[9]/td[2]")).Text;

                                        string contactName = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[2]/div[2]/table/tbody/tr[1]/td[2]")).Text;
                                        string contactTelf = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[2]/div[2]/table/tbody/tr[3]/td[2]")).Text;
                                        string contactEmail = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[2]/div[2]/table/tbody/tr[4]/td[2]")).Text;
                                        string forma_entrega = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[4]/div/div/div/table/tbody/tr[1]/td[2]")).Text;
                                        string dias_entrega = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[4]/div/div/div/table/tbody/tr[2]/td[2]")).Text;
                                        #endregion
                                        #region verifica si es de interes o no para GBM
                                        //verifica si la descripción de la cotizacion y/o descripcion del producto contiene alguna de las palabras claves
                                        string interes_gbm = "NO"; //busca si la descripción de la cotización contiene alguna de las frases o palabras clave de la lista anterior
                                        interes_gbm = keyMatch(descripcion, words);
                                        #endregion
                                        #region buscar el AM de la entidad
                                        string AM = "";
                                        System.Data.DataRow[] entityInfo = entitiesInfo.Select($"entities ='{entidad}'");
                                        if (entityInfo.Length != 0)
                                        {
                                            AM = entityInfo[0]["salesRepCoti"].ToString();
                                        }
                                        else
                                        {
                                            AM = "MAMBURG";
                                        }
                                        #endregion
                                        #region Buscar si es Interes o No por medio de la descripción del producto
                                        //primero se debe indicar si es de interes o no, aunque es un doble loop
                                        //solo si por la descripción general se concluyo que no es de interes
                                        if (interes_gbm == "NO")
                                        {
                                            try
                                            {

                                                string pd = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr/td[6]")).Text;
                                                int c = 1;
                                                while (!string.IsNullOrEmpty(pd))
                                                {
                                                    try
                                                    {
                                                        pd = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + c + "]/td[6]")).Text;
                                                        if (interes_gbm == "NO")
                                                        {
                                                            //en caso de que siga como NO busca en el siguiente producto description
                                                            interes_gbm = keyMatch(pd, words);
                                                        }
                                                        else
                                                        {
                                                            break;
                                                        }
                                                    }
                                                    catch (Exception)
                                                    { pd = ""; }
                                                    c++;
                                                }
                                            }
                                            catch (Exception)
                                            {
                                                //NO TIENE ITEMS
                                            }

                                        }
                                        #endregion
                                        #region Descarga el adjunto solo si es de interes
                                        string attachment = "";
                                        if (interes_gbm == "SI")
                                        {
                                            #region descargar adjunto
                                            try
                                            {
                                                string hayad = "";
                                                try
                                                {
                                                    //trata de tomar el texto cuando no hay documentos adjuntos
                                                    hayad = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[6]/div[2]/table/tbody/tr/td/center")).Text;

                                                }
                                                catch (Exception)
                                                {
                                                }
                                                if (hayad != "No hay documentos adjuntos")
                                                {
                                                    //descarga el archivo
                                                    chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[6]/div[2]/table/tbody/tr/td[4]/a")).Click();
                                                    //no tengo la manera de saber el nombre del archivo
                                                    System.Threading.Thread.Sleep(5000);
                                                    var directory = new DirectoryInfo(root.FilesDownloadPath);
                                                    var myFile = directory.GetFiles()
                                                                 .OrderByDescending(f => f.LastWriteTime)
                                                                 .First();
                                                    attachment = myFile.Name.ToString();
                                                    string fullAttach = root.FilesDownloadPath + "\\" + myFile.Name.ToString();
                                                    adjunto[cont_adj] = fullAttach;
                                                    cont_adj++;
                                                    Array.Resize(ref adjunto, adjunto.Length + 1);

                                                    //agregar la ruta y nombre del archivo como parte de los adjuntos del array del AM
                                                    adj_names[AM + "_" + cont] = fullAttach;
                                                    cont++;

                                                    #region subir FTP
                                                    if (File.Exists(fullAttach))
                                                    {
                                                        string user = "";
                                                        if (SSMandante == "QAS")
                                                        {
                                                            user = cred.QA_SS_APP_SERVER_USER;
                                                        }
                                                        else if (SSMandante == "PRD")
                                                        {
                                                            user = cred.PRD_SS_APP_SERVER_USER;
                                                        }
                                                        bool subir_files = wb2.uploadSftp(fullAttach, $"/home/{user}/projects/smartsimple/gbm-hub-api/src/assets/files/PanamaBids/QuickQuotes", $"Request #{singleOrderRecord}");
                                                        //llenar tabla
                                                        lpsql.insertFile(fullAttach.Split(new string[] { @"downloads\" }, StringSplitOptions.None)[1], singleOrderRecord, "quickQuotes");

                                                    }
                                                    #endregion

                                                }
                                                else
                                                {
                                                    attachment = "No hay documentos adjuntos";
                                                }
                                            }
                                            catch (Exception)
                                            {

                                            }
                                            #endregion
                                        }
                                        else
                                        {
                                            attachment = "No hay documentos adjuntos";
                                        }
                                        #endregion
                                        if (interes_gbm == "SI")
                                        {
                                            #region Extrae la información de producto y agrega info a excel

                                            IWebElement pedidoDetalle = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table"));
                                            int contador_subtotal = pedidoDetalle.FindElements(By.TagName("tr")).Count;

                                            List<productsQuickQuote> PoproductInfo = new List<productsQuickQuote>();
                                            for (int x = 2; x < contador_subtotal; x++)
                                            {
                                                #region get product info
                                                productsQuickQuote pInfo = new productsQuickQuote();
                                                string prod_descrip = cleanText(chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + x + "]/td[6]")).Text);
                                                string prod_clasi = cleanText(chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + x + "]/td[3]")).Text);
                                                string prod_cant = cleanText(chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + x + "]/td[4]")).Text);
                                                string prod_umed = cleanText(chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + x + "]/td[5]")).Text);

                                                pInfo.singleOrderRecord = singleOrderRecord;
                                                pInfo.productService = prod_descrip;
                                                pInfo.ammount = prod_cant;
                                                pInfo.unit = prod_umed;
                                                pInfo.clasification = prod_clasi;
                                                pInfo.active = "1";
                                                pInfo.createdBy = "databot";
                                                PoproductInfo.Add(pInfo);
                                                #endregion
                                                #region Agregar coti al excel principal si es de interes de GBM

                                                #region aregar la cotización al archivo general 
                                                DataRow rRow = excelResults.Rows.Add();
                                                rRow["Entidad"] = entidad;
                                                rRow["Dependencia"] = dependencia;
                                                rRow["Número de cotización"] = singleOrderRecord;
                                                rRow["Descripción de la Solicitud"] = descripcion;
                                                rRow["Fecha y Hora Presentación de Cotizaciones:"] = fecha_presentacion;
                                                rRow["Precio Estimado:"] = precio;
                                                rRow["Nombre del Contacto"] = contactName;
                                                rRow["Telefono de contacto"] = contactTelf;
                                                rRow["Correo de contacto"] = contactEmail;
                                                rRow["Forma de Entrega"] = forma_entrega;
                                                rRow["Días de Entrega"] = dias_entrega;
                                                rRow["Descripcion del Bien, Servicio u Obra"] = prod_descrip;
                                                rRow["Cantidad del Bien, Servicio u Obra"] = prod_cant;
                                                rRow["Unidad Medida del Bien, Servicio u Obra"] = prod_umed;
                                                rRow["Link a Cotizacion en línea"] = coti_url;
                                                rRow["Documentos adjuntos (si tiene)"] = attachment;
                                                excelResults.AcceptChanges();

                                                if (excelsAM.ContainsKey(AM))
                                                {
                                                    //si ya existe un Datatable (futuro excel) en el diccionario de AM
                                                    //se agrega una nueva fila a la tabla
                                                    DataTable excelAm = excelsAM[AM];
                                                    DataRow rRowAm = excelAm.Rows.Add(rRow.ItemArray);
                                                    //rRowAm = rRow;
                                                    excelAm.AcceptChanges();
                                                    //actualiza el diccionario
                                                    excelsAM[AM] = excelAm;
                                                }
                                                else
                                                {
                                                    //no existe por lo que se crea de cero y se agrega al diccionario
                                                    DataTable excelAm = new DataTable();
                                                    excelAm = excelResults.Clone();
                                                    DataRow rRowAm = excelAm.Rows.Add(rRow.ItemArray);
                                                    //rRowAm = rRow;
                                                    excelAm.AcceptChanges();
                                                    excelsAM[AM] = excelAm;
                                                }
                                                //string file = "Cotizaciones Rapidas - " + AM + ".xlsx";
                                                #endregion

                                                #endregion
                                            }
                                            #endregion
                                            #region agrega info a BD

                                            Dictionary<string, string> quoteFields = new Dictionary<string, string>
                                            {
                                                ["description"] = descripcion,
                                                ["quoteLink"] = coti_url,
                                                ["entity"] = entidad,
                                                ["singleOrderRecord"] = singleOrderRecord,
                                                ["purchaseUnit"] = unidad_compra,
                                                ["dependence"] = dependencia,
                                                ["quoteDate"] = fecha_presentacion,
                                                ["estimatePrice"] = precio,
                                                ["contactName"] = contactName,
                                                ["contactPhone"] = contactTelf,
                                                ["mailContact"] = contactEmail,
                                                ["deliveryMethod"] = (forma_entrega == "Total") ? "1" : (forma_entrega == "Parcial") ? "2" : "3",
                                                ["deliveryDays"] = dias_entrega,
                                                ["interestGBM"] = interes_gbm,
                                                ["attachments"] = attachment,
                                                ["active"] = "1",
                                                ["createdBy"] = "databot"
                                            };

                                            console.WriteLine("  Agregando información a la base de datos");
                                            bool add_sql = lpsql.insertInfoPurchaseOrder(quoteFields, "quickQuoteReport");
                                            log.LogDeCambios("Creacion", root.BDProcess, "Cotizaciones Rápidas Panama", "Agregar Cotización Rapida Convenio", singleOrderRecord, add_sql.ToString());
                                            respFinal = respFinal + "\\n" + "Agregar Cotización Rápida Convenio: " + singleOrderRecord + " " + add_sql.ToString();

                                            if (add_sql == false)
                                            {
                                                //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                                //para enviarlo por email y agregarla
                                                resp_sql = resp_sql + singleOrderRecord + "<br>";
                                                resp_add_sql = false;
                                            }
                                            else
                                            {
                                                //se insertan los productos
                                                bool addProducts = lpsql.insertInfoProductsQuotes(PoproductInfo);
                                                if (!addProducts)
                                                {
                                                    //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                                    //para enviarlo por email y agregarla
                                                    resp_sql = resp_sql + coti_url + "<br>";
                                                    resp_add_sql = false;
                                                }
                                            }

                                            #endregion
                                            //si es de interes y si el AM pertenece a estos 2 usuarios se debe crear el mensaje para enviar por wteams

                                            if (AM == "MAMBURG")
                                            {
                                                respMAMBURG = respMAMBURG + "- **" + entidad + "**: [" + singleOrderRecord + "](" + coti_url + ")" + " - " + descripcion + "\n";
                                            }
                                            else if (AM == "ERUIZ")
                                            {
                                                respERUIZ = respERUIZ + "- **" + entidad + "**: [" + singleOrderRecord + "](" + coti_url + ")" + " - " + descripcion + "\n";
                                            }
                                        }

                                        #region finaliza todo en la nueva hoja por lo que cierra 
                                        try
                                        {
                                            System.Threading.Thread.Sleep(1000);
                                            chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                                            chrome.Close();
                                            System.Threading.Thread.Sleep(1000);

                                        }
                                        catch (Exception)
                                        {
                                        }
                                        #endregion
                                    }

                                }

                                chrome.SwitchTo().Window(chrome.WindowHandles[0]); //regreso a la pagina principal y sigo con la siguiente fila
                            }
                            catch (Exception EX)
                            {
                                resp_sql = resp_sql + "Error al descargar cotización: " + rows + "<br>";
                                resp_add_sql = false;
                                console.WriteLine(EX.Message);
                            }
                        }

                    }
                    catch (Exception EX)
                    {
                        resp_sql = resp_sql + "Error en la página: " + i + "<br>";
                        resp_add_sql = false;
                        console.WriteLine(EX.Message);
                    }
                }

            }
            #endregion
            console.WriteLine("  Guardar el reporte");
            #region guardar reporte general y enviar a usuarios principales

            chrome.Close();
            proc.KillProcess("chromedriver", true);
            string fecha_file = DateTime.Now.ToString("dd_MM_yyyy");
            string fileNameGeneral = root.FilesDownloadPath + "\\" + "Cotizaciones Rápidas - Gobierno del Panamá - " + fecha_file + ".xlsx";
            MsExcel.CreateExcel(excelResults, "Cotizaciones Programadas", fileNameGeneral);

            fecha_file = DateTime.Now.ToString("dd/MM/yyyy");
            console.WriteLine("  Enviando Reporte");
            string mes_text = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")));
            JArray j_copias = JArray.Parse(lpsql.getEmail("LICPA"));
            string html = Properties.Resources.emailtemplate1;
            html = html.Replace("{subject}", "Cotizaciones Rápidas Programadas");
            html = html.Replace("{cuerpo}", "Se adjunta el reporte de Cotizaciones Rápidas del Gobierno de Panamá programadas al día de mañana");
            html = html.Replace("{contenido}", "");
            string sub = "Cotizaciones rápidas programadas del Gobierno de Panamá - Fecha: " + fecha_file;
            //string[] Users = j_copias.ToObject<string[]>();
            string jmail1 = j_copias[0]["email"].ToString();
            string jmail2 = j_copias[1]["email"].ToString();
            string[] Users = new string[] { jmail1, jmail2 };
            mail.SendHTMLMail(html, Users, sub, null, new string[] { fileNameGeneral });

            #endregion
            #region enviar reportes a AM y chats 

            string header = "Se le notifica que se han publicado las siguientes **cotizaciones rápidas**:\r\n ";

            foreach (KeyValuePair<string, DataTable> item in excelsAM)
            {
                string am = item.Key;
                DataTable excelAm = item.Value;
                string Nfile = "Cotizaciones Rapidas - " + am + ".xlsx";
                string filePath = root.FilesDownloadPath + "\\" + Nfile;
                string subject = "Cotizaciones rápidas programadas del Gobierno de Panamá – Cuentas " + am + " - Fecha: " + fecha_file;
                string body = "Se adjunta el reporte de Cotizaciones Rápidas del Gobierno de Panamá en sus cuentas asignadas programadas al día de mañana";
                MsExcel.CreateExcel(excelAm, "Cotizaciones Programadas", filePath);
                html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Cotizaciones Rápidas Programadas");
                html = html.Replace("{cuerpo}", body);
                html = html.Replace("{contenido}", "");

                int x = 0;
                string[] adj = new string[1];
                foreach (KeyValuePair<string, string> pair in adj_names)
                {
                    //si la llave contiene el id del AM
                    if (pair.Key.ToString().Contains(am))
                    {
                        //se adjunta el archivo al array para el email
                        string archivo = pair.Value.ToString();
                        if (File.Exists(archivo))
                        {
                            adj[x] = archivo;
                            x++;
                            Array.Resize(ref adj, adj.Length + 1);
                        }

                    }
                }
                adj[adj.Length - 1] = filePath;
                mail.SendHTMLMail(html, new string[] { am + "@gbm.net" }, subject, null, adj);

                //Enviar notificaciones solamente a estos 2 usuarios
                if (am == "MAMBURG" || am == "ERUIZ")
                {
                    string mensaje = "";
                    mensaje = (am == "MAMBURG") ? respMAMBURG : (am == "ERUIZ") ? respERUIZ : "";
                    wt.SendNotification(am + "@gbm.net", "Nuevas cotizaciones rápidas LCPA", header + mensaje.ToString());

                }
            }
            #endregion
            #region enviar correos de error
            if (validar_lineas == false)
            {
                string[] cc = { "appmanagement@gbm.net" };
                string[] adj = { fileNameGeneral };
                mail.SendHTMLMail("A continuacion se adjunta el archivo con el conglomerado de las nuevas cotizaciones rapidas de GBM de Panama", new string[] { "kvanegas@gbm.net" }, "Reporte diario de Cotizaciones Rapidas - " + fecha_file, cc, adj);

            }
            if (resp_add_sql == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                string[] adj = { fileNameGeneral };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Cotizaciones a la base de datos: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Reporte diario de Cotizaciones Rapidas - " + fecha_file, cc, adj);

            }
            #endregion

            root.BDUserCreatedBy = "MAMBURG@gbm.net";
            root.requestDetails = respFinal;

            return "";
        }
        public string keyMatch(string texto, List<string> words)
        {
            string interes_gbm = "NO";
            try
            {
                texto = texto.ToLower();
                texto = texto.Replace("á", "a"); texto = texto.Replace("é", "e"); texto = texto.Replace("í", "i"); texto = texto.Replace("ó", "o"); texto = texto.Replace("ú", "u");
                texto = val.RemoveSpecialChars(texto, 1);

                var result = words.Where(x => texto.Contains(x)).ToList();
                if (result.Count > 0)
                { interes_gbm = "SI"; }
            }
            catch (Exception)
            {
                interes_gbm = "SI";
            }

            return interes_gbm;
        }
        public string cleanText(string text)
        {
            text = text.Replace("\"", "");
            text = text.Replace("'", "");
            return text;
        }

    }
    public class productsQuickQuote
    {
        public string singleOrderRecord { get; set; }
        public string productService { get; set; }
        public string clasification { get; set; }
        public string ammount { get; set; }
        public string unit { get; set; }
        public string active { get; set; }
        public string createdBy { get; set; }
    }
}
