using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using System.IO;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Globalization;
using OpenQA.Selenium.Interactions;
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

namespace DataBotV5.Automation.MASS.PanamaBids
{
    /// <summary>
    /// Clase RPA Automation encargada de extraer la información de los registros de GBPA que se encuentran publicados en el convenio macro del Gobierno de Panamá (una vez al día).
    /// </summary>
    class PurchaseOrderMarco
    {
        #region variables_globales
        Stats estadisticas = new Stats();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        ProcessAdmin padmin = new ProcessAdmin();
        WebInteraction sel = new WebInteraction();
        PanamaBidsLogical paBids = new PanamaBidsLogical();
        Log log = new Log();

        string respuesta = "";

        string respFinal = "";
        #endregion
        public void Main()
        {

            console.WriteLine(" Procesando...");
            respuesta = getPurchaseOrderPanamaGBM(); //Metodo en la capa de logical para extraer la información y actualizar la base de datos
            console.WriteLine(" Creando Estadisticas");
            if (respuesta == "true")
            {
                root.requestDetails = "No hay registros";
            }
            else
            {

                console.WriteLine("Si hay registros");


                //Se deja en el main porque se ejecuta con planificador.
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }

            }



        }
        private string getPurchaseOrderPanamaGBM()
        {
            #region variables privadas
            double filas = 0;
            string[] adjunto = new string[1],
                CopyCC = new string[1];
            string cant_filas = "",
                main_web_page = "",
                vendor_text = "",
                resp_sql = "",
                singleOrderRecord = "",
                sector = "";
            bool resp_add_sql = true,
                validar_lineas = true,
                cisco_add = false,
                no_registros = false;
            int cont_adj = 0,
                cont_am = 0;
            DateTime file_date = DateTime.MinValue, file_date_before = DateTime.MinValue;
            Dictionary<string, string> adj_names = new Dictionary<string, string>(),
                AMs = new Dictionary<string, string>();
            List<string> newEntities = new List<string>();
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
            #region excel
            console.WriteLine("  Creando Excel");
            DataTable excelResult = new DataTable();


            DataTable columnsPoMacro = lpsql.columnsPoMacro();
            #region titulos_excel
            string[] columns = {
                "Registro Unico de Pedido",
                "Sector",
                "Convenio",
                "Entidad",
                "Producto/Servicio",
                "Cantidad de Producto",
                "Marca del Producto",
                "Total del Producto",
                "Sub Total de la orden",
                "Orden de Compra",
                "Fecha de Registro",
                "Fecha de Publicacion",
                "Fianza por cumplimiento",
                "Oportunidad",
                "Quote",
                "Tipo de Pedido",
                "Sales Order",
                "Estado de GBM",
                "Estado de Orden",
                "Dias de Entrega",
                "Fecha Maxima de Entrega",
                "Dias Faltantes",
                "Forecast",
                "Provincia",
                "Lugar de Entrega",
                "Contacto de la Empresa",
                "Telefono del Contacto",
                "Email del Contacto",
                "Confirmación de Orden",
                "Fecha Real de Entrega",
                "Monto de Multa",
                "Unidad Solicitante",
                "Contacto Cuenta",
                "Email Cuenta",
                "Tipo de Forecast",
                "Vendor Order",
                "Nombre del adjunto",
                "Account Manager Asignado",
                "Link al documento",
                "Comentarios"
            };
            foreach (string item in columns)
            {
                excelResult.Columns.Add(item);
            }
            DataTable excelCisco = excelResult.Copy();
            #endregion
            #endregion
            #region Ingreso al website
            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#/busqueda-tienda-virtual");
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#/busqueda-tienda-virtual"); }
            IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
            #endregion
            #region Buscar convenio y proveedor
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='IdConvenio']"))); }
            catch { }
            chrome.Manage().Cookies.DeleteAllCookies();

            //eliminar pop del gobierno
            // 
            try
            {
                chrome.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/ngbd-modal-content/button")).Click();
            }
            catch (Exception)
            {
            }


            //eliminar header para que no estorbe
            string deleteElement = @"
           var l = document.getElementsByClassName('flex-none position-fixed w-100')[0];
           l.parentNode.removeChild(l);
        ";
            jsup.ExecuteScript(deleteElement);

            console.WriteLine("Seleccionar Convenio");
            SelectElement convenio_select = new SelectElement(chrome.FindElement(By.XPath("//*[@id='IdConvenio']"))); //*[@id="IdConvenio"]
            System.Threading.Thread.Sleep(1000);
            convenio_select.SelectByValue("134");  //EQUIPOS INFORMÁTICOS Y TELECOMUNICACIONES 
            System.Threading.Thread.Sleep(3000);
            string convenio_name = convenio_select.SelectedOption.Text;

            SelectElement vendor_select2 = new SelectElement(chrome.FindElement(By.XPath("//*[@id='IdProveedor']"))); //*[@id="IdProveedor"]
            vendor_select2.SelectByValue("1979769");
            vendor_text = vendor_select2.SelectedOption.Text;
            System.Threading.Thread.Sleep(1000);
            #endregion
            #region calendario
            System.Threading.Thread.Sleep(250);
            chrome.FindElement(By.XPath("//*[@id='fd']")).Clear();
            chrome.FindElement(By.XPath("//*[@id='fd']")).SendKeys(DateTime.Now.ToString("dd-MM-yyyy"));
            chrome.FindElement(By.XPath("//*[@id='fh']")).Clear();
            chrome.FindElement(By.XPath("//*[@id='fh']")).SendKeys(DateTime.Now.ToString("dd-MM-yyyy"));
            //*[@id="fd"]
            //try
            //{ chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/span/button")).Click(); } //fecha desde 
            //catch (Exception ex)
            //{
            //    console.WriteLine("Error en click calendario" + ex.Message);
            //    chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/span/button")).Click();
            //}
            //System.Threading.Thread.Sleep(250);
            //try
            //{
            //    //TODAY
            //    chrome.FindElement(By.XPath("//*[@id='imageLazyContainer']/div[3]/ul/li[2]/span/button[1]")).Click();
            //}
            //catch (Exception ex)
            //{
            //    console.WriteLine("Error en click today" + ex.Message);
            //    System.Threading.Thread.Sleep(2000);
            //    chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/div/ul/li[2]/span/button[1]")).Click();
            //}

            file_date = DateTime.Today;
            string fecha_hasta = "'" + file_date.ToString("dd-MM-yyyy");
            string fecha_desde = fecha_hasta;
            #endregion
            #region saca la lista de proveedores
            //IWebElement vendor_list = chrome.FindElement(By.XPath("//*[@id='IdProveedor']"));
            //System.Threading.Thread.Sleep(1000);
            //SelectElement selectList = new SelectElement(vendor_list);
            //IList<IWebElement> voptions = selectList.Options;
            //console.WriteLine("  Cantidad de proveedores en lista: " + voptions.Count);
            #endregion
            #region Extraer Ordenes de Compra de GBM

            #region buscar y tomar la cantidad de filas
            try
            {
                IWebElement buscar = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button"));
                new Actions(chrome).MoveToElement(buscar).Perform();
                chrome.Manage().Window.Minimize();
                chrome.Manage().Window.Maximize();
            }
            catch (Exception)
            {
                try
                {
                    chrome.Manage().Window.Minimize();
                    chrome.Manage().Window.Maximize();
                }
                catch (Exception)
                {

                }
            }
            try
            {
                chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button")).Click(); //BUSCAR
            }
            catch (Exception)
            {
                chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button")).Click(); //BUSCAR

            }
            console.WriteLine("   Buscar");
            System.Threading.Thread.Sleep(3000);
            main_web_page = chrome.Url;
            string noResults = "";
            try
            {

                noResults = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div/div/table/tbody/tr/td")).Text;
            }
            catch (Exception)
            {

            }


            #endregion
            if (noResults != "No hay registro")
            {
                cant_filas = "";
                try
                { cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text; }
                catch (Exception)
                { cant_filas = ""; }

                console.WriteLine("   " + cant_filas);

                while (cant_filas == "Se encontraron Pedidos Publicados")
                {
                    System.Threading.Thread.Sleep(1000);
                    try
                    { cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text; }
                    catch (Exception)
                    { cant_filas = ""; }
                }

                int mas = 0;
                #region Buscar si son más de 50 resultados
                //La pagina no tira la cantidad de ordenes en total cuando son mas de 50 por lo que tira +50
                //lo que se debe de hacer es dar click hasta el final de la tabla para asi mostrar la cantidad real de ordenes
                if (cant_filas.Contains("+"))
                {
                    IWebElement pagination_last = chrome.FindElement(By.CssSelector("[aria-label=Last]"));
                    try
                    {

                        new Actions(chrome).MoveToElement(pagination_last).Perform();
                        chrome.Manage().Window.Minimize();
                        chrome.Manage().Window.Maximize();
                    }
                    catch (Exception)
                    {
                        try
                        {

                            new Actions(chrome).MoveToElement(pagination_last).Perform();
                            chrome.Manage().Window.Minimize();
                            chrome.Manage().Window.Maximize();
                        }
                        catch (Exception)
                        {

                        }
                    }
                    do
                    {

                        try
                        {
                            pagination_last.Click();

                        }
                        catch (Exception)
                        {
                            pagination_last.Click();
                        }
                        System.Threading.Thread.Sleep(1500);
                        cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text;
                        mas++;
                    } while (cant_filas.Contains("+"));

                    if (mas > 0)
                    {
                        IWebElement pagination_first = chrome.FindElement(By.CssSelector("[aria-label=First]"));
                        try
                        {
                            new Actions(chrome).MoveToElement(pagination_first).Perform();
                            chrome.Manage().Window.Minimize();
                            chrome.Manage().Window.Maximize();
                        }
                        catch (Exception)
                        {
                            try
                            {
                                new Actions(chrome).MoveToElement(pagination_first).Perform();
                                chrome.Manage().Window.Minimize();
                                chrome.Manage().Window.Maximize();
                            }
                            catch (Exception)
                            {

                            }
                        }
                        //var pag_first = pagination_first.FindElement(By.TagName("a"));
                        try
                        {
                            pagination_first.Click();
                        }
                        catch (Exception)
                        {
                            pagination_first.Click();
                        }

                        System.Threading.Thread.Sleep(1000);
                        cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text;
                    }
                }
                #endregion
                #region Calcular la cantidad 
                cant_filas = cant_filas.Split(new string[] { ": " }, StringSplitOptions.None)[1];
                double num1 = double.Parse(cant_filas);
                double num2 = 10;
                filas = (num1 / num2);
                double pag_row2 = Math.Ceiling(filas);
                if (pag_row2 == 0)
                { pag_row2 = 1; }
                #endregion
                //por cantidad de paginas de la tabla principal
                for (int i = 1; i <= pag_row2; i++)
                {
                    #region Clickear la siguiente pagina cuando termina
                    //next en pagination

                    if (i != 1)
                    {
                        System.Threading.Thread.Sleep(1000);
                        console.WriteLine("Siguiente pagina");
                        IWebElement pagination_next = chrome.FindElement(By.CssSelector("[aria-label=Next]"));
                        try
                        {
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 40)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector("[aria-label=Next]"))); }
                            catch { }
                            new Actions(chrome).MoveToElement(pagination_next).Perform();
                            chrome.Manage().Window.Minimize();
                            chrome.Manage().Window.Maximize();
                            //System.Threading.Thread.Sleep(1000);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                new Actions(chrome).MoveToElement(pagination_next).Perform();
                                chrome.Manage().Window.Minimize();
                                chrome.Manage().Window.Maximize();

                            }
                            catch (Exception ex)
                            {
                                console.WriteLine(ex.ToString());
                            }
                        }
                        System.Threading.Thread.Sleep(2000);
                        try
                        {
                            pagination_next.Click();
                        }
                        catch (Exception)
                        {
                            pagination_next.Click();
                        }
                    }
                    #endregion

                    //toma la cantidad de filas que tiene en la tabla en la pagina
                    IWebElement tableElement = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[1]/div[2]/table/tbody"));
                    IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));
                    int e = 0;
                    int w = 1;
                    //por cada fila de la tabla
                    foreach (IWebElement element in trCollection)
                    {
                        List<PoProductMacro> PoproductInfo = new List<PoProductMacro>();
                        List<calculateData> datosCalculados = new List<calculateData>();
                        try
                        {
                            //extraer las columnas de la fila
                            PoproductInfo.Clear();
                            datosCalculados.Clear();
                            chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                            IList<IWebElement> tdCollection;
                            tdCollection = element.FindElements(By.TagName("td"));
                            if (tdCollection.Count > 0)
                            {
                                #region extraer información general
                                string id = tdCollection[0].Text;
                                string entidad = tdCollection[0].Text;
                                string descripcion = tdCollection[2].Text;
                                string fecha = tdCollection[3].Text;
                                singleOrderRecord = tdCollection[1].Text; //Registro único de pedido

                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Extraer PO", $"Id: {id}, Entidad: {entidad}, Descripción: {descripcion}, Fecha: {fecha}", root.Subject);
                                respFinal = respFinal + "\\n" + "Extraer PO " + $"Id: {id}, Entidad: {entidad}, Descripción: {descripcion}, Fecha: {fecha}";


                                console.WriteLine("    Click en el Codigo Unico Pedido: " + singleOrderRecord);
                                #endregion
                                #region Ir la nueva pestaña de la orden

                                IWebElement link = tdCollection[1];
                                var linkhref = link.FindElement(By.TagName("a"));

                                try
                                {

                                    if (w == 1)
                                    {
                                        try
                                        {
                                            //despues de darle siguiene pagina, es decir la primera fila se va mas arriba al convenio
                                            new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='IdConvenio']")));
                                            chrome.Manage().Window.Minimize();
                                            chrome.Manage().Window.Maximize();

                                        }
                                        catch (Exception)
                                        {

                                        }
                                    }
                                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 40)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath($"/html/body/app-root/main/ng-component/div/div/div/div[1]/div[2]/table/tbody/tr[{w}]/td[2]/a"))); }
                                    catch { }
                                    w++;
                                    new Actions(chrome).MoveToElement(link).Perform();
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        new Actions(chrome).MoveToElement(link).Perform();
                                        chrome.Manage().Window.Minimize();
                                        chrome.Manage().Window.Maximize();
                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                                try
                                { linkhref.Click(); }
                                catch (Exception)
                                {
                                    linkhref.Click();
                                }

                                System.Threading.Thread.Sleep(2000);

                                chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab
                                #endregion
                                //extrae la información de la PO, inserta los datos en las DB y devuelve el excel lleno
                                string cantidad_opp = "";
                                PoInfoMacro poInfo = paBids.GetPoInfo(chrome, excelResult, excelCisco, convenio_name, entidad, adjunto, AMs, singleOrderRecord, cont_adj, cantidad_opp, newEntities);
                                excelResult = poInfo.excel.Copy();
                                excelResult.AcceptChanges();
                                adjunto = poInfo.adjunto;
                                AMs = poInfo.AMs;
                                cont_adj = poInfo.contAdj;
                                newEntities = poInfo.newEntities;

                                if (poInfo.isCisco)
                                {
                                    cisco_add = true;
                                    excelCisco = poInfo.excelCisco.Copy();
                                    excelCisco.AcceptChanges();
                                }
                                #region Cerrar pestaña de chrome y seguir con la siguiente orden de la lista
                                System.Threading.Thread.Sleep(1000);
                                chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                                chrome.Close();
                                System.Threading.Thread.Sleep(1000);
                                chrome.SwitchTo().Window(chrome.WindowHandles[0]); //regreso a la pagina principal y sigo con la siguiente fila
                                #endregion

                            } //la fila tiene td
                        }
                        catch (Exception ex)
                        {
                            //console.WriteLine(ex.ToString());
                            DataRow rRow = excelResult.Rows.Add();
                            rRow["Registro Unico de Pedido"] = singleOrderRecord;
                            rRow["Sector"] = sector;
                            rRow["Convenio"] = convenio_name;
                            rRow["Entidad"] = ex;
                            excelResult.AcceptChanges();
                        }
                    } //foreach fila en la tabla main

                } //for cantidad de paginas de la tabla main

            }
            else
            {
                //no hay registros
                DataRow rRow = excelResult.Rows.Add();
                rRow["Registro Unico de Pedido"] = singleOrderRecord;
                rRow["Sector"] = sector;
                rRow["Convenio"] = convenio_name;
                rRow["Entidad"] = "No hay registros el día de hoy";
                excelResult.AcceptChanges();
            }




            #endregion
            #region cerrar chrome
            console.WriteLine("  Guardar el reporte");
            chrome.Close();
            proc.KillProcess("chromedriver", true);
            #endregion
            #region Crear Excel
            console.WriteLine("Save Excel...");
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add(excelResult, "Ordenes Convenio");
            ws.Columns().AdjustToContents();
            ws.Range($"K2:L{excelResult.Rows.Count + 1}").Style.NumberFormat.Format = "dd/MM/yyyy";
            ws.Range($"U2:U{excelResult.Rows.Count + 1}").Style.NumberFormat.Format = "dd/MM/yyyy";

            for (int i = 2; i <= excelResult.Rows.Count + 1; i++)
            {
                try
                {
                    ws.Cell(i, 39).Hyperlink = new XLHyperlink(ws.Cell(i, 39).Value.ToString());
                }
                catch (Exception)
                {

                }
            }



            string ruta = root.FilesDownloadPath + $"\\Reporte de Convenio GBM  {DateTime.Now.ToString("dd_MM_yyyy")}.xlsx";
            if (File.Exists(ruta))
            {
                File.Delete(ruta);
            }
            wb.SaveAs(ruta);

            //cisco
            string rutaCisco = "";
            if (cisco_add)
            {

                console.WriteLine("Save Excel...");
                XLWorkbook wbCisco = new XLWorkbook();
                IXLWorksheet wsCisco = wbCisco.Worksheets.Add(excelCisco, "Ordenes Convenio");
                wsCisco.Columns().AdjustToContents();
                wsCisco.Range($"K2:L{excelResult.Rows.Count + 1}").Style.NumberFormat.Format = "dd/MM/yyyy";
                wsCisco.Range($"U2:U{excelResult.Rows.Count + 1}").Style.NumberFormat.Format = "dd/MM/yyyy";
                for (int i = 2; i <= excelResult.Rows.Count + 1; i++)
                {
                    try
                    {
                        wsCisco.Cell(i, 39).Hyperlink = new XLHyperlink(wsCisco.Cell(i, 39).Value.ToString());
                    }
                    catch (Exception)
                    {

                    }
                }

                rutaCisco = root.FilesDownloadPath + $"\\Reporte de Convenio GBM Networking {DateTime.Now.ToString("dd_MM_yyyy")}.xlsx";
                if (File.Exists(rutaCisco))
                {
                    File.Delete(rutaCisco);
                }
                wbCisco.SaveAs(rutaCisco);

            }
            #endregion

            #region entidades sin crear
            if (newEntities.Count > 0)
            {
                newEntities = newEntities.Distinct().ToList();
                string cuerpo = "<table class='myCustomTable' width='100 %'>";
                cuerpo += "<thead><tr><th>Entidad</th></tr></thead>";
                cuerpo += "<tcuerpo>";
                foreach (string item in newEntities)
                {
                    //crear tabla html
                    cuerpo += "<tr><td>";
                    cuerpo += item;
                    cuerpo += "</td></tr>";
                }
                cuerpo += "</tcuerpo>";
                cuerpo += "</table>";
                //sendemail
                string email_gen = Properties.Resources.emailtemplate1;
                email_gen = email_gen.Replace("{subject}", "Nuevas entidades sin crear - Convenio Marco");
                email_gen = email_gen.Replace("{cuerpo}", "A Continuación se detallan los emails de cuenta, unidad solicitantes y/o entidades sin crear.");
                email_gen = email_gen.Replace("{contenido}", cuerpo);

                mail.SendHTMLMail(email_gen, new string[] { "kvanegas@gbm.net" }, "Nuevas entidades sin crear - Convenio Marco", new string[] { "dmeza@gbm.net" }, null);
            }
            #endregion

            #region Enviar Reporte

            console.WriteLine("  Enviando Reporte");
            string mes_text = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(DateTime.Now.AddMonths(-1).ToString("MMMM", CultureInfo.CreateSpecificCulture("es")));
            if (!no_registros)
            {
                string[] cc = new string[1];
                cc[0] = (validar_lineas == false) ? "appmanagement@gbm.net" : "";
                string msj = $"A continuacion se adjunta el archivo con el conglomerado de las nuevas ordenes publicadas en Convenio Marco a favor de GBM de Panama";
                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Reporte diario de ordenes nuevas publicadas por Convenio Marco");
                html = html.Replace("{cuerpo}", msj);
                html = html.Replace("{contenido}", "");
                console.WriteLine("Send Email...");
                adjunto[adjunto.Length - 1] = ruta;
                mail.SendHTMLMail(html, new string[] { "kvanegas@gbm.net" }, "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + DateTime.Now.ToString("dd/MM/yyyy"), cc, adjunto);
                root.BDUserCreatedBy = "KVANEGAS";

            }
            if (resp_add_sql == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + DateTime.Now.ToString("dd/MM/yyyy"), cc, adjunto);
                root.BDUserCreatedBy = "DMEZA";

            }

            #region enviar reporte Cisco
            if (cisco_add)
            {
                string subject = "Reporte diario de ordenes nuevas publicadas por Convenio Marco Networking - " + DateTime.Now.ToString("dd/MM/yyyy");
                string body = "A continuacion se adjunta el archivo con el conglomerado de las nuevas ordenes de Networking publicadas en Convenio Marco a favor de GBM de Panama";
                string[] adj = { rutaCisco };
                foreach (KeyValuePair<string, string> pair in AMs)
                {
                    string email = pair.Value.ToString() + "@gbm.net";
                    CopyCC[cont_am] = email;
                    cont_am++;
                    Array.Resize(ref CopyCC, CopyCC.Length + 1);

                }
                Array.Resize(ref CopyCC, CopyCC.Length - 1);
                string[] senders = { "mmedina@gbm.net", "jleal@gbm.net" };
                root.BDUserCreatedBy = "MMEDINA";

                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", subject);
                html = html.Replace("{cuerpo}", body);
                html = html.Replace("{contenido}", "");
                console.WriteLine("Send Email...");

                mail.SendHTMLMail(html, senders , subject, CopyCC, adj);
            }
            #endregion

            #endregion


            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "KVANEGAS";
            return validar_lineas.ToString();

        }
        public int businessDays(int cantidad)
        {
            int dias = 0;

            if (cantidad >= 1 && cantidad <= 15)
            { dias = 18; }
            else if (cantidad >= 16 && cantidad <= 30)
            { dias = 28; }
            else if (cantidad >= 31 && cantidad <= 50)
            { dias = 38; }
            else if (cantidad >= 51 && cantidad <= 100)
            { dias = 50; }
            else if (cantidad >= 101)
            { dias = 70; }
            else
            { dias = 0; }

            return dias;
        }
        public string get_ext(string href)
        {
            string ext = "";
            href = href.ToLower();
            if (href.Contains(".pdf"))
            { ext = ".pdf"; }
            else if (href.Contains(".jpeg"))
            { ext = ".jpeg"; }
            else if (href.Contains(".jpg"))
            { ext = ".jpg"; }
            else if (href.Contains(".png"))
            { ext = ".png"; }
            else if (href.Contains(".xlsx"))
            { ext = ".xlsx"; }
            else if (href.Contains(".docx"))
            { ext = ".docx"; }
            else if (href.Contains(".bmp"))
            { ext = ".bmp"; }
            else if (href.Contains(".rar"))
            { ext = ".rar"; }
            else
            { ext = ""; }

            return ext;

        }
        public DateTime AddWorkdays(DateTime originalDate, int workDays)
        {
            DateTime tmpDate = originalDate;
            DateTime[] feriados = lpsql.getholidays();
            try
            {
                while (workDays > 0)
                {
                    tmpDate = tmpDate.AddDays(1); //agregar un dia a la fecha
                    DateTime ntmpDate = new DateTime(tmpDate.Year, tmpDate.Month, tmpDate.Day); //para quitarle las horas
                    bool feriado = Array.Exists(feriados, x => x == ntmpDate); //para saber si el ntmpDate es feriado en la lista de feriados
                    if (tmpDate.DayOfWeek < DayOfWeek.Saturday && tmpDate.DayOfWeek > DayOfWeek.Sunday && feriado == false)
                    {
                        workDays--;
                    }
                }
            }
            catch (Exception)
            {

            }

            return tmpDate;
        }


    }
    public class PoProductMacro
    {
        public string singleOrderRecord { get; set; }
        public string productCode { get; set; }
        public string quantity { get; set; }
        public string totalProduct { get; set; }
        public string orderType { get; set; }
        public string active { get; set; }
        public string createdBy { get; set; }
    }
    public class calculateData
    {
        public int deliveryDay { get; set; }
        public double daysRemaining { get; set; }
        public DateTime maximumDeliveryDate { get; set; }
    }
}
