using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Globalization;
using System.Net;
using System.Windows.Forms;
using System.Threading;
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
using OpenQA.Selenium.Interactions;

namespace DataBotV5.Automation.RPA2.PanamaBids

{
    /// <summary>
    /// Clase RPA Automation encargada de extraer la información 
    /// de compras de la competencia del convenio macro del Gobierno de Panamá (una vez al mes).
    /// </summary>
    class PurchaseOrderCompetition
    {
        #region variables_globales
        public static bool completed = false;
        public static WebBrowser wb2;
        PanamaPurchase pa_compra = new PanamaPurchase();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        ProcessAdmin padmin = new ProcessAdmin();
        WebInteraction sel = new WebInteraction();
        Log log = new Log();

        string respuesta = "";


        string respFinal = "";

        #endregion

        public void Main()
        {
            console.WriteLine(" Procesando...");
            respuesta = getPurchaseOrderPanama();

            if (respuesta == "true")
            {
                console.WriteLine(" Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }


        private string getPurchaseOrderPanama()
        {
            try
            {

                #region variables privadas
                string respuesta = "", cant_filas = "", vendor_text = "", resp_sql = "", id_producto = "", main_web_page = "", precio_unitario = "";
                double filas = 0, prod_cant = 0, prod_total = 0;
                bool resp_add_sql = true, validar_lineas = true;
                int pag_row = 0, cont_adj = 0;

                DateTime file_date = DateTime.MinValue;
                DateTime file_date_before = DateTime.MinValue;

                DataTable productTypes = lpsql.productType();
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
                // titulos_excel
                excelResult.Columns.Add("Convenio");
                excelResult.Columns.Add("Proveedor");
                excelResult.Columns.Add("Fecha Desde");
                excelResult.Columns.Add("Fecha Hasta");
                excelResult.Columns.Add("Entidad");
                excelResult.Columns.Add("Registro único de pedido");
                excelResult.Columns.Add("Fecha de Registro");
                excelResult.Columns.Add("Fecha de Publicacion");
                //excelResult.Columns.Add("Empresa");
                excelResult.Columns.Add("Producto/Servicio");
                excelResult.Columns.Add("Cantidad");
                excelResult.Columns.Add("Total del Producto");
                excelResult.Columns.Add("Precio Unitario");
                excelResult.Columns.Add("Sub-Total de la PO");
                excelResult.Columns.Add("Línea de producto");
                excelResult.Columns.Add("Tipo de producto (Lenguaje GBM)");
                excelResult.Columns.Add("GBM participa?");
                excelResult.Columns.Add("Link");
                DataTable xlSheet = new DataTable();
                #endregion
                #region Ingreso al website
                IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath, true);
                try
                {
                    console.WriteLine("  Ingresando al website");
                    chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#/busqueda-tienda-virtual");
                }
                catch (Exception)
                { chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#/busqueda-tienda-virtual"); }

                //js executor para subir al inicio de pagina
                IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;

                #endregion
                #region Buscar las ordenes
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

                try
                {
                    chrome.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/ng-component/button")).Click();
                }
                finally
                {

                }

                //eliminar header para que no estorbe
                string deleteElement = @"
           var l = document.getElementsByClassName('flex-none position-fixed w-100')[0];
           l.parentNode.removeChild(l);
        ";
                jsup.ExecuteScript(deleteElement);

                console.WriteLine("  Seleccionar Convenio");
                //seleccionar una opcion de un dropdown
                IWebElement convenioSelect = chrome.FindElement(By.XPath("//*[@id='IdConvenio']"));
                SelectElement convenio_select = new SelectElement(convenioSelect);

                System.Threading.Thread.Sleep(1000);
                convenio_select.SelectByValue("134"); //EQUIPOS INFORMÁTICOS Y TELECOMUNICACIONES 
                System.Threading.Thread.Sleep(2000);
                string convenio_name = convenio_select.SelectedOption.Text;
                string fecha_desde = chrome.FindElement(By.XPath("//*[@id='fd']")).Text; //*[@id="fd"]
                string fecha_hasta = chrome.FindElement(By.XPath("//*[@id='fh']")).Text;
                if (fecha_hasta == "" || fecha_hasta == null)
                {
                    file_date = DateTime.Today;
                    file_date_before = file_date.AddMonths(-1);
                    fecha_hasta = file_date.ToString("dd-MM-yyyy");
                    fecha_desde = file_date_before.ToString("dd-MM-yyyy");
                }
                System.Threading.Thread.Sleep(1000);
                SelectElement selectList = new SelectElement(chrome.FindElement(By.XPath("//*[@id='IdProveedor']")));
                IList<IWebElement> voptions = selectList.Options;
                console.WriteLine("  Cantidad de proveedores en lista: " + voptions.Count);
                main_web_page = "https://www.panamacompra.gob.pa/Inicio/#/busqueda-tienda-virtual";
                #endregion
                #region Extraer ordenes por cada proveedor
                console.WriteLine("  Extraer Ordenes de Compra por cada proveedor");
                //int z = 1;
                //por cada proveedores de la categoria BIENES INFORMÁTICOS, REDES Y COMUNICACIONES
                //foreach (IWebElement selectElement in voptions.Skip(1))
                for (int z = 1; z < voptions.Count(); z++)
                {
                    try
                    {
                        //No se puede utilizar la variable selectElement debido a que en la segunda interacion del foreach da error ya que no encuentra la página igual
                        //por lo que se utiliza una variable Z que inicia en 1 y aumenta en cada vuelta
                        //cuando z es igual al numero de opciones en el dropdown de proveedores significa que ya termino con el ultimo proveedor
                        if (z == voptions.Count())
                        {
                            break;
                        }

                        //se le suma un 1 a Z debido a que la primera opcion es "Todos" y no cuenta
                        IWebElement prov_option = chrome.FindElement(By.XPath($"//*[@id='IdProveedor']/option[{z + 1}]"));
                        vendor_text = prov_option.Text.ToString();
                        console.WriteLine("   Vendor: " + vendor_text);

                        //Se debe de inicializar dentro del foreach para que se pueda seleccionar la opción ya que el boton de buscar desliga el HTML del chrome
                        SelectElement vendor_select2 = new SelectElement(chrome.FindElement(By.XPath("//*[@id='IdProveedor']")));
                        System.Threading.Thread.Sleep(1000);
                        chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
                        try
                        { vendor_select2.SelectByIndex(z); }
                        catch (Exception)
                        { vendor_select2.SelectByIndex(z); }
                        chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);

                        try
                        {
                            // se debe de mover el cursor al boton de buscar ya que a veces no lo encuentra
                            IWebElement buscar = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button"));
                            //a aveces da error en mover al elemento porque la página cambia constantemente en su estructura
                            new Actions(chrome).MoveToElement(buscar).Perform();
                            //se minimiza y maximiza la ventana para que así el movetoElement se vea reflejado
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
                            //a veces da error el primer intento del click, el segundo usualmente no falla
                            chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button")).Click(); //BUSCAR
                        }
                        //chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button")).Click(); //BUSCAR
                        console.WriteLine("Buscar");
                        System.Threading.Thread.Sleep(3000);
                        main_web_page = chrome.Url;

                        string noResults = "";
                        try
                        {
                            //toma la primera columna de la primer fila de la tabla para verificar si tiene registros
                            noResults = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div/div/table/tbody/tr/td")).Text;
                        }
                        catch (Exception)
                        {

                        }

                        if (noResults == "No hay registro")
                        {
                            continue;
                        }
                        cant_filas = "";
                        try
                        { cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text; }
                        catch (Exception)
                        { cant_filas = ""; }

                        //int ct = 0;
                        while (cant_filas == "Se encontraron Pedidos Publicados")
                        {
                            System.Threading.Thread.Sleep(1000);
                            try
                            { cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text; }
                            catch (Exception)
                            { cant_filas = ""; }
                        }


                        console.WriteLine("   " + cant_filas);
                        int mas = 0;
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

                                //var pag_last = pagination_last.FindElement(By.TagName("a"));
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

                        cant_filas = cant_filas.Split(new string[] { ": " }, StringSplitOptions.None)[1];

                        double num1 = 1;
                        try
                        {
                            num1 = double.Parse(cant_filas);
                        }
                        catch (Exception ex)
                        {
                            console.WriteLine(ex.Message);
                            chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/form[2]/div[6]/button")).Click(); //BUSCAR
                            console.WriteLine("   Buscar");
                            System.Threading.Thread.Sleep(3000);
                            cant_filas = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[3]/div/small[2]")).Text;
                            if (cant_filas.Contains("+"))
                            {
                                do
                                {
                                    IWebElement pagination_last = chrome.FindElement(By.CssSelector("[aria-label=Last]"));
                                    try
                                    {

                                        new Actions(chrome).MoveToElement(pagination_last).Perform();
                                    }
                                    catch (Exception)
                                    {
                                        try
                                        {

                                            new Actions(chrome).MoveToElement(pagination_last).Perform();
                                        }
                                        catch (Exception)
                                        {

                                        }
                                    }
                                    //var pag_last = pagination_last.FindElement(By.TagName("a"));
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

                                    }
                                    catch (Exception)
                                    {
                                        try
                                        {
                                            new Actions(chrome).MoveToElement(pagination_first).Perform();
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
                            //cant_filas = cant_filas.Substring(15, 3).Trim();
                            //if (cant_filas.Contains(" "))
                            //{
                            //    cant_filas = cant_filas.Substring(0, 2).Trim();
                            //}
                            cant_filas = cant_filas.Split(new string[] { ": " }, StringSplitOptions.None)[1];

                            num1 = double.Parse(cant_filas);
                        }

                        double num2 = 10;
                        filas = (num1 / num2);
                        double pag_row2 = Math.Ceiling(filas);
                        if (pag_row2 == 0)
                        { pag_row2 = 1; }

                        for (int i = 1; i <= pag_row2; i++)
                        {

                            int rows = excelResult.Rows.Count;

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
                            System.Threading.Thread.Sleep(1000);
                            IWebElement tableElement = chrome.FindElement(By.XPath("/html/body/app-root/main/ng-component/div/div/div/div[1]/div[2]/table/tbody"));
                            IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));
                            int e = 0;
                            int w = 1;
                            //for por cada fila de la tabla principal de la pagina web
                            foreach (IWebElement element in trCollection)
                            {
                                IList<IWebElement> tdCollection;
                                tdCollection = element.FindElements(By.TagName("td")); //toma las columnas de la fila 
                                                                                       //if (tdCollection.Count > 0)
                                                                                       //{
                                string id = tdCollection[0].Text;
                                string entidad = tdCollection[0].Text;
                                string descripcion = tdCollection[2].Text;
                                string fecha = tdCollection[3].Text;
                                string registroUnicoPedido = tdCollection[1].Text; //Registro único de pedido

                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Extraer PO", $"Id: {id}, Entidad: {entidad}, Descripción: {descripcion}, Fecha: {fecha}, Registro único de pedido: {registroUnicoPedido}", root.Subject);
                                respFinal = respFinal + "\\n" + "Extraer PO " + $"Id: {id}, Entidad: {entidad}, Descripción: {descripcion}, Fecha: {fecha}, Registro único de pedido: {registroUnicoPedido}";


                                IWebElement link = tdCollection[1];
                                var linkhref = link.FindElement(By.TagName("a"));
                                // /html/body/app-root/main/ng-component/div/div/div/div[1]/table/tbody/tr[3]/td[2]/a
                                // /html/body/app-root/main/ng-component/div/div/div/div[1]/table/tbody/tr[1]/td[2]/a
                                //moverse al final de la pagina
                                try
                                {
                                    if (w == 1)
                                    {
                                        try
                                        {
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
                                    //chrome.Manage().Window.Minimize();
                                    //chrome.Manage().Window.Maximize();
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
                                        try
                                        {

                                            chrome.Manage().Window.Minimize();
                                            chrome.Manage().Window.Maximize();
                                        }
                                        catch (Exception)
                                        {

                                        }
                                    }
                                }
                                string linkHref = linkhref.GetAttribute("href");
                                console.WriteLine("    Click en el Codigo Unico Pedido: " + registroUnicoPedido);
                                string html = "";

                                //meterme en la pagina
                                //chrome.Navigate().GoToUrl(linkHref);

                                try
                                { linkhref.Click(); }
                                catch (Exception)
                                {
                                    linkhref.Click();
                                }

                                System.Threading.Thread.Sleep(2000);

                                chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab
                                #region extrae la información de la orden
                                //hacer todo lo que tenga que hacer en la nueva hoja-----
                                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[3]/td"))); }
                                catch { }


                                string po_url = linkHref;

                                string fecha_registro = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[3]/td")).Text;
                                //fecha_registro = fecha_registro.Remove(fecha_registro.Length - 12);
                                DateTime RDate = Convert.ToDateTime(fecha_registro);
                                fecha_registro = RDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                //fecha_registro = "'" + fecha_registro;

                                string fecha_doc = "";
                                try
                                {
                                    fecha_doc = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[6]/div[2]/table/tbody/tr[2]/td[3]")).Text;

                                }
                                catch (Exception)
                                {

                                    fecha_doc = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[7]/div[2]/table/tbody/tr[2]/td[3]")).Text;
                                }
                                DateTime oDate = DateTime.Now;
                                if (fecha_doc != "")
                                {

                                    //fecha_doc = fecha_doc.Remove(fecha_doc.Length - 12);
                                    oDate = Convert.ToDateTime(fecha_doc);
                                }

                                string mes2 = oDate.Month.ToString();
                                string ano = oDate.Year.ToString();
                                fecha_doc = mes2 + "-" + ano;

                                string producto = "";
                                string total = "";
                                string cantidad = "";


                                //int contador_subtotal = doc.DocumentNode.SelectNodes("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle']/tr").Count();
                                IWebElement pedidoDetalle = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[4]/div[2]/table"));
                                int contador_subtotal = pedidoDetalle.FindElements(By.TagName("tr")).Count;

                                string sub_total = "";

                                List<PoProductInfo> PoproductInfo = new List<PoProductInfo>();
                                for (int x = 3; x < contador_subtotal; x++)
                                {
                                    PoProductInfo pInfo = new PoProductInfo();
                                    try
                                    {
                                        string productId = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[1]")).Text;
                                        if (string.IsNullOrWhiteSpace(productId))
                                        {
                                            //*[@id="elementToPDF"]/div/div[4]/div[2]/table/tbody/tr[4]/td[3]/div/span[2]
                                            sub_total = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[3]/div")).Text.Replace("B/.\r\n", "");
                                            break;
                                        }
                                        producto = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[2]")).Text;
                                        cantidad = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[4]")).Text;
                                        precio_unitario = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[3]/div")).Text.Replace("B/.\r\n", "");
                                        total = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[7]/div")).Text.Replace("B/.\r\n", "");

                                        //buscar en tabla#1
                                        System.Data.DataRow[] productInfo = productTypes.Select($"productCode ='{productId}'"); //like '%" + institu + "%'"
                                        string linea_producto = "";
                                        string tipo_producto = "";
                                        string gbm_participa = "";

                                        if (productInfo.Count() > 0)
                                        {
                                            linea_producto = productInfo[0]["productLine"].ToString();
                                            tipo_producto = productInfo[0]["typeProduct"].ToString();
                                            gbm_participa = productInfo[0]["gbmParticipate"].ToString();
                                        }



                                        console.WriteLine("     Agregar informacion al excel, producto " + producto);
                                        DataRow rRow = excelResult.Rows.Add();
                                        rRow["Convenio"] = convenio_name;
                                        rRow["Proveedor"] = vendor_text;
                                        rRow["Fecha Desde"] = fecha_desde;
                                        rRow["Fecha Hasta"] = fecha_hasta;
                                        rRow["Entidad"] = entidad;
                                        rRow["Registro único de pedido"] = registroUnicoPedido;
                                        rRow["Fecha de Registro"] = fecha_registro;
                                        rRow["Fecha de Publicacion"] = fecha_doc;
                                        rRow["Producto/Servicio"] = producto;
                                        rRow["Cantidad"] = cantidad;
                                        rRow["Total del Producto"] = total;
                                        rRow["Precio Unitario"] = precio_unitario;
                                        rRow["Sub-Total de la PO"] = sub_total;
                                        rRow["Línea de producto"] = linea_producto;
                                        rRow["Tipo de producto (Lenguaje GBM)"] = tipo_producto;
                                        rRow["GBM participa?"] = gbm_participa;
                                        rRow["Link"] = po_url;

                                        excelResult.AcceptChanges();

                                        pInfo.singleOrderRecord = registroUnicoPedido;
                                        pInfo.product = productId;
                                        pInfo.amount = cantidad;
                                        pInfo.total = total;
                                        pInfo.unitPrice = precio_unitario.ToString();
                                        pInfo.subtotal = sub_total;
                                        pInfo.typeProduct = tipo_producto;
                                        pInfo.gbmParticipate = gbm_participa;
                                        pInfo.productLine = linea_producto;
                                        pInfo.active = "1";
                                        pInfo.createdBy = "databot";
                                        PoproductInfo.Add(pInfo);

                                        e++;

                                    }
                                    catch (Exception ex)
                                    { }
                                }
                                //agregar información a la base de datos
                                Dictionary<string, string> info = new Dictionary<string, string>
                                {
                                    ["singleOrderRecord"] = registroUnicoPedido,
                                    ["agreement"] = convenio_name,
                                    ["vendor"] = vendor_text,
                                    ["dateFrom"] = file_date_before.ToString("yyyy-MM-dd"),
                                    ["dateTo"] = file_date.ToString("yyyy-MM-dd"),
                                    ["entity"] = entidad,
                                    ["registrationDate"] = RDate.Date.ToString("yyyy-MM-dd"),
                                    ["publicationDate"] = oDate.Date.ToString("yyyy-MM-dd"),
                                    ["documentLink"] = po_url,
                                    ["active"] = "1",
                                    ["createdBy"] = "databot"
                                };

                                //true todo bien, false significa que dio un error
                                bool add_sql = lpsql.insertInfoPurchaseOrder(info, "purchaseOrderCompetition");
                                if (!add_sql)
                                {
                                    //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                    //para enviarlo por email y agregarla
                                    resp_sql = resp_sql + po_url + "<br>";
                                    resp_add_sql = false;
                                }
                                else
                                {
                                    //insertar productos
                                    bool addProducts = lpsql.insertInfoProductsCompetition(PoproductInfo);
                                    if (!addProducts)
                                    {
                                        //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                        //para enviarlo por email y agregarla
                                        resp_sql = resp_sql + po_url + "<br>";
                                        resp_add_sql = false;
                                    }
                                }
                                #endregion
                                System.Threading.Thread.Sleep(1000);
                                chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                                chrome.Close();
                                System.Threading.Thread.Sleep(1000);
                                chrome.SwitchTo().Window(chrome.WindowHandles[0]); //regreso a la pagina principal y sigo con la siguiente fila
                            } //foreach fila en la tabla main
                        } //for cantidad de paginas de la tabla main


                    }
                    catch (Exception ex)
                    {
                        #region catch

                        console.WriteLine(ex.ToString());
                        validar_lineas = false;
                        resp_add_sql = false;

                        DataRow rRow = excelResult.Rows.Add();
                        rRow["Convenio"] = convenio_name;
                        rRow["Proveedor"] = vendor_text;
                        rRow["Fecha Desde"] = fecha_desde;
                        rRow["Fecha Hasta"] = fecha_hasta;
                        rRow["Entidad"] = "Error al descargar informacion";
                        excelResult.AcceptChanges();

                        #endregion
                    }
                    //z++;
                } //for vendor

                #endregion
                #region cerrar chrome
                console.WriteLine("  Guardar el reporte");
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                #endregion
                #region Crear Excel
                console.WriteLine("Save Excel...");
                XLWorkbook wb = new XLWorkbook();
                IXLWorksheet ws = wb.Worksheets.Add(excelResult, "Ordenes Competencia");
                ws.Columns().AdjustToContents();
                ws.Range($"C2:D{excelResult.Rows.Count + 1}").Style.NumberFormat.Format = "dd/MM/yyyy";
                ws.Range($"G2:G{excelResult.Rows.Count + 1}").Style.NumberFormat.Format = "dd/MM/yyyy";
                for (int i = 2; i <= excelResult.Rows.Count + 1; i++)
                {
                    try
                    {
                        ws.Cell(i, 17).Hyperlink = new XLHyperlink(ws.Cell(i, 17).Value.ToString());
                    }
                    catch (Exception ex)
                    {

                    }
                }

                string ruta = root.FilesDownloadPath + $"\\Reporte de Convenio Competencia {DateTime.Now.ToString("dd_MM_yyyy")}.xlsx";
                if (File.Exists(ruta))
                {
                    File.Delete(ruta);
                }
                wb.SaveAs(ruta);
                #endregion
                #region Enviar reporte
                console.WriteLine("  Enviando Reporte");
                string[] adjunto = { ruta };
                string mes_text = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(DateTime.Now.AddMonths(-1).ToString("MMMM", CultureInfo.CreateSpecificCulture("es")));
                string year = DateTime.Now.Year.ToString();
                if (DateTime.Now.AddMonths(-1).Month == 12)
                { year = DateTime.Now.AddMonths(-1).Year.ToString(); }

                if (validar_lineas == false)
                {
                    root.BDUserCreatedBy = "KVANEGAS";
                    //enviar email de repuesta de error
                    string[] cc = { "appmanagement@gbm.net", "kvanegas@gbm.net" };
                    mail.SendHTMLMail("A continuacion se adjunta reporte de las ordenes publicas en Convenio Marco de la competencia en el Mes " + mes_text + " del año " + year, new string[] { "dmeza@gbm.net" }, "Reporte de Competencia mensual de las ordenes de Convenio Marco - " + DateTime.Now.ToString("dd/MM/yyyy"), cc, adjunto);
                }
                else
                {
                    //enviar email de repuesta de exito 
                    string[] cc = { "frivas@gbm.net", "tdiaz@gbm.net", "lotero@gbm.net" };

                    string msj = $"A continuacion se adjunta reporte de las ordenes publicas en Convenio Marco de la competencia en el mes de {mes_text} del año {year}";
                    string html = Properties.Resources.emailtemplate1;
                    html = html.Replace("{subject}", "Reporte de Competencia mensual de las órdenes de Convenio Marco");
                    html = html.Replace("{cuerpo}", msj);
                    html = html.Replace("{contenido}", "");
                    console.WriteLine("Send Email...");

                    mail.SendHTMLMail(html, new string[] { "kvanegas@gbm.net" }, "Reporte de Competencia mensual de las órdenes de Convenio Marco - " + DateTime.Now.ToString("dd/MM/yyyy"), cc, adjunto);
                    root.BDUserCreatedBy = "KVANEGAS";


                }
                if (resp_add_sql == false)
                {
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + DateTime.Now.ToString("dd/MM/yyyy"), cc, adjunto);
                    root.BDUserCreatedBy = "DMEZA";

                }

                #endregion
                root.requestDetails = respFinal;
                return validar_lineas.ToString();


            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                return("Error: " + ex.Message);
            }
        }

    }
    public class PoProductInfo
    {
        public string singleOrderRecord { get; set; }
        public string product { get; set; }
        public string amount { get; set; }
        public string total { get; set; }
        public string unitPrice { get; set; }
        public string subtotal { get; set; }
        public string typeProduct { get; set; }
        public string gbmParticipate { get; set; }
        public string productLine { get; set; }
        public string active { get; set; }
        public string createdBy { get; set; }
    }
}
