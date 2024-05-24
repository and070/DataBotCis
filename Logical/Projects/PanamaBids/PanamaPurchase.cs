using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DataBotV5.Data;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using SAP.Middleware.Connector;
using System.Net.Mail;
using System.Net;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.Data.Projects.PanamaBids;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;

namespace DataBotV5.Logical.Projects.PanamaBids
{
    /// <summary>
    /// Clase Logical encargada de compras de panamá.
    /// </summary>
    class PanamaPurchase
    {
        #region variables globales
        WebInteraction sel = new WebInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ProcessAdmin padmin = new ProcessAdmin();
        Rooting roots = new Rooting();
        Credentials cred = new Credentials();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();
        Log log = new Log();

        WebexTeams wt = new WebexTeams();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        SharePoint sharep = new SharePoint();
        Database db2 = new Database();
        ValidateData val = new ValidateData();
        CRUD crud = new CRUD();
        SapVariants sap = new SapVariants();
        public List<info_vendor> lista_vendors = new List<info_vendor>();
        DataTable registo_unico_info = new DataTable();
        public Excel.Workbook xlWorkBookAM;
        public Excel.Worksheet xlWorkSheetAM;
        string sapSys = "ERP";
        #endregion
        /// <summary>
        /// Reporte de la competencia
        /// </summary>
        /// <returns></returns>
        public string PaPurchaseWeb()
        {
            #region variables privadas
            string respuesta = "";
            string cant_filas = "";
            double filas = 0;
            int pag_row = 0;
            string id_producto = "";
            string main_web_page = "";
            double precio_unitario = 0;
            double prod_cant = 0;
            double prod_total = 0;
            string vendor_text = "";
            string resp_sql = "";
            bool resp_add_sql = true;
            int cont_adj = 0;
            bool validar_lineas = true;
            DateTime file_date = DateTime.MinValue;
            DateTime file_date_before = DateTime.MinValue;

            DataTable productTypes = lpsql.TypeProduct();
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
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            #region titulos_excel
            xlWorkSheet.Cells[1, 1].value = "Convenio";
            xlWorkSheet.Cells[1, 2].value = "Proveedor";
            xlWorkSheet.Cells[1, 3].value = "Fecha Desde";
            xlWorkSheet.Cells[1, 4].value = "Fecha Hasta";
            xlWorkSheet.Cells[1, 5].value = "Entidad";
            xlWorkSheet.Cells[1, 6].value = "Registro único de pedido";
            xlWorkSheet.Cells[1, 7].value = "Fecha de Registro";
            xlWorkSheet.Cells[1, 8].value = "Fecha de Publicacion";
            //xlWorkSheet.Cells[1, 9].value = "Empresa";
            xlWorkSheet.Cells[1, 9].value = "Producto/Servicio";
            xlWorkSheet.Cells[1, 10].value = "Cantidad";
            xlWorkSheet.Cells[1, 11].value = "Total del Producto";
            xlWorkSheet.Cells[1, 12].value = "Precio Unitario";
            xlWorkSheet.Cells[1, 13].value = "Sub-Total de la PO";
            xlWorkSheet.Cells[1, 14].value = "Línea de producto";
            xlWorkSheet.Cells[1, 15].value = "Tipo de producto (Lenguaje GBM)";
            xlWorkSheet.Cells[1, 16].value = "GBM participa?";
            //xlWorkSheet.Cells[1, 17].value = "Nombre del PDF adjunto";
            xlWorkSheet.Cells[1, 17].value = "Link";
            Excel.Worksheet xlSheet = xlWorkBook.ActiveSheet;
            for (int i = 1; i <= 17; i++)
            {
                Excel.Range rango = (Excel.Range)xlSheet.Cells[1, i];
                rango.Interior.Color = Excel.XlRgbColor.rgbRoyalBlue;
                rango.Font.Color = Excel.XlRgbColor.rgbWhite;
            }
            Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
            Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;

            #endregion
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo");
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl("https://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo"); }
            //System.Threading.Thread.Sleep(5000);
            //js executor para subir al inicio de pagina
            IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
            #endregion
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='apc.fields.IdConvenio']"))); }
            catch { }
            chrome.Manage().Cookies.DeleteAllCookies();

            console.WriteLine("  Seleccionar Convenio");
            SelectElement convenio_select = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdConvenio']")));
            System.Threading.Thread.Sleep(1000);
            convenio_select.SelectByValue("number:107"); //BIENES INFORMÁTICOS, REDES Y COMUNICACIONES
            System.Threading.Thread.Sleep(3000);
            string convenio_name = convenio_select.SelectedOption.Text;

            string fecha_desde = chrome.FindElement(By.XPath("//*[@id='apc.fields.fd']")).Text;
            string fecha_hasta = chrome.FindElement(By.XPath("//*[@id='apc.fields.fh']")).Text;
            if (fecha_hasta == "" || fecha_hasta == null)
            {
                file_date = DateTime.Today;
                file_date_before = file_date.AddMonths(-1);
                fecha_hasta = "'" + file_date.ToString("dd-MM-yyyy");
                fecha_desde = "'" + file_date_before.ToString("dd-MM-yyyy");
            }


            System.Threading.Thread.Sleep(1000);
            SelectElement selectList = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdProveedor']")));
            IList<IWebElement> voptions = selectList.Options;
            console.WriteLine("  Cantidad de proveedores en lista: " + voptions.Count);
            main_web_page = "http://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo";


            console.WriteLine("  Extraer Ordenes de Compra por cada proveedor");
            int z = 1;
            //por cada proveedores de la categoria BIENES INFORMÁTICOS, REDES Y COMUNICACIONES
            foreach (IWebElement selectElement in voptions.Skip(1))
            {
                try
                {
                    //chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    //Se debe de inicializar dentro del foreach para que se pueda seleccionar la opción ya que el boton de buscar desliga el HTML del chrome
                    SelectElement vendor_select2 = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdProveedor']")));

                    IWebElement prov_option = chrome.FindElement(By.XPath($"//*[@id='apc.fields.IdProveedor']/option[{z + 1}]"));

                    vendor_text = prov_option.Text.ToString();

                    //if (vendor_text != "-- Seleccione --") //&& vendor_text != "GBM de Panamá, S.A"
                    //{
                    console.WriteLine("   Vendor: " + vendor_text);


                    System.Threading.Thread.Sleep(1000);
                    chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
                    try
                    { vendor_select2.SelectByIndex(z); }
                    catch (Exception)
                    { vendor_select2.SelectByIndex(z); }
                    chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
                    #region calendario
                    //revisar calendario
                    //try
                    //{
                    //console.WriteLine("...................................................................");
                    //chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/span/button")).Click();
                    //chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/div/ul/li[2]/span/button[1]")).Click();

                    //*[@id="datepicker-1721-3270-7"]/button
                    //*[@id="datepicker-1925-7847-41"]/button
                    //*[@id="datepicker-1925-7847-41"]/button/span
                    //*[@id="datepicker-1925-7847-41"]
                    //*[@id="datepicker-1925-7847-9"]/button
                    //*[@id="datepicker-1925-7847-8"]
                    //*[@id="datepicker-2279-8891-8"]
                    //*[@id="datepicker-2381-4911-8"]/button
                    //*[@id="datepicker-2567-43-8"]


                    //IWebElement datetable = chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/div/ul/li[1]/div/div/div/table/tbody"));


                    // List<IWebElement> columns = datetable.FindElement(By.TagName("td"));

                    //}
                    //catch (Exception)
                    //{ }
                    #endregion

                    chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[4]/div/center/button")).Click(); //BUSCAR
                    console.WriteLine("   Buscar");
                    System.Threading.Thread.Sleep(3000);
                    main_web_page = chrome.Url;

                    cant_filas = "";
                    try
                    { cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text; }
                    catch (Exception)
                    { cant_filas = ""; }

                    //int ct = 0;
                    while (cant_filas == "Se encontraron Pedidos Publicados")
                    {
                        System.Threading.Thread.Sleep(1000);
                        try
                        { cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text; }
                        catch (Exception)
                        { cant_filas = ""; }
                    }


                    console.WriteLine("   " + cant_filas);
                    if (cant_filas != "Se encontraron 0 Pedidos Publicados")
                    {
                        int mas = 0;
                        if (cant_filas.Contains("+"))
                        {
                            do
                            {
                                IWebElement pagination_last = chrome.FindElement(By.ClassName("pagination-last"));
                                var pag_last = pagination_last.FindElement(By.TagName("a"));
                                pag_last.Click();
                                System.Threading.Thread.Sleep(1500);
                                cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                mas++;
                            } while (cant_filas.Contains("+"));

                            if (mas > 0)
                            {
                                IWebElement pagination_first = chrome.FindElement(By.ClassName("pagination-first"));
                                var pag_first = pagination_first.FindElement(By.TagName("a"));
                                pag_first.Click();
                                System.Threading.Thread.Sleep(1000);
                                cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                            }
                        }
                        cant_filas = cant_filas.Substring(15, 3).Trim();
                        if (cant_filas.Contains(" "))
                        {
                            cant_filas = cant_filas.Substring(0, 2).Trim();
                        }
                        double num1 = 1;
                        try
                        {
                            num1 = double.Parse(cant_filas);
                        }
                        catch (Exception ex)
                        {
                            console.WriteLine(ex.Message);
                            chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[4]/div/center/button")).Click(); //BUSCAR
                            console.WriteLine("   Buscar");
                            System.Threading.Thread.Sleep(3000);
                            cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                            if (cant_filas.Contains("+"))
                            {
                                do
                                {
                                    IWebElement pagination_last = chrome.FindElement(By.ClassName("pagination-last"));
                                    var pag_last = pagination_last.FindElement(By.TagName("a"));
                                    pag_last.Click();
                                    System.Threading.Thread.Sleep(1500);
                                    cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                    mas++;
                                } while (cant_filas.Contains("+"));



                                if (mas > 0)
                                {
                                    IWebElement pagination_first = chrome.FindElement(By.ClassName("pagination-first"));
                                    var pag_first = pagination_first.FindElement(By.TagName("a"));
                                    pag_first.Click();
                                    System.Threading.Thread.Sleep(1000);
                                    cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                }
                            }
                            cant_filas = cant_filas.Substring(15, 3).Trim();
                            if (cant_filas.Contains(" "))
                            {
                                cant_filas = cant_filas.Substring(0, 2).Trim();
                            }
                            num1 = double.Parse(cant_filas);
                        }

                        double num2 = 10;
                        filas = (num1 / num2);
                        double pag_row2 = Math.Ceiling(filas);
                        if (pag_row2 == 0)
                        { pag_row2 = 1; }

                        for (int i = 1; i <= pag_row2; i++)
                        {

                            int rows = xlWorkSheet.UsedRange.Rows.Count + 1;

                            //next en pagination
                            if (i != 1)
                            {
                                console.WriteLine("     Siguiente pagina");
                                IWebElement pagination_next = chrome.FindElement(By.ClassName("pagination-next"));
                                var pag_next = pagination_next.FindElement(By.TagName("a"));
                                pag_next.Click();
                            }
                            IWebElement tableElement = chrome.FindElement(By.XPath("//*[@id='toTopBA']"));
                            IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));
                            int e = 0;
                            //for por cada fila de la tabla principal de la pagina web
                            foreach (IWebElement element in trCollection.Skip(1))
                            {
                                IList<IWebElement> tdCollection;
                                tdCollection = element.FindElements(By.TagName("td")); //toma las columnas de la fila 
                                                                                       //if (tdCollection.Count > 0)
                                                                                       //{
                                string id = tdCollection[0].Text;
                                string entidad = tdCollection[1].Text;
                                string descripcion = tdCollection[3].Text;
                                string fecha = tdCollection[4].Text;
                                string registroUnicoPedido = tdCollection[2].Text; //Registro único de pedido
                                IWebElement link = tdCollection[2];
                                IWebElement linkElement = link.FindElement(By.TagName("a"));
                                string linkHref = linkElement.GetAttribute("href");
                                //new Actions(chrome).MoveToElement(link).Perform(); //mover la vista del chrome al elemento indicado
                                console.WriteLine("    Click en el Codigo Unico Pedido: " + registroUnicoPedido);

                                //try
                                //{ linkhref.Click(); }
                                //catch (Exception)
                                //{ linkhref.Click(); }

                                //var descargar_link = linkhref.GetAttribute("href");
                                ////Console.WriteLine(descargar_link.ToString());
                                //chrome.Navigate().GoToUrl(descargar_link.ToString());

                                //System.Threading.Thread.Sleep(2000);

                                //chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab

                                string html = "";
                                using (WebClient client = new WebClient())
                                {
                                    client.Encoding = UTF8Encoding.UTF8;
                                    html = client.DownloadString(linkHref);
                                }
                                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                doc.LoadHtml(html);

                                string po_url = linkHref;

                                string fecha_registro = doc.DocumentNode.SelectSingleNode("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']").InnerText;

                                //string fecha_registro = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']")).Text;
                                fecha_registro = fecha_registro.Remove(fecha_registro.Length - 5);
                                DateTime RDate = Convert.ToDateTime(fecha_registro);
                                fecha_registro = RDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                fecha_registro = "'" + fecha_registro;

                                string fecha_doc = doc.DocumentNode.SelectSingleNode("//*[@id='ctl00_ContentPlaceHolder1_gvAdjuntos_ctl02_lblFecha']").InnerText;

                                fecha_doc = fecha_doc.Remove(fecha_doc.Length - 5);
                                DateTime oDate = Convert.ToDateTime(fecha_doc);
                                string mes2 = oDate.Month.ToString();
                                string ano = oDate.Year.ToString();
                                fecha_doc = mes2 + "-" + ano;
                                #region descargas

                                //IWebElement pdf = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvAdjuntos_ctl02_lbtUrl']"));
                                //string href_pdf = pdf.GetAttribute("onclick").ToString();
                                //string ext = "";
                                //if (href_pdf.Contains(".pdf"))
                                //{ ext = ".pdf"; }
                                //else if (href_pdf.Contains(".jpeg"))
                                //{ ext = ".jpeg"; }
                                //else if (href_pdf.Contains(".jpg"))
                                //{ ext = ".jpg"; }
                                //else if (href_pdf.Contains(".png"))
                                //{ ext = ".png"; }
                                //else
                                //{ ext = ""; }
                                //string pdf_full_name = "";
                                //string pdf_name = "";
                                //if (ext != "")
                                //{
                                //    pdf_full_name = @"C:\Users\" + Environment.UserName + @"\Downloads\" + pdf.Text + ext;
                                //    pdf_name = pdf.Text + ext;
                                //}
                                //tabla con los productos de la PO
                                //*[@id="ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl02_lblNombre"] 'cada fila de la tabla
                                //*[@id="ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl03_lblNombre"] 'cada fila de la tabla
                                #endregion
                                string producto = "";
                                string total = "";
                                string cantidad = "";



                                int contador_subtotal = doc.DocumentNode.SelectNodes("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle']/tr").Count();
                                string sub_total = doc.DocumentNode.SelectSingleNode("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador_subtotal + "_lblSubtotal']").InnerText;
                                producto = doc.DocumentNode.SelectSingleNode("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl02_lblNombre']").InnerText;
                                int contador = 2;
                                while (producto != "")
                                {
                                    try
                                    {
                                        producto = doc.DocumentNode.SelectSingleNode($"//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0{contador}_lblNombre']").InnerText;
                                        cantidad = doc.DocumentNode.SelectSingleNode($"//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0{contador}_lblCantidad']").InnerText;
                                        total = doc.DocumentNode.SelectSingleNode($"//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0{contador}_lbltotal']").InnerText;

                                        //precio unitario por producto
                                        try
                                        {
                                            float total2 = float.Parse(total, new NumberFormatInfo() { NumberDecimalSeparator = ".", NumberGroupSeparator = "," });

                                            prod_cant = double.Parse(cantidad);
                                            prod_total = double.Parse(total2.ToString());

                                            precio_unitario = (prod_total / prod_cant);
                                        }
                                        catch (Exception)
                                        {

                                        }



                                        string product_id = producto.Substring(0, 2);
                                        if (product_id.Contains("-"))
                                        { product_id = producto.Substring(0, 1); }

                                        //buscar en tabla#1
                                        System.Data.DataRow[] productInfo = productTypes.Select($"id ='{product_id}'"); //like '%" + institu + "%'"

                                        string linea_producto = productInfo[0]["linea_producto"].ToString();
                                        string tipo_producto = productInfo[0]["tipo_producto"].ToString();
                                        string gbm_participa = productInfo[0]["gbm_part"].ToString();

                                        console.WriteLine("     Agregar informacion al excel, producto " + producto);
                                        xlWorkSheet.Cells[rows + e, 1].value = convenio_name;
                                        xlWorkSheet.Cells[rows + e, 2].value = vendor_text;
                                        xlWorkSheet.Cells[rows + e, 3].value = fecha_desde;
                                        xlWorkSheet.Cells[rows + e, 4].value = fecha_hasta;
                                        xlWorkSheet.Cells[rows + e, 5].value = entidad;
                                        xlWorkSheet.Cells[rows + e, 6].value = registroUnicoPedido;
                                        xlWorkSheet.Cells[rows + e, 7].value = fecha_registro;
                                        xlWorkSheet.Cells[rows + e, 8].value = fecha_doc;
                                        //xlWorkSheet.Cells[rows + e, 9].value = cliente;
                                        xlWorkSheet.Cells[rows + e, 9].value = producto;
                                        xlWorkSheet.Cells[rows + e, 10].value = cantidad;
                                        xlWorkSheet.Cells[rows + e, 11].value = total;
                                        xlWorkSheet.Cells[rows + e, 12].value = precio_unitario;
                                        xlWorkSheet.Cells[rows + e, 13].value = sub_total;
                                        try
                                        {
                                            xlWorkSheet.Range["K" + (rows + e)].Copy();
                                            xlWorkSheet.Range["M" + (rows + e)].PasteSpecial(paste, pasteop, false, false);
                                            xlWorkSheet.Range["L" + (rows + e)].PasteSpecial(paste, pasteop, false, false);
                                        }
                                        catch (Exception)
                                        { }
                                        xlWorkSheet.Cells[rows + e, 14].value = linea_producto;
                                        xlWorkSheet.Cells[rows + e, 15].value = tipo_producto;
                                        xlWorkSheet.Cells[rows + e, 16].value = gbm_participa;
                                        //xlWorkSheet.Cells[rows + e, 17].value = pdf_name;
                                        xlWorkSheet.Cells[rows + e, 17].value = po_url;
                                        xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[rows + e, 17], po_url);
                                        //agregar información a la base de datos
                                        //true todo bien, false significa que dio un error
                                        bool add_sql = lpsql.InfoSqlAddCompetencia(convenio_name, vendor_text, file_date_before, file_date, entidad, registroUnicoPedido, RDate.Date, oDate.Date, producto, cantidad, total, precio_unitario, sub_total, linea_producto, tipo_producto, gbm_participa, po_url);
                                        if (add_sql == false)
                                        {
                                            //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                            //para enviarlo por email y agregarla
                                            resp_sql = resp_sql + po_url + "<br>";
                                            resp_add_sql = false;
                                        }
                                        e++;
                                        //log.LogdeCambios("Creacion", roots.BD_Proceso, "Licitaciones Panama", "Crear reporte de convenios de Competencia Panama", vendor_text + ": " + producto, registroUnicoPedido);

                                    }
                                    catch (Exception)
                                    { producto = ""; }
                                    contador++;
                                }
                                #region descargas


                                #endregion
                                #region nchrome
                                #endregion

                                //} //la fila tiene td
                            } //foreach fila en la tabla main
                        } //for cantidad de paginas de la tabla main
                    }
                    else
                    {
                        Console.WriteLine("");
                    }



                    //} //if proveedor drop down
                    //resh++;
                }
                catch (Exception ex)
                {
                    #region catch

                    Console.WriteLine(ex.ToString());
                    validar_lineas = false;
                    resp_add_sql = false;
                    int row = xlWorkSheet.UsedRange.Rows.Count + 1;
                    if (row == 1)
                    { row = 2; }
                    xlWorkSheet.Cells[row, 1].value = convenio_name;
                    xlWorkSheet.Cells[row, 2].value = vendor_text;
                    xlWorkSheet.Cells[row, 3].value = fecha_desde;
                    xlWorkSheet.Cells[row, 4].value = fecha_hasta;
                    xlWorkSheet.Cells[row, 5].value = "Error al descargar informacion";
                    #endregion
                }
                z++;
            } //for vendor

            console.WriteLine("  Guardar el reporte");
            chrome.Close();

            proc.KillProcess("chromedriver", true);

            xlWorkSheet.Columns.AutoFit();

            string mes = "";
            string dia = "";
            mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

            string fecha_file = dia + "_" + mes + "_" + DateTime.Now.Year.ToString();
            string nombre_fila = root.FilesDownloadPath + "\\" + "Reporte de Convenio Competencia " + fecha_file + ".xlsx";
            xlWorkBook.SaveAs(nombre_fila);
            xlWorkBook.Close();
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);
            string[] adjunto = { nombre_fila };

            fecha_file = dia + "/" + mes + "/" + DateTime.Now.Year.ToString();
            console.WriteLine("  Enviando Reporte");

            string mes_text = DeterminarMes(DateTime.Now.AddMonths(-1).Month);
            string year = DateTime.Now.Year.ToString();
            if (DateTime.Now.AddMonths(-1).Month == 12)
            { year = DateTime.Now.AddMonths(-1).Year.ToString(); }

            if (validar_lineas == false)
            {
                //enviar email de repuesta de error
                string[] cc = { "appmanagement@gbm.net", "acolina@gbm.net" };
                mail.SendHTMLMail("A continuacion se adjunta reporte de las ordenes publicas en Convenio Marco de la competencia en el Mes " + mes_text + " del año " + year, new string[] { "dmeza@gbm.net" }, "Reporte de Competencia mensual de las ordenes de Convenio Marco - " + fecha_file, cc, adjunto);
            }
            else
            {
                //enviar email de repuesta de exito 
                string[] cc = { "frivas@gbm.net", "tdiaz@gbm.net", "lotero@gbm.net" };
                mail.SendHTMLMail("A continuacion se adjunta reporte de las ordenes publicas en Convenio Marco de la competencia en el mes de " + mes_text + " del año " + year, new string[] { "kvanegas@gbm.net" }, "Reporte de Competencia mensual de las órdenes de Convenio Marco - " + fecha_file, cc, adjunto);

            }
            if (resp_add_sql == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] { "appmanagement@gbm.net" }, "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + fecha_file, cc, adjunto);

            }

            return validar_lineas.ToString();
        }
        /// <summary>
        /// Reporte de GBMPA
        /// </summary>
        /// <returns></returns>
        public string PaPurchaseWebGBMV2()
        {
            string cant_filas = "";
            double filas = 0;
            string[] adjunto = new string[1];
            int cont_adj = 0;
            string main_web_page = "";
            string vendor_text = "";
            string resp_sql = "";
            bool resp_add_sql = true;
            bool validar_lineas = true;
            bool cisco_add = false;
            bool no_registros = false;
            string[] CopyCC = new string[1];
            int cont = 0;
            int cont_am = 0;
            DateTime file_date = DateTime.MinValue;
            DateTime file_date_before = DateTime.MinValue;
            Dictionary<string, string> adj_names = new Dictionary<string, string>();
            Dictionary<string, string> newrow = new Dictionary<string, string>();
            Dictionary<string, string> AMs = new Dictionary<string, string>();

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
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Workbook xlWorkBookCisco;
            Excel.Worksheet xlWorkSheetCisco;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];

            xlWorkBookCisco = xlApp.Workbooks.Add();
            xlWorkSheetCisco = (Excel.Worksheet)xlWorkBookCisco.Sheets[1];
            string[] columnas = lpsql.ConvenioColumns("reporte_convenio_gbmpa");
            #region titulos_excel
            for (int i = 0; i < columnas.Length; i++)
            {
                xlWorkSheet.Cells[1, i + 1].value = columnas[i].ToString();
                xlWorkSheetCisco.Cells[1, i + 1].value = columnas[i].ToString();

                Excel.Range rango = (Excel.Range)xlWorkSheet.Cells[1, i + 1];
                rango.Interior.Color = Excel.XlRgbColor.rgbRoyalBlue;
                rango.Font.Color = Excel.XlRgbColor.rgbWhite;

                Excel.Range rango2 = (Excel.Range)xlWorkSheetCisco.Cells[1, i + 1];
                rango2.Interior.Color = Excel.XlRgbColor.rgbRoyalBlue;
                rango2.Font.Color = Excel.XlRgbColor.rgbWhite;
            }
            xlWorkSheetCisco.Cells[1, columnas.Length + 1].value = "Account Manager Asignado";
            Excel.Range rango3 = (Excel.Range)xlWorkSheetCisco.Cells[1, columnas.Length + 1];
            rango3.Interior.Color = Excel.XlRgbColor.rgbRoyalBlue;
            rango3.Font.Color = Excel.XlRgbColor.rgbWhite;
            Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
            Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;

            #endregion
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo");
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl("https://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo"); }
            //System.Threading.Thread.Sleep(5000);
            //js executor para subir al inicio de pagina
            IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
            #endregion

            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='apc.fields.IdConvenio']"))); }
            catch { }
            chrome.Manage().Cookies.DeleteAllCookies();

            console.WriteLine("  Seleccionar Convenio");
            SelectElement convenio_select = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdConvenio']")));
            System.Threading.Thread.Sleep(1000);
            convenio_select.SelectByValue("number:107"); //BIENES INFORMÁTICOS, REDES Y COMUNICACIONES
            System.Threading.Thread.Sleep(3000);
            string convenio_name = convenio_select.SelectedOption.Text;

            #region calendario
            System.Threading.Thread.Sleep(250); //*[@id="busquedaC2"]/div[3]/div[x]/p/span/button
            try
            { chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/span/button")).Click(); } //fecha desde 
            catch (Exception ex)
            {
                Console.WriteLine("Error en click calendario" + ex.Message);
                chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/span/button")).Click();
            }
            //*[@id="busquedaC2"]/div[3]/div[1]/p/div/ul/li[2]/span/button[1]
            System.Threading.Thread.Sleep(250); //*[@id="busquedaC2"]/div[3]/div[2]/p/div/ul/li[2]/span/button[1]
            try
            {
                //TODAY

                chrome.FindElement(By.XPath("//*[@id='imageLazyContainer']/div[3]/ul/li[2]/span/button[1]")).Click(); ////*[@id='busquedaC2']/div[3]/div[1]/p/div/ul/li[2]/span/button[1]
            }
            catch (Exception ex)
            {
                console.WriteLine("Error en click today" + ex.Message);
                System.Threading.Thread.Sleep(2000);
                chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[3]/div[1]/p/div/ul/li[2]/span/button[1]")).Click();
            }

            file_date = DateTime.Today;
            string fecha_hasta = "'" + file_date.ToString("dd-MM-yyyy");
            string fecha_desde = fecha_hasta;
            #endregion

            #region saca la lista de proveedores
            IWebElement vendor_list = chrome.FindElement(By.XPath("//*[@id='apc.fields.IdProveedor']"));
            System.Threading.Thread.Sleep(1000);
            SelectElement selectList = new SelectElement(vendor_list);
            IList<IWebElement> voptions = selectList.Options;
            console.WriteLine("  Cantidad de proveedores en lista: " + voptions.Count);
            #endregion

            console.WriteLine("  Extraer Ordenes de Compra por cada proveedor");

            //int val = 0;
            for (int z = 5; z <= voptions.Count; z++) //lista_vendors.Count vendor_lists.Length - 1
            {
                //int z = 6;
                try
                {
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    SelectElement vendor_select2 = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdProveedor']")));
                    IWebElement prov_option = chrome.FindElement(By.XPath("//*[@id='apc.fields.IdProveedor']/option[" + z + "]"));
                    vendor_text = prov_option.Text.ToString();

                    if (vendor_text == "GBM de Panamá, S.A") //&& vendor_text != "GBM de Panamá, S.A"
                    {
                        console.WriteLine("   Vendor: " + vendor_text);
                        System.Threading.Thread.Sleep(1000);
                        chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
                        try
                        { vendor_select2.SelectByIndex(z - 1); }
                        catch (Exception)
                        { vendor_select2.SelectByIndex(z - 1); }
                        chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);

                        chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[4]/div/center/button")).Click(); //BUSCAR
                        console.WriteLine("   Buscar");
                        System.Threading.Thread.Sleep(3000);
                        main_web_page = chrome.Url;

                        cant_filas = "";
                        try
                        { cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text; }
                        catch (Exception)
                        { cant_filas = ""; }

                        console.WriteLine("   " + cant_filas);
                        if (cant_filas != "Se encontraron 0 Pedidos Publicados")
                        {
                            int mas = 0;
                            if (cant_filas.Contains("+"))
                            {
                                do
                                {
                                    IWebElement pagination_last = chrome.FindElement(By.ClassName("pagination-last"));
                                    var pag_last = pagination_last.FindElement(By.TagName("a"));
                                    pag_last.Click();
                                    System.Threading.Thread.Sleep(1500);
                                    cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                    mas++;
                                } while (cant_filas.Contains("+"));



                                if (mas > 0)
                                {
                                    IWebElement pagination_first = chrome.FindElement(By.ClassName("pagination-first"));
                                    var pag_first = pagination_first.FindElement(By.TagName("a"));
                                    pag_first.Click();
                                    System.Threading.Thread.Sleep(1000);
                                    cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                }
                            }
                            cant_filas = cant_filas.Substring(15, 3).Trim();
                            if (cant_filas.Contains(" "))
                            {
                                cant_filas = cant_filas.Substring(0, 2).Trim();
                            }

                            double num1 = double.Parse(cant_filas);
                            double num2 = 10;
                            filas = (num1 / num2);
                            double pag_row2 = Math.Ceiling(filas);
                            if (pag_row2 == 0)
                            { pag_row2 = 1; }

                            for (int i = 1; i <= pag_row2; i++)
                            {
                                //next en pagination
                                if (i != 1)
                                {
                                    console.WriteLine("     Siguiente pagina");
                                    IWebElement pagination_next = chrome.FindElement(By.ClassName("pagination-next"));
                                    var pag_next = pagination_next.FindElement(By.TagName("a"));
                                    pag_next.Click();
                                }
                                IWebElement tableElement = chrome.FindElement(By.XPath("//*[@id='toTopBA']"));
                                IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));
                                int e = 0;

                                foreach (IWebElement element in trCollection)
                                {
                                    IList<IWebElement> tdCollection;
                                    tdCollection = element.FindElements(By.TagName("td"));
                                    if (tdCollection.Count > 0)
                                    {
                                        string id = tdCollection[0].Text;
                                        string entidad = tdCollection[1].Text;
                                        newrow["ENTIDAD"] = entidad;
                                        string descripcion = tdCollection[3].Text;
                                        string fecha = tdCollection[4].Text;
                                        string link_string = tdCollection[2].Text; //Registro único de pedido
                                        newrow["REGISTRO_UNICO_DE_PEDIDO"] = link_string;
                                        IWebElement link = tdCollection[2];
                                        var linkhref = link.FindElement(By.TagName("a"));
                                        new Actions(chrome).MoveToElement(link).Perform();
                                        console.WriteLine("    Click en el Codigo Unico Pedido: " + link_string);

                                        try
                                        { linkhref.Click(); }
                                        catch (Exception)
                                        { linkhref.Click(); }

                                        System.Threading.Thread.Sleep(2000);

                                        chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab

                                        //hacer todo lo que tenga que hacer en la nueva hoja-----
                                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']"))); }
                                        catch { }

                                        string lugar = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblLugarEntrega']")).Text;

                                        string funcionario = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblNameFuncionario']")).Text;
                                        string telefono = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblCellFuncionario']")).Text;
                                        string email = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblEmailFuncionario']")).Text;
                                        string provincia = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblProvincia']")).Text;

                                        newrow["LUGAR_DE_ENTREGA"] = lugar; newrow["CONTACTO_DE_ENTREGA"] = funcionario;
                                        newrow["TELEFONO"] = telefono; newrow["EMAIL"] = email;
                                        newrow["PROVINCIA"] = provincia;

                                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']"))); }
                                        catch { }
                                        //string cliente = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblNombreProveedor']")).Text;
                                        string po_url = chrome.Url;
                                        string fecha_registro = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']")).Text;
                                        fecha_registro = fecha_registro.Remove(fecha_registro.Length - 5);
                                        DateTime RDate = Convert.ToDateTime(fecha_registro);
                                        fecha_registro = RDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                        fecha_registro = "'" + fecha_registro;

                                        string unidad_solicitante = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblUnidadSolicitante']")).Text,
                                           contactocuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblRegistradoPor']")).Text,
                                           emailcuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblEmail']")).Text;

                                        newrow["UNIDAD_SOLICITANTE"] = unidad_solicitante; newrow["CONTACTO_CUENTA"] = contactocuenta; newrow["EMAIL_CUENTA"] = emailcuenta;
                                        newrow["FECHA_DE_REGISTRO"] = RDate.Date.ToString("yyyy-MM-dd"); newrow["LINK_AL_DOCUMENTO"] = po_url;

                                        try
                                        { new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ibtnExportarPDF']"))).Perform(); }
                                        catch (Exception) { }

                                        string fecha_doc = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvAdjuntos_ctl02_lblFecha']")).Text;
                                        fecha_doc = fecha_doc.Remove(fecha_doc.Length - 5);
                                        DateTime oDate = Convert.ToDateTime(fecha_doc);
                                        fecha_doc = oDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                        fecha_doc = "'" + fecha_doc;

                                        IWebElement pdf = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvAdjuntos_ctl02_lbtUrl']"));
                                        string href_pdf = pdf.GetAttribute("onclick").ToString();
                                        string ext = "";
                                        ext = get_ext(href_pdf);
                                        string pdf_full_name = "";
                                        string pdf_name = "";
                                        if (ext != "")
                                        {
                                            pdf_full_name = root.FilesDownloadPath + "\\" + pdf.Text + ext;
                                            pdf_name = pdf.Text + ext;
                                        }
                                        else
                                        {
                                            pdf_full_name = root.FilesDownloadPath + "\\" + pdf.Text + ".pdf";
                                            pdf_name = pdf.Text + ".pdf";
                                        }

                                        //buscar el AM de la entidad
                                        string[] entidad_info = lpsql.DetermineSector(entidad);
                                        string AM = entidad_info[4];
                                        string sector = entidad_info[0];
                                        string user = "";


                                        string producto = "";
                                        string total = "";
                                        string cantidad = "";

                                        producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl02_lblNombre']")).Text;
                                        int contador_subtotal = 2;
                                        while (producto != "")
                                        {
                                            try
                                            {
                                                producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador_subtotal + "_lblNombre']")).Text;
                                                contador_subtotal++;
                                            }
                                            catch (Exception)
                                            { producto = ""; }

                                        }
                                        //*[@id="ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl04_lblSubtotal"] 
                                        string sub_total = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador_subtotal + "_lblSubtotal']")).Text;

                                        newrow["SUB_TOTAL_ORDEN"] = sub_total; newrow["NOMBRE_DEL_ADJUNTO"] = pdf_name;
                                        newrow["FECHA_DE_PUBLICACION"] = oDate.Date.ToString("yyyy-MM-dd");

                                        producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl02_lblNombre']")).Text;
                                        int contador = 2;
                                        while (producto != "")
                                        {
                                            try
                                            {
                                                producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador + "_lblNombre']")).Text;
                                                cantidad = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador + "_lblCantidad']")).Text;
                                                total = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador + "_lbltotal']")).Text;

                                                //extrae los dias habiles de acuerdo a la cantidad
                                                int dias_h = dias_habiles(int.Parse(cantidad));

                                                //extrae la marca del producto
                                                string product_id = producto.Substring(0, 2);
                                                if (product_id.Contains("-"))
                                                { product_id = producto.Substring(0, 1); }

                                                //buscar en tabla#1
                                                string[] marca_array = lpsql.MarcaProduct(product_id);
                                                string marca = marca_array[0];

                                                //fecha maxima de entrega
                                                DateTime fecha_max_entrega = DateTime.MinValue;
                                                string fecha_max = "";
                                                if (fecha_doc != "")
                                                {
                                                    int dias_ent = dias_h + 2;
                                                    //Excel.IWorksheetFunction workday = (Excel.WorksheetFunction)Excel.WorksheetFunction.WorkDay(fecha_documento, dias_h,"");
                                                    fecha_max_entrega = AddWorkdays(oDate, dias_ent);
                                                    if (fecha_max_entrega != oDate)
                                                    {
                                                        fecha_max = fecha_max_entrega.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                                        fecha_max = "'" + fecha_max;
                                                    }

                                                }

                                                //dias faltantes
                                                double dias_faltantes = 0;
                                                if (fecha_max != "")
                                                {

                                                    try
                                                    {
                                                        DateTime hoy = DateTime.Now;
                                                        dias_faltantes = (fecha_max_entrega.Date - hoy.Date).TotalDays;
                                                        dias_faltantes = Math.Ceiling(dias_faltantes);
                                                    }
                                                    catch (Exception)
                                                    { }

                                                }

                                                newrow["FECHA_MAXIMA_ENTREGA"] = fecha_max_entrega.Date.ToString("yyyy-MM-dd"); newrow["SECTOR"] = sector;
                                                newrow["MARCA"] = marca; newrow["DIAS_ENTREGA"] = dias_h.ToString(); newrow["PRODUCTO_SERVICIO"] = producto;
                                                newrow["CANTIDAD"] = cantidad; ; newrow["TOTAL_DEL_PRODUCTO"] = total;
                                                newrow["DIAS_FALTANTES"] = dias_faltantes.ToString();

                                                newrow["FIANZA_CUMPLIMIENTO"] = "No Aplica";
                                                newrow["ORDEN_COMPRA"] = "";
                                                newrow["OPORTUNIDAD"] = "";
                                                newrow["QUOTE"] = "";
                                                newrow["TIPO_PEDIDO"] = "";
                                                newrow["SALES_ORDER"] = "";
                                                newrow["ESTADO_GBM"] = "Pendiente de Procesar";
                                                newrow["ESTATUS_DE_ORDEN"] = "Refrendado";
                                                newrow["MONTO_MULTA"] = "";
                                                newrow["FORECAST"] = DateTime.MinValue.Date.ToString("yyyy-MM-dd");
                                                newrow["CONVENIO"] = convenio_name;
                                                newrow["CONFIRMACION_ORDEN"] = "";
                                                newrow["COMENTARIOS"] = "";
                                                newrow["FECHA_REAL_ENTREGA"] = "";
                                                newrow["TIPO_FORECAST"] = "E2E";
                                                newrow["VENDOR_ORDER"] = "";

                                                //agrega la información en el excel
                                                int rows = xlWorkSheet.UsedRange.Rows.Count + 1;
                                                int rowsCisco = xlWorkSheetCisco.UsedRange.Rows.Count + 1;
                                                //for por columna

                                                for (int x = 0; x < columnas.Length; x++)
                                                {
                                                    //columnas[i].ToString();
                                                    string columna = xlWorkSheetCisco.Cells[1, x + 1].text.ToString();
                                                    string valor = "";
                                                    try
                                                    {
                                                        valor = newrow[columna];
                                                    }
                                                    catch (Exception)
                                                    {
                                                        valor = "N/A";
                                                    }

                                                    xlWorkSheet.Cells[rows, x + 1].value = valor;
                                                    xlWorkSheet.Range["G" + (rows)].Copy();
                                                    xlWorkSheet.Range["H" + (rows)].PasteSpecial(paste, pasteop, false, false);
                                                    if (columna == "LINK_AL_DOCUMENTO")
                                                    {
                                                        xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[rows, x + 1], po_url);
                                                    }



                                                    if (marca == "Cisco")
                                                    {
                                                        //otro row?
                                                        xlWorkSheetCisco.Cells[rowsCisco, x + 1].value = valor;
                                                        xlWorkSheetCisco.Range["G" + (rowsCisco)].Copy();
                                                        xlWorkSheetCisco.Range["H" + (rowsCisco)].PasteSpecial(paste, pasteop, false, false);
                                                        if (columna == "LINK_AL_DOCUMENTO")
                                                        {
                                                            xlWorkSheetCisco.Hyperlinks.Add(xlWorkSheetCisco.Cells[rows, x + 1], po_url);
                                                        }

                                                    }

                                                }

                                                if (marca == "Cisco")
                                                {
                                                    cisco_add = true;

                                                    //if (AMarrays.All(AM.Contains))
                                                    //{

                                                    //}
                                                    //else
                                                    //{
                                                    //    AMarrays[cont_am] = AM;
                                                    //    cont_am++;
                                                    //    Array.Resize(ref AMarrays, AMarrays.Length + 1);
                                                    //}
                                                    try
                                                    {
                                                        user = AMs[AM];
                                                    }
                                                    catch (Exception)
                                                    {
                                                        string AMemail = am_email(AM);
                                                        string[] sep = new string[] { "@" };
                                                        string[] split = AMemail.Split(sep, StringSplitOptions.None);
                                                        user = split[0].ToString().Trim();
                                                        AMs[AM] = user;
                                                    }


                                                    xlWorkSheetCisco.Cells[rowsCisco, columnas.Length + 1].value = user;
                                                }

                                                //agregar información a la base de datos
                                                //true todo bien, false significa que dio un error
                                                bool add_sql = lpsql.InfoSqlAddGbmV2(newrow);
                                                //bool add_sql = lpsql.info_sql_add_gbm(convenio_name, entidad, producto, cantidad, marca, total, sub_total, orden_compra, link_string, RDate.Date, oDate.Date, fianza, opp, quote, tipo_pedido, sales_order, estado_gbm, estado_orden, dias_h, fecha_max_entrega.Date, dias_faltantes, forecast, provincia, lugar, funcionario, telefono, email, monto_multa, pdf_name, po_url);
                                                if (!add_sql)
                                                {
                                                    //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                                    //para enviarlo por email y agregarla
                                                    resp_sql = resp_sql + po_url + "<br>";
                                                    resp_add_sql = false;
                                                }

                                                log.LogDeCambios("Creacion", roots.BDProcess, "Ventas Panama", "Crear reporte de convenios de Competencia Panama", vendor_text + ": " + producto, link_string);

                                            }
                                            catch (Exception)
                                            { producto = ""; }
                                            contador++;
                                        }

                                        //aqui termina la tabla

                                        //descargable pdf
                                        try
                                        {
                                            if (pdf_full_name != "")
                                            {
                                                adjunto[cont_adj] = pdf_full_name;
                                                cont_adj++;
                                                Array.Resize(ref adjunto, adjunto.Length + 1);
                                            }


                                            pdf.Click();
                                            for (var x = 0; x < 40; x++)
                                            {
                                                if (File.Exists(pdf_full_name)) { break; }
                                                System.Threading.Thread.Sleep(1000);
                                            }
                                            //System.Threading.Thread.Sleep(7000);
                                            chrome.SwitchTo().Window(chrome.WindowHandles[2]);
                                            System.Threading.Thread.Sleep(1000);
                                            chrome.Close();

                                            //agregar la ruta y nombre del archivo como parte de los adjuntos del array del AM
                                            adj_names[AM + "_" + cont] = pdf_full_name;
                                            cont++;

                                        }
                                        catch (Exception ex)
                                        {
                                            console.WriteLine("Error al descargar adjunto: " + ex.Message);
                                            try
                                            {
                                                chrome.SwitchTo().Window(chrome.WindowHandles[2]);
                                                System.Threading.Thread.Sleep(1000);
                                                chrome.Close();
                                            }
                                            catch (Exception)
                                            { }
                                        }
                                        System.Threading.Thread.Sleep(1000);
                                        chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                                        chrome.Close();
                                        System.Threading.Thread.Sleep(1000);
                                        chrome.SwitchTo().Window(chrome.WindowHandles[0]); //regreso a la pagina principal y sigo con la siguiente fila

                                    } //la fila tiene td
                                } //foreach fila en la tabla main
                            } //for cantidad de paginas de la tabla main
                        }
                        else
                        {
                            console.WriteLine("");

                            //no hay registros
                            int rows = xlWorkSheet.UsedRange.Rows.Count + 1;
                            if (rows == 1)
                            { rows = 2; }
                            xlWorkSheet.Cells[rows, 1].value = convenio_name;
                            xlWorkSheet.Cells[rows, 2].value = vendor_text;
                            xlWorkSheet.Cells[rows, 3].value = "No hay registro";
                            no_registros = true;
                        }

                        break;


                    } //if proveedor drop down
                      //resh++;
                }
                catch (Exception ex)
                {
                    //console.WriteLine(ex.ToString());
                    int row = xlWorkSheet.UsedRange.Rows.Count + 1;
                    if (row == 1)
                    { row = 2; }
                    xlWorkSheet.Cells[row, 1].value = convenio_name;
                    xlWorkSheet.Cells[row, 2].value = vendor_text;
                    xlWorkSheet.Cells[row, 3].value = fecha_desde;
                    xlWorkSheet.Cells[row, 4].value = fecha_hasta;
                    xlWorkSheet.Cells[row, 5].value = "Error al descargar informacion";

                }
            } //for vendor

            console.WriteLine("  Guardar el reporte");
            chrome.Close();

            proc.KillProcess("chromedriver", true);

            xlWorkSheet.Columns.AutoFit();
            xlWorkSheetCisco.Columns.AutoFit();

            string mes = "";
            string dia = "";

            mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

            string fecha_file = dia + "_" + mes + "_" + DateTime.Now.Year.ToString();
            string nombre_fila = root.FilesDownloadPath + "\\" + "Reporte de Convenio GBM " + fecha_file + ".xlsx";
            xlWorkBook.SaveAs(nombre_fila);
            xlWorkBook.Close();
            string nombre_fila_Cisco = root.FilesDownloadPath + "\\" + "Reporte de Convenio GBM Networking " + fecha_file + ".xlsx";
            xlWorkBookCisco.SaveAs(nombre_fila_Cisco);
            xlWorkBookCisco.Close();

            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);
            adjunto[cont_adj] = nombre_fila;

            fecha_file = dia + "/" + mes + "/" + DateTime.Now.Year.ToString();
            console.WriteLine("  Enviando Reporte");

            string mes_text = DeterminarMes(DateTime.Now.Month);
            if (no_registros == false)
            {
                string[] cc = new string[1];
                cc[0] = (validar_lineas == false) ? "appmanagement@gbm.net" : "";

                mail.SendHTMLMail("A continuacion se adjunta el archivo con el conglomerado de las nuevas ordenes publicadas en Convenio Marco a favor de GBM de Panama", new string[] { "kvanegas@gbm.net" }, "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + fecha_file, cc, adjunto);

            }
            if (resp_add_sql == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] { "appmanagement@gbm.net" }, "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + fecha_file, cc, adjunto);
            }

            #region enviar reporte Cisco
            if (cisco_add == true)
            {
                string subject = "Reporte diario de ordenes nuevas publicadas por Convenio Marco - " + fecha_file;
                string body = "A continuacion se adjunta el archivo con el conglomerado de las nuevas ordenes de Networking publicadas en Convenio Marco a favor de GBM de Panama";
                string[] adj = { nombre_fila_Cisco };
                foreach (KeyValuePair<string, string> pair in AMs)
                {
                    string email = pair.Value.ToString() + "@gbm.net";
                    CopyCC[cont_am] = email;
                    cont_am++;
                    Array.Resize(ref CopyCC, CopyCC.Length + 1);

                }
                Array.Resize(ref CopyCC, CopyCC.Length - 1);
                string[] senders = { "mmedina@gbm.net", "jleal@gbm.net" };
                mail.SendHTMLMail(body, senders, subject, CopyCC, adj);
            }

            #endregion


            return validar_lineas.ToString();
        }
        /// <summary>
        /// Metodo correos aprobados o publicados
        /// </summary>
        /// <param name="body">el body del correo electronico del gobierno de Panamá (ver carpta Solicitudes convenio gbmpa)</param>
        /// <returns></returns>
        public bool AgreementPaEmail(string body)
        {
            #region variable privadas
            bool validar_lineas = true;
            string sectorf = "";
            string body_clean = ""; string[] link;
            string subject = root.Subject;
            string convenio = "BIENES INFORMÁTICOS, REDES Y COMUNICACIONES";
            string resp_sql = "";
            bool resp_add_sql = true;
            string[] adjunto = new string[1];
            int cont_adj = 0;
            //string[] valores = new string[1];
            //int cont_val = 0;
            Dictionary<string, string> valores = new Dictionary<string, string>();
            Dictionary<string, string> campos_opp = new Dictionary<string, string>();
            string tipo = "";
            #endregion
            #region aprobada o publicada
            tipo = (subject.Contains("Orden de Compra Aprobada")) ? "aprobada" : "publicada";
            #endregion
            #region extrae info del body
            console.WriteLine(" Extrayendo info del Body y Subject");
            Regex reg;
            reg = new Regex("[*'\"_&+^><@]");
            body_clean = reg.Replace(body, string.Empty);
            #endregion
            if (tipo == "aprobada")
            {
                int limite = 0;
                //extrae entidad
                string[] stringSeparators0 = new string[] { "Por este medio les deseamos informar que la entidad " };
                link = body_clean.Split(stringSeparators0, StringSplitOptions.None);
                string entidad = link[1].ToString().Trim();
                limite = entidad.IndexOf(" ha generado una Orden de compra");
                if (entidad.Length >= limite)
                { entidad = entidad.Substring(0, limite).Trim(); }

                //registro unico
                string[] stringSeparators1 = new string[] { "Trámite No. " };
                link = body_clean.Split(stringSeparators1, StringSplitOptions.None);
                string registro_unico = link[1].ToString().Trim();
                limite = registro_unico.IndexOf(" a través del Catálogo Electrónico");
                if (registro_unico.Length >= limite)
                { registro_unico = registro_unico.Substring(0, limite).Trim(); }

                //producto y cantidad
                string[] stringSeparators = new string[] { "Cantidad\t\r\n" };
                link = body.Split(stringSeparators, StringSplitOptions.None);
                string tabla = "";
                try
                {
                    tabla = link[1].ToString().Trim();
                }
                catch (Exception ex)
                {
                    stringSeparators = new string[] { "Cantidad\t \r\n" };
                    link = body.Split(stringSeparators, StringSplitOptions.None);
                    tabla = link[1].ToString().Trim();
                }

                limite = tabla.IndexOf("Referente:");
                if (tabla.Length >= limite)
                { tabla = tabla.Substring(0, limite).Trim(); }

                string[] stringSeparators3 = new string[] { "\t" };
                link = tabla.Split(stringSeparators3, StringSplitOptions.None);
                string fianza = "No Aplica";
                string estado_gbm = "Pendiente de Procesar";
                string estado_orden = "En Refrendo";

                //fecha de registro
                string fecha_registro = root.ReceivedTime.ToString("yyyy-MM-dd");
                float sub_total = 0;
                for (int i = 0; i < link.Length; i++)
                {
                    string producto = link[i + 1].ToString().Trim();
                    string cantidad = link[i + 2].ToString().Trim();
                    //extrae los dias habiles de acuerdo a la cantidad
                    int dias_h = dias_habiles(int.Parse(cantidad));

                    //extrae la marca del producto
                    string product_id = producto.Substring(0, 2);
                    if (product_id.Contains("-"))
                    { product_id = producto.Substring(0, 1); }

                    //buscar en tabla#1
                    //string marca = marca_product(product_id);
                    string[] marca_array = lpsql.MarcaProduct(product_id);
                    string marca = marca_array[0];
                    //total= cantidad * precio y subtotal = suma de totales
                    string precio = marca_array[1];
                    string total = "0";
                    try
                    {
                        precio = precio.Replace(".", ",");
                        float total_prod = (float.Parse(precio) * float.Parse(cantidad));
                        total = total_prod.ToString();
                        total = total.Replace(",", ".");
                        sub_total += total_prod;

                    }
                    catch (Exception ex)
                    {

                    }

                    //buscar sector 
                    string[] entidad_info = lpsql.DetermineSector(entidad);
                    string sector = entidad_info[0];

                    string fecha_publi = "0001-01-01";

                    console.WriteLine("  Agregando información a la base de datos");
                    bool add_sql = lpsql.InfoSqlAddAprobadas(convenio, entidad, producto, cantidad, total, marca, registro_unico, fecha_registro, fecha_publi, fianza, estado_gbm, estado_orden, dias_h, sector);
                    log.LogDeCambios("Creacion", roots.BDProcess, "Ventas Panama", "Agregar info a convenio GBPA - Orden Aprobada", registro_unico + "-" + producto, add_sql.ToString());
                    if (add_sql == false)
                    {
                        //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                        //para enviarlo por email y agregarla
                        resp_sql = resp_sql + registro_unico + " - " + producto + "<br>";
                        resp_add_sql = false;
                    }
                    i++; i++;
                }
                if (sub_total != 0)
                {
                    resp_add_sql = lpsql.UpdateSubtotal(registro_unico, sub_total.ToString());
                }
                if (resp_add_sql == false)
                {
                    string[] cc = { "appmanagement@gbm.net", "dmeza@gbm.net" };
                    mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] { "appmanagement@gbm.net" }, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + registro_unico, cc);
                }
                else
                {
                    //todo salio bien, se envia notificación.
                    JArray j_copias = JArray.Parse(lpsql.EmailCpa("LICPA"));

                    for (int i = 0; i < j_copias.Count; i++)
                    {
                        string email = j_copias[i]["email"].ToString();

                        wt.SendNotification(email, "Nueva Orden de Compra en Refrendo", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + registro_unico + "**, la cual se encuentra En Refrendo.<br><br>Haga click en el siguiente enlace: <a href=\"https://databot.gbm.net/lp\" > Orden de Compra - Portal GBM</a> para ver el documento");
                    }

                }


            }
            else //publicada
            {
                string brand_opp = "";
                string registro_opp = "";
                string cantidad_opp = "";
                //extrae entidad
                string[] stringSeparators0 = new string[] { "Por este medio les deseamos informar que el (la) " };
                link = body_clean.Split(stringSeparators0, StringSplitOptions.None);
                string entidad = link[1].ToString().Trim();
                int limite = entidad.IndexOf("ha generado la orden de compra") - 1;
                if (entidad.Length >= limite)
                { entidad = entidad.Substring(0, limite).Trim(); }
                valores.Add("ENTIDAD", entidad);

                //registro unico 
                string[] stringSeparators1 = new string[] { "Orden de Compra " };
                link = subject.Split(stringSeparators1, StringSplitOptions.None);
                string registro_unico = link[1].ToString().Trim();
                valores.Add("REGISTRO_UNICO_DE_PEDIDO", registro_unico);

                //text opp
                string[] stringSeparators3 = new string[] { "-RC" };
                link = registro_unico.Split(stringSeparators3, StringSplitOptions.None);
                registro_opp = link[1].ToString().Trim();
                registro_opp = "-RC" + registro_opp;

                //extrae link href para ir a la orden en web
                string[] stringSeparators2 = new string[] { "<" };
                link = body.Split(stringSeparators2, StringSplitOptions.None);
                string link_po = link[1].ToString().Trim();
                limite = link_po.IndexOf("> ,");
                if (link_po.Length >= limite)
                { link_po = link_po.Substring(0, limite).Trim(); }
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

                IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

                try
                {
                    #region Ingreso al website
                    console.WriteLine("  Ingresando al website");
                    try
                    { chrome.Navigate().GoToUrl(link_po); }
                    catch (Exception)
                    { chrome.Navigate().GoToUrl(link_po); }
                    IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
                    #endregion

                    chrome.Manage().Cookies.DeleteAllCookies();

                    console.WriteLine("  Extrayendo información de la página web");

                    #region extraer información general  
                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblLugarEntrega']"))); }
                    catch { }

                    string lugar = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblLugarEntrega']")).Text;
                    string funcionario = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblNameFuncionario']")).Text;
                    string telefono = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblCellFuncionario']")).Text;
                    string email = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblEmailFuncionario']")).Text;
                    string provincia = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblProvincia']")).Text;

                    valores.Add("LUGAR_DE_ENTREGA", lugar); valores.Add("CONTACTO_DE_ENTREGA", funcionario);
                    valores.Add("TELEFONO", telefono); valores.Add("EMAIL", email);
                    valores.Add("PROVINCIA", provincia);

                    try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']"))); }
                    catch { }
                    string po_url = chrome.Url;
                    valores.Add("LINK_AL_DOCUMENTO", po_url);
                    string fecha_registro = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']")).Text;
                    fecha_registro = fecha_registro.Remove(fecha_registro.Length - 5);
                    DateTime RDate = Convert.ToDateTime(fecha_registro);
                    valores.Add("FECHA_DE_REGISTRO", RDate.Date.ToString("yyyy-MM-dd"));
                    fecha_registro = RDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                    fecha_registro = "'" + fecha_registro;

                    string unidad_solicitante = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblUnidadSolicitante']")).Text,
                    contactocuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblRegistradoPor']")).Text,
                    emailcuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblEmail']")).Text;
                    valores.Add("UNIDAD_SOLICITANTE", unidad_solicitante); valores.Add("CONTACTO_CUENTA", contactocuenta); valores.Add("EMAIL_CUENTA", emailcuenta);
                    #endregion

                    #region Extraer información de la entidad                    
                    //extrae sector, id cliente, id contacto, usuario
                    string sector = "", cliente_opp = "", contacto_opp = "", sales_rep = "";

                    //si la entidad es alguna de estas se determina el cliente con base a la unidad solicitante
                    if (entidad == "MINISTERIO DE EDUCACION" || entidad == "CAJA DE SEGURO SOCIAL" || entidad == "MUNICIPIO DE ANTÓN")
                    {
                        string[] entidad_info = lpsql.DetermineSector(unidad_solicitante);
                        sector = entidad_info[0];
                        cliente_opp = entidad_info[1];
                        contacto_opp = entidad_info[2];
                        sales_rep = entidad_info[3];

                        if (cliente_opp == "0010067544") //no encontro con la unidad entonces busca por entidad madre
                        {

                            string[] entidad_info2 = lpsql.DetermineSector(entidad);
                            sector = entidad_info2[0];
                            cliente_opp = entidad_info2[1];
                            contacto_opp = entidad_info2[2];
                            sales_rep = entidad_info2[3];

                        }
                    }
                    else if (entidad == "MINISTERIO DE GOBIERNO" || entidad == "MINISTERIO DE SALUD" || entidad == "Ministerio de Seguridad Pública" || entidad == "MINISTERIO DE LA PRESIDENCIA") //email, unidad, entidad
                    {
                        try
                        {
                            MailAddress address = new MailAddress(emailcuenta);
                            string host = address.Host.ToString();
                            //string[] domain = host.Split('.');
                            //host = domain[0].ToString();
                            string[] entidad_info = lpsql.DetermineSector(host);
                            sector = entidad_info[0];
                            cliente_opp = entidad_info[1];
                            contacto_opp = entidad_info[2];
                            sales_rep = entidad_info[3];
                        }
                        catch (Exception)
                        {
                            cliente_opp = "0010067544";
                        }

                        if (cliente_opp == "0010067544") //no encontro con el email entonces busca por unidad
                        {

                            string[] entidad_info = lpsql.DetermineSector(unidad_solicitante);
                            sector = entidad_info[0];
                            cliente_opp = entidad_info[1];
                            contacto_opp = entidad_info[2];
                            sales_rep = entidad_info[3];

                            if (cliente_opp == "0010067544") //no encontro con la unidad entonces busca por entidad madre
                            {

                                string[] entidad_info2 = lpsql.DetermineSector(entidad);
                                sector = entidad_info2[0];
                                cliente_opp = entidad_info2[1];
                                contacto_opp = entidad_info2[2];
                                sales_rep = entidad_info2[3];

                            }

                        }
                    }
                    else
                    {
                        string[] entidad_info = lpsql.DetermineSector(entidad);
                        sector = entidad_info[0];
                        cliente_opp = entidad_info[1];
                        contacto_opp = entidad_info[2];
                        sales_rep = entidad_info[3];

                    }
                    #endregion

                    #region extraer información al final

                    try
                    { new Actions(chrome).MoveToElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ibtnExportarPDF']"))).Perform(); }
                    catch (Exception) { }

                    string fecha_doc = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvAdjuntos_ctl02_lblFecha']")).Text;
                    fecha_doc = fecha_doc.Remove(fecha_doc.Length - 5);
                    DateTime oDate = Convert.ToDateTime(fecha_doc);
                    valores.Add("FECHA_DE_PUBLICACION", oDate.Date.ToString("yyyy-MM-dd"));
                    string mes2 = oDate.Month.ToString();
                    string ano = oDate.Year.ToString();
                    fecha_doc = mes2 + "-" + ano;


                    IWebElement pdf = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvAdjuntos_ctl02_lbtUrl']"));
                    string href_pdf = pdf.GetAttribute("onclick").ToString();
                    string ext = "";
                    ext = get_ext(href_pdf);
                    string pdf_full_name = "";
                    string pdf_name = "";
                    if (ext != "")
                    {
                        pdf_full_name = root.FilesDownloadPath + "\\" + pdf.Text + ext;
                        pdf_name = pdf.Text + ext;
                    }
                    else
                    {
                        pdf_full_name = root.FilesDownloadPath + "\\" + pdf.Text + ".pdf";
                        pdf_name = pdf.Text + ".pdf";
                    }
                    valores.Add("NOMBRE_DEL_ADJUNTO", pdf_name);

                    #endregion

                    #region extraer subtotal        
                    string producto = "";
                    string total = "";
                    string cantidad = "";

                    producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl02_lblNombre']")).Text;
                    int contador_subtotal = 2;
                    while (producto != "")
                    {
                        try
                        {
                            producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador_subtotal + "_lblNombre']")).Text;
                            contador_subtotal++;
                        }
                        catch (Exception)
                        { producto = ""; }

                    }
                    //*[@id="ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl04_lblSubtotal"] 
                    string sub_total = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador_subtotal + "_lblSubtotal']")).Text;
                    valores.Add("SUB_TOTAL_ORDEN", sub_total);

                    #endregion

                    #region extraer productos y actualizar base de datos

                    producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl02_lblNombre']")).Text;
                    int contador = 2;
                    while (producto != "")
                    {
                        try
                        {
                            producto = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador + "_lblNombre']")).Text;
                            cantidad = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador + "_lblCantidad']")).Text;
                            total = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvPedidoDetalle_ctl0" + contador + "_lbltotal']")).Text;
                            //Se agrega diferente el valor al dictionary debido a que cambia
                            valores["PRODUCTO_SERVICIO"] = producto; valores["CANTIDAD"] = cantidad; ; valores["TOTAL_DEL_PRODUCTO"] = total;
                            //extrae los dias habiles de acuerdo a la cantidad
                            int dias_h = dias_habiles(int.Parse(cantidad));
                            valores["DIAS_ENTREGA"] = dias_h.ToString();
                            //extrae la marca del producto
                            string product_id = producto.Substring(0, 2);
                            if (product_id.Contains("-"))
                            { product_id = producto.Substring(0, 1); }



                            //buscar en tabla#1
                            //string marca = marca_product(product_id);
                            string[] marca_array = lpsql.MarcaProduct(product_id);
                            string marca = marca_array[0];
                            valores["MARCA"] = marca;
                            if (marca != "")
                            { brand_opp = (marca == "TrippLite") ? "U" : marca.Substring(0, 1); }

                            //cantidad para text opp
                            cantidad_opp += "," + "R" + brand_opp + product_id + "-" + cantidad;

                            //buscar sector 

                            valores["SECTOR"] = sector;
                            sectorf = sector;

                            //fecha maxima de entrega
                            DateTime fecha_max_entrega = DateTime.MinValue;
                            string fecha_max = "";
                            if (fecha_doc != "")
                            {
                                int dias_ent = dias_h + 2;
                                //Excel.IWorksheetFunction workday = (Excel.WorksheetFunction)Excel.WorksheetFunction.WorkDay(fecha_documento, dias_h,"");
                                fecha_max_entrega = AddWorkdays(oDate, dias_ent);
                                if (fecha_max_entrega != oDate)
                                {
                                    fecha_max = fecha_max_entrega.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    fecha_max = "'" + fecha_max;
                                }

                            }
                            valores["FECHA_MAXIMA_ENTREGA"] = fecha_max_entrega.Date.ToString("yyyy-MM-dd");
                            //dias faltantes
                            double dias_faltantes = 0;
                            if (fecha_max != "")
                            {

                                try
                                {
                                    DateTime hoy = DateTime.Now;
                                    dias_faltantes = (fecha_max_entrega.Date - hoy.Date).TotalDays;
                                    dias_faltantes = Math.Ceiling(dias_faltantes);
                                }
                                catch (Exception)
                                { }

                            }

                            valores["DIAS_FALTANTES"] = dias_faltantes.ToString();

                            valores["FIANZA_CUMPLIMIENTO"] = "No Aplica";
                            valores["ORDEN_COMPRA"] = "";
                            valores["OPORTUNIDAD"] = "";
                            valores["QUOTE"] = "";
                            valores["TIPO_PEDIDO"] = "";
                            valores["SALES_ORDER"] = "";
                            valores["ESTADO_GBM"] = "Pendiente de Procesar";
                            valores["ESTATUS_DE_ORDEN"] = "Refrendado";
                            valores["MONTO_MULTA"] = "";
                            valores["FORECAST"] = DateTime.MinValue.Date.ToString("yyyy-MM-dd");
                            valores["CONVENIO"] = convenio;
                            valores["CONFIRMACION_ORDEN"] = "";
                            valores["COMENTARIOS"] = "";
                            valores["FECHA_REAL_ENTREGA"] = "";
                            valores["TIPO_FORECAST"] = "E2E";
                            valores["VENDOR_ORDER"] = "";
                            //agregar información a la base de datos
                            //true todo bien, false significa que dio un error
                            console.WriteLine("  Agregando información a la base de datos");

                            bool add_sql = lpsql.InfoSqlAddGbmV2(valores);

                            //bool add_sql = info_sql_add_gbm(convenio, entidad, producto, cantidad, marca, total, sub_total, orden_compra, registro_unico, RDate.Date, oDate.Date, fianza, opp, quote, tipo_pedido, sales_order, estado_gbm, estado_orden, dias_h, fecha_max_entrega.Date, dias_faltantes, forecast, provincia, lugar, funcionario, telefono, email, monto_multa, pdf_name, po_url);
                            log.LogDeCambios("Creacion", roots.BDProcess, "Ventas Panama", "Agregar info a convenio GBPA - Orden Publicada", registro_unico + "-" + producto, add_sql.ToString());
                            if (add_sql == false)
                            {
                                //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                //para enviarlo por email y agregarla
                                resp_sql = resp_sql + po_url + "<br>";
                                resp_add_sql = false;
                            }

                        }
                        catch (Exception)
                        { producto = ""; }
                        contador++;
                    }
                    #endregion

                    #region descargable pdf
                    try
                    {
                        if (pdf_full_name != "")
                        {
                            adjunto[cont_adj] = pdf_full_name;
                            cont_adj++;
                            Array.Resize(ref adjunto, adjunto.Length + 1);
                        }

                        pdf.Click();
                        for (var x = 0; x < 60; x++)
                        {
                            if (File.Exists(pdf_full_name)) { break; }
                            System.Threading.Thread.Sleep(1000);
                        }
                        //System.Threading.Thread.Sleep(7000);
                        chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                        System.Threading.Thread.Sleep(1000);
                        chrome.Close();

                    }
                    catch (Exception ex)
                    {
                        console.WriteLine("Error al descargar adjunto: " + ex.Message);
                        try
                        {
                            chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                            System.Threading.Thread.Sleep(1000);
                            chrome.Close();
                        }
                        catch (Exception)
                        { }
                    }
                    System.Threading.Thread.Sleep(1000);
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]); //regreso a la pagina principal
                    #endregion

                    #region crear opp
                    string opp_actual = lpsql.OppExist(registro_unico);
                    if (String.IsNullOrEmpty(opp_actual))
                    {
                        campos_opp["tipo"] = "ZOPS"; //Standard

                        string opp_text = "CM GOB-";
                        if (cantidad_opp != "")
                        {
                            if (cantidad_opp.Substring(0, 1) == ",")
                            {
                                cantidad_opp = cantidad_opp.Substring(1, cantidad_opp.Length - 1);
                            }
                        }

                        string opp_descripcion = opp_text + cantidad_opp + registro_opp;
                        if (opp_descripcion.Length > 40)
                        {
                            int cant_caract = (opp_text.Length + registro_opp.Length);
                            //opp descip = 40 caract
                            int cantf = 40 - cant_caract;
                            if (cantidad_opp.Length > cantf)
                            {
                                cantidad_opp = cantidad_opp.Substring(0, cantf);
                            }
                            opp_descripcion = opp_text + cantidad_opp + registro_opp;

                        }
                        campos_opp["descripcion"] = opp_descripcion; //"CM GOB-RL3-1-RC-003598"; //opp_text;


                        campos_opp["fecha_inicio"] = DateTime.Now.Date.ToString("yyyy-MM-dd");
                        campos_opp["Fecha_Final"] = DateTime.Now.AddDays(5).Date.ToString("yyyy-MM-dd");

                        campos_opp["Ciclo"] = "Y3"; //quotation
                        campos_opp["Origen"] = "Y08"; //Public Bid - licitaciones

                        string grupo_opp = "";
                        grupo_opp = (cliente_opp == "0010067544") ? grupo_opp = "0001" : grupo_opp = "0002"; //0001 new 0002 exist
                        campos_opp["grupo_opp"] = grupo_opp;

                        campos_opp["Cliente"] = cliente_opp; // "0010004721"; 
                        campos_opp["Contacto"] = contacto_opp;  // "0070012034";// contacto_opp;
                        campos_opp["Usuario"] = sales_rep;  //"AA70000134"; // sales_rep;

                        campos_opp["OrgVentas"] = "O 50000142"; //Panama
                        campos_opp["OrgServicios"] = "50003612"; //Panama Service Delivery

                        string id_opp = crear_opp(campos_opp);

                        if (!String.IsNullOrEmpty(id_opp))
                        {
                            if (id_opp.Contains("Error"))
                            {
                                validar_lineas = false;
                                //resp_add_sql = false;
                                resp_sql = "Error al crear la opp";
                            }
                            else
                            {
                                log.LogDeCambios("Creacion", "Creacion de Oportunidad", "Ventas Panama", "Oportunidad de convenio GBPA - Orden Publicada", registro_unico, id_opp);
                                if (cliente_opp == "0010067544") //no encontro con la unidad entonces busca por entidad
                                {
                                    JArray j_copias = JArray.Parse(lpsql.EmailCpa("LICPA"));
                                    string jmail = j_copias[0]["email"].ToString();
                                    wt.SendNotification(jmail, "Nueva Oportunidad Creada", "Se le notifica que se ha creado una nueva oportunidad con el id: **" + id_opp + "** cuya entidad/Unidad Solicitante/email: **" + entidad + "** no se encuentra creada en SAP, <br><br> Haga click en el siguiente enlace: <a href =\"https://databot.gbm.net\">Portal de Datos Maestros</a> para crear el cliente y una vez creadoo agreguelo en la <a href=\"https://databot.gbm.net/lp/home/entidades\">Tabla de Mantenimiento de entidades</a>");

                                }
                                Dictionary<string, string> oppo = new Dictionary<string, string>();
                                oppo.Add("OPORTUNIDAD", id_opp);
                                //bool opp_update = update_reg("", "", "", "", "", "", registro: registro_unico, opp: id_opp); //modificar
                                bool opp_update = lpsql.UpdateRegister(registro_unico, oppo, 1);


                                if (opp_update == false)
                                {
                                    JArray j_copias = JArray.Parse(lpsql.EmailCpa("LICPA"));
                                    string jmail = j_copias[0]["email"].ToString();
                                    wt.SendNotification(jmail, "Nueva Oportunidad Creada", "Se le notifica que se ha creado una nueva oportunidad con el id: **" + id_opp + "** del registro: **" + registro_unico + "** no se pudo actualizar en la base de datos.");

                                }
                            }

                        }
                    }
                    #endregion


                }
                catch (Exception ex)
                {
                    //dio error sacando la info con selenium
                    validar_lineas = false;
                }

                console.WriteLine("  Finalizando");
                try { chrome.Close(); } catch (Exception) { }

                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);

                #region subir archivos FTP
                Array.Resize(ref adjunto, adjunto.Length - 1);
                for (int i = 0; i < adjunto.Length; i++)
                {
                    //bool subir_files = db2.UploadFtp("ftp://10.7.60.72/licitaciones_files/", "gbmadmin", cred.password_server_web, adjunto[i].ToString());
                    //bool subir_files = wb2.upload_ftp("https://databot.gbm.net/lp/home/assets/adjuntos/", "databot", cred.pass_db1, adjunto[i].ToString());
                    sharep.UploadFileToSharePoint("https://gbmcorp.sharepoint.com/sites/licitaciones_panama", adjunto[i].ToString());
                }
                #endregion

                //dio error en la opp
                if (validar_lineas == false)
                {
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] { "appmanagement@gbm.net" }, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + registro_unico, cc);

                }

                if (resp_add_sql == false)
                {
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] { "appmanagement@gbm.net" }, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + registro_unico, cc);

                }
                else
                {
                    //todo salio bien, se envia notificación.
                    //try
                    //{ PushN.GenerarNotificacion(root.BD_Proceso, "AColina@gbm.net", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + registro_unico + "**, la cual se encuentra Refrendado<br><br>Haga click en el siguiente enlace: <a href=\"https://databot.gbm.net/lp\">Orden de Compra - Portal GBM</a> para ver el documento"); }
                    //catch (Exception ex)
                    //{ }
                    //try
                    //{ PushN.GenerarNotificacion(root.BD_Proceso, "aleayala@gbm.net", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + registro_unico + "**, la cual se encuentra Refrendado<br><br>Haga click en el siguiente enlace: <a href=\"https://databot.gbm.net/lp\">Orden de Compra - Portal GBM</a> para ver el documento"); }
                    //catch (Exception ex)
                    //{ }

                    JArray j_copias = JArray.Parse(lpsql.EmailCpa("LICPA"));

                    for (int i = 0; i < j_copias.Count; i++)
                    {
                        string email = j_copias[i]["email"].ToString();

                        wt.SendNotification(email, "Nueva Orden de Compra Refrendada", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + registro_unico + "**, la cual se encuentra Refrendado<br><br>Haga click en el siguiente enlace: <a href=\"https://databot.gbm.net/lp\">Orden de Compra - Portal GBM</a> para ver el documento");
                    }

                    if (sectorf == "BF")
                    {
                        JArray JBF = JArray.Parse(lpsql.EmailCpa("LPBF"));
                        string bfemail = JBF[0]["email"].ToString();
                        wt.SendNotification(bfemail, "Nueva Orden de Compra Refrendada", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + registro_unico + "**, la cual se encuentra Refrendado<br><br>Haga click en el siguiente enlace: <a href=\"https://databot.gbm.net/lp\">Orden de Compra - Portal GBM</a> para ver el documento");
                    }

                }
            } //else de tipo publicada

            return resp_add_sql;
        }
        /// <summary>
        /// Actualizar precios todos los miercoles
        /// </summary>
        /// <returns></returns>
        public string AgreementGetPrice()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            Int32 fila_producto = -1;
            Int32 fila_vendor = -1;
            string respuesta = "";
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
            catch (Exception ex)
            { respuesta = ex.Message; }
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("http://catalogo.panamacompra.gob.pa/forms/Publico/documentosPrecioProveedor.aspx");

                chrome.Manage().Cookies.DeleteAllCookies();

                console.WriteLine("  Extrayendo información de la página web");

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlFecha0']"))); }
                catch { }

                SelectElement fecha_select = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlFecha0']")));
                System.Threading.Thread.Sleep(1000);
                fecha_select.SelectByValue(today);
                System.Threading.Thread.Sleep(3000);

                SelectElement convenio = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlConvenioFiltro']")));
                System.Threading.Thread.Sleep(1000);
                convenio.SelectByValue("107"); //Bienes informaticos
                System.Threading.Thread.Sleep(3000);

                SelectElement region = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlRegion']")));
                System.Threading.Thread.Sleep(1000);
                region.SelectByValue("33"); //Provincia de panama
                System.Threading.Thread.Sleep(3000);

                chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_btnPdf']")).Click(); //BUSCAR
                console.WriteLine("Buscar...");
                System.Threading.Thread.Sleep(3000);

                chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                System.Threading.Thread.Sleep(1000);
                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[1]/tbody/tr[2]/td[1]/div"))); }
                catch { }
                console.WriteLine("  Descargando");
                chrome.FindElement(By.XPath("//*[@id='toolbar']/table/tbody/tr/td/input[3]")).Click(); //descargar
                System.Threading.Thread.Sleep(1000);
                string ruta = root.FilesDownloadPath + "\\" + "ReportePrecio.xls";
                for (var x = 0; x < 40; x++)
                {
                    if (File.Exists(ruta)) { break; }
                    System.Threading.Thread.Sleep(1000);
                }
                try
                {
                    chrome.SwitchTo().Window(chrome.WindowHandles[2]);
                    chrome.Close();
                    chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                    chrome.Close();
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    chrome.Close();
                }
                catch (Exception)
                {
                    chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                    chrome.Close();
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    chrome.Close();
                }
                proc.KillProcess("chromedriver", true);
                #region Abrir excel y extraer precios

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                console.WriteLine("  Abriendo excel");
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(ruta);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
                int rows = xlWorkSheet.UsedRange.Rows.Count;

                SqlConnection myConn = new SqlConnection();
                string sql_select = "";
                string sql_update = "";
                DataTable mytable = new DataTable();
                string precio = "";
                string producto = "";

                try
                {
                    #region Connection DB     
                    sql_select = "select * from productos_compras_gbpa";
                    //mytable = crud.Select("Databot", sql_select, "ventas");
                    #endregion

                    if (mytable.Rows.Count > 0)
                    {
                        for (int i = 0; i < mytable.Rows.Count; i++)
                        {
                            fila_producto = 0; int contador = 0; precio = ""; producto = "";
                            producto = mytable.Rows[i][0].ToString() + " - ";
                            console.WriteLine("  Producto: " + mytable.Rows[i][1].ToString());
                            Excel.Range rango = xlWorkSheet.Range["A1:H" + rows];
                            Excel.Range valor_encontrado = rango.Find(producto);
                            if (valor_encontrado != null)
                            {
                                fila_producto = valor_encontrado.Row + 2; //+2 para quitar el titulo del prooducto
                                string vendor = xlWorkSheet.Cells[fila_producto, 1].text.ToString().Trim();
                                while (vendor != "Proveedor")
                                {
                                    if (vendor == "GBM de Panamá, S.A")
                                    {
                                        precio = xlWorkSheet.Cells[fila_producto + contador, 6].text.ToString().Trim();
                                        break;
                                    }
                                    contador++;
                                    vendor = xlWorkSheet.Cells[fila_producto + contador, 1].text.ToString().Trim();
                                }
                                try { precio = precio.Replace("$", ""); precio = precio.Replace(".", ""); precio = precio.Replace(",", "."); } catch (Exception) { }
                                if (precio != "")
                                {
                                    try
                                    {
                                        console.WriteLine("  Actualizando...");
                                        int id = int.Parse(mytable.Rows[i][0].ToString());
                                        sql_update = "UPDATE `productos_compras_gbpa` SET `precio`='" + precio + "' WHERE `id`=" + id;
                                        //crud.Update("Databot", sql_update, "ventas");



                                        log.LogDeCambios("Modificacion", roots.BDProcess, "Ventas Panama", "Actualizar Precio de Producto", id.ToString(), precio.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        respuesta = ex.Message;
                                    }

                                }
                            }
                            else
                            {

                            }

                        }
                    }


                    console.WriteLine("  Cerrando excel.");
                    xlApp.DisplayAlerts = false;
                    xlWorkBook.Close();
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                }
                catch (Exception ex)
                {
                    respuesta = ex.Message;
                }

                #endregion
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                respuesta = ex.Message;
            }

            proc.KillProcess("EXCEL", true);
            return respuesta;
        }
        /// <summary>
        /// Extraer la información de cotizaciones programadas y abiertas
        /// </summary>
        /// <returns></returns>
        public string QuickQuote()
        {
            //crea una lista con todas las palabras claves de la base de datos.
            List<string> words = lpsql.KeyWord("");
            string[] CotiAll = lpsql.AllQuotes();
            Dictionary<string, string> entidad_info = lpsql.getAllEntity();

            Dictionary<string, string> campos_coti = new Dictionary<string, string>();
            Dictionary<string, string> adj_names = new Dictionary<string, string>();
            string[] adjunto = new string[1];
            string[] AMarrays = new string[1];
            int cont_adj = 0;
            int cont_am = 0;
            int cont = 1;
            string resp_sql = "";
            string resp_AA70000134 = "";
            string resp_AA00070471 = "";
            bool resp_add_sql = true;
            bool validar_lineas = true;
            Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll;
            Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone;
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
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            Excel.Range xlRango;
            //Excel.Range xlRangoCopy;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(root.quickQuoteReport);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            int cant_row = xlWorkSheet.UsedRange.Rows.Count;
            if (cant_row > 1)
            {
                xlRango = xlWorkSheet.get_Range("A2", "A1048576");
                xlRango.EntireRow.Delete(Type.Missing);
            }
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#!/busquedaAvanzada?BusquedaTipos=True&IdTipoBusqueda=53&estado=51&title=Cotizaciones%20en%20L%C3%ADnea");
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl("https://www.panamacompra.gob.pa/Inicio/#!/busquedaAvanzada?BusquedaTipos=True&IdTipoBusqueda=53&estado=51&title=Cotizaciones%20en%20L%C3%ADnea"); }
            //System.Threading.Thread.Sleep(5000);
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

            System.Threading.Thread.Sleep(2000);
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[2]/label[4]"))); }
            catch { }
            chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[2]/label[4]")).Click(); //click en el boton Programadas
            //System.Threading.Thread.Sleep(8000);

            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/center/pre/small[2]"))); }
            catch { }

            //*[@id="toTopBA"]/h5/b
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='toTopBA']/h5/b"))); }
            catch { }
            string cant_filas_prog = "";
            try
            { cant_filas_prog = chrome.FindElement(By.XPath("//*[@id='toTopBA']/h5/b")).Text; }
            catch (Exception)
            { cant_filas_prog = ""; }


            if (cant_filas_prog == "Se encontraron 0 Cotizaciones Programadas")
            {
                //no hay cotizaciones programadas
            }
            else //programadas
            {
                //si hay cotizaciones programadas
                int mas = 0;
                if (cant_filas_prog.Contains("+"))
                {
                    do
                    {
                        IWebElement pagination_last = chrome.FindElement(By.ClassName("pagination-last"));
                        var pag_last = pagination_last.FindElement(By.TagName("a"));
                        pag_last.Click();
                        System.Threading.Thread.Sleep(1500);
                        cant_filas_prog = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                        mas++;
                    } while (cant_filas_prog.Contains("+"));



                    if (mas > 0)
                    {
                        IWebElement pagination_first = chrome.FindElement(By.ClassName("pagination-first"));
                        var pag_first = pagination_first.FindElement(By.TagName("a"));
                        pag_first.Click();
                        System.Threading.Thread.Sleep(1000);
                        cant_filas_prog = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
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

                int e = 0;
                //FOR por cada página de la tabla principal
                for (int i = 1; i <= pag_row2; i++)
                {
                    int rows = xlWorkSheet.UsedRange.Rows.Count + 1;
                    try
                    {

                        //next en pagination
                        if (i != 1)
                        {
                            console.WriteLine("     Siguiente pagina");
                            IWebElement pagination_next = chrome.FindElement(By.ClassName("pagination-next"));
                            var pag_next = pagination_next.FindElement(By.TagName("a"));
                            pag_next.Click();
                        }
                        IWebElement tableElement = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[4]/table"));
                        IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));

                        //LOOP por cada fila de la tabla principal de cotizaciones programadas
                        foreach (IWebElement element in trCollection)
                        {
                            try
                            {

                                IList<IWebElement> tdCollection;
                                tdCollection = element.FindElements(By.TagName("td"));
                                if (tdCollection.Count > 0)
                                {
                                    string num_coti = tdCollection[1].Text; //Registro único de pedido
                                                                            //verifica si ya el registro esta en la base de datos
                                    if (!CotiAll.Any(num_coti.Contains))
                                    {
                                        //if (!lpsql.cotiexiste(num_coti))
                                        //{
                                        string descripcion = tdCollection[2].Text;
                                        descripcion = cleantext(descripcion);

                                        string id = tdCollection[0].Text;
                                        string entidad = tdCollection[3].Text;

                                        string fecha_publicacion = tdCollection[5].Text;

                                        IWebElement link = tdCollection[1];
                                        var linkhref = link.FindElement(By.TagName("a"));
                                        string href = linkhref.GetAttribute("href");
                                        //new Actions(chrome).MoveToElement(link).Perform();
                                        console.WriteLine("    Click en el Codigo Unico Pedido: #" + tdCollection[0].Text + " " + num_coti);

                                        IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                                        js.ExecuteScript("window.open('{0}', '_blank');");
                                        chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab
                                        chrome.Navigate().GoToUrl(href); //ir al link de la cotizacion
                                        System.Threading.Thread.Sleep(2000);

                                        //hacer todo lo que tenga que hacer en la nueva hoja-----
                                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='elementToPDF']/div[1]/table/tbody/tr[1]/td[2]"))); }
                                        catch { }

                                        string coti_url = chrome.Url;
                                        entidad = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]")).Text;
                                        string dependencia = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]")).Text;
                                        string unidad_compra = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[2]")).Text;
                                        num_coti = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[1]/table/tbody/tr[1]/td[2]")).Text;
                                        //mover hacia abajo //*[@id="elementToPDF"]/div[2]/div[3]/div[2]/table/tbody/tr[15]/td[2]
                                        string fecha_presentacion = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[6]/td[2]")).Text;
                                        string precio = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[9]/td[2]")).Text;
                                        string nom_cont = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[12]/td[2]")).Text;
                                        string nom_telf = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[14]/td[2]")).Text;
                                        string nom_email = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[3]/div[2]/table/tbody/tr[15]/td[2]")).Text;
                                        //bajar un poco mas //*[@id="elementToPDF"]/div[2]/div[6]/div[1]/b
                                        string forma_entrega = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[4]/div/div/div/table/tbody/tr[1]/td[2]")).Text;
                                        string dias_entrega = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[4]/div/div/div/table/tbody/tr[2]/td[2]")).Text;

                                        campos_coti["DESCRIPCION"] = descripcion;
                                        campos_coti["LINK_COTI"] = coti_url;
                                        campos_coti["ENTIDAD"] = entidad;
                                        campos_coti["NUM_COTIZACION"] = num_coti; campos_coti["UNIDAD_COMPRA"] = unidad_compra; campos_coti["DEPENDENCIA"] = dependencia;
                                        campos_coti["FECHA_COTI"] = fecha_presentacion; campos_coti["PRECIO_ESTIMADO"] = precio; campos_coti["NOMBRE_CONTACTO"] = nom_cont;
                                        campos_coti["TELEFONO_CONTACTO"] = nom_telf; campos_coti["CORREO_CONTACTO"] = nom_email;
                                        campos_coti["FORMA_ENTREGA"] = forma_entrega; campos_coti["DIAS_ENTREGA"] = dias_entrega;

                                        //verifica si la descripción de la cotizacion y/o descripcion del producto contiene alguna de las palabras claves
                                        string interes_gbm = "NO";
                                        string categoria = "";
                                        //la categoría si es Servicio o Bien
                                        string objc = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[1]/table/tbody/tr[3]/td[2]")).Text;


                                        //busca si la descripción de la cotización contiene alguna de las frases o palabras clave de la lista anterior
                                        interes_gbm = key_match(descripcion, words);


                                        //buscar el AM de la entidad
                                        string AM = "";
                                        try
                                        {
                                            AM = entidad_info[entidad];
                                        }
                                        catch (Exception)
                                        {
                                            AM = "AA70000134";
                                        }

                                        #region Es de Interes o No por medio de la descripción del producto
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
                                                            interes_gbm = key_match(pd, words);
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
                                        campos_coti["INTERES_GBM"] = interes_gbm;

                                        //Descarga el adjunto solo si es de interes
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
                                                    campos_coti["DOC_ADJUNTOS"] = myFile.Name.ToString();
                                                    adjunto[cont_adj] = root.FilesDownloadPath + "\\" + myFile.Name.ToString();
                                                    cont_adj++;
                                                    Array.Resize(ref adjunto, adjunto.Length + 1);

                                                    //agregar la ruta y nombre del archivo como parte de los adjuntos del array del AM
                                                    adj_names[AM + "_" + cont] = root.FilesDownloadPath + "\\" + myFile.Name.ToString();
                                                    cont++;

                                                }
                                                else
                                                {
                                                    campos_coti["DOC_ADJUNTOS"] = "No hay documentos adjuntos";
                                                }
                                            }
                                            catch (Exception)
                                            {

                                            }
                                            #endregion
                                        }
                                        else
                                        {
                                            campos_coti["DOC_ADJUNTOS"] = "No hay documentos adjuntos";
                                        }

                                        campos_coti["COMENTARIOS"] = "";

                                        #region Extrae la información de producto, agrega a BD y agrega info a excel
                                        string prod_descrip = "noitems";
                                        try
                                        {
                                            prod_descrip = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr/td[6]")).Text;

                                        }
                                        catch (Exception)
                                        {
                                            //no posee items
                                        }
                                        int contador = 1;
                                        JArray productos = new JArray();

                                        while (!string.IsNullOrEmpty(prod_descrip))
                                        {
                                            try
                                            {
                                                string prod_clasi = "";
                                                string prod_cant = "";
                                                string prod_umed = "";
                                                JObject producto = new JObject();
                                                if (prod_descrip == "noitems")
                                                {
                                                    prod_descrip = "";
                                                }
                                                else
                                                {
                                                    prod_descrip = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + contador + "]/td[6]")).Text;
                                                    prod_clasi = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + contador + "]/td[3]")).Text;
                                                    prod_cant = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + contador + "]/td[4]")).Text;
                                                    prod_umed = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div[2]/div[5]/div[2]/div/table/tbody/tr[" + contador + "]/td[5]")).Text;

                                                    prod_descrip = cleantext(prod_descrip);
                                                    prod_clasi = cleantext(prod_clasi);
                                                    prod_cant = cleantext(prod_cant);
                                                    prod_umed = cleantext(prod_umed);
                                                }


                                                producto["PROD_DESCRIPCION"] = prod_descrip;
                                                producto["PROD_CLASIFICACION"] = prod_clasi;
                                                producto["PROD_CANTIDAD"] = prod_cant;
                                                producto["PROD_UN"] = prod_umed;

                                                productos.Add(producto);

                                                #region Agregar coti al excel principal si es de interes de GBM
                                                if (interes_gbm == "SI")
                                                {
                                                    rows = xlWorkSheet.UsedRange.Rows.Count + 1;

                                                    //aregar la cotización al archivo general 
                                                    xlWorkSheet.Cells[rows, 1].value = entidad;
                                                    xlWorkSheet.Cells[rows, 2].value = dependencia;
                                                    xlWorkSheet.Cells[rows, 3].value = num_coti;
                                                    xlWorkSheet.Cells[rows, 4].value = descripcion;
                                                    xlWorkSheet.Cells[rows, 5].value = fecha_presentacion;
                                                    xlWorkSheet.Cells[rows, 6].value = precio;
                                                    xlWorkSheet.Cells[rows, 7].value = nom_cont;
                                                    xlWorkSheet.Cells[rows, 8].value = nom_telf;
                                                    xlWorkSheet.Cells[rows, 9].value = nom_email;
                                                    xlWorkSheet.Cells[rows, 10].value = forma_entrega;
                                                    xlWorkSheet.Cells[rows, 11].value = dias_entrega;
                                                    xlWorkSheet.Cells[rows, 12].value = prod_descrip;
                                                    xlWorkSheet.Cells[rows, 13].value = prod_cant;
                                                    xlWorkSheet.Cells[rows, 14].value = prod_umed;
                                                    xlWorkSheet.Cells[rows, 15].value = coti_url;

                                                    //agregar hypervinculo al link de la PO 
                                                    xlWorkSheet.Hyperlinks.Add(xlWorkSheet.Cells[rows, 15], coti_url);
                                                    xlWorkSheet.Cells[rows, 16].value = campos_coti["DOC_ADJUNTOS"];

                                                    //e++;

                                                    //string file = "Cotizaciones Rapidas " + AM + ".xlsx";
                                                    string file = "Cotizaciones Rapidas - " + AM + ".xlsx";
                                                    int rowam = 2;
                                                    if (!File.Exists(root.FilesDownloadPath + "\\" + file))
                                                    {
                                                        //el excel del AM no existe por lo que se crea
                                                        AMarrays[cont_am] = AM;
                                                        cont_am++;
                                                        Array.Resize(ref AMarrays, AMarrays.Length + 1);
                                                        xlWorkBookAM = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                                                        xlWorkSheetAM = (Excel.Worksheet)xlWorkBookAM.Sheets[1];
                                                        rowam = 2;
                                                        xlWorkSheet.Range["A1", "P1"].Copy();
                                                        //xlWorkSheetAM.Range["A1", "P1"].PasteSpecial(paste, pasteop, false, false);
                                                        xlWorkSheetAM.Range["A1", "P1"].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                                                        xlWorkSheetAM.Columns.AutoFit();
                                                    }
                                                    else
                                                    {
                                                        //el excel existe por lo que se le agrega una nueva fila

                                                        foreach (Excel.Workbook item in xlApp.Workbooks)
                                                        {
                                                            //Select the excel target 'NAME'
                                                            if (item.Name == file)
                                                            {
                                                                xlWorkBookAM = item;
                                                                xlWorkSheetAM = (Excel.Worksheet)xlWorkBookAM.Sheets[1];
                                                                rowam = xlWorkSheetAM.UsedRange.Rows.Count + 1;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    xlWorkSheetAM.Cells[rowam, 1].value = entidad;
                                                    xlWorkSheetAM.Cells[rowam, 2].value = dependencia;
                                                    xlWorkSheetAM.Cells[rowam, 3].value = num_coti;
                                                    xlWorkSheetAM.Cells[rowam, 4].value = descripcion;
                                                    xlWorkSheetAM.Cells[rowam, 5].value = fecha_presentacion;
                                                    xlWorkSheetAM.Cells[rowam, 6].value = precio;
                                                    xlWorkSheetAM.Cells[rowam, 7].value = nom_cont;
                                                    xlWorkSheetAM.Cells[rowam, 8].value = nom_telf;
                                                    xlWorkSheetAM.Cells[rowam, 9].value = nom_email;
                                                    xlWorkSheetAM.Cells[rowam, 10].value = forma_entrega;
                                                    xlWorkSheetAM.Cells[rowam, 11].value = dias_entrega;
                                                    xlWorkSheetAM.Cells[rowam, 12].value = prod_descrip;
                                                    xlWorkSheetAM.Cells[rowam, 13].value = prod_cant;
                                                    xlWorkSheetAM.Cells[rowam, 14].value = prod_umed;
                                                    xlWorkSheetAM.Cells[rowam, 15].value = coti_url;
                                                    //agregar hypervinculo al link
                                                    xlWorkSheetAM.Hyperlinks.Add(xlWorkSheetAM.Cells[rowam, 15], coti_url);
                                                    xlWorkSheetAM.Cells[rowam, 16].value = campos_coti["DOC_ADJUNTOS"];

                                                    xlWorkSheetAM.Columns.AutoFit();
                                                    xlWorkBookAM.SaveAs(root.FilesDownloadPath + "\\" + file);

                                                }
                                                #endregion

                                            }
                                            catch (Exception)
                                            { prod_descrip = ""; }
                                            contador++;
                                        }
                                        #endregion

                                        campos_coti["PROD_DESCRIPCION"] = productos.ToString();
                                        campos_coti["ACCOUNT_MANAGER"] = AM;
                                        console.WriteLine("  Agregando información a la base de datos");
                                        bool add_sql = lpsql.AddQuote(campos_coti);
                                        log.LogDeCambios("Creacion", roots.BDProcess, "Cotizaciones Rápidas Panama", "Agregar Cotización Rapida Convenio", num_coti, add_sql.ToString());
                                        if (add_sql == false)
                                        {
                                            //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                                            //para enviarlo por email y agregarla
                                            resp_sql = resp_sql + num_coti + "<br>";
                                            resp_add_sql = false;
                                        }

                                        //si es de interes y si el AM pertenece a estos 2 usuarios se debe crear el mensaje para enviar por wteams
                                        if (interes_gbm == "SI")
                                        {
                                            if (AM == "AA70000134")
                                            {
                                                //resp_AA70000134 = resp_AA70000134 + "**" + entidad + "**: <a href=\"" + coti_url + "\">" + num_coti + "</a> - " + descripcion + ".<br>";
                                                resp_AA70000134 = resp_AA70000134 + "- **" + entidad + "**: [" + num_coti + "](" + coti_url + ")" + " - " + descripcion + "\n";
                                            }
                                            else if (AM == "AA00070471")
                                            {
                                                //resp_AA00070471 = resp_AA00070471 + "**" + entidad + "**: <a href=\"" + coti_url + "\">" + num_coti + "</a> - " + descripcion + ".<br>";
                                                resp_AA00070471 = resp_AA00070471 + "- **" + entidad + "**: [" + num_coti + "](" + coti_url + ")" + " - " + descripcion + "\n";
                                            }
                                        }

                                        //finaliza todo en la nueva hoja por lo que cierra 
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
            chrome.Close();

            proc.KillProcess("chromedriver", true);

            xlWorkSheet.Columns.AutoFit();


            string mes = "";
            string dia = "";

            mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

            string fecha_file = dia + "_" + mes + "_" + DateTime.Now.Year.ToString();
            string nombre_fila = root.FilesDownloadPath + "\\" + "Cotizaciones Rápidas - Gobierno del Panamá - " + fecha_file + ".xlsx";
            xlWorkBook.SaveAs(nombre_fila);
            xlWorkBook.Close();
            if (xlWorkBookAM != null)
            {
                xlWorkBookAM.Close();
            }
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);
            //adjunto[cont_adj] = nombre_fila;

            fecha_file = dia + "/" + mes + "/" + DateTime.Now.Year.ToString();
            console.WriteLine("  Enviando Reporte");

            string mes_text = DeterminarMes(DateTime.Now.Month);
            #region subir archivos FTP
            Array.Resize(ref adjunto, adjunto.Length - 1);
            for (int i = 0; i < adjunto.Length; i++)
            {
                if (!String.IsNullOrEmpty(adjunto[i].ToString()))
                {
                    //se adjunta el archivo al array para el email
                    string archivo = adjunto[i].ToString();
                    if (File.Exists(archivo))
                    {
                        //bool subir_files = db2.UploadFtp("ftp://10.7.60.72/licitaciones_files/", "gbmadmin", cred.password_server_web, adjunto[i].ToString());
                        ////bool subir_files = wb2.upload_ftp("https://databot.gbm.net/lp/home/assets/adjuntos/", "databot", cred.pass_db1, adjunto[i].ToString());
                        //sharep.UploadFileToSharePoint("https://gbmcorp.sharepoint.com/sites/licitaciones_panama", adjunto[i].ToString(), "databot@gbm.net", cred.password_outlook);
                    }
                }
            }
            #endregion

            #region enviar reportes a AM y chats 

            string header = "Se le notifica que se han publicado las siguientes **cotizaciones rápidas**:\r\n ";
            Array.Resize(ref AMarrays, AMarrays.Length - 1);
            //el AMarrays contiene todos los ID de empleado de los Acount Manager que haya encontrado de acuerdo a la entidad de cada cotizacion de interes
            for (int i = 0; i < AMarrays.Length; i++)
            {
                try
                {
                    //toma el email de SAP de acuerdo al ID AAxxxxx
                    string email = am_email(AMarrays[i].ToString());
                    string[] sep = new string[] { "@" };
                    string[] link = email.Split(sep, StringSplitOptions.None);
                    string user = link[0].ToString().Trim();

                    string file = "Cotizaciones Rapidas - " + AMarrays[i].ToString() + ".xlsx";
                    string Nfile = "Cotizaciones Rapidas - " + user + ".xlsx";
                    File.Move(root.FilesDownloadPath + "\\" + file, root.FilesDownloadPath + "\\" + Nfile);

                    //string subject = "Cotizaciones Rápidas del Gobierno de Panamá - " + AMarrays[i].ToString() + " - " + fecha_file;
                    string subject = "Cotizaciones rápidas programadas del Gobierno de Panamá – Cuentas " + user + " - Fecha: " + fecha_file;
                    //string[] adj = { root.Google_Download + "\\" + file };
                    string[] adj = new string[1];
                    int x = 0;
                    //se realza un for en el array de los archivos que se han descargado
                    //los keys del diccionario contiene el ID del AM
                    foreach (KeyValuePair<string, string> pair in adj_names)
                    {
                        //console.WriteLine("FOREACH KEYVALUEPAIR: {0}, {1}", pair.Key, pair.Value);
                        //si la llave contiene el id del AM
                        if (pair.Key.ToString().Contains(AMarrays[i].ToString()))
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
                    //al final que un campo vacio en el array de adjuntos del email y es donde se adjunta el file
                    adj[adj.Length - 1] = root.FilesDownloadPath + "\\" + Nfile;
                    string body = "Se adjunta el reporte de Cotizaciones Rápidas del Gobierno de Panamá en sus cuentas asignadas programadas al día de mañana";
                    mail.SendHTMLMail(body, new string[] { email }, subject, root.CopyCC, adj);

                    //Enviar notificaciones solamente a estos 2 usuarios
                    if ("AA00070471" == AMarrays[i].ToString() || "AA70000134" == AMarrays[i].ToString())
                    {
                        string mensaje = ""; //this.GetType().GetField("resp_" + AMarrays[i].ToString()).GetValue(this);
                        mensaje = ("AA00070471" == AMarrays[i].ToString()) ? resp_AA00070471 : ("AA70000134" == AMarrays[i].ToString()) ? resp_AA70000134 : "";
                        //try
                        //{ PushN.GenerarNotificacion(root.BD_Proceso, email, header + mensaje.ToString()); }
                        //catch (Exception ex)
                        //{ }
                        wt.SendNotification(email, "Nuevas cotizaciones rápidas LCPA", header + mensaje.ToString());

                    }
                }
                catch (Exception ex)
                {
                    string[] adj = new string[1];
                    if (File.Exists("Cotizaciones Rapidas - " + AMarrays[i].ToString() + ".xlsx"))
                    {
                        adj[0] = "Cotizaciones Rapidas - " + AMarrays[i].ToString() + ".xlsx";
                    }

                    mail.SendHTMLMail(
                        "Error al enviar el reporte del usuario: " + AMarrays[i].ToString(),
                        new string[] { "dmeza@gbm.net" }, 
                        "Error: Cotizaciones rápidas programadas del Gobierno de Panamá",
                        root.CopyCC,
                        adj);
                    console.WriteLine(ex.Message);
                }
            }
            #endregion

            if (validar_lineas == false)
            {
                string[] cc = { "appmanagement@gbm.net" };
                string[] adj = { nombre_fila };
                mail.SendHTMLMail("A continuacion se adjunta el archivo con el conglomerado de las nuevas cotizaciones rapidas de GBM de Panama", new string[] { "kvanegas@gbm.net" }, "Reporte diario de Cotizaciones Rapidas - " + fecha_file, cc, adj);

            }
            if (resp_add_sql == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                string[] adj = { nombre_fila };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Cotizaciones a la base de datos: " + "<br>" + resp_sql, new string[] { "appmanagement@gbm.net" }, "Reporte diario de Cotizaciones Rapidas - " + fecha_file, cc, adj);

            }
            return "";
        }

        #region Metodos de apoyo
        public string crear_opp(Dictionary<string, string> campos)
        {
            string idopp = "";
            try
            {
                RfcDestination destination = sap.GetDestRFC("CRM");

                console.WriteLine(" Conectado con SAP CRM");
                RfcRepository repo = destination.Repository;
                IRfcFunction func = repo.CreateFunction("ZOPP_VENTAS");
                IRfcTable general = func.GetTable("GENERAL");
                IRfcTable partners = func.GetTable("PARTNERS");
                // IRfcTable items = func.GetTable("ITEMS");
                console.WriteLine(" Llenando informacion general de oportunidad");
                general.Append();
                general.SetValue("TIPO", campos["tipo"].ToString());
                general.SetValue("DESCRIPCION", campos["descripcion"].ToString());
                general.SetValue("FECHA_INICIO", campos["fecha_inicio"].ToString());
                general.SetValue("FECHA_FIN", campos["Fecha_Final"].ToString());
                general.SetValue("FASE_VENTAS", campos["Ciclo"].ToString());
                // general.SetValue("CICLO_VENTAS", datos_oportunidad.DATA_GENERAL.ORIGEN);
                general.SetValue("PORCENTAJE", "100");
                general.SetValue("REVENUE", "");
                general.SetValue("MONEDA", "USD");
                general.SetValue("GRUPO_OPP", campos["grupo_opp"].ToString());
                general.SetValue("ORIGEN", campos["Origen"].ToString());
                general.SetValue("PRIORIDAD", "4");
                console.WriteLine(" Llenando informacion de cliente y equipo de ventas");
                partners.Append();
                partners.SetValue("PARTNER", campos["Cliente"].ToString());
                partners.SetValue("FUNCTION", "00000021");
                partners.Append();
                partners.SetValue("PARTNER", campos["Contacto"].ToString());
                partners.SetValue("FUNCTION", "00000015");
                partners.Append();
                partners.SetValue("PARTNER", campos["Usuario"].ToString());
                partners.SetValue("FUNCTION", "00000014");
                console.WriteLine(" Llenando Org de Servicios y Ventas");
                func.SetValue("SALES_ORG", campos["OrgVentas"].ToString());
                func.SetValue("SRV_ORG", campos["OrgServicios"].ToString());
                console.WriteLine(" Creando Oportunidad en SAP CRM");
                func.Invoke(destination);



                IRfcTable validate = func.GetTable("VALIDATE");

                if (func.GetValue("RESPONSE").ToString() != "")
                {
                    console.WriteLine(" Response of the request: " + func.GetValue("RESPONSE").ToString());
                }
                if (func.GetValue("OPP_ID").ToString() != "")
                {
                    console.WriteLine(" ID de la oportunidad creada: " + func.GetValue("OPP_ID").ToString());
                    idopp = func.GetValue("OPP_ID").ToString();
                }
                else
                {
                    idopp = "Error: creating the opportunity";
                    console.WriteLine(" Error creating the opportunity");
                }
                for (int i = 0; i < validate.Count; i++)
                {
                    console.WriteLine(" Generated errors:");
                    console.WriteLine(DateTime.Now + " > > >  " + validate[i].GetValue("MENSAJE") + "\r\n");
                }
                console.WriteLine("");

            }
            catch (Exception ex)
            {
                idopp = "Error: " + ex.Message;
            }

            return idopp;
        }
        public string am_email(string AM)
        {
            string email = "";
            try
            {
                RfcDestination destination = new SapVariants().GetDestRFC("ERP");

                console.WriteLine(" Conectado con SAP ERP");

                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters["BP"] = AM;
                IRfcFunction func = new SapVariants().ExecuteRFC(sapSys, "ZDM_READ_BP", parameters);

                email = func.GetValue("EMAIL").ToString();
            }
            catch (Exception ex)
            {
                email = "acolina@gbm.net";
            }

            return email;
        }
        public int dias_habiles(int cantidad)
        {
            int dias = 0;
            if (cantidad >= 1 && cantidad <= 10)
            { dias = 15; }
            else if (cantidad >= 11 && cantidad <= 25)
            { dias = 25; }
            else if (cantidad >= 11 && cantidad <= 25)
            { dias = 25; }
            else if (cantidad >= 26 && cantidad <= 50)
            { dias = 35; }
            else if (cantidad >= 51 && cantidad <= 100)
            { dias = 45; }
            else if (cantidad >= 101 && cantidad <= 250)
            { dias = 60; }
            else if (cantidad >= 251 && cantidad <= 500)
            { dias = 70; }
            else if (cantidad >= 501)
            { dias = 90; }
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
        private string DeterminarMes(int mes)
        {
            string mes_text = "";
            switch (mes)
            {
                case 1:
                    mes_text = "Enero";
                    break;
                case 2:
                    mes_text = "Febrero";
                    break;
                case 3:
                    mes_text = "Marzo";
                    break;
                case 4:
                    mes_text = "Abril";
                    break;
                case 5:
                    mes_text = "Mayo";
                    break;
                case 6:
                    mes_text = "Junio";
                    break;
                case 7:
                    mes_text = "Julio";
                    break;
                case 8:
                    mes_text = "Agosto";
                    break;
                case 9:
                    mes_text = "Setiembre";
                    break;
                case 10:
                    mes_text = "Octubre";
                    break;
                case 11:
                    mes_text = "Noviembre";
                    break;
                case 12:
                    mes_text = "Diciembre";
                    break;
                default:
                    mes_text = "";
                    break;
            }
            return mes_text;
        }
        public DateTime AddWorkdays(DateTime originalDate, int workDays)
        {
            DateTime tmpDate = originalDate;
            DateTime[] feriados = getholidays();
            try
            {
                while (workDays > 0)
                {
                    tmpDate = tmpDate.AddDays(1); //agregar un dia a la fecha
                    DateTime ntmpDate = new DateTime(tmpDate.Year, tmpDate.Month, tmpDate.Day); //para quitarle las horas
                    bool feriado = Array.Exists(feriados, x => x == ntmpDate); //para saber si el ntmpDate es feriado en la lista de feriados
                    //si el dayofweek de tmpDate esta entre L-V si es menor a sabado pero mayor a domingo (o bien si no es feriado)
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
        public DateTime[] getholidays()
        {
            DateTime[] feriados = new DateTime[1];
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            DataTable mytable = new DataTable();

            try
            {
                #region Connection DB   
                sql_select = "select * from feriados_panama";
                //mytable = crud.Select("Databot", sql_select, "ventas");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        DateTime fecha = DateTime.Parse(mytable.Rows[i][1].ToString());
                        DateTime nfecha = new DateTime(DateTime.Now.Year, fecha.Month, fecha.Day);
                        feriados[i] = nfecha; //fechas feriadas
                        Array.Resize(ref feriados, feriados.Length + 1);
                    }
                }
                else
                {
                    feriados[0] = DateTime.MinValue;

                }

            }
            catch (Exception ex)
            {
                //marca[0] = "No se encontró este producto en la lista";
            }
            Array.Resize(ref feriados, feriados.Length - 1);
            return feriados;
        }
        public string key_match(string texto, List<string> words)
        {
            string interes_gbm = "NO";
            try
            {
                texto = texto.ToLower();
                texto = texto.Replace("á", "a"); texto = texto.Replace("é", "e"); texto = texto.Replace("í", "i"); texto = texto.Replace("ó", "o"); texto = texto.Replace("ú", "u");
                //texto = val.QuitarEnne(texto);
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
        public string cleantext(string text)
        {
            text = text.Replace("\"", "");
            text = text.Replace("'", "");
            return text;
        }
        #endregion

        #region robots locales
        public Dictionary<string, string> extraerinfolink(string link)
        {
            Dictionary<string, string> campos = new Dictionary<string, string>();

            #region eliminar cache and cookies chrome
            try
            {
            }
            catch (Exception)
            { }
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                try
                { chrome.Navigate().GoToUrl(link); }
                catch (Exception)
                { chrome.Navigate().GoToUrl(link); }
                IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
                #endregion

                chrome.Manage().Cookies.DeleteAllCookies();

                console.WriteLine("  Extrayendo información de la página web");

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblLugarEntrega']"))); }
                catch { }

                string unidad_solicitante = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblUnidadSolicitante']")).Text,
                  contactocuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblRegistradoPor']")).Text,
                  emailcuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblEmail']")).Text;
                campos.Add("UNIDAD_SOLICITANTE", unidad_solicitante); campos.Add("CONTACTO_CUENTA", contactocuenta); campos.Add("EMAIL_CUENTA", emailcuenta);


            }
            catch (Exception ex)
            {
            }

            try { chrome.Close(); } catch (Exception) { }

            proc.KillProcess("chromedriver", true);

            return campos;
        }
        public void get_customer()
        {
            string respuesta = "";
            string cant_filas = "";
            double filas = 0;
            int pag_row = 0;
            string id_producto = "";
            string main_web_page = "";
            double precio_unitario = 0;
            double prod_cant = 0;
            double prod_total = 0;
            string vendor_text = "";
            string resp_sql = "";
            bool resp_add_sql = true;
            int cont_adj = 0;
            bool validar_lineas = true;
            DateTime file_date = DateTime.MinValue;
            DateTime file_date_before = DateTime.MinValue;

            #region eliminar cache and cookies chrome
            try
            {
            }
            catch (Exception)
            { }
            #endregion

            #region excel
            console.WriteLine("  Creando Excel");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            #region titulos_excel
            xlWorkSheet.Cells[1, 1].value = "ENTIDAD";
            xlWorkSheet.Cells[1, 2].value = "CONVENIO";
            xlWorkSheet.Cells[1, 3].value = "PROVEEDOR";
            xlWorkSheet.Cells[1, 4].value = "REGISTRO";
            xlWorkSheet.Cells[1, 5].value = "UNIDAD SOLICITANTE";
            Excel.Worksheet xlSheet = xlWorkBook.ActiveSheet;
            for (int i = 1; i <= 17; i++)
            {
                Excel.Range rango = (Excel.Range)xlSheet.Cells[1, i];
                rango.Interior.Color = Excel.XlRgbColor.rgbRoyalBlue;
                rango.Font.Color = Excel.XlRgbColor.rgbWhite;
            }
            Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
            Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;

            #endregion
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            #region Ingreso al website
            try
            {
                console.WriteLine("  Ingresando al website");
                chrome.Navigate().GoToUrl("https://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo");
            }
            catch (Exception)
            { chrome.Navigate().GoToUrl("https://panamacompra.gob.pa/Inicio/#!/busquedaCatalogo"); }
            IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;
            #endregion
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='apc.fields.IdConvenio']"))); }
            catch { }
            chrome.Manage().Cookies.DeleteAllCookies();
            console.WriteLine("  Seleccionar Convenio");
            SelectElement convenio_select = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdConvenio']")));
            System.Threading.Thread.Sleep(1000);
            convenio_select.SelectByValue("number:107"); //BIENES INFORMÁTICOS, REDES Y COMUNICACIONES
            System.Threading.Thread.Sleep(3000);
            string convenio_name = convenio_select.SelectedOption.Text;

            string fecha_desde = chrome.FindElement(By.XPath("//*[@id='apc.fields.fd']")).Text;
            string fecha_hasta = chrome.FindElement(By.XPath("//*[@id='apc.fields.fh']")).Text;
            if (fecha_hasta == "" || fecha_hasta == null)
            {
                file_date = DateTime.Today;
                file_date_before = file_date.AddMonths(-1);
                fecha_hasta = "'" + file_date.ToString("dd-MM-yyyy");
                fecha_desde = "'" + file_date_before.ToString("dd-MM-yyyy");
            }
            SelectElement vendor = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdProveedor']")));
            System.Threading.Thread.Sleep(1000);
            vendor.SelectByValue("number:1979769"); //GBM PANAMA
            string vendor_name = vendor.SelectedOption.Text;
            IWebElement entidad_list = chrome.FindElement(By.XPath("//*[@id='apc.fields.IdEmpresa']"));
            System.Threading.Thread.Sleep(1000);
            SelectElement selectList = new SelectElement(entidad_list);
            IList<IWebElement> voptions = selectList.Options;
            string ent_text = "";
            for (int z = 1; z <= voptions.Count; z++) //lista_vendors.Count vendor_lists.Length - 1
            {
                try
                {
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    SelectElement entidades = new SelectElement(chrome.FindElement(By.XPath("//*[@id='apc.fields.IdEmpresa']")));

                    IWebElement ent_option = chrome.FindElement(By.XPath("//*[@id='apc.fields.IdEmpresa']/option[" + z + "]"));

                    ent_text = "";
                    ent_text = ent_option.Text.ToString();

                    if (ent_text == "-- Seleccione --") //&& vendor_text != "GBM de Panamá, S.A"
                    {
                        continue;
                    }
                    if (ent_text == "MINISTERIO DE GOBIERNO" || ent_text == "MINISTERIO DE EDUCACION" || ent_text == "MINISTERIO DE SALUD" || ent_text == "MINISTERIO DE LA PRESIDENCIA"
                        || ent_text == "Ministerio de Seguridad Pública" || ent_text == "Ministerio de Seguridad Pública" || ent_text == "CAJA DE SEGURO SOCIAL" || ent_text == "MUNICIPIO DE ANTÓN")
                    {
                        console.WriteLine("   Vendor: " + vendor_text);


                        System.Threading.Thread.Sleep(1000);
                        //string chrometime = chrome.Manage().Timeouts().PageLoad.ToString();
                        chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
                        try
                        { entidades.SelectByIndex(z - 1); }
                        catch (Exception)
                        { entidades.SelectByIndex(z - 1); }
                        chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);
                        #region calendario

                        #endregion

                        chrome.FindElement(By.XPath("//*[@id='busquedaC2']/div[4]/div/center/button")).Click(); //BUSCAR
                        console.WriteLine("   Buscar");
                        System.Threading.Thread.Sleep(3000);
                        main_web_page = chrome.Url;

                        cant_filas = "";
                        try
                        { cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text; }
                        catch (Exception)
                        { cant_filas = ""; }

                        console.WriteLine("   " + cant_filas);
                        if (cant_filas != "Se encontraron 0 Pedidos Publicados")
                        {
                            int mas = 0;
                            if (cant_filas.Contains("+"))
                            {
                                do
                                {
                                    IWebElement pagination_last = chrome.FindElement(By.ClassName("pagination-last"));
                                    var pag_last = pagination_last.FindElement(By.TagName("a"));
                                    pag_last.Click();
                                    System.Threading.Thread.Sleep(1500);
                                    cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                    mas++;
                                } while (cant_filas.Contains("+"));



                                if (mas > 0)
                                {
                                    IWebElement pagination_first = chrome.FindElement(By.ClassName("pagination-first"));
                                    var pag_first = pagination_first.FindElement(By.TagName("a"));
                                    pag_first.Click();
                                    System.Threading.Thread.Sleep(1000);
                                    cant_filas = chrome.FindElement(By.XPath("//*[@id='body']/div/div[2]/div/div/div[2]/div[2]/div[1]/h5/b")).Text;
                                }
                            }
                            cant_filas = cant_filas.Substring(15, 3).Trim();
                            if (cant_filas.Contains(" "))
                            {
                                cant_filas = cant_filas.Substring(0, 2).Trim();
                            }

                            double num1 = double.Parse(cant_filas);
                            double num2 = 10;
                            filas = (num1 / num2);
                            double pag_row2 = Math.Ceiling(filas);
                            if (pag_row2 == 0)
                            { pag_row2 = 1; }

                            for (int i = 1; i <= pag_row2; i++)
                            {

                                int rows = xlWorkSheet.UsedRange.Rows.Count + 1;

                                //next en pagination
                                if (i != 1)
                                {
                                    console.WriteLine("     Siguiente pagina");
                                    IWebElement pagination_next = chrome.FindElement(By.ClassName("pagination-next"));
                                    var pag_next = pagination_next.FindElement(By.TagName("a"));
                                    pag_next.Click();
                                }
                                //*[@id="body"]/div/div[2]/div/div/div[2]/div[2]/center/ul/li[5]/a
                                IWebElement tableElement = chrome.FindElement(By.XPath("//*[@id='toTopBA']"));
                                IList<IWebElement> trCollection = tableElement.FindElements(By.TagName("tr"));
                                int e = 0;

                                foreach (IWebElement element in trCollection)
                                {
                                    IList<IWebElement> tdCollection;
                                    tdCollection = element.FindElements(By.TagName("td"));
                                    if (tdCollection.Count > 0)
                                    {
                                        string entidad = tdCollection[1].Text;
                                        string registro_unico = tdCollection[2].Text; //Registro único de pedido
                                        IWebElement link = tdCollection[2];
                                        var linkhref = link.FindElement(By.TagName("a"));
                                        new Actions(chrome).MoveToElement(link).Perform();

                                        try
                                        { linkhref.Click(); }
                                        catch (Exception)
                                        { linkhref.Click(); }

                                        System.Threading.Thread.Sleep(2000);

                                        chrome.SwitchTo().Window(chrome.WindowHandles[1]); //ir al nuevo tab

                                        //hacer todo lo que tenga que hacer en la nueva hoja-----
                                        try { new WebDriverWait(chrome, new TimeSpan(0, 0, 60)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblFechaRegistro']"))); }
                                        catch { }

                                        string unidad_solicitante = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblUnidadSolicitante']")).Text;
                                        string contacto_cuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblRegistradoPor']")).Text;
                                        string email_cuenta = chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_lblEmail']")).Text;

                                        xlWorkSheet.Cells[rows + e, 1].value = ent_text;
                                        xlWorkSheet.Cells[rows + e, 2].value = convenio_name;
                                        xlWorkSheet.Cells[rows + e, 3].value = vendor_name;
                                        xlWorkSheet.Cells[rows + e, 4].value = registro_unico;
                                        xlWorkSheet.Cells[rows + e, 5].value = unidad_solicitante;
                                        //xlWorkSheet.Cells[rows + e, 1].value = convenio_name;

                                        chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                                        chrome.Close();

                                        chrome.SwitchTo().Window(chrome.WindowHandles[0]); //regreso a la pagina principal y sigo con la siguiente fila
                                        e++;

                                    } //la fila tiene td
                                } //foreach fila en la tabla main
                            } //for cantidad de paginas de la tabla main
                        }
                        else
                        {
                            console.WriteLine("");
                        }



                    } //if proveedor drop down
                      //resh++;
                }
                catch (Exception ex)
                {
                    //console.WriteLine(ex.ToString());
                    int row = xlWorkSheet.UsedRange.Rows.Count + 1;
                    if (row == 1)
                    { row = 2; }
                    xlWorkSheet.Cells[row, 1].value = ent_text;
                    xlWorkSheet.Cells[row, 2].value = convenio_name;
                    xlWorkSheet.Cells[row, 3].value = vendor_name;
                    xlWorkSheet.Cells[row, 4].value = "Error al descargar informacion";
                    xlWorkSheet.Cells[row, 5].value = ex.Message;

                }
            } //for vendor

            console.WriteLine("  Guardar el reporte");
            chrome.Close();

            proc.KillProcess("chromedriver", true);

            xlWorkSheet.Columns.AutoFit();

            string nombre_fila = root.FilesDownloadPath + "\\" + "Reporte de entidades.xlsx";
            xlWorkBook.SaveAs(nombre_fila);
            xlWorkBook.Close();
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);
        }
        public bool add_ent(Dictionary<string, string> campos)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_insert = "";
            string sql_insert2 = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable update = new DataTable();
            bool respuesta = true;


            try
            {
                //#region Connection DB     
                //MySqlConnection conn = new Database().Conn("ventas");
                //#endregion

                sql_insert = "INSERT INTO `sector`(`id`, `entidad`, `sector`, `cliente`, `contacto`, `sales_rep`) VALUES ('" + campos["id"] + "','" + campos["entidad"] + "','" + campos["sector"] + "','" + campos["cliente"] + "','" + campos["contacto"] + "','" + campos["sales_rep"] + "')";
                //crud.Insert("Databot", sql_insert, "ventas");

                return true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;
                return false;
            }
        }
        /// <summary>
        /// es importante que los keys del diccionario sea el mismo nombre (incluyendo mayusculas) que las columnas de la BD
        /// </summary>
        /// <param name="registro"></param>
        /// <param name="dictionary"></param>
        /// <returns></returns>
        #endregion

    }
    public class info_vendor
    {
        public string texto { get; set; }
        public string valor { get; set; }
    }
}
