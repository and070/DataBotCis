using System;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Projects.PanamaBids;
using DataBotV5.App.Global;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace DataBotV5.Automation.WEB.PanamaBids
{
    /// <summary>
    /// Clase Web Automation encargada de obtener precios del convenio GBPA.
    /// </summary>
    class AgreementGPBAGetPrices
    {
        #region variables_globales
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();
        ProcessAdmin padmin = new ProcessAdmin();
        WebInteraction sel = new WebInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        Log log = new Log();
        CRUD crud = new CRUD();
        MsExcel MsExcel = new MsExcel();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        string mand = "QAS";
        private string mandante = "QAS";
        string respFinal = "";

        #endregion
        /// <summary>
        /// Metodo para actualizar precio de productos se realiza todos los miercoles
        /// </summary>
        public void Main()
        {
            console.WriteLine("Procesando...");
            string respuesta = GetProductsPrices();
            if (respuesta != "")
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar descargar los precios: " + "<br>" + respuesta, new string[] {"appmanagement@gbm.net"}, "Error: Get Prices of Convenio Macro", cc);

            }
            else
            {
                console.WriteLine("Creando estadísticas...");

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }



        }
        /// <summary>
        /// Actualizar precios todos los miercoles
        /// </summary>
        /// <returns></returns>
        private string GetProductsPrices()
        {
            bool valRows = true;
            string errorMsj = @"<table class='myCustomTable' width='100 %'>
                <thead><tr><th>Orden</th><th>Producto</th><th>Precio</th><th>Respuesta</th></tr></thead>
                <tbody>";
            //string today = DateTime.Now.ToString("yyyy-MM-dd");
            DateTime today = DateTime.Today;
            string wend = today.AddDays(-(int)today.DayOfWeek).AddDays(3).ToString("yyyy-MM-dd"); //sacar el miercoles de la semana

            Int32 fila_producto = -1;
            Int32 fila_vendor = -1;
            string respuesta = "";
            DataTable products = lpsql.productsInfo();
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
                console.WriteLine("Ingresando al website");
                chrome.Navigate().GoToUrl("http://catalogo.panamacompra.gob.pa/forms/Publico/documentosPrecioProveedor.aspx");

                chrome.Manage().Cookies.DeleteAllCookies();

                console.WriteLine("Extrayendo información de la página web");

                try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlFecha0']"))); }
                catch { }

                SelectElement fecha_select = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlFecha0']")));
                System.Threading.Thread.Sleep(1000);
                fecha_select.SelectByValue(wend);
                System.Threading.Thread.Sleep(3000);

                SelectElement convenio = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlConvenioFiltro']")));
                System.Threading.Thread.Sleep(1000);
                convenio.SelectByValue("134"); //Bienes informaticos
                System.Threading.Thread.Sleep(3000);

                SelectElement region = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlRegion']")));
                System.Threading.Thread.Sleep(1000);
                region.SelectByValue("33"); //Provincia de panama
                System.Threading.Thread.Sleep(3000);

                SelectElement reglones = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlRenglon']")));
                System.Threading.Thread.Sleep(1000);
                IList<IWebElement> voptions = reglones.Options;

                string ruta = root.FilesDownloadPath + "\\" + "ReportePrecio.xls";
                int z = 1;
                string body = @"<table class='myCustomTable' width='100 %'>
                <thead><tr><th>Reglon</th><th>Descripción</th><th>Precio</th><th>Respuesta</th></tr></thead>
                <tbody>";
                string price = "", reglonText = "", productCode = "";
                foreach (IWebElement selectElement in voptions.Skip(1)) //por cada reglon
                {
                    try
                    {
                        price = ""; reglonText = ""; productCode = "";
                        IWebElement rOption = chrome.FindElement(By.XPath($"//*[@id='ctl00_ContentPlaceHolder1_ddlRenglon']/option[{z + 1}]"));
                        reglonText = rOption.Text.ToString();

                        productCode = reglonText.Split(new char[] { '-' })[0].ToString().Trim();
                        EnumerableRowCollection<DataRow> result = products.AsEnumerable().Where(myRow => myRow.Field<string>("productCode") == productCode);
                        DataRow[] pBrand = products.Select($"productCode ='{productCode}'");
                        if (pBrand.Length != 0)
                        {
                            console.WriteLine("Reglon: " + reglonText);

                            SelectElement reglon = new SelectElement(chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_ddlRenglon']")));
                            System.Threading.Thread.Sleep(1000);
                            chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
                            try
                            { reglon.SelectByIndex(z); }
                            catch (Exception)
                            { reglon.SelectByIndex(z); }
                            chrome.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120);

                            chrome.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_btnPdf']")).Click(); //BUSCAR
                            console.WriteLine("Buscar...");
                            System.Threading.Thread.Sleep(3000);

                            chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                            System.Threading.Thread.Sleep(1000);
                            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[1]/tbody/tr[2]/td[1]/div"))); }
                            catch { }

                            //*[@id="__bookmark_2"]
                            HtmlAgilityPack.HtmlDocument sourceUrl = new HtmlAgilityPack.HtmlDocument();
                            sourceUrl.LoadHtml(chrome.PageSource);
                            HtmlAgilityPack.HtmlNodeCollection eptable = sourceUrl.DocumentNode.SelectNodes("//*[@id='__bookmark_2']/tbody/tr"); //tomar toda la tabla
                                                                                                                                                 //por cada fila de la tabla principal de licitaciones
                            if (eptable != null)
                            {


                                for (int g = 4; g < eptable.Count; g++)
                                {
                                    string vendor = chrome.FindElement(By.XPath($"//*[@id='__bookmark_2']/tbody/tr[{g}]/td[1]/div")).Text;
                                    if (vendor.Contains("(Precio De 25 a 60 unidades)"))
                                    {
                                        break;
                                    }
                                    if (vendor == "GBM de Panamá, S.A")
                                    {
                                        price = chrome.FindElement(By.XPath($"//*[@id='__bookmark_2']/tbody/tr[{g}]/td[3]/div")).Text.ToString().Replace("$", "").Replace(".", "");
                                        bool up = crud.Update($"UPDATE products SET price = '{price}' WHERE productCode = '{productCode}'", "panama_bids_db");
                                        if (up)
                                        {

                                            log.LogDeCambios("Actualizar", root.BDProcess,  root.BDUserCreatedBy, "Actualizar los totales de producto", $"Se actualiza el producto del código: {productCode}, con el precio: {price}", root.Subject);
                                            respFinal = respFinal + "\\n" + $"Se actualiza el producto del código: {productCode}, con el precio: {price}";

                                            //si actualizó bien el producto que actualice los totales de producto en la tabla purchaseOrderProducts de las ordenes en refrendo
                                            string sqlSelect = $"SELECT purchaseOrderProduct.productId, purchaseOrderProduct.singleOrderRecord, purchaseOrderProduct.productCode, purchaseOrderProduct.quantity, purchaseOrderProduct.totalProduct FROM `purchaseOrderProduct` INNER JOIN purchaseOrderMacro ON purchaseOrderProduct.singleOrderRecord = purchaseOrderMacro.singleOrderRecord WHERE purchaseOrderProduct.productCode = {productCode} and purchaseOrderMacro.orderStatus = 3";
                                            DataTable poCodeEnRefrendo = crud.Select( sqlSelect, "panama_bids_db");
                                            if (poCodeEnRefrendo.Rows.Count > 0)
                                            {
                                                foreach (DataRow rRow in poCodeEnRefrendo.Rows)
                                                {
                                                    string productId = rRow["productId"].ToString();
                                                    string singleOrderRecord = rRow["singleOrderRecord"].ToString();
                                                    int quantity = int.Parse(rRow["quantity"].ToString());
                                                    int priceInt = int.Parse(price);
                                                    int newTotal = priceInt * quantity;
                                                    string sqlUp = $"UPDATE `purchaseOrderProduct` SET `totalProduct`= '{newTotal}' WHERE `productId` = '{productId}'";
                                                    string sqlUp2 = $"UPDATE `purchaseOrderMacro` SET `orderSubtotal`= '{newTotal}' WHERE `singleOrderRecord` = '{singleOrderRecord}'";
                                                    bool up2 = crud.Update(sqlUp, "panama_bids_db");
                                                    bool up3 = crud.Update(sqlUp2, "panama_bids_db");
                                                    if (!up2 || !up3)
                                                    {
                                                        valRows = false;
                                                        errorMsj = errorMsj + $@"<tr>
                                                    <td>{singleOrderRecord}</td>
                                                    <td>{productCode}</td>
                                                    <td>{price}</td>
                                                    <td>Error al actualizar registro de orden</td>
                                                    </tr>";
                                                    }
                                                }
                                            }
                                        }

                                        body = body + $@"<tr>
                                        <td>{productCode}</td>
                                        <td>{reglonText}</td>
                                        <td>{price}</td>
                                        <td>{((!up) ? "Error al actualizar" : "Ok")}</td>
                                        </tr>";

                                        break;
                                    }
                                }


                                //*[@id="__bookmark_2"]/tbody/tr[4]/td[1]/div
                                //*[@id="__bookmark_2"]/tbody/tr[5]/td[1]/div

                                //console.WriteLine("Descargando");
                                //chrome.FindElement(By.XPath("//*[@id='toolbar']/table/tbody/tr/td/input[3]")).Click(); //descargar
                                //System.Threading.Thread.Sleep(1000);

                                //for (var x = 0; x < 40; x++)
                                //{
                                //    if (File.Exists(ruta)) { break; }
                                //    System.Threading.Thread.Sleep(1000);
                                //}

                            }
                            try
                            {
                                //chrome.SwitchTo().Window(chrome.WindowHandles[2]);
                                //chrome.Close();
                                //chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                                chrome.Close();
                                chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                                //chrome.Close();
                            }
                            catch (Exception)
                            {

                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        body = body + $@"<tr>
                                        <td>{productCode}</td>
                                        <td>{reglonText}</td>
                                        <td>{price}</td>
                                        <td>{$"Error al actualizar {ex.ToString()}"}</td>
                                        </tr>";

                        try
                        {
                            //chrome.SwitchTo().Window(chrome.WindowHandles[2]);
                            //chrome.Close();
                            //chrome.SwitchTo().Window(chrome.WindowHandles[1]);
                            chrome.Close();
                            chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                            //chrome.Close();
                        }
                        catch (Exception)
                        {

                        }
                    }
                    z++;
                }
                body = body + "</tbody>";
                body = body + "</table>";

                proc.KillProcess("chromedriver", true);

                string emailhtml = Properties.Resources.emailtemplate1;
                #region Enviar correo de error
                if (!valRows)
                {
                    errorMsj = errorMsj + "</tbody>";
                    errorMsj = errorMsj + "</table>";
                    emailhtml = emailhtml.Replace("{subject}", "Error al Actualizar Precios de Productos");
                    emailhtml = emailhtml.Replace("{cuerpo}", $"Se registraron error al actualizar el total de las siguientes ordenes, al {DateTime.Now.ToString("dd/MM/yyyy")}");
                    emailhtml = emailhtml.Replace("{contenido}", errorMsj);
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail(emailhtml, new string[] {"appmanagement@gbm.net"}, $"Error: actualización de Precios - Convenio Marco {DateTime.Now.ToString("dd/MM/yyyy")}", cc);
                }
                #endregion

                #region enviar correo
                emailhtml = Properties.Resources.emailtemplate1;
                emailhtml = emailhtml.Replace("{subject}", "Notificación Actualizar Precios de Productos");
                emailhtml = emailhtml.Replace("{cuerpo}", $"A Continuación los precios actualizados del Convenio Marco, al {DateTime.Now.ToString("dd/MM/yyyy")}");
                emailhtml = emailhtml.Replace("{contenido}", body);

                mail.SendHTMLMail(emailhtml, new string[] { "kvanegas@gbm.net" }, $"Actualización de Precios - Convenio Marco {DateTime.Now.ToString("dd/MM/yyyy")}", null);

                root.BDUserCreatedBy = "KVANEGAS";

                #endregion
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                respuesta = ex.Message;
            }

            root.requestDetails = respFinal;

            return respuesta;
        }

    }
}
