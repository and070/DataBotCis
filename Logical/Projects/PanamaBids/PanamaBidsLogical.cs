using DataBotV5.App.Global;
using DataBotV5.Automation.MASS.PanamaBids;
using DataBotV5.Automation.WEB.PanamaBids;
using DataBotV5.Data;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Projects.PanamaBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;

namespace DataBotV5.Logical.Projects.PanamaBids
{
    class PanamaBidsLogical
    {
        BidsGbPaSql lpsql = new BidsGbPaSql();
        Rooting root = new Rooting();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        ConsoleFormat console = new ConsoleFormat();
        public PoInfoMacro GetPoInfo(IWebDriver chrome, DataTable excelResult, DataTable excelCisco, string convenio, string entidad, string[] adjunto, Dictionary<string, string> AMs, string singleOrderRecord, int cont_adj, string cantidadOpp, List<string> newEntities)
        {
            #region variables privadas

            PoInfoMacro poInfo = new PoInfoMacro();

            string[] CopyCC = new string[1];
            string
                resp_sql = "",
                sector = "",
                brandOpp = "";
            bool resp_add_sql = true,
                validar_lineas = true,
                cisco_add = false,
                no_registros = false,
                newEntity = false;
            int cont = 0;
            DateTime file_date = DateTime.MinValue, file_date_before = DateTime.MinValue;
            Dictionary<string, string> adj_names = new Dictionary<string, string>();
            #endregion

            DataTable entities = lpsql.entitiesInfo();
            DataTable productBrand = lpsql.productsInfo();

            #region Extraer información de la orden 
            //esperar hasta que la fecha de registro salga (es decir esperar que la pagina termine de cargar)
            try { new WebDriverWait(chrome, new TimeSpan(0, 0, 30)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[3]/td"))); }
            catch { }

            string lugar = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[2]/div[2]/table/tbody/tr[4]/td")).Text;
            string funcionario = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[2]/div[2]/table/tbody/tr[6]/td")).Text;
            string telefono = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[2]/div[2]/table/tbody/tr[7]/td")).Text;
            string email = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[2]/div[2]/table/tbody/tr[8]/td")).Text;
            string provincia = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[2]/div[2]/table/tbody/tr[2]/td")).Text;
            string po_url = chrome.Url;
            string fecha_registro = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[3]/td")).Text;
            fecha_registro = fecha_registro.Remove(fecha_registro.Length - 12);
            DateTime RDate = Convert.ToDateTime(fecha_registro);
            fecha_registro = RDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

            string unidad_solicitante = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[6]/td")).Text,
               contactocuenta = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[4]/td")).Text,
               emailCuenta = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[1]/div[2]/table/tbody/tr[7]/td")).Text;

            //hay veces que la pagina posee un registro más entonces los adjuntos pasan del div6 al div7
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
                fecha_doc = fecha_doc.Remove(fecha_doc.Length - 12);
                oDate = Convert.ToDateTime(fecha_doc);
            }
            fecha_doc = oDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            #endregion
            #region buscar el AM de la entidad
            //buscar el AM de la entidad
            string AM = "";
            sector = "";
            string user = "";
            string sectorId = "";
            try
            {

                //si la entidad es alguna de estas se determina el cliente con base a la unidad solicitante
                if (entidad == "MINISTERIO DE EDUCACION" || entidad == "CAJA DE SEGURO SOCIAL" || entidad == "MUNICIPIO DE ANTÓN")
                {

                    System.Data.DataRow[] entityInfo = entities.Select($"entities ='{unidad_solicitante}'"); //like '%" + institu + "%'"
                    if (entityInfo.Length == 0)
                    {
                        newEntity = true;
                        newEntities.Add(unidad_solicitante);

                        entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
                        if (entityInfo.Length == 0)
                        {
                            Array.Resize(ref entityInfo, 1);
                            entityInfo[0]["salesRep"] = "MAMBURG";
                            entityInfo[0]["sectorText"] = "Public Sector";
                            entityInfo[0]["sector"] = "PS";
                            newEntity = true;
                            newEntities.Add(entidad);
                        }


                    }
                    AM = entityInfo[0]["salesRep"].ToString();
                    sector = entityInfo[0]["sectorText"].ToString();
                    sectorId = entityInfo[0]["sector"].ToString();
                    user = entityInfo[0]["salesRep"].ToString();

                }
                //busca por el email de la cuenta
                //email, unidad, entidad
                else if (entidad == "MINISTERIO DE GOBIERNO" || entidad == "MINISTERIO DE SALUD" || entidad == "Ministerio de Seguridad Pública" || entidad == "MINISTERIO DE LA PRESIDENCIA")
                {
                    MailAddress address = new MailAddress(emailCuenta);
                    string host = address.Host.ToString();
                    System.Data.DataRow[] entityInfo = entities.Select($"entities ='{host}'"); //like '%" + institu + "%'"
                    if (entityInfo.Length == 0)
                    {
                        newEntity = true;
                        newEntities.Add(host);

                        entityInfo = entities.Select($"entities ='{unidad_solicitante}'"); //like '%" + institu + "%'"
                        if (entityInfo.Length == 0)
                        {
                            newEntity = true;
                            newEntities.Add(unidad_solicitante);
                            entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
                            if (entityInfo.Length == 0)
                            {
                                Array.Resize(ref entityInfo, 1);
                                Array.Resize(ref entityInfo, 1);
                                entityInfo[0]["salesRep"] = "MAMBURG";
                                entityInfo[0]["sectorText"] = "Public Sector";
                                entityInfo[0]["sector"] = "PS";
                                newEntity = true;
                                newEntities.Add(entidad);
                            }
                        }

                    }
                    AM = entityInfo[0]["salesRep"].ToString();
                    sector = entityInfo[0]["sectorText"].ToString();
                    sectorId = entityInfo[0]["sector"].ToString();
                    user = entityInfo[0]["salesRep"].ToString();

                }
                else
                {
                    System.Data.DataRow[] entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
                    if (entityInfo.Length == 0)
                    {
                        Array.Resize(ref entityInfo, 1);
                        entityInfo[0]["salesRep"] = "MAMBURG";
                        entityInfo[0]["sectorText"] = "Public Sector";
                        entityInfo[0]["sector"] = "PS";
                        newEntity = true;
                        newEntities.Add(entidad);
                    }

                    AM = entityInfo[0]["salesRep"].ToString();
                    sector = entityInfo[0]["sectorText"].ToString();
                    sectorId = entityInfo[0]["sector"].ToString();
                    user = entityInfo[0]["salesRep"].ToString();


                }


            }
            catch (Exception)
            {

            }
            #endregion

            #region descargar los adjuntos
            string attachNames = "";
            IWebElement adjuntos;
            try
            {
                adjuntos = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[7]/div[2]/table"));
            }
            catch (Exception)
            {
                adjuntos = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[6]/div[2]/table"));
            }
            IList<IWebElement> trCollection = adjuntos.FindElements(By.TagName("tr"));
            int tdCount = trCollection.Count();

            try
            { new Actions(chrome).MoveToElement(adjuntos).Perform(); }
            catch (Exception) { }
            for (int i = 2; i <= tdCount; i++)
            {
                IWebElement pdf;
                try
                {
                    pdf = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[7]/div[2]/table/tbody/tr[{i}]/td[1]/a"));
                }
                catch (Exception)
                {
                    pdf = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[6]/div[2]/table/tbody/tr[{i}]/td[1]/a"));
                }
                string href_pdf = pdf.Text;
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

                attachNames += $", {pdf_name}";
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
                    try
                    {
                        chrome.SwitchTo().Window(chrome.WindowHandles[2]);
                        System.Threading.Thread.Sleep(1000);
                        chrome.Close();
                    }
                    catch (Exception)
                    {
                    }


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
            }
            #endregion


            #region Extraer la información de los productos y llenar exceles

            IWebElement pedidoDetalle = chrome.FindElement(By.XPath("//*[@id='elementToPDF']/div/div[4]/div[2]/table"));
            int contador_subtotal = pedidoDetalle.FindElements(By.TagName("tr")).Count;
            string sub_total = "";
            List<PoProductMacro> PoproductInfo = new List<PoProductMacro>();
            List<calculateData> datosCalculados = new List<calculateData>();
            for (int x = 3; x < contador_subtotal; x++)
            {
                PoProductMacro pInfo = new PoProductMacro();
                calculateData datoCalculados = new calculateData();
                try
                {
                    #region extraer info de producto

                    string productId = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[1]")).Text;
                    if (string.IsNullOrWhiteSpace(productId))
                    {
                        sub_total = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[3]/div")).Text.Replace("B/.\r\n", "");
                        break;
                    }
                    string producto = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[2]")).Text;
                    string cantidad = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[4]")).Text;
                    string precio_unitario = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[3]/div")).Text.Replace("B/.\r\n", "");
                    string total = chrome.FindElement(By.XPath($"//*[@id='elementToPDF']/div/div[4]/div[2]/table/tbody/tr[{x}]/td[7]/div")).Text.Replace("B/.\r\n", "");
                    #endregion
                    #region Calcular campos a base del producto
                    //extrae los dias habiles de acuerdo a la cantidad
                    int dias_h = businessDays(int.Parse(cantidad));

                    //buscar la marca
                    string marcaCode = "";
                    string marca = "";
                    try
                    {

                        System.Data.DataRow[] pBrand = productBrand.Select($"productCode ='{productId}'"); //like '%" + institu + "%'"
                        marcaCode = pBrand[0]["brand"].ToString();
                        marca = pBrand[0]["brandText"].ToString();
                    }
                    catch (Exception)
                    {
                        marcaCode = "1";
                        marca = "Lenovo";
                    }

                    if (marca != "")
                    { brandOpp = (marca == "TrippLite") ? "U" : marca.Substring(0, 1); }

                    //cantidad para text opp
                    cantidadOpp += "," + "R" + brandOpp + productId + "-" + cantidad;

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

                    //extraer el mayor de todos ya que es el que manda
                    datoCalculados.deliveryDay = dias_h;
                    datoCalculados.maximumDeliveryDate = fecha_max_entrega;
                    datoCalculados.daysRemaining = dias_faltantes;
                    datosCalculados.Add(datoCalculados);
                    #endregion
                    #region Agregar info al excel

                    console.WriteLine(DateTime.Now + " > > > " + "     Agregar informacion al excel, producto " + producto);
                    DataRow rRow = excelResult.Rows.Add();
                    rRow["Registro Unico de Pedido"] = singleOrderRecord;
                    rRow["Sector"] = sector;
                    rRow["Convenio"] = convenio;
                    rRow["Entidad"] = entidad;
                    rRow["Producto/Servicio"] = producto;
                    rRow["Cantidad de Producto"] = cantidad;
                    rRow["Marca del Producto"] = marca;
                    rRow["Total del Producto"] = total;
                    rRow["Sub Total de la orden"] = sub_total;
                    rRow["Orden de Compra"] = "";
                    rRow["Fecha de Registro"] = RDate.Date.ToString("dd/MM/yyyy");
                    rRow["Fecha de Publicacion"] = oDate.Date.ToString("dd/MM/yyyy");
                    rRow["Fianza por cumplimiento"] = "No Aplica";
                    rRow["Oportunidad"] = "";
                    rRow["Quote"] = "";
                    rRow["Tipo de Pedido"] = "";
                    rRow["Sales Order"] = "";
                    rRow["Estado de GBM"] = "Pendiente de Procesar";
                    rRow["Estado de Orden"] = "Refrendado";
                    rRow["Dias de Entrega"] = dias_h.ToString();
                    rRow["Fecha Maxima de Entrega"] = fecha_max_entrega.ToString("dd/MM/yyyy");
                    rRow["Dias Faltantes"] = dias_faltantes.ToString();
                    rRow["Forecast"] = "";
                    rRow["Provincia"] = provincia;
                    rRow["Lugar de Entrega"] = lugar;
                    rRow["Contacto de la Empresa"] = funcionario;
                    rRow["Telefono del Contacto"] = telefono;
                    rRow["Email del Contacto"] = email;
                    rRow["Confirmación de Orden"] = "";
                    rRow["Fecha Real de Entrega"] = "";
                    rRow["Monto de Multa"] = "";
                    rRow["Unidad Solicitante"] = unidad_solicitante;
                    rRow["Contacto Cuenta"] = contactocuenta;
                    rRow["Email Cuenta"] = emailCuenta;
                    rRow["Tipo de Forecast"] = "E2E";
                    rRow["Vendor Order"] = "";
                    rRow["Nombre del adjunto"] = attachNames;
                    rRow["Account Manager Asignado"] = user;
                    rRow["Link al documento"] = po_url;
                    rRow["Comentarios"] = "";

                    excelResult.AcceptChanges();

                    if (marca == "Cisco")
                    {
                        AMs[AM] = user;
                        poInfo.isCisco = true;
                        DataRow cRow = excelCisco.Rows.Add();
                        cRow = rRow;
                        excelCisco.AcceptChanges();
                    }

                    #endregion
                    #region Llenar lista de productos

                    pInfo.singleOrderRecord = singleOrderRecord;
                    pInfo.productCode = productId;
                    pInfo.quantity = cantidad;
                    pInfo.totalProduct = total;
                    pInfo.orderType = "";
                    pInfo.active = "1";
                    pInfo.createdBy = "databot";
                    PoproductInfo.Add(pInfo);

                    #endregion
                    //log.LogdeCambios("Creacion", roots.BD_Proceso, "Licitaciones Panama", "Crear reporte de convenios de Competencia Panama", vendor_text + ": " + producto, registroUnicoPedido);

                }
                catch (Exception ex)
                { }
            }

            #endregion


            #region Insertar la información en purchaseOrderMacro y purchaseOrderProduct
            //extrae los valores mayores
            double maxDaysRemaining = datosCalculados.Max(t => t.daysRemaining);
            DateTime maxMaximumDeliveryDate = datosCalculados.Max(t => t.maximumDeliveryDate);
            double maxDeliveryDay = datosCalculados.Max(t => t.deliveryDay);

            Dictionary<string, string> newrow = new Dictionary<string, string>
            {
                ["singleOrderRecord"] = singleOrderRecord,
                ["sector"] = sectorId,
                ["agreement"] = convenio,
                ["entity"] = entidad.Replace("'", ""),
                ["orderSubtotal"] = sub_total,
                ["purchaseOrder"] = "",
                ["registrationDate"] = RDate.Date.ToString("yyyy-MM-dd"),
                ["publicationDate"] = oDate.Date.ToString("yyyy-MM-dd"),
                ["performanceBond"] = "1",
                ["oportunity"] = "",
                ["quote"] = "",
                ["salesOrder"] = "",
                ["deliveryDay"] = maxDeliveryDay.ToString(),
                ["maximumDeliveryDate"] = maxMaximumDeliveryDate.ToString("yyyy-MM-dd"),
                ["daysRemaining"] = maxDaysRemaining.ToString(),
                ["forecast"] = DateTime.MinValue.Date.ToString("yyyy-MM-ddy"),
                ["forecastType"] = "4", //E2E
                ["gbmStatus"] = "3",
                ["orderStatus"] = "1",
                ["state"] = provincia.Replace("'", ""),
                ["deliveryLocation"] = lugar.Replace("'", ""),
                ["deliveryContact"] = funcionario.Replace("'", ""),
                ["phone"] = telefono,
                ["email"] = email,
                ["orderConfirmation"] = "",
                ["actualDeliveryDate"] = DateTime.MinValue.Date.ToString("yyyy-MM-ddy"),
                ["fineAmount"] = "",
                ["requestingUnit"] = unidad_solicitante.Replace("'", ""),
                ["accountContact"] = contactocuenta,
                ["emailAccount"] = emailCuenta,
                ["vendorOrder"] = "",
                ["documentLink"] = po_url,
                ["comment"] = "",
                ["active"] = "1",
                ["createdBy"] = "databot"
            };




            //true todo bien, false significa que dio un error
            bool add_sql = lpsql.insertInfoPurchaseOrder(newrow, "purchaseOrderMacro");
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
                bool addProducts = lpsql.insertInfoProductsMacro(PoproductInfo);
                if (!addProducts)
                {
                    //adjunto los URL de cada PO que dio error agregando la info a la base de datos
                    //para enviarlo por email y agregarla
                    resp_sql = resp_sql + po_url + "<br>";
                    resp_add_sql = false;
                }
            }

            #endregion

            poInfo.excel = excelResult.Copy();
            poInfo.excelCisco = excelCisco.Copy();
            poInfo.adjunto = adjunto;
            poInfo.AMs = AMs;
            poInfo.cantidadOpp = cantidadOpp;
            poInfo.contAdj = cont_adj;
            poInfo.newEntities = newEntities;

            return poInfo;
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
        public string oppCreate(Dictionary<string, string> campos)
        {
            string idopp = "";
            try
            {
                RfcDestination destination = sap.GetDestRFC("CRM");
                console.WriteLine(DateTime.Now + " > > > " + " Conectado con SAP CRM");
                RfcRepository repo = destination.Repository;
                IRfcFunction func = repo.CreateFunction("ZOPP_VENTAS");
                IRfcTable general = func.GetTable("GENERAL");
                IRfcTable partners = func.GetTable("PARTNERS");
                // IRfcTable items = func.GetTable("ITEMS");
                console.WriteLine(DateTime.Now + " > > > " + " Llenando informacion general de oportunidad");
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
                console.WriteLine(DateTime.Now + " > > > " + " Llenando informacion de cliente y equipo de ventas");
                partners.Append();
                partners.SetValue("PARTNER", campos["Cliente"].ToString());
                partners.SetValue("FUNCTION", "00000021");
                partners.Append();
                partners.SetValue("PARTNER", campos["Contacto"].ToString());
                partners.SetValue("FUNCTION", "00000015");
                partners.Append();
                partners.SetValue("PARTNER", campos["Usuario"].ToString());
                partners.SetValue("FUNCTION", "00000014");
                console.WriteLine(DateTime.Now + " > > > " + " Llenando Org de Servicios y Ventas");
                //console.WriteLine(DateTime.Now + " > > > " + " Items data has been added");
                func.SetValue("SALES_ORG", campos["OrgVentas"].ToString());
                //console.WriteLine(DateTime.Now + " > > > " + " Sales Org data has been added");
                func.SetValue("SRV_ORG", campos["OrgServicios"].ToString());
                //console.WriteLine(DateTime.Now + " > > > " + " Service Org data has been added");
                console.WriteLine(DateTime.Now + " > > > " + " Creando Oportunidad en SAP CRM");
                func.Invoke(destination);



                IRfcTable validate = func.GetTable("VALIDATE");

                if (func.GetValue("RESPONSE").ToString() != "")
                {
                    console.WriteLine(DateTime.Now + " > > > " + " Response of the request: " + func.GetValue("RESPONSE").ToString());
                }
                if (func.GetValue("OPP_ID").ToString() != "")
                {
                    console.WriteLine(DateTime.Now + " > > > " + " ID de la oportunidad creada: " + func.GetValue("OPP_ID").ToString());
                    idopp = func.GetValue("OPP_ID").ToString();
                }
                else
                {
                    idopp = "Error: creating the opportunity";
                    console.WriteLine(DateTime.Now + " > > > " + " Error creating the opportunity");
                }
                for (int i = 0; i < validate.Count; i++)
                {
                    console.WriteLine(DateTime.Now + " > > > " + " Generated errors:");
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

    }
    public class PoInfoMacro
    {
        public DataTable excel { get; set; }
        public DataTable excelCisco { get; set; }
        public List<PoProductMacro> PoproductInfo { get; set; }
        public bool isCisco { get; set; }
        public string[] adjunto { get; set; }
        public Dictionary<string, string> AMs { get; set; }
        public string cantidadOpp { get; set; }
        public int contAdj { get; set; }
        public List<string> newEntities { get; set; }
    }
}
