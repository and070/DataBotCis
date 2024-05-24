using ClosedXML.Excel;
using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Routing;

namespace DataBotV5.Automation.RPA2.LenovoReports
{
    class BsLenovoInPlantReport
    {
        ProcessInteraction proc = new ProcessInteraction();
        Rooting root = new Rooting();
        MsExcel MsExcel = new MsExcel();
        MailInteraction mail = new MailInteraction();
        Credentials cred = new Credentials();
        Log log = new Log();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        Stats stats = new Stats();
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();
        BsSQL bsql = new BsSQL();

        string respFinal = "";

        string mandante = "ERP";
        public void Main()
        {

            if (mail.GetAttachmentEmail("Solicitudes Backlog InPlant", "Procesados", "Procesados Backlog InPlant"))
            {
                root.BDUserCreatedBy = "GAHERRERA";

                //existen 3 reportes distintos, dependiendo del subject del correo se procesa como tal:
                if (root.Subject.Contains("Reporte de Backlog / In transit Actualizados - MIAMI DIRECT"))
                {
                    //nombre del archiv: BKG - In Transit - MIAMI DIRECT.xlsx
                    //se ejecuta Lunes, Miercoles y Viernes
                    //los miercoles se manda a los admin supports
                    //cada vez que se ejecuta el robot guarda una copia con el pivot Table para enviarlo el dia viernes
                    //aqui verifica los día viernes si todos los archivos se encuentran para ser enviados los viernes
                    
                    console.WriteLine("Procesando In Plant Report");
                    inPlantReport(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    updatePlant(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
                else if (root.Subject.Contains("Transit to Miami Warehouse USA Lenovo"))
                {
                    //llegan solo los Lunes a las 7:00 am
                    //nombre del archivo: Transit to Miami Warehouse USA Lenovo 2021.xlsx
                    //lo que hace es guardar el archivo solamente ya que se envía hasta el viernes
                    console.WriteLine("Procesando Transit to Miami Report");
                    transitReport(root.FilesDownloadPath + "\\" + root.ExcelFile);

                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }

                }
                else if (root.Subject.Contains("A new version of Auxiliar de Inventarios MM MD"))
                {
                    //Llega todos los días a las 5:00 am redireccionado de Cognos por Tanya, 
                    //se ejecuta los dias
                    //nombre del archivo: Auxiliar de Inventarios MM MD.xlsx
                    if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday || DateTime.Now.DayOfWeek == DayOfWeek.Friday || DateTime.Now.DayOfWeek == DayOfWeek.Monday)
                    {
                        console.WriteLine("Procesando Inventario");
                        inventoryReport(root.FilesDownloadPath + "\\" + root.ExcelFile);

                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }

                    }
                }

            }
        }
        /// <summary>
        /// Depurar el reporte de productos lenovo en transito y en whareHouse para Business System
        /// </summary>
        /// <param name="ruta">el directorio con el nombre del archivo a manipular</param>
        private void inPlantReport(string ruta)
        {
            bool valLines = true;

            #region Abrir excel
            DataSet excelBook = MsExcel.GetExcelBook(ruta);
            DataTable excel = new DataTable();

            excel.Columns.Add("PSD", System.Type.GetType("System.DateTime"));
            excel.Columns.Add("UnitPrc", System.Type.GetType("System.Double"));
            excel.Columns.Add("Extended Prc / Revenue", System.Type.GetType("System.Double"));

            excel.Columns["PSD"].DataType = System.Type.GetType("System.DateTime");
            excel = excelBook.Tables["Backlog"];
            if (excel == null)
            {
                mail.SendHTMLMail("Error al leer la plantilla de BKG - In transit - MIAMI DIRECT, verifique el nombre de la hoja sea \"Backlog\" o bien el título de las columnas sea el correcto", new string[] { "vaarrieta@gbm.net" }, "Error al leer la plantilla de BKG - In transit - MIAMI DIRECT", new string[] { "appmanagement@gbm.net", "dmeza@gbm.net" });
                return;
            }
            #endregion
            #region extraer el sistema de documentos
            DataTable documentSystem = new DataTable();
            try
            {
                string sql = "SELECT `document_system`.documentId, `document_system`.customerName, `document_trading`.`poTrading`, document_country.countryName FROM `document_system` INNER JOIN `document_trading` on `document_system`.documentId = `document_trading`.documentId INNER JOIN document_country on document_system.country = document_country.countryCode";
                documentSystem = crud.Select(sql, "document_system");

            }
            catch (Exception EX)
            {
            }
            #endregion
            #region agregar el país, cliente y PO pais
            excel.Columns.Add("País");
            excel.Columns.Add("Cliente");
            excel.Columns.Add("PO País");
            #endregion
            #region eliminar columnas
            string[] columnsDelete = {
                "Ctry1",
                "COUNTRY",
                "PO_NUM",
                "MFG_SO_NUM",
                "SO_ITEM",
                "BRAND",
                "SOLD_TO",
                "PLANT",
                "OPEN_QTY",
                "FG_DATE",
                "INCO1",
                "SEGMENT",
                "Ship To Name",
                "CUSTOMER_NAME",
                "CRM Contract",
                "Following / WL",
                "Rango Aging",
                "IN / OUT",
                "Quarter",
                "Sold To Name",
                "MOT",
            };
            foreach (string item in columnsDelete)
            {
                excel.Columns.Remove(item);
            }
            #endregion
            #region ordenar las columnas
            Dictionary<string, int> columnas = new Dictionary<string, int>();
            columnas["País"] = 0;
            columnas["Cliente"] = 1;
            columnas["PO País"] = 2;
            columnas["CREATE_DATE"] = 3;
            columnas["CUST_PO"] = 4;
            columnas["SALES_ORDER"] = 5;
            columnas["MATERIAL"] = 6;
            columnas["ORDER_QTY"] = 7;
            columnas["FAMILY"] = 8;
            columnas["PSD"] = 9;
            columnas["UnitPrc"] = 10;
            columnas["Extended Prc / Revenue"] = 11;
            columnas["Comments"] = 12;

            foreach (KeyValuePair<string, int> pair in columnas)
            {
                string campo = pair.Key.ToString();
                int valor = pair.Value;
                excel.Columns[campo].SetOrdinal(valor);

            }
            #endregion
            #region crear excel para Administradores
            DataTable excelAdmis = excel.Copy();
            excelAdmis.Columns.Remove("UnitPrc");
            excelAdmis.Columns.Remove("Extended Prc / Revenue");
            excelAdmis.Columns.Remove("Comments");
            #endregion
            #region crear excel para pivotTable  monto total por mes de las ordenes 
            DataTable excelTotal = excel.Copy();
            excelTotal.Columns.Add("Mes");
            #endregion
            //contador para extraer la fila respectiva del excel de administradores
            int contAdmis = 0;
            Dictionary<string, DataTable> books = new Dictionary<string, DataTable>();
            //por cada fila del reporte original
            string response = "";
            foreach (DataRow row in excel.Rows)
            {

                try
                {
                    DataRow rowAdmis = excelAdmis.Rows[contAdmis];
                    DataRow rowTotal = excelTotal.Rows[contAdmis];

                    string purchOrder = "";
                    string notaF = "";

                    string salesOrder = row["SALES_ORDER"].ToString();
                    if (salesOrder.Substring(0, 2) == "46")
                    {
                        row.Delete();
                        rowAdmis.Delete();
                        rowTotal.Delete();
                        contAdmis++;
                        continue;
                    }

                    purchOrder = row["CUST_PO"].ToString();
                    if (purchOrder != "")
                    {
                        //quitar el prefijo en la PO
                        if (!int.TryParse(purchOrder, out _))
                        {
                            try
                            {
                                purchOrder = string.Concat(purchOrder.Where(c => Char.IsDigit(c)));
                                row["CUST_PO"] = purchOrder;
                                rowAdmis["CUST_PO"] = purchOrder;
                                rowTotal["CUST_PO"] = purchOrder;
                            }
                            catch (Exception) { }

                        }

                        PoInfo infoSS = getPoInfo(purchOrder, documentSystem);
                        if (infoSS == null)
                        {
                            infoSS = getPoInfoSAP(purchOrder, row["MATERIAL"].ToString(), row["ORDER_QTY"].ToString());
                        }
                        string pais = (infoSS.country == null) ? "" : infoSS.country;
                        string cliente = (infoSS.customer == "" || infoSS.customer == null) ? "STOCK MIAMI" : infoSS.customer;
                        string countryReport = (pais == "" || pais == null) ? "Miami Direct, Inc." : (pais == "Miami Direct") ? pais + ", Inc." : "GBM de " + pais + ", S. A.";
                        string doc = (infoSS.documentId == null) ? "" : infoSS.documentId;
                        row["País"] = countryReport;
                        row["Cliente"] = cliente;
                        row["PO País"] = doc;

                        rowAdmis["País"] = countryReport;
                        rowAdmis["Cliente"] = cliente;
                        rowAdmis["PO País"] = doc;

                        rowTotal["País"] = countryReport;
                        rowTotal["Cliente"] = cliente;
                        rowTotal["PO País"] = doc;


                        //CONVERTIR FECHA
                        try
                        {
                            DateTime psd = DateTime.Parse(row["PSD"].ToString());
                            double psdD = psd.ToOADate();
                            DateTime dt = DateTime.FromOADate(psdD);
                            DateTime dtAdmis = dt.AddDays(4);
                            row["PSD"] = dt;
                            rowAdmis["PSD"] = dtAdmis;

                            rowTotal["PSD"] = dt;

                            rowTotal["Mes"] = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(dt.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")));

                        }
                        catch (Exception EX)
                        {
                            //PSD no es una fecha
                            rowTotal["Mes"] = (cliente == "STOCK MIAMI") ? "No Status Stock" : "No Status Clientes";
                        }

                        //convertir a moneda
                        try
                        {
                            decimal unitPrice = Convert.ToDecimal(row["UnitPrc"].ToString());
                            decimal revenue = Convert.ToDecimal(row["Extended Prc / Revenue"].ToString());

                            row["UnitPrc"] = unitPrice;
                            row["Extended Prc / Revenue"] = revenue;

                            rowTotal["UnitPrc"] = unitPrice;
                            rowTotal["Extended Prc / Revenue"] = revenue;
                            //Los admis no llevan

                        }
                        catch (Exception)
                        {

                        }



                        DataTable book = new DataTable();
                        if (books.ContainsKey(pais))
                        {
                            book = books[pais];
                        }
                        else
                        {
                            book = excelAdmis.Clone();
                        }
                        book.Rows.Add(rowAdmis.ItemArray);
                        book.AcceptChanges();
                        books[pais] = book;




                    }
                    contAdmis++;
                }
                catch (Exception ex)
                {
                    valLines = false;
                    response = response + "<br>" + ex.Message;
                }

            }
            excel.AcceptChanges();
            excelAdmis.AcceptChanges();
            excelTotal.AcceptChanges();

            if (!valLines)
            {
                mail.SendHTMLMail(response, new string[] {"appmanagement@gbm.net"}, $"Error: Reporte Órdenes en Planta escaladas a Lenovo", new string[] { "dmeza@gbm.net" }, new string[] { ruta });
            }

            console.WriteLine("Save Excel...");

            #region guardar excel y enviar a BS
            try
            {
                XLWorkbook wb = new XLWorkbook();
                IXLWorksheet ws = wb.Worksheets.Add(excel, "BL Lenovo");
                ws.Columns().AdjustToContents();
                ws.Range($"K2:L{excel.Rows.Count + 1}").Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                ruta = root.FilesDownloadPath + $"\\GBM Report {DateTime.Now.ToString("dd.MM")}.xlsx";
                if (File.Exists(ruta)) File.Delete(ruta);
                wb.SaveAs(ruta);

                //SEND EMAIL
                string msj = $"Estimado(a) se le adjunta el reporte Órdenes en Planta escaladas a Lenovo actualizado al {DateTime.Now.ToString("dd")} {DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"))}**.";
                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Reporte Órdenes en Planta escaladas a Lenovo");
                html = html.Replace("{cuerpo}", msj);
                html = html.Replace("{contenido}", "");
                console.WriteLine("Send Email...");
                try
                {
                    BsSQL bs = new BsSQL();
                    string[] cc = bs.EmailAddress(9);
                    mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Reporte Órdenes en Planta escaladas a Lenovo {DateTime.Now.ToString("dd.MM")}", cc, new string[] { ruta });
                }
                catch (Exception ex)
                {
                    mail.SendHTMLMail("Error al responder email de reporte de Lenovo en Transito " + ex.Message, new string[] { "dmeza@gbm.net" }, "Error", null, new string[] { ruta });
                }

                root.requestDetails = msj;

            }
            catch (Exception ex)
            {
                mail.SendHTMLMail(ex.Message, new string[] {"appmanagement@gbm.net"}, $"Error: Reporte al guardar/envíar Órdenes en Planta escaladas a Lenovo", new string[] { "dmeza@gbm.net" }, new string[] { ruta });
            }
            #endregion

            #region crear los exceles por país y enviarlo

            if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
            {

                console.WriteLine("crear los exceles por país y enviarlo.");
                DataTable administrators = new DataTable();

                string sql = "SELECT * FROM `adminSupport`";
                administrators = crud.Select(sql, "business_system_db");

                foreach (KeyValuePair<string, DataTable> pair in books)
                {
                    string adminSupport = "";
                    try
                    {
                        string country = pair.Key.ToString();
                        DataTable valor = pair.Value;
                        ruta = root.FilesDownloadPath + $"\\GBM Report {country} {DateTime.Now.ToString("dd.MM")}.xlsx";
                        MsExcel.CreateExcel(valor, "BL Lenovo", ruta);
                        //send email


                        try { adminSupport = administrators.Select("country = '" + country + "'")[0]["email"].ToString(); } catch { adminSupport = ""; }

                        if (adminSupport != "")
                        {
                            string[] admisEmail = adminSupport.Split(',');
                            string msjAdmis = $"Estimado(a) se le adjunta el reporte Órdenes en Planta escaladas a Lenovo actualizado al {DateTime.Now.ToString("dd")} {DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"))}**.";
                            string htmlAdmis = Properties.Resources.emailtemplate1;
                            htmlAdmis = htmlAdmis.Replace("{subject}", "Reporte Órdenes en Planta escaladas a Lenovo");
                            htmlAdmis = htmlAdmis.Replace("{cuerpo}", msjAdmis);
                            htmlAdmis = htmlAdmis.Replace("{contenido}", "");
                            try
                            {
                                BsSQL bs = new BsSQL();
                                string[] cc = bs.EmailAddress(9);
                                mail.SendHTMLMail(htmlAdmis, admisEmail, $"Reporte Órdenes en Planta escaladas a Lenovo {DateTime.Now.ToString("dd.MM")}", cc, new string[] { ruta });
                            }
                            catch (Exception ex)
                            {
                                mail.SendHTMLMail($"Error al envíar email de reporte de Lenovo en Planta a {adminSupport} :" + ex.Message, new string[] { "dmeza@gbm.net" }, "Error", null, new string[] { ruta });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        mail.SendHTMLMail($"Error al envíar email de reporte de Lenovo en Planta a {adminSupport} :" + ex.Message, new string[] { "dmeza@gbm.net" }, "Error", null, new string[] { ruta });
                    }
                }
            }

            #endregion

            #region guardar excel para reporte del viernes
            string sheetName = "BL Lenovo";
            XLWorkbook wbTotal = new XLWorkbook();
            IXLWorksheet wsTotal = wbTotal.Worksheets.Add(excelTotal, sheetName);
            wsTotal.Columns().AdjustToContents();
            wsTotal.Range($"K2:L{excelTotal.Rows.Count + 1}").Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
            ruta = root.LenovoReports + "\\" + CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday) + "-" + DateTime.Now.ToString("yyyy");
            string name = $"\\GBM Report {DateTime.Now.ToString("dd.MM")}.xlsx";
            string pathName = ruta + name;
            //si la ruta existe borre todo para dejar solo el ultimo de la semana
            fileExist(ruta, "GBM Report");

            wbTotal.SaveAs(pathName);

            #region Crear Pivot Table
            try
            {
                PivotTableParameters pivotTableParameters = new PivotTableParameters();
                pivotTableParameters.route = pathName;
                pivotTableParameters.sourceSheetName = sheetName;
                pivotTableParameters.newSheet = true;
                pivotTableParameters.newSheetName = "Monto Total Por Mes";
                pivotTableParameters.title = "Monto Total Por Mes";
                pivotTableParameters.destinationRange = "A2";
                pivotTableParameters.firstSourceCell = "A1";
                pivotTableParameters.endSourceCell = "N" + (excelTotal.Rows.Count + 1);
                pivotTableParameters.pivotTableName = "MontoMes";
                List<PivotTableFields> ptFields = new List<PivotTableFields>();

                PivotTableFields ptField = new PivotTableFields();
                ptField.sourceColumnName = "Mes";
                ptField.pivotFieldOrientation = "xlRowField";
                ptFields.Add(ptField);

                PivotTableFields ptField2 = new PivotTableFields();
                ptField2.sourceColumnName = "Extended Prc / Revenue";
                ptField2.pivotFieldOrientation = "xlDataField";
                ptField2.consolidationFunction = "xlSum";
                ptField2.fieldName = "MES Salida";
                ptField2.numberFormat = "[$$-en-US] #,##0.00";
                ptFields.Add(ptField2);
                pivotTableParameters.pivotTableFields = ptFields;
                MsExcel.CreatePivotTable(pivotTableParameters);
            }
            catch (Exception ex)
            {
                mail.SendHTMLMail(ex.Message, new string[] {"appmanagement@gbm.net"}, $"Error: Reporte al guardar/envíar Órdenes en Planta escaladas a Lenovo", new string[] { "dmeza@gbm.net" }, new string[] { ruta });
            }
            #endregion

            #endregion

            #region verifica si se pueden envíar todos los archivos
            try
            {

                bool sendReport = sendReports();
            }
            catch (Exception ex)
            {
                mail.SendHTMLMail("Error al enviar el consolidado de reportes: <br><br>" + ex, new string[] {"appmanagement@gbm.net"}, $"Error al enviar reportes consolidados de LENOVO", new string[] { "dmeza@gbm.net" }, new string[] { ruta });
            }
            #endregion

        }

        /// <summary>
        /// homologar el detalle de planta
        /// </summary>
        /// <param name="path">el directorio con el nombre del archivo a manipular</param>
        private void updatePlant(string path)
        {
            bool valLines = true;

            #region Abrir excel
            console.WriteLine("Abriendo Excel - Actualizando Plantas SAP");
            DataSet excelBook = MsExcel.GetExcelBook(path);
            DataTable excel = new DataTable();
            DataTable plantRejected = new DataTable();
            plantRejected.Columns.Add("PurchaseOrder");
            plantRejected.Columns.Add("Material");
            plantRejected.Columns.Add("Quantity");
            plantRejected.Columns.Add("PlantLenovo");
            plantRejected.Columns.Add("PlantSap");
            plantRejected.Columns.Add("Response");


            excel = excelBook.Tables["Backlog"];
            if (excel == null)
            {
                mail.SendHTMLMail("Error al leer la plantilla de BKG - In transit - MIAMI DIRECT, verifique el nombre de la hoja sea \"Backlog\" o bien el título de las columnas sea el correcto", new string[] { "vaarrieta@gbm.net" }, "Error al leer la plantilla de BKG - In transit - MIAMI DIRECT",  new string[] { "appmanagement@gbm.net", "dmeza@gbm.net" });
                return;
            }
            #endregion

            #region Recorrer excel

            string response = "";
            sap.LogSAP("ERP");
            int totalCount = 0;
            int rejectedCount = 0;
            int modifyCount = 0;

            DataTable lenovoPlantInfo = new DataTable();
            foreach (DataRow row in excel.Rows)
            {
                string custPo = row["CUST_PO"].ToString();
                string custPoNumber = string.Concat(custPo.Where(cust => Char.IsDigit(cust)));
                console.WriteLine($"PO: {custPoNumber}");
                string material = row["MATERIAL"].ToString().Trim();
                string openQty = row["OPEN_QTY"].ToString();
                string plant = row["PLANT"].ToString();
                string salesOrder = row["SALES_ORDER"].ToString();
                string sapPlant = "";
                int scroll = 0;
                string materialSap = "";
                string qtySap = "";
                string mensaje_sap = "";
                totalCount++;
                try
                {
                    if (salesOrder.Substring(0, 2) == "46")
                    {
                        console.WriteLine("Ignorar Sales Order 46");
                        continue;
                    }

                    if (salesOrder.Substring(0, 2) == "43")
                    {
                        string inco1 = row["INCO1"].ToString();
                        string sql = $"SELECT lpi.sapPlant FROM lenovo_plant_info AS lpi WHERE lpi.salesOrderId = 43 AND lpi.plant = '{plant}'";
                        lenovoPlantInfo = crud.Select(sql, "document_system");
                        sapPlant = lenovoPlantInfo.Rows[0]["sapPlant"].ToString() + " " + inco1;
                    }

                    if (salesOrder.Substring(0, 2) == "44")
                    {
                        string sql = $"SELECT lpi.sapPlant FROM lenovo_plant_info AS lpi WHERE lpi.salesOrderId = 44 AND lpi.plant = '{plant}'";
                        lenovoPlantInfo = crud.Select(sql, "document_system");
                        sapPlant = lenovoPlantInfo.Rows[0]["sapPlant"].ToString();
                    }

                    try
                    {
                        SapVariants.frame.Iconify();
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nYMMBL";
                        SapVariants.frame.SendVKey(0);
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtYBL_DATA-EBELN_9004")).Text = custPoNumber;
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btnP1")).Press();
                        ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB3")).Select();
                    }
                    catch (Exception ex)
                    {
                        mensaje_sap = "No se encontro el numero de orden en SAP";

                        DataRow rowTable = plantRejected.Rows.Add();
                        rowTable["PurchaseOrder"] = custPoNumber;
                        rowTable["Material"] = material;
                        rowTable["PlantLenovo"] = plant;
                        rowTable["PlantSap"] = sapPlant;
                        rowTable["Quantity"] = openQty;
                        rowTable["Response"] = mensaje_sap + "<br>" + ex.Message; ;
                        rejectedCount++;
                        continue;
                    }

                    console.WriteLine($"Planta: {sapPlant}");
                    for (int i = 0; i <= 1000; i++)
                    {
                        try
                        {
                            materialSap = "";
                            qtySap = "";

                            if (i == 12)
                            {
                                ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB3/ssubSUB3:Y_MM_BACKLOG:8003/tblY_MM_BACKLOGTABCONTROL3")).VerticalScrollbar.Position = scroll;
                                i = 0;
                                scroll = scroll + 12;
                            }

                            //[columna, fila]
                            materialSap = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB3/ssubSUB3:Y_MM_BACKLOG:8003/tblY_MM_BACKLOGTABCONTROL3/txtWA_YBL_DATA1-MATNR[2," + i + "]")).Text.ToString();
                            qtySap = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB3/ssubSUB3:Y_MM_BACKLOG:8003/tblY_MM_BACKLOGTABCONTROL3/txtWA_YBL_DATA1-MENGE[3," + i + "]")).Text.ToString();
                            qtySap = qtySap.Replace(",000", "");
                            if (materialSap.Replace("_", "") == "" || qtySap.Replace("_", "") == "")
                            {
                                mensaje_sap = "No se encontro el material o la cantidad en SAP";

                                DataRow rowTable = plantRejected.Rows.Add();
                                rowTable["PurchaseOrder"] = custPoNumber;
                                rowTable["Material"] = material;
                                rowTable["PlantLenovo"] = plant;
                                rowTable["PlantSap"] = sapPlant;
                                rowTable["Quantity"] = openQty;
                                rowTable["Response"] = mensaje_sap;
                                rejectedCount++;
                                console.WriteLine(mensaje_sap);
                                break;
                            }

                            if (materialSap == material && qtySap == openQty)
                            {
                                try
                                {
                                    ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB3/ssubSUB3:Y_MM_BACKLOG:8003/tblY_MM_BACKLOGTABCONTROL3/cmbWA_YBL_DATA1-NAMPL[8," + i + "]")).Key = sapPlant;
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                                    SAPFEWSELib.GuiFrameWindow frame1 = (SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[1]");
                                    frame1.Iconify();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();
                                    modifyCount++;
                                    console.WriteLine("Planta Actualizada");
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    valLines = false;

                                    DataRow rowTable = plantRejected.Rows.Add();
                                    rowTable["PurchaseOrder"] = custPoNumber;
                                    rowTable["Material"] = material;
                                    rowTable["PlantLenovo"] = plant;
                                    rowTable["PlantSap"] = sapPlant;
                                    rowTable["Quantity"] = openQty;
                                    rowTable["Response"] = $"La planta generada por la plantilla no se encuentra en SAP: {sapPlant}";
                                    rejectedCount++;
                                    console.WriteLine($"La planta generada por la plantilla no se encuentra en SAP: {sapPlant}");
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            valLines = false;
                            response = response + "<br>" + ex.Message;

                            DataRow rowTable = plantRejected.Rows.Add();
                            rowTable["PurchaseOrder"] = custPoNumber;
                            rowTable["Material"] = material;
                            rowTable["PlantLenovo"] = plant;
                            rowTable["PlantSap"] = sapPlant;
                            rowTable["Quantity"] = openQty;
                            rowTable["Response"] = ex.Message;
                            rejectedCount++;
                            console.WriteLine(ex.Message);
                            break;
                        }

                    }

                }
                catch (Exception ex)
                {
                    valLines = false;
                    response = response + "<br>" + ex.Message;
                    DataRow rowTable = plantRejected.Rows.Add();
                    rowTable["PurchaseOrder"] = custPoNumber;
                    rowTable["Material"] = material;
                    rowTable["PlantLenovo"] = plant;
                    rowTable["PlantSap"] = sapPlant;
                    rowTable["Quantity"] = openQty;
                    rowTable["Response"] = ex.Message;
                    rejectedCount++;
                    console.WriteLine(ex.Message);
                }

            }
            sap.KillSAP();
            #endregion

            plantRejected.AcceptChanges();

            if (!valLines)
            {
                mail.SendHTMLMail(response, new string[] {"appmanagement@gbm.net"}, $"Error: Reporte carga de datos de planta en BACKLOG", new string[] { "dmeza@gbm.net" }, new string[] { path });
            }

            #region envio correo resultados
            string route = root.FilesDownloadPath + "\\" + "sapPlantRejected.xlsx";
            MsExcel.CreateExcel(plantRejected, "sheet1", route);
            string[] cc = bsql.EmailAddress(12);

            string msj = $"Estimado/a, adjunto el archivo Excel con los casos de error generados durante la carga de plantas en BACKLOG, conforme al informe de Lenovo.<br> <br>Detalle de la actualización. <br> Número total de ordenes analizadas: {totalCount}<br> Número total de plantas actualizadas: {modifyCount}<br> Número total de registros con fallos: {rejectedCount}";
            string html = Properties.Resources.emailtemplate1;
            html = html.Replace("{subject}", "Informe carga plantas BACKLOG");
            html = html.Replace("{cuerpo}", msj);
            html = html.Replace("{contenido}", "");
            console.WriteLine("Send Email...");
            mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Reporte carga de datos de planta en BACKLOG", cc, new string[] { route });
            #endregion
        }

        private void transitReport(string ruta)
        {

            bool valLines = true;

            #region Abrir excel
            DataTable excel = MsExcel.GetExcel(ruta);
            excel.Columns.Add("MES ETA");
            if (excel == null)
            {
                mail.SendHTMLMail("Error al leer la plantilla de Transit to Miami Warehouse USA Lenovo, verifique el nombre de la hoja sea \"Página1_1\" o bien el título de las columnas sea el correcto", new string[] { "vaarrieta@gbm.net" }, "Error al leer la plantilla de Transit to Miami Warehouse USA Lenovo", new string[] { "appmanagement@gbm.net", "dmeza@gbm.net" });
                return;
            }
            #endregion
            excel.AcceptChanges();
            foreach (DataRow row in excel.Rows)
            {
                try
                {
                    //DataRow rowChange = excel.Rows[cont];

                    string codProduct = row["Cod Product"].ToString();
                    if (string.IsNullOrWhiteSpace(codProduct))
                    {
                        row.Delete();
                    }
                    if (codProduct.Substring(0, 1) == "5" || codProduct.Substring(0, 1) == "7" || codProduct.Length < 9)
                    {
                        //se elimina la fila
                        row.Delete();
                        //cont--;
                    }
                    else
                    {
                        try
                        {
                            DateTime psd = DateTime.Parse(row["Fecha Est Miami"].ToString());//.Replace(" 00:00:00", "")).ToString("yyyy-MM-dd");
                            row["MES ETA"] = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(psd.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"))) + ", " + psd.ToString("yyyy");
                        }
                        catch (Exception EX)
                        {
                            row["MES ETA"] = "No Date";
                        }
                    }

                    //excelSource.AcceptChanges();

                    //cont++;
                }
                catch (Exception ex)
                {

                }
            }
            excel.AcceptChanges();
            string sheetName = "Página1_1";
            //ruta = root.FilesDownloadPath;
            string name = $"\\Transit to Miami Warehouse USA Lenovo {DateTime.Now.ToString("dd")} {CultureInfo.InvariantCulture.TextInfo.ToTitleCase(DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")))}.xlsx";
            string pathName = ruta + name;
            //MsExcel.CreateExcel(excel, sheetName, pathName);

            #region Guardar para envíar con el resto
            ruta = root.LenovoReports + "\\" + CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday) + "-" + DateTime.Now.ToString("yyyy");
            pathName = ruta + name;
            //si la ruta existe borre todo para dejar solo el ultimo de la semana
            fileExist(ruta, "Transit to Miami Warehouse USA Lenovo");

            MsExcel.CreateExcel(excel, sheetName, pathName);
            #endregion

            #region Crear Pivot Table

            try
            {
                PivotTableParameters pivotTableParameters = new PivotTableParameters();
                pivotTableParameters.route = ruta + name;
                pivotTableParameters.sourceSheetName = sheetName;
                pivotTableParameters.newSheet = true;
                pivotTableParameters.newSheetName = "Monto Total";
                pivotTableParameters.title = "Monto Total Por Mes y Status";
                pivotTableParameters.destinationRange = "A2";
                pivotTableParameters.firstSourceCell = "A1";
                pivotTableParameters.endSourceCell = "P" + (excel.Rows.Count + 1);
                pivotTableParameters.pivotTableName = "MontoMes";
                List<PivotTableFields> ptFields = new List<PivotTableFields>();

                PivotTableFields ptField = new PivotTableFields();
                ptField.sourceColumnName = "MES ETA";
                ptField.pivotFieldOrientation = "xlRowField";
                ptFields.Add(ptField);

                PivotTableFields ptField1 = new PivotTableFields();
                ptField1.sourceColumnName = "Status";
                ptField1.pivotFieldOrientation = "xlColumnField";
                ptFields.Add(ptField1);

                PivotTableFields ptField2 = new PivotTableFields();
                ptField2.sourceColumnName = "Trd Po Total Cost";
                ptField2.pivotFieldOrientation = "xlDataField";
                ptField2.consolidationFunction = "xlSum";
                ptField2.fieldName = "MONTO TOTAL POR MES Y STATUS";
                ptField2.numberFormat = "[$$-en-US] #,##0.00";
                ptFields.Add(ptField2);
                pivotTableParameters.pivotTableFields = ptFields;


                bool pivotCreate = MsExcel.CreatePivotTable(pivotTableParameters);
                root.requestDetails = "Tabla dinámica creada con éxito";

            }
            catch (Exception ex)
            {
                mail.SendHTMLMail("Error al leer la plantilla de Transit to Miami Warehouse USA Lenovo <BR>" + ex.Message, new string[] {"vaarrieta@gbm.net"}, "Error en Transit to Miami Warehouse USA Lenovo", new string[] { "appmanagement@gbm.net", "dmeza@gbm.net" });
            }

            #endregion

        }

        private void inventoryReport(string ruta)
        {

            bool valLines = true;
            #region Workarround para eliminar primer fila de titulo
            removeFirstRow(ruta);
            #endregion
            #region Abrir excel
            DataTable excel = MsExcel.GetExcel(ruta);
            if (excel == null)
            {
                mail.SendHTMLMail("Error al leer la plantilla de Auxiliar de Inventarios MM MD, verifique el nombre de la hoja sea \"Page1_1\" o bien el título de las columnas sea el correcto", new string[] {"vaarrieta@gbm.net"}, "Error al leer la plantilla de Auxiliar de Inventarios MM MD", new string[] { "appmanagement@gbm.net", "dmeza@gbm.net" });
                return;
            }
            #endregion
            excel.AcceptChanges();
            foreach (DataRow row in excel.Rows)
            {
                try
                {
                    string document = row["Document"].ToString();
                    string materialDescription = row["Material Description"].ToString().ToLower();
                    string docFour = "";
                    string docSix = "";
                    if (document.Length >= 4)
                    {
                        docFour = document.Substring(document.Length - 4);
                    }
                    if (document.Length >= 6)
                    {
                        docSix = document.Substring(document.Length - 6);
                    }
                    //128448
                    if (docFour != "0325" && docSix != "104034" && docSix != "128448" && document != " ")
                    {
                        //se elimina la fila
                        row.Delete();
                    }
                    if (materialDescription.Contains("dell") || materialDescription.Contains("ibm")
                       || materialDescription.Contains("lexmark") || materialDescription.Contains("kingston")
                       || materialDescription.Contains("tripp lite") || materialDescription.Contains("toshiba"))
                    {
                        //se elimina la fila
                        row.Delete();
                    }


                }
                catch (Exception ex)
                {

                }
            }

            #region eliminar columnas
            string[] columnsDelete = {
                "Plant",
                "Storage Location",
                "Document",
                "Document Item",
                "Fecha Ultima Entrada",
                "Fecha Ultima Salida",
                "Cantidad de Días",
                "Unit of Measure",
                "Material Group",
                "Product Hierarchy",
                "Nivel Significancia",
                "De 0-90 dias",
                "De 0-91 días (USD)",
                "De 91-180 dias",
                "De 91-180 Días (USD)",
                "De 181-270 dias",
                "De 181-270 (USD)",
                "De 271-360 dias",
                "De 271-360 Días (USD)",
                "Mayor 360 dias",
                "Mayor 360 días (USD)",
                "Price Unit Local",
                "Price Local"
            };
            foreach (string item in columnsDelete)
            {
                try
                {
                    excel.Columns.Remove(item);

                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.Message);
                }
            }
            #endregion

            excel.AcceptChanges();

            #region Guardar para envíar con el resto
            string sheetName = "Page1_1";
            //ruta = root.FilesDownloadPath;
            string name = $"\\Inventario MIA {DateTime.Now.ToString("dd")} {CultureInfo.InvariantCulture.TextInfo.ToTitleCase(DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")))}.xlsx";
            //MsExcel.CreateExcel(excel, sheetName, ruta + name);
            ruta = root.LenovoReports + "\\" + CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday) + "-" + DateTime.Now.ToString("yyyy");
            //si la ruta existe borre todo para dejar solo el ultimo de la semana
            fileExist(ruta, "Inventario MIA");
            MsExcel.CreateExcel(excel, sheetName, ruta + name);
            #endregion

            root.requestDetails = "Plantilla de inventarios creada con éxito";
        }
        #region Metodos de Apoyo
        /// <summary>
        /// Extrae la PO país, cliente y país de una PO trading, a traves del item (material + cantidad)
        /// </summary>
        /// <param name="po">po trading</param>
        /// <param name="material">codigo de material</param>
        /// <param name="quantity">la cantidad del material en la po trading</param>
        /// <returns></returns>
        private PoInfo getPoInfoSAP(string po, string material, string quantity)
        {

            PoInfo info = new PoInfo();
            Dictionary<string, string> parametros = new Dictionary<string, string>();
            parametros["PO_TRADING"] = po;
            parametros["MATERIAL"] = material;
            parametros["QUANTITY"] = quantity;
            try
            {
                IRfcFunction sapFunction = sap.ExecuteRFC(mandante, "ZFI_GET_INFO_PO_TRADING", parametros);
                string resp = "";

                resp = sapFunction.GetValue("RESPONSE").ToString().Trim();
                if (resp == "El material y la cantidad no coinciden")
                {
                    info.country = "";
                    info.customer = "";
                    return info;
                }
                string poPais = sapFunction.GetValue("PO_COUNTRY").ToString().Trim();
                poPais = (poPais != "") ? (poPais.Substring(0, 4) != "1000" && poPais.Substring(0, 1) != "5") ? "" : poPais : "";
                string country = sapFunction.GetValue("COUNTRY").ToString().Trim();
                info.country = (country == "Rep. Dominicana") ? "República Dominicana" : (country == "Panama") ? "Panamá" : country;
                info.customer = sapFunction.GetValue("CUSTOMER_NAME").ToString().Trim();
                info.documentId = poPais;
            }
            catch (Exception)
            { }
            return info;
        }
        /// <summary>
        /// Extrae la PO país, cliente y país de una PO trading del sistema de documentos de S&S
        /// </summary>
        /// <param name="po">PO Trading</param>
        /// <param name="pos">El datatable con la información de document_system</param>
        /// <returns></returns>
        private PoInfo getPoInfo(string po, DataTable pos)
        {
            PoInfo info = new PoInfo();
            string pais = "";
            string cliente = "";
            string poPais = "";
            DataRow[] dr = pos.Select($"poTrading = '{po}'");
            if (dr.Length > 1)
            {
                Dictionary<string, string> vall = new Dictionary<string, string>();
                foreach (DataRow dataRow in dr)
                {
                    string customerName = dataRow["customerName"].ToString();
                    if (!vall.ContainsKey(customerName))
                    {
                        vall[customerName] = customerName;
                    }
                }
                if (vall.Count > 1)
                {
                    return null;
                }

            }
            try { pais = pos.Select("poTrading = '" + po + "'")[0]["countryName"].ToString(); } catch { pais = ""; }
            try { cliente = pos.Select("poTrading = '" + po + "'")[0]["customerName"].ToString(); } catch { cliente = ""; }
            try { poPais = pos.Select("poTrading = '" + po + "'")[0]["documentId"].ToString(); } catch { poPais = ""; }

            poPais = (poPais != "") ? (poPais.Substring(0, 4) != "1000" && poPais.Substring(0, 1) != "5") ? "" : poPais : "";
            if (pais == "")
            {
                info = null;
            }
            else
            {

                info.country = pais;
                info.customer = cliente;
                info.documentId = poPais;
            }
            return info;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">la ruta del folder donde buscar</param>
        /// <param name="fileName">el nombre del archivo que se desea buscar y eliminar</param>
        private void fileExist(string path, string fileName)
        {
            //si la carpeta de la semana existe, 
            if (Directory.Exists(path))
            {
                //verifica si existe un reporte que contenga la palabra de la variable filename
                System.IO.DirectoryInfo di = new DirectoryInfo(path);
                foreach (FileInfo file in di.EnumerateFiles())
                {
                    //si lo encuentra lo elimina, con el fin de dejar el ultimo actualizado para el reporte del viernes
                    if (file.Name.Contains(fileName))
                    {
                        file.Delete();
                    }
                }
            }
            else
            {
                //si la carpeta no existe lo crea para que pueda ser guardado el archivo
                Directory.CreateDirectory(path);
            }
        }
        public bool sendReports()
        {
            string path = root.LenovoReports + "\\" + CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday) + "-" + DateTime.Now.ToString("yyyy");
            bool valPlant = false;
            bool valTransit = false;
            bool valInventory = false;
            if (DateTime.Now.DayOfWeek == DayOfWeek.Monday)
            {
                //si la carpeta de la semana existe, 
                if (!Directory.Exists(path))
                {
                    return false;
                }
                //verifica si existe un reporte que contenga la palabra de la variable filename
                System.IO.DirectoryInfo di = new DirectoryInfo(path);
                string files = "";
                foreach (FileInfo file in di.EnumerateFiles())
                {
                    if (file.Name.Contains("GBM Report"))
                    {
                        valPlant = true;
                    }
                    if (file.Name.Contains("Transit to Miami Warehouse USA Lenovo"))
                    {
                        valTransit = true;
                    }
                    if (file.Name.Contains("Inventario MIA"))
                    {
                        valInventory = true;
                    }
                    files = files + file.FullName + ",";

                }
                if (valInventory && valTransit && valInventory)
                {
                    try
                    {
                        files = (files.Substring(files.Length - 1) == ",") ? files.Remove(files.Length - 1) : files;
                        string[] filesNames = files.Split(',');

                        string msj = $"Estimado(a) se le adjunta los reportes consolidados de Lenovo al día {DateTime.Now.ToString("dd")} {DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"))}**.";
                        string plantPivotTable = "";
                        string transitPivotTable = "";
                        foreach (string pathFile in filesNames)
                        {
                            DataSet plant = MsExcel.GetExcelBook(pathFile);
                            if (pathFile.Contains("GBM Report"))
                            {
                                DataTable excel = plant.Tables["Monto Total Por Mes"];
                                plantPivotTable = "<ul style='text-align: left !important;'><li>Órdenes en Planta</li></ul>";
                                plantPivotTable = plantPivotTable + "<br>" + ExportDatatableToHtml(excel);
                                plantPivotTable = plantPivotTable + "<br>" + $"<p>Actualizado {DateTime.Now.ToString("dd")} {DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"))}**<p/>";
                            }
                            if (pathFile.Contains("Transit to Miami Warehouse USA Lenovo"))
                            {
                                DataTable excel = plant.Tables["Monto Total"];
                                transitPivotTable = "<ul style='text-align: left !important;'><li>Órdenes en Transito a MIA y órdenes que ya se encuentran en MIA pendientes de facturar (Warehouse USA)</li></ul>";
                                transitPivotTable = transitPivotTable + "<br>" + ExportDatatableToHtml(excel);
                                transitPivotTable = transitPivotTable + "<br>" + $"<p>Actualizado {DateTime.Now.ToString("dd")} {DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"))}**<p/>";

                            }
                        }


                        string html = Properties.Resources.emailtemplate1;
                        html = html.Replace("{subject}", "Reporte Consolidados de Lenovo");
                        html = html.Replace("{cuerpo}", msj);
                        html = html.Replace("{contenido}", plantPivotTable + "<br>" + transitPivotTable);
                        BsSQL bsql = new BsSQL();
                        string[] cc = bsql.EmailAddress(9);

                        mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Reporte Consolidados de Lenovo", cc, filesNames);
                    }
                    catch (Exception ex)
                    {
                        mail.SendHTMLMail($"Error al envíar email los reportes de Lenovo" + ex.Message, new string[] { "dmeza@gbm.net" }, "Error", null, null);
                    }
                }
                return valInventory;

            }
            return false;
        }
        protected string ExportDatatableToHtml(DataTable dt)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<table class='myCustomTable' width='100 %'>"); //border='1px' cellpadding='1' cellspacing='1' bgcolor='white' style='font-family:Garamond; font-size:smaller'
            strHTMLBuilder.Append("<thead>");
            strHTMLBuilder.Append("<tr>"); // bgcolor='grey' style='color: white !important;'
            foreach (DataColumn myColumn in dt.Columns)
            {
                strHTMLBuilder.Append("<th>");
                if (!myColumn.ColumnName.Contains("Column"))
                {
                    strHTMLBuilder.Append(myColumn.ColumnName);
                }
                strHTMLBuilder.Append("</th>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</thead>");
            strHTMLBuilder.Append("<tbody>");
            foreach (DataRow myRow in dt.Rows)
            {


                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append($"<t{(((dt.Rows.IndexOf(myRow) == dt.Rows.Count - 1 || dt.Rows.IndexOf(myRow) == 0 || dt.Rows.IndexOf(myRow) == 1)) ? "h" : "d")}>");
                    strHTMLBuilder.Append(((dt.Columns.IndexOf(myColumn) == 0 || dt.Rows.IndexOf(myRow) == 0 || dt.Rows.IndexOf(myRow) == 1) ? "" : "$") + myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append($"</t{(((dt.Rows.IndexOf(myRow) == dt.Rows.Count - 1 || dt.Rows.IndexOf(myRow) == 0 || dt.Rows.IndexOf(myRow) == 1)) ? "h" : "d")}>");

                }
                strHTMLBuilder.Append("</tr>");
            }

            strHTMLBuilder.Append("</tbody>");
            strHTMLBuilder.Append("</table>");
            string Htmltext = strHTMLBuilder.ToString();
            return Htmltext;
        }
        private void removeFirstRow(string path)
        {
            XLWorkbook wb = new XLWorkbook(path);
            IXLWorksheet ws = wb.Worksheet("Page1_1");

            IXLRangeRow firstRow = ws.Row(1).AsRange().Row(1);
            if (firstRow.FirstCell().Value.ToString() == "Inventario")
            {
                firstRow.Delete();
            }
            wb.Save();
            proc.KillProcess("EXCEL", true);
        }
        #endregion
    }
    public class PoInfo
    {
        public string country { get; set; }
        public string customer { get; set; }
        public string documentId { get; set; }
    }
}
