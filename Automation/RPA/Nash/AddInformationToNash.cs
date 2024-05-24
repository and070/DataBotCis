using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using System.Globalization;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Threading;
using System.Xml.Linq;
using System.Linq;
using System.Data;
using System.IO;
using System;

namespace DataBotV5.Automation.RPA.Nash
{
    /// <summary>
    /// Clase RPA Automation encargada  de agregar información a NASH.
    /// </summary>
    class AddInformationToNash
    {
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        string respFinal = "";
        Log log = new Log();


        string endDate = DateTime.Now.Date.ToString("dd.MM.yyyy");
        string startDate = DateTime.Now.Date.AddDays(-3).ToString("dd.MM.yyyy");
        string folderPath = @"\\Nashdb\nash\GBM\Import\";

        public void Main()
        {
            if (mail.GetAttachmentEmail("Solicitudes Nash", "Procesados", "Procesados Nash"))
            {
                console.WriteLine("Procesando...");

                //string region = Thread.CurrentThread.CurrentCulture.ToString();
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-Us");

                ProcessNash(root.FilesDownloadPath + "\\" + root.ExcelFile);

                Thread.CurrentThread.CurrentCulture = new CultureInfo("es-CR");
                //region = Thread.CurrentThread.CurrentCulture.ToString();


                root.requestDetails = respFinal;
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        public void ProcessNash(string route)
        {
            #region Variables Privadas

            string mandante = "ERP";
            IRfcTable productList, vendorProdList, nashReceiptList, nashOrderList, nashUsosList, nashSalesList;

            string matGroupDesc;
            string[] lines = { };
            string[] lines2 = { };
            string[] lines3 = { };
            string[] lines4 = { };
            string[] lines5 = { };
            string[] lines6 = { };

            Program program1 = new Program();

            bool validateLines = true;
            string response = "Se han creados y cargado todos los XML de NASH. Con excepción de: <br><br>";
            string validation = "";

            Application xlApp;
            Worksheet xlWorkSheet;
            Workbook xlWorkbook;
            Worksheet xlWorksheet;

            xlApp = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            #endregion

            #region----------------------------------------( Inventario )--------------------------------------------------------------------------------

            try
            {
                console.WriteLine("Calculando Inventarios...");
                Workbook stockWorkbook = xlApp.Workbooks.Open(route);
                xlWorkSheet = (Worksheet)stockWorkbook.Sheets[1];

                validation = xlWorkSheet.Cells[1, 1].text.ToString().Trim();

                if (validation != "Storage Location")
                {
                    //string returnMsg = "Utilizar la plantilla oficial de la pagina de inventarios";
                    validateLines = false;
                }
                else
                {
                    console.WriteLine("... Plantilla Correcta");

                    xlWorkSheet.Range["1:1"].Delete();
                    int rows = xlWorkSheet.UsedRange.Rows.Count;
                    xlWorkSheet.Range[rows - 2 + ":" + rows].Delete();

                    string month = "";
                    string day = "";
                    month = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                    day = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                    string fileName1 = "TRAN" + DateTime.Now.Year + month + day + "0001.csv";
                    string fileName2 = "TRAN" + DateTime.Now.Year + month + day + "0001.xml";

                    console.WriteLine("... Guardando CSV");

                    xlApp.DisplayAlerts = false;
                    stockWorkbook.SaveAs(root.FilesDownloadPath + "\\" + fileName1, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                    stockWorkbook.Close();
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    console.WriteLine("... Creando XML");

                    lines4 = File.ReadAllLines(root.FilesDownloadPath + "\\" + fileName1);
                    XNamespace texto = "urn:transaccion-schema";
                    XElement xml = new XElement(
                        texto + "ITransaccion",
                        from str in lines4
                        let columns = str.Split(',')
                        select new XElement(texto + "ITransaccionRow",
                            new XElement(texto + "TipoTrans", "INV"),
                            new XElement(texto + "Accion", "S"),
                            new XElement(texto + "Ubicacion", columns[0]),
                            new XElement(texto + "Fecha", root.ReceivedTime.ToString("yyyy-MM-dd")),
                            new XElement(texto + "Gtin", columns[2]),
                            new XElement(texto + "UOM", columns[3]),
                            new XElement(texto + "Costo", columns[4]),
                            new XElement(texto + "Cantidad", columns[5])
                            ));

                    console.WriteLine("... Guardando XML en Server");
                    FileDelete(folderPath + fileName2);
                    xml.Save(folderPath + fileName2);

                    log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", fileName2, "con exito");
                    respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:"  + fileName2;

                }
            }
            catch (Exception ex)
            {
                console.WriteLine(" Error process " + ex.Message);
                response += "Error Inventarios: " + ex.Message + "<br>";
                validateLines = false;
            }
            #endregion

            #region----------------------------------------( Correr WS de SAP )--------------------------------------------------------------------------------
            try
            {
                console.WriteLine("Corriendo WS de SAP...");

                //***************Se van a sacar primero Productos, recibos, ordenes y por aparte usos y ventas*******************
                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["PRODUCTOS"] = "X",
                    ["RECIBOS"] = "X",
                    ["ORDENES"] = "X",
                    ["FECHA_INICIAL"] = startDate,
                    ["FECHA_FINAL"] = endDate// fecha para correr el robot (recibos 16.08.2017/xxxx) (Producto 15.03.2017/16.11.2015) (Ordenes 19.07.2017/20.08.2017)
                };
                //                                                      // Fechas (Usos 26.07.2017)  estas fechas son de Desarrollo
                //                                                      //--------correr WS------------

                IRfcFunction fm = sap.ExecuteRFC(mandante, "ZDM_GET_DATA_NASH", parameters);
                productList = fm.GetTable("PRODUCT_LIST");
                vendorProdList = fm.GetTable("PROD_PROV_LIST");
                nashReceiptList = fm.GetTable("RECIBOS_NASH_LIST");
                nashOrderList = fm.GetTable("ORDENES_NASH_LIST");
                xlApp = new Application();

                #region Procesar Salidas del FM

                #region----------------------------------------( Producto )--------------------------------------------------------------------------------
                try
                {
                    if (productList.RowCount > 0)
                    {
                        console.WriteLine("Calculando Producto...");
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        console.WriteLine("... Creando CSV");
                        for (int i = 0; i < productList.RowCount; i++)
                        {
                            xlWorksheet.Cells[i + 1, 1] = "S";
                            xlWorksheet.Cells[i + 1, 2] = productList[i].GetValue("CODIGOPROD").ToString();
                            xlWorksheet.Cells[i + 1, 3] = productList[i].GetValue("TIPO_PROD").ToString();
                            string descripcion = productList[i].GetValue("PROD_DESCR").ToString();
                            descripcion = descripcion.Replace(",", " ");
                            descripcion = descripcion.Replace("\n", " ");
                            xlWorksheet.Cells[i + 1, 4] = descripcion;
                            xlWorksheet.Cells[i + 1, 5] = " ";  //PRODUCT_LIST[i].GetValue("MAT_GROUP").ToString(); //categoria
                            xlWorksheet.Cells[i + 1, 6] = productList[i].GetValue("DESC_MAT_GROUP").ToString();
                            matGroupDesc = productList[i].GetValue("DESC_MAT_GROUP").ToString();
                            if (matGroupDesc.Length > 6)
                            {
                                xlWorksheet.Cells[i + 1, 8] = matGroupDesc.Substring(0, 6); //marca
                            }
                            else
                            {
                                xlWorksheet.Cells[i + 1, 8] = matGroupDesc;
                            }
                            string unidad = productList[i].GetValue("UOM").ToString();
                            if (unidad == "ST")
                            {
                                unidad = "UN";
                            }
                            xlWorksheet.Cells[i + 1, 9] = unidad;
                            xlWorksheet.Cells[i + 1, 10] = unidad;
                            xlWorksheet.Cells[i + 1, 14] = "U";
                            xlWorksheet.Cells[i + 1, 18] = unidad;
                            xlWorksheet.Cells[i + 1, 15] = "0";
                            string fecha_order = productList[i].GetValue("FECHA_CREACION").ToString();
                            DateTime DT = new DateTime();
                            DT = Convert.ToDateTime(fecha_order);
                            var fecha_nash = DT.ToString("yyyy-MM-dd");
                            xlWorksheet.Cells[i + 1, 16] = "'" + fecha_nash;
                            //xlWorksheet.Cells[i + 1, 17] = "1";
                            xlWorksheet.Cells[i + 1, 19] = "1";
                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "PROD" + DateTime.Now.Year + mes + dia + "0001.csv";
                        string nomarchivo2 = "PROD" + DateTime.Now.Year + mes + dia + "0001.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);
                        //Programacion XML
                        console.WriteLine("... Creando XML");
                        Program program = new Program();
                        XNamespace texto = "urn:producto-schema";
                        XElement xml = new XElement(texto + "IProducto",
                        from str in lines
                        let columns = str.Split(',')
                        select new XElement(texto + "IProductoRow",
                        new XElement(texto + "Accion", columns[0]),
                        new XElement(texto + "GtinTipo", columns[2]),
                        new XElement(texto + "Gtin", columns[1]),
                        new XElement(texto + "Descripcion", columns[3]),
                        new XElement(texto + "Categoria", columns[4]),
                        new XElement(texto + "Departamento", columns[5]),
                        new XElement(texto + "Marca", columns[7]),
                        new XElement(texto + "Uom", columns[8]),
                        new XElement(texto + "UomTamano", columns[9]),
                        new XElement(texto + "PesoUnidad", columns[13]),
                        new XElement(texto + "AceptaFracciones", columns[14]),
                        new XElement(texto + "FechaProductoNuevo", columns[15]),
                        new XElement(texto + "UOMTraspaso", columns[17]),
                        new XElement(texto + "UOMTraspasoCantidad", columns[18])
                        )
                        // new XElement(texto + "IProducto"

                        );
                        console.WriteLine("... Guardando XML en Server");
                        FileDelete(folderPath + nomarchivo2);
                        xml.Save(folderPath + nomarchivo2);
                        log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");

                        respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;
                    }
                    else
                    {
                        response += "Aviso Producto: no hay data" + "<br>";
                        console.WriteLine("Aviso Producto: no hay data");
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Producto: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #region----------------------------------------( Producto-Proveedor )--------------------------------------------------------------------------------

                try
                {
                    if (productList.RowCount > 0)
                    {
                        console.WriteLine("Calculando Producto-Proveedor...");
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        System.Data.DataTable TablaSQL = new System.Data.DataTable();
                        int contador = 0;
                        console.WriteLine("... Creando CSV");
                        for (int i = 0; i < productList.RowCount; i++)
                        {
                            string MG = productList[i].GetValue("MAT_GROUP").ToString();
                            TablaSQL = Getvendor(MG);
                            if (TablaSQL.Rows.Count > 0)
                            {
                                for (int h = 0; h <= TablaSQL.Rows.Count - 1; h++)
                                {
                                    xlWorksheet.Cells[contador + 1 + h, 1] = "S";
                                    xlWorksheet.Cells[contador + 1 + h, 2] = productList[i].GetValue("CODIGOPROD").ToString();
                                    xlWorksheet.Cells[contador + 1 + h, 3] = TablaSQL.Rows[h][2].ToString(); // proveedor
                                    xlWorksheet.Cells[contador + 1 + h, 4] = "UN";
                                    xlWorksheet.Cells[contador + 1 + h, 5] = "1";
                                    xlWorksheet.Cells[contador + 1 + h, 6] = TablaSQL.Rows[h][3].ToString(); // proveedor primario
                                    contador++;
                                }
                            }
                            else
                            {
                                //Enviar alerta a DM que el MG no esta en la base de datos
                            }
                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "PRPR" + DateTime.Now.Year + mes + dia + "0001.csv";
                        string nomarchivo2 = "PRPR" + DateTime.Now.Year + mes + dia + "0001.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);

                        Program program = new Program();
                        console.WriteLine("... Creando XML");
                        XNamespace texto = "urn:prodprov-schema";
                        XElement xml = new XElement(texto + "IProdProv",
                        from str in lines
                        let columns = str.Split(',')
                        select new XElement(texto + "IProdProvRow",
                        new XElement(texto + "Accion", columns[0]),
                        new XElement(texto + "Gtin", columns[1]),
                        new XElement(texto + "Proveedor", columns[2]),
                        new XElement(texto + "UomCaja", columns[3]),
                        new XElement(texto + "UomCajaCantidad", columns[4]),
                        new XElement(texto + "ProvPrimario", columns[5])  //nota 1 cuando es primario 0 cuando no es 
                       )

                        );
                        console.WriteLine("... Guardando XML en Server");
                        FileDelete(folderPath + nomarchivo2);
                        xml.Save(folderPath + nomarchivo2);
                        log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");
                        respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;

                    }
                    else
                    {
                        response += "Aviso Producto-Proveedor: no hay data" + "<br>";
                        console.WriteLine("Aviso Producto-Proveedor: no hay data");
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Producto-Proveedor: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #region----------------------------------------( Recibos )--------------------------------------------------------------------------------
                try
                {
                    if (nashReceiptList.RowCount > 0)
                    {
                        console.WriteLine("Calculando Recibos...");
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        console.WriteLine("... Creando CSV");
                        for (int i = 0; i < nashReceiptList.RowCount; i++)
                        {
                            xlWorksheet.Cells[i + 1, 1] = "S";
                            string fecha_order = nashReceiptList[i].GetValue("FECHARECIBO").ToString();
                            DateTime DT = new DateTime();
                            DT = Convert.ToDateTime(fecha_order);
                            var fecha_nash = DT.ToString("yyyy-MM-dd");
                            xlWorksheet.Cells[i + 1, 2] = "'" + fecha_nash;
                            xlWorksheet.Cells[i + 1, 3] = nashReceiptList[i].GetValue("UBICACION").ToString();
                            xlWorksheet.Cells[i + 1, 4] = nashReceiptList[i].GetValue("MUMPROVEEDOR").ToString();

                            xlWorksheet.Cells[i + 1, 5] = nashReceiptList[i].GetValue("MATERIALDOC").ToString();
                            string status = nashReceiptList[i].GetValue("STATUS").ToString();
                            if (status == "101")
                            {
                                status = "4";
                            }
                            xlWorksheet.Cells[i + 1, 6] = status;
                            xlWorksheet.Cells[i + 1, 7] = nashReceiptList[i].GetValue("PURCHASEORDER").ToString();
                            xlWorksheet.Cells[i + 1, 8] = nashReceiptList[i].GetValue("ITEM").ToString();
                            xlWorksheet.Cells[i + 1, 9] = " ";
                            xlWorksheet.Cells[i + 1, 10] = nashReceiptList[i].GetValue("PRODUCTO").ToString();
                            string unidad = nashReceiptList[i].GetValue("UOM").ToString();
                            if (unidad == "ST")
                            {
                                unidad = "UN";
                            }
                            xlWorksheet.Cells[i + 1, 11] = unidad;

                            float oc = float.Parse(nashReceiptList[i].GetValue("RECIBOCANTIDAD").ToString());
                            float recibo_cantidad = (float)Math.Round(oc * 100f) / 100f;
                            string recibo_cantidads = recibo_cantidad.ToString().Replace(",", ".");
                            xlWorksheet.Cells[i + 1, 12] = recibo_cantidads.ToString();
                            xlWorksheet.Cells[i + 1, 13] = "0";
                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "RECI" + DateTime.Now.Year + mes + dia + "0001.csv";
                        string nomarchivo2 = "RECI" + DateTime.Now.Year + mes + dia + "0001.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines2 = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);

                        Program program = new Program();
                        console.WriteLine("... Creando XML");
                        XNamespace texto1 = "urn:recibo-schema";
                        XElement xml_root = new XElement(texto1 + "IRecibo");     // ESTE ES EL XML PRINCIPAL
                        XElement xml_items;
                        XElement xml_encabezado;
                        IEnumerable<string> Valores_unicos = (from str in lines2 let columns = str.Split(',') select columns[4]).Distinct(); // SE TOMA SOLO LOS VALORES UNICOS (el tipo debe se el de columns)
                        foreach (var orden in Valores_unicos)   //FOR POR TODOS LOS VALORES UNICOS
                        {
                            IEnumerable<string[]> tabla1 = from str in lines2 let columns = str.Split(',') where columns[4] == orden select columns; //SE TOMA TODA LA TABLA DE LA "ORDEN" ACTUAL
                            string[] primera_linea = tabla1.ElementAt(0); //TOMO LOS DATOS DE LA PRIMERA LINEA YA QUE LAS DEMAS IGUAL SE REPITEN
                            xml_encabezado = new XElement(texto1 + "IReciboRow",                                        //SE LLENA EL FRAGMENTO DEL "ENCABEZADO"
                                                          new XElement(texto1 + "Accion", primera_linea[0]),
                                                          new XElement(texto1 + "Fecha", primera_linea[1]),
                                                          new XElement(texto1 + "Ubicacion", primera_linea[2]),
                                                          new XElement(texto1 + "Proveedor", primera_linea[3]),
                                                          new XElement(texto1 + "NumRecibo", primera_linea[4]),
                                                          new XElement(texto1 + "Estatus", primera_linea[5]),
                                                          new XElement(texto1 + "OrdenCompra", primera_linea[6])
                                                         );
                            xml_root.Add(xml_encabezado);       //SE INSERTA AL XML PRINCIPAL
                            foreach (var item in tabla1)     //FOR POR TODOS LOS ITEMS DE ESA ORDEN
                            {
                                xml_items = new XElement(texto1 + "IReciboDetalleRow",                        //SE LLENA EL FRAGMENTO DE LOS ITEMS
                                                         new XElement(texto1 + "NumLinea", item[7]),
                                                         new XElement(texto1 + "Gtin", item[9]),
                                                         new XElement(texto1 + "Uom", item[10]),
                                                         new XElement(texto1 + "Cantidad", item[11]),
                                                         new XElement(texto1 + "Costo", item[12])
                                                         );
                                xml_encabezado.Add(xml_items);    //SE INSERTA EN EL XML DEL ENCABEZADO
                            }
                        }
                        console.WriteLine("... Guardando XML en Server");
                        FileDelete(folderPath + nomarchivo2);
                        xml_root.Save(folderPath + nomarchivo2);
                        log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");
                        respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;

                    }
                    else
                    {
                        console.WriteLine("Aviso Recibos: no hay data");
                        response += "Aviso Recibos: no hay data" + "<br>";
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Recibos: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #region----------------------------------------( Recibos internos )--------------------------------------------------------------------------------
                try
                {
                    if (nashReceiptList.RowCount > 0)
                    {
                        console.WriteLine("Calculando Recibos internos...");
                        bool existe = false;
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        console.WriteLine("... Creando CSV");
                        for (int i = 0; i < nashReceiptList.RowCount; i++)
                        {
                            string vendor = nashReceiptList[i].GetValue("MUMPROVEEDOR").ToString();
                            bool vendor_GBM = false;
                            switch (vendor)
                            {
                                case "0010000663":
                                    vendor_GBM = true;
                                    break;
                                case "0010000681":
                                    vendor_GBM = true;
                                    break;
                                case "0010000731":
                                    vendor_GBM = true;
                                    break;
                                case "0010000735":
                                    vendor_GBM = true;
                                    break;
                                case "0010000799":
                                    vendor_GBM = true;
                                    break;
                                case "0010000811":
                                    vendor_GBM = true;
                                    break;
                                case "0010000829":
                                    vendor_GBM = true;
                                    break;
                                default:
                                    vendor_GBM = false;
                                    break;
                            }
                            if (vendor_GBM == true)
                            {
                                existe = true;
                                xlWorksheet.Cells[i + 1, 1] = "S";
                                xlWorksheet.Cells[i + 1, 5] = nashReceiptList[i].GetValue("MATERIALDOC").ToString();
                                xlWorksheet.Cells[i + 1, 7] = nashReceiptList[i].GetValue("PURCHASEORDER").ToString();

                                string ubicacion = nashReceiptList[i].GetValue("UBICACION").ToString();
                                ubicacion = ubicacion.Substring(0, 2) + "02";
                                xlWorksheet.Cells[i + 1, 3] = ubicacion;

                                xlWorksheet.Cells[i + 1, 4] = nashReceiptList[i].GetValue("MUMPROVEEDOR").ToString();
                                string fecha_order = nashReceiptList[i].GetValue("FECHARECIBO").ToString();
                                DateTime DT = new DateTime();
                                DT = Convert.ToDateTime(fecha_order);
                                var fecha_nash = DT.ToString("yyyy-MM-dd");
                                xlWorksheet.Cells[i + 1, 2] = "'" + fecha_nash;
                                string status = nashReceiptList[i].GetValue("STATUS").ToString();
                                if (status == "101")
                                {
                                    status = "4";
                                }
                                xlWorksheet.Cells[i + 1, 6] = status;
                                xlWorksheet.Cells[i + 1, 9] = " ";
                                xlWorksheet.Cells[i + 1, 8] = nashReceiptList[i].GetValue("ITEM").ToString();
                                xlWorksheet.Cells[i + 1, 10] = nashReceiptList[i].GetValue("PRODUCTO").ToString();
                                string unidad = nashReceiptList[i].GetValue("UOM").ToString();
                                if (unidad == "ST")
                                {
                                    unidad = "UN";
                                }
                                xlWorksheet.Cells[i + 1, 11] = unidad;
                                float oc = float.Parse(nashReceiptList[i].GetValue("RECIBOCANTIDAD").ToString());
                                float recibo_cantidad = (float)Math.Round(oc * 100f) / 100f;

                                string recibo_cantidads = recibo_cantidad.ToString().Replace(",", ".");

                                xlWorksheet.Cells[i + 1, 12] = recibo_cantidads.ToString();
                                xlWorksheet.Cells[i + 1, 13] = "0";
                            }
                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "RECI" + DateTime.Now.Year + mes + dia + "0002.csv";
                        string nomarchivo2 = "RECI" + DateTime.Now.Year + mes + dia + "0002.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines2 = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);

                        if (existe == true)
                        {
                            console.WriteLine("... Creando XML");
                            Program program = new Program();
                            XNamespace texto1 = "urn:recibo-schema";
                            XElement xml_root = new XElement(texto1 + "IRecibo");     // ESTE ES EL XML PRINCIPAL
                            XElement xml_items;
                            XElement xml_encabezado;
                            IEnumerable<string> Valores_unicos = (from str in lines2 let columns = str.Split(',') select columns[4]).Distinct(); // SE TOMA SOLO LOS VALORES UNICOS (el tipo debe se el de columns)
                            foreach (var orden in Valores_unicos)   //FOR POR TODOS LOS VALORES UNICOS
                            {
                                IEnumerable<string[]> tabla1 = from str in lines2 let columns = str.Split(',') where columns[4] == orden select columns; //SE TOMA TODA LA TABLA DE LA "ORDEN" ACTUAL
                                string[] primera_linea = tabla1.ElementAt(0); //TOMO LOS DATOS DE LA PRIMERA LINEA YA QUE LAS DEMAS IGUAL SE REPITEN
                                xml_encabezado = new XElement(texto1 + "IReciboRow",                                        //SE LLENA EL FRAGMENTO DEL "ENCABEZADO"
                                                              new XElement(texto1 + "Accion", primera_linea[0]),
                                                              new XElement(texto1 + "Fecha", primera_linea[1]),
                                                              new XElement(texto1 + "Ubicacion", primera_linea[2]),
                                                              new XElement(texto1 + "Proveedor", primera_linea[3]),
                                                              new XElement(texto1 + "NumRecibo", primera_linea[4]),
                                                              new XElement(texto1 + "Estatus", primera_linea[5]),
                                                              new XElement(texto1 + "OrdenCompra", primera_linea[6])
                                                             );
                                xml_root.Add(xml_encabezado);       //SE INSERTA AL XML PRINCIPAL
                                foreach (var item in tabla1)     //FOR POR TODOS LOS ITEMS DE ESA ORDEN
                                {
                                    xml_items = new XElement(texto1 + "IReciboDetalleRow",                        //SE LLENA EL FRAGMENTO DE LOS ITEMS
                                                             new XElement(texto1 + "NumLinea", item[7]),
                                                             new XElement(texto1 + "Gtin", item[9]),
                                                             new XElement(texto1 + "Uom", item[10]),
                                                             new XElement(texto1 + "Cantidad", item[11]),
                                                             new XElement(texto1 + "Costo", item[12])
                                                             );
                                    xml_encabezado.Add(xml_items);    //SE INSERTA EN EL XML DEL ENCABEZADO
                                }
                            }
                            console.WriteLine("... Guardando XML en Server");
                            FileDelete(folderPath + nomarchivo2);
                            xml_root.Save(folderPath + nomarchivo2); //OJO ruta servidor NASH
                            log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");
                            respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;

                        }
                        else
                        {
                            response += "Aviso Recibos internos: no hay data" + "<br>";
                            console.WriteLine("... Aviso Recibos internos: no hay data");
                        }
                        //System.IO.File.Delete(root.Google_Download + "\\" + nomarchivo);
                    }
                    else
                    {
                        response += "Aviso Recibos internos: no hay data" + "<br>";
                        console.WriteLine("Aviso Recibos internos: no hay data");
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Recibos internos: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #region----------------------------------------( Ordenes )--------------------------------------------------------------------------------
                try
                {
                    int last1;
                    if (nashOrderList.RowCount > 0)
                    {
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        for (int i = 0; i < nashOrderList.RowCount; i++)
                        {
                            string ubicacion = nashOrderList[i].GetValue("COMPANYCODE").ToString();
                            if (ubicacion.Substring(2, 2) == "02")
                            {
                                last1 = xlApp.Cells[xlApp.Rows.Count, "A"].End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row;

                                string fecha_order = nashOrderList[i].GetValue("FECHAORDEN").ToString();
                                DateTime DT = new DateTime();
                                DT = Convert.ToDateTime(fecha_order);
                                var fecha_nash = DT.ToString("yyyy-MM-dd");
                                xlWorksheet.Cells[last1 + 1, 1] = "S";
                                xlWorksheet.Cells[last1 + 1, 2] = "'" + fecha_nash;

                                xlWorksheet.Cells[last1 + 1, 3] = ubicacion;

                                xlWorksheet.Cells[last1 + 1, 4] = nashOrderList[i].GetValue("NUMPROVEEDOR").ToString();
                                xlWorksheet.Cells[last1 + 1, 7] = "'" + fecha_nash;
                                xlWorksheet.Cells[last1 + 1, 8] = nashOrderList[i].GetValue("PURCHASEORDER").ToString();
                                string status = nashOrderList[i].GetValue("STATUS").ToString();
                                if (status == "9")
                                {
                                    status = "1";
                                }
                                xlWorksheet.Cells[last1 + 1, 11] = status;
                                xlWorksheet.Cells[last1 + 1, 12] = nashOrderList[i].GetValue("POCREATOR").ToString();
                                xlWorksheet.Cells[last1 + 1, 13] = " ";
                                xlWorksheet.Cells[last1 + 1, 18] = nashOrderList[i].GetValue("POITEM").ToString();
                                xlWorksheet.Cells[last1 + 1, 19] = nashOrderList[i].GetValue("PRODUCTO").ToString();
                                string unidad = nashOrderList[i].GetValue("UOM").ToString();
                                if (unidad == "ST")
                                {
                                    unidad = "UN";
                                }
                                xlWorksheet.Cells[last1 + 1, 20] = unidad;
                                xlWorksheet.Cells[last1 + 1, 21] = nashOrderList[i].GetValue("POCANTIDAD").ToString();
                                xlWorksheet.Cells[last1 + 1, 22] = nashOrderList[i].GetValue("COSTO").ToString();
                            }

                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "ORDE" + DateTime.Now.Year + mes + dia + "0001.csv";
                        string nomarchivo2 = "ORDE" + DateTime.Now.Year + mes + dia + "0001.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines5 = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);
                        console.WriteLine("... Creando XML");
                        Program program = new Program();
                        XNamespace texto1 = "urn:orden-schema";
                        XElement xml_root = new XElement(texto1 + "IOrden");     // ESTE ES EL XML PRINCIPAL
                        XElement xml_items;
                        XElement xml_encabezado;
                        IEnumerable<string> Valores_unicos = (from str in lines5 let columns = str.Split(',') select columns[7]).Distinct(); // SE TOMA SOLO LOS VALORES UNICOS (el tipo debe se el de columns)
                        foreach (var orden in Valores_unicos)   //FOR POR TODOS LOS VALORES UNICOS
                        {
                            IEnumerable<string[]> tabla1 = from str in lines5 let columns = str.Split(',') where columns[7] == orden select columns; //SE TOMA TODA LA TABLA DE LA "ORDEN" ACTUAL
                            string[] primera_linea = tabla1.ElementAt(0); //TOMO LOS DATOS DE LA PRIMERA LINEA YA QUE LAS DEMAS IGUAL SE REPITEN
                            xml_encabezado = new XElement(texto1 + "IOrdenRow",                                        //SE LLENA EL FRAGMENTO DEL "ENCABEZADO"
                                                          new XElement(texto1 + "Accion", primera_linea[0]),
                                                          new XElement(texto1 + "Fecha", primera_linea[1]),
                                                          new XElement(texto1 + "Ubicacion", primera_linea[2]),
                                                          new XElement(texto1 + "Proveedor", primera_linea[3]),
                                                          new XElement(texto1 + "FechaEntrega", primera_linea[6]),
                                                          new XElement(texto1 + "NumOrden", primera_linea[7]),
                                                          new XElement(texto1 + "Estatus", primera_linea[10]),
                                                          new XElement(texto1 + "Comprador", primera_linea[11])
                                                         );
                            xml_root.Add(xml_encabezado);       //SE INSERTA AL XML PRINCIPAL
                            foreach (var item in tabla1)     //FOR POR TODOS LOS ITEMS DE ESA ORDEN
                            {
                                xml_items = new XElement(texto1 + "IOrdenDetalleRow",                        //SE LLENA EL FRAGMENTO DE LOS ITEMS
                                                         new XElement(texto1 + "NumLinea", item[17]),
                                                         new XElement(texto1 + "Gtin", item[18]),
                                                         new XElement(texto1 + "Uom", item[19]),
                                                         new XElement(texto1 + "Cantidad", item[20]),
                                                         new XElement(texto1 + "Costo", item[21])
                                                         );
                                xml_encabezado.Add(xml_items);    //SE INSERTA EN EL XML DEL ENCABEZADO
                            }
                        }
                        console.WriteLine("... Guardando XML en Server");
                        FileDelete(folderPath + nomarchivo2);
                        xml_root.Save(folderPath + nomarchivo2); //OJO ruta servidor NASH
                        log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");
                        respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;

                    }
                    else
                    {
                        response += "Aviso Ordenes: no hay data" + "<br>";
                        console.WriteLine("Aviso Ordenes: no hay data");
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Ordenes: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #region----------------------------------------( Ordenes internas )--------------------------------------------------------------------------------
                try
                {
                    if (nashOrderList.RowCount > 0)
                    {
                        console.WriteLine("Calculando Ordenes internas...");
                        bool existe = false;
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        for (int i = 0; i < nashOrderList.RowCount; i++)
                        {
                            string vendor = nashOrderList[i].GetValue("NUMPROVEEDOR").ToString();
                            bool vendor_GBM = false;
                            switch (vendor)
                            {
                                case "0010000663":
                                    vendor_GBM = true;
                                    break;
                                case "0010000681":
                                    vendor_GBM = true;
                                    break;
                                case "0010000731":
                                    vendor_GBM = true;
                                    break;
                                case "0010000735":
                                    vendor_GBM = true;
                                    break;
                                case "0010000799":
                                    vendor_GBM = true;
                                    break;
                                case "0010000811":
                                    vendor_GBM = true;
                                    break;
                                case "0010000829":
                                    vendor_GBM = true;
                                    break;
                                default:
                                    vendor_GBM = false;
                                    break;
                            }
                            if (vendor_GBM == true)
                            {
                                existe = true;
                                string fecha_order = nashOrderList[i].GetValue("FECHAORDEN").ToString();
                                DateTime DT = new DateTime();
                                DT = Convert.ToDateTime(fecha_order);
                                var fecha_nash = DT.ToString("yyyy-MM-dd");
                                xlWorksheet.Cells[i + 1, 1] = "S";
                                xlWorksheet.Cells[i + 1, 2] = "'" + fecha_nash;//ORDENES_NASH_LIST[i].GetValue("FECHAORDEN").ToString();
                                string ubicacion = nashOrderList[i].GetValue("COMPANYCODE").ToString();
                                ubicacion = ubicacion.Substring(0, 2) + "02";
                                xlWorksheet.Cells[i + 1, 3] = ubicacion;
                                xlWorksheet.Cells[i + 1, 4] = nashOrderList[i].GetValue("NUMPROVEEDOR").ToString();
                                xlWorksheet.Cells[i + 1, 7] = "'" + fecha_nash; //ORDENES_NASH_LIST[i].GetValue("FECHAORDEN").ToString();
                                xlWorksheet.Cells[i + 1, 8] = nashOrderList[i].GetValue("PURCHASEORDER").ToString();
                                string status = nashOrderList[i].GetValue("STATUS").ToString();
                                if (status == "9")
                                {
                                    status = "1";
                                }
                                xlWorksheet.Cells[i + 1, 11] = status;
                                xlWorksheet.Cells[i + 1, 12] = nashOrderList[i].GetValue("POCREATOR").ToString();
                                xlWorksheet.Cells[i + 1, 13] = " ";
                                xlWorksheet.Cells[i + 1, 18] = nashOrderList[i].GetValue("POITEM").ToString();
                                xlWorksheet.Cells[i + 1, 19] = nashOrderList[i].GetValue("PRODUCTO").ToString();
                                string unidad = nashOrderList[i].GetValue("UOM").ToString();
                                if (unidad == "ST")
                                {
                                    unidad = "UN";
                                }
                                xlWorksheet.Cells[i + 1, 20] = unidad;

                                float oc = float.Parse(nashOrderList[i].GetValue("POCANTIDAD").ToString());
                                float ordenes_cantidad = (float)Math.Round(oc * 100f) / 100f;
                                float oco = float.Parse(nashOrderList[i].GetValue("COSTO").ToString());
                                float ordenes_costo = (float)Math.Round(oco * 100f) / 100f;

                                string ordenes_cantidads = ordenes_cantidad.ToString().Replace(",", ".");
                                string ordenes_costos = ordenes_costo.ToString().Replace(",", ".");

                                xlWorksheet.Cells[i + 1, 21] = ordenes_cantidads.ToString();
                                xlWorksheet.Cells[i + 1, 22] = ordenes_costos.ToString();
                            }
                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "ORDE" + DateTime.Now.Year + mes + dia + "0002.csv";
                        string nomarchivo2 = "ORDE" + DateTime.Now.Year + mes + dia + "0002.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines5 = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);

                        if (existe == true)
                        {
                            console.WriteLine("... Creando XML");
                            XNamespace texto1 = "urn:orden-schema";
                            XElement xml_root = new XElement(texto1 + "IOrden");     // ESTE ES EL XML PRINCIPAL
                            XElement xml_items;
                            XElement xml_encabezado;
                            IEnumerable<string> Valores_unicos = (from str in lines5 let columns = str.Split(',') select columns[7]).Distinct(); // SE TOMA SOLO LOS VALORES UNICOS (el tipo debe se el de columns)
                            foreach (var orden in Valores_unicos)   //FOR POR TODOS LOS VALORES UNICOS
                            {
                                IEnumerable<string[]> tabla1 = from str in lines5 let columns = str.Split(',') where columns[7] == orden select columns; //SE TOMA TODA LA TABLA DE LA "ORDEN" ACTUAL
                                string[] primera_linea = tabla1.ElementAt(0); //TOMO LOS DATOS DE LA PRIMERA LINEA YA QUE LAS DEMAS IGUAL SE REPITEN
                                xml_encabezado = new XElement(texto1 + "IOrdenRow",                                        //SE LLENA EL FRAGMENTO DEL "ENCABEZADO"
                                                              new XElement(texto1 + "Accion", primera_linea[0]),
                                                              new XElement(texto1 + "Fecha", primera_linea[1]),
                                                              new XElement(texto1 + "Ubicacion", primera_linea[2]),
                                                              new XElement(texto1 + "Proveedor", primera_linea[3]),
                                                              new XElement(texto1 + "FechaEntrega", primera_linea[6]),
                                                              new XElement(texto1 + "NumOrden", primera_linea[7]),
                                                              new XElement(texto1 + "Estatus", primera_linea[10]),
                                                              new XElement(texto1 + "Comprador", primera_linea[11])
                                                             );
                                xml_root.Add(xml_encabezado);       //SE INSERTA AL XML PRINCIPAL
                                foreach (var item in tabla1)     //FOR POR TODOS LOS ITEMS DE ESA ORDEN
                                {
                                    xml_items = new XElement(texto1 + "IOrdenDetalleRow",                        //SE LLENA EL FRAGMENTO DE LOS ITEMS
                                                             new XElement(texto1 + "NumLinea", item[17]),
                                                             new XElement(texto1 + "GTIN", item[18]),
                                                             new XElement(texto1 + "Uom", item[19]),
                                                             new XElement(texto1 + "Cantidad", item[20]),
                                                             new XElement(texto1 + "Costo", item[21])
                                                             );
                                    xml_encabezado.Add(xml_items);    //SE INSERTA EN EL XML DEL ENCABEZADO
                                }
                            }
                            console.WriteLine("... Guardando XML en Server");
                            FileDelete(folderPath + nomarchivo2);
                            xml_root.Save(folderPath + nomarchivo2); //OJO ruta servidor NASH
                            log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");
                            respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;

                        }
                        else
                        {
                            response += "Aviso Ordenes internas: no hay data" + "<br>";
                            console.WriteLine("... Aviso Ordenes internas: no hay data");
                        }
                        //System.IO.File.Delete(root.Google_Download + "\\" + nomarchivo);
                    }
                    else
                    {
                        response += "Aviso Ordenes internas: no hay data" + "<br>";
                        console.WriteLine("Aviso Ordenes internas: no hay data");
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Ordenes internas: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #region----------------------------------------( Traspaso )--------------------------------------------------------------------------------
                try
                {
                    if (nashOrderList.RowCount > 0)
                    {
                        console.WriteLine("Calculando Traspaso...");
                        bool existe = false;
                        xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                        console.WriteLine("... Creando CSV");
                        for (int i = 0; i < nashOrderList.RowCount; i++)
                        {
                            string vendor = nashOrderList[i].GetValue("NUMPROVEEDOR").ToString();
                            bool vendor_GBM = false;
                            switch (vendor)
                            {
                                case "0010000663":
                                    vendor_GBM = true;
                                    break;
                                case "0010000681":
                                    vendor_GBM = true;
                                    break;
                                case "0010000731":
                                    vendor_GBM = true;
                                    break;
                                case "0010000735":
                                    vendor_GBM = true;
                                    break;
                                case "0010000799":
                                    vendor_GBM = true;
                                    break;
                                case "0010000811":
                                    vendor_GBM = true;
                                    break;
                                case "0010000829":
                                    vendor_GBM = true;
                                    break;
                                default:
                                    vendor_GBM = false;
                                    break;
                            }
                            if (vendor_GBM == true)
                            {
                                existe = true;
                                string fecha_order = nashOrderList[i].GetValue("FECHAORDEN").ToString();
                                DateTime DT = new DateTime();
                                DT = Convert.ToDateTime(fecha_order);
                                var fecha_nash = DT.ToString("yyyy-MM-dd");
                                xlWorksheet.Cells[i + 1, 1] = "TRP";
                                xlWorksheet.Cells[i + 1, 2] = "S";
                                xlWorksheet.Cells[i + 1, 4] = "'" + fecha_nash; //ORDENES_NASH_LIST[i].GetValue("FECHAORDEN").ToString();
                                string ubicacion = nashOrderList[i].GetValue("COMPANYCODE").ToString();
                                ubicacion = ubicacion.Substring(0, 2) + "02";
                                xlWorksheet.Cells[i + 1, 3] = ubicacion;
                                xlWorksheet.Cells[i + 1, 5] = nashOrderList[i].GetValue("PRODUCTO").ToString();
                                string unidad = nashOrderList[i].GetValue("UOM").ToString();
                                if (unidad == "ST")
                                {
                                    unidad = "UN";
                                }
                                xlWorksheet.Cells[i + 1, 6] = unidad;

                                float oc = float.Parse(nashOrderList[i].GetValue("POCANTIDAD").ToString());
                                float ordenes_cantidad = (float)Math.Round(oc * 100f) / 100f;
                                float oco = float.Parse(nashOrderList[i].GetValue("COSTO").ToString());
                                float ordenes_costo = (float)Math.Round(oco * 100f) / 100f;

                                string ordenes_cantidads = ordenes_cantidad.ToString().Replace(",", ".");
                                string ordenes_costos = ordenes_costo.ToString().Replace(",", ".");

                                xlWorksheet.Cells[i + 1, 7] = ordenes_cantidads.ToString();
                                xlWorksheet.Cells[i + 1, 8] = ordenes_costos.ToString();

                                xlWorksheet.Cells[i + 1, 9] = "1";
                            }
                        }
                        xlApp.DisplayAlerts = false;
                        string mes = "";
                        string dia = "";
                        mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                        dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                        string nomarchivo = "TRAN" + DateTime.Now.Year + mes + dia + "0003.csv";
                        string nomarchivo2 = "TRAN" + DateTime.Now.Year + mes + dia + "0003.xml";
                        xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nomarchivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        xlWorkbook.Close();
                        xlApp.Quit();
                        lines5 = File.ReadAllLines(root.FilesDownloadPath + "\\" + nomarchivo);

                        if (existe == true)
                        {
                            console.WriteLine("... Creado XML");
                            Program program = new Program();
                            XNamespace texto = "urn:orden-schema";
                            XElement xml = new XElement(texto + "ITransaccion",
                            from str in lines5
                            let columns = str.Split(',')
                            select new XElement(texto + "ITransaccionRow",
                            new XElement(texto + "TipoTrans", columns[0]),
                            new XElement(texto + "Accion", columns[1]),
                            new XElement(texto + "FechaOrden", GetDate(columns[3])),
                            new XElement(texto + "Ubicacion", columns[2]),
                            new XElement(texto + "GTIN", columns[4]),
                            new XElement(texto + "UOM", columns[5]),
                            new XElement(texto + "OrdenCantidad", columns[6]),
                            new XElement(texto + "OrdenCosto", columns[7]),
                            new XElement(texto + "Picks", columns[8])
                            )

                            );
                            console.WriteLine("... Guardando XML en Server");
                            FileDelete(folderPath + nomarchivo2);
                            xml.Save(folderPath + nomarchivo2); //OJO ruta servidor NASH
                            log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nomarchivo2, "con exito");
                            respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nomarchivo2;

                        }
                        else
                        {
                            response += "Aviso Traspaso: no hay data" + "<br>";
                            console.WriteLine("... Aviso Traspaso: no hay data");
                        }
                        //System.IO.File.Delete(root.Google_Download + "\\" + nomarchivo);
                    }
                    else
                    {
                        response += "Aviso Traspaso: no hay data" + "<br>";
                        console.WriteLine("Aviso Traspaso: no hay data");
                    }
                }
                catch (Exception ex)
                {
                    console.WriteLine(" Error process " + ex.Message);
                    response += "Error Traspaso: " + ex.Message + "<br>";
                    validateLines = false;
                }
                #endregion

                #endregion
            }
            catch (Exception ex)
            {
                console.WriteLine(" Error process " + ex.Message);
                response += "Error Sacando información de WS SAP: " + ex.Message + "<br>";
                validateLines = false;
            }
            #endregion

            #region----------------------------------------( Usos )--------------------------------------------------------------------------------
            try
            {
                console.WriteLine("Corriendo WS de SAP para Usos...");

                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["USOS"] = "X",
                    ["FECHA_INICIAL"] = startDate,
                    ["FECHA_FINAL"] = endDate// fecha para correr el robot (recibos 16.08.2017/xxxx) (Producto 15.03.2017/16.11.2015) (Ordenes 19.07.2017/20.08.2017)
                };
                //                                                           // Fechas (Usos 26.07.2017)  estas fechas son de Desarrollo-

                IRfcFunction usosFm = sap.ExecuteRFC(mandante, "ZDM_GET_DATA_NASH", parameters);

                nashUsosList = usosFm.GetTable("USOS_NASH_LIST");

                if (nashUsosList.RowCount > 0)
                {
                    console.WriteLine("Calculando Usos...");
                    xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    xlWorksheet = (Worksheet)xlWorkbook.Sheets[1];

                    //El primer valor se pone normalmente en la fila 1
                    console.WriteLine("... Creando CSV");
                    #region primer valor
                    xlWorksheet.Cells[1, 1] = "USO";
                    xlWorksheet.Cells[1, 2] = "S";
                    xlWorksheet.Cells[1, 3] = nashUsosList[0].GetValue("UBICACION").ToString().Substring(0, 2) + "02";
                    DateTime dateTime = new DateTime();
                    dateTime = Convert.ToDateTime(nashUsosList[0].GetValue("FECHADOCUMENTO").ToString());
                    xlWorksheet.Cells[1, 4] = "'" + dateTime.ToString("yyyy-MM-dd");
                    xlWorksheet.Cells[1, 5] = nashUsosList[0].GetValue("PRODUCTO").ToString();
                    xlWorksheet.Cells[1, 6] = "UN";

                    float ioQuantity = float.Parse(nashUsosList[0].GetValue("USOSCANTIDAD").ToString());
                    float iUsosquantity = (float)Math.Round(ioQuantity * 100f) / 100f;
                    float ioCost = float.Parse(nashUsosList[0].GetValue("COSTO").ToString());
                    float iUsosCost = (float)Math.Round(ioCost * 100f) / 100f;

                    string iUsosQuantities = iUsosquantity.ToString().Replace(",", ".");
                    string iUsosCosts = iUsosCost.ToString().Replace(",", ".");

                    xlWorksheet.Cells[1, 7] = iUsosQuantities;
                    xlWorksheet.Cells[1, 8] = iUsosCosts;
                    xlWorksheet.Cells[1, 9] = 1; //picks
                    #endregion

                    int cont = 2; //inicializamos un contador que nos va a ir diciendo en que fila pones el valor (inicia en fila 2 del excel)
                    int picks = 1;
                    //ahora inicia en el campo 01 de la tabla USOS_NASH_LIST
                    for (int i = 1; i < nashUsosList.RowCount; i++)
                    {
                        string currentProduct = nashUsosList[i].GetValue("PRODUCTO").ToString();

                        string orderdate = nashUsosList[i].GetValue("FECHADOCUMENTO").ToString();
                        DateTime dt = new DateTime();
                        dt = Convert.ToDateTime(orderdate);
                        string currentNashDate = dt.ToString("yyyy-MM-dd");

                        string lastProduct = nashUsosList[i - 1].GetValue("PRODUCTO").ToString(); //Se toma el anterior en la lista por eso se le resta 1

                        string lastDate = nashUsosList[i - 1].GetValue("FECHADOCUMENTO").ToString();
                        var lastNashDate = Convert.ToDateTime(lastDate).ToString("yyyy-MM-dd");

                        if (currentProduct == lastProduct && currentNashDate == lastNashDate)
                        {
                            picks++;
                            xlWorksheet.Cells[cont - 1, 9] = picks;

                            try
                            {
                                string currentQuantityStr = nashUsosList[i].GetValue("USOSCANTIDAD").ToString(); //.Replace(".000","")
                                float currentQuantity = float.Parse(currentQuantityStr);
                                string currentCostStr = nashUsosList[i].GetValue("COSTO").ToString();
                                float currentCost = float.Parse(currentCostStr);

                                float lastQuantity = float.Parse(xlWorksheet.Cells[cont - 1, 7].value.ToString());
                                float lastCost = float.Parse(xlWorksheet.Cells[cont - 1, 8].value.ToString());

                                float cat = currentQuantity + lastQuantity;
                                float cot = currentCost + lastCost;

                                float totalQuantity = (float)Math.Round(cat * 100f) / 100f;
                                float totalCost = (float)Math.Round(cot * 100f) / 100f;

                                string totalQuantities = totalQuantity.ToString().Replace(",", ".");
                                string totalCosts = totalCost.ToString().Replace(",", ".");

                                xlWorksheet.Cells[cont - 1, 7] = totalQuantities.ToString();
                                xlWorksheet.Cells[cont - 1, 8] = totalCosts.ToString();
                            }
                            catch (Exception) { }

                        }
                        else
                        {
                            picks = 1;
                            xlWorksheet.Cells[cont, 1] = "USO";
                            xlWorksheet.Cells[cont, 2] = "S";

                            string location = nashUsosList[i].GetValue("UBICACION").ToString();
                            location = location.Substring(0, 2) + "02";
                            xlWorksheet.Cells[cont, 3] = location;

                            xlWorksheet.Cells[cont, 4] = "'" + currentNashDate;
                            xlWorksheet.Cells[cont, 5] = currentProduct;
                            string unidad = nashUsosList[i].GetValue("UOM").ToString();
                            if (unidad == "ST")
                                unidad = "UN";

                            xlWorksheet.Cells[cont, 6] = unidad;
                            float oc = float.Parse(nashUsosList[i].GetValue("USOSCANTIDAD").ToString());
                            float usosQuantity = (float)Math.Round(oc * 100f) / 100f;
                            float oco = float.Parse(nashUsosList[i].GetValue("COSTO").ToString());
                            float usosCost = (float)Math.Round(oco * 100f) / 100f;
                            string usosQuantities = usosQuantity.ToString().Replace(",", ".");
                            string usosCosts = usosCost.ToString().Replace(",", ".");
                            xlWorksheet.Cells[cont, 7] = usosQuantities.ToString();
                            xlWorksheet.Cells[cont, 8] = usosCosts.ToString();
                            xlWorksheet.Cells[cont, 9] = 1; //picks
                            cont++;
                        }
                    }

                    xlApp.DisplayAlerts = false;

                    string mes = "";
                    string dia = "";
                    mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                    dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

                    string nameFile1 = "TRAN" + DateTime.Now.Year + mes + dia + "0002.csv";
                    string nameFile2 = "TRAN" + DateTime.Now.Year + mes + dia + "0002.xml";

                    xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + nameFile1, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                    xlWorkbook.Close();
                    xlApp.Quit();

                    lines3 = File.ReadAllLines(root.FilesDownloadPath + "\\" + nameFile1);

                    console.WriteLine("... Creando XML");

                    Program program = new Program();
                    XNamespace texto = "urn:transaccion-schema";
                    XElement xml = new XElement(texto + "ITransaccion",
                    from str in lines3
                    let columns = str.Split(',')
                    select new XElement(texto + "ITransaccionRow",
new XElement(texto + "TipoTrans", columns[0]),
new XElement(texto + "Accion", columns[1]),
new XElement(texto + "Ubicacion", columns[2]),
new XElement(texto + "Fecha", columns[3]),
new XElement(texto + "Gtin", columns[4]),
new XElement(texto + "UOM", columns[5]),
new XElement(texto + "Cantidad", columns[6]),
new XElement(texto + "Costo", columns[7]),
new XElement(texto + "Picks", columns[8])
));

                    console.WriteLine("... Guardando XML en Server");
                    FileDelete(folderPath + nameFile2);
                    xml.Save(folderPath + nameFile2); //OJO ruta servidor NASH
                    log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", nameFile2, "con exito");
                    respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + nameFile2;

                }
                else
                {
                    response += "Aviso Usos: no hay data" + "<br>";
                    console.WriteLine("Aviso Usos: no hay data");
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(" Error process " + ex.Message);
                response += "Error Usos: " + ex.Message + "<br>";
                validateLines = false;
            }

            #endregion

            #region----------------------------------------( Ventas )--------------------------------------------------------------------------------
            try
            {
                console.WriteLine("Corriendo WS de SAP para Ventas...");


                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["VENTA"] = "X",
                    ["FECHA_INICIAL"] = startDate,
                    ["FECHA_FINAL"] = endDate                        // fecha para correr el robot (recibos 16.08.2017/xxxx) (Producto 15.03.2017/16.11.2015) (Ordenes 19.07.2017/20.08.2017)
                };
                //                                                             // Fechas (Usos 26.07.2017)  estas fechas son de Desarrollo
                //                                                             //--------correr WS------------

                IRfcFunction salesFm = sap.ExecuteRFC(mandante, "ZDM_GET_DATA_NASH", parameters);


                nashSalesList = salesFm.GetTable("VENTAS_NASH_LIST");

                if (nashSalesList.RowCount > 0)
                {
                    console.WriteLine("... Creando CSV");
                    xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    xlWorksheet = (Worksheet)xlWorkbook.Sheets[1];
                    for (int i = 0; i < nashSalesList.RowCount; i++)
                    {
                        xlWorksheet.Cells[i + 1, 1] = "VEN";
                        xlWorksheet.Cells[i + 1, 2] = "S";

                        string location = nashSalesList[i].GetValue("UBICACION").ToString();
                        location = location.Substring(0, 2) + "02";
                        xlWorksheet.Cells[i + 1, 3] = location;

                        string orderDate = nashSalesList[i].GetValue("FECHADOCUMENTO").ToString();
                        DateTime dt = new DateTime();
                        dt = Convert.ToDateTime(orderDate);
                        string nashDate = dt.ToString("yyyy-MM-dd");
                        xlWorksheet.Cells[i + 1, 4] = "'" + nashDate;
                        xlWorksheet.Cells[i + 1, 5] = nashSalesList[i].GetValue("PRODUCTO").ToString();
                        xlWorksheet.Cells[i + 1, 6] = "UN";

                        float oc = float.Parse(nashSalesList[i].GetValue("VENTACANTIDAD").ToString());
                        float salesQuantity = (float)Math.Round(oc * 100f) / 100f;
                        float oco = float.Parse(nashSalesList[i].GetValue("COSTO").ToString());
                        float salesCost = (float)Math.Round(oco * 100f) / 100f;
                        string salesQuantities = salesQuantity.ToString().Replace(",", ".");
                        string salesCosts = salesCost.ToString().Replace(",", ".");
                        xlWorksheet.Cells[i + 1, 7] = salesQuantities.ToString();
                        xlWorksheet.Cells[i + 1, 8] = salesCosts.ToString();

                        xlWorksheet.Cells[i + 1, 9] = nashSalesList[i].GetValue("INGRESO").ToString();
                        xlWorksheet.Cells[i + 1, 10] = 1;
                    }

                    xlApp.DisplayAlerts = false;
                    string mes = "";
                    string dia = "";
                    mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                    dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

                    string fileName1 = "TRAN" + DateTime.Now.Year + mes + dia + "0004.csv";
                    string fileName2 = "TRAN" + DateTime.Now.Year + mes + dia + "0004.xml";

                    xlWorkbook.SaveAs(root.FilesDownloadPath + "\\" + fileName1, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                    xlWorkbook.Close();
                    xlApp.Quit();

                    lines6 = File.ReadAllLines(root.FilesDownloadPath + "\\" + fileName1);

                    console.WriteLine("... Creando XML");

                    Program program = new Program();
                    XNamespace texto = "urn:transaccion-schema";
                    XElement xml = new XElement(texto + "ITransaccion",
                    from str in lines6
                    let columns = str.Split(',')
                    select new XElement(texto + "ITransaccionRow",
new XElement(texto + "TipoTrans", columns[0]),
new XElement(texto + "Accion", columns[1]),
new XElement(texto + "Ubicacion", columns[2]),
new XElement(texto + "Fecha", columns[3]),
new XElement(texto + "Gtin", columns[4]),
new XElement(texto + "UOM", columns[5]),
new XElement(texto + "Cantidad", columns[6]),
new XElement(texto + "Costo", columns[7]),
new XElement(texto + "Ingreso", columns[8]),
new XElement(texto + "Picks", columns[9])
));

                    console.WriteLine("... Guardando XML en Server");
                    FileDelete(folderPath + fileName2);
                    xml.Save(folderPath + fileName2); //OJO ruta servidor NASH  
                    log.LogDeCambios("Creacion", root.BDProcess, "Departamento de Inv", "Cargar XML:", fileName2, "con exito");
                    respFinal = respFinal + "\\n" + "Se cargó el XML con éxito:" + fileName2;

                }
                else
                {
                    response += "Aviso Ventas: no hay data" + "<br>";
                    console.WriteLine("Aviso Ventas: no hay data");
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(" Error process " + ex.Message);
                response += "Error Usos: " + ex.Message + "<br>";
                validateLines = false;
            }
            #endregion

            #region Finalizando
            console.WriteLine("Respondiendo solicitud");

            xlApp.Quit();
            proc.KillProcess("EXCEL", true);

            if (validateLines == false)
            {
                //enviar email de repuesta de error
                string[] cc = { "jearaya@gbm.net" };
                mail.SendHTMLMail(response, new string[] { "internalcustomersrvs@gbm.net" }, "Error: Carga data NASH - " + DateTime.Now, cc);
            }
            else
            {
                string[] cc = { "jearaya@gbm.net" };
                //enviar email de repuesta de éxito
                mail.SendHTMLMail(response, new string[] { "internalcustomersrvs@gbm.net" }, "Carga data NASH - " + DateTime.Now, cc);
            }
            #endregion
        }
        public string GetDate(string date)
        {
            try
            {
                DateTime DT = Convert.ToDateTime(date);
                return DT.ToString("yyyy-MM-dd");
            }
            catch (Exception)
            {
                return "";
            }
        }
        /// <summary>
        /// Método que tiene un proceso de buscar los datos de proveedor por GM.
        /// </summary>
        public DataTable Getvendor(string mg)
        {
            DataTable sqlTable = new DataTable();
            try
            {
                //sqlTable = new CRUD().Select("Databot", "SELECT * FROM mg_vendor WHERE material_group = '" + mg + "'", "nash");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
            }

            return sqlTable;
        }
        public void FileDelete(string fileName)
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }
    }
}
