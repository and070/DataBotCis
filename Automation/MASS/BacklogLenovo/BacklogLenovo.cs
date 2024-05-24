using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Data;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Data.Database;

namespace DataBotV5.Automation.MASS.BacklogLenovo

{
    /// <summary>
    /// Clase MASS Automation encargada de solicitudes Backlog Leonovo.
    /// </summary>
    class BacklogLenovo
    {
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        SapVariants sap = new SapVariants();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        SharePoint sharepoint = new SharePoint();
        BsSQL bsql = new BsSQL();
        Log log = new Log();
        MsExcel MsExcel = new MsExcel();
        Stats estadisticas = new Stats();
        object[] columnas_duplicate;
        object[] columnas_duplicate2;
        string respFinal = "";
        string mandante = "ERP";
        CRUD crud = new CRUD();
        string mand = "QAS";


        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
            if (!sap.CheckLogin(mandante))
            {
                //leer correo y descargar archivo
                if (mail.GetAttachmentEmail("Solicitudes Backlog Lenovo", "Procesados", "Procesados Backlog Lenovo"))
                {
                    console.WriteLine("Procesando...");

                    ProcessBLLenovoV2(root.FilesDownloadPath + "\\" + root.ExcelFile);

                    root.requestDetails = respFinal;
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }

        public void ProcessBLLenovoV2(string route)
        {
            #region Variables Privadas
            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;
            var valor = "";
            string respuesta = "";
            long contador; long contador_while;
            string sheetname = "";
            bool noshipdate = false;
            string name = "Órdenes en Stock In Transit 1.xlsx";

            DateTime fecha; DateTime fecha_corte; DateTime fecha_anterior;
            long a; int dia = 0; int mes = 0; string today_sap = ""; string fecha_sap = ""; string fecha_file = "";
            string order = ""; string order_status = ""; string mensaje_sap = ""; string validacion = "";

            //columnas necesarias para cargar en SAP
            string[] readColumns = {
                "Sales Order Number",
                "Actual Delivery Date"
            };
            //contador para verificar el excel
            int contTrue = 0;
            //PLantilla en html para el envío de email
            string htmlEmail = Properties.Resources.emailtemplate1;
            #endregion

            console.WriteLine("Abrir Excel y modificando");

            #region abrir excel

            DataSet xlWorkBook2 = MsExcel.GetExcelBook(route);
            //DataTable xlWorkSheetColumns = xlWorkBook2.Tables["Order List"];

            DataTable xlWorkSheetColumns = new DataTable();

            xlWorkSheetColumns.Columns.Add("Actual Ship Date", System.Type.GetType("System.DateTime"));
            xlWorkSheetColumns.Columns.Add("Actual Delivery Date", System.Type.GetType("System.DateTime"));
            xlWorkSheetColumns.Columns.Add("Estimated Delivery Date", System.Type.GetType("System.DateTime"));
            xlWorkSheetColumns.Columns.Add("Firm Ship Date", System.Type.GetType("System.DateTime"));
            xlWorkSheetColumns.Columns.Add("Estimated Ship Date", System.Type.GetType("System.DateTime"));
            xlWorkSheetColumns.Columns.Add("Order Entry Date", System.Type.GetType("System.DateTime"));
            xlWorkSheetColumns.Columns.Add("Order Receipt Date", System.Type.GetType("System.DateTime"));

            xlWorkSheetColumns.Columns["Actual Ship Date"].DataType = System.Type.GetType("System.DateTime");
            xlWorkSheetColumns.Columns["Actual Delivery Date"].DataType = System.Type.GetType("System.DateTime");
            xlWorkSheetColumns.Columns["Estimated Delivery Date"].DataType = System.Type.GetType("System.DateTime");
            xlWorkSheetColumns.Columns["Firm Ship Date"].DataType = System.Type.GetType("System.DateTime");
            xlWorkSheetColumns.Columns["Estimated Ship Date"].DataType = System.Type.GetType("System.DateTime");
            xlWorkSheetColumns.Columns["Order Entry Date"].DataType = System.Type.GetType("System.DateTime");
            xlWorkSheetColumns.Columns["Order Receipt Date"].DataType = System.Type.GetType("System.DateTime");

            xlWorkSheetColumns = xlWorkBook2.Tables["Order List"];

            //columnas del excel
            DataColumnCollection columns = xlWorkSheetColumns.Columns;
            #endregion

            #region definicion de fechas

            fecha_corte = DateTime.Now.AddMonths(-12);
            fecha_anterior = DateTime.Now.AddMonths(-12);
            //fecha_corte = DateTime.Now.AddDays(1).AddMonths(-12);
            //fecha_anterior = DateTime.Now.AddDays(1).AddMonths(-12);

            dia = fecha_anterior.Day;
            mes = fecha_anterior.Month;
            fecha_sap = dia + "." + mes + "." + fecha_anterior.Year.ToString();

            dia = DateTime.Now.Day;
            mes = DateTime.Now.Month;
            today_sap = dia + "." + mes + "." + DateTime.Now.Year.ToString();
            fecha_file = dia + "_" + mes + "_" + DateTime.Now.Year.ToString();
            #endregion

            #region validacion
            foreach (string columnName in readColumns)
            {
                //verifica si la columna esta en el excel
                if (columns.Contains(columnName))
                {
                    contTrue++;
                }
            }
            //si es diferente a 4 significa que no encontro una de las columnas necesarias para cargar la reconocimiento
            if (contTrue != readColumns.Length)
            {
                respuesta = "No es plantilla de Lenovo";
                htmlEmail = htmlEmail.Replace("{subject}", root.Subject).Replace("{cuerpo}", respuesta).Replace("{contenido}", "");
                mail.SendHTMLMail(htmlEmail, new string[] { root.BDUserCreatedBy }, "Error: " + root.Subject, root.CopyCC, new string[] { root.FilesDownloadPath + "\\" + root.ExcelFile });
                return;
            }
            #endregion

            #region por cada hoja del excel
            //for para hacer el mismo paso por cada hoja que tenga el excel
            int i = 1;
            foreach (DataTable dataSheet in xlWorkBook2.Tables)
            {
                DataTable workSheet = new DataTable();

                #region poner las columnas de fechas como custom
                workSheet.Columns.Add("Actual Ship Date", System.Type.GetType("System.DateTime"));
                workSheet.Columns.Add("Actual Delivery Date", System.Type.GetType("System.DateTime"));
                workSheet.Columns.Add("Estimated Delivery Date", System.Type.GetType("System.DateTime"));
                workSheet.Columns.Add("Firm Ship Date", System.Type.GetType("System.DateTime"));
                workSheet.Columns.Add("Estimated Ship Date", System.Type.GetType("System.DateTime"));
                workSheet.Columns.Add("Order Entry Date", System.Type.GetType("System.DateTime"));
                workSheet.Columns.Add("Order Receipt Date", System.Type.GetType("System.DateTime"));

                workSheet.Columns["Actual Ship Date"].DataType = System.Type.GetType("System.DateTime");
                workSheet.Columns["Actual Delivery Date"].DataType = System.Type.GetType("System.DateTime");
                workSheet.Columns["Estimated Delivery Date"].DataType = System.Type.GetType("System.DateTime");
                workSheet.Columns["Firm Ship Date"].DataType = System.Type.GetType("System.DateTime");
                workSheet.Columns["Estimated Ship Date"].DataType = System.Type.GetType("System.DateTime");
                workSheet.Columns["Order Entry Date"].DataType = System.Type.GetType("System.DateTime");
                workSheet.Columns["Order Receipt Date"].DataType = System.Type.GetType("System.DateTime");
                #endregion

                workSheet = dataSheet.Copy();

                columns = xlWorkSheetColumns.Columns;
                int filas = workSheet.Rows.Count;
                string nombreHoja = workSheet.TableName;
                workSheet.Columns.Remove("Direct or Indirect");

                #region eliminar duplicados

                console.WriteLine(" Eliminar Duplicados");
                DataTable xLworkSheet = workSheet.AsEnumerable()
               .OrderBy(x => x.Field<string>("Sales Order Number"))
               .GroupBy(x => new
               {
                   oeDate = x.Field<DateTime>("Order Entry Date"),
                   sOrder = x.Field<string>("Sales Order Number"),
                   cPoNum = x.Field<string>("Customer Purchase Order Number"),
                   proId = x.Field<string>("Product ID"),
                   oQuan = x.Field<double>("Order Quantity"),
                   sQuan = x.Field<double>("Shipped Quantity"),
                   iNum = x.Field<string>("Invoice Number"),
                   fSDate = x.Field<DateTime?>("Firm Ship Date"),
                   cName = x.Field<string>("Carrier Name"),
                   cTrackNum = x.Field<string>("Carrier Tracking Number"),
               })
               .Select(x => x.First())
               .CopyToDataTable();

                #endregion

                #region dejar los ultimos 12 meses
                console.WriteLine(" Dejar los ultimos 12 meses");
                DataRow[] rowsToRemove = xLworkSheet.AsEnumerable()
                 .Where(row => DateTime.Parse(row["Order Entry Date"].ToString()) < fecha_corte)
                    .ToArray();// Required to prevent "Collection was modified" exception in foreach below

                foreach (DataRow row in rowsToRemove)
                {
                    xLworkSheet.Rows.Remove(row);
                }
                #endregion

                #region Cambio de fechas y cambiar productos
                console.WriteLine(" Cambio de Fechas");

                foreach (DataRow dr in xLworkSheet.Rows)
                {
                    string productId = dr["Product ID"].ToString();
                    if (productId.Contains("-"))
                    {
                        dr["Product ID"] = dr["Product ID"].ToString().Replace("-", "");
                    }

                    //copiar el valor de la columna W en la P si es diferente a blanco
                    if (!string.IsNullOrWhiteSpace(dr["Actual Ship Date"].ToString()))
                    {
                        dr["Firm Ship Date"] = dr["Actual Ship Date"];
                    }

                    //si la columna F es Non-shippable poner LSEF en la columna U
                    if (dr["Line Item Status"].ToString() == "Non-shippable")
                    {
                        dr["Carrier Tracking Number"] = "L.S.E.F";
                    }

                    //si la columna O es diferente a vacio y la columna P es igual a vacio
                    if (!string.IsNullOrWhiteSpace(dr["Estimated Ship Date"].ToString()) && string.IsNullOrWhiteSpace(dr["Firm Ship Date"].ToString()))
                    {
                        dr["Firm Ship Date"] = dr["Estimated Ship Date"];
                    }

                    //agregar 3 días firm ship date
                    DateTime firm_ship_date = DateTime.MinValue;

                    try
                    {
                        firm_ship_date = DateTime.Parse(dr["Firm Ship Date"].ToString());
                        firm_ship_date = firm_ship_date.AddDays(4);
                        dr["Firm Ship Date"] = firm_ship_date;
                    }
                    catch (Exception ex)
                    {
                        //Console.WriteLine(ex.Message);
                    }

                    //agregar 15 dias a fechas si el product id empieza con 6

                    if (productId.Substring(0, 1).ToString() == "6")
                    {
                        DateTime estimatedDeliveryDate = DateTime.MinValue;
                        try
                        {
                            estimatedDeliveryDate = DateTime.Parse(dr["Estimated Delivery Date"].ToString());

                            estimatedDeliveryDate = estimatedDeliveryDate.AddDays(15);

                            dr["Estimated Delivery Date"] = estimatedDeliveryDate;
                        }
                        catch (Exception)
                        {

                        }

                        DateTime actualDeliveryDate = DateTime.MinValue;
                        try
                        {
                            actualDeliveryDate = DateTime.Parse(dr["Actual Delivery Date"].ToString());

                            actualDeliveryDate = actualDeliveryDate.AddDays(15);

                            dr["Actual Delivery Date"] = actualDeliveryDate;
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
                #endregion

                //se guarda el archivo y se envía a Business System, luego se hacen más validaciones pero para subir a SAP 
                xLworkSheet.AcceptChanges();

                string fileBeforeSapPath = root.FilesDownloadPath + "\\" + "BL" + " - " + fecha_file + ".xlsx";
                MsExcel.CreateExcel(xLworkSheet, nombreHoja, fileBeforeSapPath, true);

                #region Subir archivo a Sharepoint para portal en S&S
                try
                {
                    sharepoint.UploadFileToSharePointV2("https://gbmcorp.sharepoint.com/sites/PurchasingLenovo/", "Documentos/Backlog Lenovo/", fileBeforeSapPath);

                }
                catch (Exception ex)
                {
                    mail.SendHTMLMail($"Error{ex.Message}", new string[] { "dmeza@gbm.net" }, "Error al subir archivo a sharepoint", null, null);
                }
                #endregion

                //Convierte las fechas al formato gringo sin tocar la hoja original xLworkSheet
                convertDates(fileBeforeSapPath);

                #region Validaciones para subir a SAP
                foreach (DataRow dr in xLworkSheet.Rows)
                {
                    //prefijo PO
                    string po = dr["Customer Purchase Order Number"].ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(po))
                    {
                        if (!int.TryParse(po, out _))
                        {
                            try
                            {
                                po = string.Concat(po.Where(c => Char.IsDigit(c)));
                                //po = po.Substring(2, po.Length - 2);
                                dr["Customer Purchase Order Number"] = po;
                            }
                            catch (Exception)
                            {
                            }

                        }
                    }

                    //cambio de fecha no tracking number
                    string val2 = dr["Firm Ship Date"].ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(val2))
                    {
                        DateTime ship_date2 = DateTime.Parse(dr["Firm Ship Date"].ToString());

                        Int64 addedDays = 1;
                        DateTime endDate = DateTime.Today.AddDays(addedDays);
                        if (ship_date2 < DateTime.Today)
                        { noshipdate = true; }
                        // sumarle un dia a la fecha de hoy
                        if (dr["Shipped Quantity"].ToString() == "0" && ship_date2 <= DateTime.Today && string.IsNullOrWhiteSpace(dr["Carrier Tracking Number"].ToString()))
                        {
                            dr["Firm Ship Date"] = endDate;
                        }
                    }

                    //Elimina los 0 con partially shipped y delivery
                    if (dr["Line Item Status"].ToString() == "Partially Shipped" || dr["Line Item Status"].ToString() == "Partially Delivered")
                    {
                        string valora = dr["Shipped Quantity"].ToString();
                        if (valora == "0")
                        {
                            dr.Delete();
                        }
                    }

                }
                #endregion

                string fileSapPath = root.FilesDownloadPath + "\\" + "BLSAP" + i + ".xlsx";
                xLworkSheet.AcceptChanges();
                MsExcel.CreateExcel(xLworkSheet, nombreHoja, fileSapPath, true);

                //despues de unos cambios vuelve a convertir las fechas a formato gringo sin tocar la hoja xLworkSheet
                convertDates(fileSapPath);
                i++;


                #region cargar en SAP
                console.WriteLine("Cargando a SAP");
                //se espera hasta que el mandante este desbloqueado
                while (sap.CheckLogin(mandante))
                {
                    System.Threading.Thread.Sleep(1000);
                }
                //bloquea mandante
                sap.BlockUser(mandante, 1);
                sap.LogSAP(mandante);

                try
                {
                    SapVariants.frame.Iconify();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "Y_MM_BACKLOG_PLANTILLA_LENOVO";
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtP_FNAME")).Text = fileSapPath;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSO_DATE-LOW")).Text = fecha_sap;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSO_DATE-HIGH")).Text = today_sap;
                    try
                    {
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                        mensaje_sap = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[1]/usr/txtMESSTXT1")).Text;
                        SAPFEWSELib.GuiFrameWindow frame1 = (SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[1]");
                        frame1.Iconify();
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                    }
                    catch (Exception)
                    { }
                    SapVariants.frame.Iconify();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/n";
                    SapVariants.frame.SendVKey(0);
                    sap.KillSAP();
                }
                catch (Exception ex)
                {
                    mail.SendHTMLMail("Error al subir el archivo a SAP", new string[] {"appmanagement@gbm.net"}, root.Subject);
                }

                sap.BlockUser(mandante, 0);
                #endregion cargar en SAP

                if (mensaje_sap == "")
                { mensaje_sap = "Se ha cargado con exito"; }

                console.WriteLine(mensaje_sap);

                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear backLog Lenovo", mensaje_sap, fileSapPath);
                respFinal = respFinal + "\\n" + "Crear backLog Lenovo: " + mensaje_sap + " " + fileSapPath;

                string rutaEmailFinal = root.FilesDownloadPath + "\\" + "BL SAP" + " - " + fecha_file + ".xlsx";
                xLworkSheet.AcceptChanges();
                MsExcel.CreateExcel(xLworkSheet, nombreHoja, rutaEmailFinal, true);

                //por ultimo convierte las fechas del ultimo archivo que sería el que se envia por email
                convertDates(rutaEmailFinal);

                console.WriteLine("Enviando Archivo a BS");

                //enviar email con el archivo root.Google_Download + "\\" + "BL" + " - " + fecha_file + ".xlsx" a la gente y copia que esta en la base de datos
                string[] cc = bsql.EmailAddress(2);
                //string[] cc = { root.f_copy1, root.f_copy2, root.f_copy3, root.f_copy4, root.f_copy5, root.f_copy6 };

                string[] adjunto = { fileBeforeSapPath, rutaEmailFinal };
                mail.SendHTMLMail(mensaje_sap + " el BackLog en SAP", new string[] { root.f_sender }, root.Subject, cc, adjunto);
            }



            #endregion


        }
        private void convertDates(string route)
        {
            #region abrir excel
            Excel.Range xlRango;
            Excel.Range xlRangoDuplicate;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];

            #endregion
            string[] ranges =
            {
                "C:C" ,
                "D:D" ,
                "O:O" ,
                "P:P" ,
                "Q:Q" ,
                "R:R" ,
                "W:W"

            };



            Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
            Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;
            foreach (string range in ranges)
            {
                xlWorkSheet.Range[range].NumberFormat = "MM/DD/YYYY";
                xlWorkSheet.Range[range].Copy();
                xlWorkSheet.Range[range].PasteSpecial(paste, pasteop, false, false);
            }



            xlWorkBook.SaveAs(route);
            xlWorkBook.Close();

            xlApp.DisplayAlerts = false;
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);


        }

        private void FormatAndSetDate(DataRow row, string columnName)
        {
            if (row[columnName] is DateTime dateValue)
            {
                row[columnName] = dateValue.ToString("MM/dd/yyyy");
            }
            else if (row[columnName] is string stringValue && DateTime.TryParse(stringValue, out dateValue))
            {
                row[columnName] = dateValue.ToString("MM/dd/yyyy");
            }
            // Handle other cases or errors if needed
        }

        public void ProcessBLLenovo(string route)
        {
            #region Variables Privadas
            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;
            var valor = "";
            string respuesta = "";
            long contador; long contador_while;
            string sheetname = "";
            bool noshipdate = false;

            DateTime fecha; DateTime fecha_corte; DateTime fecha_anterior;
            long a; int dia = 0; int mes = 0; string today_sap = ""; string fecha_sap = ""; string fecha_file = "";
            string order = ""; string order_status = ""; string mensaje_sap = ""; string validacion = "";
            #endregion

            console.WriteLine("Abrir Excel y modificando");

            #region abrir excel
            Excel.Range xlRango;
            Excel.Range xlRangoDuplicate;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];

            #endregion

            #region definicion de fechas

            fecha_corte = DateTime.Now.AddMonths(-12);
            fecha_anterior = DateTime.Now.AddMonths(-12);

            dia = fecha_anterior.Day;
            mes = fecha_anterior.Month;
            fecha_sap = dia + "." + mes + "." + fecha_anterior.Year.ToString();

            dia = DateTime.Now.Day;
            mes = DateTime.Now.Month;
            today_sap = dia + "." + mes + "." + DateTime.Now.Year.ToString();
            fecha_file = dia + "_" + mes + "_" + DateTime.Now.Year.ToString();
            #endregion

            #region validacion
            validacion = xlWorkSheet.Cells[1, 24].text;
            if (validacion != "")
            { validacion = validacion.ToString().Trim(); }
            if (validacion != "Mode Of Transportation")
            {
                mensaje_devolucion = "No es plantilla de Lenovo";
                respuesta = "No es plantilla de Lenovo";
                mail.SendHTMLMail(respuesta, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
                //enviar email a usuario y datos maestros

            }

            else
            {

                //for para hacer el mismo paso por cada hoja que tenga el excel
                contador = xlWorkBook.Worksheets.Count;
                for (int i = 1; i <= contador; i++)
                {

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[i];
                    rows = xlWorkSheet.UsedRange.Rows.Count;
                    sheetname = xlWorkSheet.Name;
                    #region poner las columnas de fechas como custom
                    xlWorkSheet.Range["Y:Y"].Delete();
                    xlWorkSheet.Range["C:C"].NumberFormat = "MM/DD/YYYY";
                    xlWorkSheet.Range["D:D"].NumberFormat = "MM/DD/YYYY";
                    xlWorkSheet.Range["O:O"].NumberFormat = "MM/DD/YYYY";
                    xlWorkSheet.Range["P:P"].NumberFormat = "MM/DD/YYYY";
                    xlWorkSheet.Range["Q:Q"].NumberFormat = "MM/DD/YYYY";
                    xlWorkSheet.Range["R:R"].NumberFormat = "MM/DD/YYYY";
                    xlWorkSheet.Range["W:W"].NumberFormat = "MM/DD/YYYY";
                    #endregion

                    #region eliminar duplicados
                    console.WriteLine(" Eliminar Duplicados");
                    //for para agregar todas las columnas respectivas al array para eliminar duplicados

                    for (int x = 0; x <= 9; x++)
                    {
                        //ReDim Preserve columnas_duplicate(x);
                        Array.Resize(ref columnas_duplicate, x + 1);
                        switch (x)
                        {
                            case 0:
                                columnas_duplicate[x] = 4;
                                break;
                            case 1:
                                columnas_duplicate[x] = 7;
                                break;
                            case 2:
                                columnas_duplicate[x] = 8;
                                break;
                            case 3:
                                columnas_duplicate[x] = 9;
                                break;
                            case 4:
                                columnas_duplicate[x] = 10;
                                break;
                            case 5:
                                columnas_duplicate[x] = 13;
                                break;
                            case 6:
                                columnas_duplicate[x] = 14;
                                break;
                            case 7:
                                columnas_duplicate[x] = 16;
                                break;
                            case 8:
                                columnas_duplicate[x] = 20;
                                break;
                            case 9:
                                columnas_duplicate[x] = 21;
                                break;
                        }
                    }
                    Array.Resize(ref columnas_duplicate, 10);
                    //eliminar duplicados
                    xlWorkSheet.Range["A1:Y" + rows].RemoveDuplicates(columnas_duplicate, Excel.XlYesNoGuess.xlYes);
                    #endregion eliminar duplicados

                    #region dejar los ultimos 6 meses
                    console.WriteLine(" Dejar los ultimos 6 meses");
                    rows = xlWorkSheet.UsedRange.Rows.Count; //nuevos rows sin duplicados


                    a = 2;
                    contador_while = 0;
                    order = xlWorkSheet.Cells[a, 7].text;

                    while (order != "")
                    {
                        fecha = xlWorkSheet.Cells[a, 4].value;
                        if (fecha < fecha_corte)
                        {
                            xlRango = xlWorkSheet.Rows[a + ":" + a];
                            xlRango.Select();
                            xlRango.Delete();
                            xlRango = null;
                            a = a - 1;
                            rows = xlWorkSheet.UsedRange.Rows.Count;
                        }
                        a = a + 1;
                        order = xlWorkSheet.Cells[a, 1].text;
                    }
                    #endregion dejar los ultimos 6 meses

                    #region validaciones extras
                    rows = xlWorkSheet.UsedRange.Rows.Count;

                    Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
                    Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;

                    xlWorkSheet.Range["W:W"].Copy();
                    xlWorkSheet.Range["P:P"].PasteSpecial(paste, pasteop, false, false);

                    xlWorkSheet.Range["O:O"].Copy();
                    xlWorkSheet.Range["P:P"].PasteSpecial(paste, pasteop, false, false);
                    xlWorkSheet.Range["P:P"].NumberFormat = "MM/DD/YYYY";


                    console.WriteLine(" Cambio de Fechas");
                    for (int e = 2; e <= rows; e++)
                    {
                        //'copiar el valor de la columna W en la P
                        if (xlWorkSheet.Cells[e, 23].text != "")
                        {
                            xlWorkSheet.Range["P" + e].NumberFormat = "MM/DD/YYYY";
                            string ship_date = xlWorkSheet.Cells[e, 23].text;
                            xlWorkSheet.Cells[e, 16].value = xlWorkSheet.Cells[e, 23].text;
                        }

                        //si la columna F es Non-shippable poner LSEF en la columna U
                        order_status = xlWorkSheet.Cells[e, 6].text;
                        if (order_status == "Non-shippable")
                        {
                            xlWorkSheet.Cells[e, 21].value = "L.S.E.F";
                        }

                        //'columna O y P
                        if (xlWorkSheet.Cells[e, 15].text != "" && xlWorkSheet.Cells[e, 16].text == "")
                        {
                            xlWorkSheet.Range["P" + e].NumberFormat = "MM/DD/YYYY";
                            string estimate_ship_date = xlWorkSheet.Cells[e, 15].text;
                            xlWorkSheet.Cells[e, 16].value = estimate_ship_date;
                        }

                        //agregar 3 días firm ship date
                        try
                        {
                            DateTime firm_ship_date = DateTime.MinValue;

                            try
                            {
                                firm_ship_date = xlWorkSheet.Cells[e, 16].value;
                                firm_ship_date = firm_ship_date.AddDays(4);
                                xlWorkSheet.Range["P" + e].NumberFormat = "MM/DD/YYYY";
                                xlWorkSheet.Cells[e, 16].value = firm_ship_date;
                            }
                            catch (Exception ex)
                            {
                                //Console.WriteLine(ex.Message);
                            }

                        }
                        catch (Exception)
                        { }


                    } //for por cada fila del excel
                    #endregion

                    xlWorkBook.SaveAs(root.FilesDownloadPath + "\\" + "BL" + " - " + fecha_file + ".xlsx");


                    //###################################
                    // Robot que alimenta el portal de S&S de registro de fechas de ordenes Lenovo
                    // sumarle 3 días a la columna Q y R sin embargo puede ser que mejor lo hagamos en S&S ya que aqui nos cagamos en el otro
                    //
                    //
                    //###################################



                    #region validaciones extras despues de
                    rows = xlWorkSheet.UsedRange.Rows.Count;
                    console.WriteLine(" Validaciones extras");
                    for (int e = 2; e <= rows; e++)
                    {
                        //prefijo PO
                        string po = xlWorkSheet.Cells[e, 8].text.Trim();
                        if (po != "")
                        {
                            if (!int.TryParse(po, out _))
                            {
                                try
                                {
                                    po = string.Concat(po.Where(c => Char.IsDigit(c)));
                                    //po = po.Substring(2, po.Length - 2);
                                    xlWorkSheet.Cells[e, 8].value = po;
                                }
                                catch (Exception)
                                {
                                }

                            }
                        }

                        //cambio de fecha no tracking number
                        string val2 = xlWorkSheet.Cells[e, 16].text.Trim();
                        if (val2 != "")
                        {
                            DateTime ship_date2 = /*DateTime.Parse(*/ xlWorkSheet.Cells[e, 16].value;

                            Int64 addedDays = 1;
                            DateTime endDate = DateTime.Today.AddDays(addedDays);
                            if (ship_date2 < DateTime.Today)
                            { noshipdate = true; }
                            // sumarle un dia a la fecha de hoy
                            if (xlWorkSheet.Cells[e, 13].text == "0" && ship_date2 <= DateTime.Today && xlWorkSheet.Cells[e, 21].text == "")
                            {
                                xlWorkSheet.Range["P" + e].NumberFormat = "MM/DD/YYYY";
                                xlWorkSheet.Cells[e, 16].value = endDate;
                            }
                        }

                    } //for por cada fila del excel
                    #endregion

                    #region eliminar ceros
                    console.WriteLine(" Eliminar Ceros");
                    rows = xlWorkSheet.UsedRange.Rows.Count; //nuevos rows sin duplicados
                    a = 2;
                    contador_while = 0;
                    order = xlWorkSheet.Cells[a, 7].text;

                    while (order != "")
                    {
                        //Elimina los 0 con partially shipped y delivery

                        if (xlWorkSheet.Cells[a, 6].text == "Partially Shipped" || xlWorkSheet.Cells[a, 6].text == "Partially Delivered")
                        {
                            string valora = xlWorkSheet.Cells[a, 13].text;
                            if (valora == "0")
                            {
                                xlRango = xlWorkSheet.Rows[a + ":" + a];
                                xlRango.Select();
                                xlRango.Delete();
                                xlRango = null;
                                a = a - 1;
                                rows = xlWorkSheet.UsedRange.Rows.Count;
                            }
                        }

                        a = a + 1;
                        order = xlWorkSheet.Cells[a, 1].text;
                    }
                    #endregion

                    #region Especifica el archivo con su direccion
                    string File_sap = root.FilesDownloadPath + "\\" + "BLSAP" + i + ".xlsx";
                    if (File.Exists(File_sap))
                    {
                        File.Delete(File_sap);
                    }
                    xlWorkSheet.Copy(Type.Missing, Type.Missing);
                    xlApp.Workbooks[2].SaveAs(File_sap);
                    xlApp.Workbooks[2].Close();
                    #endregion


                    #region cargar en SAP
                    console.WriteLine("Cargando a SAP");

                    //se espera hasta que el mandante este desbloqueado
                    while (sap.CheckLogin(mandante))
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    //bloquea mandante
                    sap.BlockUser(mandante, 1);
                    sap.LogSAP(mandante.ToString());

                    try
                    {
                        SapVariants.frame.Iconify();
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
                        SapVariants.frame.SendVKey(0);
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "Y_MM_BACKLOG_PLANTILLA_LENOVO";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtP_FNAME")).Text = File_sap;
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSO_DATE-LOW")).Text = fecha_sap;
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtSO_DATE-HIGH")).Text = today_sap;
                        try
                        {
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                            mensaje_sap = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[1]/usr/txtMESSTXT1")).Text;
                            SAPFEWSELib.GuiFrameWindow frame1 = (SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[1]");
                            frame1.Iconify();
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                        }
                        catch (Exception)
                        { }
                        SapVariants.frame.Iconify();
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/n";
                        SapVariants.frame.SendVKey(0);
                        sap.KillSAP();
                    }
                    catch (Exception ex)
                    {
                        mail.SendHTMLMail("Error al subir el archivo a SAP", new string[] {"appmanagement@gbm.net"}, root.Subject);
                    }
                    sap.BlockUser(mandante, 0);
                    #endregion cargar en SAP

                    if (mensaje_sap == "")
                    { mensaje_sap = "Se ha cargado con exito"; }

                    console.WriteLine(mensaje_sap);

                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear backLog Lenovo", mensaje_sap, File_sap);
                    respFinal = respFinal + "\\n" + "Crear backLog Lenovo: " + mensaje_sap + " " + File_sap;


                } //for modificar y cargar a SAP cada hoja del excel
                console.WriteLine("Enviando Archivo a BS");
                xlWorkBook.SaveAs(root.FilesDownloadPath + "\\" + "BL SAP" + " - " + fecha_file + ".xlsx");
                xlWorkBook.Close();

                xlApp.DisplayAlerts = false;
                xlApp.Workbooks.Close();
                xlApp.Quit();
                proc.KillProcess("EXCEL", true);

                //enviar email con el archivo root.Google_Download + "\\" + "BL" + " - " + fecha_file + ".xlsx" a la gente y copia que esta en la base de datos
                string[] cc = bsql.EmailAddress(2);
                //string[] cc = { root.f_copy1, root.f_copy2, root.f_copy3, root.f_copy4, root.f_copy5, root.f_copy6 };

                string[] adjunto = { root.FilesDownloadPath + "\\" + "BL" + " - " + fecha_file + ".xlsx", root.FilesDownloadPath + "\\" + "BL SAP" + " - " + fecha_file + ".xlsx" };
                mail.SendHTMLMail(mensaje_sap + " el BackLog en SAP", new string[] { root.f_sender }, root.Subject, cc, adjunto);
            }
            #endregion validacion


        }
    }
}
