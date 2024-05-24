using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Automation.RPA2.OrdersToAssign
{
    public class OrdersToAssign
    {
        ConsoleFormat console = new ConsoleFormat();
        SharePoint sharepoint = new SharePoint();
        Log log = new Log();
        MsExcel MsExcel = new MsExcel();
        Stats estadisticas = new Stats();
        CRUD crud = new CRUD();
        Rooting root = new Rooting();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        string mand = "QAS";

        /// <summary>
        /// Robot encargado de la obtención de informacion de ordenes para inventario (Archivo que es generado por el bot de Backlog Lenovo) y ordenes pendientes por asignar de Lenovo (Archivo obtenido del sharepoint del equipo de BusinessSystem),
        /// al obtener la información es cargada en la base de datos "ordersForInventory", para el manejo de dicha información desde el portal de Pedidos para inventario en Smart Simple, que es administrado por que el equipo de Business System.
        /// </summary>
        public void Main()
        {
            string orderInStockName = "Órdenes en Stock In Transit 1.xlsx";

            string fecha = DateTime.Now.ToString("d_M_yyyy");
            string name = "BL" + " - " + fecha + ".xlsx";

            bool validate = sharepoint.DownloadFileFromSharePoint("https://gbmcorp.sharepoint.com/sites/PurchasingLenovo/", "Documentos", name, "Backlog Lenovo");

            if (validate)
            {
                readAndInsertOrderPending(orderInStockName);
                readAndInsertReport(root.FilesDownloadPath + "\\" + name, root.FilesDownloadPath + $"\\{orderInStockName}");

            }
            else
            {
                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Error al descargar");
                html = html.Replace("{cuerpo}", $"No se pudo descargar el archivo: {name}, del SharePoint: PurchasingLenovo");
                html = html.Replace("{contenido}", "");
                console.WriteLine("Send Email...");
                BsSQL bs = new BsSQL();
                string[] cc = bs.EmailAddress(10);
                mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Error al descargar archivo: {name}", cc, null);
            }
        }
        /// <summary>
        /// Metodo encargado de la lectura de archivos de excel que contienen las ordenes para inventario de Lenovo y subir dicha informacion a la tabla "reportUS" en la base de datos "ordersForInventory".
        /// </summary>
        /// <param name="ruta">Ruta para obtener el excel de ordenes para inventario</param>
        /// <param name="ruta2">Ruta para obtener las ordenes pendientes para asignar</param>
        private void readAndInsertReport(string ruta, string ruta2)
        {

            string actualShipDate;
            string deliveryDate;
            string firmShipDate;
            string orderEntryDate;
            double quantityReserved = 0;
            bool error = false;
            string errorRows = "";
            #region Abrir excel y filtrao inicial
            DataTable excel = MsExcel.GetExcel(ruta);
            DataSet excelBook = MsExcel.GetExcelBook(ruta2);
            DataTable excel2 = excelBook.Tables["Órdenes para asignar"];
            for (int i = excel.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = excel.Rows[i];
                if (!dr["Customer Purchase Order Number"].ToString().Contains("ST") || dr["Order Status"].ToString() == "Delivered" || dr["Order Status"].ToString() == "Cancelled")
                {
                    dr.Delete();
                }
            }
            excel.AcceptChanges();

            #endregion

            #region eliminacion de columnas
            excel.Columns.Remove("Sold To Customer Number");
            excel.Columns.Remove("Sold To Customer Name");
            excel.Columns.Remove("Order Receipt Date");
            excel.Columns.Remove("Order Status");
            excel.Columns.Remove("Line Item Status");
            excel.Columns.Remove("Product Description");
            excel.Columns.Remove("Unit Price");
            excel.Columns.Remove("Shipped Quantity");
            excel.Columns.Remove("Invoice Number");
            excel.Columns.Remove("Estimated Ship Date");
            excel.Columns.Remove("Total Amount in Document Currency");
            excel.Columns.Remove("Carrier Name");
            excel.Columns.Remove("Carrier Tracking Number");
            excel.Columns.Remove("Serial Number");
            excel.Columns.Remove("Mode Of Transportation");


            #endregion


            string sql = "SELECT * FROM reportUS";
            string getOrdersAssign = "SELECT * FROM ordesAssign";

            DataTable registerDB = crud.Select(sql, "ordersForInventory", mand);


            foreach (DataRow item in excel2.Rows)
            {
                object po = item["PO País"];
                object so = item["SO"];
                object pn = item["PN"];

                if (po is string)
                {
                    item["PO País"] = Convert.ToDouble(0);
                }
                if (so is string)
                {
                    item["SO"] = Convert.ToDouble(0);
                }
                if (pn.ToString() == "Pendiente")
                {
                    item["PN"] = Convert.ToDouble(0);
                }
            }
            excel2.AcceptChanges();


            foreach (DataRow row in excel.Rows)
            {
                try
                {
                    #region Validaciones de Fechas
                    if (row["Actual Ship Date"].ToString() == "")
                    {

                        if (row["Firm Ship Date"].ToString() == "")
                        {
                            actualShipDate = "9999-12-12 00:00:00";
                        }
                        else
                        {
                            actualShipDate = DateTime.Parse(row["Firm Ship Date"].ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    else
                    {
                        DateTime date = DateTime.Parse(row["Actual Ship Date"].ToString());
                        if (date.ToString() == "")
                        {
                            actualShipDate = "9999-12-12 00:00:00";
                        }
                        else
                        {
                            date.AddDays(4);
                            actualShipDate = DateTime.Parse(date.ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }

                    if (row["Actual Delivery Date"].ToString() == "")
                    {
                        if (row["Estimated Delivery Date"].ToString() == "")
                        {
                            deliveryDate = "9999-12-12 00:00:00";
                        }
                        else
                        {
                            deliveryDate = DateTime.Parse(row["Estimated Delivery Date"].ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    else
                    {
                        DateTime date = DateTime.Parse(row["Actual Delivery Date"].ToString());
                        if (date.ToString() == "")
                        {
                            deliveryDate = "9999-12-12 00:00:00";
                        }
                        else
                        {
                            date.AddDays(3);
                            deliveryDate = DateTime.Parse(date.ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    if (row["Order Entry Date"].ToString() == "")
                    {
                        orderEntryDate = "9999-12-12 00:00:00";
                    }
                    else
                    {
                        orderEntryDate = DateTime.Parse(row["Order Entry Date"].ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                    }

                    if (row["Firm Ship Date"].ToString() == "")
                    {
                        firmShipDate = "9999-12-12 00:00:00";
                    }
                    else
                    {
                        firmShipDate = DateTime.Parse(row["Firm Ship Date"].ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                    }

                    #endregion

                    console.WriteLine("Agrupando Registros");

                    #region Agrupar Registros
                    var groupedRows = from line in excel2.AsEnumerable()
                                      group line by new { PN = line.Field<string>("PN"), SO = line.Field<double>("SO Lenovo") }
                                      into grp
                                      select new { PN = grp.Key.PN, SO = grp.Key.SO, TotalQUA = grp.Sum(r => r.Field<double>("Cantidad Reservada")) };
                    string rowPN = row["Product ID"].ToString();
                    string rowSO = row["Sales Order Number"].ToString();

                    foreach (var group in groupedRows)
                    {
                        if (rowPN == group.PN && rowSO == group.SO.ToString())
                        {
                            quantityReserved = group.TotalQUA;
                            break;
                        }
                    }

                    #endregion


                    string getAllReportUs = "SELECT * FROM reportUS WHERE POTrad = '" + row["Customer Purchase Order Number"].ToString() + "' AND PN = '" + row["Product ID"].ToString() + "' AND SO = '" + row["Sales Order Number"].ToString() + "'";

                    DataTable reportUS = crud.Select(getAllReportUs, "ordersForInventory", mand);
                    if (reportUS.Rows.Count != 0)
                    {
                        console.WriteLine("Actualizando Registros de Rerport US: SO: " + row["Sales Order Number"].ToString() + " PN: " + row["Product ID"].ToString() + " POTRAD: " + row["Customer Purchase Order Number"].ToString() + " cantidad a restar: " + quantityReserved);

                        string sqlUpdate = "UPDATE `reportUS` SET `ordenEntryDate`='" + orderEntryDate + "',`firmShipDate`='" + firmShipDate + "',`ordenQuantity`='" + row["Order Quantity"].ToString() + "',`actualShipDate`='" + actualShipDate + "'" +
                            ",`deliveryDate`='" + deliveryDate + "',`quantityAvailable`= " + Convert.ToInt64((Convert.ToDouble(row["Order Quantity"].ToString()) - quantityReserved)) + ",`quantityReserved`= " + quantityReserved +
                            " WHERE POTrad = '" + row["Customer Purchase Order Number"].ToString() + "' AND PN = '" + row["Product ID"].ToString() + "' AND SO = '" + row["Sales Order Number"].ToString() + "'";

                        crud.Update(sqlUpdate, "ordersForInventory", mand);
                        quantityReserved = 0;
                    }
                    else if (registerDB.Select("POTrad = '" + row["Customer Purchase Order Number"].ToString() + "'").Count() == 0 || registerDB.Select("PN = '" + row["Product ID"].ToString() + "'").Count() == 0 || registerDB.Select("SO = '" + row["Sales Order Number"].ToString() + "'").Count() == 0)
                    {

                        console.WriteLine("Insertando Registros de Rrport US: SO: " + row["Sales Order Number"].ToString() + " PN: " + row["Product ID"].ToString() + " POTRAD: " + row["Customer Purchase Order Number"].ToString());
                        string sqlInsert = "INSERT INTO `reportUS`(`ordenEntryDate`, `firmShipDate`, `SO`, `POTrad`, `PN`, `ordenQuantity`, `actualShipDate`, `deliveryDate`, `quantityAvailable`, `quantityReserved`) VALUES (" +
                    "'" + orderEntryDate + "','" + firmShipDate + "','" + row["Sales Order Number"].ToString() + "','" + row["Customer Purchase Order Number"].ToString() + "','" + row["Product ID"].ToString()
                    + "','" + row["Order Quantity"].ToString() + "','" + actualShipDate + "','" + deliveryDate + "'," + Convert.ToInt64((Convert.ToDouble(row["Order Quantity"].ToString()) - quantityReserved)) + "," + quantityReserved + ")";

                        crud.Insert(sqlInsert, "ordersForInventory", mand);
                        quantityReserved = 0;
                    }
                }
                catch (Exception ex)
                {
                    error = true;
                    StringBuilder errorBuilder = new StringBuilder();
                    errorBuilder.Append("<tr>");
                    errorBuilder.Append($"<td>{row["Customer Purchase Order Number"]}</td>");
                    errorBuilder.Append($"<td>{row["Product ID"]}</td>");
                    errorBuilder.Append($"<td>{ row["Sales Order Number"]}</td>");
                    errorBuilder.Append($"<td>{row["Order Quantity"]}</td>");
                    errorBuilder.Append($"<td>{ex.Message}</td>");
                    errorBuilder.Append($"</tr>");
                    errorRows += errorBuilder;

                }
            }

            if (error == true)
            {

                StringBuilder strHTMLBuilder = new StringBuilder();
                strHTMLBuilder.Append("<table class='myCustomTable' width='100 %'>");
                strHTMLBuilder.Append("<thead>");
                strHTMLBuilder.Append("<tr>");
                strHTMLBuilder.Append("<th>PO</th>");
                strHTMLBuilder.Append("<th>PN</th>");
                strHTMLBuilder.Append("<th>SO</th>");
                strHTMLBuilder.Append("<th>Order Quantity</th>");
                strHTMLBuilder.Append($"<th>Error</th>");
                strHTMLBuilder.Append("</thead>");
                strHTMLBuilder.Append("<tbody>");
                strHTMLBuilder.Append(errorRows);
                strHTMLBuilder.Append("</tbody>");
                strHTMLBuilder.Append("</table>");
                string htmlTable = strHTMLBuilder.ToString();

                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Error en insertar las siguientes ordenes para asignar");
                html = html.Replace("{cuerpo}", $"Error en insertar las siguientes ordenes para asignar");
                html = html.Replace("{contenido}", htmlTable);
                console.WriteLine("Send Email...");
                BsSQL bs = new BsSQL();
                string[] cc = bs.EmailAddress(10);
                mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Error en insertar las siguientes ordenes para asignar", cc, null);
            }


        }
        /// <summary>
        /// Metodo encargado de la lectura de un excel que contiene las ordenes pendientes por asignar de Lenovo y luego insertarlo en la tabla "ordesAssign" en la base de datos "ordersForInventory".
        /// </summary>
        /// <param name="name">Nombre del archivo de ordenes en stock en transito</param>
        private void readAndInsertOrderPending(string name)
        {
            bool validate = sharepoint.DownloadFileFromSharePoint("https://gbmcorp.sharepoint.com/sites/PurchasingLenovo/", "Documentos", name, null);
            bool error = false;
            string errorRows = "";
            if (validate == false)
            {
                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Error al descargar");
                html = html.Replace("{cuerpo}", $"No se pudo descargar el archivo: {name}, del SharePoint: PurchasingLenovo");
                html = html.Replace("{contenido}", "");
                console.WriteLine("Send Email...");
                BsSQL bs = new BsSQL();
                string[] cc = bs.EmailAddress(10);
                mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Error al descargar archivo: {name}", cc, null);
            }
            else
            {

                string valueDate;
                #region Abrir excel
                DataSet excelBook = MsExcel.GetExcelBook(root.FilesDownloadPath + $"\\{name}");
                DataTable excel = excelBook.Tables["Órdenes para asignar"];
                console.WriteLine("");

                #endregion

                #region eliminacion de columnas
                excel.Columns.Remove("Status");
                excel.Columns.Remove("PN reemplazo a sugerir");
                #endregion

                string sql = "SELECT * FROM ordesAssign";

                DataTable registerDB = crud.Select(sql, "ordersForInventory", mand);

                foreach (DataRow row in excel.Rows)
                {
                    try
                    {
                        int country = getCountry(row["País"].ToString());
                        if (row["Fecha de recibo de la orden"].ToString() == "Reserva")
                        {
                            valueDate = "9999-12-12 00:00:00";
                        }
                        else
                        {
                            valueDate = row["Fecha de recibo de la orden"].ToString();
                        }
                        string date = DateTime.Parse(valueDate).ToString("yyyy-MM-dd HH:mm:ss");
                        string getAllOrdersToAssign = "SELECT * FROM ordesAssign WHERE countryPO = '" + row["PO País"].ToString() + "' AND PN = '" + row["PN"].ToString() + "' AND SO = '" + row["SO"].ToString() + "' AND SOLenovo = '" + row["SO Lenovo"].ToString() + "'";
                        DataTable ordersToAssign = crud.Select(getAllOrdersToAssign, "ordersForInventory", mand);

                        if (ordersToAssign.Rows.Count != 0)
                        {
                            string sqlUpdate = "UPDATE `ordesAssign` SET `countryOA`='" + country + "',`customer`='" + row["Cliente"].ToString() + "',`dateToRevibeOrder`='" + date + "',`SO`='" + row["SO"].ToString() +
                                "',`countryPO`='" + row["PO País"].ToString() + "',`PN`='" + row["PN"].ToString() + "',`SOLenovo` ='" + row["SO Lenovo"].ToString() + "',`reservedAmount`='" + row["Cantidad Reservada"].ToString() + "',`costSO`='" + row["Costo SO"].ToString() +
                                "',`post`='" + row["Post"].ToString() + "',`user`='" + row["Usuario"].ToString() + "',`comment`='" + row["Comentarios "].ToString() + "'" +
                                "WHERE countryPO = '" + row["PO País"].ToString() + "' AND PN = '" + row["PN"].ToString() + "' AND SO = '" + row["SO"].ToString() + "' AND SOLenovo = '" + row["SO Lenovo"].ToString() + "'";

                            console.WriteLine("Actualizando Orden pendiente para asignar: SO: " + row["SO"].ToString() + " PN: " + row["PN"].ToString() + " PO: " + row["PO País"].ToString() + " SO LEnovo:: " + row["SO Lenovo"].ToString());
                            crud.Update(sqlUpdate, "ordersForInventory", mand);
                        }
                        else
                        {
                            string sqlInsert = "INSERT INTO `ordesAssign`(`countryOA`, `customer`, `dateToRevibeOrder`, `SO`, `countryPO`, `PN`,`SOLenovo`, `reservedAmount`, `costSO`,`StatusOrderAssign`, `post`, `user`, `comment`) VALUES (" +
                                country + ",'" + row["Cliente"].ToString() + "','" + date + "','" + row["SO"].ToString() + "','" + row["PO País"].ToString() + "','" + row["PN"].ToString() + "','" + row["SO Lenovo"].ToString() + "','"
                                + row["Cantidad Reservada"].ToString() + "','" + row["Costo SO"].ToString() + "'," + 1 + ",'" + row["Post"].ToString() + "','" + row["Usuario"].ToString() + "','" + row["Comentarios "].ToString() + "')";

                            console.WriteLine("Insertando Orden pendiente para asignar: SO: " + row["SO"].ToString() + " PN: " + row["PN"].ToString() + " PO: " + row["PO País"].ToString() + " SO LEnovo:: " + row["SO Lenovo"].ToString());
                            crud.Insert(sqlInsert, "ordersForInventory", mand);
                        }
                    }
                    catch (Exception ex)
                    {
                        error = true;
                        StringBuilder errorBuilder = new StringBuilder();
                        errorBuilder.Append("<tr>");
                        errorBuilder.Append($"<td>{row["PO País"]}</td>");
                        errorBuilder.Append($"<td>{row["PN"]}</td>");
                        errorBuilder.Append($"<td>{ row["SO"]}</td>");
                        errorBuilder.Append($"<td>{ex.Message}</td>");
                        errorBuilder.Append($"</tr>");
                        errorRows += errorBuilder;
                    }
                }
            }
            if (error == true)
            {
                StringBuilder strHTMLBuilder = new StringBuilder();
                strHTMLBuilder.Append("<table class='myCustomTable' width='100 %'>");
                strHTMLBuilder.Append("<thead>");
                strHTMLBuilder.Append("<tr>");
                strHTMLBuilder.Append("<th>PO</th>");
                strHTMLBuilder.Append("<th>PN</th>");
                strHTMLBuilder.Append("<th>SO</th>");
                strHTMLBuilder.Append($"<th>Error</th>");
                strHTMLBuilder.Append("</thead>");
                strHTMLBuilder.Append("<tbody>");
                strHTMLBuilder.Append(errorRows);
                strHTMLBuilder.Append("</tbody>");
                strHTMLBuilder.Append("</table>");
                string htmlTable = strHTMLBuilder.ToString();

                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "Error en insertar las siguientes ordenes pendientes asignar");
                html = html.Replace("{cuerpo}", $"Error en insertar las siguientes ordenes pendientes asignar");
                html = html.Replace("{contenido}", htmlTable);
                console.WriteLine("Send Email...");
                BsSQL bs = new BsSQL();
                string[] cc = bs.EmailAddress(10);
                mail.SendHTMLMail(html, new string[] { root.f_sender }, $"Error en insertar las siguientes ordenes pendientes para asignar", cc, null);
            }


        }
        /// <summary>
        /// Metodo de ayuda, encargado de la devolucioón del número de pais que se tiene en la base de datos de "ordersForInventory".
        /// </summary>
        /// <param name="country">Nombre del país que se busca obtener el número en la base de datos</param>
        /// <returns></returns>
        private int getCountry(string country)
        {
            Object[,] countries = new Object[,] { { "COSTA RICA", 1 }, { "GUATEMALA", 2 }, { "PANAMÁ", 3 }, { "PANAMA", 3 }, { "DOMINICANA", 4 }, { "NICARAGUA", 5 }, { "HONDURAS", 6 }, { "SALVADOR", 7 }, { "MIAMI", 8 }, { "COLOMBIA", 9 } };

            for (int i = 0; i < countries.GetLength(0); i++)
            {
                string temp = countries[i, 0].ToString();
                if (temp.Equals(country))
                {
                    return Int32.Parse(countries[i, 1].ToString());
                }
            }
            console.WriteLine(country);

            return 0;
        }
    }

}
