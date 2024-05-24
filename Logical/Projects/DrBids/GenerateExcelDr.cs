using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using static DataBotV5.Automation.MASS.DrBids.GetDrBids;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Projects.DrBids;
using static DataBotV5.Data.Projects.DrBids.BidsGbDrSql;
using DataBotV5.App.Global;

namespace DataBotV5.Logical.Projects.DrBids
{
    /// <summary>
    /// Clase Logical encargada de generar excel de Républica Dominicana.
    /// </summary>
    class GenerateExcelDr 
    {

        Rooting rooting = new Rooting();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proccess = new ProcessInteraction();
        /// <summary>
        /// Genera los excel de cada vendedor asociado con sus licitaciones y le envía el reporte por email.
        /// </summary>
        /// <param name="list"> Lista de vendedores que a los que se le enviara el informe de excel</param>
        public void FileExcel(List<SellerBid> list)
        {
            for (int a = 0; a < list.Count; a++)
            {
                try
                {
                    Application xlApp;
                    Workbook xlWorkBook;
                    Worksheet xlWorkSheet;
                    xlApp = new Application();
                    xlApp.Visible = false;
                    xlApp.DisplayAlerts = false;
                    xlWorkBook = xlApp.Workbooks.Add();
                    xlWorkSheet = (Worksheet)xlWorkBook.Sheets[1];
                    int cont = 2;
                    List<string> datos = ColumnsExcel();
                    for (int b = 0; b < datos.Count; b++)
                    {
                        xlWorkSheet.Cells[1, b + 1].value = datos[b];
                        xlWorkSheet.Cells[1, b + 1].EntireRow.Font.Bold = true;
                    }
                    xlWorkSheet.Range["D:D"].NumberFormat = "DD/MM/YYYY hh:mm";
                    xlWorkSheet.Range["M:AZ"].NumberFormat = "DD/MM/YYYY hh:mm";

                    for (int i = 1; i < list[a].listDatosGenerales.Count + 1; i++)
                    {
                        //Agregar columnas de Cronograma
                        foreach (var item in list[a].listDatosGenerales[i - 1].cronograma)
                        {
                            var result = datos.Where(x => x.Equals(item.Key.Replace("_", " "))).ToList();
                            if (result.Count == 0)
                            {
                                xlWorkSheet.Cells[1, datos.Count + 1].value = item.Key.Replace("_", " ");
                                datos.Add(item.Key.Replace("_", " "));
                            }
                        }
                        //Agregar articulos y datos generales
                        for (int j = 0; j < list[a].listDatosGenerales[i - 1].listaArticulos.Count; j++)
                        {
                            xlWorkSheet.Cells[cont, 1].value = list[a].listDatosGenerales[i - 1].cliente;
                            xlWorkSheet.Cells[cont, 2].value = list[a].listDatosGenerales[i - 1].referencia;
                            xlWorkSheet.Cells[cont, 3].value = list[a].listDatosGenerales[i - 1].descripcion;
                            xlWorkSheet.Cells[cont, 4].value = list[a].listDatosGenerales[i - 1].fechaPublicacion;
                            xlWorkSheet.Cells[cont, 5].value = list[a].listDatosGenerales[i - 1].presupuesto;
                            xlWorkSheet.Cells[cont, 6].value = list[a].listDatosGenerales[i - 1].interesGBM;
                            xlWorkSheet.Cells[cont, 7].value = list[a].listDatosGenerales[i - 1].listaArticulos[j].Codigo;
                            xlWorkSheet.Cells[cont, 8].value = list[a].listDatosGenerales[i - 1].listaArticulos[j].descripcion_Articulo;
                            xlWorkSheet.Cells[cont, 9].value = list[a].listDatosGenerales[i - 1].listaArticulos[j].cantidad;
                            xlWorkSheet.Cells[cont, 10].value = list[a].listDatosGenerales[i - 1].listaArticulos[j].unidad;
                            xlWorkSheet.Cells[cont, 11].value = list[a].listDatosGenerales[i - 1].listaArticulos[j].precio_Unitario;
                            xlWorkSheet.Cells[cont, 12].value = list[a].listDatosGenerales[i - 1].listaArticulos[j].Precio_total;

                            int indextemp = 0;
                            List<string> listKeys = new List<string>();
                            List<DateTime> listValues = new List<DateTime>();
                            //Agrega el cronograma 
                            foreach (var item in list[a].listDatosGenerales[i - 1].cronograma)
                            {
                                listKeys.Add(item.Key.Replace("_", " "));
                                listValues.Add(item.Value);
                            }
                            for (int h = 0; h < listKeys.Count; h++)
                            {
                                int index = datos.IndexOf(listKeys[h]) + 1;
                                xlWorkSheet.Cells[cont, index].value = listValues[h];
                                indextemp++;
                            }
                            cont++;
                        }
                    }
                    //
                    xlWorkSheet.Columns.AutoFit();
                    xlWorkSheet.Columns[3].ColumnWidth = 120;
                    xlWorkSheet.Columns[8].ColumnWidth = 100;

                    string FilePath = rooting.ExcelDr[0] + "Licitaciones_" + list[a].nombreVendedor + $"_{DateTime.Now.Day}_{DateTime.Now.Month}_{DateTime.Now.Year}" + ".xlsx";
                    if (File.Exists(FilePath))
                    {
                        File.Delete(FilePath);
                    }
                    xlWorkBook.SaveAs(FilePath);
                    xlWorkBook.Close();
                    xlApp.Workbooks.Close();
                    xlApp.Quit();
                    proccess.KillProcess("EXCEL",true);
                    string html = Properties.Resources.emailtemplate1;
                    html = html.Replace("{subject}", "Licitaciones Dominicana");
                    html = html.Replace("{cuerpo}", "Estimado " + list[a].nombreVendedor + " se le adjunta el reporte de los procesos de compras del dia " + DateTime.Now.ToString("dd/MM/yyyy"));
                    html = html.Replace("{contenido}", "");
                    string[] ruta = { FilePath };
                    mail.SendHTMLMail(html, new string[] { list[a].nombreVendedor + "@gbm.net" }, "Informe de Excel con licitaciones del Gobierno de Republica Dominicana", null, ruta);
                }
                catch (Exception e)
                {
                    new ConsoleFormat().WriteLine(e.ToString());
                }
            }
        }
        /// <summary>
        /// Método encargado de agregar nombre a las futuras columnas del informe de excel.
        /// </summary>
        /// <returns></returns>
        private List<string> ColumnsExcel()
        {
            List<string> columnas = new List<string>();
            columnas.Add("Nombre de Cliente");
            columnas.Add("Id de licitación");
            columnas.Add("Descripción de licitación");
            columnas.Add("Fecha de publicación");
            columnas.Add("Presupuesto");
            columnas.Add("Intereses de GBM");
            columnas.Add("Codigo Articulo");
            columnas.Add("Descripcion Articulo");
            columnas.Add("Cantidad de Articulos");
            columnas.Add("Unidades");
            columnas.Add("Precion Unitario");
            columnas.Add("Precio Total");
            return columnas;
        }
        /// <summary>
        /// Metodo para traer los vendedores desde la base de datos para el informe de excel y asignarles las licitaciones que tienen
        /// </summary>
        /// <param name="list"> Lista con todos los datos de lsa licitaicones, traidos de la pagina de licitaciones de dominicana</param>
        /// <returns></returns>
        public List<SellerBid> GenerateSellers(List<GeneralData> list)
        {

            BidsGbDrSql db_Do_sql = new BidsGbDrSql();
            List<SellerBid> vendedoresLicitacion = new List<SellerBid>();
            List<ClientSAP> listaVendedoresSAP = new List<ClientSAP>();
            listaVendedoresSAP = db_Do_sql.FetchSAPCostumers();

            #region Client
            var validarVendedorDefault = listaVendedoresSAP.Where(x => "Unidad de Compra".Contains(x.unidadCompra)).ToList();
            string vendedorDefault = validarVendedorDefault[0].nombreVendedor;
            SellerBid vendedor = new SellerBid()
            {
                nombreVendedor = vendedorDefault,
                listDatosGenerales = new List<GeneralData>()
            };
            vendedoresLicitacion.Add(vendedor);
            #endregion

            for (int i = 0; i < list.Count; i++)
            {
                //Buscar si el vendedor esta en la tabla de omologacion
                var result = listaVendedoresSAP.Where(x => x.unidadCompra.Contains(list[i].cliente)).ToList();
                SellerBid VendedorLicitacion = new SellerBid()
                {
                    nombreVendedor = "",
                    listDatosGenerales = new List<GeneralData>()
                };
                if (result.Count > 0)
                {
                    //Buscar si el vendedor ya esta ingresado en la lista de vendedores
                    var validarExisteVendedor = vendedoresLicitacion.Where(x => result[0].nombreVendedor.Equals(x.nombreVendedor)).ToList();

                    if (validarExisteVendedor.Count > 0)
                    {
                        //Agrega una licitacion a un vendedor que ya esta agregado en la lista de vendedores
                        validarExisteVendedor[0].listDatosGenerales.Add(list[i]);
                    }
                    else
                    {
                        //Si el vendedor no existe en la lista de vendedores lo crea
                        VendedorLicitacion.listDatosGenerales.Add(list[i]);
                        VendedorLicitacion.nombreVendedor = result[0].nombreVendedor;
                        vendedoresLicitacion.Add(VendedorLicitacion);
                    }
                }
                else
                {
                    //Si la licitacion no tiene vendedor asociado se le asigna al vendedor por default
                    var validarExisteVendedor = vendedoresLicitacion.Where(x => vendedorDefault.Contains(x.nombreVendedor)).ToList();
                    validarExisteVendedor[0].listDatosGenerales.Add(list[i]);
                }
            }
            return vendedoresLicitacion;
        }

        public class SellerBid
        {
            public List<GeneralData> listDatosGenerales { get; set; }
            public string nombreVendedor { get; set; }
        }
    }
}
