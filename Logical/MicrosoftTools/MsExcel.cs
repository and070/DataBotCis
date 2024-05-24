using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using HtmlAgilityPack;
using ClosedXML.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.Processes;

namespace DataBotV5.Logical.MicrosoftTools
{
    /// <summary>
    /// Clase Logical encargada de MsExcel.
    /// </summary>
    class MsExcel
    {
        ProcessInteraction proc = new ProcessInteraction();
        /// <summary>
        /// Convertir una hoja de excel a un datatable
        /// </summary>
        /// <param name="ruta">el directorio + el nombre del archivo de excel</param>
        /// <param name="headerRow">false: No usar la primera fila como nombre de las columnas</param>
        /// <returns></returns>
        public DataTable GetExcel(string ruta, bool headerRow = true)
        {
            DataTable excel = new DataTable();
            DataSet result = null;
            FileStream stream = null;
            try
            {
                stream = File.Open(ruta, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream);

                result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false,
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = headerRow
                    }
                });

                stream.Close();
                excel = result.Tables[0];

            }
            catch (Exception ex)
            {
                stream.Close();
                if (ex is ExcelDataReader.Exceptions.HeaderException)
                {
                    #region fake xls (html)

                    HtmlDocument doc = new HtmlDocument();
                    try
                    {
                        doc.LoadHtml(File.ReadAllText(ruta));

                        HtmlNodeCollection columns = doc.DocumentNode.SelectNodes("//*[@id='tableContacts']/thead/tr/th");
                        HtmlNodeCollection rows = doc.DocumentNode.SelectNodes("//*[@id='tableContacts']/tbody/tr");

                        foreach (HtmlNode column in columns)
                            try { excel.Columns.Add(column.InnerText.Trim()); }
                            catch (DuplicateNameException) { excel.Columns.Add(column.InnerText.Trim() + "_1"); }

                        foreach (HtmlNode row in rows)
                        {
                            HtmlNodeCollection campo = row.SelectNodes(".//th");
                            DataRow fila = excel.NewRow();
                            for (int i = 0; i < campo.Count; i++)
                                fila[i] = campo[i].InnerText.Trim();
                            excel.Rows.Add(fila);
                        }
                    }
                    catch (Exception)
                    {
                        excel = null;
                    }

                    #endregion
                }
                else
                    excel = null;
            }

            return excel;

        }
        /// <summary>
        /// Convertir un libro de excel a un DataSet
        /// </summary>
        /// <param name="ruta">el directorio + el nombre del archivo de excel</param>
        /// <returns></returns>
        public DataSet GetExcelBook(string ruta)
        {
            DataTable excel = new DataTable();
            DataSet result = null;
            FileStream stream = null;
            try
            {
                stream = File.Open(ruta, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream);

                result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false,
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                stream.Close();

            }
            catch (Exception ex)
            {
                stream.Close();
                if (ex is ExcelDataReader.Exceptions.HeaderException)
                {
                    #region fake xls (html)

                    HtmlDocument doc = new HtmlDocument();
                    try
                    {
                        doc.LoadHtml(File.ReadAllText(ruta));

                        HtmlNodeCollection columns = doc.DocumentNode.SelectNodes("//*[@id='tableContacts']/thead/tr/th");
                        HtmlNodeCollection rows = doc.DocumentNode.SelectNodes("//*[@id='tableContacts']/tbody/tr");

                        foreach (HtmlNode column in columns)
                            try { excel.Columns.Add(column.InnerText.Trim()); }
                            catch (DuplicateNameException) { excel.Columns.Add(column.InnerText.Trim() + "_1"); }

                        foreach (HtmlNode row in rows)
                        {
                            HtmlNodeCollection campo = row.SelectNodes(".//th");
                            DataRow fila = excel.NewRow();
                            for (int i = 0; i < campo.Count; i++)
                                fila[i] = campo[i].InnerText.Trim();
                            excel.Rows.Add(fila);
                        }
                        result.Tables.Add(excel);
                    }
                    catch (Exception)
                    {
                        excel = null;
                        result = null;
                    }

                    #endregion
                }
                else
                    result = null;
            }

            return result;

        }
        /// <summary>
        /// Convierte un datatable a un excel
        /// </summary>
        /// <param name="dt">el datatable con la información</param>
        /// <param name="wsName">el nombre de la hoja de excel</param>
        /// <param name="rute">la ruta y el nombre del archivo donde se desea guardar</param>
        /// <param name="withoutTable">True para crear el excel sin formato de tabla</param>
        public void CreateExcel(DataTable dt, string wsName, string rute, [Optional] bool withoutTable)
        {
            XLWorkbook wbAdmis = new XLWorkbook();
            IXLWorksheet wsAdmis = null;
            if (withoutTable)
            {
                //crear excel sin formato tabla
                wsAdmis = wbAdmis.Worksheets.Add(wsName);
                wsAdmis.Cell(1, 1).InsertTable(dt, false);
            }
            else
            {
                //crear excel con formato de tabla
                wsAdmis = wbAdmis.Worksheets.Add(dt, wsName);
            }
            wsAdmis.Columns().AdjustToContents();
            if (File.Exists(rute))
            {
                File.Delete(rute);
            }
            wbAdmis.SaveAs(rute);
        }
        /// <summary>
        /// Crear una tabla dinámica, ver las referencias en la clase PivotTableParameters para ver los parametros
        /// </summary>
        /// <param name="pivotTableParameters"></param>
        /// <returns></returns>
        public bool CreatePivotTable(PivotTableParameters pivotTableParameters)
        {
            try
            {


                Excel.Application xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(pivotTableParameters.route);
                Excel.Worksheet WSource = (Excel.Worksheet)xlWorkBook.Sheets[pivotTableParameters.sourceSheetName];
                Excel.Worksheet xlWorkSheet = null;

                if (pivotTableParameters.newSheet)
                {
                    xlWorkSheet = xlWorkBook.Sheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);
                    xlWorkSheet.Name = pivotTableParameters.newSheetName;
                    xlWorkSheet.Cells[1, 1] = pivotTableParameters.title;
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 2]].Merge(Type.Missing);

                }
                else
                {
                    xlWorkSheet = WSource;
                }
                Excel.Range sRange = xlWorkSheet.Range[pivotTableParameters.destinationRange];
                Excel.Range oRange = WSource.Range[pivotTableParameters.firstSourceCell, pivotTableParameters.endSourceCell];

                Excel.PivotCache cache = (Excel.PivotCache)xlWorkBook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, oRange);
                Excel.PivotTable pivot = (Excel.PivotTable)xlWorkSheet.PivotTables().Add(PivotCache: cache, TableDestination: sRange, TableName: pivotTableParameters.pivotTableName);
                pivot.NullString = "0";
                bool currency = false;
                foreach (PivotTableFields ptField in pivotTableParameters.pivotTableFields)
                {
                    Excel.PivotField mField = (Excel.PivotField)pivot.PivotFields(ptField.sourceColumnName);


                    switch (ptField.pivotFieldOrientation)
                    {
                        case "xlRowField":
                            mField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                            break;
                        case "xlColumnField":
                            mField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                            break;
                        case "xlPageField":
                            mField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                            break;
                        case "xlHidden":
                            mField.Orientation = Excel.XlPivotFieldOrientation.xlHidden;
                            break;
                        case "xlDataField":

                            if (ptField.numberFormat.Contains("$"))
                            {
                                //mField.NumberFormat = ptField.numberFormat;
                                currency = true;
                            }
                            mField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                            if (ptField.consolidationFunction == "xlSum")
                            {
                                mField.Function = Excel.XlConsolidationFunction.xlSum;
                            }
                            else if (ptField.consolidationFunction == "xlCount")
                            {
                                mField.Function = Excel.XlConsolidationFunction.xlCount;
                            }
                            else if (ptField.consolidationFunction == "xlAverage")
                            {
                                mField.Function = Excel.XlConsolidationFunction.xlAverage;
                            }
                            else if (ptField.consolidationFunction == "xlProduct")
                            {
                                mField.Function = Excel.XlConsolidationFunction.xlProduct;
                            }
                            mField.Name = ptField.fieldName;
                            break;
                        default:
                            break;
                    }


                }


                pivot.TableStyle = "PivotStyleMedium9";
                pivot.TableStyle2 = "PivotStyleMedium9";
                xlWorkBook.SaveAs(pivotTableParameters.route);

                if (currency)
                {

                    string column = pivotTableParameters.destinationRange.Substring(0, 1).ToUpper();
                    int columns = xlWorkSheet.UsedRange.Columns.Count;
                    char let = char.Parse(column);

                    char sChar = (char)(((int)let) + 1);
                    string startColumn = sChar.ToString();

                    char nextChar = (char)(((int)let) + columns - 1);
                    string endColumn = nextChar.ToString();

                    int rows = xlWorkSheet.UsedRange.Rows.Count;
                    xlWorkSheet.Range[$"{startColumn}:{endColumn}"].NumberFormat = "[$$-en-US] #,##0.00";

                    xlWorkBook.SaveAs(pivotTableParameters.route);
                }

                xlWorkBook.Close();

                proc.KillProcess("EXCEL", true);
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }
    }
    /// <summary>
    /// <param name="route">La ruta + nombre + extension del archivo de excel donde se toma la información y se guarda nuevamente al final</param>
    /// <param name="sourceSheetName">el nombre de la hoja donde se toma la información de la tabla dinámica</param>
    /// <param name="newSheet">true = se crea una nueva hoja, false = se utiliza la misma hoja donde esta la tabla fuente</param>
    /// <param name="newSheetName">[Optional]Solo en caso newSheet = True: el nombre de la nueva hoja de excel </param>
    /// <param name="pivotTableName">Nombre de la tabla dinámica </param>
    /// <param name="firstSourceCell">la primera celda de la tabla fuente, ejemplo A1</param>
    /// <param name="endSourceCell">última celda de la tabla fuente donde se toma la información, ejemplo N205</param>
    /// <param name="destinationRange">La celda donde se pondra la tabla dinamica, ejemplo A2 (no poner A1 en caso de que newSheet = True) </param>
    /// <param name="title">Solo en caso newSheet = True: el titulo de la nueva tabla</param>
    /// <param name="pivotTableFields">Una Lista de pivotTableFields de la tabla dinámica</param>
    /// </summary>
    public class PivotTableParameters
    {
        public string route { get; set; }
        public string sourceSheetName { get; set; }
        public bool newSheet { get; set; }
        public string newSheetName { get; set; }
        public string pivotTableName { get; set; }
        public string firstSourceCell { get; set; }
        public string endSourceCell { get; set; }
        public string destinationRange { get; set; }
        public string title { get; set; }
        public List<PivotTableFields> pivotTableFields { get; set; }
    }
    /// <summary>
    /// <param name="sourceColumnName">El nombre de la columna de la tabla donde se va a tomar la información</param>
    /// <param name="PivotFieldOrientation">indicar el tipo del campo: xlRowField, xlColumnField, xlPageField, xlDataField</param>
    /// <param name="ConsolidationFunction">Solo en caso PivotFieldOrientation = xlDataField, es la formula del campo valor: xlSum, xlCount, xlAverage, xlProduct</param>
    /// <param name="fieldName">Solo en caso PivotFieldOrientation = xlDataField, el nombre del campo data</param>
    /// </summary>
    public class PivotTableFields
    {
        public string sourceColumnName { get; set; }
        public string pivotFieldOrientation { get; set; }
        public string consolidationFunction { get; set; }
        public string fieldName { get; set; }
        public string numberFormat { get; set; }
    }
}
