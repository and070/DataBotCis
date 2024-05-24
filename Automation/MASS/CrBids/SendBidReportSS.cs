using ClosedXML.Excel;
using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Projects.CrBids;
using Newtonsoft.Json.Linq;
using System;
using System.Data;
using System.Linq;

namespace DataBotV5.Automation.MASS.CrBids
{
    /// <summary>
    /// Clase MASS Automation "Robot 6" encargada de enviar un reporte semanal YTD(1 de enero a fecha actual) al DMO de concursos SICOP de la base de datos actuales y de Backup de CostaRicaBids de SmartAndSimple.
    /// </summary>
    class SendBidReportSS
    {
        #region variables_globales
        string enviroment = "QAS";

        CrBidsLogical cr_licitaciones = new CrBidsLogical();
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();
        BidsGbCrSql lcsql = new BidsGbCrSql();
        CRUD crud = new CRUD();
        Log log = new Log();
        string respFinal = "";


        internal CrBidsLogical Cr_licitaciones { get => cr_licitaciones; set => cr_licitaciones = value; }
        internal Stats Estadisticas { get => estadisticas; set => estadisticas = value; }
        internal BidsGbCrSql Lcsql { get => lcsql; set => lcsql = value; }
        #endregion

        public void Main()
        {
            console.WriteLine(" Procesando...");
            SendBidReportMethod();
            using (Stats stats = new Stats())
            {
                stats.CreateStat();
            }
        }


        /// <summary>
        /// Este robot envía un reporte de Excel al correo del DMOLíder (extraído de la BD de Databot en FabricaDeOFertas-Empleados), en el reporte 
        /// lleva todas las licitaciones de Costa Rica existentes tanto activas como backup en el lapso del primer día del presente año a la fecha actual, 
        /// donde cada fila lleva el product de la licitación junto a la información relacionada tanto de la purchaseOrder, 
        /// purchaseOrderAdditionalData, evaluations (el cual se junta los diferentes registros en una sola celda), entre otros.
        /// </summary>
        public void SendBidReportMethod()
        {
            console.WriteLine("Generando reporte de bids YTD actuales y backup al DMO de Costa Rica Bids");

            #region Extracción principal de datos de DB.

                //se debe extraer el backup en 2 tractos debido a la cantidad de data.
                //por lo que se extrae la fecha intermedia entre el 01/01 del presente año a la fecha de hoy
                DateTime firstDayYear = DateTime.Parse($"{DateTime.Now.Year}-01-01 00:00:00");
                DateTime currentDay = DateTime.Parse($"{DateTime.Now.ToString("yyyy-MM-dd")} 23:59:00");
                string firstDayYearS = firstDayYear.ToString("yyyy-MM-dd HH:mm:ss");
                string currentDayS = currentDay.ToString("yyyy-MM-dd HH:mm:ss");

                console.WriteLine("Extraer Purchase Orders actuales");

                    #region PurchaseOrders actuales del 1 de enero a la fecha
                
                    string sqlPurchaseOrders = $"SELECT * FROM `purchaseOrder` WHERE publicationDate BETWEEN '{firstDayYearS}' and '{currentDayS}'";
                    DataTable purchaseOrders = crud.Select( sqlPurchaseOrders, "costa_rica_bids_db");

                    string auxSelectPO = $"SELECT id FROM `purchaseOrder` WHERE publicationDate BETWEEN '{firstDayYearS}' and '{currentDayS}'";

                    string sqlPOAditionalData = $"SELECT * FROM `purchaseOrderAdditionalData` WHERE bidNumber in ({auxSelectPO})";
                    DataTable POAditionalData = crud.Select( sqlPOAditionalData, "costa_rica_bids_db");

                    string sqlEvaluations = $"SELECT * FROM `evaluations` WHERE bidNumber in ({auxSelectPO})";
                    DataTable evaluations = crud.Select( sqlEvaluations, "costa_rica_bids_db");

                    string sqlProducts = $"SELECT * FROM `products` WHERE bidNumber in ({auxSelectPO})";
                    DataTable products = crud.Select( sqlProducts, "costa_rica_bids_db");

                #endregion

                    #region PurchaseOrders Backup extraída en 2 tractos.

                    DataTable purchaseOrdersBackup1 = new DataTable();
                    DataTable POAditionalDataBackup1 = new DataTable();
                    DataTable evaluationsBackup1 = new DataTable();
                    DataTable productsBackup1 = new DataTable();

                    //Este caso es por si es 1 de enero para únicamente extraer este día de los backups.
                    if (firstDayYear.ToString("dd/MM/yyyy") == currentDay.ToString("dd/MM/yyyy"))
                    {
                        string sqlPurchaseOrdersBackup1 = $"SELECT * FROM `purchaseOrderBackup` WHERE publicationDate BETWEEN '{firstDayYearS}' and '{currentDayS}'";
                        purchaseOrdersBackup1 = crud.Select( sqlPurchaseOrdersBackup1, "costa_rica_bids_db");

                        string auxSelectPOBackUp = $"SELECT id FROM `purchaseOrderBackup` WHERE publicationDate BETWEEN '{firstDayYearS}' and '{currentDayS}'";

                        string sqlPOAditionalDataBackup1 = $"SELECT * FROM `purchaseOrderAdditionalDataBackup` WHERE bidNumber in ({auxSelectPOBackUp})";
                        POAditionalDataBackup1 = crud.Select( sqlPOAditionalDataBackup1, "costa_rica_bids_db");

                        string sqlEvaluationsBackup1 = $"SELECT * FROM `evaluationsBackup` WHERE bidNumber in ({auxSelectPOBackUp})";
                        evaluationsBackup1 = crud.Select( sqlEvaluationsBackup1, "costa_rica_bids_db");

                        string sqlProductsBackup1 = $"SELECT * FROM `products` WHERE bidNumber in ({auxSelectPOBackUp})";
                        productsBackup1 = crud.Select( sqlProductsBackup1, "costa_rica_bids_db");


                    }
                    else //De lo contrario extrae la información en diferentes tractos debido a la gran cantidad de Data.
                    {
                        TimeSpan currentAmountDaysYear = currentDay.Subtract(firstDayYear);
                        DateTime middleTime = firstDayYear.AddMinutes(currentAmountDaysYear.TotalMinutes / 2);
                        DateTime middleTimeNext = middleTime.AddDays(1);

                        string middleTimeS = middleTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string middleTimeF = middleTimeNext.ToString("yyyy-MM-dd") + " 00:00:00";

                        #region Primer extracto de Data de Backup
                        string sqlPurchaseOrdersBackup1 = $"SELECT * FROM `purchaseOrderBackup` WHERE publicationDate BETWEEN '{firstDayYearS}' and '{middleTimeS}'";
                        purchaseOrdersBackup1 = crud.Select( sqlPurchaseOrdersBackup1, "costa_rica_bids_db");

                        string auxSelectPOBackUp = $"SELECT id FROM `purchaseOrderBackup` WHERE publicationDate BETWEEN '{firstDayYearS}' and '{middleTimeS}'";

                        string sqlPOAditionalDataBackup1 = $"SELECT * FROM `purchaseOrderAdditionalDataBackup` WHERE bidNumber in ({auxSelectPOBackUp})";
                        POAditionalDataBackup1 = crud.Select( sqlPOAditionalDataBackup1, "costa_rica_bids_db");

                        string sqlEvaluationsBackup1 = $"SELECT * FROM `evaluationsBackup` WHERE bidNumber in ({auxSelectPOBackUp})";
                        evaluationsBackup1 = crud.Select( sqlEvaluationsBackup1, "costa_rica_bids_db");

                        string sqlProductsBackup1 = $"SELECT * FROM `productsBackup` WHERE bidNumber in ({auxSelectPOBackUp})";
                        productsBackup1 = crud.Select( sqlProductsBackup1, "costa_rica_bids_db");
                        #endregion

                        #region Segundo extracto de Data de Backup
                        string sqlPurchaseOrdersBackup2 = $"SELECT * FROM `purchaseOrderBackup` WHERE publicationDate BETWEEN '{middleTimeF}' and '{currentDayS}'";
                        DataTable purchaseOrdersBackup2 = crud.Select( sqlPurchaseOrdersBackup2, "costa_rica_bids_db");

                        string auxSelectPOBackUp2 = $"SELECT id FROM `purchaseOrderBackup` WHERE publicationDate BETWEEN '{middleTimeF}' and '{currentDayS}'";

                        string sqlPOAditionalDataBackup2 = $"SELECT * FROM `purchaseOrderAdditionalDataBackup` WHERE bidNumber in ({auxSelectPOBackUp2})";
                        DataTable POAditionalDataBackup2 = crud.Select( sqlPOAditionalDataBackup2, "costa_rica_bids_db");

                        string sqlEvaluationsBackup2 = $"SELECT * FROM `evaluationsBackup` WHERE bidNumber in ({auxSelectPOBackUp2})";
                        DataTable evaluationsBackup2 = crud.Select( sqlEvaluationsBackup2, "costa_rica_bids_db");

                        string sqlProductsBackup2 = $"SELECT * FROM `productsBackup` WHERE bidNumber in ({auxSelectPOBackUp2})";
                        DataTable productsBackup2 = crud.Select( sqlProductsBackup2, "costa_rica_bids_db");

                        #endregion

                        //Merge de backups
                        purchaseOrdersBackup1.Merge(purchaseOrdersBackup2);
                        POAditionalDataBackup1.Merge(POAditionalDataBackup2);
                        evaluationsBackup1.Merge(evaluationsBackup2);
                        productsBackup1.Merge(productsBackup2);
                    }

                #endregion
            
                
                //Merge de Tablas principales con backups.
                purchaseOrders.Merge(purchaseOrdersBackup1);
                POAditionalData.Merge(POAditionalDataBackup1);
                evaluations.Merge(evaluationsBackup1);
                products.Merge(productsBackup1);

            #endregion

            #region  Extracción adicional de datos de BD.
            console.WriteLine("Extracción datos adicionales");

            string sqlEntities = "SELECT * FROM `institutions`";
            DataTable entities = crud.Select( sqlEntities, "costa_rica_bids_db");
            string sqlEmployees = "SELECT * FROM `digital_sign`";
            DataTable employees = crud.Select(sqlEmployees, "MIS");
            string sqlTitles = "SELECT * FROM `sicopFields`";
            DataTable titles = crud.Select( sqlTitles, "costa_rica_bids_db");
            string sqlValueTeam = "SELECT * FROM `valueTeam`";
            DataTable valueTeam = crud.Select( sqlValueTeam, "costa_rica_bids_db");
            string sqlProcessType = "SELECT * FROM `processType`";
            DataTable processType = crud.Select( sqlProcessType, "costa_rica_bids_db");


            #endregion


            //DataTable clave donde se almacenarán todas las purchaseOrder al excel. 
            DataTable bidReport = new DataTable();

            #region Agregar columnas al DataTable con Titulos establecidos de la tabla camposSicop. 

            titles.Rows.Cast<DataRow>().ToList().ForEach(dataRow =>
            {
                string val = dataRow["csicop"].ToString();
                if (val != "[ 11. Información de bien, servicio u obra ]" && val != "Sistema de Evaluación de Ofertas")
                {
                    bidReport.Columns.Add(val);
                }
            });

            #endregion

            #region Ciclo para recorrer product por product e insertarlo a una fila con su información adicional al datatable bidReport.

            console.WriteLine("Recorriendo cada purchase order...");

            //Recorre cada PurchaseOrder extraída del rango de tiempo.
            for (int i = 0; i < purchaseOrders.Rows.Count; i++)
            {
                
                string idPurchaseOrder = purchaseOrders.Rows[i]["id"].ToString();
                console.WriteLine("#" + (i + 1) + " de #" + purchaseOrders.Rows.Count + ". id PO: " + idPurchaseOrder +".");


                DataTable ProductsByPO;
                #region Crear un DataTable auxiliar (ProductsByPO), únicamente con los products relacionados al purchaseOrder actual.
                ProductsByPO = products.Clone(); //Copiar las columnas del Datable Products original.           
                DataRow[] drAuxProducts = products.Select($"bidNumber={idPurchaseOrder}");
                ProductsByPO = drAuxProducts.CopyToDataTable();
                #endregion

                string resp = "Se agrega todos los datos al reporte de excel del número de licitación: " + idPurchaseOrder;
                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Agregar fila reporte ", resp, root.Subject);
                respFinal = respFinal + "\\n" + "Agregar fila reporte: " + resp;


                //Recorre los ProductsByPO para insertar una fila de excel por cada uno, junto a la demás información extraída.
                ProductsByPO.Rows.Cast<DataRow>().ToList().ForEach(productByPO =>
                {
                    
                    DataRow bidRow = bidReport.Rows.Add();

                    DataTable EvaluationsByPO;
                    #region Crear un DataTable auxiliar (EvaluationsByPO), únicamente con los Evaluations relacionados al purchaseOrder actual.
                    EvaluationsByPO = evaluations.Clone(); //Copiar las columnas del Datable Evaluations original.
                    try
                    { //Existen purchaseOrder que no tienen evaluations.
                        DataRow[] drAuxEvaluations = evaluations.Select($"bidNumber={idPurchaseOrder}");
                        EvaluationsByPO = drAuxEvaluations.CopyToDataTable();                        
                    }
                    catch (Exception) { }
                    #endregion

                    DataTable POAdditionalDataByPO;
                    #region Crear un DataTable auxiliar (POAditionalData), únicamente con los PurchaseOrderAdditionalData relacionados al purchaseOrder actual.
                    POAdditionalDataByPO = POAditionalData.Clone(); //Copiar las columnas del Datable PurchaseOrder original.           
                    DataRow[] drAuxPOAditionalData = POAditionalData.Select($"bidNumber={idPurchaseOrder}");
                    POAdditionalDataByPO = drAuxPOAditionalData.CopyToDataTable();
                    #endregion


                    //Crea un String de los campos de Evaluación para cada fila del product, junta las diferentes evaluaciones en una celda de cada campo.
                    EvaluationsByPO.Columns.Cast<DataColumn>().ToList().ForEach(evaColumn =>
                    {
                        string eval = "";
                        string nameColumnSicopField = "";
                        try
                        {
                            DataRow[] auxDrColumnSicopField = titles.Select($"poColumn='{evaColumn}'");
                            nameColumnSicopField = auxDrColumnSicopField[0]["csicop"].ToString();
                        }
                        catch (Exception) { };

                        //Validación para verificar si la columna de de evaluations actual, está en sicopFields,
                        //el cual es la plantilla de columnas para el reporte, si no está lo ignora.
                        if (!string.IsNullOrEmpty(nameColumnSicopField) && nameColumnSicopField != "Número de procedimiento")
                        {
                            EvaluationsByPO.Rows.Cast<DataRow>().ToList().ForEach(evaRow =>
                            {
                                eval += evaRow[evaColumn].ToString() + Environment.NewLine;
                            });
                            bidRow[nameColumnSicopField] = eval;
                        }
                    });

                    //Llenar los campos de products en el bidRow auxiliar.
                    ProductsByPO.Columns.Cast<DataColumn>().ToList().ForEach(ProductColumn =>
                    {
                        DataRow[] auxDrColumnSicopField = titles.Select($"poColumn='{ProductColumn}'");
                        string nameColumnSicopField = "";
                        try//Hay veces que consulta un campo en sicopFields que no existe ahí.
                        {
                            nameColumnSicopField = auxDrColumnSicopField[0]["csicop"].ToString();
                        }
                        catch (Exception) { };

                        //Validación para verificar si la columna de de evaluations actual, está en sicopFields,
                        //el cual es la plantilla de columnas para el reporte, si no está lo ignora.
                        if (!string.IsNullOrEmpty(nameColumnSicopField) && nameColumnSicopField != "Número de procedimiento")
                        {
                            bidRow[nameColumnSicopField] = productByPO[ProductColumn];
                        }
                    });

                    //Llenar los campos de PurchaseOrderAdditionalData en el bidRow auxiliar.
                    POAdditionalDataByPO.Columns.Cast<DataColumn>().ToList().ForEach(POAdditionalColumn =>
                    {
                        DataRow[] auxDrColumnSicopField = titles.Select($"poColumn='{POAdditionalColumn}'");
                        string nameColumnSicopField = "";
                        try //Hay veces que consulta un campo en sicopFields que no existe ahí.
                        {
                            nameColumnSicopField = auxDrColumnSicopField[0]["csicop"].ToString();
                        }
                        catch (Exception) { };

                        //Validación para verificar si la columna de de POAdditionalData actual, está en sicopFields,
                        //el cual es la plantilla de columnas para el reporte, si no está lo ignora.
                        if (!string.IsNullOrEmpty(nameColumnSicopField) && nameColumnSicopField != "Número de procedimiento")
                        {
                            if (POAdditionalColumn.ToString() == "valueTeam") //En caso que la columna sea de ValueTeam cambiar el formato de 1 al nombre para fines estéticos.
                            {
                                bidRow[nameColumnSicopField] =
                                valueTeam.Select($"id={POAdditionalDataByPO.Rows[0][POAdditionalColumn]}")[0]["valueTeam"].ToString();
                            }
                            else if (POAdditionalColumn.ToString() == "gbmStatus")//En caso que la columna sea de GBMStatus cambiar el formato de 1 al nombre para fines estéticos.
                            {
                                string value = (POAdditionalColumn.ToString() == "1") ? "Si" : "No";
                                bidRow[nameColumnSicopField] = value;
                            }
                            else//En este caso es una posición normal de POAdditionalData
                            {
                                //En la posición 0 porque siempre hay solo una POAdditionalData siempre.
                                bidRow[nameColumnSicopField] = POAdditionalDataByPO.Rows[0][POAdditionalColumn];
                            }
                        }
                    });

                    //Llenar los campos de PurchaseOrder final en el bidRow auxiliar.
                    purchaseOrders.Columns.Cast<DataColumn>().ToList().ForEach(PurchaseOrdersColumn =>
                    {
                        DataRow[] auxDrColumnSicopField = titles.Select($"poColumn='{PurchaseOrdersColumn}'");
                        string nameColumnSicopField = "";
                        try//Hay veces que consulta un campo en sicopFields que no existe ahí.
                        {
                            nameColumnSicopField = auxDrColumnSicopField[0]["csicop"].ToString();
                        }
                        catch (Exception) { };

                        //Validación para verificar si la columna de de PurchaseOrder actual, está en sicopFields,
                        //el cual es la plantilla de columnas para el reporte, si no está lo ignora.
                        if (!string.IsNullOrEmpty(nameColumnSicopField))
                        {
                            if (PurchaseOrdersColumn.ToString() == "processType") //En caso que la columna sea de ProcessType cambiar el formato de 1 al nombre para fines estéticos.
                            {
                                bidRow[nameColumnSicopField] =
                                processType.Select($"id={purchaseOrders.Rows[i][PurchaseOrdersColumn]}")[0]["processType"].ToString();
                            }
                            else//En este caso es una posición normal de PurchaseOrders.
                            {
                                bidRow[nameColumnSicopField] = purchaseOrders.Rows[i][PurchaseOrdersColumn];
                            }
                        }
                    });
                    bidReport.AcceptChanges();
                });
            }

            bidReport.Columns.Remove("Detalle de partida");
            bidReport.Columns.Remove("Detalle de linea");
            #endregion

            console.WriteLine("Guardar el excel");
            #region Guardar el Excel
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(bidReport, "Concursos SICOP");
            string fecha_file = $"{DateTime.Now.Day.ToString().PadLeft(2, '0')}_{DateTime.Now.Month.ToString().PadLeft(2, '0')}_{DateTime.Now.Year}";
            string ruta = root.FilesDownloadPath + $"\\Reporte Concursos SICOP {fecha_file}.xlsx";
            wb.Worksheet("Concursos SICOP").Columns().AdjustToContents(); //Es exacto a la función autofit del excel antiguo 
            wb.SaveAs(ruta);
            #endregion

            console.WriteLine("Enviar el email");
            #region Envío email 

            string[] cc = { "" };
            string[] adj = { ruta };
            string sub = "Concursos SICOP - Licitaciones de Costa Rica - " + fecha_file.Replace("_", "/");
            string msj = "A continuación, se adjunta el archivo de las licitaciones públicas del portal SICOP a la fecha " + DateTime.Today.ToString();
            //DMO
            string dmoemail = "";
            try
            {
                //DataTable DMO = Lcsql.SelectRow("licitaciones_cr", "SELECT * FROM `email_address` WHERE `CATEGORIA` = 'DMOLIDER'");
                DataTable DMO = crud.Select("SELECT * FROM `emailAddress` WHERE `category` = 'DMOLIDER'", "costa_rica_bids_db");
                JObject row = JObject.Parse(DMO.Rows[0]["jemail"].ToString().Trim());
                dmoemail = row["email"].Value<string>();
            }
            catch (Exception)
            {
                dmoemail = "dmeza@gbm.net";
            }
            string html = Properties.Resources.emailLpCr;
            html = html.Replace("{subject}", "Concursos SICOP YTD");
            html = html.Replace("{cuerpo}", msj);
            html = html.Replace("{contenido}", "");

            mail.SendHTMLMail(html, new string[] { dmoemail }, sub, null, adj);

            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Envio de reporte", "Se envia reporte por email", root.Subject);
            respFinal = respFinal + "\\n" + "Se envia reporte por email";


            root.BDUserCreatedBy = dmoemail;
            root.requestDetails = respFinal;



            #endregion

        }


    }
}
