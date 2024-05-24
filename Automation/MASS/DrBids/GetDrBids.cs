using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.DrBids;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Projects.DrBids;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.MASS.DrBids
{
    /// <summary>
    /// Clase MASS Automation encargada de extraer licitaciones del portal web de compras en República Dominicana.
    /// </summary>
    class GetDrBids
    {
        #region variables_globales
        Stats estadisticas = new Stats();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        Log log = new Log();
        MailInteraction mail = new MailInteraction();
        
        ProcessInteraction proc = new ProcessInteraction();
        BidsGbCrSql lcsql = new BidsGbCrSql();
        Database db2 = new Database();
        ProcessAdmin padmin = new ProcessAdmin();
        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
        BidsGbCrSql liccr = new BidsGbCrSql();
        WebexTeams wt = new WebexTeams();
        WebInteraction webInteraction = new WebInteraction();
        BidsGbDrSql Dr_Sql = new BidsGbDrSql();
        ProcessInteraction proccess = new ProcessInteraction();
        Rooting rooting = new Rooting();
        GenerateExcelDr informeExcelDr = new GenerateExcelDr();
        ProcessAdmin process_Admin = new ProcessAdmin();
        BidsFilter filtrar = new BidsFilter();
        ValidateData val = new ValidateData();

        string respFinal = "";

        #endregion

        public void Main()
        {
            console.WriteLine("Procesando...");
            ProcessDr();

            using (Stats stats = new Stats())
            {
                stats.CreateStat();
            }

        }
        /// <summary>
        /// Método encargado de ingresar a la página de compras de república dominicana, extraer las licitaciones junto a su información necesaria para extraerla y almacenarlas.
        /// </summary>
        private void ProcessDr()
        {

            //Cerrar el proceso de Google Chrome
            proccess.KillProcess("chromedriver", true);
            proccess.KillProcess("chrome", true);

            //Se hace la coneccion y se accesa a la pagina de licitaciones Dominicana
            IWebDriver chrome = webInteraction.NewSeleniumChromeDriver(root.borrarArchivo);
            chrome.Navigate().GoToUrl("https://comunidad.comprasdominicana.gob.do/Public/Tendering/ContractNoticeManagement/Index");
            try
            {
                if (chrome.FindElement(By.XPath("//*[@id='frmMainForm_tblContainer_trRow1_tdCell2_tblTable']")).Displayed)
                {
                    chrome.Navigate().GoToUrl("https://comunidad.comprasdominicana.gob.do/Public/Tendering/ContractNoticeManagement/Index");
                }
            }
            catch { }

            //Se procede a hacer una busqueda avanzada y se le manda el rango de fechas por el cual va a filtrar
            chrome.FindElement(By.XPath("//*[@id='lnkAdvancedSearch']")).SendKeys(Keys.Return);
            chrome.FindElement(By.XPath("//*[@id='dtmbOfficialPublishDateFrom_txt']")).SendKeys(DateTime.Now.ToString("dd/MM/yyyy 00:00"));
            chrome.FindElement(By.XPath("//*[@id='dtmbOfficialPublishDateTo_txt']")).SendKeys(DateTime.Now.ToString("dd/MM/yyyy 23:59"));
            SelectElement selectElement = new SelectElement(chrome.FindElement(By.XPath("//*[@id='selRequestStatus']")));
            selectElement.SelectByValue("3");
            chrome.FindElement(By.XPath("//*[@id='btnSearchButton']")).SendKeys(Keys.Return);
            Thread.Sleep(5000);
            try
            {
                string noData = chrome.FindElement(By.ClassName("VortalGridEmptyResult")).Text;
                chrome.Close();
                proccess.KillProcess("chromedriver", true);
                console.WriteLine("No hay licitaciones");
                return;
            }
            catch (Exception)
            { }

            WaitLoadingPagWeb(30000, chrome);

            //Empezar a tocar el boton ver mas para desplegar la tabla por completo
            try
            {
                IWebElement verMas = chrome.FindElement(By.XPath("//*[@id='tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_Paginator_goToPage_MoreItems']"));

                while (verMas.Displayed == true)
                {
                    verMas.SendKeys(Keys.Return);
                    Thread.Sleep(1000);
                    verMas = chrome.FindElement(By.XPath("//*[@id='tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_Paginator_goToPage_MoreItems']"));
                }
            }
            catch { }
            //Tomar las filas de la tabla de licitaciones
            IWebElement tablaLicitaciones = chrome.FindElement(By.XPath("//*[@id='tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tbl']"));
            IList<IWebElement> Filas = tablaLicitaciones.FindElements(By.TagName("tr"));
            int cont = 0;
            List<GeneralData> addExcel = new List<GeneralData>();
            //Empesar a recorer las filas de la tabla para tomar los datos de las columnas
            string nombre_archivo = "";
            List<string> listaRepetidos = Dr_Sql.TraerIdBids();
            //REcorre todas las filas de la tabla de licitaciones
            foreach (IWebElement item in Filas.Skip(1))
            {
                List<string> nombresArchivos = new List<string>();
                //Toma los datos Generales para almacenarlos y convertirlo en un JSON
                string cliente = chrome.FindElement(By.Id("tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tdAuthorityNameCol_spnMatchingResultAuthorityName_" + cont)).Text;
                string referencia = chrome.FindElement(By.Id("tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tdUniqueIdentifierCol_spnMatchingResultReference_" + cont)).Text;
                string descripcion = chrome.FindElement(By.Id("tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tdDescriptionCol_spnMatchingResultDescription_" + cont)).Text.Replace(@"""", "").Replace("''", "");
                string fechaPubli = (chrome.FindElement(By.Id("dtmbNationalOfficialPublishingDate_" + cont + "_txt")).Text.Replace(" (UTC -4 hours)", ""));
                DateTime.TryParse(fechaPubli, out DateTime fechaPublicacion);
                string presupuesto = chrome.FindElement(By.Id("cbxBasePriceValue_" + cont)).Text.Replace("Dominican Pesos", "");
                Thread.Sleep(1000);
                //Valida si el registro ya esta en la base de datos
                console.WriteLine("Leyendo licitacion número: " + cont + " de " + (Filas.Count - 3) + " - " + referencia);
                if (Dr_Sql.ValidateRepetidos(listaRepetidos, referencia) == false)
                {
                    string resp = $"Cliente: {cliente} - Referencia: {referencia} - Descripción: {descripcion} - Fecha de publicación: {fechaPubli} - Presupuesto: {presupuesto}";
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Extraer Licitación", resp, root.Subject);
                    respFinal = respFinal + "\\n" + "Extraer licitación: " + resp;


                    //Abre el modal y se cambia al elemento para obtener informacion que esa contenida ahi
                    chrome.FindElement(By.Id("tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tdDetailColumn_lnkDetailLink_" + cont)).SendKeys(Keys.Enter);
                    chrome.SwitchTo().Frame("OpportunityDetailModal_iframe");
                    chrome.FindElement(By.XPath("//*[@id='fdsRequestSummaryInfo_tblDetail']"));
                    string tipoProceso = chrome.FindElement(By.XPath("//*[@id='fdsRequestSummaryInfo_tblDetail_trRowProcedureType']")).Text;
                    IWebElement tablaScheduling = chrome.FindElement(By.XPath("//*[@id='fdsSchedulingP2Gen_tblDetail']"));
                    IList<IWebElement> FilasScheduling = tablaScheduling.FindElements(By.TagName("tr"));
                    List<string> listaRegistros = new List<string>();
                    List<string> key = new List<string>();

                    //Obtener datos de tiempo de la tabla de cronogramas, almacenandolo en un diccionario para luego convertirlo a un JSON
                    Dictionary<string, DateTime> diccionariotemp = new Dictionary<string, DateTime>();
                    foreach (IWebElement scheduling in FilasScheduling.Where(scheduling => scheduling.Displayed))
                    {
                        IList<IWebElement> columnasScheduling;
                        columnasScheduling = scheduling.FindElements(By.TagName("td"));
                        try
                        {
                            string[] valorCronograma = columnasScheduling[1].Text.Split('(');
                            DateTime.TryParse(valorCronograma[1], out DateTime fechaCronograma);
                            diccionariotemp.Add(val.RemoveAccents(columnasScheduling[0].Text), fechaCronograma);
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    string Json_scheduling = JsonConvert.SerializeObject(diccionariotemp);

                    List<string> ListaJsonsArticulos = new List<string>();
                    List<Article> listaArticulos = new List<Article>();
                    Article nuevo_articulo = null;
                    IList<IWebElement> tablaArticulos = chrome.FindElements(By.ClassName("PriceListLineTable"));
                    int isValid = 0;
                    int contReferencias = 1;

                    //Recorre los articulos de cada licitacion
                    foreach (IWebElement findTr in tablaArticulos.Where(findTr => findTr.Displayed))
                    {
                        IList<IWebElement> filasArticulos = findTr.FindElements(By.TagName("tr"));

                        foreach (IWebElement articulo in filasArticulos.Where(articulo => articulo.Displayed))
                        {

                            IList<IWebElement> columnasArticulos;
                            columnasArticulos = articulo.FindElements(By.TagName("td"));
                            string Codigo = columnasArticulos[0].Text;
                            if (Codigo == "")
                            {
                                Codigo = Convert.ToString(contReferencias);
                            }
                            string descripcion_Articulo = columnasArticulos[4].Text;
                            string cantidad = columnasArticulos[5].Text;
                            string unidad = columnasArticulos[6].Text;
                            string precio_Unitario = columnasArticulos[7].Text;
                            string Precio_total = columnasArticulos[8].Text;
                            if (filtrar.KeyMatch(descripcion_Articulo, filtrar.KeyWord()) == "SI")
                            {
                                isValid++;
                            }
                            //Generar JSON de las columnas de Articulos
                            nuevo_articulo = new Article()
                            {
                                Codigo = Codigo,
                                descripcion_Articulo = descripcion_Articulo,
                                cantidad = cantidad,
                                unidad = unidad,
                                precio_Unitario = precio_Unitario,
                                Precio_total = Precio_total
                            };
                            listaArticulos.Add(nuevo_articulo);
                            string json_articulo = JsonConvert.SerializeObject(nuevo_articulo);
                            ListaJsonsArticulos.Add(json_articulo);
                            contReferencias++;
                            break;
                        }
                    }
                    string json_Lista_Articulos = JsonConvert.SerializeObject(ListaJsonsArticulos);

                    //crear array con nombre de archivo y extension y pasarlo a un string 
                    IWebElement tablaArchivosDescarga = chrome.FindElement(By.XPath("//*[@id='grdGridDocumentList_tbl']"));
                    IList<IWebElement> filasArchivosDescarga = tablaArchivosDescarga.FindElements(By.TagName("tr"));
                    int indexArchivos = 0;
                    byte[] archivoZip = { };
                    filtrar.KeyWord();
                    string interesGBM = "";
                    int descargado = 0;
                    //Filtrar Licitaciones por palabras
                    if (filtrar.KeyMatch(descripcion, filtrar.KeyWord()) == "SI" || isValid != 0)
                    {
                        descargado = 1;
                        interesGBM = "SI";
                        //Crear y obtener ruta de la carpeta para almacenar archivos 
                        List<ArchivoBinario> archivoBinarios = new List<ArchivoBinario>();
                        foreach (IWebElement descarga in filasArchivosDescarga.Skip(1))
                        {
                            IList<IWebElement> columnasArchivosDescarga;
                            columnasArchivosDescarga = descarga.FindElements(By.TagName("td"));
                            nombre_archivo = columnasArchivosDescarga[0].Text;
                            int cantidadArchivo = 1;
                            if (nombresArchivos.Count == 0)
                            {
                                nombresArchivos.Add(nombre_archivo);
                            }
                            else
                            {
                                for (int i = 0; i < nombresArchivos.Count; i++)
                                {
                                    if (nombresArchivos[i] == columnasArchivosDescarga[0].Text)
                                    {
                                        int index = nombre_archivo.IndexOf(".");
                                        nombre_archivo = nombre_archivo.Insert(index, " (" + cantidadArchivo + ")");
                                        nombresArchivos.Add(nombre_archivo);
                                        cantidadArchivo++;
                                    }
                                    else if (i == nombresArchivos.Count - 1)
                                    {
                                        nombresArchivos.Add(nombre_archivo);
                                        break;
                                    }
                                }
                            }

                            chrome.FindElement(By.XPath("//*[@id='lnkDownloadLinkP3Gen_" + indexArchivos + "']")).SendKeys(Keys.Enter);
                            Thread.Sleep(1000);

                            // esperar que descargue por completo el archivo
                            int con = 0;
                            while (File.Exists(rooting.borrarArchivo + "\\" + nombre_archivo) == false)
                            {
                                Thread.Sleep(1000);
                                con++;
                                if (con > 60) { break; }
                            }
                            string rutaArchivo = Path.GetFullPath(rooting.borrarArchivo + "\\" + nombre_archivo);
                            using (BinaryFiles binaryFiles = new BinaryFiles())
                            {
                                if (archivoBinarios.Count == 0)
                                {
                                    archivoBinarios.Add(binaryFiles.Convert(rutaArchivo, nombre_archivo));
                                }
                                else
                                {
                                    string temp = nombresArchivos[nombresArchivos.Count - 1];
                                    if (temp != nombre_archivo)
                                    {
                                        archivoBinarios.Add(binaryFiles.Convert(rutaArchivo, nombre_archivo));
                                    }
                                }
                            }
                            indexArchivos++;
                        }
                        if (archivoBinarios.Count > 0)
                        {
                            using (BinaryFiles binary = new BinaryFiles())
                            {
                                archivoZip = binary.CrearBytesZIP(archivoBinarios);
                            }
                        }
                    }

                    else
                    {
                        interesGBM = "NO";
                    }
                    //Generar Objeto Datos Generales
                    Thread.Sleep(1000);
                    GeneralData datos_Generales = new GeneralData()
                    {
                        cliente = cliente,
                        referencia = referencia,
                        descripcion = val.RemoveAccents(descripcion),
                        fechaPublicacion = fechaPublicacion,
                        presupuesto = presupuesto,
                        tipoProceso = tipoProceso,
                        interesGBM = interesGBM
                    };
                    SapData datoSAP = new SapData() { };

                    string json_datos = JsonConvert.SerializeObject(datos_Generales);
                    chrome.FindElement(By.XPath("//*[@id='tbToolBarPlaceholder_btnClose']")).SendKeys(Keys.Enter);
                    chrome.SwitchTo().Window(chrome.WindowHandles[0]);
                    Thread.Sleep(1500);
                    //Generar Objeto Licitaciones
                    Bids licitaciones = new Bids
                    {
                        DG = datos_Generales,
                        DS = datoSAP,
                        AT = json_Lista_Articulos,
                        PL = Json_scheduling,
                        AJ = archivoZip
                    };
                    //Insertar en la base de datos
                    Dr_Sql.InsertRow(licitaciones);
                    if (descargado == 1)
                    {
                        Dr_Sql.InsertFiles(licitaciones, referencia + ".zip");
                    }
                    datos_Generales.cronograma = diccionariotemp;
                    datos_Generales.listaArticulos = listaArticulos;
                    addExcel.Add(datos_Generales);
                    process_Admin.DeleteFiles(rooting.borrarArchivo);
                }
                cont++;
                if (Filas.Count - 3 == cont)
                {
                    break;
                }
            }
            console.WriteLine("Creando Exceles...");
            informeExcelDr.FileExcel(informeExcelDr.GenerateSellers(addExcel));
            process_Admin.DeleteFiles(rooting.ExcelDr[0]);
            chrome.Close();
            //Cerrar el proceso de Google Chrome
            proccess.KillProcess("chromedriver", true);

            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "RGUERRERO";

        }
        /// <summary>
        /// Método destinado a esperar que cargue la página web si da un error de carga.
        /// </summary>
        /// <param name="waitTime">Tiempo de espera.</param>
        /// <param name="chromeDriverInstance">Instancia de ChromeDrive.r</param>
        void WaitLoadingPagWeb(int waitTime, IWebDriver chromeDriverInstance)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            while (chromeDriverInstance.FindElements(By.ClassName("loadingCursor")).Count > 0)
            {
                System.Threading.Thread.Sleep(50);
                if (sw.ElapsedMilliseconds > waitTime) throw new TimeoutException();
            }
        }
        #region Objetos
        public class Article
        {
            public string Codigo { get; set; }
            public string descripcion_Articulo { get; set; }
            public string cantidad { get; set; }
            public string unidad { get; set; }
            public string precio_Unitario { get; set; }
            public string Precio_total { get; set; }
        }
        public class GeneralData
        {
            public string cliente { get; set; }
            public string referencia { get; set; }
            public string descripcion { get; set; }
            public DateTime fechaPublicacion { get; set; }
            public string presupuesto { get; set; }
            public string tipoProceso { get; set; }
            public List<Article> listaArticulos { get; set; }
            public string interesGBM { get; set; }
            public Dictionary<string, DateTime> cronograma { get; set; }
        }
        public class SapData
        {
            public string participa { get; set; }
            public int opp { get; set; }
            public int quote { get; set; }
            public int salesOrder { get; set; }
            public string salesTeam { get; set; }
            public string tipoOpp { get; set; }
            public string salesType { get; set; }
        }
        public class Bids
        {
            public GeneralData DG { get; set; }
            public SapData DS { get; set; }
            public string AT { get; set; }

            public string PL { get; set; }

            public byte[] AJ { get; set; }
        }
        #endregion
    }
}
