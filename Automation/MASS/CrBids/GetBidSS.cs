using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Process;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.CrBids;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

namespace DataBotV5.Automation.MASS.CrBids
{
    /// <summary>
    /// Clase RPA "Robot 1" Automation encargada de guardar
    /// la información de las licitaciones SICOP de CR y notificar a los AM.
    /// </summary>
    class GetBidSS
    {
        BidsGbCrSql liccr = new BidsGbCrSql();
        ProcessAdmin padmin = new ProcessAdmin();
        CrBidsLogical crBids = new CrBidsLogical();
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        WebInteraction web = new WebInteraction();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        WebexTeams wt = new WebexTeams();
        CRUD crud = new CRUD();
        internal CrBidsLogical bidLogical { get => crBids; set => crBids = value; }
        string mandante = "QAS";

        string respFinal = "";

        public void Main()
        {

            console.WriteLine("Procesando...");
            string cantFilas = ExtractInfo(bidLogical.SelConn("https://www.sicop.go.cr/index.jsp"));

            if (cantFilas != "0")
            {

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }

        }
        /// <summary>
        /// Extrae la información de los concursos publicados el dia actual y envia las notificaciones correspondientes
        /// </summary>
        /// <param name="chrome">Webdriver de la página de SICOP</param>
        private string ExtractInfo(IWebDriver chrome)
        {
            #region variables privadas
            IDictionary<string, string> lista_am = new Dictionary<string, string>();
            IDictionary<string, string> lista_g = new Dictionary<string, string>();
            bool validar_lineas = true;
            string bidNumberFromWebPage = "";
            string tabla_correo_error = "";
            string lista_gerente_general = "";
            byte[] zip = null;
            //titulos de la tabla para enviar correo
            string tabla_correo_th = "<table class='myCustomTable' width='100 %'>" +
                       "<tr>" +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Concurso         </b></p></td>	   " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Descripcion      </b></p></td>	   " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Institucion      </b></p></td>	   " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Publicacion      </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Aclaraciones     </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Apertura         </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Presupuesto      </b></p></td></tr> ";
            //titulos de la tabla para enviar correo para DMO
            string tabla_correog_th = "<table class='myCustomTable' width='100 %'>" +
                       "<tr>" +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Concurso         </b></p></td>	   " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Descripcion      </b></p></td>	   " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Institucion      </b></p></td>	   " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Publicacion      </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Aclaraciones     </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Apertura         </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Presupuesto      </b></p></td>      " +
                       "<td style='padding:0.75pt;'><p align='center' style='font-size:11pt;font-family:Calibri,sans-serif;text-align:center;margin:0;'><b>Account Manager  </b></p></td></tr> ";
            #endregion
            #region Borrar las descargas anteriores
            try
            {
                Directory.Delete(root.downloadfolder, true);
                Directory.CreateDirectory(root.downloadfolder);
            }
            catch (Exception) { }
            #endregion
            #region Tomar info necesaria de la DB

            DataTable bidIdsPurchaseOrder = crud.Select( "SELECT bidNumber FROM purchaseOrder", "costa_rica_bids_db");//tomar los concursos actuales
            DataTable bidIdspurchaseOrderBackup = crud.Select( "SELECT bidNumber FROM purchaseOrderBackup", "costa_rica_bids_db");//tomar los concursos actuales
            DataTable institutions = crud.Select( "SELECT * FROM institutions", "costa_rica_bids_db");//tomar los encargados de las entidades
            DataTable empleados = crud.Select("SELECT * FROM `digital_sign`", "MIS"); //tabla de empleados, para busacar los AM
            DataTable customers = crud.Select("SELECT idClient, name FROM `clients`", "databot_db");
            //DataTable full_columns = liccr.select_row("licitaciones_cr", "SHOW FULL COLUMNS FROM concursos"); //tabla campos/comments

            DataTable sicopFields = crud.Select( "SELECT * FROM sicopFields WHERE active = 1", "costa_rica_bids_db"); //tabla campos/comments
            DataTable idjsonSicopFields = crud.Select( "SELECT DISTINCT idjson FROM sicopFields WHERE active = 1 GROUP BY idjson", "costa_rica_bids_db"); //tabla campos/comments
            DataTable emaildefault = crud.Select( "SELECT * FROM emailAddress", "costa_rica_bids_db");
            DataTable processTypeData = crud.Select( "SELECT * FROM `processType`", "costa_rica_bids_db");
            List<string> words = liccr.KeyWord(mandante);
            #endregion


            #region buscador avanzado
            IWebElement topFrame = chrome.FindElement(By.XPath("//*[@id='topFrame']"));
            chrome.SwitchTo().Frame(topFrame);
            System.Threading.Thread.Sleep(1000);

            //click en el tab concursos
            chrome.FindElement(By.XPath("/html/body/div[2]/div/div/div[4]/ul/li[2]/div[1]/a[3]")).Click();

            #region regresa a la pagina/frame principal
            chrome.SwitchTo().ParentFrame();
            chrome.SwitchTo().Window(chrome.WindowHandles[0]);
            chrome.SwitchTo().DefaultContent();
            #endregion

            #region entro al mainFrame de la pagina de concursos
            IWebElement mainFrame = chrome.FindElement(By.XPath("//*[@id='mainFrame']"));
            chrome.SwitchTo().Frame(mainFrame);
            System.Threading.Thread.Sleep(1000);
            #endregion
            #region primero entrar a mainFrame para luego entrar al rightFrame (que es el buscador)
            IWebElement frame = chrome.FindElement(By.XPath("//*[@id='rightFrame']"));
            chrome.SwitchTo().Frame(frame);
            System.Threading.Thread.Sleep(1000);
            #endregion
            //Aqui van los filtros de busqueda
            #region limpiar el campo desde de Rango de fechas de publicación
            chrome.FindElement(By.Id("regDtFrom")).Clear();
            if (root.planner[root.BDClass]["conta"][0] == "1") //la primera vez que se ejecuta este bot en el día
                                                               //llena con el día de ayer por ser la primera vez del día 
                chrome.FindElement(By.Id("regDtFrom")).SendKeys(DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"));
            else
                //llenar con el día de hoy
                chrome.FindElement(By.Id("regDtFrom")).SendKeys(DateTime.Now.ToString("dd/MM/yyyy"));
            #endregion
            #region limpiar el campo hasta de rango de fechas de publicacion
            chrome.FindElement(By.Id("regDtTo")).Clear();
            chrome.FindElement(By.Id("regDtTo")).SendKeys(DateTime.Now.ToString("dd/MM/yyyy"));
            #endregion

            #region seleccionar el Estado del concurso
            SelectElement registration_applicationId = new SelectElement(chrome.FindElement(By.Name("biddocRcvYn")));
            registration_applicationId.SelectByValue("N"); //TODOS
            #endregion

            chrome.FindElement(By.XPath("/html/body/div[1]/div/div[2]/p/span/a")).Click(); //consultar boton

            #endregion
            #region cacula la cantidad de licitaciones que hay.
            string cant_filas = "";
            try
            { cant_filas = chrome.FindElement(By.XPath("//*[@id='total']/span[1]")).Text; }
            catch (Exception)
            { cant_filas = "0"; }
            #endregion
            if (cant_filas != "0")
            {
                #region Leer un concurso

                #region calculando la cantidad de paginas que hay
                int cf = int.Parse(cant_filas);
                double num1 = double.Parse(cant_filas), num2 = 10, filas = (num1 / num2), pageNumbers = Math.Ceiling(filas);
                if (pageNumbers == 0)
                { pageNumbers = 1; }
                int cont = 2;
                #endregion

                //for por cada pagina que hay
                for (int i = 1; i <= pageNumbers; i++)
                {
                    try
                    {

                        console.WriteLine("PAGINA #" + i + " DE " + pageNumbers);
                        cont++;
                        #region next en pagination
                        if (i != 1)
                        {
                            console.WriteLine("Siguiente pagina");
                            if (cont == 13)
                            {
                                IWebElement pagination_next = chrome.FindElement(By.ClassName("page02"));
                                pagination_next.Click();
                                cont = 3;
                            }
                            else
                            {
                                IWebElement pagination_next2 = chrome.FindElement(By.CssSelector("#paging > ul > li > :nth-child(" + cont + ")"));
                                pagination_next2.Click();
                            }

                        }
                        #endregion

                        IJavaScriptExecutor js = (IJavaScriptExecutor)chrome;
                        int e = 1;

                        #region Leer pagina actual

                        #region extrae la tabla de licitiaciones de la pagin actual
                        HtmlAgilityPack.HtmlDocument pagina_actual = new HtmlAgilityPack.HtmlDocument();
                        pagina_actual.LoadHtml(chrome.PageSource);
                        HtmlAgilityPack.HtmlNodeCollection eptable = pagina_actual.DocumentNode.SelectNodes("//*[@class='eptable']/tbody/tr"); //tomar toda la tabla
                        #endregion
                        //por cada fila de la tabla principal de licitaciones
                        for (int g = 1; g < eptable.Count; g++)
                        {
                            //los json de la BD, de acuerdo al id del comentario de la columna se escoge el json y se va agregando la info
                            IDictionary<string, JObject> jaisons = new Dictionary<string, JObject>();
                            //jainson va a guardar purchaseOrder, purcharseOrderAddData, products, evaluations
                            //idjsonSicopFields = sicopFields from DB
                            for (int w = 0; w < idjsonSicopFields.Rows.Count; w++)
                            {
                                jaisons[idjsonSicopFields.Rows[w]["idjson"].ToString()] = new JObject();
                            }

                            //estos 2 ya son json entonces es solo copiarlos
                            string bienes_servicios = "";
                            string evaluacion = "";

                            string columnas = "", valores = "", AM = "", gerente = "", cliente_institucion = "", contacto_institucion = "",
                                descripcion = "", publicacion = "", aclaraciones = "", apertura = "", presupuesto = "", valueTeam = "";

                            console.WriteLine("Leyendo registro #" + (e + ((i - 1) * (eptable.Count - 1))) + " de " + cf + "<br>");

                            string institution = eptable[g].SelectSingleNode("td[1]").LastChild.InnerText.Trim();
                            bidNumberFromWebPage = eptable[g].SelectSingleNode("td[1]").FirstChild.InnerText.Trim();
                            string concursoStatus = eptable[g].SelectSingleNode("td[5]").FirstChild.InnerText.Trim();

                            string interes = "";


                            //hacer una validacion primero para ver si la licitaciones esta o no en la BD IF
                            if (bidIdsPurchaseOrder.Select("bidNumber = '" + bidNumberFromWebPage + "'").Count() == 0 && bidIdspurchaseOrderBackup.Select("bidNumber = '" + bidNumberFromWebPage + "'").Count() == 0)
                            {
                                if (concursoStatus == "En recepción de ofertas" || concursoStatus == "Publicado")
                                {
                                    console.WriteLine(" > " + bidNumberFromWebPage + " No existe, extrayendo");
                                    //click en cada licitaciones, abrir en nueva pestaña y extraer la info
                                    try
                                    {
                                        //Click en el hypervinculo del concurso
                                        string link = eptable[g].SelectSingleNode("*//a[contains(@href, 'js_cartelSearch')]").GetAttributeValue("href", null).Replace("javascript:", "");
                                        js.ExecuteScript(link);

                                        #region Obtener datos de la licitacion en la pagina web

                                        #region variables para almacenar los datos
                                        DataTable factores_de_evaluacion = new DataTable();
                                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                        DataTable data_table = new DataTable();
                                        List<KeyValuePair<string, string>> datos = new List<KeyValuePair<string, string>>();//basicamente es un dic con keys repetidos
                                        List<KeyValuePair<string, string>> entrega = new List<KeyValuePair<string, string>>();//basicamente es un dic con keys repetidos

                                        DataTable licitacion = new DataTable();
                                        licitacion.Columns.Add("nombre_tabla");
                                        licitacion.Columns.Add("campo");
                                        licitacion.Columns.Add("value");
                                        licitacion.Columns.Add("data_table", typeof(DataTable));
                                        #endregion

                                        //extrae todos los titulos de los campos de la página SICOP
                                        ReadOnlyCollection<IWebElement> titulos = chrome.FindElements(By.CssSelector("body > div > div > div.cl_context > p.epsubtitle")); //titulos de la tablas
                                        //extrae todos las tablas de una licitación que hay en la página de SICOP
                                        ReadOnlyCollection<IWebElement> tablas = chrome.FindElements(By.CssSelector("body > div > div > div.cl_context > p.epsubtitle + table")); //las tablas

                                        List<string> adjuntos = new List<string>();
                                        string CartelNo = "", CartelSeq = "";
                                        bool fecha_en_producto = false;

                                        if (titulos.Count == tablas.Count)
                                        {
                                            IDictionary<string, string> tablas_eptdc = new Dictionary<string, string>();

                                            for (int q = 0; q < titulos.Count; q++)
                                            {
                                                string nombre_tabla = titulos[q].Text;
                                                //verificar count igual y si cambia el formato de la tabla

                                                doc.LoadHtml(tablas[q].GetAttribute("innerHTML"));
                                                HtmlAgilityPack.HtmlNodeCollection data_names2 = doc.DocumentNode.SelectNodes("//*[@class='epcthl']"); //porclase
                                                HtmlAgilityPack.HtmlNodeCollection data_values2 = doc.DocumentNode.SelectNodes("//*[@class='eptdl']"); //porclase

                                                if (data_names2 != null)
                                                {
                                                    #region extrae todos los campos con la clase epcthl y guarda cada uno de los valores de eptdl
                                                    for (int z = 0; z <= (data_names2.Count - 1); z++) //si son campos      
                                                    {
                                                        datos.Add(new KeyValuePair<string, string>(HttpUtility.HtmlDecode(data_names2[z].InnerText).Trim(), HttpUtility.HtmlDecode(data_values2[z].InnerText).Trim()));

                                                        DataRow row = licitacion.NewRow();
                                                        row["nombre_tabla"] = nombre_tabla;

                                                        string campo = HttpUtility.HtmlDecode(data_names2[z].InnerText).Trim();
                                                        row["campo"] = campo;


                                                        string valor = HttpUtility.HtmlDecode(data_values2[z].InnerText).Trim();
                                                        valor = valor.Replace(Convert.ToChar(160), ' '); //quitar caracter &nbsp;
                                                                                                         //valor = valor.Replace("\"", "").Replace("“", "").Replace("”", "");
                                                        valor = valor.Replace(Convert.ToChar(8220), ' ').Replace(Convert.ToChar(8221), ' ').Replace(Convert.ToChar(34), ' ');

                                                        row["value"] = valor;


                                                        licitacion.Rows.Add(row);

                                                    }
                                                    #endregion
                                                    #region Obtener el Numero SICOP y buscar FECHA DE RECEPCION si no la tiene aqui
                                                    if (nombre_tabla.Trim() == "[ 1. Información general ]")
                                                    {
                                                        CartelNo = doc.DocumentNode.SelectSingleNode("//*[@class='epreadc readonly fixCartelNo']").GetAttributeValue("value", null);
                                                        CartelSeq = doc.DocumentNode.SelectSingleNode("//*[@class='epreadc readonly fixCartelSeq']").GetAttributeValue("value", null);

                                                        if (doc.Text.Contains("La fecha y hora de la apertura del producto de la partida detallada podrá ser visualizada"))
                                                            fecha_en_producto = true;
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    #region leer las tablas
                                                    //hay secciones en la licitación que son tablas por ende no tiene elementos con la clase epcthl x lo que se descarga toda la tabla
                                                    //Nueva forma de tomar los adjuntos, ya que la numeracion puede cambiar por alguna razon
                                                    if (nombre_tabla.Trim() == "[ F. Documento del cartel ]")
                                                    {
                                                        HtmlAgilityPack.HtmlNodeCollection adjuntos_nodes = doc.DocumentNode.SelectNodes("//a[contains(@href, 'js_downloadFile')]");
                                                        if (adjuntos_nodes != null)
                                                        {
                                                            foreach (HtmlAgilityPack.HtmlNode adjunto in adjuntos_nodes)
                                                            {
                                                                string href = adjunto.GetAttributeValue("href", string.Empty);
                                                                adjuntos.Add(href.Replace("javascript:", ""));
                                                            }
                                                        }
                                                    }

                                                    data_table = web.TableToDatatable(tablas[q]);

                                                    DataRow row = licitacion.NewRow();
                                                    row["nombre_tabla"] = nombre_tabla;
                                                    row["value"] = "tabla";
                                                    row["data_table"] = data_table;
                                                    licitacion.Rows.Add(row);
                                                    #endregion
                                                }

                                            }
                                        }
                                        else//no son campos son tablas
                                        {
                                            //algo no cuadró
                                        }


                                        #region buscar si tiene ese mensaje raro sino buscar en la ventana del producto
                                        if (fecha_en_producto == true)
                                        {
                                            string temp = "";
                                            using (WebClient client = new WebClient())
                                            {
                                                client.Encoding = UTF8Encoding.UTF8;
                                                temp = client.DownloadString("https://www.sicop.go.cr/moduloBid/cartel/EP_CTJ_EXQ004.jsp" + "?cartelNo=" + CartelNo + "&cartelSeq=" + CartelSeq + "&cartelCate=1");
                                            }
                                            doc.LoadHtml(temp);
                                            string ss = sicopFields.Select("poColumn = 'receptionClosing'")[0]["csicop"].ToString();
                                            if (ss != "Cierre de recepción de ofertas")
                                            {
                                                ss = ss.Remove(0, 2);
                                            }

                                            temp = doc.DocumentNode.SelectSingleNode("//*[text() = '" + ss + "']/following-sibling::td").InnerText;
                                            temp = HttpUtility.HtmlDecode(temp).Trim();

                                            licitacion.Select("campo='" + ss + "'")[0]["value"] = temp;
                                            licitacion.AcceptChanges();

                                            string ao = sicopFields.Select("poColumn = 'offerOpening'")[0]["csicop"].ToString();
                                            if (ao != "Fecha/hora de apertura de ofertas")
                                            {
                                                ao = ao.Remove(0, 2); //elimina los 2 primeros caracteres
                                            }
                                            temp = doc.DocumentNode.SelectSingleNode("//*[text() = '" + ao + "']/following-sibling::td").InnerText;
                                            temp = HttpUtility.HtmlDecode(temp).Trim();

                                            licitacion.Select("campo='" + ao + "'")[0]["value"] = temp;
                                            licitacion.AcceptChanges();

                                        }
                                        #endregion
                                        #region FACTORES DE EVALUACION
                                        try
                                        {
                                            chrome.FindElement(By.XPath("//a[contains(@href, 'js_evalItemSearch')]")).Click();
                                            factores_de_evaluacion = web.TableToDatatable(chrome.FindElement(By.Id("fieldTable")));
                                            js.ExecuteScript("history.back();");
                                        }
                                        catch (Exception)
                                        { }
                                        #endregion
                                        #region PLAZO ENTREGA
                                        string html = "";
                                        using (WebClient client = new WebClient())
                                        {
                                            client.Encoding = UTF8Encoding.UTF8;
                                            html = client.DownloadString("https://www.sicop.go.cr/moduloBid/cartel/EP_CTJ_EXQ005.jsp" + "?cartelNo=" + CartelNo + "&cartelSeq=" + CartelSeq + "&cartelCate=" + "1" + "&cateSeqno=" + "1");
                                        }
                                        doc.LoadHtml(html);

                                        try
                                        {
                                            if (doc.DocumentNode.SelectNodes("//*[@id=\"delvDurId\"]/tr").Count() > 1)
                                            {
                                                entrega.Add(new KeyValuePair<string, string>("el_plazo_de_entrega", HttpUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//*[@id=\"delvDurId\"]/tr[2]/td[2]").InnerText).Trim()));
                                                entrega.Add(new KeyValuePair<string, string>("fecha_de_entrega", HttpUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//*[@id=\"delvDurId\"]/tr[2]/td[3]").InnerText).Trim()));
                                                entrega.Add(new KeyValuePair<string, string>("plazo_maximo_de_entrega", HttpUtility.HtmlDecode(doc.DocumentNode.SelectSingleNode("//*[@id=\"delvDurId\"]/tr[2]/td[4]").InnerText).Trim()));
                                            }
                                        }
                                        catch (ArgumentNullException) { }
                                        #endregion
                                        #endregion

                                        //JObject licit = JObject.Parse(JsonConvert.SerializeObject(licitacion));

                                        #region Obtener datos de la tabla de entidades
                                        //buscar el AM y si no lo encuentra Omar Mota
                                        //ingresar Gerente en licitacion mismo nombre de la columna de la BD
                                        //ingresar Value Team en licitacion mismo nombre de la columna de la BD
                                        string emailam = "";
                                        string nombream = "";
                                        DataRow[] insti = institutions.Select("institution = '" + institution + "'");
                                        if (insti.Count() == 1)
                                        {
                                            AM = insti[0]["salesRepresentative"].ToString();
                                            gerente = insti[0]["salesManager"].ToString();
                                            contacto_institucion = insti[0]["contactId"].ToString();
                                            cliente_institucion = insti[0]["customerId"].ToString();
                                            valueTeam = insti[0]["valueTeam"].ToString();
                                            DataRow[] amInfo = empleados.Select("user = '" + AM + "'");
                                            if (amInfo.Count() > 0)
                                            {
                                                nombream = amInfo[0]["name"].ToString();
                                            }
                                            emailam = $"{AM}@gbm.net";
                                        }
                                        else
                                        {
                                            //tomar valores por default
                                            //AM = "AA60000135";  //Omar Mota
                                            DataRow[] amDefault = emaildefault.Select("category = 'AMDEFAULT'");
                                            string jemail = amDefault[0]["jemail"].ToString();
                                            AM = JObject.Parse(jemail)["AM"].Value<string>();
                                            gerente = JObject.Parse(jemail)["manager"].Value<string>();
                                            valueTeam = JObject.Parse(jemail)["VT"].Value<string>();
                                            emailam = JObject.Parse(jemail)["email"].Value<string>();
                                            nombream = JObject.Parse(jemail)["name"].Value<string>();
                                        }

                                        JObject jdSAP = jaisons["purchaseOrderAdditionalData"];
                                        jdSAP["accountManager"] = AM.ToUpper();
                                        jdSAP["managerSector"] = gerente;
                                        jdSAP["valueTeam"] = valueTeam;
                                        jdSAP["customerInstitute"] = cliente_institucion;
                                        jdSAP["contactId"] = contacto_institucion;

                                        if (!string.IsNullOrWhiteSpace(cliente_institucion))
                                        {

                                            string cust = cliente_institucion.Substring(2, cliente_institucion.Length - 2);
                                            DataRow[] infoCustomer = customers.Select("idClient = '" + cust + "'");
                                            string customerName = "", contactName = "";
                                            if (infoCustomer.Count() == 1)
                                            {
                                                customerName = infoCustomer[0]["name"].ToString();
                                            }
                                            else
                                            {
                                                //busca el nombre en SAP
                                                customerName = crBids.GetInfoBP(cliente_institucion, 1);
                                            }

                                            //buscar el nombre del contacto
                                            contactName = crBids.GetInfoBP(contacto_institucion, 2);

                                            jdSAP["customerName"] = customerName;
                                            jdSAP["contactName"] = contactName;
                                        }

                                        //guardar el JSON con los datos de sap en poAddData
                                        jaisons["purchaseOrderAdditionalData"] = jdSAP;
                                        #endregion

                                        #region Guardar en BD

                                        for (int b = 0; b < sicopFields.Rows.Count; b++)
                                        {
                                            //string field = full_columns.Rows[b]["Field"].ToString(); //el nombre de la columna
                                            //string comment = full_columns.Rows[b]["Comment"].ToString(); //el comentario de la columna

                                            //el nombre de la columna de todas las DB
                                            string columnn = sicopFields.Rows[b]["poColumn"].ToString();
                                            //es el campo de SICOP o del comentario en la BD (para los campos que no son de sicop)
                                            string sicopField = sicopFields.Rows[b]["csicop"].ToString();
                                            //el nombre de la tabla de la columna purchaseOrder, purchaseOrderAdditionalData, products, evaluations
                                            string tableDb = sicopFields.Rows[b]["idjson"].ToString();

                                            JObject jObject = jaisons[tableDb];
                                            if (columnn == "products")
                                            {
                                                //llenar bienes
                                                DataRow[] fila_select = licitacion.Select("nombre_tabla = '" + sicopField + "'");
                                                string bs = JsonConvert.SerializeObject((DataTable)fila_select[0]["data_table"]);
                                                //bienes_servicios = bs;
                                                //DataTable productos = (DataTable)licitacion.Select($"nombre_tabla = '{sicopField}'")[0]["data_table"];
                                                //foreach (DataRow row in productos.Rows)
                                                //{

                                                //}

                                                jObject[columnn] = bs;
                                            }
                                            else if (columnn == "accountManager" || columnn == "managerSector" || columnn == "valueTeam" || columnn == "contactId" || columnn == "customerInstitute")
                                            {
                                                //ya esta lleno
                                            }
                                            else if (columnn == "timeDelivery")
                                            {
                                                string plazo = "";
                                                foreach (var item in entrega)
                                                {
                                                    plazo = plazo + item.Value;
                                                }
                                                jObject[columnn] = plazo;
                                            }
                                            else if (columnn == "participation")
                                            {
                                                //agregar espacio en blanco para la busqueda en S&S
                                                jObject[columnn] = " ";
                                            }
                                            else if (columnn == "evaluations")
                                            {
                                                evaluacion = JsonConvert.SerializeObject(factores_de_evaluacion).ToString(); // JObject.Parse(JsonConvert.SerializeObject(factores_de_evaluacion));
                                                try
                                                {
                                                    evaluacion = evaluacion.Replace("\t", " ");
                                                    evaluacion = evaluacion.Replace("\n", " ");
                                                }
                                                catch (Exception) { }
                                                jObject[columnn] = evaluacion;
                                            }
                                            else if (columnn == "participationAmount" || columnn == "complianceAmount")
                                            {
                                                DataRow[] fila_select = licitacion.Select("campo = '" + "Monto o porcentaje" + "'");
                                                int aRow = (columnn == "participationAmount") ? 0 : 1;
                                                string valor = fila_select[aRow]["value"].ToString();
                                                jObject[columnn] = valor;
                                            }
                                            else if (columnn == "offerValidity")
                                            {
                                                DataRow[] fila_select = licitacion.Select("nombre_tabla = '[ 5. Oferta ]' and campo = '" + sicopField + "'");
                                                jObject[columnn] = fila_select[0]["value"].ToString();
                                            }
                                            else if (columnn == "receptionClarification")
                                            {
                                                DataRow[] fila_select = licitacion.Select("campo = '" + sicopField + "'");
                                                aclaraciones = fila_select[0]["value"].ToString().Replace("Solicitud de Aclaracion", "").Replace("Consulta de Aclaracion", "").Trim();
                                                try { aclaraciones = aclaraciones.Substring(0, 16); } catch (Exception) { }
                                                string aclarDate = "";
                                                try
                                                {
                                                    aclarDate = DateTime.Parse(aclaraciones).ToString("yyyy-MM-dd HH:mm:ss");
                                                }
                                                catch (Exception)
                                                { aclarDate = "NULL"; }
                                                jObject[columnn] = aclarDate;
                                            }
                                            else if (columnn == "month")
                                            {
                                                string mess = "";
                                                try
                                                {
                                                    DataRow[] fila_select = licitacion.Select("campo = '" + "Fecha/hora de apertura de ofertas" + "'");
                                                    int.TryParse(fila_select[0]["value"].ToString().Substring(3, 2), out int mes);
                                                    mess = CultureInfo.GetCultureInfo("es-CR").DateTimeFormat.GetMonthName(mes);
                                                }
                                                catch (Exception) { }
                                                jObject[columnn] = mess;
                                            }
                                            else if (columnn == "processType")
                                            {
                                                string processType = "0";
                                                try
                                                {
                                                    DataRow[] fila_select = licitacion.Select("campo = '" + sicopField + "'");
                                                    processType = fila_select[0]["value"].ToString();
                                                    DataRow[] pType = processTypeData.Select($"processType = '{processType}'");
                                                    processType = (pType.Count() > 0) ? processType = pType[0]["id"].ToString() : "0";
                                                }
                                                catch (Exception)
                                                {

                                                }
                                                jObject[columnn] = processType;

                                            }
                                            else if (columnn == "receptionObjections" || columnn == "changeDate" || columnn == "productLine"
                                                || columnn == "oppType" || columnn == "salesType" || columnn == "noParticipationReason" || columnn == "oppType")
                                            {
                                                jObject[columnn] = "NULL";

                                            }
                                            else if (columnn == "receptionClosing" || columnn == "offerOpening" || columnn == "publicationDate")
                                            {
                                                DataRow[] fila_select = licitacion.Select("campo = '" + sicopField + "'");
                                                string date = "";
                                                try
                                                {
                                                    date = DateTime.Parse(fila_select[0]["value"].ToString()).ToString("yyyy-MM-dd HH:mm:ss");

                                                }
                                                catch (Exception)
                                                { date = "NULL"; }
                                                jObject[columnn] = date;

                                            }
                                            else if (columnn == "budget")
                                            {
                                                DataRow[] fila_select = licitacion.Select("campo = '" + sicopField + "'");
                                                presupuesto = fila_select[0]["value"].ToString();

                                                if (presupuesto == "")
                                                {
                                                    //el presupuesto en colones esta en blanco, tomar dolares y convertirlo
                                                    fila_select = licitacion.Select("campo = 'Presupuesto total estimado USD (Opcional)'");
                                                    presupuesto = fila_select[0]["value"].ToString();

                                                    string[] split2 = presupuesto.Split(',');
                                                    split2 = split2[0].Split('[');
                                                    long.TryParse(split2[0].Replace(".", ""), out long presupuesto_long2);

                                                    long tipo_cambio = 1;
                                                    try
                                                    {
                                                        string htmlTC = "";
                                                        string today = DateTime.Today.ToString("yyyy/MM/dd");
                                                        using (WebClient client = new WebClient())
                                                        {
                                                            client.Encoding = UTF8Encoding.UTF8;
                                                            htmlTC = client.DownloadString("https://gee.bccr.fi.cr/indicadoreseconomicos/Cuadros/frmVerCatCuadro.aspx?CodCuadro=400&Idioma=1&FecInicial=" + today + "&FecFinal=" + today + "&Filtro=0");
                                                        }
                                                        HtmlAgilityPack.HtmlDocument doctc = new HtmlAgilityPack.HtmlDocument();
                                                        doctc.LoadHtml(htmlTC);
                                                        string tipc = HttpUtility.HtmlDecode(doctc.DocumentNode.SelectSingleNode("//*[@id=\"theTable400\"]/tr[2]/td[2]/table/tr/td/table/tr/td").InnerText).Trim().Split(',')[0];
                                                        tipo_cambio = long.Parse(tipc);
                                                    }
                                                    catch (Exception)
                                                    {
                                                        tipo_cambio = 600;
                                                    }

                                                    presupuesto_long2 = tipo_cambio * presupuesto_long2;
                                                    presupuesto = presupuesto_long2.ToString("N", new CultureInfo("is-IS")) + " [CRC]";

                                                }

                                                jObject[columnn] = presupuesto;
                                            }
                                            else //sin reglas especificas
                                            {
                                                string v = "";
                                                try
                                                {
                                                    DataRow[] fila_select = licitacion.Select("campo = '" + sicopField + "'");
                                                    if (fila_select.Count() > 0)
                                                    {
                                                        v = fila_select[0]["value"].ToString();
                                                    }
                                                }
                                                catch (Exception) { }
                                                JToken vv = jObject[columnn];

                                                if (vv == null)
                                                {
                                                    jObject[columnn] = v;
                                                }
                                                else
                                                {
                                                    if (string.IsNullOrEmpty(vv.ToString()))
                                                    {
                                                        jObject[columnn] = v;
                                                    }
                                                }

                                            }

                                            jaisons[tableDb] = jObject;

                                            if (columnn == "description")
                                            {
                                                #region Identificar si es de Interes para GBM y descargar los adjuntos
                                                DataRow[] fila_select = licitacion.Select("campo = '" + sicopField + "'");
                                                descripcion = fila_select[0]["value"].ToString();
                                                //Filtrar si es de interes o no de GBM con las palabras claves y llenar Interes para GBM
                                                //si es de interes descargue los adjuntos

                                                interes = bidLogical.KeyMatch(descripcion, words);
                                                if (interes != "SI")
                                                {
                                                    //Si no es de interes, buscarlo tambien en la descripcion de los productos
                                                    DataTable productos = (DataTable)licitacion.Select("nombre_tabla = '[ 11. Información de bien, servicio u obra ]'")[0]["data_table"];
                                                    foreach (DataRow row in productos.Rows)
                                                    {
                                                        interes = bidLogical.KeyMatch(row["Nombre"].ToString(), words);
                                                        if (interes == "SI")
                                                        {
                                                            break;
                                                        }
                                                    }
                                                }

                                                if (interes == "SI")
                                                {
                                                    //de interes
                                                    //descargar
                                                    //es una lista con los nombres de los archivos descargados
                                                    List<string> fileList = bidLogical.DownloadAttachment(chrome, adjuntos, bidNumberFromWebPage);

                                                    string json_files = JsonConvert.SerializeObject(fileList);

                                                    fileList.ForEach(delegate (string name)
                                                    {
                                                        bool ins = liccr.InsertFile(bidNumberFromWebPage, root.downloadfolder + "\\" + bidNumberFromWebPage + "\\" + name, mandante);
                                                    });

                                                    JObject jint = jaisons["purchaseOrderAdditionalData"];
                                                    jint["gbmStatus"] = "1";
                                                    jaisons["purchaseOrderAdditionalData"] = jint;

                                                    //JObject jad = jaisons["nom"];
                                                    //jad["adjuntos"] = json_files;
                                                    //jaisons["nom"] = jad;
                                                }
                                                else
                                                {
                                                    JObject jint = jaisons["purchaseOrderAdditionalData"];
                                                    jint["gbmStatus"] = "2";
                                                    jaisons["purchaseOrderAdditionalData"] = jint;
                                                }

                                                #endregion
                                            }
                                            else if (columnn == "publicationDate") //es para crear la tr de la tabla del email
                                            {
                                                publicacion = licitacion.Select("campo = '" + sicopField + "'")[0]["value"].ToString();
                                            }
                                            else if (columnn == "offerOpening") //es para crear la tr de la tabla del email
                                            {
                                                apertura = licitacion.Select("campo = '" + sicopField + "'")[0]["value"].ToString();
                                            }
                                        }



                                        bool insert = liccr.InsertRowSS(jaisons);
                                        if (!insert)
                                        {
                                            validar_lineas = false;
                                        }
                                        #endregion

                                        #region Enviar_noti_idividuales

                                        //notifiación para AM (gerentes) wteams, correo
                                        //Institución, concurso, descripción, Fecha/hora de publicación, fecha límite de aclaraciones y la fecha de apertura, el presupuesto.
                                        //enviar notificación por Wteams,

                                        string mensaje = "Se le notifica que se ha publicado el siguiente concurso en SICOP:\r\n - **" + institution + "**: " + bidNumberFromWebPage + " - " + descripcion + ". Inicia: " + publicacion + " al " + aclaraciones + " - Fecha de apertura: " + apertura + " - Presupuesto: " + presupuesto + "\n\n" + "Por favor hacer click [aqui](https://smartsimple.gbm.net) para indicar su participación.";

                                        log.LogDeCambios("Nueva solicitud", root.BDProcess, emailam, "Nueva solicitud de SICOP", mensaje, "");

                                        wt.SendNotification(emailam, "Nuevo Concurso SICOP", mensaje);
                                        respFinal = respFinal + "\\n" + "Extracción nuevo Concurso SICOP: " + mensaje;

                                        #endregion

                                        #region Crear_correo

                                        //crear la tabla del email para luego enviarselo  
                                        //una tabla de email por AM, diccionario[AM] = tabla;

                                        string tabla_correo_tr = "<tr>" +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + bidNumberFromWebPage + "</p></td>     " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + descripcion + "</p></td>  " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + institution + "</p></td>      " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + publicacion + "</p></td>  " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + aclaraciones + "</p></td> " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + apertura + "</p></td>     " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + presupuesto + "</p></td>  " +
                                                                   "</tr>";


                                        string tabla_correog_tr = "<tr>" +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + bidNumberFromWebPage + "</p></td>     " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + descripcion + "</p></td>  " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + institution + "</p></td>      " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + publicacion + "</p></td>  " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + aclaraciones + "</p></td> " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + apertura + "</p></td>     " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + presupuesto + "</p></td>  " +
                                                                   //"<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + AM + "</p></td>  " +
                                                                   "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + nombream + "</p></td>  " +
                                                                   "</tr>";


                                        //revisar si ya tiene el AM?
                                        if (lista_am.ContainsKey(AM))
                                            //Update
                                            lista_am[AM] = lista_am[AM] + tabla_correo_tr;
                                        else
                                            //Insert
                                            lista_am.Add(AM, tabla_correo_th + tabla_correo_tr);


                                        //verificar el presupuesto para enviarselo al gerente de ventas (gerente_ventas de la tabla entidades) en colones!!
                                        //ejemplo 1.422.393,36 [CRC]
                                        string[] split = presupuesto.Split(',');
                                        split = split[0].Split('[');
                                        long.TryParse(split[0].Replace(".", ""), out long presupuesto_long);

                                        if (interes == "SI")
                                        {
                                            if (presupuesto_long > 60000000)
                                            {
                                                //if mayor a 60 M al gerente

                                                // revisar si ya tiene el gerente ?
                                                if (lista_g.ContainsKey(gerente))
                                                    lista_g[gerente] = lista_g[gerente] + tabla_correog_tr;
                                                else
                                                    lista_g.Add(gerente, tabla_correog_th + tabla_correog_tr);

                                                // lista_gerente = lista_gerente + tabla_correog_tr;
                                                if (presupuesto_long > 120000000)
                                                {
                                                    //elseif mayor a 120m al Ggeneral
                                                    lista_gerente_general = lista_gerente_general + tabla_correog_tr;
                                                }

                                            }
                                        }
                                        #endregion

                                        js.ExecuteScript("history.back();");
                                    }
                                    catch (Exception ex)
                                    {
                                        validar_lineas = false;
                                        tabla_correo_error = tabla_correo_error + "<tr>" +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + bidNumberFromWebPage + "</p></td>     " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + ex.Message.ToString() + "</p></td>  " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + ex.ToString() + "</p></td>      " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>  " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td> " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>     " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>  " +
                                                                           //"<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + AM + "</p></td>  " +
                                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>  " +
                                                                           "</tr>";
                                        js.ExecuteScript("history.back();");
                                    }
                                }
                            }
                            else
                            {
                                console.WriteLine(" > " + bidNumberFromWebPage + " YA existe, omitiendo");
                            }
                            e++;
                        }

                        #endregion

                    }
                    catch (Exception ex)
                    {
                        validar_lineas = false;
                        tabla_correo_error = tabla_correo_error + "<tr>" +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + bidNumberFromWebPage + "</p></td>     " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + ex.Message.ToString() + "</p></td>  " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + ex.ToString() + "</p></td>      " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>  " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td> " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>     " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>  " +
                                                           //"<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + AM + "</p></td>  " +
                                                           "<td style='padding:0.75pt;'><p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0;'>" + "" + "</p></td>  " +
                                                           "</tr>";

                    }
                }

                #endregion

                #region Cerrar chrome
                chrome.Close();
                proc.KillProcess("chromedriver", true);
                #endregion

                #region Enviar los emails

                console.WriteLine("Enviar correos");

                string htmlpage = Properties.Resources.emailLpCr;

                #region correo de error

                if (validar_lineas == false)
                {
                    console.WriteLine("Ocurrió un error ya que es posible que no se pudieron insertar algunas lineas en la base de datos.");
                    string tabla = tabla_correo_th + tabla_correo_error + "</table>";
                    string htmlpage5 = htmlpage;
                    htmlpage5 = htmlpage5.Replace("{subject}", "Notificaciones Nuevos Concursos SICOP");
                    htmlpage5 = htmlpage5.Replace("{cuerpo}", "Estimado Account Manager, se le notifica que se han publicado los siguiente concursos en SICOP, por favor hacer click en el botón abajo para indicar su participación.");
                    htmlpage5 = htmlpage5.Replace("{contenido}", tabla);
                    string[] cc = { "dmeza@gbm.net" };
                    mail.SendHTMLMail(htmlpage5, new string[] {"appmanagement@gbm.net"}, "Error en concursos SICOP del Gobierno de Costa Rica - Fecha: " + DateTime.Now.ToString("D"), cc);
                }
                #endregion

                #region enviar correo a los AM
                console.WriteLine("Enviando correo a los Account Managers...");
                foreach (var item in lista_am)
                {
                    string stringCopias = "";
                    string htmlpage2 = htmlpage;
                    string emailam = item.Key + "@gbm.net";
                    string tabla = item.Value + "</table>";
                    htmlpage2 = htmlpage2.Replace("{subject}", "Notificaciones Nuevos Concursos SICOP");
                    htmlpage2 = htmlpage2.Replace("{cuerpo}", "Estimado Account Manager, se le notifica que se han publicado los siguiente concursos en SICOP, por favor hacer click en el botón abajo para indicar su participación.");
                    htmlpage2 = htmlpage2.Replace("{contenido}", tabla);
                    if (emailam.ToLower() == "rmena@gbm.net")
                    {
                        try
                        {
                            string sql = @"SELECT 
                                            MIS.digital_sign.email
                                            FROM SS_Access_Permissions.UserAccess
                                            INNER JOIN MIS.digital_sign ON SS_Access_Permissions.UserAccess.fk_SignID = MIS.digital_sign.id
                                            INNER JOIN SS_Access_Permissions.Permissions ON SS_Access_Permissions.UserAccess.fk_Permissions = SS_Access_Permissions.Permissions.id
                                            WHERE SS_Access_Permissions.Permissions.name = 'CostaRicaBids VT TELCO'";

                            DataTable ccdt = crud.Select( sql, "SS_Access_Permissions");
                            string[] cc = new string[ccdt.Rows.Count];
                            for (int i = 0; i < ccdt.Rows.Count; i++)
                            {
                                cc[i] = ccdt.Rows[i]["email"].ToString();
                                stringCopias += ccdt.Rows[i]["email"].ToString() + ", ";
                            }

                            console.WriteLine($"Enviando correo a {emailam} (Account Manager), con copia a: {stringCopias}.");
                            mail.SendHTMLMail(htmlpage2, new string[] { emailam }, "Concursos SICOP del Gobierno de Costa Rica - Fecha: " + DateTime.Now.ToString("D"), cc);

                        }
                        catch (Exception e)
                        {
                            string msg = $"No se pudo enviar correo al sender: {emailam}, ni a sus respectivas copias. Detalle del error: " + e.ToString();
                            console.WriteLine(msg);
                            mail.SendHTMLMail(msg, new string[] {"appmanagement@gbm.net"}, "No he podido enviar el correo electrónico - GetBidSS - Databot", new string[] { "dmeza@gbm.net", "epiedra@gbm.net" });
                        }
                    }
                    else if (emailam.ToLower() == "@gbm.net")
                    //Significa que no hay ningún Account Manager asociado, por lo tanto se le envía el correo al AMDEFAULT de la tabla emailAddress
                    //de la DB Costa_Rica_Bids en 10.7.60.137 de PHPMMyAdmin.
                    {
                        try
                        {
                            #region Extraer Email del AM Default  
                            string emailAMDefault = "[" + emaildefault.Select("category = 'AMDEFAULT'")[0]["jemail"].ToString() + "]";
                            DataTable auxDtAMEmail = (DataTable)JsonConvert.DeserializeObject(emailAMDefault, (typeof(DataTable)));
                            emailAMDefault = auxDtAMEmail.Rows[0]["email"].ToString();
                            #endregion

                            emailam = emailAMDefault;

                            console.WriteLine($"Enviando correo a {emailam} (Account Manager Default), esto debido a que no existe ningún AM asociado a la cuenta.");
                            mail.SendHTMLMail(htmlpage2, new string[] { emailam }, "Concursos SICOP del Gobierno de Costa Rica - Fecha: " + DateTime.Now.ToString("D"));
                        }
                        catch (Exception e)
                        {
                            console.WriteLine($"No se pudo enviar correo al sender: {emailam}, ni a sus respectivas copias.");
                            console.WriteLine("Detalle del error: " + e.ToString());
                        }
                    }
                    else
                    {
                        try
                        {

                            console.WriteLine($"Enviando correo a {emailam} (Account Manager), sin copias (esto porque no es a rmena@gbm.net).");
                            mail.SendHTMLMail(htmlpage2, new string[] { emailam }, "Concursos SICOP del Gobierno de Costa Rica - Fecha: " + DateTime.Now.ToString("D"));

                        }
                        catch (Exception e)
                        {
                            string msg = $"No se pudo enviar correo al sender: {emailam}, sin copias. Detalle del error: " + e.ToString();
                            console.WriteLine(msg);
                            mail.SendHTMLMail(msg, new string[] {"appmanagement@gbm.net"}, "No he podido enviar el correo electrónico - GetBidSS - Databot", new string[] { "dmeza@gbm.net", "epiedra@gbm.net" });
                        }
                    }

                }
                #endregion

                #region enviar los emails a los gerentes
                console.WriteLine("Enviando correo a los gerentes...");
                if (lista_g.Count > 0)
                {
                    string stringCopias = "";
                    JArray j_copias = JArray.Parse(emaildefault.Select("category = 'PRESUPUESTO100K'")[0]["jemail"].ToString());
                    string[] a_copias = new string[j_copias.Count];

                    for (int i = 0; i < j_copias.Count; i++)
                    {
                        a_copias[i] = j_copias[i]["email"].ToString();
                        stringCopias += j_copias[i]["email"].ToString() + ", ";
                    }

                    foreach (var item in lista_g)
                    {
                        string htmlpage2 = htmlpage;
                        string sender1 = item.Key;
                        string tabla = item.Value + "</table>";

                        htmlpage2 = htmlpage2.Replace("{subject}", "Notificaciones Nuevos Concursos SICOP");
                        htmlpage2 = htmlpage2.Replace("{cuerpo}", "Estimado, se le notifica que se han publicado los siguiente concursos en SICOP cuyo presupuesto es de su interes.");
                        htmlpage2 = htmlpage2.Replace("{contenido}", tabla);
                        console.WriteLine($"Enviando correo a {sender1} (Gerente), con copia a: {stringCopias}.");
                        mail.SendHTMLMail(htmlpage2, new string[] { sender1 + "@gbm.net" }, "Concursos SICOP del Gobierno de Costa Rica - Fecha: " + DateTime.Now.ToString("D"), a_copias);
                    }

                }
                #endregion
                #region enviar coreo a GERENTE GENERAL
                console.WriteLine("Enviando correo al Gerente General...");
                if (lista_gerente_general != "")
                {
                    string stringCopias = "";
                    JArray j_copias = JArray.Parse(emaildefault.Select("category = 'PRESUPUESTO200K'")[0]["jemail"].ToString());
                    string[] a_copias = new string[j_copias.Count];

                    for (int i = 0; i < j_copias.Count; i++)
                    {
                        a_copias[i] = j_copias[i]["email"].ToString();
                        stringCopias += j_copias[i]["email"].ToString() + ",";
                    }



                    htmlpage = htmlpage.Replace("{subject}", "Notificaciones Nuevos Concursos SICOP");
                    htmlpage = htmlpage.Replace("{cuerpo}", "Estimado, se le notifica que se han publicado los siguiente concursos en SICOP cuyo presupuesto es de su interes.");
                    htmlpage = htmlpage.Replace("{contenido}", tabla_correog_th + lista_gerente_general + "</table>");
                    //rrivera@gbm.net
                    console.WriteLine("Enviando correo a rrivera@gbm.net (Gerente General),con copias a " + stringCopias + ".");
                    mail.SendHTMLMail(htmlpage, new string[] { "rrivera@gbm.net" }, "Concursos SICOP del Gobierno de Costa Rica - Fecha: " + DateTime.Now.ToString("D"), a_copias);

                }
                #endregion
                #region enviar correo de resumen a licitaciones de CR
                if (root.planner[root.BDClass]["conta"][0] == "13")
                {
                    console.WriteLine("Enviando correo DMO...");
                    //el robot se ejecuta 13 veces al dia desde las 7 am hasta las 7 pm, cuando se ejecuta la ultima, extrae todo lo del dia y le envia un consolidado a licitaciones cr
                    DataTable bids = crud.Select( $@"SELECT * FROM `purchaseOrder` WHERE createdAt >= '{DateTime.Now.ToString("yyyy-MM-dd")}'
UNION
SELECT * FROM `purchaseOrderBackup` WHERE createdAt >= '2023-02-01'", "costa_rica_bids_db");

                    string ruta = root.FilesDownloadPath + "\\" + $"licitaciones_{DateTime.Now.ToString("yyyy_MM_dd")}.xlsx";
                    MsExcel ms = new MsExcel();
                    ms.CreateExcel(bids, "bids", ruta, false);
                    #region Extraer Email del AM Default  
                    string emailDmo = emaildefault.Select("category = 'DMOLIDER'")[0]["jemail"].ToString();
                    JObject jsonEmail = JObject.Parse(emailDmo);
                    string dmoEmail = jsonEmail["email"].ToString();
                    #endregion
                    string htmlpageEnd = Properties.Resources.emailLpCr;
                    htmlpageEnd = htmlpageEnd.Replace("{subject}", $"Notificaciones Nuevos Concursos SICOP del día {DateTime.Now.ToString("yyyy-MM-dd")}");
                    htmlpageEnd = htmlpageEnd.Replace("{cuerpo}", "Estimados, se le notifica que se han publicado los siguiente concursos en SICOP.");
                    htmlpageEnd = htmlpageEnd.Replace("{contenido}", "Ver Excel Adjunto");
                    mail.SendHTMLMail(htmlpageEnd, new string[] { dmoEmail }, $"Notificaciones Nuevos Concursos SICOP del día {DateTime.Now.ToString("yyyy-MM-dd")}", null, new string[] { ruta });


                }
                #endregion
                #endregion

                #region Borrar las descargas anteriores
                try
                {
                    Directory.Delete(root.downloadfolder, true);
                    Directory.CreateDirectory(root.downloadfolder);
                }
                catch (Exception) { }
                #endregion

                root.requestDetails = respFinal;
                root.BDUserCreatedBy = "MGARCIA";

            }

            return cant_filas;
        }

    }

}
