using System;
using System.Collections.Generic;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Projects.PanamaBids;
using DataBotV5.Data.Projects.PanamaBids;
using System.Text;
using System.Text.RegularExpressions;
using DataBotV5.Automation.MASS.PanamaBids;
using System.Net.Mail;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using System.Linq;
using System.IO;
using System.Data;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.Web;
using DataBotV5.Data.Process;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.WEB.PanamaBids
{
    /// <summary>
    /// Clase Web Automation encargada de extraer licitaciones de GBPA vía email.
    /// </summary>
    class GetBidEmail
    {
        #region variables_globales
        PanamaBidsLogical paBids = new PanamaBidsLogical();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();
        BidsGbPaSql lpsql = new BidsGbPaSql();
        Credentials cred = new Credentials();
        ProcessInteraction proc = new ProcessInteraction();
        ProcessAdmin padmin = new ProcessAdmin();
        WebInteraction sel = new WebInteraction();
        Log log = new Log();
        WebexTeams wt = new WebexTeams();
        SharePoint sharep = new SharePoint();
        Database dBase = new Database();
        ConsoleFormat console = new ConsoleFormat();
        string ssMandante = "PRD";

        string respFinal = "";
        #endregion

        /// <summary>
        /// extraer la información por correo electronico: aprobadas y publicadas (depende del subject)
        /// </summary>
        public void Main()
        {
            if (mail.GetAttachmentEmail("Solicitudes convenio gbmpa", "Procesados", "Procesados convenio gbmpa"))
            {
                console.WriteLine("Procesando....");
                bool respuesta = getBid(root.Email_Body);
                console.WriteLine("Creando Estadisticas");
                if (respuesta == false)
                { root.requestDetails = "Error al cargar a información"; }
                else { root.requestDetails = "Se cargo la data"; }
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        private bool getBid(string body)
        {

            #region variable privadas
            bool validar_lineas = true;
            string sectorf = "";
            string bodyClean = ""; string[] link;
            string subject = root.Subject;
            string convenio = "EQUIPOS INFORMÁTICOS Y TELECOMUNICACIONES";// "BIENES INFORMÁTICOS, REDES Y COMUNICACIONES";
            string resp_sql = "";
            bool resp_add_sql = true;
            string[] adjunto = new string[1];
            int cont_adj = 0;
            Dictionary<string, string> valores = new Dictionary<string, string>();
            Dictionary<string, string> campos_opp = new Dictionary<string, string>();
            string tipo = "";
            #endregion
            #region aprobada o publicada
            tipo = (subject.Contains("Orden de Compra Aprobada")) ? "aprobada" : "publicada";
            #endregion
            #region extrae info del body
            console.WriteLine(" Extrayendo info del Body y Subject");
            Regex reg;
            reg = new Regex("[*'\"_&+^><@]");
            bodyClean = reg.Replace(body, string.Empty);
            #endregion
            if (tipo == "aprobada")
            {
                getBidEnRefrendo(bodyClean, body, convenio);
            }
            else //publicada
            {
                getBidRefrendada(bodyClean, body, convenio);
            }

            return resp_add_sql;
        }
        private void getBidEnRefrendo(string bodyClean, string body, string convenio)
        {
            #region variables privadas
            DataTable entities = lpsql.entitiesInfo();
            DataTable productBrand = lpsql.productsInfo();
            bool newEntity = false;
            List<string> newEntities = new List<string>();
            string resp_sql = "";
            bool resp_add_sql = true;
            string[] link;
            int limite = 0;
            #endregion
            #region extrae entidad
            string[] stringSeparators0 = new string[] { "Por este medio les deseamos informar que la entidad " };
            link = bodyClean.Split(stringSeparators0, StringSplitOptions.None);
            string entidad = link[1].ToString().Trim();
            limite = entidad.IndexOf(" ha generado una Orden de compra");
            if (entidad.Length >= limite)
            { entidad = entidad.Substring(0, limite).Trim(); }
            #endregion
            #region extrae registro unico 
            string[] stringSeparators1 = new string[] { "Trámite No. " };
            link = bodyClean.Split(stringSeparators1, StringSplitOptions.None);
            string singleOrderRecord = link[1].ToString().Trim();
            limite = singleOrderRecord.IndexOf(" a través del Catálogo Electrónico");
            if (singleOrderRecord.Length >= limite)
            { singleOrderRecord = singleOrderRecord.Substring(0, limite).Trim(); }
            #endregion
            #region extraer producto y cantidad
            string[] stringSeparators = new string[] { "Cantidad\t\r\n" };
            link = body.Split(stringSeparators, StringSplitOptions.None);
            string tabla = "";
            try
            {
                tabla = link[1].ToString().Trim();
            }
            catch (Exception ex)
            {
                stringSeparators = new string[] { "Cantidad\t \r\n" };
                link = body.Split(stringSeparators, StringSplitOptions.None);
                tabla = link[1].ToString().Trim();
            }

            limite = tabla.IndexOf("Referente:");
            if (tabla.Length >= limite)
            { tabla = tabla.Substring(0, limite).Trim(); }

            string[] stringSeparators3 = new string[] { "\t" };
            link = tabla.Split(stringSeparators3, StringSplitOptions.None);
            #endregion
            #region extraer fecha de registro
            string fecha_registro = root.ReceivedTime.ToString("yyyy-MM-dd");
            float sub_total = 0;
            #endregion
            #region buscar sector 
            string sector = "";

            System.Data.DataRow[] entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
            if (entityInfo.Length != 0)
            {
                sector = entityInfo[0]["sector"].ToString();
            }

            else
            {
                sector = "PS";
                newEntity = true;
                newEntities.Add(entidad);
            }
            #endregion
            #region extrae cada producto
            List<calculateData> datosCalculados = new List<calculateData>();
            List<PoProductMacro> PoproductInfo = new List<PoProductMacro>();
            for (int i = 0; i < link.Length; i++)
            {
                calculateData datoCalculados = new calculateData();
                PoProductMacro pInfo = new PoProductMacro();
                string producto = link[i + 1].ToString().Trim();
                string cantidad = link[i + 2].ToString().Trim();
                //extrae los dias habiles de acuerdo a la cantidad
                int dias_h = paBids.businessDays(int.Parse(cantidad));
                datoCalculados.deliveryDay = dias_h;
                datosCalculados.Add(datoCalculados);
                //extrae la marca del producto
                string product_id = producto.Substring(0, 2);
                if (product_id.Contains("-"))
                { product_id = producto.Substring(0, 1); }

                //buscar en tabla#1
                //string marca = marca_product(product_id);
                System.Data.DataRow[] pBrand = productBrand.Select($"productCode ='{product_id}'"); //like '%" + institu + "%'"
                string marca = "";
                string precio = "";
                if (pBrand.Length != 0)
                {
                    marca = pBrand[0]["brandText"].ToString();
                    precio = pBrand[0]["price"].ToString();
                }
                else
                {
                    marca = "Lenovo";
                    precio = "1";
                }

                string total = "0";
                try
                {
                    precio = precio.Replace(".", ",");
                    float total_prod = (float.Parse(precio) * float.Parse(cantidad));
                    total = total_prod.ToString();
                    total = total.Replace(",", ".");
                    sub_total += total_prod;

                }
                catch (Exception ex)
                {

                }
                pInfo.singleOrderRecord = singleOrderRecord;
                pInfo.productCode = product_id;
                pInfo.quantity = cantidad;
                pInfo.totalProduct = total;
                pInfo.orderType = "";
                pInfo.active = "1";
                pInfo.createdBy = "databot";
                PoproductInfo.Add(pInfo);
                i++; i++;
            }
            #endregion
            console.WriteLine("  Agregando información a la base de datos");
            int maxDeliveryDay = datosCalculados.Max(t => t.deliveryDay);
            bool add_sql = lpsql.insertInfoApproved(convenio, entidad, singleOrderRecord, fecha_registro, maxDeliveryDay, sector, sub_total);
           
            log.LogDeCambios("Creacion", root.BDProcess, "Ventas Panama", "Agregar info a convenio GBPA - Orden Aprobada", singleOrderRecord, add_sql.ToString());
            respFinal = respFinal + "\\n" + "Agregar info a convenio GBPA - Orden Aprobada: "+ singleOrderRecord + " " + add_sql.ToString();

            if (!add_sql)
            {
                string[] cc = { "appmanagement@gbm.net", "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + singleOrderRecord, cc);
            }
            else
            {
                //insertar productos
                bool addProducts = lpsql.insertInfoProductsMacro(PoproductInfo);
                if (!addProducts)
                {
                    string[] cc = { "appmanagement@gbm.net", "dmeza@gbm.net" };
                    mail.SendHTMLMail("Dio error al intentar ingresar los productos de la siguiente orden: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + singleOrderRecord, cc);
                }
                //todo salio bien, se envia notificación.
                JArray j_copias = JArray.Parse(lpsql.getEmail("LICPA"));

                for (int i = 0; i < j_copias.Count; i++)
                {
                    string email = j_copias[i]["email"].ToString();

                    wt.SendNotification(email, "Nueva Orden de Compra en Refrendo", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + singleOrderRecord + "**, la cual se encuentra En Refrendo.<br><br>Haga click en el siguiente enlace: <a href=\"https://smartsimple.gbm.net/admin/PanamaBids/agreement\" > Orden de Compra - Portal GBM</a> para ver el documento");
                }

            }

        }
        private void getBidRefrendada(string bodyClean, string body, string convenio)
        {
            #region variables privadas
            DataTable entities = lpsql.entitiesInfo();
            DataTable productBrand = lpsql.productsInfo();
            bool newEntity = false;
            List<string> newEntities = new List<string>();
            string resp_sql = "";
            bool resp_add_sql = true;
            string[] link;
            string brand_opp = "";
            string registro_opp = "";
            string cantidad_opp = "";
            bool validar_lineas = true;
            string sectorf = "";
            string subject = root.Subject;
            string[] adjunto = new string[1];
            int cont_adj = 0;
            //string[] valores = new string[1];
            //int cont_val = 0;
            Dictionary<string, string> valores = new Dictionary<string, string>();
            Dictionary<string, string> campos_opp = new Dictionary<string, string>();
            string tipo = "";
            #endregion
            #region extrae entidad
            string[] stringSeparators0 = new string[] { "Por este medio les deseamos informar que el (la) " };
            link = bodyClean.Split(stringSeparators0, StringSplitOptions.None);
            string entidad = link[1].ToString().Trim();
            int limite = entidad.IndexOf("ha generado la orden de compra") - 1;
            if (entidad.Length >= limite)
            { entidad = entidad.Substring(0, limite).Trim(); }
            valores.Add("ENTIDAD", entidad);
            #endregion
            #region registro unico 
            string[] stringSeparators1 = new string[] { "Orden de Compra Publicada - " };
            link = subject.Split(stringSeparators1, StringSplitOptions.None);
            string singleOrderRecord = link[1].ToString().Trim();
            valores.Add("REGISTRO_UNICO_DE_PEDIDO", singleOrderRecord);
            //text opp
            string[] stringSeparators3 = new string[] { "-RC" };
            link = singleOrderRecord.Split(stringSeparators3, StringSplitOptions.None);
            registro_opp = link[1].ToString().Trim();
            registro_opp = "-RC" + registro_opp;
            #endregion
            #region extrae link href para ir a la orden en web
            string[] stringSeparators2 = new string[] { "<" };
            link = body.Split(stringSeparators2, StringSplitOptions.None);
            string link_po = link[1].ToString().Trim();
            limite = link_po.IndexOf("> ,");
            if (link_po.Length >= limite)
            { link_po = link_po.Substring(0, limite).Trim(); }
            #endregion

            //datatable que trae la info de la PO
            DataTable excelResult = new DataTable();
            DataTable excelCisco = new DataTable();
            #region titulos_excel
            string[] columns = {
                        "Registro Unico de Pedido",
                        "Sector",
                        "Convenio",
                        "Entidad",
                        "Producto/Servicio",
                        "Cantidad de Producto",
                        "Marca del Producto",
                        "Total del Producto",
                        "Sub Total de la orden",
                        "Orden de Compra",
                        "Fecha de Registro",
                        "Fecha de Publicacion",
                        "Fianza por cumplimiento",
                        "Oportunidad",
                        "Quote",
                        "Tipo de Pedido",
                        "Sales Order",
                        "Estado de GBM",
                        "Estado de Orden",
                        "Dias de Entrega",
                        "Fecha Maxima de Entrega",
                        "Dias Faltantes",
                        "Forecast",
                        "Provincia",
                        "Lugar de Entrega",
                        "Contacto de la Empresa",
                        "Telefono del Contacto",
                        "Email del Contacto",
                        "Confirmación de Orden",
                        "Fecha Real de Entrega",
                        "Monto de Multa",
                        "Unidad Solicitante",
                        "Contacto Cuenta",
                        "Email Cuenta",
                        "Tipo de Forecast",
                        "Vendor Order",
                        "Nombre del adjunto",
                        "Account Manager Asignado",
                        "Link al documento",
                        "Comentarios"
                    };
            foreach (string item in columns)
            {
                excelResult.Columns.Add(item);
            }
            excelCisco = excelResult.Copy();
            #endregion

            #region eliminar cache and cookies chrome
            try
            {
                proc.KillProcess("chromedriver", true);
                proc.KillProcess("chrome", true);
                padmin.DeleteFiles(@"C:\Users\" + Environment.UserName + @"\AppData\Local\Google\Chrome\User Data\Default\Cache\");
                string cookies = @"C:\Users\" + Environment.UserName + @"\AppData\Local\Google\Chrome\User Data\Default\Cookies";
                string cookiesj = @"C:\Users\" + Environment.UserName + @"\AppData\Local\Google\Chrome\User Data\Default\Cookies-journal";
                if (File.Exists(cookies))
                { File.Delete(cookies); }
                if (File.Exists(cookiesj))
                { File.Delete(cookiesj); }
            }
            catch (Exception)
            { }
            #endregion

            IWebDriver chrome = sel.NewSeleniumChromeDriver(root.FilesDownloadPath);

            try
            {
                #region Ingreso al website
                console.WriteLine("  Ingresando al website");
                try
                { chrome.Navigate().GoToUrl(link_po); }
                catch (Exception)
                { chrome.Navigate().GoToUrl(link_po); }
                IJavaScriptExecutor jsup = (IJavaScriptExecutor)chrome;

                string url = chrome.Url;
                string urlId = url.Split(new char[] { '=' })[1];
                url = $"https://www.panamacompra.gob.pa/Inicio/v2/#!/PedidoPublicado/" + urlId + "?esap=0";

                try
                { chrome.Navigate().GoToUrl(url); }
                catch (Exception)
                { chrome.Navigate().GoToUrl(url); }
                #endregion

                chrome.Manage().Cookies.DeleteAllCookies();

                console.WriteLine("  Extrayendo información de la página web");

                Dictionary<string, string> AMs = new Dictionary<string, string>();

                PoInfoMacro poInfo = paBids.GetPoInfo(chrome, excelResult, excelCisco, convenio, entidad, adjunto, AMs, singleOrderRecord, cont_adj, cantidad_opp, newEntities);
                excelResult = poInfo.excel.Copy();
                excelResult.AcceptChanges();
                adjunto = poInfo.adjunto;
                cantidad_opp = poInfo.cantidadOpp;
                cont_adj = poInfo.contAdj;

                #region crear opp

                string user = "", cliente_opp = "", contacto_opp = "", sales_rep = "";

                //si la entidad es alguna de estas se determina el cliente con base a la unidad solicitante
                if (entidad == "MINISTERIO DE EDUCACION" || entidad == "CAJA DE SEGURO SOCIAL" || entidad == "MUNICIPIO DE ANTÓN")
                {

                    System.Data.DataRow[] entityInfo = entities.Select($"entities ='{excelResult.Rows[0]["Unidad Solicitante"].ToString()}'"); //like '%" + institu + "%'"
                    if (entityInfo.Length == 0)
                    {
                        entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
                        if (entityInfo.Length == 0)
                        {
                            Array.Resize(ref entityInfo, 1);
                            entityInfo[0]["customerId"] = "0010067544";
                            entityInfo[0]["contact"] = "0070032409";
                            entityInfo[0]["salesRep"] = "MAMBURG";
                        }

                    }
                    cliente_opp = entityInfo[0]["customerId"].ToString();
                    contacto_opp = entityInfo[0]["contact"].ToString();
                    sales_rep = entityInfo[0]["salesRep"].ToString();

                }
                //busca por el email de la cuenta
                //email, unidad, entidad
                else if (entidad == "MINISTERIO DE GOBIERNO" || entidad == "MINISTERIO DE SALUD" || entidad == "Ministerio de Seguridad Pública" || entidad == "MINISTERIO DE LA PRESIDENCIA")
                {
                    MailAddress address = new MailAddress(excelResult.Rows[0]["Email Cuenta"].ToString());
                    string host = address.Host.ToString();
                    System.Data.DataRow[] entityInfo = entities.Select($"entities ='{host}'"); //like '%" + institu + "%'"
                    if (entityInfo.Length == 0)
                    {
                        entityInfo = entities.Select($"entities ='{excelResult.Rows[0]["Unidad Solicitante"]}'"); //like '%" + institu + "%'"
                        if (entityInfo.Length == 0)
                        {
                            entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
                            if (entityInfo.Length == 0)
                            {
                                Array.Resize(ref entityInfo, 1);
                                entityInfo[0]["customerId"] = "0010067544";
                                entityInfo[0]["contact"] = "0070032409";
                                entityInfo[0]["salesRep"] = "MAMBURG";
                            }
                        }

                    }
                    cliente_opp = entityInfo[0]["customerId"].ToString();
                    contacto_opp = entityInfo[0]["contact"].ToString();
                    sales_rep = entityInfo[0]["salesRep"].ToString();

                }
                else
                {
                    System.Data.DataRow[] entityInfo = entities.Select($"entities ='{entidad}'"); //like '%" + institu + "%'"
                    if (entityInfo.Length == 0)
                    {
                        Array.Resize(ref entityInfo, 1);
                        entityInfo[0]["customerId"] = "0010067544";
                        entityInfo[0]["contact"] = "0070032409";
                        entityInfo[0]["salesRep"] = "MAMBURG";
                    }


                    cliente_opp = entityInfo[0]["customerId"].ToString();
                    contacto_opp = entityInfo[0]["contact"].ToString();
                    sales_rep = entityInfo[0]["salesRep"].ToString();


                }


                string opp_actual = lpsql.oppExist(singleOrderRecord);
                if (string.IsNullOrWhiteSpace(opp_actual))
                {
                    campos_opp["tipo"] = "ZOPS"; //Standard

                    string opp_text = "CM GOB-";
                    if (cantidad_opp != "")
                    {
                        if (cantidad_opp.Substring(0, 1) == ",")
                        {
                            cantidad_opp = cantidad_opp.Substring(1, cantidad_opp.Length - 1);
                        }
                    }

                    string opp_descripcion = opp_text + cantidad_opp + registro_opp;
                    if (opp_descripcion.Length > 40)
                    {
                        int cant_caract = (opp_text.Length + registro_opp.Length);
                        //opp descip = 40 caract
                        int cantf = 40 - cant_caract;
                        if (cantidad_opp.Length > cantf)
                        {
                            cantidad_opp = cantidad_opp.Substring(0, cantf);
                        }
                        opp_descripcion = opp_text + cantidad_opp + registro_opp;

                    }
                    campos_opp["descripcion"] = opp_descripcion; //"CM GOB-RL3-1-RC-003598"; //opp_text;


                    campos_opp["fecha_inicio"] = DateTime.Now.Date.ToString("yyyy-MM-dd");
                    campos_opp["Fecha_Final"] = DateTime.Now.AddDays(5).Date.ToString("yyyy-MM-dd");

                    campos_opp["Ciclo"] = "Y3"; //quotation
                    campos_opp["Origen"] = "Y08"; //Public Bid - licitaciones

                    string grupo_opp = "";
                    grupo_opp = (cliente_opp == "0010067544") ? "0001" : "0002"; //0001 new 0002 exist
                    campos_opp["grupo_opp"] = grupo_opp;

                    campos_opp["Cliente"] = cliente_opp.PadLeft(10, '0'); // "0010004721"; 
                    campos_opp["Contacto"] = contacto_opp.PadLeft(10, '0');  // "0070012034";// contacto_opp;
                    string salesRep = lpsql.getUserEmail(sales_rep);
                    campos_opp["Usuario"] = salesRep;  //"AA70000134"; // sales_rep;

                    campos_opp["OrgVentas"] = "O 50000142"; //Panama
                    campos_opp["OrgServicios"] = "50003612"; //Panama Service Delivery

                    string id_opp = paBids.oppCreate(campos_opp);

                    if (!String.IsNullOrEmpty(id_opp))
                    {
                        if (id_opp.Contains("Error"))
                        {
                            validar_lineas = false;
                            resp_sql = "Error al crear la opp";
                        }
                        else
                        {
                            bool opp_update = lpsql.OppUpdate(singleOrderRecord, id_opp);


                            if (opp_update == false)
                            {
                                JArray j_copias = JArray.Parse(lpsql.getEmail("LICPA"));
                                string jmail = j_copias[0]["email"].ToString();
                                wt.SendNotification(jmail, "Nueva Oportunidad Creada", "Se le notifica que se ha creado una nueva oportunidad con el id: **" + id_opp + "** del registro: **" + singleOrderRecord + "** no se pudo actualizar en la base de datos.");

                            }

                            log.LogDeCambios("Creacion", "Creacion de Oportunidad", "Ventas Panama", "Oportunidad de convenio GBPA - Orden Publicada", singleOrderRecord, id_opp);
                            respFinal = respFinal + "\\n" + "Agregar info a convenio GBPA - Orden Aprobada: " + singleOrderRecord + " " + id_opp;

                            if (cliente_opp == "0010067544") //no encontro con la unidad entonces busca por entidad
                            {
                                JArray j_copias = JArray.Parse(lpsql.getEmail("LICPA"));
                                string jmail = j_copias[0]["email"].ToString();
                                wt.SendNotification(jmail, "Nueva Oportunidad Creada", "Se le notifica que se ha creado una nueva oportunidad con el id: **" + id_opp + "** cuya entidad/Unidad Solicitante/email: **" + entidad + "** no se encuentra creada en SAP, <br><br> Haga click en el siguiente enlace: <a href =\"https://databot.gbm.net\">Portal de Datos Maestros</a> para crear el cliente y una vez creado agreguelo en la <a href=\"https://smartsimple.gbm.net/admin/PanamaBids/entities\">Tabla de Mantenimiento de entidades</a>");

                            }
                        }




                    }
                }
                #endregion


            }
            catch (Exception ex)
            {
                //dio error sacando la info con selenium
                validar_lineas = false;
            }

            #region cerrar chrome
            console.WriteLine("  Finalizando");
            try { chrome.Close(); } catch (Exception) { }

            proc.KillProcess("chromedriver", true);
            proc.KillProcess("chrome", true);
            #endregion

            #region subir archivos FTP

            Array.Resize(ref adjunto, adjunto.Length - 1);
            for (int i = 0; i < adjunto.Length; i++)
            {

                string user = "";
                if (ssMandante == "QAS")
                {
                    user = cred.QA_SS_APP_SERVER_USER;
                }
                else if (ssMandante == "PRD")
                {
                    user = cred.PRD_SS_APP_SERVER_USER;
                }
                bool subir_files = dBase.uploadSftp(adjunto[i].ToString(), $"/home/{user}/projects/smartsimple/gbm-hub-api/src/assets/files/PanamaBids/Agreement", $"Request #{singleOrderRecord}");
                sharep.UploadFileToSharePoint("https://gbmcorp.sharepoint.com/sites/licitaciones_panama", adjunto[i].ToString());
                //llenar tabla
                lpsql.insertFile(adjunto[i].ToString().Split(new string[] { @"downloads\" }, StringSplitOptions.None)[1], singleOrderRecord, "agreement");
            }

            #endregion

            //dio error en la opp
            if (validar_lineas == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + singleOrderRecord, cc);

            }
            //si da error
            if (resp_add_sql == false)
            {
                string[] cc = { "dmeza@gbm.net" };
                mail.SendHTMLMail("Dio error al intentar adjuntar las siguientes Ordenes a la base de datos: " + "<br>" + resp_sql, new string[] {"appmanagement@gbm.net"}, "Error: email de ordenes nuevas publicadas por Convenio Marco, registro: " + singleOrderRecord, cc);

            }
            else
            {
                #region Enviar notificación

                //todo salio bien, se envia notificación.
                JArray j_copias = JArray.Parse(lpsql.getEmail("LICPA"));

                for (int i = 0; i < j_copias.Count; i++)
                {
                    string email = j_copias[i]["email"].ToString();
                    wt.SendNotification(email, "Nueva Orden de Compra Refrendada", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + singleOrderRecord + "**, la cual se encuentra Refrendado<br><br>Haga click en el siguiente enlace: <a href=\"https://smartsimple.gbm.net/admin/PanamaBids/agreement\">Orden de Compra - Portal GBM</a> para ver el documento");
                }

                if (sectorf == "BF")
                {
                    JArray JBF = JArray.Parse(lpsql.getEmail("LPBF"));
                    string bfemail = JBF[0]["email"].ToString();
                    wt.SendNotification(bfemail, "Nueva Orden de Compra Refrendada", "Se le notifica que el (la) **" + entidad + "** ha generado la orden de compra No. **" + singleOrderRecord + "**, la cual se encuentra Refrendado<br><br>Haga click en el siguiente enlace: <a href=\"https://smartsimple.gbm.net/admin/PanamaBids/agreement\">Orden de Compra - Portal GBM</a> para ver el documento");
                }
                #endregion

                #region enviar correo a cliente
                string cuentaEmail = excelResult.Rows[0]["Email Cuenta"].ToString();
                string contactoEmail = excelResult.Rows[0]["Email del Contacto"].ToString();
                string lugar = excelResult.Rows[0]["Lugar de Entrega"].ToString();
                string telf = excelResult.Rows[0]["Telefono del Contacto"].ToString();
                string contacto = excelResult.Rows[0]["Contacto Cuenta"].ToString();
                string documentLink = excelResult.Rows[0]["Link al documento"].ToString();
                string fechaMax = excelResult.Rows[0]["Fecha Maxima de Entrega"].ToString();

                StringBuilder strHTMLBuilder = new StringBuilder();
                strHTMLBuilder.Append($"Estimado cliente {entidad}");
                strHTMLBuilder.Append($"<br><br>");
                strHTMLBuilder.Append($"Hemos recibido su Orden de Compra correspondiente al {singleOrderRecord} de compras en el portal de Convenio Marco de “Equipos Informáticos y Telecomunicaciones”.");
                strHTMLBuilder.Append($"<br>");
                strHTMLBuilder.Append($"El estatus de su orden es “En Proceso”, en cuanto esté facturada estaremos coordinando con usted la entrega por esta misma vía.");
                strHTMLBuilder.Append($"<br><br>");
                strHTMLBuilder.Append($"Confirmamos los datos de entrega:");
                strHTMLBuilder.Append($"<br>");
                strHTMLBuilder.Append($"<ul>");
                strHTMLBuilder.Append($"<li>");
                strHTMLBuilder.Append($"<b>Fecha máxima de entrega:</b> {fechaMax}");
                strHTMLBuilder.Append($"</li>");
                strHTMLBuilder.Append($"<li>");
                strHTMLBuilder.Append($"<b>Lugar de Entrega:</b> {lugar}");
                strHTMLBuilder.Append($"</li>");
                strHTMLBuilder.Append($"<li>");
                strHTMLBuilder.Append($"<b>Contacto de Entrega:</b> {contacto}");
                strHTMLBuilder.Append($"</li>");
                strHTMLBuilder.Append($"<li>");
                strHTMLBuilder.Append($"<b>Teléfono:</b> {telf}");
                strHTMLBuilder.Append($"</li>");
                strHTMLBuilder.Append($"<li>");
                strHTMLBuilder.Append($"<b>e-mail:</b> {contactoEmail}");
                strHTMLBuilder.Append($"</li>");
                strHTMLBuilder.Append($"</ul>");
                strHTMLBuilder.Append($"<br>");
                strHTMLBuilder.Append($"<b>link al documento:</b> {documentLink}");
                strHTMLBuilder.Append($"</li>");
                strHTMLBuilder.Append($"<br><br>");
                strHTMLBuilder.Append($"Gracias por preferir a GBM como su proveedor #1 de tecnología, cualquier consulta puede comunicarse con:");
                strHTMLBuilder.Append($"<br><br>");
                strHTMLBuilder.Append($"Margarita Amburg<br>");
                strHTMLBuilder.Append($"Account Manager<br>");
                strHTMLBuilder.Append($"Public Sector<br>");
                strHTMLBuilder.Append($"GBM Panamá<br>");
                strHTMLBuilder.Append($"e-mail: kvanegas@gbm.net<br>");
                strHTMLBuilder.Append($"T. (507) 300.4808 ext. 7198 | F. (507) 300.4899 | C. (507) 6253-6072<br>");
                strHTMLBuilder.Append($"Business Park, Boulevard Costa del Este, Torre Sur Piso 2.");
                strHTMLBuilder.Append($"<br><br><p style=\"color: grey !important; font-size: 12px !important;\">*Por favor no responda este correo, comuniquese con la persona encargada de GBM</p>");
                string Htmltext = strHTMLBuilder.ToString();

                subject = $"GBM – Orden de Compra recibida - {singleOrderRecord}";

                string html = Properties.Resources.emailtemplate1;
                html = html.Replace("{subject}", "");
                html = html.Replace("{cuerpo}", Htmltext);
                html = html.Replace("{contenido}", "");
                html = html.Replace("<td align=\"center\" style=\"color: #888888; font-size: 16px; font-family: 'Work Sans', Calibri, sans-serif; line-height: 24px;\">", " <td style=\"color: #888888; font-size: 16px; font-family: 'Work Sans', Calibri, sans-serif; line-height: 24px;\">");

                mail.SendHTMLMail(html, new string[] { cuentaEmail, contactoEmail }, subject, new string[] { "kvanegas@gbm.net" }, adjunto);
                #endregion
            }

            root.requestDetails = respFinal;

        }


    }
}
