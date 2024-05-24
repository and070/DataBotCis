using System;
using System.Data;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Database;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.CrBids;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.RPA2.CrBids
{
    /// <summary>
    /// Clase RPA "Robot 5" Automation encargada de la modificación de licitación de un concurso en SICOP.
    /// </summary>
    class ModifyBidSS
    {

        string enviroment = "QAS";
        CrBidsLogical cr_licitaciones = new CrBidsLogical();
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        MailInteraction mail = new MailInteraction();
        BidsGbCrSql liccr = new BidsGbCrSql();
        Log log = new Log();

        string respFinal = "";
        bool executeStats = false;


        public void Main()
        {

            if (mail.GetAttachmentEmail("Solicitudes Sicop", "Procesados", "Procesados Sicop"))
            {
                console.WriteLine(" Procesando....");
                ProcessToModifyBidNew(root.Email_Body);

                if (executeStats == true)
                {
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }

        /// <summary>
        /// Este método se encarga de modificar una licitación cuando sicop envia un correo electronico al Databot indicando que una licitacion existente
        /// ligada a GBM cambió, esto funciona ya que el bot lee la carpeta de solicitudes Sicop,  y cuando aparece un correo bajo el subject 
        /// de [External]Modificación del concurso, lo mueve a procesados, y extrae el body para descargar los nuevos datos y nueva información, y 
        /// actualizar- informar al Account Manager responsable del mismo, y si anteriormente se había indicado que GBM Participa entonces se le informa al SalesTeam también.
        /// Otros correos que lleguen a esa carpeta con diferente nombre de subject, solo se reenvia al Account Manager responsable.
        /// </summary>
        /// <param name="body"></param>
        private void ProcessToModifyBidNew(string body)
        {
            console.WriteLine("Tipo de subject a validar: " + root.Subject);
            Regex reg;
            reg = new Regex("[*'\"_&+^><@]");
            string body_clean = reg.Replace(body, string.Empty);
            string[] separa0 = new string[1];
            string separa1 = "";
            console.WriteLine(" Extrayendo info del Body y Subject");
            if (root.Subject == "[External]Modificación del concurso") { separa0[0] = "número "; separa1 = " para "; }
            else if (root.Subject == "[External]Registro de anuncio/aclaración al cartel") { separa0[0] = "número "; separa1 = "\r\npara "; }
            else if (root.Subject == "[External]Notificación de recepción definitiva" || root.Subject == "[External]Solicitud de subsanación") { separa0[0] = "Número "; separa1 = " para "; }
            else if (root.Subject == "[External]Respuesta de aclaración al cartel solicitada") { separa0[0] = "concurso No. "; separa1 = " para "; }
            else if (root.Subject.Contains("[External]Notificación de vencimiento de plazo de Garantía de participación número")) { separa0[0] = "concurso Número "; separa1 = " Correspondiente a "; }
            else if (root.Subject.Contains("[External]Comunicación acto final")) { separa0[0] = "concurso Número "; separa1 = " para "; }
            else if (root.Subject.Contains("[External]Invitación a concursar")) { return; }


            string concurso = "";
            try
            {
                string[] link = body_clean.Split(separa0, StringSplitOptions.None);
                concurso = link[1].ToString().Trim();
                int limite = concurso.IndexOf(separa1);
                if (concurso.Length >= limite)
                { concurso = concurso.Substring(0, limite).Trim(); }
            }
            catch (Exception)
            {
                if (root.Subject == "[External]Respuesta de aclaración al cartel solicitada")
                {
                    separa0[0] = "concurso número "; separa1 = ", para ";
                    string[] link = body_clean.Split(separa0, StringSplitOptions.None);
                    concurso = link[1].ToString().Trim();
                    int limite = concurso.IndexOf(separa1);
                    if (concurso.Length >= limite)
                    { concurso = concurso.Substring(0, limite).Trim(); }
                }
            }


            //extraer la info correspondiente de la BD

            console.WriteLine("Extrayendo información de base de datos...");
            #region Extracción de información de Base de Datos
            string sqlConcursoInfo = $"SELECT* FROM `purchaseOrder` WHERE `bidNumber` LIKE '{concurso}'";
            DataTable PurchaseOrder = crud.Select(sqlConcursoInfo, "costa_rica_bids_db");

            string sqlPOAdditionalData = $"SELECT * FROM `purchaseOrderAdditionalData` WHERE bidNumber in (select Id FROM purchaseOrder WHERE bidNumber = '{concurso}')";
            DataTable POAdditionalData = crud.Select(sqlPOAdditionalData, "costa_rica_bids_db");

            //Extraer todos los usuarios del SalesTeam enlazados a la purchaseOrderActual
            string sqlSalesTeam = $"SELECT * from salesTeam where bidNumber in (SELECT Id FROM `purchaseOrder` WHERE bidNumber='{concurso}')";
            DataTable SalesTeamDataTable = crud.Select(sqlSalesTeam, "costa_rica_bids_db");

            #endregion

            if (PurchaseOrder.Rows.Count <= 0)
            {
                //si el concurso no se encuentra en la BD significa que se movio al backup debido a su NO participación o bien por el robot4
                //o bien no se ha guardado el concurso en la BD
                console.WriteLine("El concurso no se encuentra en la BD, significa que se movió a backup y no se encuentra activa, por tanto no se procesa.");
                return;
            }

            console.WriteLine($"Licitación a procesar: id: {PurchaseOrder.Rows[0]["id"]}, BidNumber #{concurso}, Institution: {PurchaseOrder.Rows[0]["institution"]}, " +
                    $"Description: {PurchaseOrder.Rows[0]["description"]}");

            if (POAdditionalData.Rows.Count <= 0)
            {
                //Ocurrió un error ya que la PO existe pero no tiene POAddtionalData, por tanto se notifica a los admin para el mapeo de porque no tiene, y no se puede procesar la modificación.
                console.WriteLine($"Ocurrió un error ya que la licitación BidNumber #{concurso}, existe pero no tiene POAddtionalData, " +
                    "por tanto se notifica a los AppManagement y desarrolladores para el mapeo de porque no tiene, y no se puede procesar la modificación. ");
                console.WriteLine($"Datos de la licitación: Id: {PurchaseOrder.Rows[0]["id"]}, Institution: {PurchaseOrder.Rows[0]["institution"]}, Description: {PurchaseOrder.Rows[0]["description"]}.");

                string msg = $"El error ocurrió en la clase ModifyBidSS del proyecto CrBids, no se pudo modificar la licitación:<br><br> Id: {PurchaseOrder.Rows[0]["id"]}. <br>BidNumber #{concurso}. <br>Institution: {PurchaseOrder.Rows[0]["institution"]}. <br>" +
                    $"Description: {PurchaseOrder.Rows[0]["description"]}.<br><br>Esto debido a que no posee POAdditionalData. Verifica por favor el porqué no tiene POAdditionalData y proceder a notificar al AM sobre su modificación.<br><br>Databot.";
                console.WriteLine(msg);
                mail.SendHTMLMail(msg, new string[] { "appmanagement@gbm.net" }, "No se pudo modificar la licitación - ModifyBidSS - Databot", new string[] { "dmeza@gbm.net", "epiedra@gbm.net" });
                return;
            }


            if (root.Subject == "[External]Modificación del concurso")
            {

                console.WriteLine("Se procede a la modificación del concurso...");
                string descripcion = PurchaseOrder.Rows[0]["description"].ToString();
                string participa = POAdditionalData.Rows[0]["participation"].ToString();
                string interes = POAdditionalData.Rows[0]["gbmStatus"].ToString();//Interés de GBM 
                string vendedor = POAdditionalData.Rows[0]["accountManager"].ToString();
                string aperturaOfertas = PurchaseOrder.Rows[0]["offerOpening"].ToString();

                //entrar a selenium
                //verificar fecha de apertura
                //si cambia con la actual modificar fecha de apertura
                bool newDescargas = (participa == "SI" || interes == "1"/*SI*/) ? true : false;
                console.WriteLine("Participación: " + participa + ", interés de GBM: " + interes);
                ModifyConcourse ConcursoInfo = cr_licitaciones.ModConcourse2(cr_licitaciones.SelConn("https://www.sicop.go.cr/index.jsp"), concurso, newDescargas, enviroment);
                //y hacer loop para nueva documentación en caso de que sea de interes o si participa
                //verificar si en la tabla de archivos  existe el nombre del archivo con relacion al concurso, si no existe ingresar el blop
                if (ConcursoInfo.Val)
                {
                    string cuerpo = $"Estimado {vendedor} se le informa que el concurso {concurso} fue modificado por SICOP";

                    #region ParseDates
                    ConcursoInfo.OpenningDate = DateTime.Parse(ConcursoInfo.OpenningDate).ToString("yyyy-MM-dd HH:mm:ss");
                    aperturaOfertas = (aperturaOfertas == "") ? "" : DateTime.Parse(aperturaOfertas).ToString("yyyy-MM-dd HH:mm:ss");
                    #endregion
                    //Si cambió la fecha
                    if (ConcursoInfo.OpenningDate.ToString() != aperturaOfertas)
                    {
                        //actualizar campo en BD cuando la fecha cambia
                        Dictionary<string, string> dd = new Dictionary<string, string>
                        {
                            ["offerOpening"] =
                        ConcursoInfo.OpenningDate.ToString()
                        };
                        liccr.UpdateRowModifyBid("costa_rica_bids_db", "purchaseOrder", "bidNumber", concurso, dd);
                        cuerpo += $". La nueva fecha de apertura de oferta es: {ConcursoInfo.OpenningDate}";
                        console.WriteLine("Cambió la fecha de apertura de la licitación.");

                        log.LogDeCambios("Modificación", root.BDProcess, "", "Modificación de  concurso: " + concurso + ", la nueva fecha de apertura de oferta es: " + ConcursoInfo.OpenningDate, root.Subject, "");
                        respFinal = respFinal + "\\n" + "Modificación de concurso : " + concurso + ", la nueva fecha de apertura de oferta es: " + ConcursoInfo.OpenningDate;
                        executeStats = true;

                    }
                    //Si existen nuevas descargas.
                    if (ConcursoInfo.NewDownloads)
                    {
                        cuerpo += $". Y se descargó nueva documentación, ingrese <a href='https://smartsimple.gbm.net/' target='_blank'>aquí</a> para ver los nuevos documentos del concurso.";
                        console.WriteLine("Se descargó nueva documentación.");

                        log.LogDeCambios("Modificación", root.BDProcess, "", "Modificación de  concurso: " + concurso + ", se descargó nueva documentación", root.Subject, "");
                        respFinal = respFinal + "\\n" + "Modificación de  concurso: " + concurso + ", se descargó nueva documentación";
                        executeStats = true;
                    }
                    string cleanedText = Regex.Replace(body_clean, @"http[^\s]+", "");
                    string html = Properties.Resources.emailtemplate1;
                    html = html.Replace("{subject}", $"Modificación del concurso: {concurso}");
                    html = html.Replace("{cuerpo}", cuerpo);
                    html = html.Replace("{contenido}", cleanedText);
                    string[] sales_team;
                    if (participa == "SI") //Si participa se le notifica al SalesTeam
                    {
                        #region crear el sales team
                        sales_team = (SalesTeamDataTable.Rows.Count == 0) ? null : new string[SalesTeamDataTable.Rows.Count];

                        for (int i = 0; i < SalesTeamDataTable.Rows.Count; i++)
                        {
                            sales_team[i] = SalesTeamDataTable.Rows[i]["salesTeam"].ToString() + "@gbm.net";
                        }
                        #endregion

                        console.WriteLine("Si participa en la licitación, por tanto se le envia correo electrónico al AM: " + vendedor + " y a su SalesTeam en copia.");

                        mail.SendHTMLMail(html, new string[] { vendedor + "@gbm.net" }, $"Modificación del concurso {concurso} - {descripcion}", sales_team, null);

                    }
                    //else if (string.IsNullOrEmpty(participa))
                    else if (string.IsNullOrEmpty(participa) || participa == " ")
                    { //Solo se le notifica al vendedor ya que no participa.
                        mail.SendHTMLMail(html, new string[] { vendedor + "@gbm.net" }, $"Modificación del concurso {concurso} - {descripcion}", null, null);
                        console.WriteLine($"No se ha indicado si participa o no en la licitación, por tanto se notifica a el AM: {vendedor}, sin copias porque no tiene SalesTeam generado aún.");
                    }
                    else
                    {
                        console.WriteLine($"Se notifica a Diego Meza, ya que indicaron que NO participa o algún otro tipo de caracter.");
                        mail.SendHTMLMail(html, new string[] { "dmeza@gbm.net" }, $"Modificación del concurso {concurso} - {descripcion}", null, null);
                    }
                }
                else
                {
                    //error

                }



            }
            else if (root.Subject == "[External]Respuesta de aclaración al cartel solicitada" ||
                root.Subject.Contains("[External]Notificación de vencimiento de plazo de Garantía de participación número") ||
                root.Subject.Contains("[External]Registro de anuncio/aclaración al cartel") ||
                root.Subject.Contains("[External]Notificación de recepción definitiva") ||
                root.Subject.Contains("[External]Solicitud de subsanación") ||
                root.Subject.Contains("[External]Comunicación acto final"))
            {

                string vendedor = POAdditionalData.Rows[0]["accountManager"].ToString();
                string participa = POAdditionalData.Rows[0]["participation"].ToString();
                string[] sales_team = null;
                if (participa == "SI")
                {
                    #region crear el sales team
                    sales_team = (SalesTeamDataTable.Rows.Count == 0) ? null : new string[SalesTeamDataTable.Rows.Count];

                    for (int i = 0; i < SalesTeamDataTable.Rows.Count; i++)
                    {
                        sales_team[i] = SalesTeamDataTable.Rows[i]["salesTeam"].ToString() + "@gbm.net";
                    }

                    #endregion

                }
                mail.ForwardEmail(root.EmailObject, $"{vendedor}@gbm.net", sales_team, null);

                //Evaluar si habilitar lo siguiente, debido a que es solo notificación.
                //log.LogDeCambios("Notificación", root.BDProcess, "", "Notificación de  concurso: " + concurso, root.Subject, "");
                //respFinal = respFinal + "\\n" + "Notificación de concurso : " + concurso;
                //executeStats = true;
            }
            else
            {
                console.WriteLine($"El tipo de subject es: {root.Subject}, por tanto solo se procede a reenviar el correo a Diego Meza ya que no es modificación.");
                mail.ForwardEmail(root.EmailObject, "dmeza@gbm.net", null, null);
            }



            root.requestDetails = respFinal;

            //Se establece el vendedor debido a que el subject del correo es SICOP y no debería ser así.
            root.BDUserCreatedBy = POAdditionalData.Rows[0]["accountManager"].ToString();


        } //Inicio12

    }

    public class ModifyConcourse
    {
        public bool Val { get; set; }
        public string OpenningDate { get; set; }
        public bool NewDownloads { get; set; }
    }
}
