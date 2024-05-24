using DataBotV5.Data.Process;
using DataBotV5.Data.Projects.CrBids;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using DataBotV5.Logical.Projects.CrBids;
using DataBotV5.Logical.Webex;
using DataBotV5.Logical.Mail;
using System.Globalization;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using ClosedXML.Excel;
using System.IO;

namespace DataBotV5.Automation.MASS.CrBids
{
    /// <summary>
    ///Clase MASS Automation "Robot 2" encargada de la notificación de concursos en SICOP de los cuales los AccountManager no ha indicado si participa o no en el mismo,
    ///y dependiendo si ha superado el tercio de tiempo de vencimiento, notifica a los gerentes, o solo envia una notificación en WebexTeams.
    /// Corre en las mañanas, extrae la base de datos de costa_rica_bids en SmartAndSimple.
    /// </summary>
    class DueBidSS
    {
        string enviroment = "QAS";
        ProcessAdmin padmin = new ProcessAdmin();
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        BidsGbCrSql lcsql = new BidsGbCrSql();
        CrBidsLogical cr_licitaciones = new CrBidsLogical();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        WebexTeams wt = new WebexTeams();
        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

        Log log = new Log();

        string respFinal = "";



        internal CrBidsLogical Cr_licitaciones { get => cr_licitaciones; set => cr_licitaciones = value; }
        internal Rooting Root { get => root; set => root = value; }
        internal Stats Estadisticas { get => estadisticas; set => estadisticas = value; }
        internal ProcessInteraction Proc { get => proc; set => proc = value; }
        internal BidsGbCrSql Lcsql { get => lcsql; set => lcsql = value; }


        CRUD crud = new CRUD();

        /// <summary>
        /// Robot 2: corre en las mañanas, extrae la base de datos de las licitaciones (concursos) y notifica aquellos que no se haya indicado si/no en la participación
        /// </summary>
        public void Main()
        {
            console.WriteLine("Procesando Notificando Licitaciones sin participación");
            Notify();

        }

        /// <summary>
        /// La idea de este robot es notificar a los Account Manager o Gerentes que alguna licitación no ha sido marcada en la sesión de 
        /// Participa de PurchaseOrderAddtionalData como "SI" o "NO", por lo cual ha sido ignorada y a nivel de GBM no es posible dejar escapar 
        /// ninguna licitación que le pueda interesar.
        /// Tiene reglas específicas:
        /// -Si fecha actual es mayor al resultado de la fecha generada sobre de la regla de Tercios, se notifica a  los gerentes.
        /// -Si la fecha actual no sobrepasa a la fecha generada sobre de la regla de Tercio, solo envia una notificación a los AM. 
        /// -En cualquiera de los casos envia un reporte en excel al DMO Líder-licitacionescr@gbm.net con todas las licitaciones no respondidas.
        /// </summary>
        public void Notify()
        {

            #region Extracción de información de BD.
            //Extrae sólo las PO donde el campo Participation sea '' .
            string sqlPurchaseOrders = "SELECT * FROM `purchaseOrder` WHERE id in (SELECT bidNumber FROM `purchaseOrderAdditionalData` WHERE participation = ' ')";
            DataTable purchaseOrders = crud.Select( sqlPurchaseOrders, "costa_rica_bids_db");

            //Extrae las PurchaseOrderAdditionalData.
            string sqlPurchaseAdditionalData = "SELECT * FROM `purchaseOrderAdditionalData` WHERE participation = ' '";
            DataTable POAdditionalData = crud.Select( sqlPurchaseAdditionalData, "costa_rica_bids_db");

            //Extrae las entidades.
            string sqlEntities = "SELECT * FROM `institutions`";
            DataTable entities = crud.Select( sqlEntities, "costa_rica_bids_db");

            //Extrae los colaboradores.
            string slqEmployees = "SELECT * FROM `digital_sign`";
            DataTable employees = crud.Select(slqEmployees, "MIS"); //Siempre en PRD, porque Databot no tiene QA.

            #endregion

            #region Variables y diccionarios locales.

            //Diccionario para guardar los mensajes de cada AccountManager que cumplan con los criterios para wteams.
            Dictionary<string, string> AmMsj = new Dictionary<string, string>();
            //Diccionaro para guardar los msj de cada AccountManager para email .
            Dictionary<string, string> AmEmail = new Dictionary<string, string>();
            //Diccionario para guardar los msj para los Gerentes.
            Dictionary<string, string[]> MaEmail = new Dictionary<string, string[]>();
            //Guardar los value team de los Managers.
            Dictionary<string, string> ValueTeamDictionary = new Dictionary<string, string>();

            #endregion

            #region Creación de archivo de Excel y su estructura de celdas.



            DataTable excelResult = new DataTable();
            excelResult.Columns.Add("Número de Concurso");
            excelResult.Columns.Add("Descripción");
            excelResult.Columns.Add("Entidad");
            excelResult.Columns.Add("Presupuesto");
            excelResult.Columns.Add("Fecha Publicación");
            excelResult.Columns.Add("Fecha Límite para aclaraciones");
            excelResult.Columns.Add("Fecha de Apertura de Ofertas");
            excelResult.Columns.Add("Account Manager");


            #endregion

            if (purchaseOrders.Rows.Count > 0)
            {
                #region Verifica las purchase sin participar y llena los mensajes del diccionario AmMsj
                console.WriteLine("Inicio de proceso de validación si existen Bids superado el tercio de tiempo, y la participación que no haya sido indicada por los AM. ");
                //Todas las PO cumplen con que el AM no le ha dado "SI" al participation(gracias al select), aqui se verifica si supero el tercio de tiempo y guardar mensajes en diccionarios
                for (int i = 0; i < purchaseOrders.Rows.Count; i++)
                {
                    if ((i + 1) == 101)
                    {
                        console.WriteLine("Bid con duda");
                    }
                    try
                    {
                        string idPurchaseOrder = purchaseOrders.Rows[i]["id"].ToString();

                        string bidNumber = purchaseOrders.Rows[i]["bidNumber"].ToString();
                        string description = purchaseOrders.Rows[i]["description"].ToString();
                        string institution = purchaseOrders.Rows[i]["institution"].ToString();
                        string budget = purchaseOrders.Rows[i]["budget"].ToString(); //Presupuesto.

                        //Se pone [0], ya que es un Datarow[] y se selecciona la primer posición (aunque siempre devuelve solo una posición).
                        string participa = POAdditionalData.Select("bidNumber =" + idPurchaseOrder)[0]["participation"].ToString();
                        string valueTeam = POAdditionalData.Select("bidNumber =" + idPurchaseOrder)[0]["valueTeam"].ToString();
                        string managerSector = POAdditionalData.Select("bidNumber =" + idPurchaseOrder)[0]["managerSector"].ToString(); //gerente_sector


                        console.WriteLine("Validando Bid #" + (i + 1) + " , de " + (purchaseOrders.Rows.Count) + ". Concurso: " + bidNumber + ", institution: " + institution + ".");

                        //Enviar  notificación si no se ha indicado status de participación.
                        if (participa == " ")
                        {
                            //buscar el AM de la entidad

                            string AM = "";
                            string AM_name = "";
                            AM = POAdditionalData.Select("bidNumber =" + idPurchaseOrder)[0]["accountManager"].ToString();
                            if (AM == "")
                            {
                                console.WriteLine("Esta licitación no posee un Account Manager asignado.");
                                continue; //Se sale de la iteración del for actual.
                            }
                            try
                            {
                                //Extrae el Datarow con el nombre de empleado que es igual a AM user.
                                System.Data.DataRow[] enmp_info = employees.Select("user ='" + AM + "'");
                                AM_name = enmp_info[0]["name"].ToString();

                            }
                            catch (Exception)
                            {

                            }

                            string pdate = purchaseOrders.Rows[i]["publicationDate"].ToString();
                            string adate = purchaseOrders.Rows[i]["receptionClarification"].ToString();
                            string odate = purchaseOrders.Rows[i]["offerOpening"].ToString();


                            string msjact = "";
                            string mensaje = "- **" + institution + "**: [" + bidNumber + "](" + "https://smartsimple.gbm.net/" + ")" + " - " + description + ". Inicia: " + pdate + " al " + adate + "\n";

                            string mensaje_email = "";
                            mensaje_email = mensaje_email + "<tr>";
                            mensaje_email = mensaje_email + "<td>" + institution + "</td>";
                            mensaje_email = mensaje_email + "<td>" + "<a href=\"https://smartsimple.gbm.net/\" > " + bidNumber + "</a>" + "</td>";
                            mensaje_email = mensaje_email + "<td>" + description + "</td>";
                            mensaje_email = mensaje_email + "<td>" + pdate + "</td>";
                            mensaje_email = mensaje_email + "<td>" + adate + "</td>";
                            mensaje_email = mensaje_email + "<td>" + budget + "</td>";
                            //mensaje_email = mensaje_email + "<td>" + AM_name + "</td>";
                            mensaje_email = mensaje_email + "</tr>";

                            #region calculo del tercio

                            DateTime.TryParse(pdate, out DateTime fecha_publi);
                            DateTime.TryParse(adate, out DateTime fecha_aclar);
                            DateTime.TryParse(odate, out DateTime fecha_apertura);

                            TimeSpan vt = fecha_aclar.Subtract(fecha_publi);
                            TimeSpan dias_apertura = fecha_apertura.Subtract(fecha_publi);

                            double venc_t = ((fecha_aclar - fecha_publi).TotalDays) / 2; //50% del tercio
                            double apert_d = (fecha_apertura - fecha_publi).TotalDays; //???

                            DateTime vencimiento_tercio = fecha_publi.AddDays(venc_t);
                            DateTime fecha_apert = fecha_publi.AddDays(apert_d);


                            //  Reglas: antes de n-2 de cumplimiento del tercio
                            //  "Faltan n-2" días para que supere el tercio, se notifica al Gerente
                            //  Presupuesto > 100k ??

                            #endregion

                            //si no se sobrepasa con el tercio se agrega la licitación al mensaje por email para el AM y se reporta el gerente
                            if (DateTime.Today >= vencimiento_tercio)
                            {
                                console.WriteLine("Superó el tercio.");
                                //es para agregar el mensaje anterior en el diccionario del email que se le envia al AM
                                try
                                {
                                    //extrae el mensaje que tenga actualmente en el diccionario
                                    msjact = AmEmail[AM];
                                    //y lo vuelve a ingresar concatenado con el nuevo mensaje
                                    AmEmail[AM] = msjact + mensaje_email;
                                }
                                catch (Exception)
                                {
                                    //cae en catch cuando el AM no esta en el diccionario
                                    //por lo que se llena el diccionario con el primero mensaje
                                    AmEmail[AM] = mensaje_email;
                                }

                                #region notificar a gerencia
                                ValueTeamDictionary[managerSector] = valueTeam;

                                try
                                {
                                    //extrae el mensaje que tenga actualmente en el diccionario
                                    string[] AMarray = MaEmail[managerSector];
                                    //Verifica si ya el AM esta en el array del manager
                                    if (!AMarray.Any(AM.Contains))
                                    {
                                        Array.Resize(ref AMarray, AMarray.Length + 1);
                                        AMarray[AMarray.Length - 1] = AM;
                                    }

                                    //y lo vuelve a ingresar concatenado con el nuevo mensaje
                                    MaEmail[managerSector] = AMarray;
                                }
                                catch (Exception)
                                {
                                    //cae en catch cuando el Manager no esta en el diccionario
                                    //por lo que se llena el diccionario con el primero mensaje
                                    string[] AMarray = { AM };
                                    MaEmail[managerSector] = AMarray;
                                }
                                #endregion


                            }
                            else
                            {
                                console.WriteLine("No ha superado el tercio");
                                try
                                {
                                    //extrae el mensaje que tenga actualmente en el diccionario del webex teams
                                    msjact = AmMsj[AM];
                                    //y lo vuelve a ingresar concatenado con el nuevo mensaje
                                    AmMsj[AM] = msjact + mensaje;
                                }
                                catch (Exception)
                                {
                                    //cae en catch cuando el AM no esta en el diccionario
                                    //por lo que se llena el diccionario con el primero mensaje
                                    AmMsj[AM] = mensaje;
                                }
                            }

                            #region Guardar en excel

                            string user = AM.ToUpper();

                            DataRow rRow = excelResult.Rows.Add();
                            rRow["Número de Concurso"] = bidNumber;
                            rRow["Descripción"] = description;
                            rRow["Entidad"] = institution;
                            rRow["Presupuesto"] = budget;
                            rRow["Fecha Publicación"] = fecha_publi;
                            if (!string.IsNullOrEmpty(adate))
                            {
                                rRow["Fecha Límite para aclaraciones"] = fecha_aclar;
                            }
                            if (!string.IsNullOrEmpty(odate))
                            {
                                rRow["Fecha de Apertura de Ofertas"] = fecha_apertura;
                            }
                            rRow["Account Manager"] = user.ToUpper();
                            excelResult.AcceptChanges();

                            log.LogDeCambios("Modificación", root.BDProcess, user, $"Notificación de concurso {bidNumber} sin participación", bidNumber, description);
                            respFinal = respFinal + "\\n" + $"Notificación de concurso {bidNumber} sin participación. " ;

                            #endregion


                        } //no ha dicho si participa o no
                        else
                        {
                            console.WriteLine("next");
                        }




                    }
                    catch (Exception ex)
                    {
                        console.WriteLine(ex.ToString());
                    }

                } //for concursos

                console.WriteLine("Se ha finalizado el proceso de validación de concursos vencidos.");
                #endregion

                #region Enviar notificaciones  a los AM
                console.WriteLine("Se inicia el proceso de notificación a los Account Managers.");

                string header = "Se le notifica que **no** ha indicado si participa en los siguientes concursos:\r\n ";
                string footer = " \r\n Por favor indicar su participación haciendo click [aquí](" + "https://smartsimple.gbm.net/" + ")";
                //Se notifica por Webex Teams al AM que NO haya superado el tercio de tiempo, y no se notifica al gerente.
                //se envia por teams todas las que no han sido respondidas y se guardaron en el diccionario am_msj en el FOR anterior. 
                foreach (KeyValuePair<string, string> pair in AmMsj)
                {
                    string AM = pair.Key.ToString();
                    string mensaje = pair.Value.ToString();

                    wt.SendNotification(AM + "@gbm.net", "No participación", header + mensaje + footer);
                }
                //for the email
                header = "Se le notifica que <b>no</b> ha indicado si participa en los siguientes concursos y ya se cumplio el 50% del tercio. Por favor indicar su participación haciendo click en el botón abajo. ";
                footer = "<br>Por favor indicar su participación haciendo click <a href=\"https://smartsimple.gbm.net/\" > aqui</a>";
                //se envia por email todas aquellas que se este al 50% del vencimiento del tercio 
                foreach (KeyValuePair<string, string> pair in AmEmail)
                {
                    string AM = pair.Key.ToString();
                    string body = "";

                    body = body + "<table class='myCustomTable' width='100 %'>";
                    body = body + "<thead><tr><th>Institución</th><th>Concurso</th><th style='width:50%;'>Descripción</th><th>Fecha de publicación</th><th>Fecha de aclaración</th><th>Presupuesto</th></tr></thead>";
                    body = body + "<tbody>";
                    body = body + pair.Value.ToString(); //+ "<br><br>"
                    body = body + "</tbody>";
                    body = body + "</table>";

                    string am_email2 = AM + "@gbm.net";
                    string user = AM;

                    string htmlpage = Properties.Resources.emailLpCr;
                    htmlpage = htmlpage.Replace("{subject}", "Notificación de Concursos sin participación SICOP");
                    htmlpage = htmlpage.Replace("{cuerpo}", header);
                    htmlpage = htmlpage.Replace("{contenido}", body);

                    mail.SendHTMLMail(htmlpage, new string[] { am_email2 }, "Concursos sin participación - Licitaciones Costa Rica - " + user.ToUpper(), null);

                }
                #endregion

                #region Enviar notificaciones a Gerencia
                console.WriteLine("Se inicia el proceso de notificación a Gerencia.");
                header = "Se le notifica que los siguiente Account Manager <b>no</b> han indicado si participa en los siguientes concursos, y ya se cumplio el 50% del tercio:<br>";
                foreach (KeyValuePair<string, string[]> pair in MaEmail)
                {
                    string manager = pair.Key.ToString();
                    string valueteam = ValueTeamDictionary[manager];
                    string body = "";

                    string[] AMarrray = pair.Value;
                    for (int i = 0; i < AMarrray.Length; i++)
                    {
                        string user = "";
                        try
                        {
                            System.Data.DataRow[] empl_info = employees.Select("user ='" + AMarrray[i].ToString().ToUpper() + "'"); //like '%" + institu + "%'"
                            user = myTI.ToTitleCase(empl_info[0]["name"].ToString());
                        }
                        catch (Exception)
                        {
                            user = AMarrray[i].ToString().ToUpper();
                        }
                        body = body + "<b>" + user + "</b><br>";
                        body = body + "<table class='myCustomTable' width='100 %'>";
                        body = body + "<thead><tr><th>Institución</th><th>Concurso</th><th style='width:50%;'>Descripción</th><th>Fecha de publicación</th><th>Fecha de aclaración</th><th>Presupuesto</th></tr></thead>";
                        body = body + "<tbody>";
                        body = body + AmEmail[AMarrray[i].ToString()]; //+ "<br><br>"
                        body = body + "</tbody>";
                        body = body + "</table>";
                        body = body + "<br><br>";
                    }

                    string htmlemail = Properties.Resources.emailLpCr;
                    htmlemail = htmlemail.Replace("{subject}", "Notificación de Concursos sin participación SICOP");
                    htmlemail = htmlemail.Replace("{cuerpo}", header);
                    htmlemail = htmlemail.Replace("{contenido}", body);

                    mail.SendHTMLMail(htmlemail, new string[] { manager + "@gbm.net" }, "Concursos sin participación - Licitaciones Costa Rica - " + valueteam, null);
                }
                #endregion

                #region Enviar Reporte al DMO
                console.WriteLine("Se inicia el proceso de notificación a DMO con el respectivo excel.");
                string mes = (DateTime.Now.Month.ToString().Length == 1) ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                string dia = (DateTime.Now.Day.ToString().Length == 1) ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                string fecha_file = dia + "_" + mes + "_" + DateTime.Now.Year.ToString();
                string nombre_file = Root.FilesDownloadPath + "\\" + "Concursos sin participar - " + fecha_file + ".xlsx"; //???

                excelResult.AcceptChanges();

                #region Guardado Excel
                console.WriteLine("Save Excel...");
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(excelResult, "Concursos sin participar");
                wb.Worksheet("Concursos sin participar").Columns().AdjustToContents(); //Es exacto a la función autofit del excel antiguo 
                wb.Worksheet("Concursos sin participar").Columns("B:B").Width = 75; //Ancho del campo descripción en excel. 

                if (File.Exists(nombre_file))
                {
                    File.Delete(nombre_file);
                }
                wb.SaveAs(nombre_file);

                #endregion
                string[] cc = { "" };
                string[] adj = { nombre_file };
                string sub = "Concursos sin participar - Licitaciones de Costa Rica - " + fecha_file;
                string msj = "A continuación, se adjunta el archivo de las licitaciones sin participación de GBM Costa Rica a la fecha " + DateTime.Today.ToString();
                //DMO
                string dmoemail = "";
                try
                {
                    DataTable DMO = crud.Select( "SELECT * FROM `emailAddress` WHERE `category` LIKE 'DMOLIDER'", "costa_rica_bids_db");
                    JObject row = JObject.Parse(DMO.Rows[0]["jemail"].ToString().Trim());
                    dmoemail = row["email"].Value<string>();
                }
                catch (Exception)
                {
                    dmoemail = "appmanagement@gbm.net";
                }
                string html = Properties.Resources.emailLpCr;
                html = html.Replace("{subject}", "Notificación de Concursos sin participación SICOP");
                html = html.Replace("{cuerpo}", msj);
                html = html.Replace("{contenido}", "");

                mail.SendHTMLMail(html, new string[] { dmoemail }, sub, null, adj);
                #endregion

                root.requestDetails = respFinal;
                root.BDUserCreatedBy = dmoemail;


                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

    }
}
