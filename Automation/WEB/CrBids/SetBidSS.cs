using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Process;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.CrBids;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace DataBotV5.Automation.WEB.CrBids
{
    /// <summary>
    /// Clase WEB "Robot 3" Automation encargada de la participación de concurso SICOP. Se activa cuando el AM indica si participa o no en una licitación, mediante el status del concurso columna extra.
    /// </summary>
    class SetBidSS
    {
        #region variables_globales
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        Log log = new Log();
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        BidsGbCrSql lcsql = new BidsGbCrSql();
        BidsGbCrSql liccr = new BidsGbCrSql();
        ProcessAdmin padmin = new ProcessAdmin();
        CrBidsLogical crBids = new CrBidsLogical();
        WebInteraction web = new WebInteraction();
        WebexTeams wt = new WebexTeams();
        CRUD crud = new CRUD();

        string respFinal = "";


        internal Stats Estadisticas { get => estadisticas; set => estadisticas = value; }
        public Credentials Cred { get => cred; set => cred = value; }
        internal MailInteraction Mail { get => mail; set => mail = value; }
        internal BidsGbCrSql Lcsql { get => lcsql; set => lcsql = value; }
        string mandante = "QAS";
        string mandanteSAP = "CRM";
        #endregion
        /// <summary>
        /// Se activa cuando el campo statusRobot es igual a 1 en la tabla purchaseOrder
        /// De acuerdo al campo participation de la tabla purcasheOrderAddData se realizan 2 proceso
        /// si es NO: manda correos a los gerentes de acuerdo al presupuesto y mueve el bid al backup si NO es de interes de GBM
        /// si es SI: Crea la opp en SAP CRM con la información del bid y notifica al Account Manager y su Sales Teams
        /// </summary>
        public void Main()
        {
            //get_concurso() pendiente por procesar;
            DataTable respuesta = crud.Select("SELECT * FROM purchaseOrder WHERE statusRobot = 1", "costa_rica_bids_db");
            if (respuesta.Rows.Count > 0) //insertar si es que no existe para evitar duplicidad
            {
                console.WriteLine("Procesando...");
                Participate(respuesta);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void Participate(DataTable concursos)
        {
            //DataTable entidades = Lcsql.getallenti();
            string gerente = "";
            string gerente_general = "";
            string sender = "";
            string[] cc = new string[1];
            bool enviar = false;
            bool validar_lineas = true;
            string msj_dev = "";
            string user = "";
            string am_email = "";
            string am_name = "";
            string description = "";
            string budget = "";
            string bidNumber = "";
            string institution = "";
            string tipopp = "";
            string sales_type = "";
            string justificacion = "";
            string participate = "";
            string idBidNumber = "";
            try
            {
                console.WriteLine("Extrayendo información de la base de datos...");
                bidNumber = concursos.Rows[0]["bidNumber"].ToString(); //el numero del concurso
                idBidNumber = concursos.Rows[0]["Id"].ToString(); //el id del concurso
                DataTable empleados = crud.Select("SELECT * FROM `digital_sign`", "MIS"); //tabla de empleados, para busacar los AM
                DataTable purchaseOrderAdditionalData = crud.Select($"SELECT * FROM purchaseOrderAdditionalData WHERE bidNumber = '{idBidNumber}'", "costa_rica_bids_db");
                DataTable salesTeam = crud.Select($"SELECT * FROM salesTeam WHERE bidNumber = '{idBidNumber}'", "costa_rica_bids_db");
                DataTable oppType = crud.Select($"SELECT * FROM `oppType`", "costa_rica_bids_db");
                DataTable emaildefault = crud.Select("SELECT * FROM emailAddress", "costa_rica_bids_db");
                //Newtonsoft.Json.Linq.JObject json = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(concursos.Rows[0]["datos_generales"].ToString()).ToString());
                description = concursos.Rows[0]["description"].ToString();
                budget = concursos.Rows[0]["budget"].ToString();
                institution = concursos.Rows[0]["institution"].ToString();

                //json = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(concursos.Rows[0]["datos_sap"].ToString()).ToString());
                participate = purchaseOrderAdditionalData.Rows[0]["participation"].ToString();
                string accountManager = purchaseOrderAdditionalData.Rows[0]["accountManager"].ToString();
                string gbmInterest = purchaseOrderAdditionalData.Rows[0]["gbmStatus"].ToString();
                #region account manager name

                DataRow[] emp_info = empleados.Select("user ='" + accountManager + "'");
                user = accountManager;
                am_email = user + "@gbm.net";
                if (emp_info.Count() > 0)
                {
                    accountManager = emp_info[0]["UserID"].ToString();
                    am_name = emp_info[0]["name"].ToString();
                }
                else
                {
                    DataRow[] amDefault = emaildefault.Select("category = 'AMDEFAULT'");
                    string jemail = amDefault[0]["jemail"].ToString();
                    accountManager = JObject.Parse(jemail)["id"].Value<string>();
                    am_name = JObject.Parse(jemail)["name"].Value<string>();
                }

                #endregion

                console.WriteLine("Validando la licitación, id: " + idBidNumber + ", bidNumber: " + bidNumber + ", de la insititución: " + institution + ", descripción: " + description);
                if (participate == "NO")
                {
                    console.WriteLine("El account manager indicó que NO participa.");

                    gerente = purchaseOrderAdditionalData.Rows[0]["managerSector"].ToString() + "@gbm.net";
                    //json = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(concursos.Rows[0]["datos_adicionales"].ToString()).ToString());
                    string motivo = purchaseOrderAdditionalData.Rows[0]["noParticipationReason"].ToString(); //concursos.Rows[0]["no_participacion"].ToString();
                    justificacion = purchaseOrderAdditionalData.Rows[0]["notParticipate"].ToString(); //concursos.Rows[0]["motivo_noparticipacion"].ToString();
                    DataTable noParticipationReason = crud.Select($"SELECT noParticipationReason FROM `noParticipationReason` WHERE  id = {motivo}", "costa_rica_bids_db");
                    motivo = noParticipationReason.Rows[0]["noParticipationReason"].ToString();
                    //Convertir presupuesto a int ??
                    string budgetToConsole = budget;
                    budget = budget.Replace(" [CRC]", "");
                    budget = budget.Replace(".", "");
                    //presupuesto = presupuesto.Replace(",", ".");
                    float pres = float.Parse(budget);

                    string header = "Se le notifica que " + am_name + " ha decidido <b>no</b> participar en el Concurso <b>" + bidNumber + "</b> - " + description + ", debido a: " + motivo + "<br><br><b>Motivo de no participación:</b><br>" + justificacion;

                    //Si es mayor a 30 M de Colones se alerta al gerente
                    if (pres > 30000000)
                    {
                        console.WriteLine("Notificando al a gerencia ya que el monto es de: " + budgetToConsole + ".");
                        enviar = true;
                        //Extraer mediante un WS SAP el manager del AM.

                        sender = gerente;
                        cc[0] = am_email;
                        if (pres > 120000000)
                        {
                            console.WriteLine("Notificando al gerente general ya que el monto es de: " + budgetToConsole + ".");
                            //si el presupueto es mayor a 200k se pone al gerente como cc y el gerente general como sender
                            Array.Resize(ref cc, cc.Length + 1);
                            cc[cc.Length - 1] = gerente;
                            //Extraer mediante un WS SAP el gerente general del pais del AM.
                            try
                            {

                                //DataTable DMO = Lcsql.SelectRow("licitaciones_cr", "SELECT * FROM `email_address` WHERE `CATEGORIA` = 'PRESUPUESTO200K'");
                                DataTable DMO = crud.Select("SELECT * FROM `emailAddress` WHERE `category` = 'PRESUPUESTO200K'", "costa_rica_bids_db");
                                JArray jarray = JArray.Parse(DMO.Rows[0]["jemail"].ToString().Trim());
                                JObject row = JObject.Parse(jarray[0].ToString());
                                gerente_general = row["email"].Value<string>();
                            }
                            catch (Exception)
                            {
                                gerente_general = "RRIVERA@gbm.net";
                            }
                            sender = gerente_general;
                        }
                    }
                    if (enviar)
                    {
                        string html = Properties.Resources.emailLpCr;
                        html = html.Replace("{subject}", "Notificación de Concurso Rechazado SICOP");
                        html = html.Replace("{cuerpo}", header);
                        html = html.Replace("{contenido}", "");

                        mail.SendHTMLMail(html, new string[] { sender }, "Notificación de Concurso Rechazado " + bidNumber + " - " + institution, cc);

                    }
                    //mueve el consurso a una nueva tabla
                    //solo si No es de interes para GBM
                    if (gbmInterest == "2")
                    {
                        lcsql.MoveBidToBackup(idBidNumber);
                        console.WriteLine("Moviendo la licitación a backup debido a que NO es interés de GBM.");
                    }
                    else
                    {
                        //si es de interes para GBM se deja en la tabla principal con un NO en la participación
                        bool act = crud.Update($"UPDATE `purchaseOrder` SET `statusRobot` = 0 WHERE bidNumber = '{bidNumber}'", "costa_rica_bids_db");
                        if (!act) { validar_lineas = false; }
                        console.WriteLine("Se estableció el statusRobot de la licitación en la tabla purchaseOrder en 0, ya que SI es interés de GBM y no se desea mover a backup.");
                    }
                }
                else //si participa
                {
                    //actualizar el campo de "adjuntos" en la base de datos.
                    console.WriteLine("El account manager indicó que SI participa.");
                    if (gbmInterest == "1") //SI
                    {
                        console.WriteLine("La licitación es interés de GBM por tanto se procede a descargar la documentación de la licitación y subirla al FTP.");
                        if (!crBids.OnlyDownload2(bidNumber, crBids.SelConn("https://www.sicop.go.cr/index.jsp"), false, mandante))
                        {
                            mail.SendHTMLMail("", new string[] {"appmanagement@gbm.net"}, "Error al descargar los adjuntos en la si participación del concurso: " + bidNumber, new string[] { "dmeza@gbm.net" });

                            console.WriteLine("Ocurrió un error al descargar la licitación, por tanto se procede a notificar a: " + "internalcustomersrvs@gbm.net" + ", con copia: " + "dmeza@gbm.net" + ", vía email.");

                        }
                    }

                    //DataTable ent_info = Lcsql.SelectRow("licitaciones_cr", "select * from entidades where entidad = '" + institution + "'");
                    //DataTable ent_info = crud.Select( $"SELECT * FROM institutions WHERE institution = '{institution}'", "costa_rica_bids_db");//tomar los concursos actuales
                    #region crear el sales team
                    string[] salesTeamCC = new string[salesTeam.Rows.Count];
                    for (int i = 0; i < salesTeam.Rows.Count; i++)
                    {
                        salesTeamCC[i] = salesTeam.Rows[i]["salesTeam"].ToString() + "@gbm.net";
                    }
                    #endregion

                    //la institución tiene cliente de SAP en la BD
                    string body = "";
                    string sub = "";
                    string customerInstitute = purchaseOrderAdditionalData.Rows[0]["customerInstitute"].ToString();
                    if (!string.IsNullOrWhiteSpace(customerInstitute))
                    {
                        console.WriteLine("Se verificó que la institución tiene cliente en SAP en la base de datos, cliente: " + customerInstitute + ".");
                        string id_opp = purchaseOrderAdditionalData.Rows[0]["opp"].ToString();

                        if (string.IsNullOrEmpty(id_opp))
                        {
                            console.WriteLine("Se procede a crear la oportunidad en SAP.");
                            string cliente_opp = customerInstitute;
                            string contacto_opp = purchaseOrderAdditionalData.Rows[0]["contactId"].ToString(); //concursos.Rows[0]["contacto_institucion"].ToString();
                            contacto_opp = (string.IsNullOrEmpty(contacto_opp)) ? "0070037500" : contacto_opp;

                            #region crear Opp


                            try
                            {
                                tipopp = purchaseOrderAdditionalData.Rows[0]["oppType"].ToString(); //concursos.Rows[0]["tipo_opp"].ToString();

                                DataRow[] oppTypeInfo = oppType.Select("id ='" + tipopp + "'");
                                if (oppTypeInfo.Count() > 0)
                                {
                                    tipopp = oppTypeInfo[0]["key"].ToString();
                                }
                                if (tipopp == "ZOPM")
                                {
                                    sales_type = "0" + purchaseOrderAdditionalData.Rows[0]["salesType"].ToString(); //concursos.Rows[0]["sales_type"].ToString();
                                }
                            }
                            catch (Exception)
                            {
                                tipopp = "ZOPS";
                                sales_type = "";
                            }

                            Dictionary<string, string> campos_opp = new Dictionary<string, string>();

                            campos_opp["tipo"] = tipopp;
                            campos_opp["sales_type"] = sales_type;

                            string opp_descripcion = bidNumber + " - " + description;
                            if (opp_descripcion.Length > 40) { opp_descripcion = opp_descripcion.Substring(0, 40); }
                            campos_opp["descripcion"] = opp_descripcion;
                            campos_opp["fecha_inicio"] = DateTime.Now.Date.ToString("yyyy-MM-dd");
                            campos_opp["Fecha_Final"] = DateTime.Now.AddMonths(2).Date.ToString("yyyy-MM-dd");

                            campos_opp["Ciclo"] = "Y3"; //quotation ??
                            campos_opp["Origen"] = "Y08"; //Public Bid - licitaciones ??

                            campos_opp["grupo_opp"] = "0002";

                            //campos_opp["Cliente"] = cliente_opp; // "0010004721"; 
                            campos_opp["Cliente"] = (cliente_opp.Substring(0, 2) != "00") ? "00" + cliente_opp : cliente_opp; // "0010004721";
                                                                                                                              //campos_opp["Contacto"] = contacto_opp;  // "0070012034";// contacto_opp;
                            campos_opp["Contacto"] = (contacto_opp.Substring(0, 2) != "00") ? "00" + contacto_opp : contacto_opp;  //
                            campos_opp["Usuario"] = "AA" + accountManager.PadLeft(8, '0');  //"AA70000134"; // sales_rep;

                            campos_opp["OrgVentas"] = "O 50000065"; //Costa Rica        
                            campos_opp["OrgServicios"] = "50003583"; //Costa Rica Service Delivery

                            campos_opp["USER"] = user.ToUpper(); //usuario creador

                            id_opp = CreateOpp(campos_opp);
                            #endregion


                            if (!string.IsNullOrEmpty(id_opp))
                            {
                                if (id_opp.Contains("Error"))
                                {
                                    console.WriteLine("Ocurrió un error al crear la oportunidad en SAP: " + id_opp);
                                    validar_lineas = false;
                                    msj_dev = msj_dev + "Se le notifica que el concurso: <b>" + bidNumber + "</b> dio error a la hora de crear la oportunidad en SAP";
                                    body = $"Se le notifica que el concurso: <b>{bidNumber} </b>: {description}. dio error a la hora de crear la oportunidad en SAP: {id_opp}";
                                }
                                else
                                {
                                    console.WriteLine("La oportunidad fue creada exitosamente, número:" + id_opp);
                                    //actualizar opp en BD
                                    bool opp_update = crud.Update($"UPDATE `purchaseOrderAdditionalData` SET `opp` = {id_opp} WHERE bidNumber = '{idBidNumber}'", "costa_rica_bids_db");
                                    if (opp_update)
                                    {
                                        //MENSAJE AL SALES TEAM ??

                                        body = "Se ha creado una nueva Oportunidad en CRM: <b>" + id_opp + "</b> – " + bidNumber + ": " + description + ". de la institución: " + institution + "<a href=\"http://crm-prod-app.gbm.net:8001/sap(bD1lbiZjPTUwMCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?sap-client=500&sap-language=ES\" ><br><br>haga click aquí</a> para entrar a CRM";

                                    }
                                    else
                                    {
                                        msj_dev = msj_dev + "Error al actualizar la opp en la BD: " + id_opp;
                                        body = "Se ha creado una nueva Oportunidad en CRM: <b>" + id_opp + "</b> – " + bidNumber + ": " + description + ". de la institución: " + institution + "<a href=\"http://crm-prod-app.gbm.net:8001/sap(bD1lbiZjPTUwMCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?sap-client=500&sap-language=ES\" ><br><br>haga click aquí</a> para entrar a CRM";
                                        validar_lineas = false;
                                    }
                                }
                                sub = $"Participación en concurso: Nueva OPP {id_opp} – {bidNumber}: {description}";

                            }
                        }
                        else
                        {
                            console.WriteLine("La licitación ya tiene una oportunidad registrada a nivel de SAP, número: " + id_opp);
                            //ya tiene OPP
                            body = "Se ha creado una nueva Oportunidad en CRM: <b>" + id_opp + "</b> – " + bidNumber + ": " + description + ". de la institución: " + institution + "<a href=\"http://crm-prod-app.gbm.net:8001/sap(bD1lbiZjPTUwMCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm?sap-client=500&sap-language=ES\" ><br><br>haga click aquí</a> para entrar a CRM";
                            sub = $"Participación en concurso: OPP Manual – {bidNumber}: {description}";
                        }
                    }
                    else //no tiene cliente
                    {

                        //La institución no tiene cliente de SAP por lo que se procede a notificar solamente
                        body = "Se ha decidido participar en el concurso: <b>" + bidNumber + "</b>: " + description + ". de la institución: " + institution + "<br><br>Solicite el cliente en SAP y proceda a crear la Oportunidad manualmente. Una vez creado el cliente en SAP no olvide notificar a DM para actualizar el portal web de Licitaciones";
                        sub = $"Participación en concurso: OPP Manual – {bidNumber}: {description}";
                    }



                    string html = Properties.Resources.emailLpCr;
                    html = html.Replace("{subject}", "Notificación de Participación en concurso SICOP");
                    html = html.Replace("{cuerpo}", body);
                    html = html.Replace("{contenido}", "");

                    mail.SendHTMLMail(html, new string[] { am_email }, sub, salesTeamCC);

                    bool act = crud.Update($"UPDATE `purchaseOrder` SET `statusRobot` = 0 WHERE bidNumber = '{bidNumber}'", "costa_rica_bids_db");
                    console.WriteLine("Se establece el statusRobot de la licitación en '0'.");
                    if (!act) { validar_lineas = false; }

                }

            }
            catch (Exception ex)
            {
                validar_lineas = false;
                msj_dev = ex.Message;
                crud.Update($"UPDATE `purchaseOrder` SET `statusRobot` = 2 WHERE bidNumber = '{bidNumber}'", "costa_rica_bids_db");
                console.WriteLine("Existió un error al validar las líneas, verifique si la licitación tiene establecido el motivo de no participación y/o el monto de " +
                    "la misma tiene cifras correctas. A continuación se establece el 'StatusRobot' = 2 para que pueda consultarlo en base de datos.");
            }


            if (validar_lineas == false)
            {
                string[] ccopy = { "dmeza@gbm.net" };
                Mail.SendHTMLMail(msj_dev, new string[] {"appmanagement@gbm.net"}, "Error: Licitaciones de Costa Rica, concurso: " + bidNumber, ccopy);

            }
            log.LogDeCambios("Modificación", root.BDProcess, user, "Nueva Participación en el concurso: " + bidNumber, participate, justificacion);
            respFinal = respFinal + "\\n" + "Establecer participación en el concurso: " + bidNumber + " " + participate + " " + justificacion;

            root.requestDetails = respFinal;
            root.BDUserCreatedBy = user;
            console.WriteLine("Fin del proceso.");
        }



        #region Método de apoyo
        public string CreateOpp(Dictionary<string, string> campos)
        {

            string idopp = "";
            try
            {
                RfcDestination destination = new SapVariants().GetDestRFC(mandanteSAP);

                console.WriteLine(" Conectado con SAP CRM");
                RfcRepository repo = destination.Repository;
                IRfcFunction func = repo.CreateFunction("ZOPP_VENTAS");
                IRfcTable general = func.GetTable("GENERAL");
                IRfcTable partners = func.GetTable("PARTNERS");
                //IRfcTable items = func.GetTable("ITEMS");
                console.WriteLine(" Llenando informacion general de oportunidad");
                general.Append();
                general.SetValue("TIPO", campos["tipo"].ToString());
                general.SetValue("DESCRIPCION", campos["descripcion"].ToString());
                general.SetValue("FECHA_INICIO", campos["fecha_inicio"].ToString());
                general.SetValue("FECHA_FIN", campos["Fecha_Final"].ToString());
                general.SetValue("FASE_VENTAS", campos["Ciclo"].ToString());
                general.SetValue("OUTSOURCING", "");
                general.SetValue("SALES_TYPE", campos["sales_type"].ToString());

                //general.SetValue("CICLO_VENTAS", datos_oportunidad.DATA_GENERAL.ORIGEN);
                general.SetValue("PORCENTAJE", "100");
                general.SetValue("REVENUE", "");
                general.SetValue("MONEDA", "USD");
                general.SetValue("GRUPO_OPP", campos["grupo_opp"].ToString());
                general.SetValue("ORIGEN", campos["Origen"].ToString());
                general.SetValue("PRIORIDAD", "4");
                console.WriteLine(" Llenando informacion de cliente y equipo de ventas");
                partners.Append();
                partners.SetValue("PARTNER", campos["Cliente"].ToString());
                partners.SetValue("FUNCTION", "00000021");
                partners.Append();
                partners.SetValue("PARTNER", campos["Contacto"].ToString());
                partners.SetValue("FUNCTION", "00000015");
                partners.Append();
                partners.SetValue("PARTNER", campos["Usuario"].ToString());
                partners.SetValue("FUNCTION", "00000014");
                console.WriteLine(" Llenando Org de Servicios y Ventas");
                //console.WriteLine(" Items data has been added");
                func.SetValue("SALES_ORG", campos["OrgVentas"].ToString());
                //console.WriteLine(" Sales Org data has been added");
                func.SetValue("SRV_ORG", campos["OrgServicios"].ToString());

                func.SetValue("USER", campos["USER"].ToString());
                //console.WriteLine(" Service Org data has been added");
                console.WriteLine(" Creando Oportunidad en SAP CRM");
                func.Invoke(destination);



                IRfcTable validate = func.GetTable("VALIDATE");

                if (func.GetValue("RESPONSE").ToString() != "")
                {
                    console.WriteLine(" Response of the request: " + func.GetValue("RESPONSE").ToString());
                }
                if (func.GetValue("OPP_ID").ToString() != "")
                {
                    console.WriteLine(" ID de la oportunidad creada: " + func.GetValue("OPP_ID").ToString());
                    idopp = func.GetValue("OPP_ID").ToString();
                }
                else
                {
                    idopp = "Error: creating the opportunity";
                    console.WriteLine(" Error creating the opportunity");
                }
                for (int i = 0; i < validate.Count; i++)
                {
                    console.WriteLine(" Generated errors:");
                    Console.WriteLine(DateTime.Now + " > > >  " + validate[i].GetValue("MENSAJE") + "\r\n");
                }
                Console.WriteLine("");

            }
            catch (Exception ex)
            {
                idopp = "Error: " + ex.Message;
            }

            return idopp;
        }

        #endregion
    }
}
