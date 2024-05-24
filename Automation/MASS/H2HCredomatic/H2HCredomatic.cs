using System;
using System.Collections.Generic;
using WinSCP;
using System.IO;
using SAP.Middleware.Connector;
using System.Globalization;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Automation.WEB.H2HCredomatic;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.MASS.H2HCredomatic
{
    /// <summary>
    /// Clase MASS Automation encargada de la ejecución de H2H de credomatic.
    /// </summary>
    class H2HCredomatic
    {
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        SharePoint sharep = new SharePoint();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        ProcessAdmin padmin = new ProcessAdmin();
        Database db2 = new Database();
        Settings sett = new Settings();
        internal Rooting Root { get => root; set => root = value; }
        string saldo_feba = "";
        string[] mt940_files;
        string[] link;
        string[] adjunto;
        string[] adjunto_F = new string[1];
        string[] adjunto_M = new string[1];
        string[] adjunto_P = new string[1];
        string file_name = "", file_path = "";
        string mandante = "ERP";
        string account_id = "";
        string statement_number = "";
        string company_code;
        string bank = "";
        string mensaje_sap = "";
        string mensaje_sap2 = "";
        string mensaje_sap3 = "";
        string respuesta = "";
        string respuesta2 = "";
        string respuesta_M = "";
        string respuesta_F = "";
        string respuesta_P = "";
        bool validar_lineas = true;
        int cantidad_files = 0;
        int contador = 0;
        int contador_F = 0;
        int contador_M = 0;
        int contador_P = 0;
        DateTime file_date = DateTime.MinValue;
        DateTime file_date_before = DateTime.MinValue;
        string sap_date = "";
        string dia = "";
        string mes = "";
        string ano = "";
        public string fldrpath = "";
        public string fldrpathDest = "";
        DateTime datenow = DateTime.Now.Date;

        string respFinal = "";


        public void Main()
        {

            #region variables privadas
            TransferOperationResult transferResult;
            TransferOptions transferOptions = new TransferOptions();
            fldrpath = Root.h2hDownload + "\\";
            fldrpathDest = Root.h2hDownloadArchive + "\\";
            #endregion

            console.WriteLine("Hora para ejecutar H2H de Credomatic según el planificador...");

            //revisa si el usuario RPAUSER esta abierto
            if (!sap.CheckLogin(mandante))
            {
                console.WriteLine("El usuario RPAUSER está desbloqueado. Inicia proceso de H2H Credomatic.");
                sap.BlockUser(mandante, 1);
                SessionOptions sessionOptions = db2.ConnectFTP(2, "h2h.credomatic.com", 222, cred.usuario_h2h, cred.password_h2h, true, "ssh-rsa 1024 Mly1RSVfWCXCNm8Z4zwwgq6o8rGdHoC6+YLCBE5vtqk=");
                sessionOptions.AddRawSettings("ProxyPort", "0");

                using (Session session = new Session())
                {
                    console.WriteLine("Estableciendo conexion");
                    session.Open(sessionOptions);
                    console.WriteLine("Descargando archivos");
                    transferOptions.TransferMode = TransferMode.Binary;
                    transferResult = session.GetFiles("/out/MT940/*", fldrpath + "*", false, transferOptions);

                    transferResult.Check();

                    if (transferResult.Transfers.Count > 0)
                    {
                        int countFilesToProcess = transferResult.Transfers.Count;

                        console.WriteLine($" Inicio de proceso H2H para procesar {countFilesToProcess} archivos...");
                        H2HProcess(transferResult);

                        //Notificación de éxito
                        //NotifyH2HProcessedSuccesfully(countFilesToProcess);

                        if(countFilesToProcess < 5)
                        {
                            #region Notificación de Error                   
                            string errorMsg = $"Se encontraron muy pocos archivos en el servidor del BAC ({countFilesToProcess} en total), por favor verificar con el encargado del BAC";
                            string htmlpage2 = Properties.Resources.emailtemplate1;
                            string[] ccm = { "pmoreira@gbm.net",  "dmeza@gbm.net" , "rasanche@gbm.net", root.BDUserCreatedBy, "epiedra@gbm.net" };
                            string email_gen = htmlpage2;
                            email_gen = email_gen.Replace("{subject}", "Error notificación Carga de Extractos Bancarios - BAC");
                            email_gen = email_gen.Replace("{cuerpo}", errorMsg);
                            email_gen = email_gen.Replace("{contenido}", "");

                            mail.SendHTMLMail(email_gen, new string[] { "fmendez@gbm.net" }, "ERROR: Carga de cuentas: H2H de Credomatic, " + DateTime.Now.ToString(), ccm, null);
                            #endregion
                        }

                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                    }
                    else //No hubo archivos en el Server del BAC
                    {
                        #region Notificación de Error                   
                        string errorMsg = "No hay archivos en el server del BAC el día de hoy, por favor verificar con el encargado del BAC";
                        string htmlpage2 = Properties.Resources.emailtemplate1;
                        string[] ccm = { "pmoreira@gbm.net", "dmeza@gbm.net" , "rasanche@gbm.net", root.BDUserCreatedBy , "epiedra@gbm.net" };
                        string email_gen = htmlpage2;
                        email_gen = email_gen.Replace("{subject}", "Error notificación Carga de Extractos Bancarios - BAC");
                        email_gen = email_gen.Replace("{cuerpo}", errorMsg);
                        email_gen = email_gen.Replace("{contenido}", "");

                        mail.SendHTMLMail(email_gen, new string[] { "fmendez@gbm.net" }, "ERROR: Carga de cuentas: H2H de Credomatic, " + DateTime.Now.ToString(), ccm, null);
                        #endregion
                    }
                    session.Dispose();


                }
                sap.BlockUser(mandante, 0);


            }
            else //El CheckLogin está ocupado 
            {
                //NotifyNotProcessH2HYet();

                //resetear el planner para que lo vuelva a intentar ya que el mandante estaba bloqueado
                sett.setPlannerAgain();
            }

        }

        public void H2HProcess(TransferOperationResult documentos)
        {
            //se vuelven a inicializar las variables para h2hEmail
            fldrpath = Root.h2hDownload + "\\";
            fldrpathDest = Root.h2hDownloadArchive + "\\";
            //
            System.IO.DirectoryInfo di = new DirectoryInfo(fldrpath);
            cantidad_files = di.GetFiles().Length;
            adjunto = new string[cantidad_files];
            string dueno = "";
            List<accounts> cuentas = new List<accounts>();
            Dictionary<string, string> cont_email = new Dictionary<string, string>();
            Dictionary<string, string> contadores = new Dictionary<string, string>();
            string mensaje_email = "";
            string body = "";
            cuentas = val.OwnerSql();
            contadores = val.OwnerCont();

            sap.LogSAP(mandante.ToString());
            foreach (FileInfo file in di.EnumerateFiles())
            {

                mensaje_sap = "";
                mensaje_sap2 = "";
                mensaje_sap3 = "";
                respuesta = "";
                mensaje_email = "";
                company_code = "";
                account_id = "";
                statement_number = "";
                saldo_feba = "";
                try
                {
                    file_path = file.FullName;
                    file_name = Path.GetFileName(file_path);

                    //sacar la cuenta del nombre del archivo
                    string[] stringSeparators0 = new string[] { "GBM-MT940-" };
                    link = file_name.Split(stringSeparators0, StringSplitOptions.None);
                    account_id = link[1].ToString().Trim();
                    int limite = account_id.IndexOf("-" + datenow.Year);
                    if (limite == -1)
                    {
                        limite = account_id.IndexOf("-" + datenow.AddYears(-1).Year);
                    }
                    if (account_id.Length >= limite)
                    { account_id = account_id.Substring(0, limite).Trim(); }

                    //fecha  dd/mm/yyyy de cuando se creo el archivo
                    //File.GetCreationTime(file_path).Date;
                    file_date = File.GetLastWriteTime(file_path).Date;
                    file_date_before = file_date.AddDays(-1);

                    sap_date = file_date_before.ToString("dd.MM.yyyy");

                    string[] stringSeparators00 = new string[] { account_id + "-" };
                    link = file_name.Split(stringSeparators00, StringSplitOptions.None);
                    sap_date = link[1].ToString().Trim().Substring(0, 10);
                    sap_date = DateTime.Parse(sap_date).AddDays(-1).ToString("dd.MM.yyyy");

                    string fullMonthName = DateTime.Parse(sap_date).ToString("MMMM", CultureInfo.CreateSpecificCulture("es"));

                    bool upload = sharep.UploadFileToSharePoint("https://gbmcorp.sharepoint.com/sites/h2hcredomatic", file_path, fullMonthName, sap_date);

                    if (ActiveAccount(cuentas, account_id) != "False")
                    {
                        //dueño de la cuenta
                        dueno = "";
                        company_code = val.CocodeSap(account_id, mandante);
                        if (company_code == "")
                        {
                            //sacar el pais del nombre del archivo
                            string[] stringSeparators1 = new string[] { "MT940-" };
                            link = file_name.Split(stringSeparators1, StringSplitOptions.None);
                            company_code = link[1].ToString().Trim().Substring(0, 2);
                            company_code = val.Cocode(company_code);
                        }

                        //cambio de acuerdo fmendez 15.02.21
                        if (account_id == "41699")
                        {
                            company_code = "GBHQ";
                        }
                        else if (account_id == "45286")
                        {
                            company_code = "ITC0";
                        }
                        else if (account_id == "45294")
                        {
                            company_code = "WTC0";
                        }

                        #region SAP
                        console.WriteLine(DateTime.Now + " > > > " + "Corriendo Script de SAP, cuenta: " + account_id + ", " + file_date);
                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nff_5";
                        SapVariants.frame.SendVKey(0);
                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkEINLESEN")).Selected = true;
                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_KOAUSZ")).Selected = true;
                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_BUPRO")).Selected = true;
                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_STATIK")).Selected = true;
                        ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/cmbFORMAT")).Key = "S";
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtAUSZFILE")).Text = file_path;
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                        //wait
                        try { mensaje_sap3 = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtMESSTXT1")).Text.ToString(); }
                        catch (Exception) { } //mensaje de pop up puede ser error o solo un info "wnd[1]/usr/txtMESSTXT2").setFocus

                        try
                        { ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press(); }
                        catch (Exception) { } //click en ok button

                        try { mensaje_sap = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString(); }
                        catch (Exception) { } //mensaje de barra de status

                        try { mensaje_sap2 = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[2,6]")).Text.ToString(); }
                        catch (Exception) { } //mensaje de label cuando paso de pantalla y no hay nada en el movimiento
                        #endregion

                        #region Procesamiento Respuesta

                        if (mensaje_sap == "Account statement file was not updated" || mensaje_sap == "Bank statement file was not updated")
                        {
                            try
                            {
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                                ((SAPFEWSELib.GuiMenu)SapVariants.session.FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]")).Select();
                                ((SAPFEWSELib.GuiRadioButton)SapVariants.session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]")).Select();
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = fldrpath;
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "Error-MT940-" + company_code + "-" + account_id + ".txt";
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                respuesta = "Error: Account statement file was not updated - ver adjunto con error";
                                //adjuntar archivo de error
                                string errort = fldrpath + "Error-MT940-" + company_code + "-" + account_id + ".txt";
                                adjunto_M[contador_M] = errort;
                                contador_M++;
                                Array.Resize(ref adjunto_M, adjunto_M.Length + 1);
                            }
                            catch (Exception)
                            {
                                respuesta = "Error: Account statement file was not updated, pais: " + company_code + ". Cuenta: " + account_id + " - error al descargar el log de errores" + "<br>";
                            }

                        }
                        else if (mensaje_sap.Contains("does not exist in chart"))
                        {
                            respuesta = "Error: " + mensaje_sap;
                        }
                        else if (mensaje_sap.Contains("Termination in statement no"))
                        {
                            respuesta = "Error: " + mensaje_sap;
                        }
                        else if (mensaje_sap2 == "List contains no data" || mensaje_sap2 == "List does not contain any data")
                        {
                            respuesta = "Cuenta sin movimientos";
                        }
                        else if (mensaje_sap3 == "House bank table: No entry with bank key  and acct")
                        {
                            respuesta = "Error al cargar MT040, La cuenta no existe en SAP";
                        }
                        else if (mensaje_sap3.Contains("Currency USD"))
                        {
                            respuesta = "Error al cargar MT040, La cuenta no esta en dolares.";
                        }
                        else if (mensaje_sap3.Contains("not in table T028B"))
                        {
                            respuesta = "Error al cargar MT040, la cuenta no esta en la tabla T028B.";
                        }
                        else if (mensaje_sap3.Contains("already exists"))
                        {
                            respuesta = "La Cuenta ya se cargo.";
                        }
                        else if (mensaje_sap3.Contains("Acct"))
                        {
                            respuesta = "Error: subir el día anterior de la cuenta y luego volver a subir este día.";
                        }
                        else if (mensaje_sap3.Contains("Account") || mensaje_sap == "" || mensaje_sap3 == "") //correr WS para extraer el numero y saldo
                        {
                            #region SAP

                            try
                            {
                                string COCODE = company_code;
                                //Cambio ya que se sube como Panamá
                                if (account_id == "41699" || account_id == "45286" || account_id == "45294")
                                {
                                    account_id = "0000" + account_id;
                                }

                                Dictionary<string, string> parameters = new Dictionary<string, string>
                                {
                                    ["COCODE"] = COCODE,
                                    ["FECHA"] = sap_date,
                                    ["BANK_ACCOUNT"] = account_id
                                };
                                IRfcFunction func = sap.ExecuteRFC(mandante, "ZFI_GET_SALDO_FEBA", parameters);

                                saldo_feba = func.GetValue("SALDO_TOTAL").ToString();
                                statement_number = func.GetValue("STATEMENT_NO").ToString();
                                respuesta = "MT940 cargado con exito";
                            }
                            catch (Exception ex)
                            {
                                respuesta = "Se subio correctamente, sin embargo dio error extrayendo el saldo " + ex.Message;
                                validar_lineas = false;
                            }
                            #endregion
                        }
                        else
                        {
                            respuesta = mensaje_sap3;
                        }

                        #endregion

                        console.WriteLine(DateTime.Now + " > > > " + "Respuesta del robot: " + account_id + ", " + respuesta);

                        log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Movimiento H2H"," Respuesta del robot: " + account_id + ", " + respuesta, "FMENDEZ");
                        respFinal = respFinal + "\\n" + "Movimiento H2H - Respuesta del robot: " + account_id + ", " + respuesta;


                        mensaje_email = "";
                        mensaje_email = mensaje_email + "<tr>";
                        mensaje_email = mensaje_email + "<td>" + account_id + "</td>";
                        mensaje_email = mensaje_email + "<td>" + company_code + "</td>";
                        mensaje_email = mensaje_email + "<td>" + respuesta + "</td>";
                        mensaje_email = mensaje_email + "<td>" + statement_number + "</td>";
                        mensaje_email = mensaje_email + "<td>" + saldo_feba + "</td>";
                        mensaje_email = mensaje_email + "</tr>";

                        #region owner

                        //una sola respuesta para los tesoreros
                        respuesta_M = respuesta_M + mensaje_email;
                        adjunto_M[contador_M] = file_path;
                        contador_M++;
                        Array.Resize(ref adjunto_M, adjunto_M.Length + 1);

                        //por país.
                        string contador = "fmendez@gbm.net";
                        try
                        { contador = contadores[company_code]; }
                        catch (Exception ex)
                        {
                            console.WriteLine(ex.ToString());
                        }

                        //revisar si ya tiene el AM?
                        if (cont_email.ContainsKey(contador))
                        {
                            //Update
                            cont_email[contador] = cont_email[contador] + mensaje_email;
                        }
                        else
                        {
                            //Insert
                            cont_email[contador] = mensaje_email;
                        }
                        #endregion

                    }
                    else
                    {
                        respuesta = "Aviso: la siguiente cuenta no se carga:";
                    }

                }
                catch (Exception ex)
                {
                    #region catch

                    respuesta = "Error subiendo el archivo, por favor verifique y cargue a mano" + "<br>" + ex.ToString();
                    validar_lineas = false;
                    mensaje_email = "";
                    mensaje_email = mensaje_email + "<tr>";
                    mensaje_email = mensaje_email + "<td>" + account_id + "</td>";
                    mensaje_email = mensaje_email + "<td>" + company_code + "</td>";
                    if (company_code == "GBCR" || company_code == "GBHQ" || company_code == "ITC0" || company_code == "WTC0") //company_code == "BV01" ||  company_code == "SAC0"
                    {
                        mensaje_email = mensaje_email + "<td>" + respuesta + " (QA)</td>";
                    }
                    else
                    {
                        mensaje_email = mensaje_email + "<td>" + respuesta + "</td>";
                    }
                    mensaje_email = mensaje_email + "<td>" + statement_number + "</td>";
                    mensaje_email = mensaje_email + "<td>" + saldo_feba + "</td>";
                    mensaje_email = mensaje_email + "</tr>";
                    //una sola respuesta para los tesoreros
                    respuesta_M = respuesta_M + mensaje_email;
                    adjunto_M[contador_M] = file_path;
                    contador_M++;
                    Array.Resize(ref adjunto_M, adjunto_M.Length + 1);

                    //por país.
                    string contador = "fmendez@gbm.net";
                    try
                    { contador = contadores[company_code]; }
                    catch (Exception)
                    {
                    }

                    //revisar si ya tiene el AM?
                    if (cont_email.ContainsKey(contador))
                    {
                        //Update
                        cont_email[contador] = cont_email[contador] + mensaje_email;
                    }
                    else
                    {
                        //Insert
                        cont_email[contador] = mensaje_email;
                    }

                    #endregion
                }

            } //foreach files

            sap.KillSAP();

            #region adjuntar archivos
            Array.Resize(ref adjunto_F, adjunto_F.Length - 1);
            Array.Resize(ref adjunto_M, adjunto_M.Length - 1);
            Array.Resize(ref adjunto_P, adjunto_P.Length - 1);
            #endregion

            body = body + "<table class='myCustomTable' width='100 %'>";
            body = body + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
            body = body + "<tbody>";
            body = body + respuesta_M;
            body = body + "</tbody>";
            body = body + "</table>";

            string htmlpage2 = Properties.Resources.emailtemplate1;
            if (validar_lineas == false)
            {
                string[] ccm2 = { "dmeza@gbm.net", "epiedra@gbm.net" };
                mail.SendHTMLMail(body, new string[] { "fmendez@gbm.net" }, "Error: H2H de Credomatic, " + datenow, ccm2, adjunto_M);
            }

            body = "";
            body = body + "<table class='myCustomTable' width='100 %'>";
            body = body + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
            body = body + "<tbody>";



            foreach (KeyValuePair<string, string> pair in cont_email)
            {
                string contador = pair.Key.ToString();
                string cuerpo = "";

                cuerpo = cuerpo + "<table class='myCustomTable' width='100 %'>";
                cuerpo = cuerpo + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
                cuerpo = cuerpo + "<tbody>";
                cuerpo = cuerpo + pair.Value.ToString(); //+ "<br><br>"
                cuerpo = cuerpo + "</tbody>";
                cuerpo = cuerpo + "</table>";

                body = body + pair.Value.ToString();

                string emailhtml = htmlpage2;
                emailhtml = emailhtml.Replace("{subject}", "Notificación Carga de Extractos Bancarios - BAC");
                emailhtml = emailhtml.Replace("{cuerpo}", "A Continuación se encuentra el estado de cuentas de su correspondiente país, por favor su colaboración revisando la siguiente tabla, en caso de tener algún error por favor comunicarse con su Tesorero de país");
                emailhtml = emailhtml.Replace("{contenido}", cuerpo);

                mail.SendHTMLMail(emailhtml, new string[] { contador }, " Carga de cuentas: H2H de Credomatic, " + sap_date, null);


            }


            body = body + "</tbody>";
            body = body + "</table>";

            //enviar email de repuesta de exito
            //string[] cc = { "dmeza@gbm.net" };
            string[] ccm = { "pmoreira@gbm.net", "dmeza@gbm.net" , "rasanche@gbm.net", root.BDUserCreatedBy , "epiedra@gbm.net" };
            string email_gen = htmlpage2;
            email_gen = email_gen.Replace("{subject}", "Notificación Carga de Extractos Bancarios - BAC");
            email_gen = email_gen.Replace("{cuerpo}", "A Continuación se encuentra el estado de las cuentas regionales");
            email_gen = email_gen.Replace("{contenido}", body);

            mail.SendHTMLMail(email_gen, new string[] { "fmendez@gbm.net" }, "Carga de cuentas: H2H de Credomatic, " + sap_date, ccm, adjunto_M);
            //mail.SendHTMLMail(respuesta_M, "mdiaz@gbm.net", "Carga de cuentas: H2H de Credomatic, " + datenow, 1, ccm, adjunto_M, resp_type: 2);

            padmin.MoveFiles(fldrpath, fldrpathDest);

            root.requestDetails = respFinal;
            root.BDUserCreatedBy = "FMENDEZ";
        }

        /// <summary>
        /// Método que retorna si una cuenta esta activa o no.
        /// </summary>
        /// <param name="accounts"></param>
        /// <param name="account"></param>
        /// <returns>Retorna un String.</returns>
        public string ActiveAccount(List<accounts> accounts, string account)
        {
            string act = "";
            List<accounts> results = accounts.FindAll(x => x.account == account);
            if (results.Count > 0)
            {
                accounts cuent = results[0];
                act = cuent.active;
            }
            return act;

        }


        #region NotifyMethods

        /// <summary>
        /// Método para notificar que H2H se ejecutó satisfactoriamente.
        /// </summary>
        /// <param name="countFilesToProcess"></param>
        public void NotifyH2HProcessedSuccesfully(int countFilesToProcess)
        {
            string message =
                        "Hola, quiero comunicarte que he ejecutado H2H Credomatic satisfactoriamente.<br><br>" +
                        "Te adjunto datos relevantes:<br>" +
                        "Cantidad de archivos procesados: " + countFilesToProcess + "<br><br><br>Databot.";

            mail.SendHTMLMail(message, new string[] { "epiedra@gbm.net" }, "H2H Credomatic ejecutado satisfactoriamente - Databot ",  new string[] { "dmeza@gbm.net" });

            console.WriteLine($"H2H procesado exitosamente. Cantidad de archivos iniciales procesados {countFilesToProcess}");
        }

        /// <summary>
        /// Método para notificar que no se ha ejecutado H2H porque 
        /// el RPA User está bloqueado por otro robot.
        /// </summary>
        /// <param name="RPAUser"></param>
        public void NotifyNotProcessH2HYet()
        {
            console.WriteLine("El usuario RPAUSER está bloqueado en este momento.");

            string message =
                "Hola, quiero comunicarte que aún he ejecutado H2H Credomatic, ya que el RPAUSER está bloqueado, verifica si corrí el día de hoy.<br><br>" +
                "Te adjunto datos relevantes:<br>" +
                "Estado de bloqueo de RPAUSER Actual: " + sap.CheckLogin(mandante) + "<br><br><br>Databot.";

            mail.SendHTMLMail(message, new string[] { "epiedra@gbm.net" }, "No he podido ejecutar H2H Credomatic - Databot ", new string[] { "dmeza@gbm.net" });

        }
        #endregion



        #region locales
        public void H2hProcessLocal(TransferOperationResult documents)
        {
            System.IO.DirectoryInfo directory = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Q\\");
            int cantcarp = directory.GetDirectories().Length;
            string mensaje_email = "";
            string body = "";
            string dueno = "";
            Dictionary<string, string> cont_email = new Dictionary<string, string>();
            List<accounts> cuentas = new List<accounts>();
            Dictionary<string, string> contadores = new Dictionary<string, string>();
            cuentas = val.OwnerSql();
            contadores = val.OwnerCont();
            sap.LogSAP(mandante.ToString());
            foreach (DirectoryInfo direct in directory.EnumerateDirectories())
            {
                string folder = direct.Name;
                System.IO.DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Q\\" + folder + "\\");
                cantidad_files = di.GetFiles().Length;
                adjunto = new string[cantidad_files];
                mensaje_email = "";
                body = "";
                dueno = "";

                //sap.LogSAPQA("260");
                foreach (FileInfo file in di.EnumerateFiles())
                //foreach (TransferEventArgs transfer in documentos.Transfers)
                {
                    try
                    {
                        //file_name = transfer.FileName; //es el path del server

                        //file_path = transfer.Destination;
                        file_path = file.FullName;
                        file_name = Path.GetFileName(file_path);

                        //fecha  dd/mm/yyyy de cuando se creo el archivo
                        file_date = File.GetLastWriteTime(file_path).Date;
                        file_date_before = file_date.AddDays(-1);

                        sap_date = file_date_before.ToString();
                        var DMY = sap_date.Split(new char[1] { '/' });

                        dia = int.Parse(DMY[0]).ToString();
                        ano = DMY[2].Substring(0, 4).ToString();
                        if (dia.Length == 1)
                        { dia = "0" + dia; }

                        mes = int.Parse(DMY[1]).ToString();
                        if (mes.Length == 1)
                        { mes = "0" + mes; }

                        sap_date = dia + "." + mes + "." + ano;

                        sap_date = folder;

                        mensaje_sap = "";
                        mensaje_sap2 = "";
                        mensaje_sap3 = "";
                        respuesta = "";
                        mensaje_email = "";
                        company_code = "";
                        account_id = "";
                        statement_number = "";
                        saldo_feba = "";

                        //sacar la cuenta del nombre del archivo
                        string[] stringSeparators0 = new string[] { "GBM-MT940-" };
                        link = file_name.Split(stringSeparators0, StringSplitOptions.None);
                        account_id = link[1].ToString().Trim();
                        int limite = account_id.IndexOf("-" + datenow.Year);

                        if (account_id.Length >= limite)
                        { account_id = account_id.Substring(0, limite).Trim(); }

                        if (ActiveAccount(cuentas, account_id) != "0")

                        //if (account_id != "903691996" && account_id != "903692044" && account_id != "104206818" && account_id != "927391193"
                        //    && account_id != "927391201")
                        {
                            //dueño de la cuenta
                            dueno = "";
                            //try
                            //{ dueno = cuentas[account_id]; }
                            //catch (Exception)
                            //{ }

                            //sacar el company code de SAP
                            //try
                            //{

                            company_code = val.CocodeSap(account_id, mandante);

                            //}
                            //catch (Exception)
                            //{

                            //    company_code = "";

                            //}
                            //company_code = cocode(account_id);
                            if (company_code == "")
                            {
                                //sacar el pais del nombre del archivo
                                string[] stringSeparators1 = new string[] { "MT940-" };
                                link = file_name.Split(stringSeparators1, StringSplitOptions.None);
                                company_code = link[1].ToString().Trim().Substring(0, 2);
                                company_code = val.Cocode(company_code);
                            }

                            //cambio de acuerdo mdiaz 15.02.21
                            if (account_id == "41699")
                            {
                                company_code = "GBHQ";
                            }
                            else if (account_id == "45286")
                            {
                                company_code = "ITC0";
                            }
                            else if (account_id == "45294")
                            {
                                company_code = "WTC0";
                            }

                            if (company_code == "GBHQ" || company_code == "ITC0" || company_code == "WTC0") //company_code == "BV01" || company_code == "SAC0" company_code == "GBCR" ||
                            {
                                mensaje_sap3 = "QA apagado, se cargará una vez se habilite";
                                mensaje_sap = "En QA";
                                //continue;
                            }
                            else
                            {
                                #region SAP
                                console.WriteLine("Corriendo Script de SAP, cuenta: " + account_id + ", " + file_date);
                                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nff_5";
                                SapVariants.frame.SendVKey(0);
                                ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkEINLESEN")).Selected = true;
                                ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_KOAUSZ")).Selected = true;
                                ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_BUPRO")).Selected = true;
                                ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_STATIK")).Selected = true;
                                ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/cmbFORMAT")).Key = "S";
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtAUSZFILE")).Text = file_path;
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                                //wait
                                try { mensaje_sap3 = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtMESSTXT1")).Text.ToString(); }
                                catch (Exception) { } //mensaje de pop up puede ser error o solo un info "wnd[1]/usr/txtMESSTXT2").setFocus

                                try
                                { ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press(); }
                                catch (Exception) { } //click en ok button

                                try { mensaje_sap = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString(); }
                                catch (Exception) { } //mensaje de barra de status

                                try { mensaje_sap2 = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[2,6]")).Text.ToString(); }
                                catch (Exception) { } //mensaje de label cuando paso de pantalla y no hay nada en el movimiento
                                #endregion


                            }

                            #region Procesamiento Respuesta

                            if (mensaje_sap == "Account statement file was not updated")
                            {
                                try
                                {
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                                    ((SAPFEWSELib.GuiMenu)SapVariants.session.FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]")).Select();
                                    ((SAPFEWSELib.GuiRadioButton)SapVariants.session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]")).Select();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = root.FilesDownloadPath + "\\h2h_credomatic\\";
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "Error-MT940-" + company_code + "-" + account_id + ".txt";
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                    respuesta = "Error: Account statement file was not updated - ver adjunto con error";
                                    //adjuntar archivo de error
                                    string errort = root.FilesDownloadPath + "\\h2h_credomatic\\" + "Error-MT940-" + company_code + "-" + account_id + ".txt";
                                    adjunto_M[contador_M] = errort;
                                    contador_M++;
                                    Array.Resize(ref adjunto_M, adjunto_M.Length + 1);
                                }
                                catch (Exception)
                                {
                                    respuesta = "Error: Account statement file was not updated, pais: " + company_code + ". Cuenta: " + account_id + " - error al descargar el log de errores" + "<br>";
                                }

                            }
                            else if (mensaje_sap.Contains("does not exist in chart"))
                            {
                                respuesta = "Error: " + mensaje_sap;
                            }
                            else if (mensaje_sap.Contains("Termination in statement no"))
                            {
                                respuesta = "Error: " + mensaje_sap;
                            }
                            else if (mensaje_sap2 == "List contains no data")
                            {
                                respuesta = "Cuenta sin movimientos";
                            }
                            else if (mensaje_sap3 == "House bank table: No entry with bank key  and acct")
                            {
                                respuesta = "Error al cargar MT040, La cuenta no existe en SAP";
                            }
                            else if (mensaje_sap3.Contains("Currency USD"))
                            {
                                respuesta = "Error al cargar MT040, La cuenta no esta en dolares.";
                            }
                            else if (mensaje_sap3.Contains("not in table T028B"))
                            {
                                respuesta = "Error al cargar MT040, la cuenta no esta en la tabla T028B.";
                            }
                            else if (mensaje_sap3.Contains("already exists"))
                            {
                                respuesta = "La Cuenta ya se cargo.";
                            }
                            else if (mensaje_sap3.Contains("Account") || mensaje_sap == "" || mensaje_sap3 == "") //correr WS para extraer el numero y saldo
                            {
                                #region SAP

                                try
                                {
                                    string COCODE = company_code;
                                    //Cambio ya que se sube como Panamá


                                    Dictionary<string, string> parameters = new Dictionary<string, string>();
                                    parameters["COCODE"] = COCODE;
                                    parameters["FECHA"] = sap_date;
                                    parameters["BANK_ACCOUNT"] = account_id;

                                    IRfcFunction func = sap.ExecuteRFC(mandante, "ZFI_GET_SALDO_FEBA", parameters);

                                    saldo_feba = func.GetValue("SALDO_TOTAL").ToString();
                                    statement_number = func.GetValue("STATEMENT_NO").ToString();
                                    respuesta = "MT940 cargado con exito";

                                }
                                catch (Exception ex)
                                {
                                    respuesta = "Se subio correctamente, sin embargo dio error extrayendo el saldo " + ex.Message;
                                    validar_lineas = false;
                                }
                                #endregion
                            }
                            else
                            {
                                respuesta = mensaje_sap3;
                            }

                            #endregion

                            console.WriteLine("Respuesta del robot: " + account_id + ", " + respuesta);

                            mensaje_email = "";
                            mensaje_email = mensaje_email + "<tr>";
                            mensaje_email = mensaje_email + "<td>" + account_id + "</td>";
                            mensaje_email = mensaje_email + "<td>" + company_code + "</td>";
                            if (company_code == "GBHQ" || company_code == "ITC0" || company_code == "WTC0") //company_code == "BV01" ||  company_code == "SAC0" company_code == "GBCR" ||
                            {
                                mensaje_email = mensaje_email + "<td>" + respuesta + " (QA)</td>";
                            }
                            else
                            {
                                mensaje_email = mensaje_email + "<td>" + respuesta + "</td>";
                            }
                            mensaje_email = mensaje_email + "<td>" + statement_number + "</td>";
                            mensaje_email = mensaje_email + "<td>" + saldo_feba + "</td>";
                            mensaje_email = mensaje_email + "</tr>";

                            #region owner

                            //una sola respuesta para los tesoreros
                            respuesta_M = respuesta_M + mensaje_email;
                            adjunto_M[contador_M] = file_path;
                            contador_M++;
                            Array.Resize(ref adjunto_M, adjunto_M.Length + 1);

                            //por país.
                            string contador = "fmendez@gbm.net";

                            try
                            { contador = contadores[company_code]; }
                            catch (Exception ex)
                            {
                                console.WriteLine(ex.ToString());
                            }

                            //revisar si ya tiene el AM?
                            if (cont_email.ContainsKey(contador))
                            {
                                //Update
                                cont_email[contador] = cont_email[contador] + mensaje_email;
                            }
                            else
                            {
                                //Insert
                                cont_email[contador] = mensaje_email;
                            }
                            #endregion

                        }
                        else
                        {
                            respuesta = "Aviso: la siguiente cuenta no se carga:";
                        }

                    }
                    catch (Exception ex)
                    {
                        #region exception 
                        respuesta = "Error subiendo el archivo, por favor verifique y cargue a mano" + "<br>" + ex.ToString();
                        validar_lineas = false;
                        mensaje_email = "";
                        mensaje_email = mensaje_email + "<tr>";
                        mensaje_email = mensaje_email + "<td>" + account_id + "</td>";
                        mensaje_email = mensaje_email + "<td>" + company_code + "</td>";
                        if (company_code == "GBCR" || company_code == "GBHQ" || company_code == "ITC0" || company_code == "WTC0") //company_code == "BV01" ||  company_code == "SAC0"
                        {
                            mensaje_email = mensaje_email + "<td>" + respuesta + " (QA)</td>";
                        }
                        else
                        {
                            mensaje_email = mensaje_email + "<td>" + respuesta + "</td>";
                        }
                        mensaje_email = mensaje_email + "<td>" + statement_number + "</td>";
                        mensaje_email = mensaje_email + "<td>" + saldo_feba + "</td>";
                        mensaje_email = mensaje_email + "</tr>";
                        //una sola respuesta para los tesoreros
                        respuesta_M = respuesta_M + mensaje_email;
                        adjunto_M[contador_M] = file_path;
                        contador_M++;
                        Array.Resize(ref adjunto_M, adjunto_M.Length + 1);

                        //por país.
                        string contador = "fmendez@gbm.net";
                        try
                        { contador = contadores[company_code]; }
                        catch (Exception)
                        {
                        }

                        //revisar si ya tiene el AM?
                        if (cont_email.ContainsKey(contador))
                        {
                            //Update
                            cont_email[contador] = cont_email[contador] + mensaje_email;
                        }
                        else
                        {
                            //Insert
                            cont_email[contador] = mensaje_email;
                        }
                        #endregion
                    }

                } //foreach files
            }
            sap.KillSAP();
            //sap.KillSAPQA();

            #region adjuntar archivos
            Array.Resize(ref adjunto_F, adjunto_F.Length - 1);
            Array.Resize(ref adjunto_M, adjunto_M.Length - 1);
            Array.Resize(ref adjunto_P, adjunto_P.Length - 1);
            #endregion

            body = body + "<table id='concursos' width='100 %'>";
            body = body + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
            body = body + "<tbody>";
            body = body + respuesta_M;
            body = body + "</tbody>";
            body = body + "</table>";

            string htmlpage2 = Properties.Resources.emailtemplate1;

            if (validar_lineas == false)
            {
                //enviar email de repuesta de error
                //string[] cc = { "dmeza@gbm.net" };

                string[] ccm2 = { "pmoreira@gbm.net", "dmeza@gbm.net" };
                mail.SendHTMLMail(body, new string[] { "fmendez@gbm.net" }, "Error: H2H de Credomatic, " + datenow, ccm2, adjunto_M);


            }

            //}
            body = "";
            body = body + "<table id='concursos' width='100 %'>";
            body = body + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
            body = body + "<tbody>";



            foreach (KeyValuePair<string, string> pair in cont_email)
            {
                string contador = pair.Key.ToString();
                string cuerpo = "";

                cuerpo = cuerpo + "<table id='concursos' width='100 %'>";
                cuerpo = cuerpo + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
                cuerpo = cuerpo + "<tcuerpo>";
                cuerpo = cuerpo + pair.Value.ToString(); //+ "<br><br>"
                cuerpo = cuerpo + "</tcuerpo>";
                cuerpo = cuerpo + "</table>";

                body = body + pair.Value.ToString();

                string emailhtml = htmlpage2;
                emailhtml = emailhtml.Replace("{subject}", "Notificación Carga de Extractos Bancarios - BAC");
                emailhtml = emailhtml.Replace("{cuerpo}", "A Continuación se encuentra el estado de cuentas de su correspondiente país del fin de semana pasado, por favor su colaboración revisando la siguiente tabla, en caso de tener algún error por favor comunicarse con su Tesorero de país");
                emailhtml = emailhtml.Replace("{contenido}", cuerpo);

                mail.SendHTMLMail(emailhtml, new string[] { contador }, " Carga de cuentas: H2H de Credomatic, " + sap_date, null);


            }


            body = body + "</tbody>";
            body = body + "</table>";

            string[] ccm = { "pmoreira@gbm.net", "dmeza@gbm.net" };
            string email_gen = htmlpage2;
            email_gen = email_gen.Replace("{subject}", "Notificación Carga de Extractos Bancarios - BAC");
            email_gen = email_gen.Replace("{cuerpo}", "A Continuación se encuentra el estado de las cuentas regionales");
            email_gen = email_gen.Replace("{contenido}", body);

            mail.SendHTMLMail(email_gen, new string[] { "fmendez@gbm.net" }, "Carga de cuentas: H2H de Credomatic, " + sap_date, ccm, adjunto_M);

            padmin.DeleteFiles(root.FilesDownloadPath + "\\h2h_credomatic\\");

        }
        public void H2hProcessQA(TransferOperationResult documents)
        {

            System.IO.DirectoryInfo di = new DirectoryInfo(root.FilesDownloadPath + "\\h2h_credomatic\\");
            cantidad_files = di.GetFiles().Length;
            adjunto = new string[cantidad_files];
            string dueno = "";
            List<accounts> cuentas = new List<accounts>();
            Dictionary<string, string> cont_email = new Dictionary<string, string>();
            Dictionary<string, string> contadores = new Dictionary<string, string>();
            string mensaje_email = "";
            string body = "";
            string resp = "";
            contadores = val.OwnerCont();
            foreach (TransferEventArgs transfer in documents.Transfers)
            {
                try
                {

                    file_path = transfer.Destination;
                    file_name = Path.GetFileName(file_path);

                    //sacar la cuenta del nombre del archivo
                    string[] stringSeparators0 = new string[] { "GBM-MT940-" };
                    link = file_name.Split(stringSeparators0, StringSplitOptions.None);
                    account_id = link[1].ToString().Trim();
                    int limite = account_id.IndexOf("-" + datenow.Year);

                    if (account_id.Length >= limite)
                    { account_id = account_id.Substring(0, limite).Trim(); }

                    //fecha  dd/mm/yyyy de cuando se creo el archivo
                    file_date = File.GetLastWriteTime(file_path).Date;
                    file_date_before = file_date.AddDays(-1);

                    sap_date = file_date_before.ToString();
                    var DMY = sap_date.Split(new char[1] { '/' });

                    dia = int.Parse(DMY[0]).ToString();
                    ano = DMY[2].Substring(0, 4).ToString();
                    if (dia.Length == 1)
                    { dia = "0" + dia; }

                    mes = int.Parse(DMY[1]).ToString();
                    if (mes.Length == 1)
                    { mes = "0" + mes; }

                    sap_date = dia + "." + mes + "." + ano;

                    company_code = val.CocodeSap(account_id, mandante);
                    if (company_code == "")
                    {
                        //sacar el pais del nombre del archivo
                        string[] stringSeparators1 = new string[] { "MT940-" };
                        link = file_name.Split(stringSeparators1, StringSplitOptions.None);
                        company_code = link[1].ToString().Trim().Substring(0, 2);
                        company_code = val.Cocode(company_code);
                    }
                    //cambio de acuerdo fmendez 15.02.21
                    if (account_id == "41699")
                    {
                        company_code = "GBHQ";
                    }
                    else if (account_id == "45286")
                    {
                        company_code = "ITC0";
                    }
                    else if (account_id == "45294")
                    {
                        company_code = "WTC0";
                    }

                    bool upload = sharep.UploadFileToSharePoint("https://gbmcorp.sharepoint.com/sites/h2hcredomatic", file_path, sap_date);

                    if (upload)
                    {
                        resp = "Se subio correctamente a Sharepoint";
                    }
                    else
                    {
                        resp = "Error al subir el archivo al Sharepoint";
                        validar_lineas = false;
                    }


                    //FIN DEL CAMBIO----------------------------------------------------------------------------------
                }
                catch (Exception ex)
                {
                    resp = "Error desconocido: " + ex.Message;
                    validar_lineas = false;
                }
                mensaje_email = "";
                mensaje_email = mensaje_email + "<tr>";
                mensaje_email = mensaje_email + "<td>" + account_id + "</td>";
                mensaje_email = mensaje_email + "<td>" + company_code + "</td>";
                mensaje_email = mensaje_email + "<td>" + resp + "</td>";
                mensaje_email = mensaje_email + "<td>" + "" + "</td>";
                mensaje_email = mensaje_email + "<td>" + "" + "</td>";
                mensaje_email = mensaje_email + "</tr>";


                respuesta_M = respuesta_M + mensaje_email;

                //por país.
                string contador = "fmendez@gbm.net";
                try
                { contador = contadores[company_code]; }
                catch (Exception)
                { }

                //revisar si ya tiene el AM?
                if (cont_email.ContainsKey(contador))
                {
                    //Update
                    cont_email[contador] = cont_email[contador] + mensaje_email;
                }
                else
                {
                    //Insert
                    cont_email[contador] = mensaje_email;
                }

            } //foreach
            string htmlpage = Properties.Resources.emailtemplate1;
            body = "";
            body = body + "<table id='concursos' width='100 %'>";
            body = body + "<thead><tr><th>Cuenta</th><th>País</th><th>Respuesta</th><th>No FEBA</th><th>saldo total</th></tr></thead>";
            body = body + "<tbody>";


            foreach (KeyValuePair<string, string> pair in cont_email)
            {
                body = body + pair.Value.ToString();
            }


            body = body + "</tbody>";
            body = body + "</table>";

            //enviar email de repuesta de exito
            //string[] cc = { "dmeza@gbm.net" };
            string[] ccm = { "dmeza@gbm.net", "jfjimenez@gbm.net" };
            string email_gen = htmlpage;
            email_gen = email_gen.Replace("{subject}", "Notificación Carga de Extractos Bancarios - BAC");
            email_gen = email_gen.Replace("{cuerpo}", "A Continuación se encuentra el estado de las cuentas regionales");
            email_gen = email_gen.Replace("{contenido}", body);

            mail.SendHTMLMail(email_gen, new string[] { "fmendez@gbm.net" }, "Carga de cuentas: H2H de Credomatic, " + datenow, ccm, adjunto_M);
            //mail.SendHTMLMail(respuesta_M, "fmendez@gbm.net", "Carga de cuentas: H2H de Credomatic, " + datenow, 1, ccm, adjunto_M, resp_type: 2);
            if (validar_lineas)
            {
                padmin.DeleteFiles(root.FilesDownloadPath + "\\h2h_credomatic\\");
            }

            //}


        }
        #endregion
    }


}
