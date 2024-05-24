using System;
using System.IO;
using WinSCP;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.SpecialBidForms;
using DataBotV5.App.Global;
using DataBotV5.Data.Projects.SpecialBidForm;

namespace DataBotV5.Automation.WEB.SpecialBid
{
    /// <summary>
    /// NUEVA Clase RPA Automation encargada de generar un Form de una licitación especial.
    /// </summary>
    class NewSpecialBidForm
    {
        SpecialBidFormSQL spbidsql = new SpecialBidFormSQL();
        Rooting roots = new Rooting();
        Database db2 = new Database();
        ValidateData val = new ValidateData();
        ConsoleFormat console = new ConsoleFormat();
        SbForm sb = new SbForm();
        ProcessInteraction proc = new ProcessInteraction();
        CRUD crud = new CRUD();
        MailInteraction mail = new MailInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        Credentials cred = new Credentials();
        string sb_number = "";
        public string id_carpeta = "";
        public string projectname;
        public string ibm_country = "";
        public string useopp = "";
        public string opp = "";
        public string useprevbid = "";
        public string prevbid = "";
        public string priceupd = "";
        public string justi = "";
        public string customer = "";
        public string brand = "";
        public string justi2 = "";
        public string addquest = "";
        public string bpjusti = "";
        public string swma = "";
        public string renew = "";
        public string totalprice = "";
        public string customerprice = "";
        public string totalright = "";
        public string totalright2 = "";
        public string customerright = "";
        public string customerright2 = "";
        public string sb_id_gestion = "";
        public string usuario = "";
        public string alerta_text = "";



        public void Main()
        {
            //leer la base de datos a ver si hay una nueva solicitud
            id_carpeta = "";
            id_carpeta = spbidsql.NewRequestSB();
            if (id_carpeta != string.Empty && id_carpeta != "")
            {
                console.WriteLine("Procesando...");
                ProcessSBFormNew(roots.SB_FILE_Download + "\\" + id_carpeta);
                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessSBFormNew(string route)
        {
            try
            {

                #region crear ruta
                DirectoryInfo dir = new DirectoryInfo(route);
                if (dir.Exists != true)
                {
                    DirectoryInfo di = Directory.CreateDirectory(route);
                }
                #endregion

                int rows; string validacion; string respuesta = "";
                bool validar_lineas = true;
                TransferOperationResult transferResult;
                TransferOptions transferOptions = new TransferOptions();
                string subject_sb = "Special Bid - Notificacion de Finalizado: Creacion del Documento de la Gestion - #";
                console.WriteLine(" Extraer info de la base de datos");
                #region extraer info
                spbidsql.ExtractSBInfo(id_carpeta);
                #endregion
                subject_sb = subject_sb + sb_id_gestion;

                if (projectname == "" || projectname == null)
                {
                    respuesta = "No se encontro el Project de la solicitud de la carpeta: " + id_carpeta;
                    validar_lineas = false;
                }
                else
                {
                    #region validar data
                    projectname = projectname.Trim().ToUpper();
                    projectname = val.RemoveSpecialChars(projectname, 1);

                    justi = justi.Trim().ToUpper();
                    justi = val.RemoveSpecialChars(justi, 1);
                    justi = justi.Replace("\n", " ");
                    justi = justi.Replace("\r", " ");
                    if (justi.Length > 100) { justi = justi.Substring(0, 100); }

                    justi2 = justi2.Trim().ToUpper();
                    justi2 = val.RemoveSpecialChars(justi2, 1);
                    justi2 = justi2.Replace("\r", " ");
                    justi2 = justi2.Replace("\n", " ");
                    if (justi2.Length > 100) { justi2 = justi2.Substring(0, 100); }

                    bpjusti = bpjusti.Trim().ToUpper();
                    bpjusti = val.RemoveSpecialChars(bpjusti, 1);
                    bpjusti = bpjusti.Replace("\n", " ");
                    bpjusti = bpjusti.Replace("\r", " ");
                    if (bpjusti.Length > 700) { bpjusti = bpjusti.Substring(0, 700); }

                    totalprice = totalprice.ToUpper();
                    customerprice = customerprice.ToUpper();

                    totalright = totalprice.Substring(totalprice.Length - 3, 3);
                    totalright2 = totalprice.Substring(totalprice.Length - 2, 2);
                    if (totalright.Substring(0, 1) == "." || totalright2.Substring(0, 1) == ".")
                    {
                        totalprice = totalprice.Replace(",", "");
                        totalprice = totalprice.Replace(".", ",");
                    }

                    customerright = customerprice.Substring(customerprice.Length - 3, 3);
                    customerright2 = customerprice.Substring(customerprice.Length - 2, 2);
                    if (customerright.Substring(0, 1) == "." || customerright2.Substring(0, 1) == ".")
                    {
                        customerprice = customerprice.Replace(",", "");
                        customerprice = customerprice.Replace(".", ",");
                    }

                    int punto;
                    punto = totalprice.IndexOf(".") + 1;
                    if (punto == 0) { totalprice += ".00"; }
                    punto = customerprice.IndexOf(".") + 1;
                    if (punto == 0) { customerprice += ".00"; }

                    #endregion

                    #region crear Special Bid

                    bool crear = val.ValidateSBForm(ibm_country, useopp, opp, useprevbid, prevbid, priceupd, justi, customer, brand, justi2,
                                  addquest, bpjusti, swma, renew, totalprice, customerprice);

                    if (crear == true)
                    {
                        try
                        {
                            #region descarga archivos del server web

                            SessionOptions sessionOptions = db2.ConnectFTP(1, "databot.gbm.net", 21, "gbmadmin", cred.password_server_web, false, "");

                            sessionOptions.AddRawSettings("ProxyPort", "0");

                            using (Session session = new Session())
                            {
                                console.WriteLine(" Estableciendo conexion");
                                session.Open(sessionOptions);

                                console.WriteLine(" Descargando archivos");
                                transferOptions.TransferMode = TransferMode.Binary;

                                transferResult = session.GetFiles("/special_bid_files/" + id_carpeta + "/*", route + "\\*", false, transferOptions);

                                transferResult.Check();

                                if (transferResult.Transfers.Count > 0)
                                {
                                    string[] adjunto_sb = new string[1];
                                    string file_name = "";
                                    int cfgcount = 0;
                                    int cfgcount2 = 0;
                                    foreach (TransferEventArgs transfer in transferResult.Transfers)
                                    {
                                        string file_path = transfer.Destination;
                                        file_name = Path.GetFileName(file_path);

                                        string extArchivo = Path.GetExtension(file_name);
                                        if (extArchivo.ToLower() == ".cfr")
                                        {
                                            adjunto_sb[cfgcount2] = route + "\\" + file_name;
                                            cfgcount2++;
                                            Array.Resize(ref adjunto_sb, adjunto_sb.Length + 1);
                                        }
                                    }
                                    roots.cfr_list = new string[adjunto_sb.Length - 1];
                                    Array.Copy(adjunto_sb, roots.cfr_list, adjunto_sb.Length - 1);
                                }
                                session.Dispose();
                            }


                            #endregion

                            roots.BDUserCreatedBy = usuario;
                            console.WriteLine("Llenando formulario con Selenium");
                            alerta_text = sb.SbCreateFormWeb(ibm_country, projectname, useopp, opp, useprevbid, prevbid, priceupd, justi, customer, brand, justi2,
                                addquest, bpjusti, swma, renew, totalprice, customerprice);
                            if (alerta_text == "Error 500 - Internal Server Error")
                            {
                                respuesta = respuesta + projectname + ": " + "Error 500 - Internal Server Error" + "<br>";
                            }
                            else if (alerta_text == "Cliente no existe")
                            {
                                respuesta = respuesta + projectname + ": " + "Error 404 - Cliente: " + customer + " no existe" + "<br>";
                            }
                            else
                            {
                                sb_number = roots.id_special_bid;
                                respuesta = respuesta + projectname + ": " + "Creado con Exito, " + sb_number + "<br>";
                                //enviar respuesta al usuario con el id de special bid
                            }


                        }
                        catch (Exception ex)
                        {
                            console.WriteLine("Error al crear el formulario");
                            respuesta = respuesta + projectname + ": " + "Error al crear el formulario" + "<br>" + "<br>" + ex.ToString() + "<br>";
                            Console.WriteLine(ex.ToString());
                            System.Threading.Thread.Sleep(1000);
                            proc.KillProcess("chromedriver", true);
                            proc.KillProcess("chrome", true);
                        }
                    }
                    else
                    {
                        respuesta = respuesta + projectname + ": " + "Error en el formulario, verifique la informacion" + "<br>";

                    }
                    #endregion
                    log.LogDeCambios("Creacion", roots.BDProcess, usuario, "Crear Special Bid Form", respuesta + ": " + sb_number, sb_id_gestion);


                }

                console.WriteLine("Respondiendo solicitud");
                if (respuesta.Contains("Error"))
                {
                    subject_sb = "Special Bid - Notificacion de Error: Creacion del Documento de la Gestion - #" + sb_id_gestion;
                    string[] cc = { "dmeza@gbm.net", usuario };
                    spbidsql.DeleteRequestSB(id_carpeta);
                    mail.SendHTMLMail(respuesta, new string[] {"appmanagement@gbm.net"}, subject_sb, cc);
                }
                else
                {
                    mail.SendHTMLMail(respuesta + "<br>" + "<br>" + "Haga Click <a href=\'https://extbasicbpmsprd.podc.sl.edst.ibm.com/bpms/'>aqui</a> para ver el documento", new string[] { usuario }, subject_sb);
                    spbidsql.CompleteRequestSB(id_carpeta);
                    spbidsql.DeleteFolderWin(route);
                    //padmin.eliminar_files(roots.SB_FILE_Download + @"\");
                }


                proc.KillProcess("EXCEL", true);
                proc.KillProcess("chromedriver", true);

            }
            catch (Exception ex)
            {
                if (id_carpeta != "" || id_carpeta != null)
                {
                    spbidsql.DeleteRequestSB(id_carpeta);
                }
                string subject_sb = "Special Bid - Notificacion de Error: Creacion del Documento de la Gestion - #" + sb_id_gestion;
                string[] cc = { "dmeza@gbm.net", usuario };
                mail.SendHTMLMail(ex.ToString(), new string[] {"appmanagement@gbm.net"}, subject_sb, cc);

            }

        }


    }
}
