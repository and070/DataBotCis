using DataBotV5.App.Global.Interfaces;
using DataBotV5.Logical.Files;
using DataBotV5.Logical.Mail;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Newtonsoft.Json.Linq;
using Newtonsoft;
using DataBotV5.Data.Root;
using System.Data;
using Newtonsoft.Json;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using DataBotV5.App.ConsoleApp;
using DataBotV5.Data.Process;

namespace DataBotV5.App.Global
{
    class Settings : IDisposable
    {
        private bool disposedValue;
        public FileInfo LogFile;
        ConsoleFormat console = new ConsoleFormat();
        CRUD crud = new CRUD();
        Rooting root = new Rooting();
        SapVariants sap = new SapVariants();
        private System.Reflection.Assembly assembly { get; set; }
        private System.Diagnostics.FileVersionInfo fvi { get; set; }
        private string version = " 5.0.0";
        public string title { get; set; }
        public ConsoleColor color { get; set; }

        public Settings()
        {
            assembly = System.Reflection.Assembly.GetExecutingAssembly();
            fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
        }
        public void PrintLogo()
        {
            Console.Title = title;
            Console.ForegroundColor = color;
            console.WriteLine("");
            console.WriteLine(@"██████╗  █████╗ ████████╗ █████╗ ██████╗  ██████╗ ████████╗");
            console.WriteLine(@"██╔══██╗██╔══██╗╚══██╔══╝██╔══██╗██╔══██╗██╔═══██╗╚══██╔══╝");
            console.WriteLine(@"██║  ██║███████║   ██║   ███████║██████╔╝██║   ██║   ██║   ");
            console.WriteLine(@"██║  ██║██╔══██║   ██║   ██╔══██║██╔══██╗██║   ██║   ██║   ");
            console.WriteLine(@"██████╔╝██║  ██║   ██║   ██║  ██║██████╔╝╚██████╔╝   ██║   ");
            console.WriteLine(@"╚═════╝ ╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝╚═════╝  ╚═════╝    ╚═╝ Versión." + version);
            console.WriteLine("");
            console.WriteLine("MIS Automation, GBM as a Service " + DateTime.Now.Year.ToString());
            console.WriteLine("");
            console.WriteLine("");
            console.WriteLine("");
        }

        /// <summary>
        /// Método para realizar la impresión del ambiente al cual fue establecido en la variable enviroment en la clase Start.
        /// Enviroment: "PRD" apunta al 10.7.60.72. 
        /// "QAS" o "DEV" apunta al localhost.
        /// </summary>
        /// <param name="enviroment"></param>
        public void PrintEnviroment(String enviroment)
        {
            String auxText="";
            if (enviroment == "PRD")
                auxText+="Producción.";
             else if (enviroment== "QAS" || enviroment=="DEV")
                auxText += "Localhost.";
             else
                auxText += "Existe un error con el entorno de ejecución seleccionado. Por favor revisar antes de ejecutar.";
                        
            console.WriteLine("Entorno de ejecución: " + auxText);
            console.WriteLine("");
        }

        public bool Carpeta()
        {
            string ruta = string.Empty;
            using (Rooting root = new Rooting())
            {
                ruta = root.DatabotPath;
            }
            return Directory.Exists(ruta);
        }
        public void CreateLog(String msj)
        {
            string pathFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\databot\BackupLogs";
            string pathFileTxt = pathFolder + $"\\LogRobot{DateTime.Now.ToString("yyyyMMdd")}.txt";
            //Verifica si existe la carpeta.
            Directory.CreateDirectory(pathFolder);
            //Crea o agrega texto al log .txt. del día.
            using (StreamWriter sw = File.AppendText(pathFileTxt))
            {
                sw.WriteLine(msj);
            }
        }

        public void CrearDirectorio()
        {
            string ruta = @" D:\documentos nuevos\";
            // Especifica el directorio a manipular.
            try
            {
                if (Directory.Exists(ruta))                         // Revisar si existe el directorio
                {
                    Directory.Delete(ruta, true);                               // borrarlo
                }

                DirectoryInfo di = Directory.CreateDirectory(ruta);             //crearlo
            }
            catch (Exception e)
            {
                new ConsoleFormat().WriteLine(e.Message);
            }
        }

        // para crear un archivo
        void verificarSiExisteArchivo()
        {
            string rutacompletaLogs = @" D:\mi archivo.txt";
            if (File.Exists(rutacompletaLogs))
            {
                File.Delete(rutacompletaLogs);               // borrarlo
            }
            using (StreamWriter file = File.AppendText(rutacompletaLogs))   // crear el archivo.
            {
                file.WriteLine("Archivo creardo " + DateTime.Now);          //escribir el archivo.
            }
            new ConsoleFormat().WriteLine("Archivo Logs Creado " + rutacompletaLogs);
        }


        public void BlankProcess()
        {
            using (DestroyProcess kill = new DestroyProcess())
            {
                kill.KillProcess("chromedriver", true);
                kill.KillProcess("EXCEL", true);
                kill.KillProcess("saplogon", false);
                //kill.KillProcess("chrome", true);
            }
            //sap.BlockUserWithoutCnslLog(300, 0);
            //sap.BlockUserWithoutCnslLog(260, 0);
        }
        public void EndProgram(string razon)
        {
            Console.Title = "Databot is Offline";
            using (ConsoleFormat console = new ConsoleFormat())
            {
                console.WriteAnnounce(razon);
            }
            //CreateLog(razon);
            //using (DataBotV5.Data.Root.Rooting root = new DataBotV5.Data.Root.Rooting())
            //{
            //    root.Dispose();
            //}
            System.Threading.Thread.Sleep(3000);
            System.Environment.Exit(1);
        }
        public void PauseProgram()
        {
            Console.Title = "Databot is Paused";
            using (ConsoleFormat console = new ConsoleFormat())
            {
                console.WriteAnnounce("ALERT: > > > End Application? Y/N");
                string parametro = Console.ReadLine();
                while (parametro != "Y" && parametro != "N")
                {
                    console.WriteLine($"El comando: {parametro} no es válido, favor seleccionar uno de la lista anterior.");
                    parametro = Console.ReadLine();
                }
                if (parametro == "Y")
                {
                    EndProgram("END OPERATION: > > > Exit by User");
                }
                else
                {
                    console.WriteAnnounce("RESUMING OPERATION: > > > Resume by User");
                    System.Threading.Thread.Sleep(3000);
                }
            }
        }

        /// <summary>
        /// Este método notifica através de correo electrónico y por consola, si aparece un mensaje de 
        /// error en algún robot de start, con su respectiva línea, método y clase. 
        /// *Importante al momento de publicar*, en el apartado de las propiedades de DatabotV5, 
        /// publicación, archivos de publicación, incluir todos los archivos .pbd, ya que al momento de publicar,
        /// realiza una copia del debug local, con referencia de las líneas de código, es útil para que pueda indicar
        /// donde se encuentra la línea de código con error, de lo contrario no lo indica y el método no sería útil.
        /// </summary>
        /// <param name="ex">La referencia de excepción producida en Start.</param>
        /// <param name="x"> El datarow actual el while "infinito" del Start.</param>
        /// <param name="area">"El datarow actual el while "infinito" del Start."</param>
        public void ErrorProgram(Exception ex, DataRow x, string area, string enviroment)
        {
            Rooting root = new Rooting();
            string[] cc;
            string sender;

            #region Generación de mensaje de error Email.
            string InnerExceptionMsg, StackTraceMsg, TargetSiteMsg;
            StackTrace traceForGetters = new StackTrace(ex, true);

            //Valida si la excepción no está vacía
            InnerExceptionMsg = (ex.InnerException != null) ? ex.InnerException.ToString() : "";
            StackTraceMsg = (ex.StackTrace != null) ? ex.StackTrace.ToString() : "";
            TargetSiteMsg = (ex.TargetSite != null) ? ex.TargetSite.ToString() : "";

            String msgEmail =
                         "Robot: <strong>" + x["class"] + "</strong> <br>" +
                         "Proyecto: <strong>" + x["projectName"] + "</strong> <br>" +
                         "Área: <strong>" + x["area"] + "</strong> ";
            try
            {
                if (traceForGetters.GetFrame(0).GetFileLineNumber() != 0) //Verifica si el StackTrace tiene información útil.
                    msgEmail += "<br><br>Punto de referencia del error: " +
                             "<br>Clase: " + traceForGetters.GetFrame(0).GetFileName() +
                             "<br>Línea: " + traceForGetters.GetFrame(0).GetFileLineNumber() +
                             "<br>Método: " + traceForGetters.GetFrame(0).GetMethod()+ "<br>";
            }
            catch (Exception)
            {
                msgEmail += "<br><br>No es posible brindar un punto de referencia del error exacto. <br>";
            }


            msgEmail += "<br>A continuación se detalla el tipo de error generado: <br><br>" +

                         "<strong>Tipo de excepción:</strong>  " + ex.Message.ToString() + "<br><br>" +
                         "<strong>InnerException:</strong>" + InnerExceptionMsg + "<br><br>" +
                         "<strong>StackTrace: </strong>" + StackTraceMsg + "<br><br>" +
                         "<strong>TargetSite: </strong>" + TargetSiteMsg;

            #endregion

            #region Generación de mensaje de error a la consola.
            string msgConsole =
                "\n\nSe ha registrado un error en la ejecución de este robot.\n\n" ;
            try
            {
                if (traceForGetters.GetFrame(0).GetFileLineNumber() != 0) //Verifica si el StackTrace tiene información útil.
                    msgConsole +=
                "Punto de referencia del error: \n"+
                "Clase: " + traceForGetters.GetFrame(0).GetFileName() + "\n" +
                "Línea: " + traceForGetters.GetFrame(0).GetFileLineNumber() + "\n" +
                "Método: " + traceForGetters.GetFrame(0).GetMethod() + "\n\n";
            }
            catch (Exception)
            {
                msgConsole += "No es posible brindar un punto de referencia del error exacto. ";
            }

            msgConsole +=
                "A continuación se detalla el tipo de error generado: \n\n" +
                "Tipo de error:\n " + ex.Message.ToString() + "\n\n" +
                "InnerException:\n" + InnerExceptionMsg + "\n" +
                "StackTrace:\n" + StackTraceMsg + "\n" +
                "TargetSite:\n" + TargetSiteMsg + "\n" +
                "\n \nSe restablecerá el sistema y continuará con el siguiente proceso. ";
            #endregion

            #region desactivar "activate" todas las clases
            crud.Update("UPDATE `orchestrator` SET `activate`= 0", "databot_db");
            #endregion
            using (ConsoleFormat console = new ConsoleFormat())
            {
                console.WriteAnnounce(msgConsole);
            }
            BlankProcess();
            //sap.BlockUser("ERP", 0, "300");
            //sap.BlockUser("ERP", 0, "260");
            //sap.BlockUser("ERP", 0, "500");
            //sap.BlockUser("ERP", 0, "460");
            //unblock sap all mandants
            if (area == "DM" || area == "ICS")
            {
                cc = new string[] { "dmeza@gbm.net", "smarin@gbm.net" };
                sender = "internalcustomersrvs@gbm.net";
            }
            else
            {
                cc = new string[] { "dmeza@gbm.net", "epiedra@gbm.net" };
                sender = "appmanagement@gbm.net";
            }
            using (MailInteraction mail = new MailInteraction())
            {
                mail.SendHTMLMail(msgEmail, new string[] { sender }, "He registrado un error - Databot" + root.Subject, cc);
            }
            //root.Dispose();
        }


        /// <summary>
        /// Este método envía un email especificando un error y lo imprime en consola a la vez.
        /// <list type="bullet">
        /// <item>
        /// <term>NameClass (obligatorio)</term>
        /// <description>String especificando la clase donde tiene el error. 
        /// También en lugar de enviar un String puede enviar la sentencia "this.GetType()", ya que el método tiene la capacidad de extraer el nombre de la clase por el type.</description>
        /// </item>
        /// <item>
        /// <term>Subject (opcional)</term>
        /// <description>El subject para el email.</description>
        /// </item>
        ///  <item>
        /// <term>Message (opcional)</term>
        /// <description>Mensaje para el email, puede utilizar HTML y opcionalmente si lo prefiere cadenas de texto claves para no repetir etiquetas de HTML como: lx para (br), ln para (strong), lz para (/strong)).</description>
        /// </item>
        ///  <item>
        /// <term>Exception (opcional)</term>
        /// <description>La excepción que pueda tener en un catch por ejemplo.</description>
        /// </item>
        /// </list>
        /// </summary>
        /// <remarks>
        /// Este método es diseñado para notificar a AppManagement sobre un error en alguna parte del proyecto. 
        /// Ejemplo: sendError(this.GetType(), "Ocurrió un error en el select", "Este es mi mensaje. lxlx Porfavor ln verifica lz", ex); 
        /// </remarks>
        public void SendError(string nameClass, string subject = "", string message = "", Exception exception = null)
        {

            message = message.Replace("lx", "<br>").Replace("ln", "<strong>").Replace("lz", "</strong>");
            Rooting root = new Rooting();
            string[] cc;
            string sender;
            bool classFounded = true;
            string msgEmail = "";

            #region Averiguar el nombre del Área y Proyecto de la clase.

            //Establece un valor por defecto en caso que venga vacío.
            nameClass = (nameClass.Replace(" ", "") == "") ? "Nombre de clase aportada en el error fue vacía" : nameClass;

            //Averigua a cual área pertenece la clase y dependiendo de ello enviar el error en diferentes personas encargadas.
            string sql = $@"SELECT databotAreas.area, orchestrator.projectName, orchestrator.class 
                            FROM `orchestrator`
                            INNER JOIN databotAreas ON databotAreas.id = orchestrator.area
                            WHERE class = '{nameClass.Replace(" ", "")}'";
            DataTable DtClass = crud.Select(sql, "databot_db");

            string area = "", projectName = "";
            try//nameClass está en DB.
            {
                area = DtClass.Rows[0]["area"].ToString();
                projectName = DtClass.Rows[0]["projectName"].ToString();
                nameClass = DtClass.Rows[0]["class"].ToString(); //Se refresca el nombre del nameClass, para que venga igual de la DB.
            }
            catch (Exception)
            { //Caso de que nameClass no esté en DB.
                nameClass += (string.IsNullOrWhiteSpace(nameClass)) ? "La clase enviada es vacía." :
                    " (no existe en la Tabla Orchestrator).";
                classFounded = false;
            }

            #endregion

            #region Generación de mensaje de error Email.

            if (!string.IsNullOrWhiteSpace(subject)) //El subject no viene vacío.
            {
                subject += (classFounded == true) ? ($" - {nameClass} - Databot") : ($" - Databot");
            }
            else
            {
                subject = (classFounded == true) ? ($"Error registrado - {nameClass} - Databot") : ($"Error registrado - Databot");
            }

            //Palabras clave reemplazarpara etiquetas de HTML en message.
            message = message.Replace("lx", "<br>").Replace("ln", "<strong>").Replace("lz", "</strong>");

            //En caso de que el nameClass venga vacío imprime condiciones en msgEmail en cada caso.
            if (classFounded)
                msgEmail = $"Clase o Robot: <strong>{nameClass}</strong> <br>" +
                       $"Proyecto: <strong> {projectName} </strong> <br>" +
                       $"Área: <strong>{area}</strong> <br><br>{message}<br>";
            else //Es vacío
                msgEmail = $"Clase o Robot: <strong>{nameClass}</strong> <br><br>{message}<br>";

            if (exception != null) //En los parámetros viene la excepción.
            {
                try
                {   //Verifica si el StackTrace del Exception no está vacio y su info.
                    if (new StackTrace(exception, true).GetFrame(0).GetFileLineNumber() != 0)
                        msgEmail += "<br><br>Punto de referencia del error: " +
                                 "<br>Clase: " + new StackTrace(exception, true).GetFrame(0).GetFileName() +
                                 "<br>Línea: " + new StackTrace(exception, true).GetFrame(0).GetFileLineNumber() +
                                 "<br>Método: " + new StackTrace(exception, true).GetFrame(0).GetMethod() + "<br>";
                }
                catch (Exception) { msgEmail += "<br><br>No es posible brindar un punto de referencia del error exacto. <br>"; }

                //Valida si la excepción no está vacía
                string InnerExceptionMsg = (exception.InnerException != null) ? exception.InnerException.ToString() : "";
                string StackTraceMsg = (exception.StackTrace != null) ? exception.StackTrace.ToString() : "";
                string TargetSiteMsg = (exception.TargetSite != null) ? exception.TargetSite.ToString() : "";

                msgEmail += "<br>A continuación se detalla el tipo de error generado: <br><br>" +

                             "<strong>Tipo de excepción:</strong>  " + exception.Message.ToString() + "<br><br>" +
                             "<strong>InnerException:</strong>" + InnerExceptionMsg + "<br><br>" +
                             "<strong>StackTrace: </strong>" + StackTraceMsg + "<br><br>" +
                             "<strong>TargetSite: </strong>" + TargetSiteMsg;
            }


            #endregion

            #region Generación de mensaje de error a la consola.
            string msgConsole = "\n\nSe ha registrado un error en la ejecución de este robot.\n\n";
            msgConsole += msgEmail;
            msgConsole = msgConsole.Replace("<br>", "\n").Replace("</br>", "\n").Replace("<strong>", "").Replace("</strong>", "");

            console.WriteAnnounce(msgConsole);
            #endregion

            //Despedida
            msgEmail += "<br><br> Databot.";
            #region Envío de correo a Soporte según área.

            if (area == "DM" || area == "ICS")
            {
                cc = new string[] { "dmeza@gbm.net", "smarin@gbm.net" };
                sender = "internalcustomersrvs@gbm.net";
            }
            else
            {
                cc = new string[] { "dmeza@gbm.net", "epiedra@gbm.net" };
                sender = "appmanagement@gbm.net";
            }
            using (MailInteraction mail = new MailInteraction())
            {
                 mail.SendHTMLMail(msgEmail, new string[] { sender }, subject, cc);
            }
            #endregion
        }

        /// <summary>
        /// Este método envía un email especificando un error y lo imprime en consola a la vez.
        /// <list type="bullet">
        /// <item>
        /// <term>NameClass (obligatorio)</term>
        /// <description>Campo tipo Type, enviar la sentencia "this.GetType()", ya que el método tiene la capacidad de extraer el nombre de la clase por el type.</description>
        /// </item>
        /// <item>
        /// <term>Subject (opcional)</term>
        /// <description>El subject para el email.</description>
        /// </item>
        ///  <item>
        /// <term>Message (opcional)</term>
        /// <description>Mensaje para el email, puede utilizar HTML y cadenas de texto claves para no repetir etiquetas de HTML como: lx para (br), ln para (strong), lz para (/strong)).</description>
        /// </item>
        ///  <item>
        /// <term>Exception (opcional)</term>
        /// <description>La excepción que pueda tener en un catch por ejemplo.</description>
        /// </item>
        /// </list>
        /// </summary>
        /// <remarks>
        /// Este método es diseñado para notificar a AppManagement sobre un error en alguna parte del proyecto. 
        /// Ejemplo: sendError(this.GetType(), "Ocurrió un error en el select", "Este es mi mensaje. lxlx Porfavor ln verifica lz", ex); 
        /// </remarks>
        public void SendError(System.Type typeClass, string subject = "", string message = "", Exception ex = null)
        {
            //Mismo método SendError con constructores distintos.
            string nameClass = typeClass.Name;
            SendError(nameClass, subject, message, ex);
        }
        public void BreakProgram()
        {
            using (ConsoleFormat console = new ConsoleFormat())
            using (ProcessAdmin processAdmin = new ProcessAdmin())
            {
                console.WriteLine("Taking a break");
                BlankProcess();
                processAdmin.DeleteFiles(root.FilesDownloadPath);
                //root.Dispose();

                System.Threading.Thread.Sleep(20000);
                console.WriteLine("Back to duties");
                console.WriteAnnounce("");
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        #region Planner

        /// <summary>
        /// Este método se encarga de consultar la tabla orchestrator para extraer la clase y su respectivo planificador,
        /// esto con dos finalidades: 
        /// 1-Almacena en una variable *estática* (root.planner[@class]) el planificador de TODAS las clases de la área seleccionada por el usuario, con 
        /// variables adicionales como conta y datenow, el cual ayuda a la gestión de ejecución del proceso y poder ejecutar cada ciertos minutos o no correrlo
        /// más de una vez en un lapso determinado.
        /// 2-Almacenar en una variable *estática* (root.planiChange) nuevamente el planificador, pero en una variable tipo string, esto para en la ejecución del 
        /// robot poder comparar y verificar si en algún momento el planificador cambió y realizar acciones respectivas.
        /// </summary>
        public void BuildPlanner(string enviroment)
        {
            console.WriteLine("Inicializar variable del Planificador");
            string sql_select, @class, minutes, hours, weeks, months;
            int rows;
            DataTable mytable = new DataTable();
            DataTable mytable2 = new DataTable();
            try
            {
                #region Connection DB     
                sql_select = "SELECT planner, class FROM `orchestrator` WHERE JSON_LENGTH(planner) > 0";
                mytable = crud.Select(sql_select, "databot_db");
                #endregion
                rows = mytable.Rows.Count;

                root.planiChange = JsonConvert.SerializeObject(mytable); //Se utiliza para camparar si hay un cambio en la base de datos

                if (rows > 0)
                {
                    //Llenar el dictionary.
                    for (int i = 0; i < rows; i++)
                    {
                        List<string[]> lista_de_valores = new List<string[]>(3);
                        Dictionary<string, string[]> info = new Dictionary<string, string[]>();
                        info.Add("minutes", null);
                        info.Add("dateNow", null);
                        info.Add("hours", null);
                        info.Add("week", null);
                        info.Add("month", null);
                        info.Add("type", null);
                        info.Add("conta", null);


                        Newtonsoft.Json.Linq.JObject planner = Newtonsoft.Json.Linq.JObject.Parse(mytable.Rows[i]["planner"].ToString());


                        @class = mytable.Rows[i]["class"].ToString();
                        minutes = "";
                        try
                        {
                            minutes = planner["Minutes"].ToString();
                        }
                        catch (Exception)
                        {
                        }


                        hours = planner["Hour"].ToString();
                        weeks = planner["Week"].ToString();
                        months = planner["Month"].ToString();

                        string types = planner["Type"].ToString();

                        info["type"] = new string[] { types };
                        info["conta"] = new string[] { "0" };

                        if (types == "0")  //solo minutos
                        {
                            info["minutes"] = new string[] { minutes };
                            info["dateNow"] = new string[] { DateTime.UtcNow.ToString() };
                            root.planner[@class] = info;

                        }
                        else if (types == "1")  //solo horas

                        {
                            List<string> list = hours.Split(',').Select(hh => hh.Trim()).ToList();
                            info["hours"] = list.ToArray();

                            root.planner[@class] = info;
                        }
                        else if (types == "2") //dias de la semana
                        {
                            List<string> list = hours.Split(',').Select(hh => hh.Trim()).ToList();
                            info["hours"] = list.ToArray();
                            list.Clear();
                            list = weeks.Split(',').Select(hh => hh.Trim()).ToList();
                            info["week"] = list.ToArray();

                            root.planner[@class] = info;
                        }
                        else if (types == "3") //dia del mes
                        {
                            List<string> list = hours.Split(',').Select(hh => hh.Trim()).ToList();
                            info["hours"] = list.ToArray();
                            list.Clear();
                            list = months.Split(',').Select(hh => hh.Trim()).ToList();
                            info["month"] = list.ToArray();

                            root.planner[@class] = info;
                        }
                        else //cualquier otra opcion no contemplada
                        {
                            //error
                        }

                    }
                }

            }
            catch (Exception ex)
            { console.WriteLine(ex.Message); }

        }
        /// <summary>
        /// Este método fue creado para verificar si existe un cambio en el planificador de la tabla orchestrator, lo hace comparando con la variable 
        /// estática tipo String root.planiChange, el cual almacena el planificador generado inicialmente en BuildPlanner, si es así refresca los nuevos cambios.
        /// </summary>
        public void ReviewChangesPlanner(string enviroment)
        {

            DataTable mytable = new DataTable();
            string sql_select = "SELECT planner, class FROM `orchestrator` WHERE JSON_LENGTH(planner) > 0";
            mytable = crud.Select(sql_select, "databot_db");

            string string1 = Newtonsoft.Json.JsonConvert.SerializeObject(mytable);

            if (string1 != root.planiChange)
            {
                console.WriteLine("El planificador cambió.");
                BuildPlanner(enviroment);
            }
        }
        /// <summary>
        /// Este método fue creado para válidar si un proceso en base a su planificador, puede ejecutarse comparando con la hora y fecha actual, 
        /// esto según su tipo: 
        /// 0= Cantidad de minutos despúes de la última ejecución.
        /// 1= Horas. 
        /// 2= Días de la semana .
        /// 3= Día del mes.
        /// Este tipo está establecido en el JSON del planner del proceso en la tabla orchestrator.
        /// </summary>
        /// <param name="waiting"></param>
        /// <param name="process"></param>
        /// <returns></returns>
        public bool Planner(int waiting, string process)
        {
            Rooting root = new Rooting();
            bool run = false;

            Dictionary<string, string[]> info = root.planner[process];


            string sql_select;
            DataTable mytable = new DataTable();


            string[] minutes = info["minutes"];
            string[] hours = info["hours"];
            string[] week = info["week"];
            string[] month = info["month"];
            string type = info["type"][0];
            string conta_s = info["conta"][0];
            plannerResponse plannerResponse = new plannerResponse();

            if (type == "0")
            {
                //el de minutos
                DateTime end = DateTime.Parse(info["dateNow"][0]).AddMinutes(int.Parse(minutes[0]));
                if (DateTime.UtcNow > end)
                {
                    info["dateNow"] = new string[] { DateTime.UtcNow.ToString() };
                    root.planner[process] = info;
                    run = true;
                }

            }
            else if (type == "1")
            {
                //el de horas
                plannerResponse = horas(conta_s, hours, waiting);
                run = plannerResponse.run;
                info["conta"][0] = plannerResponse.conta;
                root.planner[process] = info;
            }
            else if (type == "2")
            {
                //el de dia

                for (int i = 0; i < week.Length; i++)
                {
                    if ((int)DateTime.Now.DayOfWeek == int.Parse(week[i].Trim()))
                    {
                        plannerResponse = horas(conta_s, hours, waiting);
                        run = plannerResponse.run;
                        info["conta"][0] = plannerResponse.conta;
                        root.planner[process] = info;
                    }

                }

            }
            else if (type == "3")
            {
                //el de mes
                DateTime currentDate = DateTime.Now;

                for (int i = 0; i < month.Length; i++)
                {
                    int lastDayMonth = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);
                    int dayPlanner = int.Parse(month[i].Trim());

                    //En caso que el último día del mes sea 29, y el planner diga 31.
                    if(dayPlanner> lastDayMonth)
                    {
                        dayPlanner = lastDayMonth;
                    }

                    if (currentDate.Day == dayPlanner)
                    {
                        plannerResponse = horas(conta_s, hours, waiting);
                        run = plannerResponse.run;
                        info["conta"][0] = plannerResponse.conta;
                        root.planner[process] = info;
                    }

                }


            }
            else
            {
                //tipo desconocido
                run = false;
            }

      


            return run;
        }

        //La lógica de este método tiene un eje central en conta, conta_s y waiting.
        //conta_s almacena el número de la posición del array de la hora que es próxima ejecutar y así descarta las que ya ejecutó, ejemplo [0] 07:00:00 ó [1]16:00:00 ó [2]20:00:00.
        //conta es una variable auxiliar que ayuda a poner temporalmente la variable conta_s para gestión.
        //waiting es un lapso de tiempo que le da al proceso para ejecutarse, ejemplo la hora inicial es 7:00:00 el waiting de 10 minutos, entonces el proceso debe generarse en cualquier momento de ese rango (7 a 7:10).
        //La función de conta es asegurarse que el proceso corra UNA SOLA VEZ en el rango generado de waiting, al finalizar todas las horas, el conta se establece en 0 para que la otra ejecución empiece desde 0 nuevamente.
        public plannerResponse horas(string conta_s, string[] hours, int waiting )
        {
            plannerResponse plannerResponse = new plannerResponse();
            plannerResponse.run = false;
            plannerResponse.conta = conta_s;

            int conta = Int32.Parse(conta_s);

            DateTime datenow = DateTime.Now;
            string[] timepart = hours[hours.Length - 1].Split(new char[1] { ':' }); //17 :00: 00
            DateTime Sdate;
            DateTime Edate = new DateTime(datenow.Year, datenow.Month, datenow.Day, int.Parse(timepart[0]), int.Parse(timepart[1]), int.Parse(timepart[2])).AddMinutes(waiting); //04/07/2022 17:00:00


            if (conta == hours.Length && datenow > Edate) //reset conta después de la última ejecución.
            {
                conta = 0; //lo pone en cero para que el for no entre ya que ejecuto la ultima ejecucion
                conta_s = conta.ToString(); //esto es para ponerlo en string nada mas
                plannerResponse.conta = conta_s;
            }
            // [13:00:00, 17:00:00, 20:00:000]
            // contador = 0, 1, 2
            for (int i = conta; i < hours.Length; i++) // Si ya ejecutó una hora de la lista, el for no lo toma en cuenta más hasta el reset para no ejecutarlo más si aún está en el rango de tiempo.
            {
                datenow = DateTime.Now;
                timepart = hours[i].Split(new char[1] { ':' });
                Sdate = new DateTime(datenow.Year, datenow.Month, datenow.Day, int.Parse(timepart[0]), int.Parse(timepart[1]), int.Parse(timepart[2]));
                Edate = Sdate.AddMinutes(waiting);

                if (Edate > datenow && datenow > Sdate) // Verifica si está en rango de hora establecida y el waiting.
                {
                    console.WriteLine("Si se debe de ejecutar el bot");
                    conta = i + 1;
                    conta_s = conta.ToString();
                    plannerResponse.run = true;
                    plannerResponse.conta = conta_s;
                }
            }

            return plannerResponse;

        }
        /// <summary>
        /// Metodo que resta un 1 al contador del planner de una clase en especifico para que se vuelva a ejecutar, esto cuando el mandante de SAP esta ocupado
        /// </summary>
        /// <param name="process"></param>
        /// <returns></returns>
        public bool setPlannerAgain()
        {

            Dictionary<string, string[]> planner = root.planner[root.BDClass];
            string conta_s = planner["conta"][0];
            int conta = Int32.Parse(conta_s);
            conta -= 1;
            conta_s = conta.ToString();
            planner["conta"][0] = conta_s;
            root.planner[root.BDClass] = planner;
            return false;
        }
        #endregion

        #region Dispose
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~Settings()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
    public class plannerResponse
    {
        public string conta { get; set; }
        public bool run { get; set; }
    }

}
