using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Files;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Webex;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DataBotV5.App.ConsoleApp
{
    /// <summary>
    /// Clase principal que corre todos los procesos RPA
    /// </summary>
    class Start
    {

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(System.IntPtr hWnd, int cmdShow);
        public int y;
        public static string enviroment;
        [STAThread]

        static void Main()
        {
            #region Inicializar objetos globales
            Queue<int> cola = new Queue<int>();
            ValidDataTable vdt = new ValidDataTable();
            Start me = new Start();
            Rooting root = new Rooting();
            Credentials cred = new Credentials();
            string area = string.Empty;

            #endregion
            #region Opciones de la consola
            MaximizeConsole();
            using (Settings setts = new Settings
            {
                color = ConsoleColor.Green,
                title = "Databot a GBM MIS Automation App - All rights reserved."
            })
            {
                setts.PrintLogo();
                //setts.PrintEnviroment(enviroment);
                setts.BlankProcess();
                #region Comprobar carpeta de archivos

                if (!setts.Carpeta())
                {
                    setts.EndProgram("Carpeta del Databot no detectada");

                }
                #endregion
            }
            #endregion
            #region Cola de procesos
            int colaC = 0;
            //Crea la cola de los procesos a correr con la cantida de procesos RPA existentes por tipo
            using (RPAScheduler orquestador = new RPAScheduler())
            {
                /*Importante establecer el entorno inicial de ejecución:
                -"PRD": apunta al 10.7.60.138
                -"QAS" apunta al 10.7.60.151
                -"DEV": apunta al localhost
               Esto establece en cual BD realiza las consultas a Orchestrathor (procesos, areas, planificador)*/

                enviroment = orquestador.SelectEnviroment();

                //Set the email that the robot will search and sent emails
                if (enviroment == "QAS" || enviroment == "DEV")
                {
                    root.Direccion_email = "databotqa@gbm.net";
                    cred.passOutlook = cred.password_outlook_qa;
                }
                else
                {
                    root.Direccion_email = "databot@gbm.net";
                    cred.passOutlook = cred.password_outlook;
                }

                //Seleccione el área que desea ejecutar (databotAreas de databot_db)
                area = orquestador.Robot(enviroment);
    
                vdt = orquestador.Quantity(area, enviroment);
                int cprocesos = vdt.DataTableValue.Rows.Count;
                for (int i = 0; i < cprocesos; i++)
                {
                    cola.Enqueue(i);
                }
                colaC = cola.Count;
            }
            #endregion
            #region Loop until Esc is pressed
            if (vdt.ValidTable)
            {

                using (Settings sett = new Settings())
                {
                    //verifica si se puede ejecutar, true = esta dentro de los parametros del json Planificador de la tabla orquestador
                    sett.BuildPlanner(enviroment);
                }

                while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape))
                {

                    #region Refrescar datos del ValidDataTable y revisar si hay cambios en Planificador.
                    //Refresca la tabla ValidDataTable para verificar si hubo un cambio, especialmente en Active o Activate.
                    using (RPAScheduler orquestador = new RPAScheduler())
                    {
                        vdt = orquestador.Quantity(area, enviroment);
                    }

                    //Verificar si hay algún cambio en el planificador, si es así construirlo nuevamente en rooting..
                    using (Settings sett = new Settings())
                    {
                        sett.ReviewChangesPlanner(enviroment);
                    }
                    #endregion

                    #region Foreach Row of ValidDataTable.DataTableValue
                    vdt.DataTableValue.Rows.Cast<DataRow>().ToList().ForEach(DataRow =>
                    {
                        try
                        {

                            me.y = cola.Dequeue();
                            int hnow = DateTime.Now.Hour;
                            if (hnow != 3 && hnow != 4) //No corre en la madrugada dateNow > Sdate && dateNow < Fdate
                            {
                                RunMethod(DataRow, enviroment);
                            }
                            cola.Enqueue(me.y);
                            if (vdt.DataTableValue.Rows.IndexOf(DataRow) == vdt.DataTableValue.Rows.Count - 1)
                            {
                                //la ultima fila de la tabla
                                using (Settings setts = new Settings())
                                {
                                    setts.BreakProgram();
                                }
                            }

                            if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape)

                            {
                                using (Settings setts = new Settings())
                                {
                                    setts.PauseProgram();
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                            root.requestDetails = "";
                            root.BDUserCreatedBy = "";

                            using (Settings setts = new Settings())
                            {
                                setts.ErrorProgram(ex, DataRow, area, enviroment);
                            }
                            cola.Enqueue(me.y);
                        }
                    }); //foreach
                    #endregion

                } //while
            }
            #endregion

        }
        /// <summary>
        /// Metodo para maximizar el tamaño de la consola
        /// </summary>
        private static void MaximizeConsole()
        {
            Process p = Process.GetCurrentProcess();
            ShowWindow(p.MainWindowHandle, 3);
        }
        /// <summary>
        /// Llama a la clase y metodo Main
        /// </summary>
        /// <param name="x"></param>
        /// <param name="area"></param>
        public static void RunMethod(DataRow x, string enviroment)
        {

            if (Convert.ToBoolean(x["active"])) //determina si el proceso esta activo o no
            {

                using (ConsoleFormat console = new ConsoleFormat())
                {
                    console.WriteLine($"Starting process: {x["class"]}");
                }
                bool planning = true;
                bool activate = Convert.ToBoolean(x["activate"]);
                if (x["planner"].ToString() != "{}" && !activate) //Determina si la clase tiene planificador
                {
                    Newtonsoft.Json.Linq.JObject planner = Newtonsoft.Json.Linq.JObject.Parse(x["planner"].ToString());

                    //verifica si se puede ejecutar, true = esta dentro de los parametros del json Planificador de la tabla orquestador
                    using (Settings set = new Settings())
                    {
                        planning = set.Planner(int.Parse(planner["Waiting"].ToString()), x["class"].ToString());

                    }

                }
                using (ConsoleFormat console = new ConsoleFormat())
                using (Rooting root = new Rooting())

                {
                    if (planning || activate)
                    {
                        root.BDProcess = x["projectName"].ToString();
                        root.BDIdClass = x["id"].ToString();
                        root.BDClass = x["class"].ToString();
                        root.BDArea = x["areaCode"].ToString();
                        root.BDStartDate = DateTime.Now; //se utiliza para las estadisticas
                        root.BDUserCreatedBy = "";
                        root.requestDetails = "";
                        root.BDActivate = Convert.ToBoolean(x["activate"]);

                        console.WriteLine($"Starting process: {x["class"]}");
                        //busca la clase en todo el proyecto, el problema es que no pueden haber 2 clases con el mismo nombre, el trae el primero que encuentra
                        //procedure para no insertar 2 clases con el mismo nombre en orquestador DB
                        //buscar solo en la carpeta de automation

                        //busca la clase de una forma más estatica por carpeteo
                        Type clase = Type.GetType($"DataBotV5.Automation.{root.BDArea}.{root.BDProcess}.{root.BDClass}");

                        ConstructorInfo Constructor = clase.GetConstructor(Type.EmptyTypes);
                        object ClassObject = Constructor.Invoke(new object[] { }); //aqui inicializa la clase

                        MethodInfo Method = clase.GetMethod("Main");

                        object Value = Method.Invoke(ClassObject, new object[] { }); //aqui llama el metodo, dentro del {} se pone los parametros
                                                                                     //root.Dispose(true);
                        if (activate)
                        {
                            using (RPAScheduler orquestador = new RPAScheduler())
                            {
                                orquestador.DisableProcess(root.BDClass, enviroment);
                            }
                        }

                        root.requestDetails = "";
                        root.BDUserCreatedBy = "";
                        root.CopyCC = null;


                        using (DestroyProcess kill = new DestroyProcess())
                        {
                            kill.KillProcess("chromedriver", true);
                            kill.KillProcess("EXCEL", true);
                            //kill.KillProcess("chrome", true);
                        }

                    }
                    else
                    {
                        //System.Threading.Thread.Sleep(60000);
                        Newtonsoft.Json.Linq.JObject planner = Newtonsoft.Json.Linq.JObject.Parse(x["planner"].ToString());
                        using (Settings set = new Settings())
                        {
                            planning = set.Planner(int.Parse(planner["Waiting"].ToString()), x["class"].ToString());

                        }
                        if (planning)
                        {
                            using (WebexTeams wx = new WebexTeams())
                            {
                                wx.SendNotification("dmeza@gbm.net", "Robot no se ejecutó", $"No se ejecuto el proceso del robot: {x["class"].ToString()}");
                            }
                        }
                    }

                    console.WriteAnnounce("Finishing process: " + x["class"]);
                }
            }
        }
    }
}
