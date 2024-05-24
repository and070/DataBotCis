using DataBotV5.App.Global;
using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using MySql.Data.MySqlClient;
using DataBotV5.Data.Database;

namespace DataBotV5.Data.Process
{
    /// <summary>
    /// Clase Data que contiene los métodos necesarios para correr el queue del Console App.
    /// </summary>
    class RPAScheduler : IDisposable
    {
        private bool disposedValue;


        /// <summary>
        /// Método de cantidad de procesos.
        /// </summary>
        /// <returns>Retorna un ValidDataTable (Que contiene un DataTable y una Variable Bool) con la cantidad de procesos.</returns>
        public ValidDataTable Quantity(string area, string enviroment)
        {
            ValidDataTable vdt = new ValidDataTable();
            string query = "";
            if (area == "QAS")
            {
                query = $@"SELECT 
                                    orchestrator.*,
                                    databotAreas.area as areaCode
                                    FROM `orchestrator` 
                                    INNER JOIN databotAreas ON databotAreas.id = orchestrator.area";
            }
            else
            {
                query = $@"SELECT 
                                    orchestrator.*,
                                    databotAreas.area as areaCode
                                    FROM `orchestrator` 
                                    INNER JOIN databotAreas ON databotAreas.id = orchestrator.area
                                    WHERE databotAreas.area = '{area}'";
            }
            using (CRUD crud = new CRUD())
            {
                vdt = crud.Info(query, "databot_db", enviroment);
            }
            return vdt;
        }
        public string SelectEnviroment()
        {
            using (ConsoleFormat console = new ConsoleFormat())
            {
                console.WriteAnnounce("Bienvenido a la consola del Databot V5.0.0");
                console.WriteLine("Para iniciar, indique el mandante a ejecutar:");
                List<string> enviromentsOptions = new List<string>
                {
                    "PRD",
                    "QAS",
                    "DEV"
                };
                int cont = 1;
                string[] enviromentsText =
                {
                    "1) \t PRD \t Ambiente producción (.138, 300, 500)",
                    "2) \t QAS \t Ambiente calidad (.151, 260, 460)",
                    "3) \t DEV \t Ambiente desarrollo (localhost, 120, 420)"
                };
                foreach (string enviroment in enviromentsText)
                {
                    console.WriteLine(enviroment);
                }
                //Lee la selección del usuario.
                string seleccion = Console.ReadLine();
                seleccion = seleccion.ToUpper().Trim();

                //Valida si la selección del usuario fue correcta.
                bool correctSelection = true;
                do
                {

                    string result = enviromentsOptions.SingleOrDefault(s => s == seleccion);
                    if (string.IsNullOrEmpty(result))
                    {
                        console.WriteLine($"El comando: {seleccion} no es válido, favor seleccionar uno de la lista anterior.");
                        seleccion = Console.ReadLine();
                        seleccion = seleccion.ToUpper().Trim();
                        correctSelection = false;
                    }
                    else
                    {
                        correctSelection = true;
                    }
                }
                while (!correctSelection);
                return seleccion;
            }
        }
        /// <summary>
        /// Método que consulta de cual área de operación va a ejecutar en el robot.
        /// </summary>
        /// <returns>Devuelve un String con el área seleccionada por el usuario.</returns>
        public String Robot(string enviroment)
        {
            string seleccion = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB
                string sql_select = "";
                if (enviroment == "PRD")
                {
                    sql_select = "select area, areaText from `databotAreas` WHERE area != 'QAS'";
                }
                else if (enviroment == "DEV")
                {
                    sql_select = "select area, areaText from `databotAreas`";
                }
                else if (enviroment == "QAS") 
                {
                    sql_select = "select area, areaText from `databotAreas` WHERE area = 'QAS'";
                }
                mytable = new CRUD().Select(sql_select, "databot_db");
                #endregion

                if (mytable.Rows.Count > 0) //Verifica si hay procesos de automatización.
                {
                    using (ConsoleFormat console = new ConsoleFormat())
                    {
                        console.WriteAnnounce("");
                        console.WriteLine("Indique el área de procesamiento de automatización a ejecutar: ");
                        //Imprime todos los procesos que están en la tabla orchestrator de procesos.
                        for (int i = 0; i < mytable.Rows.Count; i++)
                        {
                            string area = mytable.Rows[i]["area"].ToString();
                            console.WriteLine((i + 1) + ") \t" + area + " - " + mytable.Rows[i]["areaText"]);
                        }
                        //Lee la selección del usuario.
                        seleccion = Console.ReadLine();
                        seleccion = seleccion.ToUpper().Trim();

                        //Valida si la selección del usuario fue correcta.
                        bool correctSelection = false;
                        do
                        {
                            for (int i = 0; i < mytable.Rows.Count; i++) //Recorre el Datable para verificar si hay un valor igual a la selección.
                            {
                                if (seleccion == (mytable.Rows[i]["area"]).ToString())
                                    correctSelection = true;
                            }
                            if (!correctSelection)
                            {
                                console.WriteLine($"El comando: {seleccion} no es válido, favor seleccionar uno de la lista anterior.");
                                seleccion = Console.ReadLine();
                                seleccion = seleccion.ToUpper().Trim();
                            }
                        }
                        while (!correctSelection);
                    }
                }
                else
                {
                    using (ConsoleFormat console = new ConsoleFormat())
                    { console.WriteLine("Actualmente no existen procesos RPA."); }
                }
            }
            catch (Exception)
            {
            }
            return seleccion;
        }
        public void DisableProcess(string @class, string enviroment)
        {
            using (CRUD crud = new CRUD())
            {
                crud.Update($"UPDATE `orchestrator` SET `activate`= 0 WHERE `class` = '{@class}'", "databot_db");
            }
        }
        #region dispose
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
        // ~RPAScheduler()
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

        /*AGREGADO LOGEO DE ERRORES POR EDUARDO PIEDRA PENDIENTE APROBACION SI ESTA EN EL LUGAR CORRECTO */


    }
    #region Clases del orquestador para Json Parse
    public class Scheduler
    {
        /// <summary>
        /// Lista con las horas del día donde se debe de ejecutar este proceso.
        /// </summary>
        public List<string> DayHours { get; set; }
        /// <summary>
        /// Lista de tipo entero con los días de la semana donde se debe de ejecutar el proceso, donde 0 es Domingo, 1 es Lunes, 2 es Martes, 3 es Miercoles, 4 es Jueves 5 es viernes y 6 es Sabado.
        /// </summary>
        public List<int> DaysWeek { get; set; }
        /// <summary>
        /// Lista de tipo entero con los días del mes donde se debe de ejecutar el proceso.
        /// </summary>
        public List<int> DayMonth { get; set; }
    }
    #endregion


}
