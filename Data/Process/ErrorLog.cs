using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using MySql.Data.MySqlClient;
using System;

namespace DataBotV5.Data.Process
{
    class ErrorLog
    {
        ConsoleFormat console = new ConsoleFormat();
        /// <summary>
        /// Registra los errores producidos durante la ejecución de un proceso para futuro análisis y auditoría.
        /// </summary>
        /// <param name="process">Nombre del proceso en el cual hubo un error.</param>
        /// <param name="error">Descripción detallada del error.</param>
        public void LogErrors(string process, string error)
        {

            //string query;

            //try
            //{
                
            //    query = "INSERT INTO error_log (PROCESO, ERROR) VALUES ('"+ process + "', CURRENT_TIMESTAMP, '"+ error +"')";
            //    new CRUD().Insert("Databot", query, "databot_db");
            //}
            //catch (Exception)
            //{
            //    console.WriteLine(" Se ha producido un error ingresando reporte a log de errores");
            //}

        }

    }
}
