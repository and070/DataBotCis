using DataBotV5.App.ConsoleApp;
using DataBotV5.App.Global;
using DataBotV5.App.Global.Interfaces;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Runtime.InteropServices;

namespace DataBotV5.Data.Database
{
    /// <summary>
    /// Clase Data encargada de CRUD.
    /// </summary>
    class CRUD : IDisposable
    {
        private bool disposedValue;
        ConsoleFormat console = new ConsoleFormat();
        /// <summary>
        /// Obtiene la data table y valida si contiene objeto y si no es nulla
        /// </summary>
        /// <param name="sql">SQL Query</param>
        /// <param name="database">Database</param>
        /// <param name="ambient">Ambiente</param>
        /// <returns>Retorna la Data Table y un booleano que indica si no tiene objetos y si no es null</returns>
        public ValidDataTable Info(string sql, string database, [Optional] string ambient)
        {
            if(ambient == null)
            {
                ambient = Start.enviroment;
            }

            DataTable dt = new DataTable();
            using (Database db = new Database())
            {
                MySqlConnection conn = db.ConnSmartSimple(database, ambient);
                MySqlDataAdapter myadapter = new MySqlDataAdapter(sql, conn);
                myadapter.Fill(dt);

            }
            bool valid;
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    valid = true;
                }
                else
                {
                    valid = false;
                }
            }
            else
            {
                valid = false;
            }

            ValidDataTable dataTable = new ValidDataTable()
            {
                DataTableValue = dt,
                ValidTable = valid

            };


            return dataTable;
        }
        /// <summary>
        /// Obtiene el datatable del SELECT command. Instance: "SmartAndSimple" ó "Databot". Database: base de datos. Sql: sqlquery. Ambient: "PRD", "QA", "DEV".
        /// </summary>
        /// <param name="sql">SELECT para extraer el datatable</param>
        /// <param name="database">Base de datos a consultar</param>
        /// <param name="ambient">string DEV, QAS, o PRD</param>
        /// <returns>Retorna un objeto DataTable.</returns>
        public DataTable Select(string sql, string database, [Optional] string ambient)
        {
            if (ambient == null)
            {
                ambient = Start.enviroment;
            }
            DataTable dt = new DataTable();
            try
            {
                MySqlConnection conn = new MySqlConnection();

                conn = new Database().ConnSmartSimple(database, ambient);

                MySqlDataAdapter myadapter = new MySqlDataAdapter(sql, conn);
                myadapter.Fill(dt);
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
            }
            return dt;
        }

        /// <summary>
        /// Update en la base de datos. Instance: "SmartAndSimple" ó "Databot". Sql: sqlquery. Database: base de datos. Ambient: "PRD", "QA", "DEV".
        /// </summary>
        /// <param name="sql">Sentencia de update</param>
        /// <param name="database">Nombre de la base de datos</param>
        public bool Update(string sql, string database, [Optional] string ambient)
        {
            if (ambient == null)
            {
                ambient = Start.enviroment;
            }
            try
            {
                MySqlConnection conn = new MySqlConnection();
                conn = new Database().ConnSmartSimple(database, ambient);
                conn.Open();
                MySqlCommand execute = new MySqlCommand(sql, conn);
                execute.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                return false;
            }
        }


        /// <summary>
        /// Insertar nuevo valor en la base de datos. Instance: "SmartAndSimple" ó "Databot". Sql: sqlquery. Database: base de datos. Ambient: "PRD", "QA", "DEV".
        /// </summary>
        /// <param name="sql">Sentencia de insert</param>
        /// <param name="database">Nombre de la base de datos</param>
        /// <param name="Overload1"></param>
        public bool Insert(string sql, string database, [Optional] string ambient)
        {
            if (ambient == null)
            {
                ambient = Start.enviroment;
            }
            try
            {
                MySqlConnection conn = new MySqlConnection();
                conn = new Database().ConnSmartSimple(database, ambient);
                conn.Open();
                MySqlCommand execute = new MySqlCommand(sql, conn);
                execute.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// Eliminar row en la base de datos. Instance: "SmartAndSimple" ó "Databot". Sql: sqlquery. Database: base de datos. Ambient: "PRD", "QA", "DEV".
        /// </summary>
        /// <param name="instance">Instancia "SmartAndSimple" ó "Databot"</param>
        /// <param name="sql">Sentencia de insert</param>
        /// <param name="database">Nombre de la base de datos</param>
        /// <param name="ambient">Ambiente "PRD","QA","DEV"</param>
        public bool Delete(string sql, string database, [Optional] string ambient)
        {
            if (ambient == null)
            {
                ambient = Start.enviroment;
            }
            try
            {
                MySqlConnection conn = new MySqlConnection();
                conn = new Database().ConnSmartSimple(database, ambient);

                conn.Open();
                MySqlCommand execute = new MySqlCommand(sql, conn);
                execute.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                return false;
            }
        }
        /// <summary>
        /// Realizar un INSERT, UPDATE, DELETE y retornar el id autoincrementable del cambio
        /// </summary>
        /// <param name="instance">Instancia "SmartAndSimple" ó "Databot"</param>
        /// <param name="sql">Sentencia de insert</param>
        /// <param name="database">Nombre de la base de datos</param>
        /// <param name="ambient">Ambiente "PRD","QA","DEV"</param>
        public long NonQueryAndGetId(string sql, string database, [Optional] string ambient)
        {
            if (ambient == null)
            {
                ambient = Start.enviroment;
            }
            long id = 0;
            try
            {
                MySqlConnection conn = new MySqlConnection();

                conn = new Database().ConnSmartSimple(database, ambient);
                conn.Open();
                MySqlCommand execute = new MySqlCommand(sql, conn);
                execute.ExecuteNonQuery();
                id = execute.LastInsertedId;
                conn.Close();
            }
            catch (Exception) { }
            return id;
        }
        /// <summary>
        /// Realizar un INSERT, UPDATE, DELETE y retornar el id autoincrementable del cambio ademas de true si salio bien o false
        /// </summary>
        /// <param name="instance">Instancia "SmartAndSimple" ó "Databot"</param>
        /// <param name="sql">Sentencia de insert</param>
        /// <param name="database">Nombre de la base de datos</param>
        /// <param name="ambient">Ambiente "PRD","QA","DEV"</param>
        /// <returns></returns>
        public queryAndId NonQueryAndGetIdAndStatus(string sql, string database, [Optional] string ambient)
        {
            if (ambient == null)
            {
                ambient = Start.enviroment;
            }
            queryAndId queryAndId = new queryAndId();
            long id = 0;
            try
            {
                MySqlConnection conn = new MySqlConnection();
                conn = new Database().ConnSmartSimple(database, ambient);
                conn.Open();
                MySqlCommand execute = new MySqlCommand(sql, conn);
                execute.ExecuteNonQuery();
                id = execute.LastInsertedId;
                conn.Close();
                if (id == 0)
                {
                    queryAndId.validate = false;
                    queryAndId.id = id;
                    return queryAndId;
                }
                queryAndId.validate = true;
                queryAndId.id = id;
                return queryAndId;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                queryAndId.validate = false;
                queryAndId.id = 0;
                return queryAndId;

            }
        }


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
        // ~CRUD()
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
    }
    public class ValidDataTable : IDataTable
    {
        public DataTable DataTableValue { get; set; }
        public bool ValidTable { get; set; }
    }
    public class DataTableInfo : IDataTable
    {
        public DataTable DataTableValue { get; set; }
        public bool ValidTable { get; set; }
        public int Rows { get; set; }
    }

    public class queryAndId
    {
        public long id { get; set; }
        public bool validate { get; set; }
    }

}
