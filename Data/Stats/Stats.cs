using System;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using DataBotV5.Data.Root;
using DataBotV5.Data.Database;
using DataBotV5.App.ConsoleApp;
using DataBotV5.App.Global;

namespace DataBotV5.Data.Stats
{
    /// <summary>
    /// Clase Data creada para gestionar las Estadísticas del proyecto.
    /// </summary>
    class Stats : IDisposable
    {
        private bool disposedValue;
        Rooting root = new Rooting();
        public int MonthValue;
        public string MonthString;
        Start start = new Start();
        string enviroment = Start.enviroment;
        ConsoleFormat console = new ConsoleFormat();

        public string DetailManagement { set; get; }

        /// <summary>
        /// Método para crear un estadística.
        /// </summary>
        /// <param name="TimeToEndProcess"></param>
        /// <param name="Detail">Para mandar detalles a las estadisticas llenar primero la variable root.requestDetails</param>
        /// <param name="User">Para mandar el usuario solicitnate a las estadisticas llenar primero la variable root.BDUserCreatedBy</param>
        public void CreateStat()
        {
            console.WriteLine("Creando estadísticas...");

            string sql_insert;
            CRUD crud = new CRUD();
            #region Registro Detallado
            string sDate = root.BDStartDate.ToString("yyyy-MM-dd hh:mm:ss");
            string eDate = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

            string Detail = root.requestDetails;

            string user = root.BDUserCreatedBy;
            user = (user == "") ? "Databot" : user;
            
            sql_insert = $@"INSERT INTO `botdetails`
                            (`class`, `comments`, `startDate`, `endDate`, `createdBy`)
                            VALUES 
                            ('{root.BDIdClass}', '{Detail}', '{sDate}', '{eDate}', '{user}')";

            crud.Insert(sql_insert, "databot_db", enviroment);



            #endregion

            #region  Insertar en el .151 de S&S para pruebas de dashboard databot

    //        string sqlTest = "INSERT INTO `ConfirmContacts` (`idCustomer`, `idContact`, `CreateBy`) " +
    //$"VALUES ('{cliente}','{contacto_id}', '{usuario}')";

            crud.Insert(sql_insert, "databot_db", "QAS");
            #endregion 

            

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

    }
}
