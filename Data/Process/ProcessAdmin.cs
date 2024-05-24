using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.App.Global;

namespace DataBotV5.Data.Process
{
    /// <summary>
    /// Clase Data encargada de administración de procesos.
    /// </summary>
    class ProcessAdmin : IDisposable
    {
        Credentials.Credentials cred = new Credentials.Credentials();
        private bool disposedValue;
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();



        public void DeleteFiles(string route)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(route);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }
        }
        public void MoveFiles(string route, string sourceRute)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(route);
            foreach (FileInfo file in di.EnumerateFiles())
            {
                try
                {
                    file.MoveTo(sourceRute + file.Name);
                }
                catch (Exception)
                {
                    file.Delete();
                }
             
            }
        }
        /// <summary>
        /// este metodo sirve para eliminar los dobles espacios en el nombre de un archivo
        /// se utiliza en el metodo "Download_DM_AllFiles" debido a que a nivel de HTML no tiene espacios dobles pero en el archivo de windows si
        /// ejemplo: html--->"ejemplo de archivo.xlsx"  windows file--->"ejemplo de  archivo.xlsx" (notar los espacios despues del "de")
        /// input "ejemplo de  archivo.xlsx" - output "ejemplo de archivo.xlsx"
        /// </summary>
        public void RemoveSpacesFiles()
        {
            string file_name = "";
            DirectoryInfo d = new DirectoryInfo(root.FilesDownloadPath);
            FileInfo[] infos = d.GetFiles();
            foreach (FileInfo f in infos)
            {
                file_name = f.FullName;
                file_name = Regex.Replace(file_name, @"\s+", " ");
                File.Move(f.FullName, file_name);
            }
        }

        public bool CurrentProcess(string @class)
        {
            bool resp = false;
            string sql_select;
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB     
                //sql_select = "select * from procesos_activos2 where Metodo = '" + method + "'";
                sql_select = "select * from orchestrator where class = '" + @class + "'";
                mytable = crud.Select(sql_select, "databot_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    //resp = Convert.ToBoolean(Convert.ToInt16(mytable.Rows[0]["Activar"].ToString()));
                    resp = Convert.ToBoolean(Convert.ToInt16(mytable.Rows[0]["active"].ToString()));
                }

            }
            catch (Exception ex)
            { console.WriteLine(ex.Message); }

            return resp;
        }
        /// <summary>
        /// Desactivar el método cuando se activa en la columna "Activar" de orchestrator.
        /// </summary>
        /// <param name="clase"></param>
        public void TurnOffProcess(string clase)
        {
            string sql_update;
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB     
                //sql_update = "UPDATE `procesos_activos2` SET `Activar`= 0 WHERE `Metodo` = '" + method + "'";
                sql_update = "UPDATE `orchestrator` SET `active`= 0 WHERE `class` = '" + clase + "'";
                crud.Update(sql_update, "databot_db");
                #endregion
            }
            catch (Exception ex)
            { console.WriteLine(ex.Message); }

        }

        public bool PlanMinutes(string metodo, DateTime start, int minutes)
        {
            Rooting root = new Rooting();
            DateTime end = root.planiMinute[metodo].AddMinutes(minutes);
            if (start > end)
            {
                root.planiMinute[metodo] = DateTime.UtcNow;
                return true;
            }
            return false;
        }

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
}
