using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.App.Global
    
{
    /// <summary>
    /// Clase que contiene métodos especiales para imprimir en consola y a la vez respaldar los registros en un log.
    /// </summary>
    class ConsoleFormat : IDisposable
    {

        private bool disposedValue;
        /// <summary>
        /// Método para imprimir en consola y a la vez respaldar la línea en un Log. El formato de texto es: Date >>>> + mensaje.
        /// </summary>
        /// <param name="msg"></param>
        public void WriteLine(string msg)
        {
            Console.WriteLine(DateTime.Now + " > > > " + $" {msg}");
            //Llena el Log del robot.
            using (Settings sett = new Settings())
            {
                sett.CreateLog(DateTime.Now + " > > > " + $" {msg}");
            };
        }
        /// <summary>
        /// Método para imprimir en consola pero con una línea divisoria el cual resalta el texto como un título, además
        /// que respalda la línea en un log. El formato es una linea divisoria, y bajo a ella:  Date >>>> + mensaje.
        /// </summary>
        /// <param name="msg"></param>
        public void WriteAnnounce(string msg)
        {
            String lineLarge = "______________________________________________________________________________________________________________________";
            Console.WriteLine(DateTime.Now + " > > > " + $" {msg}");
            Console.WriteLine(lineLarge);
            

            //Llena el Log del robot.
            using (Settings sett = new Settings())
            {
                sett.CreateLog(DateTime.Now + " > > > " + $" {msg}");
                sett.CreateLog(lineLarge);
            };
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                disposedValue = true;
            }
        }


        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
