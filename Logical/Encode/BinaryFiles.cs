using Ionic.Zip;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;

namespace DataBotV5.Logical.Encode
{
    /// <summary>
    /// Clase Logical encargada de manejo, control de archivos binarios, el método es disposable, utilice USING para un uso eficiente de la memoria.
    /// </summary>
    class BinaryFiles : IDisposable
    {
        /// <summary>
        /// Variable para el Dispose de la clase
        /// </summary>
        private bool disposedValue;
        /// <summary>
        /// Crear un archivo de tipo ZIP
        /// </summary>
        /// <param name="archivos">Arreglo de tipo ArchivoBinario con los nombres y contenidos de los archivos</param>
        /// <returns>Retorna un arhivo ZIP procesado en Bytes para almacenarlo en la BD</returns>
        public byte[] CrearBytesZIP(List<ArchivoBinario> archivos)
        {
            var outputStream = new MemoryStream();
            ZipFile zip = new ZipFile();
            archivos.ForEach(x =>
            {
                zip.AddEntry(x.NombreArchivo, x.Contenido);
            });
            zip.Save(outputStream);
            outputStream.Position = 0;
            byte[] bytesOut = outputStream.ToArray();
            return bytesOut;
        }
        /// <summary>
        /// Convierte un archivo a objeto ArchivoBinario
        /// </summary>
        /// <param name="route">Ruta del archivo</param>
        /// <param name="name">Nombre del archivo</param>
        /// <returns>Retorna objeto de tipo ArchivoBinario</returns>
        public ArchivoBinario Convert(string route, string name)
        {
            ArchivoBinario archivo = new ArchivoBinario {
                NombreArchivo = name,
                Contenido = File.ReadAllBytes(route)
            };
            return archivo;
        }
        /// <summary>
        /// Realiza el Insert o el Update de un SQL en la BD
        /// </summary>
        /// <param name="sql">SQL Query</param>
        /// <param name="conn">Coneccion a la DB</param>
        /// <param name="parameter">@param debe ir inicializado en el query, EJ: UPDATE tabla SET Archivo = @param WHERE...</param>
        /// <param name="file">Bytes del archivo</param>
        public void DataBaseInsertOrUpdate(string sql, MySqlConnection conn, string parameter, byte[] file)
        {
            MySqlCommand execute = new MySqlCommand(sql, conn);
            execute.Parameters.Add($"@{parameter}", MySqlDbType.LongBlob, file.Length).Value = file;
            conn.Open();
            execute.ExecuteNonQuery();
            conn.Close();
        }
        /// <summary>
        /// Retorna una imagen en formato HTML para su insercion el Body
        /// </summary>
        /// <param name="imagen">Imagen de tipo bytes</param>
        /// <returns>Retorna un IMG en formato HTML con la imagen en base 64</returns>
        public string BinaryHTMLImage(byte[] imagen)
        {
            string html = $"<img src=\"data:image/png;base64,{System.Convert.ToBase64String(imagen)}\">";
            return html;
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
        // ~BinaryFiles()
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
    /// <summary>
    /// Clase de Archivo Binario para inicializacion
    /// </summary>
    public class ArchivoBinario
    {
        /// <summary>
        /// Nombre del archivo junto a su extencion
        /// </summary>
        public string NombreArchivo { get; set; }
        /// <summary>
        /// Contenido en byte[] del archivo
        /// </summary>
        public byte[] Contenido { get; set; }
    }
}
