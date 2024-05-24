using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Logical.Encode
{
    /// <summary>
    /// Clase Logical para codificar y decodificar strings a Base64.
    /// </summary>
    class String64 : IDisposable
    {
        private bool disposedValue;
        /// <summary>
        /// Decodifica un String Base64
        /// </summary>
        /// <param name="value">Valor del Bas64</param>
        /// <returns>Retorna un string decodificado</returns>
        public string Decode(string value)
        {
            byte[] decrypt = Convert.FromBase64String(value);
            string normalString = Encoding.ASCII.GetString(decrypt);
            return normalString;
        }
        /// <summary>
        /// Codificar string a Base64
        /// </summary>
        /// <param name="value">Valor del string</param>
        /// <returns>Devuelve un string Base64</returns>
        public string Encode(string value)
        {
            byte[] encrypt = System.Text.ASCIIEncoding.ASCII.GetBytes(value);
            string normalString = Convert.ToBase64String(encrypt);
            return normalString;
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
        // ~String64()
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
