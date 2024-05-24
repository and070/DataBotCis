using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StackExchange.Redis;

namespace DataBotV5.Data.Database
{
    /// <summary>
    /// Clase Data encargada de 
    /// </summary>
    class Redis : IDisposable
    {
        private bool disposedValue;

        public ConnectionMultiplexer RedisConn(string ip, int port)
        {
            return ConnectionMultiplexer.Connect($"{ip}:{port},abortConnect=false,allowAdmin=true");
        }

        public List<string> GetAllKeys (string pattern, ConnectionMultiplexer conn, string ip, int port)
        {
            IDatabase dbr = conn.GetDatabase();
            var keys = conn.GetServer(ip, port).Keys();
            List<string> lista = keys.Select(x => (string)x).Where(y=>y.Contains(pattern)).ToList();
            return lista;
        }

        public string GetKey(ConnectionMultiplexer conn, string key)
        {
            IDatabase dbr = conn.GetDatabase();
            return dbr.StringGet(key);
        }

        public void SetKey(ConnectionMultiplexer conn, string key, string contents)
        {
            IDatabase dbr = conn.GetDatabase();
            dbr.StringSet(key, contents);
        }

        public void DelKey(ConnectionMultiplexer conn, string key)
        {
            IDatabase dbr = conn.GetDatabase();
            dbr.KeyDelete(key);
        }

        #region hide
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
        // ~Redis()
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
