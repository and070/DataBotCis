using System;
using System.Diagnostics;
using System.Linq;

namespace DataBotV5.Logical.Files
{
    /// <summary>
    /// Clase Logical para destruiu un proceso. 
    /// </summary>
    class DestroyProcess : IDisposable
    {
        private bool disposedValue;

        public void KillProcess(string process, bool kill)
        {
            int cont = 0;
            do
            {
                Process[] procesoKill = Process.GetProcessesByName(process);
                //code block
                foreach (Process driver in procesoKill)
                {
                    try
                    {
                        if (kill)
                        {
                            driver.Kill();
                            driver.WaitForExit();
                            driver.Dispose();
                        }
                        else
                        {
                            bool s = driver.CloseMainWindow();
                            driver.Close();
                            driver.WaitForExit();
                            driver.Dispose();
                        }

                    }
                    catch (Exception ex)
                    { }

                }
                cont++;
                if (cont > 10)
                {
                    break;
                }

            } while (Process.GetProcessesByName(process).Count() > 0);

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
        // ~DestroyProcess()
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
