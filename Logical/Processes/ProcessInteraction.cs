using System;
using System.Linq;
using System.Diagnostics;

namespace DataBotV5.Logical.Processes
{
    /// <summary>
    /// Clase Logical encargada de interacción de procesos.
    /// </summary>
    class ProcessInteraction
    {

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
                            //driver.Kill();
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
        public string greeting()
        {
            string msj = "";
            TimeSpan morning = TimeSpan.Parse("05:00"); // 10 PM
            TimeSpan afternoon = TimeSpan.Parse("13:00");   // 2 AM
            TimeSpan night = TimeSpan.Parse("19:00");
            TimeSpan now = DateTime.Now.TimeOfDay;

            if (now >= morning && now <= afternoon)
            {
                msj = "Buenos días";
            }
            else if (now >= afternoon && now <= night)
            {
                msj = "Buenas tardes";
            }
            else if (now >= night)
            {
                msj = "Buenas noches";
            }
            else
            {
                msj = "Buenas";
            }
            return msj;
        }
    }
}
