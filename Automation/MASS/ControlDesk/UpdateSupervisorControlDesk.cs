using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.Projects.ControlDesk;
using System.Data;
using System.Xml;
using System;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Process;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.MASS.ControlDesk

{
    /// <summary>
    /// Clase MASS Automation encargada de actualizar el supervisor del colaborador en Control Desk.
    /// </summary>
    class UpdateSupervisorControlDesk
    {
        ControlDeskInteraction cdi = new ControlDeskInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ProcessAdmin padmin = new ProcessAdmin();
        Credentials cred = new Credentials();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();


        string response = "";

        string respFinal = "";


        public void Main()
        {
            if (mail.GetAttachmentEmail("Solicitudes CD Supervisor", "Procesados", "Procesados CD Supervisor"))
            {
                console.WriteLine("Procesando...");

                DataTable excel = ParseExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);

                if (excel != null)
                {
                    ProcessSupervisor("QAS", 460, excel);

                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }

            }
        }
        public void ProcessSupervisor(string mandCd, int mandCrm, DataTable excel)
        {
            cred.SelectCdMand(mandCd);

            foreach (DataRow fila in excel.Rows)
            {
                string usuario = fila["User(ID)"].ToString().ToUpper();
                string supervisor = fila["User(ID) Supervisor"].ToString().ToUpper();
                string status;

                usuario = usuario.Contains("@GBM.NET") ? usuario : usuario + "@GBM.NET";
                supervisor = supervisor.Contains("@GBM.NET") ? supervisor : supervisor + "@GBM.NET";

                try
                {
                    status = cdi.UpdateUserSupervisor(usuario, supervisor);
                }
                catch (Exception)
                {
                    status = "Error al conectarse a CD";
                }

                if (status == "OK")
                {
                    response = response + usuario.Replace("@GBM.NET", "") + " - Se asignó el supervisor: " + supervisor.Replace("@GBM.NET", "") + " al usuario<br>";
                   
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Asignación de supervisor: ", response, root.Subject);
                    respFinal = respFinal + "\\n" + "Asignación de supervisor:: " + response;

                }
                else if (status == "SAME")
                {
                    response = response + usuario.Replace("@GBM.NET", "") + " - El supervisor: " + supervisor.Replace("@GBM.NET", "") + " Ya esta asignado al usuario<br>";
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Asignación de supervisor: ", response, root.Subject);
                    respFinal = respFinal + "\\n" + "Asignación de supervisor:: " + response;
                }
                else
                {
                    response = response + usuario.Replace("@GBM.NET", "") + " - " + status + "<br>";
                }

            }

            console.WriteLine("Respondiendo solicitud");

            string[] cc = { "smarin@gbm.net" };
            mail.SendHTMLMail(response, new string[] { "internalcustomersrvs@gbm.net" }, "**" + mandCd + "** Actualizar Supervisor en Control Desk", cc, null);

            
            root.requestDetails = respFinal;
        }



        private DataTable ParseExcel(string route)
        {
            DataTable excel_dt = new DataTable();

            try
            {
                Excel.Application xla = new Excel.Application();
                //xla.Workbooks.Add(System.Reflection.Missing.Value);
                xla.Visible = false;
                xla.DisplayAlerts = false;
                xla.Workbooks.Open(route);

                xla.ActiveSheet.Cells.UnMerge();
                xla.Columns["F:F"].Select();
                xla.Selection.SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Select();
                xla.Selection.EntireRow.Delete();

                string[] cols = { "A", "B", "C", "D" };

                foreach (string col in cols)
                {
                    xla.Columns[col + ":" + col].Select();
                    xla.Selection.SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Select();
                    xla.Selection.FormulaR1C1 = "=+R[-1]C";
                }

                int last = xla.Cells[xla.Rows.Count, "J"].End(Excel.XlDirection.xlUp).Row;
                object[,] arr = xla.Range["A1:J" + last].Value2;

                xla.Workbooks.Close();
                xla.Quit();
                proc.KillProcess("EXCEL", true);



                //columnas
                for (int i = 1; i < arr.GetUpperBound(1); i++)
                {
                    excel_dt.Columns.Add(arr[1, i].ToString(), /*Type.GetType("System.Int32")*/ typeof(string));
                }

                //filas
                for (int i = 2; i < arr.GetUpperBound(0); i++)
                {
                    DataRow row = excel_dt.NewRow();
                    for (int j = 1; j < arr.GetUpperBound(1); j++)
                    {
                        if (arr[i, j] == null)
                        {
                            row[j - 1] = "";
                        }
                        else
                        {
                            row[j - 1] = arr[i, j].ToString();
                        }
                    }
                    //add las filas, en licitaciones lo hago :)
                    excel_dt.Rows.Add(row);
                }

            }
            catch (Exception)
            {
                excel_dt = null;
                mail.SendHTMLMail("Error al leer la plantilla", new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, null);
            }

            string validacion = excel_dt.Columns[0].ColumnName;

            if (validacion != "Compañia")
            {
                excel_dt = null;
                mail.SendHTMLMail("Utilizar el reporte de Cognos", new string[] { root.BDUserCreatedBy }, "**Supervisor CD** " + root.Subject, root.CopyCC);
            }

            return excel_dt;
        }

    }
}
