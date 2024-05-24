using System;
using System.IO;
using SAP.Middleware.Connector;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.GbmHolidays;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;

namespace DataBotV5.Automation.RPA.CoEmployees
    {
    /// <summary>
    /// Clase RPA Automation encargada dell reporte de excepciones de horas masivas de otras áreas.
    /// </summary>
    class ExceptionsHoursReports 
    {

        public string[] dias;
        public int x = 0;
        public int longitud = 0;
        public string strc_respuesta = "";
        public int fila_test;
        int mandante = 260;
        Log log = new Log();
        public int narchivo_final = 0;
        ProcessInteraction kill = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        SapVariants sap = new SapVariants();
        Log logeo = new Log();
        string ruta_respuesta;
        string erpSystem = "ERP";
        public void Main()
        {
            Stats estadisticas = new Stats();
            Log logeo = new Log();
            ProcessInteraction proc = new ProcessInteraction();

            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Solicitudes Coex", "Procesados", "Procesados Coex"))
            {
                string extArchivo = Path.GetExtension(root.FilesDownloadPath + @"\" + root.ExcelFile);
               
                if (extArchivo == ".xlsx" || extArchivo == ".xls")
                {
                    //string rutaFinal = @"C:\Users\aeramirez\Desktop\ejemplo.xlsx";
                    string rutaFinal = root.FilesDownloadPath + @"\" + root.ExcelFile;
                    console.WriteLine("Procesando...");
                    RPACoE(rutaFinal);
                    console.WriteLine("Creando Estadisticas");
                    
                    if (strc_respuesta != "")
                    {
                        console.WriteLine("Creando Excel");
                        CreateExcel();
                        string[] cc = { "dmeza@gbm.net" };
                        string[] att = { ruta_respuesta };
                        string cuerpo_correo;
                       
                        cuerpo_correo = "Adjunto Excel con detalle del proceso realizado.";
                        mail.SendHTMLMail(cuerpo_correo, new string[] { root.BDUserCreatedBy }, "Resultado de gestion Coex", cc, att);
                        console.WriteLine("Enviando Excel");
                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                    }
                }
                else
                {
                    //No es un archivo de excel valido
                }
            }
            else
            {
                //No hay attachments
            }
            proc.KillProcess("EXCEL",true);
        }


        /// <summary>
        /// Método que obtiene los días en fecha XML entre las fechas 1 y 2, y lo adjunta a un array.
        /// </summary>
        /// <param name="Date1"></param>
        /// <param name="Date2"></param>
        public void GetAvailablesDays(string Date1, string Date2)
        {
            DateTime fechaInicial = DateTime.Parse(Date1);
            DateTime fechaFinal = DateTime.Parse(Date2);
            dias = new string[longitud];
            for (DateTime i = fechaInicial; i.Date <= fechaFinal.Date; i = i.AddDays(1))
            {

                    string date = i.ToString("yyyy-MM-dd");
                    dias[x] = date;
                    x = x + 1;
            }


        }

        private List<string> AvailableDays(string from, string to, Employees employee, List<HolidaysCalendar> calendar)
        {
            List<string> lista = new List<string>();

            DateTime di = DateTime.Parse(from);
            DateTime df = DateTime.Parse(to);

            for (DateTime i = di; i.Date <= df.Date; i = i.AddDays(1))
            {

                int indx = calendar.FindIndex(x => x.Country == employee.Country);
                int zindex = calendar[indx].Calendar.FindIndex(y => y.Date == i);
                if (zindex == -1)
                {
                    lista.Add(i.ToString("yyyy-MM-dd"));
                }
            }
            return lista;
        }
        /// <summary>
        /// Obtiene la cantidad de días hbiles entre fecha1 y fecha2, para ser usado como longitud dinámica del array.
        /// </summary>
        /// <param name="date1">Fecha inicial</param>
        /// <param name="date2">Fecha Final</param>
        public void GetLengthToArrayFromDates(string date1, string date2)
        {
            DateTime fechaInicial = DateTime.Parse(date1);
            DateTime fechaFinal = DateTime.Parse(date2);
            for (DateTime i = fechaInicial; i.Date <= fechaFinal.Date; i = i.AddDays(1))
            {
                if (i.DayOfWeek.ToString() != "Sunday" && i.DayOfWeek.ToString() != "Saturday")
                {
                    longitud += 1;
                }
            }
        }
        public void CoEReport(string date, string id, string order, string hours)
        {



            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["ID"] = id;
            parameters["DATE"] = date;
            parameters["ORDER"] = order;
            parameters["HOURS"] = System.Convert.ToDecimal(hours).ToString();

            IRfcFunction func = sap.ExecuteRFC(erpSystem, "ZCOE_REP", parameters);



            strc_respuesta = strc_respuesta + id + "\t" + date + "\t" + order + "\t" + hours + "\t"  + func.GetValue("RESPONSE").ToString() + "\r\n";
            console.WriteLine(id + "\t" + date + "\t" + order + "\t" + hours + "\t" + func.GetValue("RESPONSE").ToString() + "\r\n");
            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Reporte de Horas Coex", id + ": " + func.GetValue("RESPONSE").ToString(), root.Subject);
        }
        private List<HolidaysCalendar> CalendarioFeriados()
        {
            StaticJSON json = new StaticJSON();
            List<HolidaysCalendar> calendario = new List<HolidaysCalendar>();
            List<Holidays> diasFeriados = JsonConvert.DeserializeObject<List<Holidays>>(json.Holidays);
            if (diasFeriados != null)
            {
                if (diasFeriados.Count > 0)
                {
                    List<string> paises = diasFeriados.Select(x => x.Country).Distinct().ToList();
                    for (int i = 0; i < paises.Count; i++)
                    {
                        HolidaysCalendar hc = new HolidaysCalendar();
                        hc.Country = paises[i];
                        calendario.Add(hc);
                    }

                    for (int i = 0; i < diasFeriados.Count; i++)
                    {
                        int indx = calendario.FindIndex(x => x.Country == diasFeriados[i].Country);
                        if (indx != -1)
                        {
                            Days d = new Days();
                            if (calendario[indx].Calendar == null)
                            {
                                List<Days> listado = new List<Days>();
                                listado.Add(new Days() { Date = DateTime.Parse(diasFeriados[i].Date), Description = diasFeriados[i].Description });
                                calendario[indx].Calendar = listado;
                            }
                            else {
                                calendario[indx].Calendar.Add(new Days() { Date = DateTime.Parse(diasFeriados[i].Date), Description = diasFeriados[i].Description });
                            }
                        }
                        else
                        {
                            HolidaysCalendar hc = new HolidaysCalendar();
                            hc.Country = diasFeriados[i].Country;
                            Days d = new Days();
                            d.Date = DateTime.Parse(diasFeriados[i].Date);
                            d.Description = diasFeriados[i].Description;
                            hc.Calendar.Add(d);
                            calendario.Add(hc);
                        }
                    }
                }
            }
            return calendario;
        }
        private List<Employees> Employees(List<string> employees)
        {
            List<Employees> lista = new List<Employees>();
            RfcDestination destination = sap.GetDestRFC("ERP");
            RfcRepository repo = destination.Repository;
            IRfcFunction func = repo.CreateFunction("HRPAYUS_CLD_GET_IT0105");
            IRfcTable info_pa = func.GetTable("IT_P0105_KEY"); 
            for (int i = 0; i < employees.Count; i++)
            {
                info_pa.Append();
                info_pa.SetValue("PERNR", employees[i]);
                info_pa.SetValue("ENDDA", DateTime.Now.ToString("yyyy-MM-dd"));
                info_pa.SetValue("BEGDA", DateTime.Now.ToString("yyyy-MM-dd"));
                info_pa.SetValue("SUBTYPE", "0001");
            }
            func.Invoke(destination);
            IRfcTable info_re = func.GetTable("ET_0105");
            for (int i = 0; i < info_re.Count; i++)
            {
                var x = info_re.CurrentIndex = i;
                Employees emp = new Employees();
                emp.Employee = info_re.CurrentRow[0].GetValue().ToString();
                emp.Username = info_re.CurrentRow[24].GetValue().ToString();
                lista.Add(emp);
            }

            if (lista.Count > 0)
            {
                for (int i = 0; i < lista.Count; i++)
                {
                    func = repo.CreateFunction("ZFD_GET_USER_DETAILS");
                    func.SetValue("USUARIO", lista[i].Username);
                    func.Invoke(destination);
                    lista[i].Country = func.GetValue("PAIS").ToString();
                }
            }

            return lista;

        }
        
        public void CreateExcel()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlWorkBook = xlApp.Workbooks.Open(root.ReferenciaCoe);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];

            StreamReader archivotxt = new StreamReader(root.txtCoe);
            string linea;
            int i = 2;

            while ((linea = archivotxt.ReadLine()) != null)
            {
                string[] split = linea.Split(new[] { "\t" }, StringSplitOptions.None);
                xlWorkSheet.Cells[i, 1].value = split[0];
                xlWorkSheet.Cells[i, 2].value = split[1];
                xlWorkSheet.Cells[i, 3].value = split[2];
                xlWorkSheet.Cells[i, 4].value = split[3];
                xlWorkSheet.Cells[i, 5].value = split[4];
                i++;
            }

            archivotxt.Close();
            narchivo_final = new ValidateData().RandomNumber(1000000, 2000000);
            ruta_respuesta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\CoE_Files\" + narchivo_final.ToString() +".xlsx";

            xlWorkBook.SaveAs(ruta_respuesta);
            xlWorkBook.Close();
        }
        private void ProcessList(List<string> dates, string order ,CoExcel objeto)
        {
            for (int i = 0; i < dates.Count; i++)
            {
                CoEReport(dates[i], objeto.Employee, order, objeto.Time);
            }
        }
        public void ProcessArray(string id, string order, string hours)
        {

            for (int i = 0; i < dias.Length; i++)
            {
                if (dias[i] != null)
                {
                    CoEReport(dias[i].ToString(), id, order, hours);
                }
                else
                {
                    break;
                }
            }
            Array.Clear(dias, 0, dias.Length);
        }
        private string CollaboratorFormat(string id)
        {
            string resultado = "";

            switch (id.Length)
            {
                case 4:
                    resultado = "0000" + id;
                    break;
                case 5:
                    resultado = "000" + id;
                    break;
                case 6:
                    resultado = "00" + id;
                    break;
                case 7:
                    resultado = "0" + id;
                    break;
                default:
                    resultado = id;
                    break;
            }

            return resultado;
        }
        public void RPACoE(string route)
        {


            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            //Creamos el calendarios con los dias no laborales
            List<HolidaysCalendar> calendario = CalendarioFeriados();
            List<CoExcel> datos = new List<CoExcel>();

            int rows;

            rows = xlWorkSheet.UsedRange.Rows.Count;
            //Llenamos la lista de CoExcel con los datos del excel
            for (int i = 2; i <= rows; i++)
            {
                if (xlWorkSheet.Cells[i, 1].text != "")
                {
                    CoExcel coe = new CoExcel();
                    coe.Employee = xlWorkSheet.Cells[i, 1].text.ToString();
                    coe.From = xlWorkSheet.Cells[i, 2].text.ToString();
                    coe.To = xlWorkSheet.Cells[i, 3].text.ToString();
                    coe.Order = xlWorkSheet.Cells[i, 4].text.ToString();
                    coe.Time = xlWorkSheet.Cells[i, 5].text.ToString();
                    datos.Add(coe);
                }
            }
            //Cerramos el excel para mejorar el rendimiento, ahora solo trabajaremos desde memoria
            xlApp.Quit();
            xlApp = null;
            kill.KillProcess("EXCEL",true);

            List<string> total_empleados = datos.Select(x => x.Employee).Distinct().ToList();
            List<Employees> empleados = Employees(total_empleados);
            

            for (int i = 0; i < datos.Count; i++)
            {
                //Hora de procesar objeto por objeto
                datos[i].Employee = CollaboratorFormat(datos[i].Employee);
                int idx = empleados.FindIndex(x => x.Employee == datos[i].Employee);
                List<string> diasHabiles = AvailableDays(datos[i].From, datos[i].To,empleados[idx],calendario);
                string orden = datos[i].Order;
                if (orden.Length <= 11)
                {
                    orden = "0" + orden;
                }
                ProcessList(diasHabiles, orden, datos[i]);
            }

            File.WriteAllText(root.txtCoe, string.Empty);
            System.IO.File.WriteAllText(root.txtCoe, strc_respuesta);
        }
    }
   
}
//HRPAYUS_CLD_GET_IT0105
