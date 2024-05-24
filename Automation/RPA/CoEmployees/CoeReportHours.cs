using System;
using System.IO;
using SAP.Middleware.Connector;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.GbmHolidays;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.MicrosoftTools;
using System.Data;
using DataBotV5.Data.Database;
using System.Globalization;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net;

namespace DataBotV5.Automation.RPA.CoEmployees
{
    /// <summary>
    /// Clase RPA Automation encargada del reporte de horas masivo (CoE, Orden Interna).
    /// </summary>
    class CoeReportHours
    {
        public string[] dias;
        public int x = 0;
        public int longitud = 0;
        public string strc_respuesta = "";
        public int fila_test;
        Log log = new Log();
        public int narchivo_final = 0;
        DataBotV5.Logical.Processes.ProcessInteraction kill = new DataBotV5.Logical.Processes.ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        SapVariants sap = new SapVariants();
        string mandante = "ERP";
        string ssMandante = "QAS";
        MsExcel MsExcel = new MsExcel();
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Log logeo = new Log();
        Stats stats = new Stats();
        CRUD crud = new CRUD();
        string ruta_respuesta;
        string respFinal = "";
        public void Main()
        {
            Stats estadisticas = new Stats();
            Log logeo = new Log();
            ProcessInteraction proc = new ProcessInteraction();

            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Solicitudes Coe", "Procesados", "Procesados Coe"))
            {
                string extArchivo = Path.GetExtension(root.FilesDownloadPath + @"\" + root.ExcelFile);

                if (extArchivo == ".xlsx" || extArchivo == ".xls")
                {
                    string rutaFinal = root.FilesDownloadPath + @"\" + root.ExcelFile;
                    console.WriteLine("Procesando...");
                    RPACoE(rutaFinal);
                    console.WriteLine("Creando Estadisticas");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }

        public void RPACoE(string route)
        {

            string html = Properties.Resources.emailtemplate1;

            DataTable excel = MsExcel.GetExcel(route);

            #region excel verification and Molded of the Results Excel
            //columnas necesarias para cargar en SAP
            string[] readColumns = {
                "Aplica Fin de Semana"
            };
            //columnas del excel
            DataColumnCollection columns = excel.Columns;
            //contador para verificar el excel
            int contTrue = 0;

            console.WriteLine("Checking...");
            foreach (string columnName in readColumns)
            {
                //verifica si la columna esta en el excel
                if (columns.Contains(columnName))
                {
                    contTrue++;
                }
            }
            //si es diferente a 4 significa que no encontro una de las columnas necesarias para cargar la reconocimiento
            if (contTrue != readColumns.Length)
            {
                string[] cc = root.CopyCC;
                Array.Resize(ref cc, cc.Length + 1);
                cc[cc.Length - 1] = "dmeza@gbm.net";

                html = html.Replace("{subject}", "Reporte de Horas COE").Replace("{cuerpo}", "Error: por favor utilizar la nueva plantilla de reporte de horas COE, para más información comunicarse con Rainer Morales").Replace("{contenido}", "");
                mail.SendHTMLMail(html, new string[] { root.BDUserCreatedBy }, "Error: " + root.Subject, cc, new string[] { route });
                return;
            }
            #endregion

            DataTable excelResponse = excel.Clone();
            excelResponse.Columns.Add("Respuesta");

            //Creamos el calendarios con los dias no laborales
            List<CoeHolidaysCalendar> calendario = CalendarioFeriados();
            List<string> employees = excel.AsEnumerable().Select(x => x["Colaborador"].ToString()).Distinct().ToList();
            List<CoeEmployees> employeesList = CoeEmployee(employees);

            for (int i = 0; i < excel.Rows.Count; i++)
            {
                string orden = excel.Rows[i]["Orden"].ToString().PadLeft(12, '0');
                string employee = excel.Rows[i]["Colaborador"].ToString().PadLeft(8, '0');
                if (employee == "00000000")
                {
                    continue;
                }
                string aplica = excel.Rows[i]["Aplica Fin de Semana"].ToString();
                string from = excel.Rows[i]["De"].ToString();
                string to = excel.Rows[i]["Hasta"].ToString();
                string hours = excel.Rows[i]["Horas"].ToString();

                try
                {
                    //Hora de procesar objeto por objeto
                    int idx = employeesList.FindIndex(x => x.Employee == employee);
                    if (idx == -1)
                    {
                        DataRow responseRow = excelResponse.Rows.Add();
                        responseRow["Colaborador"] = employee;
                        responseRow["Orden"] = orden;
                        responseRow["De"] = from;
                        responseRow["Hasta"] = to;
                        responseRow["Aplica Fin de Semana"] = aplica;
                        responseRow["Horas"] = hours;
                        responseRow["Respuesta"] = "Empleado no cuenta con país o datos en SAP";
                        continue;
                    }
                    List<string> diasHabiles = AvailableDays(from, to, employeesList[idx], calendario, aplica);
                    for (int e = 0; e < diasHabiles.Count; e++)
                    {
                        try
                        {

                            DataRow responseRow = excelResponse.Rows.Add();
                            responseRow["Colaborador"] = employee;
                            responseRow["Orden"] = orden;
                            responseRow["De"] = diasHabiles[e];
                            responseRow["Hasta"] = diasHabiles[e];
                            responseRow["Aplica Fin de Semana"] = aplica;
                            responseRow["Horas"] = hours;
                            string resp = CoEReport(diasHabiles[e], employee, orden, hours);
                            responseRow["Respuesta"] = resp;
                            logeo.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear reporte de horas COE", resp, root.Subject);

                        }
                        catch (Exception ex)
                        {
                            strc_respuesta = strc_respuesta + employee + "\t" + from + "\t" + orden + "\t" + to + "\t" + ex.Message + "\r\n";
                            DataRow responseRow = excelResponse.Rows.Add();
                            responseRow["Colaborador"] = employee;
                            responseRow["Orden"] = orden;
                            responseRow["De"] = diasHabiles[e];
                            responseRow["Hasta"] = diasHabiles[e];
                            responseRow["Aplica Fin de Semana"] = aplica;
                            responseRow["Horas"] = hours;
                            responseRow["Respuesta"] = ex.Message;
                            logeo.LogDeCambios("Creacion", root.BDProcess,  root.BDUserCreatedBy, "Crear reporte de horas COE", ex.Message, root.Subject);
                        }
                    }

                }
                catch (Exception exs)
                {
                    DataRow responseRow = excelResponse.Rows.Add();
                    responseRow["Colaborador"] = employee;
                    responseRow["Orden"] = orden;
                    responseRow["De"] = from;
                    responseRow["Hasta"] = to;
                    responseRow["Aplica Fin de Semana"] = aplica;
                    responseRow["Horas"] = hours;
                    responseRow["Respuesta"] = exs.Message;
                }
            }
            string ruta = root.FilesDownloadPath + "\\" + "Response_" + root.ExcelFile;
            excelResponse.AcceptChanges();
            MsExcel.CreateExcel(excelResponse, "ResponseCoe", ruta, true);

            html = html.Replace("{subject}", "Reporte de Horas COE");
            html = html.Replace("{cuerpo}", "Adjunto Excel con detalle del proceso realizado.");
            html = html.Replace("{contenido}", "");

            mail.SendHTMLMail(html, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC, new string[] { ruta });
            root.requestDetails = respFinal;

        }



        private List<string> AvailableDays(string from, string to, CoeEmployees employee, List<CoeHolidaysCalendar> calendar, string findeAplica)
        {
            List<string> lista = new List<string>();

            DateTime di = DateTime.Parse(from);
            DateTime df = DateTime.Parse(to);
            TimeSpan diffDays = df - di;
            int daysCant = diffDays.Days;

            for (DateTime i = di; i.Date <= df.Date; i = i.AddDays(1))
            {

                int indx = calendar.FindIndex(x => x.Country == employee.Country);
                int zindex = calendar[indx].Calendar.FindIndex(y => y.Date == i);
                if (findeAplica == "" && daysCant <= 2)
                {
                    //Nota si dejan en blanco la columna y es menos de un mes de carga el Bot no considera la línea
                    continue;
                }

                //hu4 Si indica que si y son menos de 7 días (o sea aplica Finde semana)
                if (findeAplica == "SI" && daysCant <= 7 && i.DayOfWeek.ToString() == "Sunday" && zindex == -1 || findeAplica == "SI" && daysCant <= 7 && i.DayOfWeek.ToString() == "Saturday" && zindex == -1)
                {
                    lista.Add(i.ToString("yyyy-MM-dd"));
                }
                //hu3 Para los rangos de fecha que son de la semana completa (Lunes a Domingo),  el bot solo debe cargar las horas en el periodo correspondiente a  la semana laboral (Lunes a Viernes)  WORKDAY aun cuando en la columna D de la plantilla tenga la indicación de “SI” o columna en blanco.
                else if (daysCant <= 31)
                {
                    if (i.DayOfWeek.ToString() != "Sunday" && i.DayOfWeek.ToString() != "Saturday" && zindex == -1)
                    {
                        lista.Add(i.ToString("yyyy-MM-dd"));
                    }
                }


            }
            return lista;
        }

        public string CoEReport(string date, string id, string order, string hours)
        {

            Dictionary<string, string> parametros = new Dictionary<string, string>();
            parametros["ID"] = id;
            parametros["DATE"] = date;
            parametros["ORDER"] = order;
            parametros["HOURS"] = System.Convert.ToDecimal(hours).ToString();

            IRfcFunction func = sap.ExecuteRFC(mandante, "ZCOE_REP", parametros);
            Console.WriteLine(id + "\t" + date + "\t" + order + "\t" + hours + "\t" + func.GetValue("RESPONSE").ToString() + "\r\n");

            return func.GetValue("RESPONSE").ToString();
            //cliente.Dispose();
            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Reporte de Horas CoE", id + ": " + func.GetValue("RESPONSE").ToString(), root.Subject);
            respFinal = respFinal + "\\n" + id + ": " + func.GetValue("RESPONSE").ToString();

        }
        private List<CoeHolidaysCalendar> CalendarioFeriados()
        {
            StaticJSON json = new StaticJSON();
            int year = DateTime.Now.Year;
            DateTime november4th = new DateTime(year, 11, 4);
            DateTime december20th = new DateTime(year, 12, 20);
            // List of Central American country codes
            string[] centralAmericanCountries = { "CR", "SV", "GT", "HN", "NI", "PA", "CO", "DO", "US" };

            List<CoeHolidaysCalendar> calendario = new List<CoeHolidaysCalendar>();

            foreach (var countryCode in centralAmericanCountries)
            {
                var holidays = GetPublicHolidays(year, countryCode);

                List<CoeDays> holidayList = new List<CoeDays>();

                foreach (var holiday in holidays)
                {

                    if (countryCode == "PA" && holiday.Date == november4th)
                    {
                        continue;
                    }
                    if (countryCode == "PA" && holiday.Date == december20th)
                    {
                        continue;
                    }
                    holidayList.Add(new CoeDays
                    {
                        Description = holiday.LocalName,
                        Date = holiday.Date
                    });
                }

                if (countryCode == "PA")
                {
                    holidayList.Add(new CoeDays
                    {
                        Description = "Dia del Duelo Nacional",
                        Date = december20th
                    });
                }

                calendario.Add(new CoeHolidaysCalendar
                {
                    Country = countryCode == "DO" ? "DR" : countryCode == "US" ? "MD" : countryCode,
                    Calendar = holidayList
                });
            }

            return calendario;
        }
        private List<CoeEmployees> CoeEmployee(List<string> employees)
        {
            List<CoeEmployees> lista = new List<CoeEmployees>();

            foreach (var item in employees)
            {
                string select = $"SELECT user, country, UserID FROM digital_sign WHERE UserID = {item} ";
                DataTable SSEmploye = crud.Select(select, "MIS");
                if (SSEmploye.Rows.Count > 0)
                {
                    foreach (DataRow employee in SSEmploye.Rows)
                    {
                        CoeEmployees emp = new CoeEmployees();
                        emp.Employee = employee["UserID"].ToString().PadLeft(8, '0');
                        emp.Username = employee["user"].ToString();
                        emp.Country = employee["country"].ToString();
                        lista.Add(emp);
                    }
                }
                else
                {
                    RfcDestination destination = sap.GetDestRFC(mandante);
                    RfcRepository repo = destination.Repository;

                    IRfcFunction func = repo.CreateFunction("HRPAYUS_CLD_GET_IT0105");

                    IRfcTable info_pa = func.GetTable("IT_P0105_KEY");
                    info_pa.Append();
                    info_pa.SetValue("PERNR", item);
                    info_pa.SetValue("ENDDA", "9999-12-31");
                    //info_pa.SetValue("BEGDA", "0001-01-01"); //DateTime.Now.ToString("yyyy-MM-dd")
                    info_pa.SetValue("SUBTYPE", "0001");
                    func.Invoke(destination);
                    IRfcTable info_re = func.GetTable("ET_0105");
                    for (int i = 0; i < info_re.Count; i++)
                    {
                        var x = info_re.CurrentIndex = i;
                        CoeEmployees emp = new CoeEmployees();
                        emp.Employee = info_re.CurrentRow[0].GetValue().ToString().PadLeft(8, '0');
                        emp.Username = info_re.CurrentRow[24].GetValue().ToString();
                        func = repo.CreateFunction("ZFD_GET_USER_DETAILS");
                        func.SetValue("USUARIO", emp.Username);
                        func.Invoke(destination);
                        emp.Country = func.GetValue("PAIS").ToString();
                        if (!string.IsNullOrWhiteSpace(emp.Country))
                        {

                            lista.Add(emp);
                        }
                    }

                }

            }


            return lista;

        }

        static List<HolidayInfo> GetPublicHolidays(int year, string countryCode)
        {
            List<HolidayInfo> holidays = new List<HolidayInfo>();

            try
            {
                // Make the GET request
                string apiUrl = $"https://date.nager.at/api/v3/PublicHolidays/{year}/{countryCode}";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(apiUrl);
                request.Method = "GET";

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream))
                {
                    string json = reader.ReadToEnd();
                    holidays = ParseHolidayResponse(json);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            return holidays;
        }

        static List<HolidayInfo> ParseHolidayResponse(string json)
        {
            // Parse the JSON response into a list of HolidayInfo objects
            var holidays = new List<HolidayInfo>();

            try
            {
                JArray jsonArray = JArray.Parse(json);
                foreach (var item in jsonArray)
                {
                    var holiday = new HolidayInfo
                    {
                        Date = DateTime.Parse(item["date"].ToString()),
                        LocalName = item["localName"].ToString(),
                        Name = item["name"].ToString()
                    };

                    holidays.Add(holiday);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing JSON: {ex.Message}");
            }

            return holidays;
        }

    }


    public class CoeReport
    {
        public string Employee { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Order { get; set; }
        public string Time { get; set; }
    }
    class CoeDays
    {
        public string Description { get; set; }
        public DateTime Date { get; set; }
    }
    public class CoeEmployees
    {
        public string Employee { get; set; }
        public string Username { get; set; }
        public string Country { get; set; }
    }

    class CoeHolidays
    {
        public string Country { get; set; }
        public string Date { get; set; }
        public string Description { get; set; }
    }
    class CoeHolidaysCalendar
    {
        public string Country { get; set; }
        public List<CoeDays> Calendar { get; set; }
    }

    class HolidayInfo
    {
        public DateTime Date { get; set; }
        public string LocalName { get; set; }
        public string Name { get; set; }
    }
}
