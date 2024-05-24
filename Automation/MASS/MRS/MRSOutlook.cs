using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using DataBotV5.Data.Database;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.App.Global;
using DataBotV5.Data.Projects.MRSOutlook;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.MASS.MRS
{
    /// <summary>
    /// Clase MASS Automation que contiene los métodos para agendar las Meeting de MRS, por demandas o Time Allocations.
    /// </summary>
    class MRSOutlook
    {
        /// <summary>
        /// Clase main del MRS Outlook
        /// </summary>
        /// 
        ConsoleFormat console = new ConsoleFormat();
        CRUD crud = new CRUD();
        MRSOutlookSQL mrs = new MRSOutlookSQL();

        public void Main()
        {
            //Selecciona el mandante donde va a agendar las Meeting, DEV, QAS, PRD
            int mandante = Selector("QAS");

            //Busca las demandas o time allocations en la tabla mrs_outlook_sap asociadas al mandante establecido
            //que tengan el estatus EN PROGRESO, y las ordena para que puedan ser consumidos por el otro metodo
            //en la tabla outlook_meetings
            CreateTableTasks(mandante);
            //Time Allocation
            //Busca, crea, elimina o actualiza las Meetings de las Demandas o Time Allocations de la tabla
            //outlook_meetings, y realiza los cambios en el Outlook del ingeniero
            ScheduleMeetings(mandante);
            using (Stats stats = new Stats())
            {
                stats.CreateStat();
            }
        }
        public void EditAllocs()
        {
            DataTable data = new DataTable();
            string sql = "SELECT * FROM outlook_meetings WHERE MAND = '300' AND MEET_TYPE = 'Time Allocation'";
            //data = crud.Select("Databot", sql, "automation");

            for (int i = 0; i < data.Rows.Count; i++)
            {
                string id = data.Rows[i][0].ToString();
                MRS_TDS info = JsonConvert.DeserializeObject<MRS_TDS>(data.Rows[i][9].ToString());

                if (info.ALLOCATION == "VACACIONES")
                {
                    info.ALLOCATION = "VACACIO";
                    CRUD update = new CRUD();
                    //update.Update("Databot", $"UPDATE outlook_meetings SET INFO = '{JsonConvert.SerializeObject(info)}' WHERE ID = '{id}'", "automation","PRD");
                    console.WriteLine($"El ID: {id} ha sido actualizado");
                }
            }
        }

        private void InsertMeeting(MRS_TDS info)
        {
            string subject = "";
            string body = "";
            string format_time = "";
            List<ConfigColor> alloc = mrs.Allocations();
            if (info.TIME_START == null)
            {
                info.TIME_START = "00:00:00";
            }
            OutlookMeetings omeets = new OutlookMeetings();
            if (info.TDS == "Demand Asign")
            {
                subject = "Asignación: " + info.DESCRIPTION + ", Orden: " + info.DEMAND_ID + ", Item: " + info.DEMAND_ITEM;
                body = "Estimado(a) " + info.USER + ", ha sido asignado a atender la orden " + info.DEMAND_ID + " con el item " + info.DEMAND_ITEM + ", en MRS por el Planner: " + info.SAP_USER + "\r\n" +
                    "la descripción es: " + info.DESCRIPTION + ", comenzando en la fecha " + info.DATE_START + " con la hora " + info.TIME_START + " y finalizando en la fecha " + info.DATE_END + " con la hora " + info.TIME_END + ".\r\n" +
                    "El tiempo agendado es de " + info.TOTAL_HOURS + " horas, tenga en cuenta que esta reunión es informativa y puede sufrir cambios acorde lo planificado en MRS.";
                if (info.TIME_START == null)
                {
                    info.TIME_START = "00:00:00";
                }
                format_time = info.DATE_START + " " + info.TIME_START;
                DateTime meetingStart = DateTime.ParseExact(format_time, "yyyy-MM-dd HH:mm:ss", null);
                string sduracion = info.TOTAL_HOURS;
                double Dduracion = double.Parse(sduracion, CultureInfo.InvariantCulture);
                TimeSpan duration = TimeSpan.FromHours(Dduracion);
                int minutos = Convert.ToInt32(duration.TotalMinutes);
                omeets.CreateMeeting(subject, body, meetingStart, minutos, info.EMAIL); //info.EMAIL
                string insert_query = "INSERT INTO outlook_meetings (MAND,MRS_GUID,EMPLOYEE,EMAIL,MEET_TYPE,DESCRIPTION,WF_ACTION,INFO,ESTATUS,PLANNER,TS, OUTLOOK_UID) " +
                                "VALUES('" + info.MAND + "','" + info.GUID + "','" + info.USER + "','" + info.EMAIL + "','" + info.TDS + "','" + info.DESCRIPTION + "','" + info.WF_ACTION + "','" + JsonConvert.SerializeObject(info) + "'" +
                                ", 'COMPLETADO','" + info.SAP_USER + "',CURRENT_TIMESTAMP, '" + omeets.MeetingUID + "')";
                CRUD insert = new CRUD();
                //insert.Insert("Databot", insert_query, "automation");
                console.WriteLine(" Meeting with description: " + subject + ", and event: " + info.TDS + " created in Outlook.");           
            }
            else
            {
                if (info.DESCRIPTION == null || info.DESCRIPTION == "")
                {
                    info.DESCRIPTION = "Sin descripcion";
                }

                if (info.ALLOCATION == "VACACIONES")
                {
                    info.ALLOCATION = "VACACIO";
                }

                subject = "Asignación: " + info.DESCRIPTION + ", Tarea: " + TASelector(info.ALLOCATION, alloc);
                body = "Estimado(a) " + info.USER + ", ha sido asignado a atender la tarea de " + TASelector(info.ALLOCATION, alloc) + ", en MRS por el Planner: " + info.SAP_USER + "\r\n" +
                    "la descripción es: " + info.DESCRIPTION + ", comenzando en la fecha " + info.DATE_START + " con la hora " + info.TIME_START + " y finalizando en la fecha " + info.DATE_END + " con la hora " + info.TIME_END + ".\r\n" +
                    "El tiempo agendado es de " + info.TOTAL_HOURS + " horas, tenga en cuenta que esta reunión es informativa y puede sufrir cambios acorde lo planificado en MRS.";

                if (info.TIME_START == null)
                {
                    info.TIME_START = "00:00:00";
                }
                format_time = info.DATE_START + " " + info.TIME_START;
                DateTime meetingStart = DateTime.ParseExact(format_time, "yyyy-MM-dd HH:mm:ss", null);
                string sduracion = info.TOTAL_HOURS;
                double Dduracion = double.Parse(sduracion, CultureInfo.InvariantCulture);
                TimeSpan duration = TimeSpan.FromHours(Dduracion);
                int minutos = Convert.ToInt32(duration.TotalMinutes);
                omeets.CreateMeeting(subject, body, meetingStart, minutos, info.EMAIL); //info.EMAIL
                string insert_query = "INSERT INTO outlook_meetings (MAND,MRS_GUID,EMPLOYEE,EMAIL,MEET_TYPE,DESCRIPTION,WF_ACTION,INFO,ESTATUS,PLANNER,TS, OUTLOOK_UID) " +
                                "VALUES('" + info.MAND + "','" + info.GUID + "','" + info.USER + "','" + info.EMAIL + "','" + info.TDS + "','" + info.DESCRIPTION + "','" + info.WF_ACTION + "','" + JsonConvert.SerializeObject(info) + "'" +
                                ", 'COMPLETADO','" + info.SAP_USER + "',CURRENT_TIMESTAMP, '" + omeets.MeetingUID + "')";
                CRUD insert = new CRUD();
                //insert.Insert("Databot", insert_query, "automation");
                console.WriteLine(" Meeting with description: " + subject + ", and event: " + info.TDS + " created in Outlook.");
            }
        }
        /// <summary>
        /// Método que crea, actualiza o elimina las Meetings en Outlook.
        /// </summary>
        /// <param name="mandante">Mandante de SAP</param>
        private void ScheduleMeetings(int mandante)
        {
            DataTable meetings = mrs.DecompressMeetings(mandante);
            List<ConfigColor> alloc = mrs.Allocations();

            string subject = "";
            string body = "";
            string format_time = "";
            if (meetings != null)
            {
                if (meetings.Rows.Count > 0)
                {
                    for (int i = 0; i < meetings.Rows.Count; i++)
                    {
                        OutlookMeetings omeets = new OutlookMeetings();
                        string guid_meeting = meetings.Rows[i][3].ToString();
                        MRS_OUT mrs_data = mrs.DecompressOutlook(guid_meeting);
                        MRS_TDS info = JsonConvert.DeserializeObject<MRS_TDS>(mrs_data.INFO);

                        switch (mrs_data.WF_ACTION)
                        {
                            case "UPDATE":
                                if (info.TDS == "Demand?Asign" || info.TDS == "Demand Asign")
                                {
                                    subject = "Asignación: " + info.DESCRIPTION + ", Orden: " + info.DEMAND_ID + ", Item: " + info.DEMAND_ITEM;
                                    body = "Estimado(a) " + info.USER + ", ha sido asignado a atender la orden " + info.DEMAND_ID + " con el item " + info.DEMAND_ITEM + ", en MRS por el Planner: " + info.SAP_USER + "\r\n" +
                                        "la descripción es: " + info.DESCRIPTION + ", comenzando en la fecha " + info.DATE_START + " con la hora " + info.TIME_START + " y finalizando en la fecha " + info.DATE_END + " con la hora " + info.TIME_END + ".\r\n" +
                                        "El tiempo agendado es de " + info.TOTAL_HOURS + " horas, tenga en cuenta que esta reunión es informativa y puede sufrir cambios acorde lo planificado en MRS.";

                                    if (info.TIME_START == null)
                                    {
                                        info.TIME_START = "00:00:00";
                                    }
                                    format_time = info.DATE_START + " " + info.TIME_START;
                                    DateTime meetingStart = DateTime.ParseExact(format_time, "yyyy-MM-dd HH:mm:ss", null);
                                    string sduracion = info.TOTAL_HOURS;
                                    double Dduracion = double.Parse(sduracion, CultureInfo.InvariantCulture);
                                    TimeSpan duration = TimeSpan.FromHours(Dduracion);
                                    int minutos = Convert.ToInt32(duration.TotalMinutes);
                                    omeets.UpdateMeeting(mrs_data.OUTLOOK_IUD, subject, body, minutos, info.EMAIL ,meetingStart);
                                    console.WriteLine(" Meeting with description: " + subject + ", and event: " + info.TDS + " updated in Outlook.");
                                    string update_query = "UPDATE outlook_meetings SET ESTATUS = 'COMPLETADO', TS = CURRENT_TIMESTAMP WHERE ID = '" + mrs_data.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                                else if (info.TDS == "Time?Allocation" || info.TDS == "Time Allocation")
                                {
                                    if (info.DESCRIPTION == null || info.DESCRIPTION == "")
                                    {
                                        info.DESCRIPTION = "Sin descripcion";
                                    }

                                    if (info.ALLOCATION == "VACACIONES")
                                    {
                                        info.ALLOCATION = "VACACIO";
                                    }

                                    subject = "Asignación: " + info.DESCRIPTION + ", Tarea: " + TASelector(info.ALLOCATION, alloc);
                                    body = "Estimado(a) " + info.USER + ", ha sido asignado a atender la tarea de " + TASelector(info.ALLOCATION, alloc) + ", en MRS por el Planner: " + info.SAP_USER + "\r\n" +
                                        "la descripción es: " + info.DESCRIPTION + ", comenzando en la fecha " + info.DATE_START + " con la hora " + info.TIME_START + " y finalizando en la fecha " + info.DATE_END + " con la hora " + info.TIME_END + ".\r\n" +
                                        "El tiempo agendado es de " + info.TOTAL_HOURS + " horas, tenga en cuenta que esta reunión es informativa y puede sufrir cambios acorde lo planificado en MRS.";
                                    if (info.TIME_START == null)
                                    {
                                        info.TIME_START = "00:00:00";
                                    }
                                    format_time = info.DATE_START + " " + info.TIME_START;
                                    DateTime meetingStart = DateTime.ParseExact(format_time, "yyyy-MM-dd HH:mm:ss", null);
                                    string sduracion = info.TOTAL_HOURS;
                                    double Dduracion = double.Parse(sduracion, CultureInfo.InvariantCulture);
                                    TimeSpan duration = TimeSpan.FromHours(Dduracion);
                                    int minutos = Convert.ToInt32(duration.TotalMinutes);
                                    omeets.UpdateMeeting(mrs_data.OUTLOOK_IUD, subject, body, minutos, info.EMAIL, meetingStart);
                                    console.WriteLine(" Meeting with description: " + subject + ", and event: " + info.TDS + " updated in Outlook.");
                                    string update_query = "UPDATE outlook_meetings SET ESTATUS = 'COMPLETADO', TS = CURRENT_TIMESTAMP WHERE ID = '" + mrs_data.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                                break;
                            case "DELETE":
                                if (info.TDS == "Demand Asign" || info.TDS == "Demand?Asign")
                                {
                                    omeets.DeleteMeeting(mrs_data.OUTLOOK_IUD);
                                    console.WriteLine(" Meeting with description: " + subject + ", and event: " + info.TDS + " deleted in Outlook.");
                                    string update_query = "UPDATE outlook_meetings SET ESTATUS = 'COMPLETADO', TS = CURRENT_TIMESTAMP WHERE ID = '" + mrs_data.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                                else if (info.TDS == "Time Allocation" || info.TDS == "Time?Allocation")
                                {
                                    omeets.DeleteMeeting(mrs_data.OUTLOOK_IUD);
                                    console.WriteLine(" Meeting with description: " + subject + ", and event: " + info.TDS + " deleted in Outlook.");
                                    string update_query = "UPDATE outlook_meetings SET ESTATUS = 'COMPLETADO', TS = CURRENT_TIMESTAMP WHERE ID = '" + mrs_data.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                                break;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Método que ordena la información consumida directamente desde SAP.
        /// </summary>
        /// <param name="mandante">Mandante de SAP</param>
        private void CreateTableTasks(int mandante)
        {
            DataTable tasks = mrs.DecompressTask(mandante);
            if (tasks != null)
            {
                if (tasks.Rows.Count > 0)
                {
                    for (int i = 0; i < tasks.Rows.Count; i++)
                    {
                        string table_id = tasks.Rows[i][0].ToString();
                        string table_mand = tasks.Rows[i][1].ToString();
                        string table_info = tasks.Rows[i][2].ToString();
                        //table_info = table_info.Replace("\"", "\\u022");

                        string table_sap = tasks.Rows[i][4].ToString();


                        try
                        {
                            MRS_TDS info = JsonConvert.DeserializeObject<MRS_TDS>(table_info);

                            if (info.DESCRIPTION == null || info.DESCRIPTION == "")
                            {
                                info.DESCRIPTION = info.ALLOCATION + " - " + info.PERSON_NAME;
                            }

                            info.DESCRIPTION = info.DESCRIPTION.Replace("\"", "");
                            info.DESCRIPTION = info.DESCRIPTION.Replace(@"\", "");

                            //Correo solo para pruebas
                            //info.EMAIL = "AERAMIREZ@GBM.NET";  //Comentar cuando se terminen las pruebas

                            if (info.WF_ACTION == "INSERT")
                            {
                                InsertMeeting(info);
                            }

                            if (info.WF_ACTION == "UPDATE")
                            {
                                console.WriteLine(" New Task with description: " + info.DESCRIPTION + ", and event: " + info.TDS + " updated.");
                                MRS_OUT meeting = mrs.DecompressOutlookT(info.GUID);
                                if (meeting.EMPLOYEE != null)
                                {
                                    string update_query = "UPDATE outlook_meetings SET EMPLOYEE = '" + info.USER + "', EMAIL = '" + info.EMAIL + "'," +
                                        " MEET_TYPE = '" + info.TDS + "', DESCRIPTION = '" + info.DESCRIPTION + "', WF_ACTION = '" + info.WF_ACTION + "'," +
                                        " INFO = '" + table_info + "', ESTATUS = 'EN PROCESO', PLANNER = '" + info.SAP_USER + "', TS = CURRENT_TIMESTAMP WHERE ID = '" + meeting.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                                else
                                {
                                    InsertMeeting(info);
                                    string update_query = "UPDATE outlook_meetings SET EMPLOYEE = '" + info.USER + "', EMAIL = '" + info.EMAIL + "'," +
                                       " MEET_TYPE = '" + info.TDS + "', DESCRIPTION = '" + info.DESCRIPTION + "', WF_ACTION = '" + info.WF_ACTION + "'," +
                                       " INFO = '" + table_info + "', ESTATUS = 'COMPLETADO', PLANNER = '" + info.SAP_USER + "', TS = CURRENT_TIMESTAMP WHERE ID = '" + meeting.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                            }

                            if (info.WF_ACTION == "DELETE")
                            {
                                MRS_OUT meeting = mrs.DecompressOutlookT(info.GUID);
                                if (meeting.EMPLOYEE != null)
                                {
                                    console.WriteLine(" New Task with description: " + info.DESCRIPTION + ", and event: " + info.TDS + " deleted.");
                                    string update_query = "UPDATE outlook_meetings SET WF_ACTION = '" + info.WF_ACTION + "'," +
                                       " ESTATUS = 'EN PROCESO', PLANNER = '" + info.SAP_USER + "', TS = CURRENT_TIMESTAMP WHERE ID = '" + meeting.ID + "'";
                                    CRUD update = new CRUD();
                                    //update.Update("Databot", update_query, "automation");
                                }
                                else
                                {
                                    console.WriteLine(" New Task with description: " + info.DESCRIPTION + ", and event: " + info.TDS + " doesn't exist, skipping.");
                                }
                            }

                            string update_query1 = "UPDATE mrs_outlook_sap SET ESTATUS = 'COMPLETADO' WHERE ID = '" + table_id + "'";
                            CRUD update1 = new CRUD();
                            //update1.Update("Databot", update_query1, "automation");
                        }
                        catch (Exception)
                        {

                            string update_query1 = "UPDATE mrs_outlook_sap SET ESTATUS = 'COMPLETADO' WHERE ID = '" + table_id + "'";
                            CRUD update1 = new CRUD();
                            //update1.Update("Databot", update_query1, "automation");
                        }

                        
                    }
                }
            }
        }

        /// <summary>
        /// Método que convierte el key de Time Allocation a formato de texto.
        /// </summary>
        /// <param name="key">Time Allocation</param>
        /// <returns>Retorna el texto del Time Allocation</returns>
        private string TASelector(string key,List<ConfigColor> lista)
        {
            string allocation = "";
            int indx = lista.FindIndex(x => x.Tipo == key);
            if (indx != -1)
            {
                allocation = lista[indx].Descripcion;
            }
            else
            {
                allocation = "";
            }
            return allocation;
        }
        /// <summary>
        /// Método para la selección del mandante.
        /// </summary>
        /// <param name="mandante">Mandante de SAP en texto</param>
        /// <returns>Retorna el valor entero del mandante.</returns>
        private int Selector(string mandante)
        {
            int mand = 0;
            switch (mandante)
            {
                case "DEV":
                    mand = 120;
                    console.WriteLine(" Connecting to DB and decompressing MAND: " + mand + " Tasks.");
                    break;
                case "QAS":
                    mand = 260;
                    console.WriteLine(" Connecting to DB and decompressing MAND: " + mand + " Tasks.");
                    break;
                case "PRD":
                    mand = 300;
                    console.WriteLine(" Connecting to DB and decompressing MAND: " + mand + " Tasks.");
                    break;
            }
            return mand; 
        }

    }
    /// <summary>
    /// Clase para descomprimir el JSON de la estructura de SAP ZMRS_TDS.
    /// </summary>
    public class MRS_TDS
    {
        public string GUID { get; set; }
        public string PERSON { get; set; }
        public string USER { get; set; }
        public string EMAIL { get; set; }
        public string DESCRIPTION { get; set; }
        public string DATE_START { get; set; }
        public string DATE_END { get; set; }
        public string TIME_START { get; set; }
        public string TIME_END { get; set; }
        public string TOTAL_HOURS { get; set; }
        public string TDS { get; set; }
        public string ALLOCATION { get; set; }
        public string DEMAND_ID { get; set; }
        public string DEMAND_ITEM { get; set; }
        public string MAND { get; set; }
        public string SAP_USER { get; set; }
        public string WF_ACTION { get; set; }
        public string CD { get; set; }
        public string CUSTOMER { get; set; }
        public string CUSTOMER_NAME { get; set; }
        public string CONTRACT { get; set; }
        public string CONTRACT_NAME { get; set; }
        public string A_TASK_TYPE { get; set; }
        public string A_TASK_PRACTICE { get; set; }
        public string A_TASK_SKILL { get; set; }
        public string PERSON_NAME { get; set; }
        public string STATUS { get; set; }
    }
    /// <summary>
    /// Clase de la tabla outlook_meetings
    /// </summary>
    public class MRS_OUT
    {
        public string ID { get; set; }
        public string MAND { get; set; }
        public string OUTLOOK_IUD { get; set; }
        public string MRS_GUID { get; set; }
        public string EMPLOYEE { get; set; }
        public string EMAIL { get; set; }
        public string MEET_TYPE { get; set; }
        public string DESCRIPTION { get; set; }
        public string WF_ACTION { get; set; }
        public string INFO { get; set; }
        public string ESTATUS { get; set; }
        public string PLANNER { get; set; }
        public string TS { get; set; }
    }
    public class ConfigColor
    {
        public string Tipo { get; set; }
        public string Descripcion { get; set; }
        public string Color { get; set; }
    }
}
