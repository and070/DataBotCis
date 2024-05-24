using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Webex;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Automation.WEB.CheckPlanner
{
    public class AlertBotPlanner
    {
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();
        Log log = new Log();
        Rooting root = new Rooting();
        string respFinal = "";
        bool executeStats = false;


        public void Main()
        {
            DateTime datenow = DateTime.Now; //01.01.1950 23:53:00
            DateTime Sdate = new DateTime(datenow.Year, datenow.Month, datenow.Day, 23, 50, 00); //01.01.1950 23:50:00
            DateTime Edate = Sdate.AddMinutes(5); //01.01.1950 23:55:00

            if (Edate > datenow && datenow > Sdate) // Verifica si está en rango de hora establecida y el waiting.
            {
                console.WriteLine("Procesando...");
                DataTable dt = crud.Select("SELECT planner, class FROM `orchestrator` WHERE JSON_LENGTH(planner) > 0", "databot_db");
                console.WriteLine($"Se encontraron {dt.Rows.Count} bots con planificador");
                foreach (DataRow bot in dt.Rows)
                {
                    console.WriteLine($"Procesando el bot {bot["class"]}");
                    Newtonsoft.Json.Linq.JObject planner = Newtonsoft.Json.Linq.JObject.Parse(bot["planner"].ToString());
                    string type = planner["Type"].ToString();
                    string[] week = planner["Week"].ToString().Split(',');
                    string[] month = planner["Month"].ToString().Split(',');
                    bool execute = false;
                    if (type == "0")
                    {


                    }
                    else if (type == "1")
                    {
                        //el de horas
                        execute = true;
                    }
                    else if (type == "2")
                    {
                        //el de dia

                        for (int i = 0; i < week.Length; i++)
                        {
                            if ((int)DateTime.Now.DayOfWeek == int.Parse(week[i].Trim()))
                            {
                                execute = true;
                            }

                        }

                    }
                    else if (type == "3")
                    {
                        //el de mes

                        for (int i = 0; i < month.Length; i++)
                        {
                            if (DateTime.Now.Day == int.Parse(month[i].Trim()))
                            {
                                execute = true;
                            }

                        }
                    }
                    else
                    {
                        //tipo desconocido

                    }


                    if (execute)
                    {
                        string[] horas = planner["Hour"].ToString().Split(',');
                        int cantHours = horas.Length;
                        console.WriteLine($"Verificando si el bot {bot["class"]} se ejecutó en el día la cantidad de {cantHours} veces");
                        DataTable dtDetails = crud.Select($"SELECT * FROM botdetails WHERE class = {bot["id"]} AND createdAt >= '{DateTime.Now.ToString("yyyy-MM-dd")}'", "databot_db");
                        if (cantHours != dtDetails.Rows.Count)
                        {

                            string resp = $"El proceso del robot: {bot["class"]} se debe de ejecutar {cantHours} veces, sin embargo solo se ejecutó una cantidad de {dtDetails.Rows.Count} veces";
                            console.WriteLine(resp);
                            using (WebexTeams wx = new WebexTeams())
                            {
                                wx.SendNotification("dmeza@gbm.net", "Robot no se ejecutó", $"El proceso del robot: {bot["class"]} se debe de ejecutar {cantHours} veces, sin embargo solo se ejecutó una cantidad de {dtDetails.Rows.Count} veces");
                                wx.SendNotification("epiedra@gbm.net", "Robot no se ejecutó", $"El proceso del robot: {bot["class"]} se debe de ejecutar {cantHours} veces, sin embargo solo se ejecutó una cantidad de {dtDetails.Rows.Count} veces");
                            }

                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Reporte de no ejecución de bot", resp, root.Subject);
                            respFinal = respFinal + "\\n" + "Reporte de no ejecución de bot: " + resp;
                            executeStats = true;


                        }
                    }
                }
            }

            if (executeStats)
            {
                root.requestDetails = respFinal;

                console.WriteLine("Creando estadísticas...");
                root.BDUserCreatedBy = "appmanagement";

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
    }
}
