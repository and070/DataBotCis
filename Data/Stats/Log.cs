using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

using MySql.Data.MySqlClient;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;
using DataBotV5.App.Global;
using DataBotV5.App.ConsoleApp;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Webex;

namespace DataBotV5.Data.Stats
{
    /// <summary>
    /// Clase Data para el login.
    /// </summary>
    class Log
    {
        string tabla;
        string sql_insert;
        string sql_vars;
        string sql_get;
        float trpa;
        float tarifa;
        int ejecuciones;
        DataTable mytable = new DataTable();
        DataTable mytable2 = new DataTable();
        ConsoleFormat console = new ConsoleFormat();
        CRUD crud = new CRUD();
        Start start = new Start();
        string enviroment = Start.enviroment;
        Rooting root = new Rooting();
        WebexTeams wt = new WebexTeams();
        //Credentials cred = new Credentials();
        /// <summary>
        /// Crear log de cambios para las modificaciones
        /// </summary>
        /// <param name="type">Tipo de proceso, Cracion o Modificacion</param>
        /// <param name="nProcess">Nombre del proceso</param>
        /// <param name="applicant">Nombre o correo del solicitante</param>
        /// <param name="change">Parametro a realizar el cambio o creacion, Ej Representante de Ventas</param>
        /// <param name="content">Contenido del dato cambiado</param>
        /// <param name="comments">Comentarios adicionales</param>
        public void LogDeCambios(string type, string nProcess, string applicant, string change,
            string content, string comments)
        {
            try
            {
                sql_insert = $@"INSERT INTO `botlog`
                                (`class`, `changeLog`, `contents`, `comments`, `createdBy`) 
                                VALUES 
                                ('{root.BDIdClass}', '{change}', '{content}', '{comments}', '{applicant}')";

                crud.Insert(sql_insert, "databot_db", enviroment);

            }
            catch (Exception ex)
            { }

            #region
            crud.Insert(sql_insert, "databot_db", "QAS");
            #endregion
        }
        public void RegisterNeewClient(int bp, string name, string country, string territory, string salesRepID, string cocode, string address, string telephone, string email)
        {
            //
            string salesRep = "";
            string manager = "";
            string digitalSignId = "";

            //MySqlConnection conn = new Database.Database().Conn("fabrica_de_ofertas", enviroment);
            if (string.IsNullOrEmpty(salesRepID))
            {
                try
                {
                    DataTable zCustomerData = new DataTable();
                    using (SapVariants sap = new SapVariants())
                    {
                        zCustomerData = sap.GetSapTable("ZCUSTOMERDATA", "ERP");
                    }
                    zCustomerData.Rows.Cast<DataRow>().ToList().ForEach(DataRow =>
                    {
                        string fila = DataRow["WA"].ToString();

                        if (fila.Contains(cocode))
                        {
                            salesRepID = "AA" + DataRow["WA"].ToString().Split(new string[] { "AA" }, StringSplitOptions.None)[1];
                            return;
                        }
                    });
                }
                catch (Exception ex)
                {

                }
            }
            if (salesRepID != "")
            {
                #region Extraer usuario del vendedor

                DataTable empleados = new DataTable();
                try
                {

                    string sqlSelect = $"SELECT * FROM digital_sign WHERE UserID = {salesRepID.ToString().Replace("AA","")} ";
                    empleados = crud.Select( sqlSelect, "MIS", enviroment);

                    if (empleados.Rows.Count > 0)
                    {
                        DataRow salesRepDr = empleados.Rows[0];
                        salesRep = salesRepDr["user"].ToString();
                        digitalSignId = salesRepDr["id"].ToString();
                    }


                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.ToString());

                }
                if (string.IsNullOrEmpty(salesRep))
                {
                    using (SapVariants sap = new SapVariants())
                    {
                        salesRep = sap.AmEmail(salesRepID).Split('@')[0];
                    }
                }

                #endregion
                #region GET Manager
                DataTable row = new DataTable();
                try
                {
                    

                    string sqlSelect = $"SELECT *  FROM clients WHERE accountManagerUser= '{salesRep}' and manager!=''";
                    row = crud.Select( sqlSelect, "databot_db", enviroment);

                    if (row.Rows.Count > 0)
                    {
                        DataRow salesRepDr = row.Rows[0];
                        manager = salesRepDr["manager"].ToString();
                    }

                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.ToString());

                }
                #endregion
                salesRepID = salesRepID.ToString().TrimStart('A');
            }

            #region Extraer el territorio
            switch (territory)
            {
                case "001":
                    territory = "Premium Account";
                    break;
                case "002":
                    territory = "GBM Direct";
                    break;
                case "003":
                    territory = "VT TELCO";
                    break;
                case "004":
                    territory = "VT Public Sector";
                    break;
                case "005":
                    territory = "VT Lgstcs & Commerce";
                    break;
                case "006":
                    territory = "VT Bnkng & Finance";
                    break;
                case "007":
                    territory = "VT Bnkng & Finance 2";
                    break;
                default:
                    territory = "GBM Direct";
                    break;
            }

            #endregion


            #region Insertar client en databot_db - .138

            string sqlCountries = $"SELECT id,countryCode  FROM `sapCountries` WHERE `countryCode` = '{country}' and active=1";
            DataTable countriesTable = crud.Select( sqlCountries, "databot_db", enviroment);

            string sqlValueTeam = $"SELECT id, valueTeam FROM `valueTeam` WHERE `valueTeam` = '{territory}' and active=1";
            DataTable valueTeamTable = crud.Select( sqlValueTeam, "databot_db", enviroment);

            string countrySS = countriesTable.Rows[0]["id"].ToString();
            string territorySS = valueTeamTable.Rows[0]["id"].ToString();

            string queryInsertClientSS = "INSERT INTO `clients` (`id`, `idClient`, `name`, `country`, `territory`, address, telephone, email, employeeResponsible, `accountManagerId`, `accountManagerUser`, `manager`, `locked`, `active`, `createdAt`, `createdBy`, `updatedAt`, `updatedBy`) " +
                $" VALUES (NULL, '{bp.ToString()}', '{name}', '{countrySS}', {territorySS}, '{address}', '{telephone}', '{email}', {(digitalSignId == "" ? "NULL" : $"'{digitalSignId}'")}, '{salesRepID}', '{salesRep}', '{manager}', '', '1', CURRENT_TIMESTAMP, 'epiedra', NULL, NULL);   ";

            bool insert = crud.Insert(queryInsertClientSS, "databot_db", enviroment);
            if (!insert)
            {
                wt.SendNotification("dmeza@gbm.net", "", $"Error al ingresar cliente dentro la DB <br><br> {queryInsertClientSS}");
            }
            #endregion
        }

    }
    #region Clases de ahorro RPA
    public class AhorroAnual
    {
        public string Nombre { get; set; }
        public double Tarifa { get; set; }
        public int Minutos { get; set; }
        public double Factor { get; set; }
        public double Ahorro { get; set; }
    }
    public class AhorroMensual
    {
        public string Nombre { get; set; }
        public int Cantidad { get; set; }
    }
    #endregion
    public class CredentialsE
    {
        public string User { get; set; }
        public string Pass { get; set; }
        public CredentialsE(string usuarioSistema)
        {
            switch (usuarioSistema)
            {
                case "databot01":
                    User = "analytics_01";
                    Pass = "9MIeFMTgfyF5K9dw";
                    break;
                case "databot02":
                    User = "analytics_02";
                    Pass = "PS2bsAVOHzXsGoXB";
                    break;
                case "databot03":
                    User = "analytics_03";
                    Pass = "3G2YE5mdYcFf6IYp";
                    break;
                case "databot04":
                    User = "analytics_04";
                    Pass = "3PwJV9gE9V8bSMGr";
                    break;
                default:
                    User = "databot";
                    Pass = "UqJkkoxRVkIXSJYf";
                    break;
            }
        }
    }
}
