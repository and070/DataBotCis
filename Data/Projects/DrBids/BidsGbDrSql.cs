using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using static DataBotV5.Automation.MASS.DrBids.GetDrBids;
using DataBotV5.Data.Database;
using DataBotV5.App.Global;
using DataBotV5.App.ConsoleApp;

namespace DataBotV5.Data.Projects.DrBids
{
    /// <summary>
    /// Clase Data encargada de sql de licitaciones de dominicana. 
    /// </summary>
    class BidsGbDrSql
    {
        ConsoleFormat console = new ConsoleFormat();
        /// <summary>
        /// Se encarga de cargar los datos traidos de un objeto licitaciones para insertarlo a la base de datos
        /// </summary>
        /// <param name="bd"></param> Nombre de la vase de datos 
        /// <param name="bids"></param> Objeto que trae los datos necesarios para insertar
        /// <returns></returns>
        public bool InsertRow( Bids bids)
        {
            //10.7.60.72
            string sql_insert = "";
            bool respuesta = true;
            try
            {
                #region Connection DB     
                MySqlConnection conn = new Database.Database().ConnSmartSimple("dr_bids_db", Start.enviroment);
                #endregion

                sql_insert = "INSERT INTO bids (generalData,sapData,status,articles,planification) VALUES (@datos,@datosSap,@estado,@articulos,@planificacion)";
                MySqlCommand mySqlCommand = new MySqlCommand(sql_insert, conn);
                mySqlCommand.Parameters.Add("@datos", MySqlDbType.LongText).Value = JsonConvert.SerializeObject(bids.DG);
                mySqlCommand.Parameters.Add("@datosSap", MySqlDbType.LongText).Value = JsonConvert.SerializeObject(bids.DS);
                mySqlCommand.Parameters.Add("@estado", MySqlDbType.Int32).Value = 0;
                mySqlCommand.Parameters.Add("@articulos", MySqlDbType.LongText).Value = bids.AT;
                mySqlCommand.Parameters.Add("@planificacion", MySqlDbType.LongText).Value = bids.PL;
                conn.Open();
                mySqlCommand.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;
            }
            return respuesta;

        }
        /// <summary>
        /// Devuelve una tabla con los datos extraidos de una base de datos 
        /// </summary>
        /// <param name="query"></param> Es la sentencia sql que se utilizara para buscar
        /// <returns></returns>
        public DataTable SelectRow(string query)
        {
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB                 
                mytable = new CRUD().Select(query, "dr_bids_db");
                #endregion

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
            }
            return mytable;
        }    /// <param name="bd"></param> Especifica el nombre de la base de datos de la cual vamos a extraer la tabla

        public bool InsertFiles(Bids bids, string fileName)
        {
            string sql_insert = "";
            bool respuesta = true;
            try
            {
                //#region Connection DB   
                //MySqlConnection conn = new Database.Database().Conn("dr_bids_db"); 
                //#endregion

                //sql_insert = "INSERT INTO documentos_dominicana(Licitacion,Nombre,Archivos) VALUES(@licitacion,@nombre,@archivos)";
                //MySqlCommand mySqlCommand = new MySqlCommand(sql_insert, conn);
                //mySqlCommand.Parameters.Add("@licitacion", MySqlDbType.Text).Value = bids.DG.referencia;
                //mySqlCommand.Parameters.Add("@nombre", MySqlDbType.Text).Value = fileName;
                //mySqlCommand.Parameters.Add("@archivos", MySqlDbType.LongBlob, bids.AJ.Length).Value = bids.AJ;
                //conn.Open();
                //mySqlCommand.ExecuteNonQuery();
                //conn.Close();
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;
            }
            return respuesta;
        }

        public List<string> TraerIdBids()
        {
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB                 
                string query = "SELECT JSON_VALUE(generalData,'$.referencia') FROM `bids`";
                mytable = new CRUD().Select(query, "dr_bids_db");
                #endregion
            }
            catch
            {
            }
            List<string> idLicitaciones = new List<string>();
            for (int i = 0; i < mytable.Rows.Count; i++)
            {
                idLicitaciones.Add(mytable.Rows[i][0].ToString());
            }

            return idLicitaciones;
        }
        public List<ClientSAP> FetchSAPCostumers()
        {
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB 
                string query = "SELECT purchasingUnit, salesRep FROM purchasingUnits";
                mytable = new CRUD().Select(query, "dr_bids_db");
                #endregion
            }
            catch
            {
            }
            List<ClientSAP> clientesSAP = new List<ClientSAP>();
            for (int i = 0; i < mytable.Rows.Count; i++)
            {
                ClientSAP cliente = new ClientSAP
                {
                    unidadCompra = mytable.Rows[i][0].ToString(),
                    nombreVendedor = mytable.Rows[i][1].ToString()
                };
                clientesSAP.Add(cliente);
            }

            return clientesSAP;
        }

        public bool ValidateRepetidos(List<string> lista, string idBid)
        {
            bool validado = false;
            var result = lista.Where(x => idBid.Contains(x)).ToList();
            if (result.Count > 0)
            {
                validado = true;
                return validado;
            }
            return validado;
        }

        public class ClientSAP
        {
            public string unidadCompra { get; set; }
            public string nombreVendedor { get; set; }
        }
    }
}
