using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using Newtonsoft.Json.Linq;
using System.Web;
using System.IO;
using Newtonsoft.Json;

namespace DataBotV5.Data.Projects.CrBids
{
    /// <summary>
    /// Clase Data con todos los métodos de apoyo para extraer data SQL.
    /// </summary>
    class BidsGbCrSql : IDisposable
    {
        Rooting root = new Rooting();
        Credentials.Credentials cred = new Credentials.Credentials();
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();
        Database.Database db = new Database.Database();
        private bool disposedValue;
        public string database = "licitaciones_cr";
        string mandante = "QAS";
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="bd"></param>
        /// <param name="tabla"></param>
        /// <param name="dictionary"></param>
        /// <param name="type">1 insert 2 update</param>
        /// <param name="key">sólo para update es la columna del WHERE ej: update xxxx where llave = valor </param>
        /// <param name="value">sólo para update es valor para filtrar del WHERE ej: update xxxx where llave = valor</param>
        /// <returns></returns>
        public string SqlCreate(string tabla, Dictionary<string, string> dictionary, int type, [Optional] string key, [Optional] string value)
        {
            string sql = "";
            if (type == 1)
            {
                sql = "INSERT INTO `" + tabla + "`(";
                string sql_insert = " VALUES(";

                foreach (KeyValuePair<string, string> pair in dictionary)
                {

                    if (sql_insert == " VALUES(")
                    {
                        sql = sql + " `" + pair.Key.ToString() + "`";
                        sql_insert = sql_insert + " '" + pair.Value.ToString() + "'";
                    }
                    else
                    {
                        sql = sql + ", `" + pair.Key.ToString() + "`";
                        sql_insert = sql_insert + ", '" + pair.Value.ToString() + "'";

                    }

                }
                sql = sql + ")";
                sql_insert = sql_insert + ")";
                sql = sql + sql_insert;

            }
            else
            {
                sql = "UPDATE `" + tabla + "` SET";
                foreach (KeyValuePair<string, string> pair in dictionary)
                {
                    if (sql == "UPDATE `" + tabla + "` SET")
                    {
                        sql = sql + " `" + pair.Key.ToString() + "` = '" + pair.Value.ToString() + "'";
                    }
                    else
                    {
                        sql = sql + ", `" + pair.Key.ToString() + "` = '" + pair.Value.ToString() + "'";
                    }
                }
                sql = sql + " WHERE " + key + " = '" + value + "'";
            }
            return sql;
        }



        /// <summary>
        /// Insertar una fila en una tabla
        /// </summary>
        /// <param name="dictionary">los campos que desea insertar, importante: los keys deben de ser igual al nombre de la columna y el valor del diccionario es el valor de la columna</param>
        /// <returns></returns>
        public bool InsertRowSS(IDictionary<string, JObject> dictionary)
        {
            bool respuesta = true;
            string idBidNumberPo = "";
            try
            {
                string sqlInsert = "";
                foreach (KeyValuePair<string, JObject> pair in dictionary)
                {
                    string table = pair.Key.ToString();
                    string json = pair.Value.ToString();

                    JObject jnom = JObject.Parse(json);
                    if (table == "products")
                    {
                        string jarray = jnom["products"].ToString();
                        DataTable productos = (DataTable)JsonConvert.DeserializeObject(jarray, (typeof(DataTable)));
                        foreach (DataRow row in productos.Rows)
                        {
                            string sql = $@"INSERT INTO `{table}` (`bidNumber`, `departure`, `line`, `code`, `name`, `amount`, `unit`, `unitPrice`, `active`, `createdBy`)
                                        VALUES ({idBidNumberPo},'{row["Partida"]}','{row["Linea"]}','{row["Codigo"]}','{row["Nombre"]}','{row["Cantidad"]}','{row["Unidad"]}','{row["Precio_Unitario"]}', 1,'databot')";
                            crud.Insert(sql, "costa_rica_bids_db");
                        }
                    }
                    else if (table == "evaluations")
                    {
                        string jarray = jnom["evaluations"].ToString();
                        DataTable evaluations = (DataTable)JsonConvert.DeserializeObject(jarray, (typeof(DataTable)));
                        foreach (DataRow row in evaluations.Rows)
                        {
                            string sql = $@"INSERT INTO `evaluations` (`bidNumber`, `numFactor`, `evaluationPercent`, `evaluationFactor`, `referenceValue`, `active`, `createdBy`)
                                        VALUES ({idBidNumberPo},'{row["Numero_de_factor"]}','{row["Porcentaje_de_evaluacion"]}','{row["Factores_de_evaluacion"]}','{row["Valor_de_referencia"]}', 1,'databot')";
                            crud.Insert(sql, "costa_rica_bids_db");
                        }
                    }
                    else
                    {

                        string sql = "INSERT INTO `" + table + "` (";
                        string sql_insert = " VALUES (";
                        foreach (JProperty x in (JToken)jnom)
                        {
                            string name = x.Name;
                            string value = x.Value.ToString();

                            if (sql_insert == " VALUES (")
                            {
                                sql = sql + " `" + name + "`";
                                if (value == "NULL")
                                {
                                    sql_insert = sql_insert + " " + value;
                                }
                                else
                                {
                                    sql_insert = sql_insert + " '" + value + "'";
                                }
                            }
                            else
                            {
                                sql = sql + ", `" + name + "`";
                                if (value == "NULL")
                                {
                                    sql_insert = sql_insert + ", " + value;
                                }
                                else
                                {
                                    sql_insert = sql_insert + ", '" + value + "'";
                                }


                            }


                        }
                        if (table == "purchaseOrderAdditionalData")
                        {
                            sql = sql + ", `active`, `createdBy`, `bidNumber`)";
                            sql_insert = sql_insert + $", 1,'databot', {idBidNumberPo})";
                        }
                        else
                        {
                            sql = sql + ", `active`, `createdBy`)";
                            sql_insert = sql_insert + ", 1,'databot')";
                        }

                        sql = sql + sql_insert;
                        long id = crud.NonQueryAndGetId(sql, "costa_rica_bids_db");
                        if (table == "purchaseOrder")
                        {
                            if (id == 0)
                            {
                                return false;
                            }
                            idBidNumberPo = id.ToString();
                        }
                    }
                }


                respuesta = true;

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;
        }


        #region ModifyBidMethods



        /// <summary>
        /// Método para insertar archivos a el FTP de SmartAndSimple, el cual según el ambiente inserta a la tabla
        /// uploadFiles de SmartAndSimple el registro del nombre del nuevo archivo, e inserta el archivo al servidor de SmartAndSimple através del FTP.
        /// </summary>
        /// <param name="bidNumber"></param>
        /// <param name="filePathName"></param>
        /// <param name="enviroment"> "PRD" o "DEV"</param>
        /// <returns></returns>
        public bool InsertFile(string bidNumber, string filePathName, string enviroment)
        {
            try
            {
                string user = "";
                string fileName = Path.GetFileName(filePathName);
           
                if (enviroment == "QAS")
                {
                    user = cred.QA_SS_APP_SERVER_USER;
                }
                else if (enviroment == "PRD")
                {
                    user = cred.PRD_SS_APP_SERVER_USER;
                }

                string pathfile = $"/home/{user}/projects/smartsimple/gbm-hub-api/src/assets/files/CrBids/{bidNumber}/{fileName}";
                string mimeType = MimeMapping.GetMimeMapping(fileName);
                string sql2 = "INSERT INTO `uploadFiles` (`name`, `bidNumber`, `user`, `codification`, `type`, `path`, `active`, `createdBy`) VALUES " +
                $"('{fileName}', '{bidNumber}', 'databot', '7bit', '{mimeType}', '{pathfile}', 1, 'databot')";


                crud.Insert(sql2, "costa_rica_bids_db");


                //subir al FTP de S&S
                bool subir_files = db.uploadSftp(filePathName, $"/home/{user}/projects/smartsimple/gbm-hub-api/src/assets/files/CrBids", bidNumber);

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        #endregion



        /// <summary>
        /// Obtiene una lista con las palabras clave para saber si el concurso es de interes
        /// </summary>
        /// <returns></returns>
        public List<string> KeyWord(string mandante)
        {
            ValidateData val = new ValidateData();
            List<string> stopWords = new List<string>();
            try
            {
                DataTable mytable = crud.Select( "SELECT * FROM keyWords", "costa_rica_bids_db");
                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        string key = mytable.Rows[i][2].ToString().ToLower();
                        key = key.Replace("á", "a"); key = key.Replace("é", "e"); key = key.Replace("í", "i"); key = key.Replace("ó", "o"); key = key.Replace("ú", "u");
                        key = val.RemoveSpecialChars(key, 1);
                        stopWords.Add(key); //palabra clave
                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
            }
            return stopWords;
        }


        /// <summary>
        /// Método especializado para mover a Backup una licitación de Costa Rica, específicamente en la base de datos de 
        /// costa_rica_bids de SmartAndSimple, el cual toma en cuenta que una purchaseOrder(Licitación), tiene relacionados products, PurchaseOrderAddtionalData,
        /// evaluations, salesTeam, los cuales esos los mueve a sus respectivas tablas de Backup también. 
        /// </summary>
        /// <param name="idBidNumber"></param>
        /// <param name="enviroment"></param>
        /// <returns>Retorna una variable tipo bool, indicando si el movimiento de la licitación a backup fue exitosa. </returns>
        public bool MoveBidToBackup(string idBidNumber)
        {
            try
            {
                //primero hace el Insert en purchaseOrderBackup
                string sql = $"INSERT INTO purchaseOrderBackup SELECT * FROM purchaseOrder WHERE id = '{idBidNumber}'";
                bool insert = crud.Insert(sql, "costa_rica_bids_db");
                if (insert)
                {
                    bool salesTeam = MoveAndDelete("salesTeam", "salesTeamBackup", idBidNumber);
                    bool evaluations = MoveAndDelete("evaluations", "evaluationsBackup", idBidNumber);
                    bool products = MoveAndDelete("products", "productsBackup", idBidNumber);
                    bool purchaseOrderAdditionalData = MoveAndDelete("purchaseOrderAdditionalData", "purchaseOrderAdditionalDataBackup", idBidNumber);
                    if (salesTeam && evaluations && products && purchaseOrderAdditionalData)
                    {
                        //si todo se movio bien se elimina el registro principal
                        string sqlDelete = $"DELETE FROM `purchaseOrder` WHERE id = '{idBidNumber}'";
                        bool delete = crud.Delete(sqlDelete, "costa_rica_bids_db");
                        return delete;
                    }
                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
            }
            return false;
        }

        /// <summary>
        /// Esta función es utilizada principalmente en MoveBidToBackup de esta misma clase de BidsGbCrSql, el cual funciona para mover 
        /// a backup un registro sea de products, evaluations, salesTeam, purchaseOrderAdditionalData, y posteriormente eliminarlo de las tablas 
        /// con los registros actuales, de la base de datos de Smart And Simple.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="tableBackup"></param>
        /// <param name="idBidNumber"></param>
        /// <param name="enviroment"></param>
        /// <returns>Retorna una variable tipo bool, indicando si el movimiento a backup fue exitosa. </returns>
        public bool MoveAndDelete(string table, string tableBackup, string idBidNumber)
        {
            //hace el primer insert en poBackup
            string sqlInsert = $"INSERT INTO {tableBackup} SELECT * FROM {table} WHERE bidNumber = '{idBidNumber}'";
            string sqlDelete = $"DELETE FROM `{table}` WHERE bidNumber = '{idBidNumber}'";
            bool insert = crud.Insert(sqlInsert, "costa_rica_bids_db");
            if (insert)
            {
                //si el insert funciona lo elimina de poBackup
                bool delete = crud.Delete(sqlDelete, "costa_rica_bids_db");
                return delete;
            }
            return false;
        }
        #region ModifyBidSQL
        public bool UpdateRowModifyBid(string bd, string tabla, string key, string valor, Dictionary<string, string> dictionary)
        {



            bool respuesta = false;
            try
            {
                #region Connection DB



                string sql_update = SqlCreate(tabla, dictionary, 2, key, valor);
                new CRUD().Update(sql_update, "costa_rica_bids_db");
                #endregion
                respuesta = true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;



            }



            return respuesta;



        }

        #endregion
        public DataTable InsertRowSSOriginal(IDictionary<string, JObject> dictionary, DataTable results, string numBid)
        {

            bool respuesta = true;
            string idBidNumberPo = "";
            DataTable dt = crud.Select( $"select id from purchaseOrder where bidNumber = '{numBid}'", "costa_rica_bids_db");
            if (dt.Rows.Count == 0)
            {
                return dt;
            }
            idBidNumberPo = dt.Rows[0]["id"].ToString();
            try
            {
                string sqlInsert = "";
                foreach (KeyValuePair<string, JObject> pair in dictionary)
                {
                    DataRow nRow = results.Rows.Add();

                    nRow["concurso"] = numBid;
                    string query = "";

                    string table = pair.Key.ToString();
                    string json = pair.Value.ToString();

                    JObject jnom = JObject.Parse(json);
                    if (table == "products")
                    {
                        string jarray = jnom["products"].ToString();
                        DataTable productos = (DataTable)JsonConvert.DeserializeObject(jarray, (typeof(DataTable)));
                        foreach (DataRow row in productos.Rows)
                        {
                            string sql = $@",'{row["Partida"]}','{row["Linea"]}','{row["Codigo"]}','{row["Nombre"]}','{row["Cantidad"]}','{row["Unidad"]}','{row["Precio_Unitario"]}', 1,'databot');";
                            //crud.Insert(sql, "costa_rica_bids_db");
                            query = sql;
                        }
                    }
                    else if (table == "evaluations")
                    {
                        string jarray = jnom["evaluations"].ToString();
                        DataTable evaluations = (DataTable)JsonConvert.DeserializeObject(jarray, (typeof(DataTable)));
                        foreach (DataRow row in evaluations.Rows)
                        {
                            string sql = $@",'{row["Numero_de_factor"]}','{row["Porcentaje_de_evaluacion"]}','{row["Factores_de_evaluacion"]}','{row["Valor_de_referencia"]}', 1,'databot');";
                            //crud.Insert(sql, "costa_rica_bids_db");
                            query = sql;
                        }
                    }
                    else
                    {

                        string sql = "INSERT INTO `" + table + "` (";
                        string sql_insert = " VALUES (";
                        foreach (JProperty x in (JToken)jnom)
                        {
                            string name = x.Name;
                            string value = x.Value.ToString();

                            if (sql_insert == " VALUES (")
                            {
                                sql = sql + " `" + name + "`";
                                if (value == "NULL")
                                {
                                    sql_insert = sql_insert + " " + value;
                                }
                                else
                                {
                                    sql_insert = sql_insert + " '" + value + "'";
                                }
                            }
                            else
                            {
                                sql = sql + ", `" + name + "`";
                                if (value == "NULL")
                                {
                                    sql_insert = sql_insert + ", " + value;
                                }
                                else
                                {
                                    sql_insert = sql_insert + ", '" + value + "'";
                                }


                            }


                        }
                        if (table == "purchaseOrderAdditionalData")
                        {
                            sql = sql + ", `active`, `createdBy`, `bidNumber`)";
                            sql_insert = sql_insert + $", 1,'databot', {idBidNumberPo});";
                            sql = sql + sql_insert;
                            bool id = crud.Insert(sql, "costa_rica_bids_db");
                            
                        }
                        else
                        {
                            sql = sql + ", `active`, `createdBy`)";
                            sql_insert = sql_insert + ", 1,'databot');";
                        }

                        sql = sql + sql_insert;
                        query = sql;

                    }

                    nRow["query"] = query;
                    nRow["tabla"] = table;
                    results.AcceptChanges();
                }


                respuesta = true;

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;
                DataRow nRow = results.Rows.Add();
                nRow["concurso"] = numBid;
                nRow["query"] = ex.ToString();
                nRow["tabla"] = "NA";
                results.AcceptChanges();

            }
            return results;
        }
        #region disposable
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
        // ~MailInteraction()
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
