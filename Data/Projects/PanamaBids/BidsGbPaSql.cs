using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Automation.MASS.PanamaBids;
using DataBotV5.Automation.RPA2.PanamaBids;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;

namespace DataBotV5.Data.Projects.PanamaBids
{
    /// <summary>
    /// Clase Data con métodos para extraer, insertar o modificar información SQL para Licitaciones de PA.
    /// </summary>
    class BidsGbPaSql
    {
        Credentials.Credentials cred = new Credentials.Credentials();
        ConsoleFormat console = new ConsoleFormat();
        CRUD crud = new CRUD();
        ValidateData val = new ValidateData();
        Log log = new Log();
        Rooting root = new Rooting();


        string ssMandante = "PRD";

        public DataTable GetInfo(string registro, [Optional] string quote)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable registo_unico_info = new DataTable();
            bool respuesta = false;
            try
            {
                #region Connection DB
                sql_select = (String.IsNullOrEmpty(quote)) ? "select * from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + registro + "'" : "select * from reporte_convenio_gbmpa where QUOTE = '" + quote + "'";
                //mytable = crud.Select("Databot", sql_select, "ventas");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    registo_unico_info = mytable;
                    respuesta = true;
                }
                else
                {
                    respuesta = false;
                }
            }
            catch (Exception ex)
            { respuesta = false; }



            return registo_unico_info;
        }
        public DataTable TypeProduct()
        {
            DataTable mytable = new DataTable();
            string[] info = new string[3];
            try
            {
                #region Connection DB
                //mytable = crud.Select("Databot", "select * from productos_compras_panama", "ventas");
                #endregion
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }
            return mytable;
        }
        public string[] MarcaProduct(string id)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            DataTable mytable = new DataTable();
            string[] marca = new string[2];

            try
            {
                #region Connection DB     
                //mytable = crud.Select("Databot", "select * from productos_compras_gbpa where id = " + id, "ventas");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    marca[0] = mytable.Rows[0][2].ToString(); //marca
                    marca[1] = mytable.Rows[0][3].ToString(); //precio
                }
                else
                {
                    marca[0] = "No se encontró este producto en la lista";

                }


            }
            catch (Exception ex)
            {
                marca[0] = "No se encontró este producto en la lista";
            }

            return marca;
        }
        public string[] DetermineSector(string entity)
        {
            DataTable mytable = new DataTable();
            string[] ent_info = new string[5];

            try
            {
                #region Connection DB
                //mytable = crud.Select("Databot", "select * from sector where entidad = '" + entity + "'", "ventas");
                #endregion

                //si lo encontró
                if (mytable.Rows.Count > 0)
                {
                    ent_info[0] = mytable.Rows[0][2].ToString(); //sector
                    ent_info[1] = mytable.Rows[0][3].ToString(); //cliente
                    ent_info[2] = mytable.Rows[0][4].ToString(); //contacto
                    ent_info[3] = mytable.Rows[0][5].ToString(); //sales rep
                    ent_info[4] = mytable.Rows[0][6].ToString(); //sales rep cotizaciones
                    //lo encontró pero esta vacio
                    if (String.IsNullOrEmpty(ent_info[1]))
                    {
                        ent_info[0] = "PS"; //sector
                        ent_info[1] = "0010067544"; //cliente
                        ent_info[2] = "0070032409"; //contacto
                        ent_info[3] = "AA70000134"; //sales rep
                        ent_info[4] = "AA70000134"; //sales rep cotizaciones
                    }
                }
                //no encontró la entidad
                else
                {
                    ent_info[0] = "PS"; //sector
                    ent_info[1] = "0010067544"; //cliente
                    ent_info[2] = "0070032409"; //contacto
                    ent_info[3] = "AA70000134"; //sales rep
                    ent_info[4] = "AA70000134"; //sales rep cotizaciones
                }

            }
            catch (Exception ex)
            {
                //si da error por X situación 
                ent_info[0] = "PS"; //sector
                ent_info[1] = "0010067544"; //cliente
                ent_info[2] = "0070032409"; //contacto
                ent_info[3] = "AA70000134"; //sales rep
                ent_info[4] = "AA70000134"; //sales rep cotizaciones
            }

            return ent_info;
        }
        public string OppExist(string registro)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            string opp_actual = "";
            DataTable mytable = new DataTable();
            bool respuesta = false;
            try
            {
                //mytable = crud.Select("Databot", "select OPORTUNIDAD from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + registro + "'", "ventas");



                if (mytable.Rows.Count > 0)
                {
                    opp_actual = mytable.Rows[0][0].ToString(); //opp

                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                opp_actual = "";

            }

            return opp_actual;
        }
        public string EmailCpa(string filter)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            string jemail = "";
            DataTable mytable = new DataTable();
            bool respuesta = false;
            try
            {
                //mytable = crud.Select("Databot", "select JEMAIL from email_address where CATEGORIA = '" + filter + "'", "licitaciones_cr");

                if (mytable.Rows.Count > 0)
                {
                    jemail = mytable.Rows[0][0].ToString(); //opp

                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                jemail = "";

            }

            return jemail;
        }
        public bool ExistQuote(string quote)
        {
            bool respuesta = true;
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            string opp_actual = "";
            DataTable mytable = new DataTable();
            try
            {
                //mytable = crud.Select("Databot", "select * from reporte_cotizaciones_linea where NUM_COTIZACION = '" + quote + "'", "ventas");

                respuesta = (mytable.Rows.Count > 0) ? true : false;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }

            return respuesta;
        }
        public List<string> KeyWord(string category)
        {
            List<string> stopWords = new List<string>();
            //string[] lines = File.ReadAllLines("C:\stopwords.txt");

            DataTable mytable = new DataTable();
            try
            {
                //mytable = crud.Select("Databot", "SELECT * FROM `key_words_convenio`", "ventas");

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        string key = mytable.Rows[i][2].ToString().ToLower();
                        key = key.Replace("á", "a"); key = key.Replace("é", "e"); key = key.Replace("í", "i"); key = key.Replace("ó", "o"); key = key.Replace("ú", "u");
                        //key = val.QuitarEnne(key);
                        key = val.RemoveSpecialChars(key, 1);
                        stopWords.Add(key); //palabra clave
                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());

            }

            return stopWords;
        }
        public string[] AllQuotes()
        {
            string[] cotis = new string[1];
            //string[] lines = File.ReadAllLines("C:\stopwords.txt");

            DataTable mytable = new DataTable();
            try
            {
                //mytable = crud.Select("Databot", "SELECT DISTINCT NUM_COTIZACION FROM `reporte_cotizaciones_linea` GROUP BY NUM_COTIZACION", "ventas");

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        string coti = mytable.Rows[i][0].ToString();
                        cotis[i] = coti;
                        Array.Resize(ref cotis, cotis.Length + 1);
                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());

            }
            Array.Resize(ref cotis, cotis.Length - 1);
            return cotis;
        }
        public Dictionary<string, string> getAllEntity()
        {
            Dictionary<string, string> entidades = new Dictionary<string, string>();
            DataTable mytable = new DataTable();
            try
            {
                //mytable = crud.Select("Databot", "SELECT entidad, sales_rep_coti FROM `sector`", "ventas");

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        string entidad = mytable.Rows[i][0].ToString();
                        string salesrep = mytable.Rows[i][1].ToString();
                        entidades[entidad] = salesrep;


                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());

            }
            return entidades;
        }
        public bool UpdateSubtotal(string registroUnico, string subTotal)
        {
            bool resp_add_sql = true;
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB     
                //mytable = crud.Select("Databot", "select * from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + registroUnico + "'", "ventas");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    //crud.Update("Databot", "UPDATE `reporte_convenio_gbmpa` SET `SUB_TOTAL_ORDEN`= '" + subTotal.ToString() + "' WHERE REGISTRO_UNICO_DE_PEDIDO = '" + registroUnico + "'", "ventas");
                }

            }
            catch (Exception ex)
            {
                resp_add_sql = false;
            }
            return resp_add_sql;
        }
        public string[] ConvenioColumns(string table)
        {
            SqlConnection myConn = new SqlConnection();
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable update = new DataTable();
            string[] column = new string[1];

            try
            {
                #region Connection DB   
                string sql_columns = "SHOW COLUMNS FROM " + table;
                //mytable = crud.Select("Databot", sql_columns, "ventas");
                #endregion

                if (columnas.Rows.Count > 0) //si hay columnas
                {
                    for (int i = 1; i < columnas.Rows.Count; i++)
                    {
                        string columna = columnas.Rows[i][0].ToString();
                        column[i - 1] = columna;
                        Array.Resize(ref column, column.Length + 1);
                    }
                    Array.Resize(ref column, column.Length - 1);

                }
            }
            catch (Exception ex)
            {

            }
            return column;
        }

        #region agregar / update info SQL
        /// <summary>
        /// agregar o actualizar informacion a la base de datos
        /// </summary>
        /// <param name="convenio"></param>
        /// <param name="entity"></param>
        /// <param name="product"></param>
        /// <param name="amount"></param>
        /// <param name="marca"></param>
        /// <param name="total"></param>
        /// <param name="subTotal"></param>
        /// <param name="purchaseOrder"></param>
        /// <param name="registroUnico"></param>
        /// <param name="registrationDate"></param>
        /// <param name="dateDoc"></param>
        /// <param name="fianza"></param>
        /// <param name="opp"></param>
        /// <param name="quote"></param>
        /// <param name="orderType"></param>
        /// <param name="salesOrder"></param>
        /// <param name="gbmState"></param>
        /// <param name="statusOrder"></param>
        /// <param name="daysDelivery"></param>
        /// <param name="dateMax"></param>
        /// <param name="diasFaltantes"></param>
        /// <param name="forecast"></param>
        /// <param name="country"></param>
        /// <param name="place"></param>
        /// <param name="funcionario"></param>
        /// <param name="telephone"></param>
        /// <param name="email"></param>
        /// <param name="montoMulta"></param>
        /// <param name="pdfName"></param>
        /// <param name="poUrl"></param>
        /// <returns></returns>
        public bool InfoSqlAddGbm(string convenio, string entity, string product, string amount, string marca, string total, string subTotal, string purchaseOrder, string registroUnico, DateTime registrationDate, DateTime dateDoc, string fianza, string opp, string quote, string orderType, string salesOrder, string gbmState, string statusOrder, int daysDelivery, DateTime dateMax, double diasFaltantes, DateTime forecast, string country, string place, string funcionario, string telephone, string email, string montoMulta, string pdfName, string poUrl)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_insert = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable update = new DataTable();
            bool respuesta = true;

            try
            {
                string fecha_registro_end = registrationDate.ToString("yyyy-MM-dd");
                string fecha_doc_end = dateDoc.ToString("yyyy-MM-dd");
                string fecha_max_end = dateMax.ToString("yyyy-MM-dd");
                string forecast_end = forecast.ToString("yyyy-MM-dd");

                #region Connection DB
                sql_select = "select * from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + registroUnico + "' and PRODUCTO_SERVICIO = '" + product + "'";
                //mytable = crud.Select("Databot", sql_select, "ventas");
                #endregion

                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    sql_insert = "INSERT INTO `reporte_convenio_gbmpa`(`CONVENIO`, `ENTIDAD`, `PRODUCTO_SERVICIO`, `CANTIDAD`, `MARCA`, " +
                        "`TOTAL_DEL_PRODUCTO`, `SUB_TOTAL_ORDEN`, `ORDEN_COMPRA`, `REGISTRO_UNICO_DE_PEDIDO`, `FECHA_DE_REGISTRO`, " +
                        "`FECHA_DE_PUBLICACION`, `FIANZA_CUMPLIMIENTO`, `OPORTUNIDAD`, `QUOTE`, `TIPO_PEDIDO`, `SALES_ORDER`, `ESTADO_GBM`, " +
                        "`ESTATUS_DE_ORDEN`, `DIAS_ENTREGA`, `FECHA_MAXIMA_ENTREGA`, `DIAS_FALTANTES`, `FORECAST`, `PROVINCIA`, `LUGAR_DE_ENTREGA`, " +
                        "`CONTACTO_DE_ENTREGA`, `TELEFONO`, `EMAIL`, `MONTO_MULTA`, `NOMBRE_DEL_ADJUNTO`, `LINK_AL_DOCUMENTO`)" +

                        " VALUES ('" + convenio + "','" + entity + "','" + product + "','" + amount + "','" + marca + "','" +
                        total + "','" + subTotal + "','" + purchaseOrder + "','" + registroUnico + "','" + fecha_registro_end + "','" +
                        fecha_doc_end + "','" + fianza + "','" + opp + "','" + quote + "','" + orderType + "','" + salesOrder + "','" + gbmState + "','" +
                        statusOrder + "','" + daysDelivery.ToString() + "','" + fecha_max_end + "','" + diasFaltantes.ToString() + "','" + forecast_end + "','" + country + "','" +
                        place + "','" + funcionario + "','" + telephone + "','" + email + "','" + montoMulta + "','" + pdfName + "','" + poUrl + "')";
                    //new CRUD().Insert("Databot", sql_insert, "business system");


                    respuesta = true;
                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;
        }
        /// <summary>
        /// agregar o actualizar informacion a la base de datos, los keys del diccionario deben ser el nombre de la columna
        /// </summary>
        /// <param name="dictionary"></param>
        /// <returns></returns>
        public bool InfoSqlAddGbmV2(Dictionary<string, string> dictionary)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_insert = "";
            string sql_insert2 = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable update = new DataTable();
            bool respuesta = true;


            try
            {
                #region Connection DB
                sql_select = "select * from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + dictionary["REGISTRO_UNICO_DE_PEDIDO"] + "' and PRODUCTO_SERVICIO = '" + dictionary["PRODUCTO_SERVICIO"] + "'";
                //mytable = new CRUD().Select("Databot", sql_select, "ventas");
                #endregion

                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    string sql_columns = "SHOW COLUMNS FROM reporte_convenio_gbmpa";
                    //columnas = new CRUD().Select("Databot", sql_columns, "ventas");

                    if (columnas.Rows.Count > 0) //si hay columnas
                    {
                        sql_insert = "INSERT INTO `reporte_convenio_gbmpa`(";
                        sql_insert2 = " VALUES(";
                        for (int i = 1; i < columnas.Rows.Count; i++)
                        {

                            string columna = columnas.Rows[i][0].ToString();
                            if (i == 1)
                            {
                                sql_insert = sql_insert + " `" + columna + "`";
                                sql_insert2 = sql_insert2 + " '" + dictionary[columna] + "'";
                            }
                            else
                            {
                                sql_insert = sql_insert + ", `" + columna + "`";
                                sql_insert2 = sql_insert2 + ", '" + dictionary[columna] + "'";
                            }

                        }
                        sql_insert = sql_insert + ")";
                        sql_insert2 = sql_insert2 + ")";
                        sql_insert = sql_insert + sql_insert2;

                        //crud.Insert("Databot", sql_insert, "ventas");


                    }

                    respuesta = true;

                }
                else //update
                {
                    string sql_columns = "SHOW COLUMNS FROM reporte_convenio_gbmpa";
                    //MySqlDataAdapter myadapter2 = new MySqlDataAdapter(sql_columns, conn);
                    //myadapter2.Fill(columnas);
                    //mytable = crud.Select("Databot", sql_columns, "ventas");

                    if (columnas.Rows.Count > 0) //si hay columnas
                    {
                        sql_update = "UPDATE `reporte_convenio_gbmpa` SET";
                        for (int i = 1; i < columnas.Rows.Count; i++)
                        {
                            string valor_actual = mytable.Rows[0][i].ToString();
                            if (String.IsNullOrEmpty(valor_actual) || valor_actual == "En Refrendo" || columnas.Rows[i][0].ToString() == "TOTAL_DEL_PRODUCTO" || columnas.Rows[i][0].ToString() == "SUB_TOTAL_ORDEN" || columnas.Rows[i][0].ToString() == "FECHA_DE_REGISTRO" || columnas.Rows[i][0].ToString() == "FECHA_MAXIMA_ENTREGA" || columnas.Rows[i][0].ToString() == "FECHA_DE_PUBLICACION" || columnas.Rows[i][0].ToString() == "TIPO_FORECAST")
                            {
                                string columna = columnas.Rows[i][0].ToString();
                                if (sql_update == "UPDATE `reporte_convenio_gbmpa` SET")
                                {
                                    sql_update = sql_update + " `" + columna + "` = '" + dictionary[columna].ToString() + "'";
                                }
                                else
                                {
                                    sql_update = sql_update + ", `" + columna + "` = '" + dictionary[columna].ToString() + "'";
                                }

                            }
                        }
                        sql_update = sql_update + " WHERE REGISTRO_UNICO_DE_PEDIDO = '" + dictionary["REGISTRO_UNICO_DE_PEDIDO"] + "' and PRODUCTO_SERVICIO = '" + dictionary["PRODUCTO_SERVICIO"] + "'";
                        //crud.Update("Databot", sql_update, "ventas");

                    }
                    respuesta = true;
                }
                try
                {
                    string eorden = "";
                    string sql_select3 = "select ESTATUS_DE_ORDEN from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + dictionary["REGISTRO_UNICO_DE_PEDIDO"] + "' and PRODUCTO_SERVICIO = '" + dictionary["PRODUCTO_SERVICIO"] + "'";
                    DataTable mytable2 = new DataTable();
                    //mytable2 = crud.Select("Databot", sql_select3, "ventas");
                    if (mytable2.Rows.Count > 0)
                    {
                        eorden = mytable2.Rows[0][0].ToString();
                        if (eorden != "Refrendado")
                        {
                            string sql_update2 = "UPDATE `reporte_convenio_gbmpa` SET `ESTATUS_DE_ORDEN`='Refrendado' WHERE REGISTRO_UNICO_DE_PEDIDO = '" + dictionary["REGISTRO_UNICO_DE_PEDIDO"] + "' and PRODUCTO_SERVICIO = '" + dictionary["PRODUCTO_SERVICIO"] + "'";
                            //new CRUD().Update("Databot", sql_update2, "ventas");

                        }
                    }
                }
                catch (Exception)
                {

                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;
        }
        /// <summary>
        /// agregar informacion a la base de datos de la competencia
        /// </summary>
        /// <param name="convenio"></param>
        /// <param name="vendor"></param>
        /// <param name="fecha_desde"></param>
        /// <param name="fecha_hasta"></param>
        /// <param name="entidad"></param>
        /// <param name="registro_unico"></param>
        /// <param name="fecha_registro"></param>
        /// <param name="fecha_publicacion"></param>
        /// <param name="producto"></param>
        /// <param name="cantidad"></param>
        /// <param name="total"></param>
        /// <param name="precio_unitario"></param>
        /// <param name="sub_total"></param>
        /// <param name="linea_producto"></param>
        /// <param name="tipo_producto"></param>
        /// <param name="gbm_participa"></param>
        /// <param name="po_url"></param>
        /// <returns></returns>
        public bool InfoSqlAddCompetencia(string convenio, string vendor, DateTime fecha_desde, DateTime fecha_hasta, string entidad, string registro_unico, DateTime fecha_registro, DateTime fecha_publicacion, string producto, string cantidad, string total, double precio_unitario, string sub_total, string linea_producto, string tipo_producto, string gbm_participa, string po_url)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_insert = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;

            try
            {
                string fecha_registro_end = fecha_registro.ToString("yyyy-MM-dd");
                string fecha_publicacion_end = fecha_publicacion.ToString("yyyy-MM-dd");
                string fecha_desde_end = fecha_desde.ToString("yyyy-MM-dd");
                string fecha_hasta_end = fecha_hasta.ToString("yyyy-MM-dd");

                #region Connection DB
                sql_select = "select * from reporte_convenio_competencia where REGISTRO_UNICO_DE_PEDIDO = '" + registro_unico + "' and PRODUCTO_SERVICIO = '" + producto + "'";
                //mytable = crud.Select("Databot", sql_select, "ventas");
                #endregion


                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    sql_insert = "INSERT INTO `reporte_convenio_competencia` (`CONVENIO`, `VENDOR`, `FECHA_DESDE`, `FECHA_HASTA`, `ENTIDAD`, `REGISTRO_UNICO_DE_PEDIDO`," +
                        " `FECHA_DE_REGISTRO`, `FECHA_DE_PUBLICACION`, `PRODUCTO_SERVICIO`, `CANTIDAD`, `TOTAL_DEL_PRODUCTO`, `PRECIO_UNI`, `SUB_TOTAL_ORDEN`," +
                        " `LINEA_PROD`, `TIPO_PROD`, `GBM_PART`, `LINK_AL_DOCUMENTO`)" +

                        " VALUES ('" + convenio + "','" + vendor + "','" + fecha_desde_end + "','" + fecha_hasta_end + "','" + entidad + "','" + registro_unico + "','" +
                        fecha_registro_end + "','" + fecha_publicacion_end + "','" + producto + "','" + cantidad + "','" + total + "','" + precio_unitario.ToString() + "','" + sub_total + "','" +
                        linea_producto + "','" + tipo_producto + "','" + gbm_participa + "','" + po_url + "')";
                    //crud.Insert("Databot", sql_insert, "ventas");


                    respuesta = true;
                }


            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;
        }
        /// <summary>
        /// Agregar información de una orden en refrendo (aprobada)
        /// </summary>
        /// <param name="convenio"></param>
        /// <param name="entidad"></param>
        /// <param name="producto"></param>
        /// <param name="cantidad"></param>
        /// <param name="precio"></param>
        /// <param name="marca"></param>
        /// <param name="registro_unico"></param>
        /// <param name="fecha_registro"></param>
        /// <param name="fecha_publi"></param>
        /// <param name="fianza"></param>
        /// <param name="estado_gbm"></param>
        /// <param name="status_orden"></param>
        /// <param name="dias_entrega"></param>
        /// <param name="sector"></param>
        /// <returns></returns>
        public bool InfoSqlAddAprobadas(string convenio, string entidad, string producto, string cantidad, string precio, string marca, string registro_unico, string fecha_registro, string fecha_publi, string fianza, string estado_gbm, string status_orden, int dias_entrega, string sector)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_insert = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;
            string fechamax = DateTime.MinValue.Date.ToString("yyyy-MM-dd");
            string fechapublic = DateTime.MinValue.Date.ToString("yyyy-MM-dd");
            string forecast = DateTime.MinValue.Date.ToString("yyyy-MM-dd");

            try
            {
                #region Connection DB
                sql_select = "select * from reporte_convenio_gbmpa where REGISTRO_UNICO_DE_PEDIDO = '" + registro_unico + "' and PRODUCTO_SERVICIO = '" + producto + "'";
                //mytable = crud.Select("Databot", sql_select, "ventas");
                #endregion

                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    //conn.Open();
                    sql_insert = "INSERT INTO `reporte_convenio_gbmpa`(`SECTOR`, `CONVENIO`, `ENTIDAD`, `PRODUCTO_SERVICIO`, `CANTIDAD`, `MARCA`, `TOTAL_DEL_PRODUCTO`, " +
                        "`REGISTRO_UNICO_DE_PEDIDO`, `FECHA_DE_REGISTRO`, `FECHA_DE_PUBLICACION`, `FIANZA_CUMPLIMIENTO`, `ESTADO_GBM`, `ESTATUS_DE_ORDEN`, `DIAS_ENTREGA`, `TIPO_FORECAST`)" +

                        " VALUES ('" + sector + "','" + convenio + "','" + entidad + "','" + producto + "','" + cantidad + "','" + marca + "','" + precio + "','" +
                        registro_unico + "','" + fecha_registro + "','" + fecha_publi + "','" + fianza + "','" + estado_gbm + "','" + status_orden + "','" +
                        dias_entrega.ToString() + "','" + "PIPE" + "')";

                    //crud.Insert("Databot", sql_insert, "ventas");

                    respuesta = true;
                }


            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;
        }

        /// <summary>
        /// Actualizar datos en la base de datos de registros de GBMPA
        /// </summary>
        /// <param name="registro">registro unico de pedido O quote</param>
        /// <param name="dictionary">los campos que desea modificar, importante: los keys deben de ser igual al nombre de la columna</param>
        /// <param name="where">1: registro, 2: quote</param>
        /// <returns></returns>
        public bool UpdateRegister(string registro, Dictionary<string, string> dictionary, int where)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_update = "";
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable update = new DataTable();
            bool respuesta = true;
            try
            {
                sql_update = "UPDATE `reporte_convenio_gbmpa` SET";

                foreach (KeyValuePair<string, string> pair in dictionary)
                {
                    //console.WriteLine("FOREACH KEYVALUEPAIR: {0}, {1}", pair.Key, pair.Value);
                    if (sql_update == "UPDATE `reporte_convenio_gbmpa` SET")
                    {
                        sql_update = sql_update + " `" + pair.Key.ToString() + "` = '" + pair.Value.ToString() + "'";
                    }
                    else
                    {
                        sql_update = sql_update + ", `" + pair.Key.ToString() + "` = '" + pair.Value.ToString() + "'";
                    }
                }
                sql_update = (where == 1) ? sql_update + " WHERE REGISTRO_UNICO_DE_PEDIDO = '" + registro + "'" : sql_update + " WHERE QUOTE = '" + registro + "'";
                //crud.Update("Databot", sql_update, "ventas");

                respuesta = true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }

            return respuesta;
        }
        /// <summary>
        /// Agregar cotización rapida para licitaciones de GBPA
        /// </summary>
        /// <param name="cotizacion">número de cotizacion</param>
        /// <param name="dictionary">los campos que desea insertar, importante: los keys deben de ser igual al nombre de la columna</param>
        /// <returns></returns>
        public bool AddQuote(Dictionary<string, string> dictionary)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_insert = "";
            string sql_insert2 = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            DataTable columnas = new DataTable();
            DataTable update = new DataTable();
            bool respuesta = true;


            try
            {
                #region Connection DB
                sql_select = "select * from reporte_cotizaciones_linea where NUM_COTIZACION = '" + dictionary["NUM_COTIZACION"] + "' and PROD_DESCRIPCION = '" + dictionary["PROD_DESCRIPCION"] + "'";
                //mytable = crud.Select("Databot", sql_select, "ventas");
                #endregion

                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    string sql_columns = "SHOW COLUMNS FROM reporte_cotizaciones_linea";

                    //columnas = crud.Select("Databot", sql_select, "ventas");

                    if (columnas.Rows.Count > 0) //si hay columnas
                    {
                        sql_insert = "INSERT INTO `reporte_cotizaciones_linea`(";
                        sql_insert2 = " VALUES(";
                        //realiza un for por cada una de las columnas de la tabla para asi crear el comando INSERT
                        for (int i = 1; i < columnas.Rows.Count; i++)
                        {

                            string columna = columnas.Rows[i][0].ToString();
                            string valor = dictionary[columna];
                            //solo el primer valor no tien comas
                            if (i == 1)
                            {
                                sql_insert = sql_insert + " `" + columna + "`";
                                sql_insert2 = sql_insert2 + " '" + valor + "'";
                            }
                            else
                            {

                                sql_insert = sql_insert + ", `" + columna + "`";
                                sql_insert2 = sql_insert2 + ", '" + valor + "'";
                            }

                            //}
                        }
                        //agregar el ultimo parentesis a ambas partes
                        sql_insert = sql_insert + ")";
                        sql_insert2 = sql_insert2 + ")";
                        //unir cada parte INSERT INTO `reporte_cotizaciones_linea`(COLUMNAS) VALUES (VALORES)
                        sql_insert = sql_insert + sql_insert2;
                        //crud.Insert("Databot", sql_insert, "ventas");


                    }

                    respuesta = true;

                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;

        }
        #endregion



        #region Nuevo sistema S&S
        /// <summary>
        /// Insertar nuave información en la tabla purchaseOrderCompetition para el portal de Smart & Simple
        /// </summary>
        /// <param name="info">un diccionario cuya llave (key) es la columna de la base de datos (exacta igual) y el valor (value) es el valor a insertar</param>
        /// <returns>true ok, false error</returns>
        public bool insertInfoPurchaseOrder(Dictionary<string, string> info, string table)
        {

            DataTable mytable = new DataTable();
            bool respuesta = true;

            try
            {
                string sql_select = "";
                string sql = "";
                #region Connecion DB     

                #endregion
                sql_select = $"select * from {table} where singleOrderRecord = '" + info["singleOrderRecord"] + "'";
                mytable = crud.Select( sql_select, "panama_bids_db");

                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    sql = $"INSERT INTO `{table}`(";
                    string sql_insert = " VALUES(";
                    foreach (KeyValuePair<string, string> pair in info)
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
                        //}
                    }


                    sql = sql + ")";
                    sql_insert = sql_insert + ")";
                    sql = sql + sql_insert;
                    bool insert = crud.Insert(sql, "panama_bids_db");
                    respuesta = insert;
                }
                else
                {
                    string sqlUpdate, sqlUpd = "";
                    string sql_columns = $"SHOW COLUMNS FROM {table}";
                    DataTable columnas = crud.Select( sql_columns, "panama_bids_db");
                    if (columnas.Rows.Count > 0) //si hay columnas
                    {
                        sqlUpdate = $"UPDATE `{table}` SET";
                        for (int i = 1; i < columnas.Rows.Count; i++)
                        {
                            //modifica solo los valores actuales que cumplan con las siguientes condiciones del if abajo
                            string valor_actual = mytable.Rows[0][i].ToString();
                            if (columnas.Rows[i][0].ToString() != "createdAt" && columnas.Rows[i][0].ToString() != "createdBy" && columnas.Rows[i][0].ToString() != "comment")
                            {

                                if (String.IsNullOrEmpty(valor_actual) || columnas.Rows[i][0].ToString() == "orderStatus" || columnas.Rows[i][0].ToString() == "orderSubtotal" || columnas.Rows[i][0].ToString() == "registrationDate" || columnas.Rows[i][0].ToString() == "maximumDeliveryDate" || columnas.Rows[i][0].ToString() == "publicationDate" || columnas.Rows[i][0].ToString() == "forecastType")
                                {
                                    string columna = columnas.Rows[i][0].ToString();
                                    sqlUpd = sqlUpd + ", `" + columna + "` = '" + info[columna].ToString() + "'";


                                }
                            }
                        }
                        sqlUpd = sqlUpd.Substring(1, sqlUpd.Length - 1);
                        sqlUpdate = sqlUpdate + sqlUpd;
                        sqlUpdate = sqlUpdate + $" WHERE singleOrderRecord = '{info["singleOrderRecord"]}'";
                        bool upd = crud.Update(sqlUpdate, "panama_bids_db");
                        respuesta = upd;
                    }
                    else
                    {
                        respuesta = true;
                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }
            return respuesta;
        }
        /// <summary>
        /// Insertar los productos en la tabla productCompetitivePanama de una orden de compra de la tabla purchaseOrderCompetition 
        /// </summary>
        /// <param name="PoListInfo">una lista de la clase PoProductInfo donde se almacen cada uno de los podcutos de una orden de compra</param>
        /// <returns>true ok, false error</returns>
        public bool insertInfoProductsCompetition(List<PoProductInfo> PoListInfo)
        {
            string sql_select = "";
            string sql = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;
            foreach (PoProductInfo item in PoListInfo)
            {
                try
                {
                    sql_select = $"select * from purchaseOrderCompetitionProducts where singleOrderRecord = '{item.singleOrderRecord}' and product = '{item.product}'";
                    mytable = crud.Select( sql_select, "panama_bids_db");

                    if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                    {
                        sql = "INSERT INTO `purchaseOrderCompetitionProducts`(`singleOrderRecord`, `product`, `amount`, `total`, `unitPrice`, `subtotal`, `active`, `createdBy`)" +
                        $"VALUES ('{item.singleOrderRecord}', '{item.product}', '{item.amount}', '{item.total.Replace(",", "")}', '{item.unitPrice.Replace(",", "")}', '{item.subtotal.Replace(",", "")}', '1', 'databot')";
                        bool insert = crud.Insert(sql, "panama_bids_db");
                        respuesta = insert;
                    }
                    else
                    {

                        respuesta = true;
                    }

                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.ToString());
                    respuesta = false;

                }

            }
            return respuesta;
        }
        /// <summary>
        /// Insertar los productos en la tabla purchaseOrderProduct de una orden de compra de la tabla purchaseOrderCompetition 
        /// </summary>
        /// <param name="PoListInfo">una lista de la clase PoProductInfo donde se almacen cada uno de los podcutos de una orden de compra</param>
        /// <returns>true ok, false error</returns>
        public bool insertInfoProductsMacro(List<PoProductMacro> PoListInfo)
        {
            string sql_select = "";
            string sql = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;
            foreach (PoProductMacro item in PoListInfo)
            {
                try
                {
                    sql_select = $"select * from purchaseOrderProduct where singleOrderRecord = '{item.singleOrderRecord}' and productCode = '{item.productCode}'";
                    mytable = crud.Select( sql_select, "panama_bids_db");

                    if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                    {
                        sql = "INSERT INTO `purchaseOrderProduct` (`singleOrderRecord`, `productCode`, `quantity`, `totalProduct`, `orderType`, `active`, `createdBy`)" +
                        $"VALUES ('{item.singleOrderRecord}', '{item.productCode}', '{item.quantity}', '{item.totalProduct.Replace(",", "")}', '0', '{item.active}', '{item.createdBy}')";
                        crud.Insert(sql, "panama_bids_db");
                    }
                    else
                    {
                        //update
                        string sqlUpdate, sqlUpd = "";
                        string sql_columns = "SHOW COLUMNS FROM purchaseOrderProduct";
                        DataTable columnas = crud.Select( sql_columns, "panama_bids_db");
                        if (columnas.Rows.Count > 0) //si hay columnas
                        {
                            sqlUpdate = "UPDATE `purchaseOrderProduct` SET";
                            for (int i = 1; i < columnas.Rows.Count; i++)
                            {
                                //modifica solo los valores actuales que cumplan con las siguientes condiciones del if abajo
                                string valor_actual = mytable.Rows[0][i].ToString();
                                if (String.IsNullOrEmpty(valor_actual) || columnas.Rows[i][0].ToString() == "quantity" || columnas.Rows[i][0].ToString() == "totalProduct" || columnas.Rows[i][0].ToString() == "registrationDate" || columnas.Rows[i][0].ToString() == "maximumDeliveryDate" || columnas.Rows[i][0].ToString() == "publicationDate" || columnas.Rows[i][0].ToString() == "forecastType")
                                {
                                    Type clase = item.GetType();
                                    string columna = columnas.Rows[i][0].ToString();
                                    string val = clase.GetProperty(columna).GetValue(item).ToString();
                                    sqlUpd = sqlUpd + ", `" + columna + "` = '" + val.ToString() + "'";


                                }
                            }
                            sqlUpd = sqlUpd.Substring(1, sqlUpd.Length - 1);
                            sqlUpdate = sqlUpdate + sqlUpd;
                            sqlUpdate = sqlUpdate + $" WHERE singleOrderRecord = '{item.singleOrderRecord}'";
                            crud.Update(sqlUpdate, "panama_bids_db");
                        }
                        respuesta = true;
                    }
                    respuesta = true;

                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.ToString());
                    respuesta = false;

                }

            }
            return respuesta;
        }
        /// <summary>
        /// Insertar los productos en la tabla purchaseOrderProduct de una orden de compra de la tabla purchaseOrderCompetition 
        /// </summary>
        /// <param name="PoListInfo">una lista de la clase PoProductInfo donde se almacen cada uno de los podcutos de una orden de compra</param>
        /// <returns>true ok, false error</returns>
        public bool insertInfoProductsQuotes(List<productsQuickQuote> PoListInfo)
        {
            string sql_select = "";
            string sql = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;
            foreach (productsQuickQuote item in PoListInfo)
            {
                try
                {
                    sql_select = $"select * from productsQuickQuote where singleOrderRecord = '{item.singleOrderRecord}' and productService = '{item.productService}'";
                    mytable = crud.Select( sql_select, "panama_bids_db");

                    if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                    {
                        sql = "INSERT INTO `productsQuickQuote`(`singleOrderRecord`, `productService`, `clasification`, `ammount`, `unit`, `active`, `createdBy`)" +
                        $"VALUES ('{item.singleOrderRecord}', '{item.productService}', '{item.clasification}', '{item.ammount}', '{item.unit}', '{item.active}', '{item.createdBy}')";
                        crud.Insert(sql, "panama_bids_db");
                    }
                    else
                    {
                        //update
                        string sqlUpdate, sqlUpd = "";
                        string sql_columns = "SHOW COLUMNS FROM productsQuickQuote";
                        DataTable columnas = crud.Select( sql_columns, "panama_bids_db");
                        if (columnas.Rows.Count > 0) //si hay columnas
                        {
                            sqlUpdate = "UPDATE `productsQuickQuote` SET";
                            for (int i = 1; i < columnas.Rows.Count; i++)
                            {
                                //modifica solo los valores actuales que cumplan con las siguientes condiciones del if abajo
                                string valor_actual = mytable.Rows[0][i].ToString();
                                if (String.IsNullOrEmpty(valor_actual) || columnas.Rows[i][0].ToString() == "productService" || columnas.Rows[i][0].ToString() == "clasification" || columnas.Rows[i][0].ToString() == "ammount" || columnas.Rows[i][0].ToString() == "unit")
                                {
                                    Type clase = item.GetType();
                                    string columna = columnas.Rows[i][0].ToString();
                                    string val = clase.GetProperty(columna).GetValue(item).ToString();
                                    sqlUpd = sqlUpd + ", `" + columna + "` = '" + val.ToString() + "'";


                                }
                            }
                            sqlUpd = sqlUpd.Substring(1, sqlUpd.Length - 1);
                            sqlUpdate = sqlUpdate + sqlUpd;
                            sqlUpdate = sqlUpdate + $" WHERE singleOrderRecord = '{item.singleOrderRecord}'";
                            crud.Update(sqlUpdate, "panama_bids_db");
                        }
                        respuesta = true;
                    }
                    respuesta = true;

                }
                catch (Exception ex)
                {
                    console.WriteLine(ex.ToString());
                    respuesta = false;

                }

            }
            return respuesta;
        }
        public bool insertInfoApproved(string convenio, string entidad, string registro_unico, string fecha_registro, int dias_entrega, string sector, float subTotal)
        {
            string sql_select = "";
            string sql = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;

            try
            {
                sql_select = $"select * from purchaseOrderMacro where singleOrderRecord = '{registro_unico}'";
                mytable = crud.Select( sql_select, "panama_bids_db");
                if (mytable.Rows.Count <= 0) //insertar si es que no existe para evitar duplicidad
                {
                    sql = "INSERT INTO `purchaseOrderMacro`(`singleOrderRecord`, `sector`, `agreement`, `entity`, `orderSubtotal`, `registrationDate`, `performanceBond`, `deliveryDay`, `forecastType`, `gbmStatus`, `orderStatus`, `active`, `createdBy`)" +
                    $"VALUES ('{registro_unico}', '{sector}', '{convenio}', '{entidad}', '{subTotal}', '{fecha_registro}', '1', '{dias_entrega.ToString()}', '2', '3', '3', '1', 'databot')";

                    crud.Insert(sql, "panama_bids_db");
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
        public bool OppUpdate(string registro, string opp)
        {
            string sql_select = "";
            string sql = "";
            DataTable mytable = new DataTable();
            bool respuesta = true;

            try
            {

                sql = $"UPDATE `purchaseOrderMacro` SET `oportunity`= '{opp}' WHERE singleOrderRecord = '{registro}'";
                crud.Update(sql, "panama_bids_db");
                respuesta = true;

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                respuesta = false;

            }


            return respuesta;
        }
        public string oppExist(string registro)
        {
            string opp_actual = "";
            DataTable mytable = new DataTable();
            try
            {
                string sql = $"SELECT oportunity FROM `purchaseOrderMacro` WHERE singleOrderRecord = '{registro}'";
                mytable = crud.Select( sql, "panama_bids_db");

                if (mytable.Rows.Count > 0)
                {
                    opp_actual = mytable.Rows[0][0].ToString(); //opp

                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());
                opp_actual = "";

            }

            return opp_actual;
        }
        /// <summary>
        /// Extraer una tabla productsCompetition para buscar los codigos de producto de las ordenes de compra de la competencia
        /// sustituir por el nuevo metodo de CRUD
        /// </summary>
        /// <returns></returns>
        public DataTable productType()
        {
            string sql_select = "";
            DataTable mytable = new DataTable();
            string[] info = new string[3];
            try
            {
                sql_select = "select * from productsCompetition";
                mytable = crud.Select( sql_select, "panama_bids_db");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }
            return mytable;
        }
        public DataTable columnsPoMacro()
        {
            string sql_select = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "SHOW FULL COLUMNS FROM `purchaseOrderMacro`";
                mytable = crud.Select( sql_select, "panama_bids_db");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }
            return mytable;
        }
        public DataTable entitiesInfo()
        {
            string sql_select = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "SELECT entities.*, sector.sector as sectorText FROM `entities` INNER JOIN sector ON sector.sectorCode = entities.sector";
                mytable = crud.Select( sql_select, "panama_bids_db");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }
            return mytable;
        }
        public DataTable productsInfo()
        {
            string sql_select = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "SELECT products.*, brand.brand as 'brandText' FROM `products` INNER JOIN brand ON products.brand = brand.id";
                mytable = crud.Select( sql_select, "panama_bids_db");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }
            return mytable;
        }
        public DateTime[] getholidays()
        {
            DateTime[] feriados = new DateTime[1];
            string sql_select = "";
            DataTable mytable = new DataTable();

            try
            {
                sql_select = "SELECT * FROM `panamaHolidays`";
                mytable = crud.Select( sql_select, "panama_bids_db");
                if (mytable.Rows.Count > 0)
                {
                    int cont = 0;
                    foreach (DataRow item in mytable.Rows)
                    {
                        DateTime fecha = DateTime.Parse(item["date"].ToString());
                        DateTime nfecha = new DateTime(DateTime.Now.Year, fecha.Month, fecha.Day);
                        feriados[cont] = nfecha; //fechas feriadas
                        Array.Resize(ref feriados, feriados.Length + 1);
                        cont++;
                    }
                }
                else
                {
                    feriados[0] = DateTime.MinValue;

                }
            }
            catch (Exception ex)
            {
                //marca[0] = "No se encontró este producto en la lista";
            }
            Array.Resize(ref feriados, feriados.Length - 1);
            return feriados;
        }
        public string getEmail(string filtro)
        {
            string sql_select = "";
            string jemail = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = $"SELECT jemail FROM `emailAddress` WHERE category = '{filtro}' AND active = 1";
                mytable = crud.Select( sql_select, "panama_bids_db");
                jemail = mytable.Rows[0][0].ToString();
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }


            return jemail;
        }
        public string getUserEmail(string user)
        {
            string sql_select = "";
            string jemail = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = $"SELECT `UserID` FROM `digital_sign` WHERE user = '{user}'";
                mytable = crud.Select( sql_select, "MIS");
                jemail = "AA" + mytable.Rows[0][0].ToString().PadLeft(8, '0');

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
                jemail = "AA70000124";
            }


            return jemail;
        }
        public bool insertFile(string fileName, string singleOrderRecord, string identifier)
        {
            string sqlInsert = "";
            string jemail = "";
            DataTable mytable = new DataTable();
            string folder = (identifier == "quickQuotes") ? "QuickQuotes" : "Agreement";
            try
            {
                string user = "";
                if (ssMandante == "QAS")
                {
                    user = cred.QA_SS_APP_SERVER_USER;
                }
                else if (ssMandante == "PRD")
                {
                    user = cred.PRD_SS_APP_SERVER_USER;
                }
                sqlInsert = $"INSERT INTO `uploadFiles`(`name`, `singleOrderRecord`, `user`, `codification`, `type`, `path`, `identifier`, `active`, `createdBy`)" +
                    $"VALUES ('{fileName}', '{singleOrderRecord}', 'databot', '7bit', 'application/pdf', '/home/{user}/projects/smartsimple/gbm-hub-api/src/assets/files/PanamaBids/{folder}/Request #{singleOrderRecord}/{fileName}', '{identifier}', 1, 'databot')";
                crud.Insert(sqlInsert, "panama_bids_db");
                return true;
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
            }
            return false;

        }
        public List<string> KeyWords(string categoria)
        {
            List<string> stopWords = new List<string>();
            DataTable mytable = new DataTable();
            try
            {
                string sqlSelect = $"SELECT * FROM `keyWords` WHERE active = 1";
                mytable = crud.Select( sqlSelect, "panama_bids_db");

                if (mytable.Rows.Count > 0)
                {
                    foreach (DataRow item in mytable.Rows)
                    {
                        string key = item["key"].ToString().ToLower();
                        key = val.RemoveSpecialChars(key, 1);
                        stopWords.Add(key); //palabra clave
                    }
                }
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.ToString());

            }

            return stopWords;
        }
        public DataTable AllQuickQuotes()
        {
            string sql_select = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "SELECT singleOrderRecord FROM `quickQuoteReport`";
                mytable = crud.Select( sql_select, "panama_bids_db");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);

            }
            return mytable;
        }

        #endregion





    }
}
