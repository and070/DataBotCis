using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using DataBotV5.Data.Root;
using MySql.Data.MySqlClient;
using DataBotV5.Data.Database;

namespace DataBotV5.Data.Projects.HcmHiring
{
    /// <summary>
    /// Clase Data encargada del SQL de la Pb10.
    /// </summary>
    class SqlPb10
    {
        Credentials.Credentials cred = new Credentials.Credentials();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();
        public Dictionary<string, string> nueva_solicitud()
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            Dictionary<string, string> solicitud_info = new Dictionary<string, string>();
            string respuesta = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "SELECT * FROM solicitudes_pb10 WHERE estado = 'EN PROCESO' OR estado = 'PENDIENTE'";
                //mytable = crud.Select("Databot", sql_select, "human_capital");


                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < 21; i++)
                    {
                        string columna = mytable.Columns[i].ToString();
                        string valor = mytable.Rows[0][i].ToString();
                        solicitud_info[columna] = valor;
                    }
                    respuesta = "OK";
                }


            }
            catch (Exception ex)
            {
                respuesta = "ERROR";
            }
            return solicitud_info;
        }
        public Dictionary<string, string> newRequest()
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            Dictionary<string, string> solicitud_info = new Dictionary<string, string>();
            string respuesta = "";
            DataTable mytable = new DataTable();
            DataTable columns = new DataTable();
            try
            {
                sql_select = "SELECT * FROM PB10Requests WHERE status = 'EP'";
                mytable = crud.Select( sql_select, "hcm_hiring_db");
                if (mytable.Rows.Count > 0)
                {
                    columns = crud.Select( "SHOW FULL COLUMNS FROM `PB10Requests`", "hcm_hiring_db");
                    for (int i = 0; i < columns.Rows.Count; i++) //Es un For para las columnas de la tabla PB10
                    {
                        string columna = mytable.Columns[i].ToString();
                        string valor = mytable.Rows[0][i].ToString(); //Solo se toma el valor de la primer fila de la tabla.
                        solicitud_info[columna] = valor;
                    }
                    respuesta = "OK";
                }




            }
            catch (Exception ex)
            {
                respuesta = "ERROR";
            }
            return solicitud_info;
        }
        /// <summary>
        /// Método para cambiar el estado de solicitudes de la PB10 a nivel de Databot.
        /// </summary>
        /// <param name="id_gestion"></param>
        /// <param name="respuesta"></param>
        /// <param name="estado"></param>
        /// <param name="fecha_creado"></param>
        /// <returns></returns>
        public bool ChangeStatePb10Databot(string id_gestion, string respuesta, string estado, DateTime fecha_creado)
        {
            respuesta = respuesta.Replace("'", "");
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            string sql_update2 = "";
            string fechaf = fecha_creado.ToString("yyyy-MM-dd HH:mm:ss");
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "select * from solicitudes_pb10 where id_gestion = " + id_gestion;
                //mytable = crud.Select("Databot", sql_select, "human_capital");
                if (mytable.Rows.Count > 0)
                {
                    sql_update = "UPDATE solicitudes_pb10 SET estado = '" + estado + "',respuesta = '" + respuesta + "',ts_finalizacion = '" + fechaf + "' where id_gestion = " + id_gestion;
                    //crud.Update("Databot", sql_update, "human_capital");
                }
                else
                {

                }

                //myadapter.Dispose();

            }
            catch (Exception ex)
            {

            }
            return false;
        }

        /// <summary>
        /// Método para cambiar el estado de solicitudes de la PB10 a nivel de SmartAnSimple.
        /// </summary>
        /// <param name="id_gestion"></param>
        /// <param name="respuesta"></param>
        /// <param name="estado"></param>
        /// <param name="fecha_creado"></param>
        /// <returns></returns>
        public bool ChangeStatePb10SmartAndSimple(string id, string response, string status, DateTime fecha_creado)
        {
            response = response.Replace("'", "");
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            string sql_update2 = "";
            string finishDate = fecha_creado.ToString("yyyy-MM-dd HH:mm:ss");
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "select * from PB10Requests where id= " + id;
                mytable = crud.Select( sql_select, "hcm_hiring_db");
                if (mytable.Rows.Count > 0)
                {
                    sql_update = "UPDATE PB10Requests SET status = '" + status + "', response = '" + response + "',finishAt = '" + finishDate + "' where id = " + id;

                    crud.Update(sql_update, "hcm_hiring_db");
                }
                else
                {



                }
            }
            catch (Exception ex)
            {



            }
            return false;
        }
        public DataTable getCandidateInfo(string table, string id)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            DataTable mytable = new DataTable();
            try
            {
                if (table == "CandidateEducation")
                {
                    sql_select = $@"SELECT {table}.*,
                                    AcademicDegree.academicDegreeCode
                                    FROM {table}
                                    INNER JOIN AcademicDegree ON {table}.academicDegree = AcademicDegree.id
                                    WHERE {table}.id = '{id}'";
                }
                else
                {
                    sql_select = $"SELECT * FROM {table} WHERE id = '{id}'";
                }
                mytable = crud.Select( sql_select, "hcm_hiring_db");



            }
            catch (Exception ex)
            {

            }
            return mytable;



        }
    }
}
