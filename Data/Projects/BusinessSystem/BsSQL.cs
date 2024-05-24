using System;
using System.Data.SqlClient;
using System.Data;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using System.Collections.Generic;
using DataBotV5.App.Global;

namespace DataBotV5.Data.Projects.BusinessSystem
{
    /// <summary>
    /// Clase Data encargada de 
    /// </summary>
    class BsSQL
    {
        Credentials.Credentials cred = new Credentials.Credentials();
        Rooting root = new Rooting();
        CRUD crud = new CRUD();
        ConsoleFormat console = new ConsoleFormat();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="table">la tabla donde se va a leer emaila_poe or emailb_poe</param>
        /// <param name="ibmPo">la PO que viene en el subject</param>
        /// <param name="ibmPdf">el nombre del PDF</param>
        public void SqlIBMPoe(string table, string ibmPo, string ibmPdf)
        {
            SqlConnection myConn = new SqlConnection();
            DataTable mytable = new DataTable();
            DataTable find_mytable = new DataTable();
            object[] find_result;
            string f_copy = "";
            string sql_select;
            string sql_insert;
            string sql_update;

            try
            {
                if (table == "emailAPoe")
                {
                    sql_select = "select * from " + table + " where purchaseOrder = '" + ibmPo + "'";
                    mytable = crud.Select(sql_select, "business_system_db");


                    if (mytable.Rows.Count > 0)
                    {
                        sql_update = "Update " + table + " set `pdfName` = '" + ibmPdf + "' where purchaseOrder = '" + ibmPo + "'";
                        crud.Update(sql_update, "business_system_db");

                    }
                    else
                    {
                        sql_insert = "Insert into emailAPoe (purchaseOrder, pdfName, createdBy) values ('" + ibmPo + "','" + ibmPdf + "', 'BS')";
                        crud.Insert(sql_insert, "business_system_db");
                    }

                }
                else if (table == "emailBPoe")
                {

                    sql_select = "select * from " + table + " where purchaseOrder = '" + ibmPo + "'";
                    mytable = crud.Select(sql_select, "business_system_db");

                    if (mytable.Rows.Count > 0)
                    {
                        sql_update = "Update " + table + " set `pdfName` = '" + ibmPdf + "' where purchaseOrder = '" + ibmPo + "'";
                        crud.Update(sql_update, "business_system_db");
                    }
                    else
                    {
                        sql_insert = $"Insert into {table} (purchaseOrder, pdfName, createdBy) values('" + ibmPo + "','" + ibmPdf + "', 'BS')";
                        crud.Insert(sql_insert, "business_system_db");
                    }

                    //busca en la tabla de Correos A la PO
                    sql_select = "SELECT `pdfName` FROM emailAPoe WHERE purchaseOrder = '" + ibmPo + "'";
                    find_mytable = crud.Select(sql_select, "business_system_db");

                    if (find_mytable.Rows.Count > 0)
                    {
                        find_result = find_mytable.Rows[0].ItemArray;
                        root.ibm_pdf2 = find_result[0].ToString();
                    }
                    else
                    {
                        root.ibm_pdf2 = "no file";
                    }

                }


            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                root.ibm_pdf2 = "no file";
            }


        }

        /// <summary>
        ///  extrae de la tabla email_address en la base de datos "reglas"
        /// </summary>
        /// param name="indice" indica la fila de la tabla de donde va a tomar la info: ejemplo 1 significa que toma la fila cuyo ID sea 1
        /// <returns> los sender y sus copias </returns>
        public string[] EmailAddress(int index)
        {

            string[] copies = null;
            try
            {
                root.f_sender = "";
                List<string> copiesList = new List<string>();
                #region Connection DB
                string sql_select = "select * from emailAddress where id = " + index;
                DataTable mytable = crud.Select( sql_select, "document_system");
                #endregion

                if (mytable.Rows.Count > 0)
                {

                    foreach (DataRow row in mytable.Rows)
                    {
                        root.f_sender = row["sender"].ToString();
                        foreach (DataColumn col in row.Table.Columns)
                        {
                            if (col.ColumnName.Contains("copy"))
                            {
                                string value = row[col.ColumnName].ToString();
                                if (value != "")
                                {
                                    copiesList.Add(value);

                                }
                            }
                        }
                    }

                    copies = copiesList.ToArray();
                }

            }
            catch (Exception ex)
            { console.WriteLine(ex.Message); }

            return copies;

        }

        public string NewRequestDMS()
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string id_carpeta = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB     
                sql_select = "select * from solicitudes_dms where ESTADO_ROBOT = 1";
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    id_carpeta = mytable.Rows[0]["CARPETA"].ToString();
                    return id_carpeta;
                }
                else
                {
                    id_carpeta = string.Empty;
                    return id_carpeta;
                }
            }
            catch (Exception ex)
            {
                id_carpeta = string.Empty;
                return id_carpeta;
            }

        }
        public DataTable ExtractDmsInfo(string folder)
        {
            //root.bs_usuario = "";
            //root.bs_po_pais = "";
            //root.bs_vendor = "";
            //root.bs_id_gestion = "";
            //root.bs_tipo_gestion = "";
            //root.bs_comentario = "";
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "select * from solicitudes_dms where ESTADO_ROBOT = 1 and CARPETA = " + folder;
                #region Connection DB 
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion
                return mytable;
            }
            catch (Exception ex)
            {
                return mytable;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="folder">carpeta que se va cambiar</param>
        /// <param name="state">0 para completado, 2 para error</param>
        public void ModifyStatusRobotDms(string folder, int state)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB 
                sql_select = "select * from solicitudes_dms where CARPETA = " + folder;
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    sql_update = "Update solicitudes_dms set ESTADO_ROBOT = " + state + " where CARPETA = '" + folder + "'";
                    //crud.Update("Databot", sql_update, "business_system_db");

                }
                else
                {

                }


            }
            catch (Exception ex)
            {

            }
        }
        public void AddReasonsDMS(string doc, string reasonCancell, string commentBs)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            try
            {
                sql_select = "select * from solicitudes_dms where ID_PO_O_SO = " + doc;
                //mytable = crud.Select("Databot", sql_select, "business_system_db");

                if (mytable.Rows.Count > 0)
                {
                    sql_update = "Update solicitudes_dms set razon_cancelado = '" + reasonCancell + "',`comentario_bs`= '" + commentBs + "' where ID_PO_O_SO = '" + doc + "'";
                    //crud.Update("Databot", sql_update, "business_system_db");
                }
                else
                {

                }

            }
            catch (Exception ex)
            {

            }
        }
        public string ExtractCountryDMS(string docNum)
        {
            root.bs_usuario = "";
            root.bs_tipo_gestion = "";
            root.bs_status = "";
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string pais = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                sql_select = "select * from solicitudes_dms where ID_PO_O_SO = " + docNum;
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    root.bs_tipo_gestion = mytable.Rows[0]["PO_O_SO"].ToString();
                    root.bs_usuario = mytable.Rows[0]["USUARIO"].ToString();
                    pais = mytable.Rows[0]["PAIS"].ToString();
                    root.bs_status = mytable.Rows[0]["ESTADO"].ToString();
                }
            }
            catch (Exception ex)
            {
                pais = "";
            }

            return pais;
        }
        public string ExtractCountryUser(string user)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string pais = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB
                sql_select = "select * from usuarios_dms where USUARIO = '" + user + "'";
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    pais = mytable.Rows[0]["PAIS"].ToString();
                }
            }
            catch (Exception ex)
            {
                pais = "";
            }

            return pais;
        }
        public void ChangeStatus(string docNum, string status)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                sql_select = "select * from solicitudes_dms where ID_PO_O_SO = " + docNum;
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    if (status == "Completed")
                    {
                        sql_update = "Update solicitudes_dms set ESTADO = '" + status + "',`FECHA_FINAL`= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where ID_PO_O_SO = '" + docNum + "'";
                        //crud.Update("Databot", sql_update, "business_system_db");
                    }
                    else if (status == "Rejected" || status == "Cancelled")
                    {
                        sql_update = "Update solicitudes_dms set ESTADO = '" + status + "',`FECHA_INTER`= '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where ID_PO_O_SO = '" + docNum + "'";
                        //crud.Update("Databot", sql_update, "business_system_db");
                    }
                    else
                    {
                        sql_update = "Update solicitudes_dms set ESTADO = '" + status + "' where ID_PO_O_SO = '" + docNum + "'";
                        //crud.Update("Databot", sql_update, "business_system_db");
                    }
                }
                else
                {

                }

            }
            catch (Exception ex)
            {

            }
        }
        public void DeleteRequest(string idFolder)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_delete = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                sql_select = "select * from solicitudes_dms where CARPETA = " + idFolder;
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    sql_delete = "DELETE FROM solicitudes_dms WHERE CARPETA = '" + idFolder + "'";
                    //crud.Delete("Databot", sql_delete, "business_system_db");
                }
                else
                {

                }
            }
            catch (Exception ex)
            {

            }
        }
        public string[] UserResponse(string userCountry, string usuario)
        {

            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            int cont_adj = 0;
            //string usuarios = "";
            string[] usuarios = new string[1];
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB
                sql_select = "select * from usuarios_dms where PAIS = '" + userCountry + "' and USER_TYPE = 'USUARIO'";
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i <= mytable.Rows.Count - 1; i++)
                    {
                        string user = mytable.Rows[i][1].ToString();
                        if (user != user)
                        {
                            usuarios[cont_adj] = mytable.Rows[i][1].ToString();
                            cont_adj++;
                            Array.Resize(ref usuarios, usuarios.Length + 1);
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                usuarios = null;
            }

            return usuarios;
        }
        public bool upDate(string docNum, string endFinal)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            bool ok = true;
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB   
                sql_select = "select * from solicitudes_dms where ID_PO_O_SO = " + docNum;
                //mytable = crud.Select("Databot", sql_select, "business_system_db");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    sql_update = "Update solicitudes_dms set `FECHA_FINAL`= '" + endFinal + "' where ID_PO_O_SO = '" + docNum + "'";
                    //crud.Update("Databot", sql_update, "business_system_db");
                }
                else
                {
                    ok = false;
                }

            }
            catch (Exception ex)
            {
                ok = false;
            }
            return ok;
        }


    }
}
