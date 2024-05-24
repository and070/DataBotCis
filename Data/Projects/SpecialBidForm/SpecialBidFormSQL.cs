using DataBotV5.Data.Database;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace DataBotV5.Data.Projects.SpecialBidForm
{
    class SpecialBidFormSQL
    {
        CRUD crud = new CRUD();

        string sb_number = "";
        public string id_carpeta = "";
        public string projectname;
        public string ibm_country = "";
        public string useopp = "";
        public string opp = "";
        public string useprevbid = "";
        public string prevbid = "";
        public string priceupd = "";
        public string justi = "";
        public string customer = "";
        public string brand = "";
        public string justi2 = "";
        public string addquest = "";
        public string bpjusti = "";
        public string swma = "";
        public string renew = "";
        public string totalprice = "";
        public string customerprice = "";
        public string totalright = "";
        public string totalright2 = "";
        public string customerright = "";
        public string customerright2 = "";
        public string sb_id_gestion = "";
        public string usuario = "";
        public string alerta_text = "";
        public string NewRequestSB()
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string id_carpeta = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB  
                sql_select = "select * from solicitudes_sb where ESTADO = 1";
                //mytable = crud.Select("Databot", sql_select, "pricing");
                #endregion


                if (mytable.Rows.Count > 0)
                {
                    id_carpeta = mytable.Rows[0][18].ToString();
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
        public void ExtractSBInfo(string folder)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            DataTable mytable = new DataTable();
            DataTable mytable2 = new DataTable();
            try
            {
                #region Connection DB
                sql_select = "select * from solicitudes_sb where ESTADO = 1 and CARPETA = " + folder;
                //mytable = crud.Select("Databot", sql_select, "pricing");
                #endregion

                if (mytable.Rows.Count > 0)
                {

                    sb_id_gestion = mytable.Rows[0][0].ToString().Trim();
                    ibm_country = mytable.Rows[0][1].ToString().Trim();
                    projectname = mytable.Rows[0][2].ToString().Trim();
                    customer = mytable.Rows[0][3].ToString().Trim();
                    brand = mytable.Rows[0][4].ToString().Trim();
                    useopp = mytable.Rows[0][5].ToString().Trim();
                    opp = mytable.Rows[0][6].ToString().Trim();
                    useprevbid = mytable.Rows[0][7].ToString().Trim();
                    prevbid = mytable.Rows[0][8].ToString().Trim();
                    priceupd = mytable.Rows[0][9].ToString().Trim();
                    justi = mytable.Rows[0][10].ToString().Trim();
                    justi2 = mytable.Rows[0][11].ToString().Trim();
                    bpjusti = mytable.Rows[0][12].ToString().Trim();
                    addquest = mytable.Rows[0][13].ToString().Trim();
                    swma = mytable.Rows[0][14].ToString().Trim();
                    renew = mytable.Rows[0][15].ToString().Trim();
                    totalprice = mytable.Rows[0][16].ToString().Trim();
                    customerprice = mytable.Rows[0][17].ToString().Trim();
                    usuario = mytable.Rows[0][19].ToString().Trim();
                }

            }
            catch (Exception ex)
            {
            }
        }
        public void DeleteFolderWin(string route)
        {
            DirectoryInfo dir = new DirectoryInfo(route);
            dir.Delete(true);
            //ultimo cambio 23_08_2019
        }
        public void CompleteRequestSB(string folder)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB  
                sql_select = "select * from solicitudes_sb where CARPETA = " + folder;
                //mytable = crud.Select("Databot", sql_select, "pricing");
                #endregion

                if (mytable.Rows.Count > 0)
                {
                    sql_update = "Update solicitudes_sb set ESTADO = 0 where CARPETA = '" + folder + "'";
                    //crud.Update("Databot", sql_update, "pricing");
                }
                else
                {

                }


            }
            catch (Exception ex)
            {

            }
        }
        public void DeleteRequestSB(string folder)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select = "";
            string sql_update = "";
            DataTable mytable = new DataTable();
            try
            {
                #region Connection DB
                sql_select = "select * from solicitudes_sb where CARPETA = " + folder;
                //mytable = crud.Select("Databot", sql_select, "pricing");
                #endregion


                if (mytable.Rows.Count > 0)
                {
                    sql_update = "Update solicitudes_sb set ESTADO = 2 where CARPETA = '" + folder + "'";
                    //crud.Update("Databot", sql_update, "pricing");

                }
                else
                {

                }

            }
            catch (Exception ex)
            {

            }
        }

    }
}
