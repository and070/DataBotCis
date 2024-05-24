using DataBotV5.Data.Database;
using System;
using System.Data;
using System.Data.SqlClient;

namespace DataBotV5.Data.Projects.OppSAM
{
    class OppSAMSQL
    {
        CRUD crud = new CRUD();
        public string[] GetManager(string country, string territory)
        {
            SqlConnection myConn = new SqlConnection();
            string sql_select;
            DataTable mytable = new DataTable();
            string[] usuarios = new string[2];
            try
            {
                #region Connection DB
                sql_select = "select * from territoryNotification where territoryCountry = '" + country + territory.TrimStart('0') + "'";
                mytable = crud.Select(sql_select, "opp_sam_db");
                #endregion

                string[] arr = new string[mytable.Columns.Count - 1];
                if (mytable.Rows.Count > 1)
                {
                    usuarios[0] = "Error: Se encontro mas de una fila en un solo pais";
                    return usuarios;
                }
                else
                {
                    arr[0] = mytable.Rows[0][1].ToString();
                    arr[1] = mytable.Rows[0][2].ToString();
                }
                usuarios = arr;
            }
            catch (Exception ex)
            {
                usuarios[0] = "Error: " + ex.Message.ToString();
                return usuarios;
            }
            return usuarios;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="opp"></param>
        public void SaveRequest(string message, string opp)
        {
            string sql_insert;
            sql_insert = "Insert into mqRequests (message, opp)" + "values('" + message + "', '" + opp + "')";
            crud.Insert(sql_insert, "opp_sam_db");

        }
        /// <summary>
        /// 
        /// </summary>
        public void TurnOffRobot()
        {
            string sql_update;
            sql_update = "UPDATE `orchestrator` SET active = '0' where class = 'OPP_SAM'";
            crud.Update(sql_update, "databot_db");

        }
    }
}
