using DataBotV5.Data.Database;
using System.Data;

namespace DataBotV5.Data.Projects.PositionLetter
{
    class PositionLetterSQL
    {

        public DataTable Processes()
        {
            DataTable mytable = new DataTable();
            string sql = "SELECT * FROM condiciones_posicion WHERE ESTADO = 'EN PROCESO'";
            //mytable = new CRUD().Select("Databot", sql, "automation");

            return mytable;
        }
    }
}
