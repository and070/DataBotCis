using DataBotV5.Data.Database;
using DataBotV5.Logical.Projects.Modals.Single;
using System.Data;

namespace DataBotV5.Data.Projects.Autopp
{
    /// <summary>
    /// Clase Data encargada de obtener las oportunidades de empleados vía email.
    /// </summary>
    class EmployeeOpp
    {
        public string GetUserOppEmail(string opp)
        {
            string resultado = "";
            string correo = "";
            DataTable mytable = new DataTable();
            string sql = "SELECT EMPLEADO FROM gestiones WHERE OPPID = '"+opp+"'";
            //mytable = new CRUD().Select("Databot", sql, "fabrica_de_ofertas");

            if (mytable.Rows.Count > 0)
            {
                resultado = mytable.Rows[0][0].ToString();
            }

            if (resultado != "")
            {
                CCEmployee empleado = new CCEmployee(resultado);
                correo = empleado.Correo;
            }

            return correo;
        }
    }
}
