using SAP.Middleware.Connector;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using System.Collections.Generic;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de FLWmployee.
    /// </summary>
    class FLEmployee
    {
        public string Nombre { set; get; }
        public string Correo { set; get; }
        public string IdEmpleado { set; get; }
        public string Usuario { set; get; }
        string sapSys = "ERP";

        public FLEmployee(string usuario)
        {
            ConsoleFormat console = new ConsoleFormat();

            Usuario = usuario;
            console.WriteLine(" Obteniendo datos de empleado.");
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["USUARIO"] = usuario;

            IRfcFunction func = new SapVariants().ExecuteRFC(sapSys, "ZFD_GET_USER_DETAILS", parameters, 260);



            Nombre = func.GetValue("NOMBRE").ToString();
            Correo = func.GetValue("EMAIL").ToString();
            IdEmpleado = "AA" + func.GetValue("IDCOLABC").ToString();

        }
    }
}
