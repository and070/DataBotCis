using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.SAP;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de CCEmployee.
    /// </summary>
    class CCEmployee
    {
        #region 
        string syst = "ERP";
        SapVariants sap = new SapVariants();
        #endregion
        public string Nombre     { set; get; }
        public string Correo     { set; get; }
        public string IdEmpleado { set; get; }
        public string Usuario    { set; get; }
        
        public CCEmployee(string usuario)
        {

            Credentials cred = new Credentials();
            ConsoleFormat console = new ConsoleFormat();

            Usuario = usuario;
            Dictionary<string, string> parametros = new Dictionary<string, string>();
            parametros["USUARIO"] = usuario;

            IRfcFunction func = sap.ExecuteRFC(syst, "ZFD_GET_USER_DETAILS", parametros);


            Nombre = func.GetValue("NOMBRE").ToString();
            Correo = func.GetValue("EMAIL").ToString();
            IdEmpleado = "AA" + func.GetValue("IDCOLABC").ToString();
        }
    }
}
