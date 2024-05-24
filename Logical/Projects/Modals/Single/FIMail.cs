using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using DataBotV5.Data.Database;
using DataBotV5.Data.SAP;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de FIEmail.
    /// </summary>
    class FIMail
    {
        string sapSys = "ERP";
        SapVariants sap = new SapVariants();
        public List<string> Correos { get; set; }
        public List<string> Copias { get; set; }
        public FIMail(string po)
        {

            string correos_db = "";
            string copias_db = "";
            string pais = "";


            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["PO"] = po;

            IRfcFunction func = sap.ExecuteRFC(sapSys, "ZCATS_INFO_PO_CTY", parameters);

            pais = func.GetValue("CTY").ToString();

            DataTable mytable = new DataTable();
            string sql = "SELECT * FROM freelance_mf WHERE PAIS = '" + pais + "'";
            //mytable = new CRUD().Select("Databot", sql, "automation");

            correos_db = mytable.Rows[0][2].ToString();
            copias_db = mytable.Rows[0][3].ToString();


            if (correos_db != "")
            {
                if (correos_db.IndexOf(',') != -1)
                {
                    //mas de un elemento
                    List<string> lista = new List<string>();
                    //mas de un elemento
                    string[] cp = correos_db.Split(',', (char)StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < cp.Length; i++)
                    {
                        lista.Add(cp[i]);
                    }
                    Correos = lista;
                }
                else
                {
                    List<string> lista = new List<string>();
                    lista.Add(correos_db);
                    Correos = lista;
                    //solo un elemento
                }
            }
            if (copias_db != "")
            {
                if (copias_db.IndexOf(',') != -1)
                {
                    List<string> lista = new List<string>();
                    //mas de un elemento
                    string[] cp = copias_db.Split(',', (char)StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < cp.Length; i++)
                    {
                        lista.Add(cp[i]);
                    }
                    Copias = lista;
                }
                else
                {
                    List<string> lista = new List<string>();
                    lista.Add(copias_db);
                    Copias = lista;
                    //solo un elemento
                }
            }
        }
    }
}
