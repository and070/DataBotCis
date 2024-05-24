using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical encargada de requerimientos de Business Process Management.
    /// </summary>
    class BPMRequirement
    {
        public List<string> requerimientos = new List<string>();
        public int n_req { set; get; }
        public string proveedor { set; get; }
        public string producto { set; get; }
        public string requerimiento { set; get; }
        public string cantidad { set; get; }
        public string integracion { set; get; }
        public string comentarios { set; get; }
        /// <summary>
        /// Inicializa y asigna los requerimientos de BPM para que estos puedan ser parseados y contabilizados.
        /// </summary>
        /// <param name="requirement1">Requerimiento de la gestión BD 1</param>
        /// <param name="requirement2">Requerimiento de la gestión BD 2</param>
        /// <param name="requirement3">Requerimiento de la gestión BD 3</param>
        /// <param name="requirement4">Requerimiento de la gestión BD 4</param>
        public BPMRequirement(string requirement1, string requirement2, string requirement3, string requirement4)
        {
            if (requirement1 != null && requirement1 != " " && requirement1 != "")
            {
                requerimientos.Add(requirement1);
            }
            if (requirement2 != null && requirement2 != " " && requirement2 != "")
            {
                requerimientos.Add(requirement2);
            }
            if (requirement3 != null && requirement3 != " " && requirement3 != "")
            {
                requerimientos.Add(requirement3);
            }
            if (requirement4 != null && requirement4 != " " && requirement4 != "")
            {
                requerimientos.Add(requirement4);
            }
            n_req = requerimientos.Count;
        }
        /// <summary>
        /// Asigna los valores acorde al numero del requerimiento (1-4)
        /// </summary>
        /// <param name="nreq">numero del requerimiento</param>
        public void AssignValues(int nreq)
        {
            if (n_req > 0)
            {
                string req = requerimientos[nreq].ToString();
                string[] parametros;

                parametros = req.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                proveedor = parametros[0].ToString();
                producto = parametros[1].ToString();
                requerimiento = parametros[2].ToString();
                cantidad = parametros[3].ToString();
                integracion = parametros[4].ToString();
                comentarios = parametros[5].ToString();

            }
        }

    }
}
