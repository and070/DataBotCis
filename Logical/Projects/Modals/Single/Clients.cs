using System;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de Clients.
    /// </summary>
    class Clients
    {
        public string Cliente { set; get; }
        public string Contacto { set; get; }
        public string OrgVentas { set; get; }
        public string OrgServicios { set; get; }

        public string Cadena { set; get; }

        public void AsignarCampos()
        {
            string[] token3 = Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            Cliente = token3[0].ToString();
            Contacto = token3[1].ToString();
            OrgVentas = token3[2].ToString();
            OrgServicios = token3[3].ToString();
        }
    }
}
