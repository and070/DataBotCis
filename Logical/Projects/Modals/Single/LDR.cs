using System;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de Ldr.
    /// </summary>
    class Ldr
    {
        public string Tipo { set; get; }
        public string Detalles { set; get; }
        public string Cadena { set; get; }

        public void AsignarCampos()
        {
            if (Cadena != "")
            {
                string[] token3 = Cadena.Split(new[] { "|||" }, StringSplitOptions.RemoveEmptyEntries);
                Tipo = token3[0].ToString();
                Detalles = Cadena.Replace(Tipo + "|||", "");
            }

        }
    }
}
