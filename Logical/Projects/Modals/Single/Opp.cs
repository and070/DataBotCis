using System;
namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de oportunidades.
    /// </summary>
    class Opp
    {
        public string Tipo { set; get; }
        public string Descripcion { set; get; }
        public string Fecha_Inicial { set; get; }
        public string Fecha_Final { set; get; }
        public string Ciclo { set; get; }
        public string Origen { set; get; }
        public string Usuario { set; get; }

        public string Cadena { set; get; }

        public void AsignarCapos()
        {
            string[] token3 = Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            Tipo = token3[0].ToString();
            Descripcion = token3[1].ToString();
            Fecha_Inicial = token3[2].ToString();
            Fecha_Final = token3[3].ToString();
            Ciclo = token3[4].ToString();
            Origen = token3[5].ToString();
            try
            {
                string[] fi = Fecha_Inicial.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                string[] ff = Fecha_Final.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);

                //11/18/2019

                Fecha_Inicial = fi[2] + fi[0] + fi[1];
                Fecha_Final = ff[2] + ff[0] + ff[1];
            }
            catch (Exception e)
            {


            }


        }
    }
}
