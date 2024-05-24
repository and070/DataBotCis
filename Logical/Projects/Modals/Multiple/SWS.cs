using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical encargada de Sws.
    /// </summary>
    class Sws
    {
        public string SWS_Cadena { set; get; }

        public string F1 { set; get; }
        public string F2 { set; get; }
        public string F3 { set; get; }
        public string F4 { set; get; }
        public string F5 { set; get; }
        public string F6 { set; get; }
        public string F7 { set; get; }



        public string Comentarios { set; get; }

        public void AsignarCampos()
        {
            try
            {
                SWS_Cadena = SWS_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = SWS_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            F1 = token3[0].ToString();
            F2 = token3[1].ToString();
            F3 = token3[2].ToString();
            F4 = token3[3].ToString();
            F5 = token3[4].ToString();
            F6 = token3[5].ToString();

            Comentarios = token3[6].ToString();
        }
    }
}
