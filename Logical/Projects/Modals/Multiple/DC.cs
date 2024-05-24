using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical Dc.
    /// </summary>
    class Dc
    {
        public string DC_Cadena { set; get; }

        public string DC1 { set; get; }
        public string DC2 { set; get; }
        public string DC3 { set; get; }
        public string DC4 { set; get; }
        public string DC5 { set; get; }
        public string DC6 { set; get; }
        public string DC7 { set; get; }
        public string DC8 { set; get; }
        public string DC9 { set; get; }
       

        public void AsignarCampos()
        {
            try
            {
                DC_Cadena = DC_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = DC_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            DC1 = token3[0].ToString();
            DC2 = token3[0].ToString();
            DC3 = token3[0].ToString();
            DC4 = token3[0].ToString();
            DC5 = token3[0].ToString();
            DC6 = token3[0].ToString();
            DC7 = token3[0].ToString();
            DC8 = token3[0].ToString();
            DC9 = token3[0].ToString();

        }
    }
}
