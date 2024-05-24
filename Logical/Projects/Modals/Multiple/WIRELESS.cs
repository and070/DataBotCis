using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical encargada de Wireless.
    /// </summary>
    class Wireless
    {
        public string WIRELESS_Cadena { set; get; }

        public string W1 { set; get; }
        public string W2 { set; get; }
        public string W3 { set; get; }
        public string W4 { set; get; }
        public string W5 { set; get; }
        public string W6 { set; get; }
        public string W7 { set; get; }
        public string W8 { set; get; }
        public string W9 { set; get; }
        public string W10 { set; get; }
        public string W11 { set; get; }
        public string W12 { set; get; }
        public string W13 { set; get; }
        public string W14 { set; get; }
        public string W15 { set; get; }
        public string W16 { set; get; }
        public string W17 { set; get; }
        public string W18 { set; get; }
        public string W19 { set; get; }
       

        public void AsignarCampos()
        {
            try
            {
                WIRELESS_Cadena = WIRELESS_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = WIRELESS_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            W1 = token3[0].ToString();
            W2 = token3[1].ToString();
            W3 = token3[2].ToString();
            W4 = token3[3].ToString();
            W5 = token3[4].ToString();
            W6 = token3[5].ToString();
            W7 = token3[6].ToString();
            W8 = token3[7].ToString();
            W9 = token3[8].ToString();
            W10 = token3[9].ToString();
            W11 = token3[10].ToString();
            W12 = token3[11].ToString();
            W13 = token3[12].ToString();
            W14 = token3[13].ToString();
            W15 = token3[14].ToString();
            W16 = token3[15].ToString();
            W17 = token3[16].ToString();
            W18 = token3[17].ToString();
            W19 = token3[18].ToString();
        
        }
    }
}
