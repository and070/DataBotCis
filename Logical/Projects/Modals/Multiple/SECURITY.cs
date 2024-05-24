using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical de security.
    /// </summary>
    class Security
    {
        public string SECURITY_Cadena { set; get; }

        public string SC1 { set; get; }
        public string SC2 { set; get; }
        public string SC3 { set; get; }
        public string SC4 { set; get; }
        public string SC5 { set; get; }
        public string SC6 { set; get; }
        public string SC7 { set; get; }
        public string SC8 { set; get; }
        public string SC9 { set; get; }
        public string SC10 { set; get; }
        public string SC11 { set; get; }
        public string SC12 { set; get; }
        public string SC13 { set; get; }
        public string SC14 { set; get; }
        public string SC15 { set; get; }
        public string SC16 { set; get; }
        public string SC17 { set; get; }
        public string SC18 { set; get; }
        public string SC19 { set; get; }
        public string SC20 { set; get; }
        public string SC21 { set; get; }
       

        public void AsignarCampos()
        {
            try
            {
                SECURITY_Cadena = SECURITY_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = SECURITY_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            SC1  = token3[0].ToString();
            SC2  = token3[1].ToString();
            SC3  = token3[2].ToString();
            SC4  = token3[3].ToString();
            SC5  = token3[4].ToString();
            SC6  = token3[5].ToString();
            SC7  = token3[6].ToString();
            SC8  = token3[7].ToString();
            SC9  = token3[8].ToString();
            SC10 = token3[9].ToString();
            SC11 = token3[10].ToString();
            SC12 = token3[11].ToString();
            SC13 = token3[12].ToString();
            SC14 = token3[13].ToString();
            SC15 = token3[14].ToString();
            SC16 = token3[15].ToString();
            SC17 = token3[16].ToString();
            SC18 = token3[17].ToString();
            SC19 = token3[18].ToString();
            SC20 = token3[19].ToString();
            SC21 = token3[20].ToString();
           
        }
    }
}
