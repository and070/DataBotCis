using System;
namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical de Power.
    /// </summary>
    class Power
    {
        public string PW_Cadena { set; get; }

        public string F1 { set; get; }
        public string F2 { set; get; }
        public string F3 { set; get; }
        public string F4 { set; get; }
        public string F5 { set; get; }
        public string F6 { set; get; }
        public string F7 { set; get; }
        public string F8 { set; get; }
        public string F9 { set; get; }
        public string F10 { set; get; }
        public string F11 { set; get; }
        public string F12 { set; get; }
        public string F13 { set; get; }
        public string F14 { set; get; }
        public string F15 { set; get; }
        public string F16 { set; get; }
        public string F17 { set; get; }
        public string F18 { set; get; }
        public string F19 { set; get; }
        public string F20 { set; get; }
        public string F21 { set; get; }
        public string F22 { set; get; }
        public string F23 { set; get; }
        public string F24 { set; get; }
        public string F25 { set; get; }
        public string F26 { set; get; }
        public string F27 { set; get; }
        public string F28 { set; get; }
        public string F29 { set; get; }
        public string F30 { set; get; }
        public string F31 { set; get; }
        public string F32 { set; get; }


        public string Comentarios { set; get; }

        public void AsignarCampos()
        {
            try
            {
                PW_Cadena = PW_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = PW_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            F1 = token3[0].ToString();
            F2 = token3[1].ToString();
            F3 = token3[2].ToString();
            F4 = token3[3].ToString();
            F5 = token3[4].ToString();
            F6 = token3[5].ToString();
            F7 = token3[6].ToString();
            F8 = token3[7].ToString();
            F9 = token3[8].ToString();
            F10 = token3[9].ToString();
            F11 = token3[10].ToString();
            F12 = token3[11].ToString();
            F13 = token3[12].ToString();
            F14 = token3[13].ToString();
            F15 = token3[14].ToString();
            F16 = token3[15].ToString();
            F17 = token3[16].ToString();
            F18 = token3[17].ToString();
            F19 = token3[18].ToString();
            F20 = token3[19].ToString();
            F21 = token3[20].ToString();
            F22 = token3[21].ToString();
            F23 = token3[22].ToString();
            F24 = token3[23].ToString();
            F25 = token3[24].ToString();
            F26 = token3[25].ToString();
            F27 = token3[26].ToString();
            F28 = token3[27].ToString();
            F29 = token3[28].ToString();
            F30 = token3[29].ToString();
            F31 = token3[30].ToString();
            F32 = token3[31].ToString();

            Comentarios = token3[32].ToString();
        }
    }
}
