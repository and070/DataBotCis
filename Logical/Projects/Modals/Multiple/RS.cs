using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical de RS.
    /// </summary>
    class RS
    {
        public string RS_Cadena { set; get; }

        public string S1 { set; get; }
        public string S2 { set; get; }
        public string S3 { set; get; }
        public string S4 { set; get; }
        public string S5 { set; get; }
        public string S6 { set; get; }
        public string S7 { set; get; }
        public string S8 { set; get; }
        public string S9 { set; get; }
        public string S10 { set; get; }
        public string S11 { set; get; }
        public string S12 { set; get; }
        public string S13 { set; get; }
        public string S14 { set; get; }
        public string S15 { set; get; }
        public string S16 { set; get; }
        public string S17 { set; get; }


        public string R1 { set; get; }
        public string R2 { set; get; }
        public string R3 { set; get; }
        public string R4 { set; get; }
        public string R5 { set; get; }
        public string R6 { set; get; }
        public string R7 { set; get; }
        public string R8 { set; get; }
        public string R9 { set; get; }
        public string R10 { set; get; }
        public string R11 { set; get; }
        public string R12 { set; get; }

        //public string TS { set; get; }
        //public string SLA { set; get; }
        //public string DS { set; get; }
        //public string I_O_M { set; get; }


        //public string Comentarios { set; get; }

        public void AsignarCampos()
        {
            try
            {
                RS_Cadena = RS_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = RS_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            S1 = token3[0].ToString();
            S2 = token3[1].ToString();
            S3 = token3[2].ToString();
            S4 = token3[3].ToString();
            S5 = token3[4].ToString();
            S6 = token3[5].ToString();
            S7 = token3[6].ToString();
            S8 = token3[7].ToString();
            S9 = token3[8].ToString();
            S10 = token3[9].ToString();
            S11 = token3[10].ToString();
            S12 = token3[11].ToString();
            S13 = token3[12].ToString();
            S14 = token3[13].ToString();
            S15 = token3[14].ToString();
            S16 = token3[15].ToString();
            S17 = token3[16].ToString();


            R1 = token3[17].ToString();
            R2 = token3[18].ToString();
            R3 = token3[19].ToString();
            R4 = token3[20].ToString();
            R5 = token3[21].ToString();
            R6 = token3[22].ToString();
            R7 = token3[23].ToString();
            R8 = token3[24].ToString();
            R9 = token3[25].ToString();
            R10 = token3[26].ToString();
            R11 = token3[27].ToString();
            R12 = token3[28].ToString();

        }
    }
}
