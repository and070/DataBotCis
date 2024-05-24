using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical collab.
    /// </summary>
    class Collab
    {
        public string COLLAB_Cadena { set; get; }

        public string T1 { set; get; }
        public string T2 { set; get; }
        public string T3 { set; get; }
        public string T4 { set; get; }
        public string T5 { set; get; }
        public string T6 { set; get; }
        public string T7 { set; get; }
        public string T8 { set; get; }
        public string T9 { set; get; }
        public string T10 { set; get; }
        public string T11 { set; get; }
        public string T12 { set; get; }
        public string T13 { set; get; }
        public string T14 { set; get; }
        public string T15 { set; get; }
        public string T16 { set; get; }
        public string T17 { set; get; }
        public string T18 { set; get; }
        public string T19 { set; get; }
        public string T20 { set; get; }
        public string T21 { set; get; }
        public string T22 { set; get; }
        public string T23 { set; get; }
        public string T24 { set; get; }
        public string T25 { set; get; }

        public string TV1 { set; get; }
        public string TV2 { set; get; }
        public string TV3 { set; get; }
        public string TV4 { set; get; }
        public string TV5 { set; get; }
        public string TV6 { set; get; }
        public string TV7 { set; get; }
        public string TV8 { set; get; }
        public string TV9 { set; get; }
        public string TV10 { set; get; }
        public string TV11 { set; get; }
        public string TV12 { set; get; }
        public string TV13 { set; get; }
        public string TV14 { set; get; }
        public string TV15 { set; get; }
        public string TV16 { set; get; }
        public string TV17 { set; get; }

        public void AsignarCampos()
        {
            try
            {
               COLLAB_Cadena = COLLAB_Cadena.Replace("undefined", "N/A");
                COLLAB_Cadena = COLLAB_Cadena.Replace(" ", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = COLLAB_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            T1 = token3[0].ToString();
            T2 = token3[1].ToString();
            T3 = token3[2].ToString();
            T4 = token3[3].ToString();
            T5 = token3[4].ToString();
            T6 = token3[5].ToString();
            T7 = token3[6].ToString();
            T8 = token3[7].ToString();
            T9 = token3[8].ToString();
            T10 = token3[9].ToString();
            T11 = token3[10].ToString();
            T12 = token3[11].ToString();
            T13 = token3[12].ToString();
            T14 = token3[13].ToString();
            T15 = token3[14].ToString();
            T16 = token3[15].ToString();
            T17 = token3[16].ToString();
            T18 = token3[17].ToString();
            T19 = token3[18].ToString();
            T20 = token3[19].ToString();
            T21 = token3[20].ToString();
            T22 = token3[21].ToString();
            T23 = token3[22].ToString();
            T24 = token3[23].ToString();
            T25 = token3[24].ToString();

            TV1 = token3[25].ToString();
            TV2 = token3[26].ToString();
            TV3 = token3[27].ToString();
            TV4 = token3[28].ToString();
            TV5 = token3[29].ToString();
            TV6 = token3[30].ToString();
            TV7 = token3[31].ToString();
            TV8 = token3[32].ToString();
            TV9 = token3[33].ToString();
            TV10 = token3[34].ToString();
            TV11 = token3[35].ToString();
            TV12 = token3[36].ToString();
            TV13 = token3[37].ToString();
            TV14 = token3[38].ToString();
            TV15 = token3[39].ToString();
            TV16 = token3[40].ToString();
            TV17 = token3[41].ToString();
            //TS = token3[21].ToString();
            //SLA = token3[22].ToString();
            //DS = token3[23].ToString();
            //I_O_M = token3[24].ToString();

            //Comentarios = token3[25].ToString();
        }

    }
}
