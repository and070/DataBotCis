using System;

namespace DataBotV5.Logical.Projects.Modals.Multiple
{
    /// <summary>
    /// Clase Logical encargada de Vmware.
    /// </summary>
    class Vmware
    {
        public string VM_Cadena { set; get; }

        public string F1 { set; get; }
        public string F2 { set; get; }
        public string F3 { set; get; }
        public string F4 { set; get; }
        public string F5 { set; get; }



        public string Comentarios { set; get; }

        public void AsignarCampos()
        {
            try
            {
                VM_Cadena = VM_Cadena.Replace("undefined", "N/A");
            }
            catch (Exception)
            {

            }
            string[] token3 = VM_Cadena.Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

            F1 = token3[0].ToString();
            F2 = token3[1].ToString();
            F3 = token3[2].ToString();
            F4 = token3[3].ToString();
            F5 = token3[4].ToString();


            Comentarios = token3[5].ToString();
        }

    }
}
