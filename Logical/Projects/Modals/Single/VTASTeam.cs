using System;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de VTASTeam
    /// </summary>
    class VTASTeam
    {
        public string Rol1 { set; get; }
        public string Empleado1 { set; get; }
        public string Rol2 { set; get; }
        public string Empleado2 { set; get; }
        public string Rol3 { set; get; }
        public string Empleado3 { set; get; }
        public string Rol4 { set; get; }
        public string Empleado4 { set; get; }
        public string Rol5 { set; get; }
        public string Empleado5 { set; get; }
        public string Rol6 { set; get; }
        public string Empleado6 { set; get; }
        public string Rol7 { set; get; }
        public string Empleado7 { set; get; }
        public string Rol8 { set; get; }
        public string Empleado8 { set; get; }
        public string Cadena { set; get; }


        public void AsignarCampos()
        {
            string[] token3 = Cadena.Split(new[] { "$" }, StringSplitOptions.RemoveEmptyEntries);

            try
            {
                string[] token4 = token3[0].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol1 = token4[0].ToString();
                Empleado1 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[1].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol2 = token4[0].ToString();
                Empleado2 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[2].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol3 = token4[0].ToString();
                Empleado3 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[3].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol4 = token4[0].ToString();
                Empleado5 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[4].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol5 = token4[0].ToString();
                Empleado5 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[5].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol6 = token4[0].ToString();
                Empleado6 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[6].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol7 = token4[0].ToString();
                Empleado7 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
            try
            {
                string[] token4 = token3[7].Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                Rol8 = token4[0].ToString();
                Empleado8 = token4[1].ToString();
            }
            catch (Exception)
            {

            }
        }
    }
}
