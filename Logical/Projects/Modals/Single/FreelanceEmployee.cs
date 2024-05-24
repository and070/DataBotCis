using DataBotV5.Data.Database;
using System.Collections.Generic;
using System.Data;

namespace DataBotV5.Logical.Projects.Modals.Single
{
    /// <summary>
    /// Clase Logical encargada de Empleado de Freelance.
    /// </summary>
    class FreelanceEmployee
    {
        public string Nombre { set; get; }
        public string Correo { set; get; }
        public string Usuario { set; get; }
        public List<string> Gerencia { set; get; }

        public FreelanceEmployee(string usuario, int pruebas)
        {
            if(usuario != "") {
                Usuario = usuario;

            DataTable vendors = GetUserVendor(Usuario.ToLower());

            if(vendors.Rows.Count > 0)
            {
                //El usuario es un vendor o externo creado para pruebas
                Nombre = vendors.Rows[0][1].ToString();
                Correo = vendors.Rows[0][2].ToString();
            }
            else
            {
                if(pruebas != 0)
                    {
                        Nombre = "Diego Meza";
                        Correo = "dmeza@GBM.NET";
                    }
                    else
                    {
                        FLEmployee datos_emp = new FLEmployee(Usuario);
                        Correo = datos_emp.Correo;
                        Nombre = datos_emp.Nombre;
                    }

                //El usuario tiene que ser freelance registrado en SAP
               
            }

            }
        }
        private DataTable GetUserVendor(string usuario)
        {
            DataTable mytable = new DataTable();
            string sql = "SELECT * FROM freelance_a WHERE USUARIO = '"+usuario+"' AND ACTIVO = 'X'";
            //mytable = new CRUD().Select("Databot", sql, "automation");
            return mytable;
        }

        public void GetManagers()
        {
            List<string> lista_usuarios = new List<string>();
            List<string> lista_correos = new List<string>();
            DataTable mytable = new DataTable();
            string sql = "SELECT USUARIO FROM roles_dm WHERE ACCESO = 'CONSULTING' AND CATEGORIA = 'GERENTE'";
            //mytable = new CRUD().Select("Databot", sql, "automation");

            for (int i = 0; i < mytable.Rows.Count; i++)
            {
                lista_usuarios.Add(mytable.Rows[i][0].ToString());
            }

            for (int i = 0; i < lista_usuarios.Count; i++)
            {
                FLEmployee datos_emp = new FLEmployee(lista_usuarios[i]);
                lista_correos.Add(datos_emp.Correo);
            }
            Gerencia = lista_correos;
        }

    }
}
