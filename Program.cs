using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5
{
    class Program
    {
        static void Main(string[] args)
        {
        }
    }

    public class Errores
    {
        public List<ErroresF> ListadoErrores { get; set; }
    }

    public class ErroresF
    {
        public string Clase { get; set; }
        public string Error { get; set; }
        public string LineaCol { get; set; }
        public int Intentos { get; set; }
    }
}
