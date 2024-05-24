using DataBotV5.Logical.Webex;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace DataBotV5.Logical.Exceptions
{
    /// <summary>
    /// Clase Logical para formato de excepciones.
    /// </summary>
    class Exceptions
    {
        public Exceptions()
        {

        }
        public void ExceptionsFormat(StackTrace trace, Errores errs, string message)
        {

            string formato = "";
            string formato_dia = DateTime.Now.ToString("tt", CultureInfo.InvariantCulture);
            string saludo = "";
            if (formato_dia == "AM")
            {
                saludo = "Buenos días, \r\n";
            }
            else
            {
                saludo = "Buenas tardes, \r\n";
            }
            
            string programa_completo = trace.GetFrame(0).GetMethod().ReflectedType.FullName;
            string programa = trace.GetFrame(0).GetMethod().ReflectedType.Name;
            string espacio_programa = trace.GetFrame(0).GetMethod().ReflectedType.Namespace;
            int linea = trace.GetFrame(0).GetFileLineNumber();
            int columna = trace.GetFrame(0).GetFileColumnNumber();
            if (linea == 0 && columna == 0)
            {
                programa = "Conector.vb";
            }
            string lineacol = linea.ToString() + columna.ToString();
            formato += saludo;
            formato += "Se ha presentando un error en la clase: " + programa + "\r\n";
            formato += "El error ocurre en la línea # " + linea + ", columna # " + columna + "\r\n";
            formato += "Y su descrición corresponde a: " + "\r\n\r\n";
            formato += message + "\r\n\r\n";
            if ((errs.ListadoErrores) != null)
            {
                if ((errs.ListadoErrores).Exists(x => x.Clase == programa && x.LineaCol == lineacol))
                {
                    if (!((errs.ListadoErrores).Exists(x => x.Clase == programa && x.LineaCol == lineacol && x.Intentos <= 8 && x.Intentos >= 0)))
                    {
                        //envie el comunicado y resetee el intento a 0

                        WebexTeams webex = new WebexTeams();
                        webex.NotificationErrors(formato, linea, columna, programa, "Y2lzY29zcGFyazovL3VzL1JPT00vNDk4MjM4NzAtMzFiMy0xMWViLTk3ZjAtYzVjODdmZTg4ZjE3", "Error en " + programa);
                        int index = (errs.ListadoErrores).FindIndex(x => x.Clase == programa && x.LineaCol == lineacol);
                        (errs.ListadoErrores)[index].Intentos = 0;
                    }
                    else
                    {
                        int index = (errs.ListadoErrores).FindIndex(x => x.Clase == programa && x.LineaCol == lineacol);
                        (errs.ListadoErrores)[index].Intentos++;

                    }
                }
                else
                {
                    ErroresF errores = new ErroresF();

                    errores.Clase = programa;
                    errores.Error = message;
                    errores.LineaCol = lineacol;
                    errores.Intentos = 0;
                    (errs.ListadoErrores).Add(errores);
                    WebexTeams webex = new WebexTeams();
                    webex.NotificationErrors(formato, linea, columna, programa, "Y2lzY29zcGFyazovL3VzL1JPT00vNDk4MjM4NzAtMzFiMy0xMWViLTk3ZjAtYzVjODdmZTg4ZjE3", "Error en " + programa);
                }
            }
            else
            {
                List<ErroresF> lista = new List<ErroresF>();
                ErroresF errores = new ErroresF();

                errores.Clase = programa;
                errores.Error = message;
                errores.LineaCol = lineacol;
                errores.Intentos = 0;
                lista.Add(errores);
                errs.ListadoErrores = lista;
                WebexTeams webex = new WebexTeams();
                webex.NotificationErrors(formato, linea, columna, programa, "Y2lzY29zcGFyazovL3VzL1JPT00vNDk4MjM4NzAtMzFiMy0xMWViLTk3ZjAtYzVjODdmZTg4ZjE3", "Error en " + programa);
            }


        }
    }
}
