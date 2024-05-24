using DataBotV5.App.Global;
using DataBotV5.Data.Projects.DrBids;
using DataBotV5.Logical.Processes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace DataBotV5.Logical.Projects.DrBids
{
    /// <summary>
    /// Clase Logical encargada de filtro de licitaciones.
    /// </summary>
    class BidsFilter
    {
        ValidateData val = new ValidateData();
        /// <summary>
        /// Metodo encargado de ver si hay palabras contenidas en un texto 
        /// </summary>
        /// <param name="texto"></param> El texto en el que buscaremos las palabras
        /// <param name="words"></param> Lista de palabras que buscaremos en el texto
        /// <returns></returns>
        public string KeyMatch(string texto, List<string> words)
        {
            string interes_gbm = "NO";
            try
            {
                texto = val.RemoveAccents(texto);
                var result = words.Where(x => texto.Contains(x)).ToList();
                if (result.Count > 0)
                { interes_gbm = "SI"; }
            }
            catch (Exception)
            {
            }
            return interes_gbm;
        }
       /// <summary>
       /// Trae palabras claves de la tabla key_words de la base de datos
       /// </summary>
       /// <returns></returns>
        public List<string> KeyWord()
        {

            List<string> stopWords = new List<string>();
            try
            {
                DataTable mytable;
                BidsGbDrSql sql = new BidsGbDrSql();
                
                   mytable = sql.SelectRow("SELECT * FROM `keyWords`");
                   
                
                if (mytable.Rows.Count > 0)
                {
                    for (int i = 0; i < mytable.Rows.Count; i++)
                    {
                        string key = mytable.Rows[i][2].ToString().ToLower();
                        key = val.RemoveSpecialChars(key, 1);
                        stopWords.Add(key); //palabra clave
                    }
                }
            }
            catch (Exception ex)
            {
                new ConsoleFormat().WriteLine(ex.Message);
            }
            return stopWords;
        }
    }
}
