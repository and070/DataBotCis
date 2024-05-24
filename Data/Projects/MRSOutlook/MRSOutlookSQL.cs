using DataBotV5.Automation.MASS.MRS;
using DataBotV5.Data.Database;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;

namespace DataBotV5.Data.Projects.MRSOutlook
{
    class MRSOutlookSQL
    {
        CRUD crud = new CRUD();

        /// <summary>
        /// Método que extrae la información de la tabla mrs_outlook_sap.
        /// </summary>
        /// <param name="mandante">Mandante de SAP</param>
        /// <returns>Retorna un DataTable con los resultados de la extracción.</returns>
        public DataTable DecompressTask(int mandante)
        {
            DataTable data = new DataTable();
            string sql = "SELECT * FROM mrs_outlook_sap WHERE MAND = '" + mandante + "' AND ESTATUS = 'EN PROCESO' ORDER BY ID asc";
            //data = crud.Select("Databot", sql, "automation");

            return data;
        }
        /// <summary>
        ///Método que extrae la información de la tabla outlook_meetings.
        /// </summary>
        /// <param name="mandante">Mandante de SAP</param>
        /// <returns>Retorna un DataTable con los resultados de la extracción.</returns>
        public DataTable DecompressMeetings(int mandante)
        {
            DataTable data = new DataTable();
            string sql = "SELECT * FROM outlook_meetings WHERE MAND = '" + mandante + "' AND ESTATUS = 'EN PROCESO' ORDER BY ID asc";
            //data = crud.Select("Databot", sql, "automation");

            return data;
        }
        /// <summary>
        /// Método que almacena la información de un query en una clase para mayor facilidad de uso.
        /// </summary>
        /// <param name="guid">GUID de MRS</param>
        /// <returns>Retorna la clase con sus parámetros.</returns>
        public MRS_OUT DecompressOutlook(string guid)
        {
            MRS_OUT mrs = new MRS_OUT();
            DataTable data = new DataTable();
            string sql = "SELECT * FROM outlook_meetings WHERE MRS_GUID = '" + guid + "' AND ESTATUS = 'EN PROCESO' ORDER BY ID asc";
            //data = crud.Select("Databot", sql, "automation");

            if (data != null)
            {
                if (data.Rows.Count > 0)
                {
                    mrs.ID = data.Rows[0][0].ToString();
                    mrs.MAND = data.Rows[0][1].ToString();
                    mrs.OUTLOOK_IUD = data.Rows[0][2].ToString();
                    mrs.MRS_GUID = data.Rows[0][3].ToString();
                    mrs.EMPLOYEE = data.Rows[0][4].ToString();
                    mrs.EMAIL = data.Rows[0][5].ToString();
                    mrs.MEET_TYPE = data.Rows[0][6].ToString();
                    mrs.DESCRIPTION = data.Rows[0][7].ToString();
                    mrs.WF_ACTION = data.Rows[0][8].ToString();
                    mrs.INFO = data.Rows[0][9].ToString();
                    mrs.ESTATUS = data.Rows[0][10].ToString();
                    mrs.PLANNER = data.Rows[0][11].ToString();
                    mrs.TS = data.Rows[0][12].ToString();
                }
            }
            return mrs;
        }
        public MRS_OUT DecompressOutlookT(string guid)
        {
            MRS_OUT mrs = new MRS_OUT();
            DataTable data = new DataTable();
            string sql = "SELECT * FROM outlook_meetings WHERE MRS_GUID = '" + guid + "' AND ESTATUS = 'COMPLETADO' ORDER BY ID asc";
            //data = crud.Select("Databot", sql, "automation");

            if (data != null)
            {
                if (data.Rows.Count > 0)
                {
                    mrs.ID = data.Rows[0][0].ToString();
                    mrs.MAND = data.Rows[0][1].ToString();
                    mrs.OUTLOOK_IUD = data.Rows[0][2].ToString();
                    mrs.MRS_GUID = data.Rows[0][3].ToString();
                    mrs.EMPLOYEE = data.Rows[0][4].ToString();
                    mrs.EMAIL = data.Rows[0][5].ToString();
                    mrs.MEET_TYPE = data.Rows[0][6].ToString();
                    mrs.DESCRIPTION = data.Rows[0][7].ToString();
                    mrs.WF_ACTION = data.Rows[0][8].ToString();
                    mrs.INFO = data.Rows[0][9].ToString();
                    mrs.ESTATUS = data.Rows[0][10].ToString();
                    mrs.PLANNER = data.Rows[0][11].ToString();
                    mrs.TS = data.Rows[0][12].ToString();
                }
            }
            return mrs;
        }


        public List<ConfigColor> Allocations()
        {
            List<ConfigColor> listado = new List<ConfigColor>();
            DataTable data = new DataTable();
            string sql = "SELECT INFO FROM analytics_config WHERE APP = 'PLANNING_CONFIG'";
            //data = crud.Select("Databot", sql, "automation");

            listado = JsonConvert.DeserializeObject<List<ConfigColor>>(data.Rows[0][0].ToString());
            return listado;
        }
    }
}
