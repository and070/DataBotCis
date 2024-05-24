using DataBotV5.App.Global;
using DataBotV5.Data.Database;
using DataBotV5.Data.Process;
using DataBotV5.Data.Projects.CrBids;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Encode;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.CrBids;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

namespace DataBotV5.Automation.RPA2.CrBids
{
    class dbMigrar
    {
        BidsGbCrSql liccr = new BidsGbCrSql();
        ProcessAdmin padmin = new ProcessAdmin();
        CrBidsLogical crBids = new CrBidsLogical();
        ConsoleFormat console = new ConsoleFormat();
        Stats estadisticas = new Stats();
        Rooting root = new Rooting();
        WebInteraction web = new WebInteraction();
        MailInteraction mail = new MailInteraction();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        WebexTeams wt = new WebexTeams();
        CRUD crud = new CRUD();
        MsExcel msTeam = new MsExcel();
        string mandante = "QAS";
        public void Main()
        {
            DataTable results = new DataTable();
            results.Columns.Add("concurso");
            results.Columns.Add("query");
            results.Columns.Add("tabla");

            //DataTable bidIdsPurchaseOrder = crud.Select("Databot", "SELECT * FROM concursos2 WHERE JSON_VALUE(datos_sap, '$.participa') != ''", "licitaciones_cr");//tomar los concursos actuales
            //DataTable camposMigrar = crud.Select("Databot", "SELECT * FROM camposMigrar", "licitaciones_cr");//tomar los concursos actuales
            //DataTable empleados = crud.Select("Databot", "SELECT * FROM `empleados`", "fabrica_de_ofertas"); //tabla de empleados, para busacar los AM
            //DataTable bidsFiles = crud.Select("Databot", "SELECT CONCURSO, NOMBRE FROM concursos_archivos", "licitaciones_cr");

            DataTable bids = crud.Select( "SELECT bidNumber FROM `purchaseOrder`", "costa_rica_bids_db");
            DataTable processTypeData = crud.Select( "SELECT * FROM `processType`", "costa_rica_bids_db");
            DataTable valueTeamData = crud.Select( "SELECT * FROM `valueTeam`", "costa_rica_bids_db");
            DataTable oppTypeData = crud.Select( "SELECT * FROM `oppType`", "costa_rica_bids_db");
            DataTable salesTypeData = crud.Select( "SELECT * FROM `salesType`", "costa_rica_bids_db");
            DataTable productLineData = crud.Select( "SELECT * FROM `productLine`", "costa_rica_bids_db");
            DataTable noParticipationReasonData = crud.Select( "SELECT * FROM `noParticipationReason`", "costa_rica_bids_db");
            string json_concursos = "";
            using (StreamReader r = new StreamReader(root.downloadfolder + "\\concursos2.json"))
            {
                json_concursos = r.ReadToEnd();
            }
            DataTable concursos = JsonConvert.DeserializeObject<DataTable>(json_concursos);

            foreach (DataRow rRow in concursos.Rows)
            {
                string numBid = rRow["num_concurso"].ToString();
                try
                {

                    //DataRow[] exist = bids.Select($"bidNumber = '{numBid}'");
                    //if (exist.Count() > 0)
                    //{
                    //    continue;
                    //}


                    IDictionary<string, JObject> jaisons = new Dictionary<string, JObject>();
                    jaisons["purchaseOrder"] = new JObject();
                    jaisons["purchaseOrderAdditionalData"] = new JObject();
                    jaisons["products"] = new JObject();
                    jaisons["evaluations"] = new JObject();
                    Newtonsoft.Json.Linq.JObject datos_generales = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(rRow["datos_generales"].ToString()).ToString());
                    Newtonsoft.Json.Linq.JObject fechas = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(rRow["fechas"].ToString()).ToString());
                    foreach (var item in fechas)
                    {
                        if (DateTime.TryParse(item.Value.ToString(), out DateTime dt))
                        {
                            fechas[item.Key] = dt.ToString("yyyy-MM-dd hh:mm:ss");
                        }
                    }
                    Newtonsoft.Json.Linq.JObject datos_sap = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(rRow["datos_sap"].ToString()).ToString());
                    Newtonsoft.Json.Linq.JObject datos_adicionales = Newtonsoft.Json.Linq.JObject.Parse(JsonConvert.DeserializeObject(rRow["datos_adicionales"].ToString()).ToString());

                    JObject licitacion = new JObject();
                    licitacion.Merge(datos_generales);
                    licitacion.Merge(fechas);
                    licitacion.Merge(datos_sap);
                    licitacion.Merge(datos_adicionales);
                    string bienes_servicios = rRow["bienes_servicios"].ToString();
                    string evaluacion = rRow["evaluacion"].ToString();
                    licitacion["num_concurso"] = numBid;
                    licitacion["bienes_servicios"] = bienes_servicios;
                    licitacion["evaluacion"] = evaluacion;

                    //foreach (DataRow item in camposMigrar.Rows)
                    //{
                    //    //el nombre de la columna de todas las DB
                    //    string newColumn = item["newColumn"].ToString();
                    //    //es el campo de SICOP o del comentario en la BD (para los campos que no son de sicop)
                    //    string oldColumn = item["oldColumn"].ToString();
                    //    //el nombre de la tabla de la columna purchaseOrder, purchaseOrderAdditionalData, products, evaluations
                    //    string tableDb = item["tableNew"].ToString();

                    //    JObject jObject = jaisons[tableDb];

                    //    if (newColumn == "products")
                    //    {

                    //        jObject[newColumn] = bienes_servicios;
                    //    }
                    //    else if (newColumn == "evaluations")
                    //    {
                    //        jObject[newColumn] = evaluacion;
                    //    }
                    //    else if (newColumn == "gbmStatus")
                    //    {
                    //        string aRow = (licitacion[oldColumn].ToString().ToUpper() == "SI") ? "1" : "2";
                    //        jObject[newColumn] = aRow;
                    //    }
                    //    else if (newColumn == "changeDate")
                    //    {
                    //        string date = licitacion[oldColumn].ToString();
                    //        string user = date.Split('-')[0].ToString().Trim();
                    //        date = date.Split('-')[1].ToString().Trim();
                    //        jObject["changeDate"] = DateTime.Parse(date).ToString("yyyy-MM-dd HH:mm:ss");
                    //        jObject["participateUser"] = user;
                    //    }
                    //    else if (newColumn == "processType")
                    //    {
                    //        string processType = licitacion[oldColumn].ToString();
                    //        DataRow[] pType = processTypeData.Select($"processType = '{processType}'");
                    //        processType = (pType.Count() > 0) ? pType[0]["id"].ToString() : "NULL";
                    //        jObject[newColumn] = processType;

                    //    }
                    //    else if (newColumn == "productLine")
                    //    {
                    //        string pLine = licitacion[oldColumn].ToString();
                    //        DataRow[] pType = productLineData.Select($"productLine = '{pLine}'");
                    //        pLine = (pType.Count() > 0) ? pType[0]["id"].ToString() : "NULL";
                    //        jObject[newColumn] = pLine;

                    //    }

                    //    else if (newColumn == "receptionObjections" || newColumn == "receptionClarification" || newColumn == "offerOpening" || newColumn == "publicationDate" || newColumn == "changeDate" || newColumn == "receptionClosing")
                    //    {
                    //        string receptionObjections = licitacion[oldColumn].ToString();
                    //        jObject[newColumn] = (receptionObjections == "") ? "NULL" : receptionObjections;



                    //    }
                    //    else if (newColumn == "valueTeam")
                    //    {
                    //        string vTeam = licitacion[oldColumn].ToString();
                    //        DataRow[] pType = valueTeamData.Select($"valueTeam = '{vTeam}'");
                    //        vTeam = (pType.Count() > 0) ? pType[0]["id"].ToString() : "NULL";
                    //        jObject[newColumn] = vTeam;

                    //    }
                    //    else if (newColumn == "oppType")
                    //    {
                    //        string oType = licitacion[oldColumn].ToString();
                    //        DataRow[] pType = oppTypeData.Select($"key = '{oType}'");
                    //        oType = (pType.Count() > 0) ? pType[0]["id"].ToString() : "NULL";
                    //        jObject[newColumn] = oType;

                    //    }
                    //    else if (newColumn == "salesType")
                    //    {
                    //        string aType = licitacion[oldColumn].ToString();
                    //        DataRow[] pType = salesTypeData.Select($"salesType = '{aType}'");
                    //        aType = (pType.Count() > 0) ? pType[0]["id"].ToString() : "NULL";
                    //        jObject[newColumn] = aType;

                    //    }
                    //    else if (newColumn == "noParticipationReason")
                    //    {
                    //        string nReason = licitacion[oldColumn].ToString();
                    //        DataRow[] pType = noParticipationReasonData.Select($"noParticipationReason = '{nReason}'");
                    //        nReason = (pType.Count() > 0) ? pType[0]["id"].ToString() : "NULL";
                    //        jObject[newColumn] = nReason;

                    //    }
                    //    else //sin reglas especificas
                    //    {
                    //        string v = "";
                    //        try
                    //        {
                    //            //DataRow[] fila_select = licitacion.Select("campo = '" + oldColumn + "'");
                    //            //if (fila_select.Count() > 0)
                    //            //{
                    //            v = licitacion[oldColumn].ToString();
                    //            //}
                    //        }
                    //        catch (Exception) { }
                    //        JToken vv = jObject[newColumn];

                    //        if (vv == null)
                    //        {
                    //            jObject[newColumn] = v;
                    //        }
                    //        else
                    //        {
                    //            if (string.IsNullOrEmpty(vv.ToString()))
                    //            {
                    //                jObject[newColumn] = v;
                    //            }
                    //        }

                    //    }

                    //    jaisons[tableDb] = jObject;
                    //}
                   
                    int before = results.Rows.Count;
                    results = liccr.InsertRowSSOriginal(jaisons, results, numBid);
                    int after = results.Rows.Count;
                    if (before == after)
                    {
                        DataRow rrRow = results.Rows.Add();
                        rrRow["concurso"] = numBid;
                        rrRow["query"] = "No se ingreso";
                        rrRow["tabla"] = "NA";
                        results.AcceptChanges();
                    }
                    //else
                    //{
                    //    //ingresar sales teams
                    //    string salesTeam = licitacion["sales_team"].ToString();
                    //    string[] sTeam = salesTeam.Split(',');
                    //    foreach (string st in sTeam)
                    //    {
                    //        DataRow[] amInfo = empleados.Select("CODIGO = '" + st + "'");
                    //        if (amInfo.Count() > 0)
                    //        {
                    //            string name = amInfo[0]["NOMBRE"].ToString();
                    //            string user = amInfo[0]["USUARIO"].ToString();
                    //            string sql = $"INSERT INTO `salesTeam`(`bidNumber`, `salesTeam`, `salesTeamName`, `active`, `createdBy`, `createdAt`) VALUES ('','{user}','{name}', 1, 'databot');";
                    //            DataRow rrRow = results.Rows.Add();
                    //            rrRow["concurso"] = numBid;
                    //            rrRow["query"] = sql;
                    //            rrRow["tabla"] = "salesTeam";
                    //            results.AcceptChanges();
                    //        }
                    //    }
                    //    //ingresar adjuntos
                    //    DataRow[] adj = bidsFiles.Select("CONCURSO = '" + numBid + "'");
                    //    if (adj.Count() > 0)
                    //    {
                    //        foreach (DataRow nnRow in adj)
                    //        {
                    //            string fileName = nnRow["NOMBRE"].ToString();
                    //            string sql = $"INSERT INTO `uploadFiles` (`name`, `bidNumber`, `user`, `codification`, `type`, `path`, `active`, `createdBy`) VALUES ('{fileName}', '{numBid}', 'databot', '7bit', '{MimeMapping.GetMimeMapping(fileName)}', '/home/tss/projects/smartsimple/gbm-hub-api/src/assets/files/CrBids/{numBid}/{fileName}', 1, 'databot');";
                    //            DataRow rrRow = results.Rows.Add();
                    //            rrRow["concurso"] = numBid;
                    //            rrRow["query"] = sql;
                    //            rrRow["tabla"] = "uploadFiles";
                    //            results.AcceptChanges();
                    //        }
                    //    }
                    //}

                }
                catch (Exception ex)
                {
                    DataRow rrRow = results.Rows.Add();
                    rrRow["concurso"] = numBid;
                    rrRow["query"] = ex.ToString();
                    results.AcceptChanges();
                }
            }

            msTeam.CreateExcel(results, "resultados", root.downloadfolder + "//resultadosMigracionLCCR.xlsx");

        }
    }
}
