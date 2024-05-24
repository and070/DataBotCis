using DataBotV5.Data.Projects.MasterData;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Mail;
using Newtonsoft.Json.Linq;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using System.Data;
using System;

namespace DataBotV5.Automation.DM.Materials
{
    /// <summary>
    /// Clase DM Automation encargada de la creación de materiales de terceros en datos maestros.
    /// </summary>
    class ThirdMaterialsSS
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        MasterDataSqlSS dm = new MasterDataSqlSS();
        Rooting root = new Rooting();
        MsExcel ms = new MsExcel();
        Log log = new Log();

        string respFinal = "";

        public void Main()
        {
            string res = dm.GetManagement("6"); //TERCEROS materiales de servicios
            if (!String.IsNullOrEmpty(res) && res != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessThirdPartyMaterials();

                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessThirdPartyMaterials()
        {
            string[] cc = { "hlherrera@gbm.net" };
            try
            {
                string[] emailAttach = new string[1];

                #region extraer datos generales (cada clase ya que es data muy personal de la solicitud 
                JArray DG = JArray.Parse(root.datagDM);
                string materialGroup;
                for (int i = 0; i < DG.Count; i++)
                {
                    JObject fila = JObject.Parse(DG[i].ToString());
                    materialGroup = fila["materialGroupCode"].Value<string>();
                }
                #endregion

                materialGroup = root.factorDM;
               
                //Por cada adjunto de la solicitud
                if (root.metodoDM == "1") //lineal
                {
                    JArray requests = JArray.Parse(root.requestDetails);

                    DataTable excelResult = new DataTable();
                    
                    excelResult.Columns.Add("Tipo de Material");
                    excelResult.Columns.Add("Codigo de Material");
                    excelResult.Columns.Add("Descripcion");
                    excelResult.Columns.Add("Unidad");
                    excelResult.Columns.Add("Grupo Articulo");
                    excelResult.Columns.Add("Grupo Material 1");
                    excelResult.Columns.Add("Grupo Tipo Posicion");
                    excelResult.Columns.Add("Descripcion Larga");
                    excelResult.Columns.Add("Costo");
                    excelResult.Columns.Add("Solicitante");

                    for (int i = 0; i < requests.Count; i++)
                    {
                        JObject row = JObject.Parse(requests[i].ToString());
                        string material = row["idMaterial"].Value<string>().Trim().ToUpper();
                        string matDesc = row["description"].Value<string>().Trim().ToUpper();
                        string matType = "ZSER";
                        string unitMeasure = row["meditUnitCode"].Value<string>().Trim().ToUpper();
                        string hierarchy = row["hierarchyCode"].Value<string>().Trim().ToUpper();
                        string gbm1 = row["materialGroup1Code"].Value<string>().Trim().ToUpper();
                        string commercialSector = row["commercialSectorCode"].Value<string>().Trim().ToUpper();
                        string price = row["price"].Value<string>().Trim().ToUpper();
                        string longDesc = row["longDescription"].Value<string>().Trim().ToUpper();

                        DataRow rRow = excelResult.Rows.Add();

                        rRow["Tipo de Material"] = matType;
                        rRow["Codigo de Material"] = material;
                        rRow["Descripcion"] = matDesc;
                        rRow["Unidad"] = unitMeasure;
                        rRow["Grupo Articulo"] = materialGroup;
                        rRow["Grupo Material 1"] = gbm1;
                        rRow["Grupo Tipo Posicion"] = "LEIS";
                        rRow["Descripcion Larga"] = longDesc;
                        rRow["Costo"] = price;
                        rRow["Solicitante"] = root.BDUserCreatedBy;

                        excelResult.AcceptChanges();

                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Agregar fila lineal material terceros a plantilla", material + ", " + matDesc, root.Subject);
                    respFinal = respFinal + "\\n" + "Agregar fila lineal material terceros a plantilla: "+ material + ", " + matDesc;

                    }
                    string attachment = "plantilla_de_materiales_terceros_" + DateTime.Now.ToString("yyyy-mm-dd") + ".xlsx";
                    ms.CreateExcel(excelResult, "Datos", root.FilesDownloadPath + "\\" + attachment, true);
                  
                    emailAttach[0] = root.FilesDownloadPath + "\\" + attachment;
                    



                }
                else //MASIVO
                {
                    string attachFile = root.ExcelFile; //ya viene 
                    if (!String.IsNullOrEmpty(attachFile))
                    {
                        //#region abrir excel
                        console.WriteLine("Abriendo excel y validando");
                        emailAttach[0] = root.FilesDownloadPath + "\\" + attachFile;
                    }

                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Material de terceros", $"La plantilla enviada por {root.BDUserCreatedBy} de material de terceros viene adjunta y se procede a enviar por correo", root.Subject);
                    respFinal = respFinal + "\\n" + $"La plantilla enviada por {root.BDUserCreatedBy} de material de terceros viene adjunta y se procede a enviar por correo";

                }

                console.WriteLine(DateTime.Now + " > > > " + "Finalizando solicitud");

                //enviar email de repuesta de error a datos maestros
                dm.ChangeStateDM(root.IdGestionDM, "En Proceso", "14"); //PENDIENTE
                mail.SendHTMLMail("Se ha creado una nueva solicitud de Materiales de Terceros, por favor su asistencia para crearlo mediante el excel adjunto", new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, cc, emailAttach);
                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                dm.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //ERROR
                mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }
    }
}
