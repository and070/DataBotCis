using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using DataBotV5.Data.Projects.MasterData;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;

using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.MicrosoftTools;
using System.Data;

namespace DataBotV5.Automation.DM.SpareParts
{
    /// <summary>
    /// Clase DM Automation encargada de la creación de repuestos de datos maestros.
    /// </summary>
    class SparePartsSS
    {
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        WebexTeams wt = new WebexTeams();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        WebInteraction web = new WebInteraction();
        Log log = new Log();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        SapVariants sap = new SapVariants();
        Stats stats = new Stats();
        
        MsExcel ms = new MsExcel();

        string fruName = "", description = "", materialGroup = "";

        string returnMsg = "";
        string fmRep = "";
        bool valData = true;
        public string resFailure = "";
        string resLog = "";
        int length;
        string erpMand = "ERP";
        bool returnRequest = false;

        string respFinal = "";



        public void Main()
        {
            string respuesta = DM.GetManagement("8"); //REPUESTOS
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessSpareParts();

                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessSpareParts()
        {
            try
            {
                int rows;
                string res1 = "", res2 = "", validate;

                #region extraer datos generales (cada clase ya que es data muy personal de la solicitud 
                JArray DG = JArray.Parse(root.datagDM);
                for (int i = 0; i < DG.Count; i++)
                {
                    JObject fila = JObject.Parse(DG[i].ToString());
                    materialGroup = fila["materialGroupSpartPartsCode"].Value<string>();
                }
                #endregion

                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII


                if (root.metodoDM == "1") //LINEAL
                {
                    JArray gestiones = JArray.Parse(root.requestDetails);
                    for (int i = 0; i < gestiones.Count; i++)
                    {
                        JObject fila = JObject.Parse(gestiones[i].ToString());

                        fruName = fila["spareId"].Value<string>().Trim().ToUpper();
                        description = fila["description"].Value<string>().Trim().ToUpper();

                        #region validaciones
                        fruName = fruName.ToUpper();

                        if (fruName.Length > 18)
                        {
                            returnMsg = "El codigo del repuesto debe de ser menor a 18 caracteres: " + fruName;
                            res2 = res2 + returnMsg + "<br>";
                            continue;
                        }

                        description = val.RemoveSpecialChars(description, 1);
                        description = description.ToUpper();
                        if (description.Length > 60)
                        { description = description.Substring(0, 60); }


                        fruName = fruName.Replace("á", "a"); fruName = fruName.Replace("é", "e"); fruName = fruName.Replace("í", "i"); fruName = fruName.Replace("ó", "o"); fruName = fruName.Replace("ú", "u"); fruName = fruName.Replace("ñ", "n");
                        #endregion

                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {

                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["MATERIAL"] = fruName,
                                ["DESCRIPCION"] = description,
                                ["GM"] = materialGroup
                            };

                            IRfcFunction func = sap.ExecuteRFC(erpMand, "ZDM_CREATE_REP", parameters);

                            #region Procesar Salidas del FM

                            res1 = res1 + fruName + ": " + func.GetValue("RESULTADO").ToString() + "<br>";

                            if (func.GetValue("RESULTADO").ToString() == "Material Bloqueado")
                            {
                                res2 = res2 + res1 + "<br>";
                                resLog = fruName + ": " + res1;
                                console.WriteLine(fruName + ": " + res1);
                            }
                            else if (func.GetValue("RESULTADO").ToString() == "Material ya existe")
                            {
                                #region Modificar Repuesto
                                Dictionary<string, string> parameters3 = new Dictionary<string, string>
                                {
                                    ["MATERIAL"] = fruName,
                                    ["MG"] = materialGroup,
                                    ["DESCRIPCION"] = description
                                };

                                IRfcFunction func_change = sap.ExecuteRFC(erpMand, "ZDM_CHANGE_MATERIAL", parameters3);

                                string res_descripcion = func_change.GetValue("RESULTADO_TEXT").ToString();
                                string res_mg = func_change.GetValue("RESULTADO_CAT").ToString();

                                if (res_descripcion == "" || res_descripcion == "Material ha sido actualizado")
                                    if (res_mg == "" || res_mg == "Material ha sido actualizado")
                                        res1 = "Material ha sido actualizado";
                                    else
                                        res1 = "Error: " + res_mg;
                                else
                                    res1 = "Error: " + res_descripcion;

                                console.WriteLine(fruName + ": " + res1);
                                res2 = res2 + fruName + ": " + res1 + "<br>";
                                resLog = fruName + ": " + res1;
                                #endregion
                            }
                            else
                            {
                                Dictionary<string, string> parameters4 = new Dictionary<string, string>
                                {
                                    ["MATERIAL"] = fruName
                                };
                                IRfcFunction func2 = sap.ExecuteRFC(erpMand, "ZDM_CREATE_EXTRA", parameters4);

                                res2 = res2 + fruName + ": " + func2.GetValue("RESULTADO").ToString() + "<br>";
                                resLog = fruName + ": " + func2.GetValue("RESULTADO").ToString();
                                console.WriteLine(fruName + ": " + func2.GetValue("RESULTADO").ToString());
                            }

                            //log de cambios base de datos
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Repuesto", resLog, root.Subject);
                            respFinal = respFinal + "\\n" + "Crear Repuesto: " + resLog;

                            if (res2.Contains("Favor contactar a Datos Maestros:"))
                            { valData = false; }
                            #endregion
                        }
                        catch (Exception ex)
                        {

                            resFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine("Finishing process " + resFailure);
                            res2 = res2 + fruName + ": " + ex.ToString() + "<br>";
                            resFailure = ex.ToString();
                            valData = false;
                        }
                        #endregion
                    }
                }
                else //MASIVO
                {
                    string attach = root.ExcelFile;
                    if (!String.IsNullOrEmpty(attach))
                    {
                        #region abrir excel
                        console.WriteLine("Abriendo excel y validando");
                        DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + attach);
                        rows = xlWorkSheet.Rows.Count;
                        #endregion

                        foreach (DataRow row in xlWorkSheet.Rows)
                        {
                            int i = xlWorkSheet.Rows.IndexOf(row);
                            fruName = row["COD MATERIAL"].ToString().Trim();
                            if (fruName != "")
                            {
                                description = row["TEXT ESPAÑOL MAYUSCULA MAX 40 CARACTERES"].ToString().Trim();
                                materialGroup = row["GRUPO ART."].ToString().Trim();

                                #region validación de datos
                               

                                if (description == "")
                                {
                                    returnMsg = "Por favor ingresar la descripcion";
                                    res2 = res2 + returnMsg + "<br>";
                                    continue;
                                }

                                if (materialGroup == "")
                                {
                                    returnMsg = "Por favor ingresar el material group";
                                    res2 = res2 + returnMsg + "<br>";
                                    continue;
                                }

                                length = (materialGroup.IndexOf("-") + 1);
                                if (length == 0)
                                    length = materialGroup.Length + 2;

                                materialGroup = materialGroup.Substring(0, length - 2);
                                materialGroup = materialGroup.Replace("#", "");

                                fruName = fruName.ToUpper();

                                if (fruName.Length > 18)
                                {
                                    returnMsg = "El codigo del repuesto debe de ser menor a 18 caracteres: " + fruName;
                                    res2 = res2 + returnMsg + "<br>";
                                    continue;
                                }

                                description = val.RemoveSpecialChars(description, 1);
                                description = description.ToUpper();
                                if (description.Length > 60)
                                    description = description.Substring(0, 60);

                                fruName = fruName.Replace("á", "a"); fruName = fruName.Replace("é", "e"); fruName = fruName.Replace("í", "i"); fruName = fruName.Replace("ó", "o"); fruName = fruName.Replace("ú", "u"); fruName = fruName.Replace("ñ", "n");
                                #endregion

                                #region SAP
                                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>
                                    {
                                        ["MATERIAL"] = fruName,
                                        ["DESCRIPCION"] = description,
                                        ["GM"] = materialGroup
                                    };

                                    IRfcFunction func = sap.ExecuteRFC(erpMand, "ZDM_CREATE_REP", parameters);

                                    #region Procesar Salidas del FM

                                    res1 = res1 + fruName + ": " + func.GetValue("RESULTADO").ToString() + "<br>";

                                    if (func.GetValue("RESULTADO").ToString() == "Material ya existe" || func.GetValue("RESULTADO").ToString() == "Material Bloqueado")
                                    {
                                        res2 = res2 + res1 + "<br>";
                                        resLog = fruName + ": " + res1;
                                        console.WriteLine(fruName + ": " + res1);
                                    }
                                    else
                                    {
                                        Dictionary<string, string> parameters2 = new Dictionary<string, string>
                                        {
                                            ["MATERIAL"] = fruName
                                        };

                                        IRfcFunction func2 = sap.ExecuteRFC(erpMand, "ZDM_CREATE_EXTRA", parameters2);

                                        res2 = res2 + fruName + ": " + func2.GetValue("RESULTADO").ToString() + "<br>";
                                        resLog = fruName + ": " + func2.GetValue("RESULTADO").ToString();
                                        console.WriteLine(fruName + ": " + func2.GetValue("RESULTADO").ToString());

                                    }

                                    //log de cambios base de datos
                                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Repuesto", resLog, root.Subject);
                                    respFinal = respFinal + "\\n" + "Crear Repuesto: " + resLog;

                                    if (res2.Contains("Favor contactar a Datos Maestros:"))
                                        valData = false;
                                    #endregion
                                }
                                catch (Exception ex)
                                {

                                    resFailure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                                    console.WriteLine(" Finishing process " + resFailure);
                                    res2 = res2 + fruName + ": " + ex.ToString() + "<br>";
                                    resFailure = ex.ToString();
                                    valData = false;
                                }

                                #endregion

                            } //IF si el repuesto esta en blanco

                        } //for para cada fila del excel

                    }
                    else
                    {
                        returnRequest = true;
                        res2 = "Error en la plantilla";
                    }

                }


                console.WriteLine("Finalizando solicitud");
                if (valData == false)
                {
                    //enviar email de repuesta de error a datos maestros
                    string[] cc = { "dmeza@gbm.net" };
                    DM.ChangeStateDM(root.IdGestionDM, res2 + "<br>" + resFailure, "4"); //ERROR
                    mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + res2 + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
                }
                else if (returnRequest == true)
                {
                    console.WriteLine("Devolviendo solicitud");
                    DM.ChangeStateDM(root.IdGestionDM, res2, "5"); //RECHAZADO
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Repuestos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + res2);
                }
                else
                {
                    //finalizar solicitud
                    DM.ChangeStateDM(root.IdGestionDM, res2, "3"); //FINALIZADO
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Repuestos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + res2);
                }


                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                string[] cc = { "dmeza@gbm.net" };
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4");//ERROR
                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }
    }
}
