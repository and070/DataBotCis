using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Projects.MasterData;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.SAP;
using System.Data;
using DataBotV5.Logical.MicrosoftTools;

namespace DataBotV5.Automation.DM.Services
{   /// <summary>
    /// Clase DM Automation encargada de los servicios de datos maestros.
    /// </summary>
    class ServicesSS
    {
        #region Variables Globales
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        WebexTeams wt = new WebexTeams();
        ProcessInteraction proc = new ProcessInteraction();
        WebInteraction web = new WebInteraction();
        Log log = new Log();
        Stats estadisticas = new Stats();
        
        MsExcel ms = new MsExcel();
        string erpMand = "ERP";
        int lenght;

        string service;
        string desc;
        string matType;
        string unit;
        string hierarchy;
        string materialGroup;
        string servProfile;
        string responseProfile;
        string gm1;
        string longText;
        string price;
        string validacion = "";
        string mensaje_devolucion = "";
        string fmrep = "";
        bool valData = true;
        public string resFailure = "";
        bool returnRequest = false;

        int rows = 0;
        int startRow = 0;
        string res1 = "", res2 = "";

        string respFinal = "";

        #endregion
        public void Main()
        {
            string respuesta = DM.GetManagement("7");
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessServices();
                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessServices()
        {
            try
            {
                #region extraer datos generales (cada clase ya que es data muy personal de la solicitud 
                JArray DG = JArray.Parse(root.datagDM);
                for (int i = 0; i < DG.Count; i++)
                {
                    JObject fila = JObject.Parse(DG[i].ToString());
                    materialGroup = fila["materialGroupCode"].Value<string>();
                }
                #endregion
                //materialGroup = root.factorDM;
                //Por cada adjunto de la solicitud

                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                if (root.metodoDM == "1")
                {
                    JArray gestiones = JArray.Parse(root.requestDetails);
                    for (int i = 0; i < gestiones.Count; i++)
                    {
                        JObject fila = JObject.Parse(gestiones[i].ToString());
                        matType = fila["materialTypeCode"].Value<string>().Trim().ToUpper();
                        service = fila["idMaterial"].Value<string>().Trim().ToUpper();
                        desc = fila["description"].Value<string>().Trim().ToUpper();
                        unit = fila["meditUnitCode"].Value<string>().Trim().ToUpper();
                        servProfile = fila["serviceProfileCode"].Value<string>().Trim().ToUpper();
                        responseProfile = fila["responseProfileCode"].Value<string>().Trim().ToUpper();
                        hierarchy = fila["hierarchyCode"].Value<string>();
                        gm1 = fila["materialGroup1Code"].Value<string>();
                        longText = fila["largeDescription"].Value<string>().Trim().ToUpper();
                        price = fila["price"].Value<string>();

                        #region validación de datos

                        if (service.Length > 18)
                        {
                            mensaje_devolucion = "El material: " + service + " La longitud del material supera los 18 caracteres";
                            res2 = res2 + mensaje_devolucion + "<br>";
                            continue;
                        }

                        if (service.Substring(service.Length - 2, 2) == "XX")
                        {
                            mensaje_devolucion = service + ": Por favor indicar el ítem con su numero completo, si no sabe el consecutivo por favor preguntarle a Datos Maestros";
                            res2 = res2 + mensaje_devolucion + "<br>";
                            continue;
                        }

                        desc = val.RemoveSpecialChars(desc, 1);
                        service = val.RemoveSpecialChars(service, 1);
                        service = service.ToUpper();
                        longText = val.RemoveSpecialChars(longText, 1);
                        if (desc.Length > 40)
                            desc = desc.Substring(0, 40);

                        if (price == "Vacio")
                            price = "";

                        if (price != "")
                        {
                            price = price.Replace("$", "");

                            if (price.Length > 3)
                            {
                                if (price.Substring(0, 3).Substring(price.Substring(0, 3).Length - 1, 1) == ","
                                    && price.Substring((price.Length - 3)).Substring(0, 1) != ","
                                    || price.Substring(0, 2).Substring(price.Substring(0, 2).Length - 1, 1) == ","
                                    && price.Substring(price.Length - 3, 1).Substring(0, 1) != ",")
                                {
                                    price = price.Replace(",", "");
                                }

                                if (((price.Substring((price.Length - 3)).Substring(0, 1) == ".") || (price.Substring((price.Length - 2)).Substring(0, 1) == ".")))
                                {
                                    // ejemplo 100.000,34 ---- 100000.34
                                    price = price.Replace(",", "");
                                    price = price.Replace(".", ",");
                                }

                            }

                            if (price == "9999999" || price == "999999.99" || price == "999,999")
                                price = "999999";

                            if (price == "0" || price == "0,00" || price == "0.00")
                                price = "";
                        }

                        if (matType == "SER_0002" && price == "")
                            price = "999999";

                        if (matType == "SER_0001" && price == "999999")
                            price = "";

                        #endregion validación de datos

                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["CATEGORIA"] = matType,
                                ["SERVICE"] = service,
                                ["STEXT"] = desc,
                                ["UNIDAD"] = unit,
                                ["SERVICE_PROFILE"] = servProfile,
                                ["RESPONSE_PROFILE"] = responseProfile,
                                ["GRUPO_ARTICULO"] = materialGroup,
                                ["JERARQUIA"] = hierarchy,
                                ["GM1"] = gm1,
                                ["LTEXT"] = longText,
                                ["PRECIO"] = price
                            };

                            IRfcFunction func = new SapVariants().ExecuteRFC(erpMand, "ZDM_CREATE_SERV", parameters);


                            #region Procesar Salidas del FM
                            res1 = res1 + service + ": " + func.GetValue("RESPUESTA").ToString() + "<br>";
                            //log de cambios base de datos
                            console.WriteLine(service + ": " + func.GetValue("RESPUESTA").ToString());
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Servicio", service + ": " + func.GetValue("RESPUESTA").ToString(), root.Subject);
                            respFinal = respFinal + "\\n" + "Crear Servicio" + service + ": " + func.GetValue("RESPUESTA").ToString();

                            if (res1.Contains("Error"))
                                valData = false;
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            resFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + resFailure);
                            res1 = res1 + service + ": " + ex.ToString() + "<br>";
                            resFailure = ex.ToString();
                            valData = false;
                        }

                        #endregion
                    }
                }
                else //MASIVO
                {
                    string adjunto = root.ExcelFile; //ya viene 
                    if (!String.IsNullOrEmpty(adjunto))
                    {
                        #region abrir excel
                        console.WriteLine("Abriendo excel y validando");
                        DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + adjunto);
                        rows = xlWorkSheet.Rows.Count;
                        #endregion

                        foreach (DataRow row in xlWorkSheet.Rows)
                        {
                            int i = xlWorkSheet.Rows.IndexOf(row);
                            matType = row["Categoria Servicio"].ToString().Trim();
                            if (matType != "")
                            {
                                if (matType.Length > 8)
                                    matType = matType.Substring(0, 8);

                                service = row["Codigo Material*"].ToString().Trim().ToUpper();
                                desc = row["Descripcion Material*"].ToString().Trim().ToUpper();
                                unit = row["Unidad Medida*"].ToString().Trim().ToUpper();
                                servProfile = row["Service Profile"].ToString().Trim().ToUpper();
                                responseProfile = row["Response Profile"].ToString().Trim().ToUpper();
                                materialGroup = row["Grupo Articulo*"].ToString().Trim();
                                hierarchy = row["Jerarquía"].ToString().Trim();
                                gm1 = row["Grupo Material 1*"].ToString().Trim();
                                longText = row["Descripción Larga*"].ToString().Trim().ToUpper();
                                price = row["Precio"].ToString().Trim();

                                #region validación de datos

                                

                                if (service == "" || desc == "" || unit == "" || materialGroup == "" || gm1 == "" || longText == "")
                                {
                                    mensaje_devolucion = service + ": " + "Por favor ingresar los campos obligatorios";
                                    res2 = res2 + mensaje_devolucion + "<br>";
                                    continue;
                                }

                                lenght = (materialGroup.IndexOf("-") + 1);
                                if (lenght == 0)
                                    lenght = materialGroup.Length + 2;

                                materialGroup = materialGroup.Substring(0, lenght - 2);
                                materialGroup = materialGroup.Replace("#", "");

                                lenght = (gm1.IndexOf("-") + 1);
                                if (lenght == 0)
                                    lenght = gm1.Length + 2;

                                gm1 = gm1.Substring(0, lenght - 2);

                                if (gm1.Length < 2 && gm1 != "")
                                    gm1 = ("0" + gm1);

                                lenght = (servProfile.IndexOf("=") + 1);
                                if ((lenght == 0))
                                    lenght = (servProfile.Length + 2);

                                servProfile = servProfile.Substring(0, (lenght - 2));

                                lenght = (unit.IndexOf("-") + 1);
                                if ((lenght == 0))
                                    lenght = (unit.Length + 2);

                                unit = unit.Substring(0, (lenght - 2));

                                lenght = (responseProfile.IndexOf("=") + 1);
                                if ((lenght == 0))
                                    lenght = (responseProfile.Length + 2);

                                responseProfile = responseProfile.Substring(0, (lenght - 2));

                                if (service.Length > 18)
                                {
                                    mensaje_devolucion = "El material: " + service + " La longitud del material supera los 18 caracteres";
                                    res2 = res2 + mensaje_devolucion + "<br>";
                                    continue;
                                }

                                if (service.Substring(service.Length - 2, 2) == "XX")
                                {
                                    mensaje_devolucion = service + ": Por favor indicar el item con su numero completo, si no sabe el consecutivo por favor preguntarle a Datos Maestros";
                                    res2 = res2 + mensaje_devolucion + "<br>";
                                    continue;
                                }

                                desc = val.RemoveSpecialChars(desc, 1);
                                service = val.RemoveSpecialChars(service, 1);
                                service = service.ToUpper();
                                longText = val.RemoveSpecialChars(longText, 1);
                                if (desc.Length > 40)
                                    desc = desc.Substring(0, 40);


                                if (price == "Vacio")
                                    price = "";


                                if (price != "")
                                {
                                    price = price.Replace("$", "");

                                    if (price.Length > 3)
                                    {
                                        if (price.Substring(0, 3).Substring(price.Substring(0, 3).Length - 1, 1) == ","

                                             && price.Substring((price.Length - 3)).Substring(0, 1) != ","

                                            || price.Substring(0, 2).Substring(price.Substring(0, 2).Length - 1, 1) == ","

                                             && price.Substring(price.Length - 3, 1).Substring(0, 1) != ",")
                                        {
                                            price = price.Replace(",", "");
                                        }

                                        if (((price.Substring((price.Length - 3)).Substring(0, 1) == ".") || (price.Substring((price.Length - 2)).Substring(0, 1) == ".")))
                                        {
                                            // ejemplo 100.000,34 ---- 100000.34
                                            price = price.Replace(",", "");
                                            price = price.Replace(".", ",");
                                        }

                                    }

                                    if (price == "9999999" || price == "999999.99" || price == "999,999")
                                        price = "999999";

                                    if (price == "0" || price == "0,00" || price == "0.00")
                                        price = "";
                                }

                                if (matType == "SER_0002" && price == "")
                                    price = "999999";

                                if (matType == "SER_0001" && price == "999999")
                                    price = "";

                                #endregion validación de datos

                                #region SAP
                                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>
                                    {
                                        ["CATEGORIA"] = matType,
                                        ["SERVICE"] = service,
                                        ["STEXT"] = desc,
                                        ["UNIDAD"] = unit,
                                        ["SERVICE_PROFILE"] = servProfile,
                                        ["RESPONSE_PROFILE"] = responseProfile,
                                        ["GRUPO_ARTICULO"] = materialGroup,
                                        ["JERARQUIA"] = hierarchy,
                                        ["GM1"] = gm1,
                                        ["LTEXT"] = longText,
                                        ["PRECIO"] = price
                                    };

                                    IRfcFunction func = new SapVariants().ExecuteRFC(erpMand, "ZDM_CREATE_SERV", parameters);

                                    #region Procesar Salidas del FM
                                    res1 = res1 + service + ": " + func.GetValue("RESPUESTA").ToString() + "<br>";
                                    //log de cambios base de datos
                                    console.WriteLine(service + ": " + func.GetValue("RESPUESTA").ToString());
                                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Servicio", service + ": " + func.GetValue("RESPUESTA").ToString(), root.Subject);
                                    respFinal = respFinal + "\\n" + "Crear Servicio" + service + ": " + func.GetValue("RESPUESTA").ToString();

                                    if (res1.Contains("Error"))
                                    { valData = false; }
                                    #endregion


                                }
                                catch (Exception ex)
                                {
                                    resFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                                    console.WriteLine("Finishing process " + resFailure);
                                    res1 = res1 + service + ": " + ex.ToString() + "<br>";
                                    resFailure = ex.ToString();
                                    valData = false;
                                }

                                #endregion

                            }

                        } //for de cada fila del excel

                       
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
                    DM.ChangeStateDM(root.IdGestionDM, res1 + "<br>" + resFailure, "4"); //error
                    mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + res1 + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, new string[] { "hlherrera@gbm.net" });
                }
                else if (returnRequest == true)
                {
                    console.WriteLine("Devolviendo solicitud");
                    DM.ChangeStateDM(root.IdGestionDM, res2, "5"); //rechazado
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Servicios:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + res2);
                }
                else
                {
                    //finalizar solicitud
                    DM.ChangeStateDM(root.IdGestionDM, res1, "3"); //finalizado
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Servicios:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + res1);
                }

                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                string[] cc = { "hlherrera@gbm.net" };
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //error

                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }
    }
}
