using DataBotV5.App.Global;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Projects.MasterData;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Mail;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Projects.SpecialBidForms;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;
using Newtonsoft.Json.Linq;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace DataBotV5.Automation.DM.Equipment
{
    /// <summary>
    ///Clase DM Automation encargada de la creación de equipos en GBM. 
    /// </summary>
    class EquipmentSS
    {
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        WebInteraction web = new WebInteraction();
        Log log = new Log();
        WebexTeams wt = new WebexTeams();
        Stats stats = new Stats();
        SbForm sbform = new SbForm();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        SapVariants sap = new SapVariants();
        MsExcel ms = new MsExcel();

        public string response = "";
        public bool failure = false;
        public string responseFailure = "";
        string month = "", year = "", day = "";
        string material;
        string erpMand = "ERP";

        string respFinal = "";


        public void Main()
        {

            string response = DM.GetManagement("4"); //EQUIPOS
            if (!String.IsNullOrEmpty(response) && response != "ERROR")
            {
                console.WriteLine("Procesando...");
                EquipmentProcessing();
                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
            else if (response == "ERROR")
            {
                string[] cc = { "smarin@gbm.net" };
                DM.ChangeStateDM(root.IdGestionDM, "Error al leer la solicitud", "4"); //ERROR
                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>Error al leer la solicitud<br>", new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }

        }
        public void EquipmentProcessing()
        {
            try
            {

                #region extraer datos generales (cada clase ya que es data muy personal de la solicitud) 
                JArray generalData = JArray.Parse(root.datagDM);
                for (int i = 0; i < generalData.Count; i++)
                {
                    JObject fila = JObject.Parse(generalData[i].ToString());
                    string factor = fila["sendingCountryCode"].Value<string>();
                }
                #endregion

                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                if (root.metodoDM == "1") //LINEAL
                {
                    #region Variables Privadas ProcesarEquipos
                    console.WriteLine("Abriendo excel y validando");
                    JArray requests = JArray.Parse(root.requestDetails);

                    int rows = requests.Count;
                    string returnMsg = "";
                    bool validateData = true;
                    response = "";

                    #endregion

                    if (rows > 100)
                    {
                        int filas = 0;
                        for (int e = 0; e <= rows; e++)
                        {
                            JObject fila = JObject.Parse(requests[e].ToString());
                            material = fila["materialId"].Value<string>();
                            if (material == "")
                            { break; }
                            filas++;
                        }
                        if (filas > 100)
                        {
                            returnMsg = "Para la creación masiva de datos, favor enviar la gestión directamente a Datos Maestros";
                            DM.ChangeStateDM(root.IdGestionDM, returnMsg, "5"); //rechazado
                            wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + returnMsg);

                            return;
                        }
                    }

                    for (int i = 0; i < rows; i++)
                    {
                        JObject fila = JObject.Parse(requests[i].ToString());
                        string descripcion = fila["description"].Value<string>();
                        string soldto = fila["soldToParty"].Value<string>();
                        string shipto = fila["shipToParty"].Value<string>();
                        string fecha = fila["instalationDate"].Value<string>();
                        string warr = fila["endOfWarranty"].Value<string>();
                        material = fila["materialId"].Value<string>();
                        string serie = fila["equipmentSeries"].Value<string>();
                        string pais = fila["companyCodeCode"].Value<string>();
                        string asset = fila["asset"].Value<string>();
                        string placa = fila["plate"].Value<string>();

                        if (material != "")
                        {
                            if (descripcion == "")
                            {
                                response = response + "Error en la linea " + (i + 1) + " la descripción no puede ser nula." + "<br>";
                                validateData = false;
                            }
                            if (soldto == "")
                            {
                                response = response + "Error en la linea " + (i + 1) + " el Sold to Party no puede ser nulo." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                if (soldto.Length < 8)
                                {
                                    response = response + "Error en la linea " + (i + 1) + " favor verificar el Sold to Party." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (shipto != "")
                            {
                                if (shipto.Length < 8)
                                {
                                    response = response + "Error en la linea " + (i + 1) + " favor verificar el Ship to Party." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (fecha == "")
                            {
                                response = response + "Error en la linea " + (i + 1) + " la fecha no puede ser nula." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                try
                                {
                                    if (fecha.Contains("-"))
                                    {
                                        string[] DMY = fecha.Split(new char[1] { '-' });

                                        day = int.Parse(DMY[0]).ToString();

                                        if (day.Length == 1)
                                            day = "0" + day;
                                        else if (day.Length == 4)
                                            day = int.Parse(DMY[2]).ToString();
                                        else
                                            day = int.Parse(DMY[0]).ToString();

                                        if (day.Length == 1)
                                            day = "0" + day;

                                        month = int.Parse(DMY[1]).ToString();
                                        if (month.Length == 1)
                                            month = "0" + month;

                                        year = int.Parse(DMY[2]).ToString();
                                        if (year.Length == 4)
                                            fecha = day + "/" + month + "/" + int.Parse(DMY[2]);
                                        else
                                            fecha = day + "/" + month + "/" + int.Parse(DMY[0]);
                                    }
                                    DateTime fechainsta = DateTime.Parse(fecha);
                                    if (fechainsta > DateTime.Today)
                                    {
                                        response = response + "Error en la linea " + (i + 1) + " la fecha no puede ser a futuro." + "<br>";
                                        validateData = false;
                                    }
                                }
                                catch (Exception) { }


                            }
                            if (material == "")
                            {
                                response = response + "Error en la linea " + (i + 1) + " el material no puede ser nulo." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                if (material.Length > 18)
                                {
                                    response = response + "Error en la linea " + (i + 1) + " favor revisar el material." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (serie == "")
                            {
                                response = response + "Error en la linea " + (i + 1) + " la serie no puede ser nula." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                if (serie.Length > 18)
                                {
                                    response = response + "Error en la linea " + (i + 1) + " favor revisar la serie." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (asset != "")
                            {
                                if (asset.Length != 12)
                                {
                                    response = response + "Error en la linea " + (i + 1) + " favor verificar el asset. Debe de contener 12 caracteres" + "<br>";
                                    validateData = false;
                                }
                            }
                            if (asset != "" && pais == "")
                            {
                                response = response + "Error en la linea " + (i + 1) + " favor ingresar la organización país, para el asset." + "<br>";
                                validateData = false;
                            }
                            if (pais != "")
                            {
                                if (pais.Substring(0, 2) != "GB")
                                {
                                    response = response + "Error en la linea " + (i + 1) + " favor verificar la organización país." + "<br>";
                                    validateData = false;
                                }
                            }
                        }
                    }

                    if (validateData == true)
                    {
                        //Todas las validaciones de las lineas son correctas
                        //Ejecute el proceso de creación
                        for (int i = 0; i < rows; i++)
                        {

                            JObject fila = JObject.Parse(requests[i].ToString());

                            string descripcion = fila["description"].Value<string>();
                            string soldto = fila["soldToParty"].Value<string>();
                            string shipto = fila["shipToParty"].Value<string>();
                            string fecha = fila["instalationDate"].Value<string>();
                            string warr = fila["endOfWarranty"].Value<string>();
                            string material = fila["materialId"].Value<string>();
                            string serie = fila["equipmentSeries"].Value<string>();
                            string pais = fila["companyCodeCode"].Value<string>();
                            string asset = fila["asset"].Value<string>();
                            string placa = fila["plate"].Value<string>();


                            if (pais != "" && asset == "")
                                pais = "";
                            if (placa.ToUpper() == "N/A" || placa.ToUpper() == "NA")
                                placa = "";
                            if (descripcion != "")
                            {
                                if (!string.IsNullOrWhiteSpace(fecha))
                                {
                                    fecha = DateTime.Parse(fecha).ToString("yyyy-MM-dd");

                                }
                                if (!string.IsNullOrWhiteSpace(warr))
                                {
                                    warr = DateTime.Parse(warr).ToString("yyyy-MM-dd");

                                }
                                #region validar data
                                //if (fecha.Contains("/"))
                                //{
                                //    var DMY = fecha.Split(new char[1] { '/' });
                                //    day = int.Parse(DMY[0]).ToString();
                                //    year = int.Parse(DMY[2]).ToString();
                                //    if (day.Length == 1)
                                //        day = "0" + day;
                                //    else if (day.Length == 4)
                                //    {
                                //        year = int.Parse(DMY[0]).ToString();
                                //        day = int.Parse(DMY[2]).ToString();
                                //        if (day.Length == 1)
                                //        { day = "0" + day; }
                                //    }
                                //    month = int.Parse(DMY[1]).ToString();
                                //    if (month.Length == 1)
                                //        month = "0" + month;
                                //    fecha = year + "-" + month + "-" + day;
                                //}
                                //else if (fecha.Contains("."))
                                //{
                                //    var DMY = fecha.Split(new char[1] { '.' });
                                //    day = int.Parse(DMY[0]).ToString();
                                //    year = int.Parse(DMY[2]).ToString();
                                //    if (day.Length == 1)
                                //        day = "0" + day;
                                //    else if (day.Length == 4)
                                //    {
                                //        year = int.Parse(DMY[0]).ToString();
                                //        day = int.Parse(DMY[2]).ToString();
                                //        if (day.Length == 1)
                                //            day = "0" + day;
                                //    }
                                //    month = int.Parse(DMY[1]).ToString();
                                //    if (month.Length == 1)
                                //        month = "0" + month;
                                //    fecha = year + "-" + month + "-" + day;
                                //}
                                //else if (fecha.Contains("-"))
                                //{
                                //    var DMY = fecha.Split(new char[1] { '-' });

                                //    day = int.Parse(DMY[0]).ToString();
                                //    year = int.Parse(DMY[2]).ToString();
                                //    if (day.Length == 1)
                                //        day = "0" + day;
                                //    else if (day.Length == 4)
                                //    {
                                //        year = int.Parse(DMY[0]).ToString();
                                //        day = int.Parse(DMY[2]).ToString();
                                //        if (day.Length == 1)
                                //        { day = "0" + day; }
                                //    }

                                //    month = int.Parse(DMY[1]).ToString();
                                //    if (month.Length == 1)
                                //        month = "0" + month;

                                //    fecha = year + "-" + month + "-" + day;
                                //}

                                //if (warr.Contains("/"))
                                //{
                                //    var DMY = warr.Split(new char[1] { '/' });
                                //    day = int.Parse(DMY[0]).ToString();
                                //    year = int.Parse(DMY[2]).ToString();
                                //    if (day.Length == 1)
                                //        day = "0" + day;
                                //    else if (day.Length == 4)
                                //    {
                                //        year = int.Parse(DMY[0]).ToString();
                                //        day = int.Parse(DMY[2]).ToString();
                                //        if (day.Length == 1)
                                //            day = "0" + day;
                                //    }
                                //    month = int.Parse(DMY[1]).ToString();
                                //    if (month.Length == 1)
                                //        month = "0" + month;
                                //    warr = year + "-" + month + "-" + day;
                                //}
                                //else if (warr.Contains("."))
                                //{
                                //    string[] DMY = warr.Split(new char[1] { '.' });
                                //    day = int.Parse(DMY[0]).ToString();
                                //    year = int.Parse(DMY[2]).ToString();
                                //    if (day.Length == 1)
                                //        day = "0" + day;
                                //    else if (day.Length == 4)
                                //    {
                                //        year = int.Parse(DMY[0]).ToString();
                                //        day = int.Parse(DMY[2]).ToString();
                                //        if (day.Length == 1)
                                //            day = "0" + day;
                                //    }
                                //    month = int.Parse(DMY[1]).ToString();
                                //    if (month.Length == 1)
                                //        month = "0" + month;
                                //    warr = year + "-" + month + "-" + day;
                                //}
                                //else if (warr.Contains("-"))
                                //{
                                //    string[] DMY = warr.Split(new char[1] { '-' });

                                //    day = int.Parse(DMY[0]).ToString();
                                //    year = int.Parse(DMY[2]).ToString();
                                //    if (day.Length == 1)
                                //        day = "0" + day;
                                //    else if (day.Length == 4)
                                //    {
                                //        year = int.Parse(DMY[0]).ToString();
                                //        day = int.Parse(DMY[2]).ToString();
                                //        if (day.Length == 1)
                                //            day = "0" + day;
                                //    }

                                //    month = int.Parse(DMY[1]).ToString();
                                //    if (month.Length == 1)
                                //        month = "0" + month;

                                //    warr = year + "-" + month + "-" + day;
                                //}

                                if (fecha != "")
                                {

                                    int month = int.Parse(DateTime.Parse(fecha).Month.ToString());
                                    if ((month > 12))
                                    {
                                        response = response + "Mes del Star Date no es valido, entre a la solicitud y modifique" + "<br>";
                                        continue;
                                    }
                                }
                                if (soldto != "")
                                {
                                    if (soldto.Substring(0, 2) != "00")
                                        soldto = ("00" + soldto);
                                }
                                if (shipto == "")
                                    shipto = soldto;
                                else
                                {
                                    if (shipto.Substring(0, 2) != "00")
                                        shipto = ("00" + shipto);
                                }

                                descripcion = val.RemoveSpecialChars(descripcion, 1);
                                descripcion = descripcion.ToUpper();


                                #endregion

                                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>
                                    {
                                        ["DESCRIPTION"] = descripcion.ToUpper(),
                                        ["SOLD_TO_PARTY"] = soldto,
                                        ["SHIP_TO_PARTY"] = shipto,
                                        ["FECHA_INSTALACION"] = fecha,
                                        ["MATERIAL"] = material.ToUpper(),
                                        ["SERIAL"] = serie.ToUpper(),
                                        ["ASSET"] = asset.ToUpper(),
                                        ["PLACA"] = placa.ToUpper(),
                                        ["PAIS"] = pais.ToUpper(),
                                        ["FIN_GARANTIA"] = warr
                                    };

                                    IRfcFunction func = sap.ExecuteRFC(erpMand, "ZDM_RPA_CE_001", parameters);

                                    #region Procesar Salidas del FM
                                    response = response + func.GetValue("RESPONSE").ToString() + "<br>";

                                    if (response.Contains("El material no existe en SAP"))
                                        validateData = false;

                                    console.WriteLine(func.GetValue("RESPONSE").ToString());
                                     log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Equipo", serie + " - " + material + ":" + func.GetValue("RESPONSE").ToString(), root.Subject);
                                    respFinal = respFinal + "\\n" + "Crear Equipo" + serie + " - " + material + ":" + func.GetValue("RESPONSE").ToString();

                                    if (response.ToLower().Contains("error"))
                                        validateData = false;
                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    responseFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                                    console.WriteLine("Finishing process " + responseFailure);
                                    response = response + material + " - " + serie + ": " + ex.ToString() + "<br>" + ex.StackTrace + "<br>";
                                    responseFailure = ex.ToString();
                                    validateData = false;
                                }
                            }
                        }
                        console.WriteLine("Finalizando solicitud");
                        if (validateData == false)
                        {
                            if (response.Contains("El material no existe en SAP"))
                            {
                                wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazado, con el siguiente resultado: <br><br> " + response);

                                response = response.Replace("<br>", " - ");
                                DM.ChangeStateDM(root.IdGestionDM, response, "5"); //RECHAZADO
                            }
                            else
                            {
                                //enviar email de repuesta de error a datos maestros
                                string[] cc = { "smarin@gbm.net" };
                                DM.ChangeStateDM(root.IdGestionDM, "Gestion: " + root.IdGestionDM + " " + response + " " + responseFailure, "4"); //ERROR
                                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + response + "<br>" + responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
                            }
                        }
                        else
                        {
                            //finalizar solicitud
                            DM.ChangeStateDM(root.IdGestionDM, response, "3"); //FINALIZADO
                            wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + response);
                        }

                    }
                    else
                    {
                        wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + response);
                        response = response.Replace("<br>", " - ");
                        //El archivo no cumple con la info entonces lo devuelve
                        DM.ChangeStateDM(root.IdGestionDM, response, "5"); //RECHAZADO
                    }


                }
                else //MASIVO
                {
                    #region Variables Privadas ProcesarEquipos
                    console.WriteLine("Abriendo excel y validando");
                    string adjunto = root.ExcelFile; //ya viene 
                    int rows;
                    string returnMsg = "";
                    string validar_strc;
                    bool validateData = true;

                    DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + adjunto);
                    #endregion

                    if (xlWorkSheet.Rows.Count > 100)
                    {
                        int filas = 0;
                        foreach (DataRow item in xlWorkSheet.Rows)
                        {
                            material = item["PRODUCTO/MATERIAL"].ToString().Trim();
                            if (material == "")
                                break;
                            filas++;
                        }
                        if (filas > 100)
                        {
                            returnMsg = "Para la creación masiva de datos, favor enviar la gestión directamente a Datos Maestros";
                            DM.ChangeStateDM(root.IdGestionDM, returnMsg, "5"); //RECHAZADO
                            wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazado, con el siguiente resultado: <br><br> " + returnMsg);

                            return;
                        }
                    }


                    //Plantilla correcta, continue las validaciones
                    foreach (DataRow item in xlWorkSheet.Rows)
                    {
                        int i = xlWorkSheet.Rows.IndexOf(item);
                        string descripcion = item["DESCRIPCION (40 CARACTERES MAXIMO)"].ToString().Trim();
                        string soldto = item["ID SOLD TO PARTY"].ToString().Trim();
                        string shipto = item["ID SHIP TO PARTY"].ToString().Trim();
                        string fecha = item["FECHA INSTALACION (AAAA-MM-DD)"].ToString().Trim();
                        string warr = item["FECHA FIN GARANTIA (AAAA-MM-DD)"].ToString().Trim();
                        material = item["PRODUCTO/MATERIAL"].ToString().Trim();
                        string serie = item["SERIE"].ToString().Trim();
                        string pais = item["Company Code"].ToString().Trim();
                        string asset = item["ASSET"].ToString().Trim();
                        string placa = item["xPLACA"].ToString().Trim();

                        if (material != "")
                        {
                            if (descripcion == "")
                            {
                                response = response + "Error en la linea " + i + " la descripcion no puede ser nula." + "<br>";
                                validateData = false;
                            }
                            if (soldto == "")
                            {
                                response = response + "Error en la linea " + i + " el Sold to Party no puede ser nulo." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                if (soldto.Length < 8)
                                {
                                    response = response + "Error en la linea " + i + " favor verificar el Sold to Party." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (shipto != "")
                            {
                                if (shipto.Length < 8)
                                {
                                    response = response + "Error en la linea " + i + " favor verificar el Ship to Party." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (fecha == "")
                            {
                                response = response + "Error en la linea " + i + " la fecha no puede ser nula." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                try
                                {
                                    if (fecha.Contains("-"))
                                    {
                                        string[] DMY = fecha.Split(new char[1] { '-' });

                                        day = int.Parse(DMY[0]).ToString();
                                        if (day.Length == 1)
                                            day = "0" + day;
                                        else if (day.Length == 4)
                                            day = int.Parse(DMY[2]).ToString();
                                        else
                                            day = int.Parse(DMY[0]).ToString();

                                        if (day.Length == 1)
                                            day = "0" + day;

                                        month = int.Parse(DMY[1]).ToString();
                                        if (month.Length == 1)
                                            month = "0" + month;

                                        year = int.Parse(DMY[2]).ToString();
                                        if (year.Length == 4)
                                            fecha = day + "/" + month + "/" + int.Parse(DMY[2]);
                                        else
                                            fecha = day + "/" + month + "/" + int.Parse(DMY[0]);
                                    }
                                    DateTime fechainsta = DateTime.Parse(fecha);
                                    if (fechainsta > DateTime.Today)
                                    {
                                        response = response + "Error en la linea " + i + " la fecha no puede ser a futuro." + "<br>";
                                        validateData = false;
                                    }
                                }
                                catch (Exception) { }


                            }
                            if (material == "")
                            {
                                response = response + "Error en la linea " + i + " el material no puede ser nulo." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                if (material.Length > 18)
                                {
                                    response = response + "Error en la linea " + i + " favor revisar el material." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (serie == "")
                            {
                                response = response + "Error en la linea " + i + " la serie no puede ser nula." + "<br>";
                                validateData = false;
                            }
                            else
                            {
                                if (serie.Length > 18)
                                {
                                    response = response + "Error en la linea " + i + " favor revisar la serie." + "<br>";
                                    validateData = false;
                                }
                            }
                            if (asset != "")
                            {
                                if (asset.Length != 12)
                                {
                                    response = response + "Error en la linea " + i + " favor verificar el asset. Debe de contener 12 caracteres" + "<br>";
                                    validateData = false;
                                }
                            }
                            if (asset != "" && pais == "")
                            {
                                response = response + "Error en la linea " + i + " favor ingresar la organizacion pais, para el asset." + "<br>";
                                validateData = false;
                            }
                            if (pais != "")
                            {
                                if (pais.Substring(0, 2) != "GB")
                                {
                                    response = response + "Error en la linea " + i + " favor verificar la organizacion pais." + "<br>";
                                    validateData = false;
                                }
                            }
                        }
                    }

                    if (validateData == true)
                    {
                        // Todas las validaciones de las lineas son correctas
                        //Ejecute el proceso de creacion
                        foreach (DataRow item in xlWorkSheet.Rows)
                        {
                            int i = xlWorkSheet.Rows.IndexOf(item);
                            string descripcion = item["DESCRIPCION (40 CARACTERES MAXIMO)"].ToString().Trim();
                            string soldto = item["ID SOLD TO PARTY"].ToString().Trim();
                            string shipto = item["ID SHIP TO PARTY"].ToString().Trim();
                            string fecha = item["FECHA INSTALACION (AAAA-MM-DD)"].ToString().Trim();
                            string warr = item["FECHA FIN GARANTIA (AAAA-MM-DD)"].ToString().Trim();
                            string material = item["PRODUCTO/MATERIAL"].ToString().Trim();
                            string serie = item["SERIE"].ToString().Trim();
                            string pais = item["Company Code"].ToString().Trim();
                            string asset = item["ASSET"].ToString().Trim();
                            string placa = item["xPLACA"].ToString().Trim();




                            if (pais != "" && asset == "")
                                pais = "";
                            if (placa.ToUpper() == "N/A" || placa.ToUpper() == "NA")
                                placa = "";
                            if (descripcion != "")
                            {
                                #region validar data
                                if (fecha.Contains("/"))
                                {
                                    var DMY = fecha.Split(new char[1] { '/' });
                                    day = int.Parse(DMY[0]).ToString();
                                    year = int.Parse(DMY[2]).ToString();
                                    if (day.Length == 1)
                                        day = "0" + day;
                                    else if (day.Length == 4)
                                    {
                                        year = int.Parse(DMY[0]).ToString();
                                        day = int.Parse(DMY[2]).ToString();
                                        if (day.Length == 1)
                                        { day = "0" + day; }
                                    }
                                    month = int.Parse(DMY[1]).ToString();
                                    if (month.Length == 1)
                                        month = "0" + month;
                                    fecha = year + "-" + month + "-" + day;
                                }
                                else if (fecha.Contains("."))
                                {
                                    var DMY = fecha.Split(new char[1] { '.' });
                                    day = int.Parse(DMY[0]).ToString();
                                    year = int.Parse(DMY[2]).ToString();
                                    if (day.Length == 1)
                                        day = "0" + day;
                                    else if (day.Length == 4)
                                    {
                                        year = int.Parse(DMY[0]).ToString();
                                        day = int.Parse(DMY[2]).ToString();
                                        if (day.Length == 1)
                                            day = "0" + day;
                                    }
                                    month = int.Parse(DMY[1]).ToString();
                                    if (month.Length == 1)
                                        month = "0" + month;
                                    fecha = year + "-" + month + "-" + day;
                                }
                                else if (fecha.Contains("-"))
                                {
                                    string[] DMY = fecha.Split(new char[1] { '-' });

                                    day = int.Parse(DMY[0]).ToString();
                                    year = int.Parse(DMY[2]).ToString();
                                    if (day.Length == 1)
                                        day = "0" + day;
                                    else if (day.Length == 4)
                                    {
                                        year = int.Parse(DMY[0]).ToString();
                                        day = int.Parse(DMY[2]).ToString();
                                        if (day.Length == 1)
                                            day = "0" + day;
                                    }

                                    month = int.Parse(DMY[1]).ToString();
                                    if (month.Length == 1)
                                        month = "0" + month;

                                    fecha = year + "-" + month + "-" + day;
                                }
                                if (warr.Contains("/"))
                                {
                                    var DMY = warr.Split(new char[1] { '/' });
                                    day = int.Parse(DMY[0]).ToString();
                                    year = int.Parse(DMY[2]).ToString();
                                    if (day.Length == 1)
                                        day = "0" + day;
                                    else if (day.Length == 4)
                                    {
                                        year = int.Parse(DMY[0]).ToString();
                                        day = int.Parse(DMY[2]).ToString();
                                        if (day.Length == 1)
                                            day = "0" + day;
                                    }
                                    month = int.Parse(DMY[1]).ToString();
                                    if (month.Length == 1)
                                        month = "0" + month;
                                    warr = year + "-" + month + "-" + day;
                                }
                                else if (warr.Contains("."))
                                {
                                    var DMY = warr.Split(new char[1] { '.' });
                                    day = int.Parse(DMY[0]).ToString();
                                    year = int.Parse(DMY[2]).ToString();
                                    if (day.Length == 1)
                                        day = "0" + day;
                                    else if (day.Length == 4)
                                    {
                                        year = int.Parse(DMY[0]).ToString();
                                        day = int.Parse(DMY[2]).ToString();
                                        if (day.Length == 1)
                                            day = "0" + day;
                                    }
                                    month = int.Parse(DMY[1]).ToString();
                                    if (month.Length == 1)
                                        month = "0" + month;
                                    warr = year + "-" + month + "-" + day;
                                }
                                else if (warr.Contains("-"))
                                {
                                    string[] DMY = warr.Split(new char[1] { '-' });

                                    day = int.Parse(DMY[0]).ToString();
                                    year = int.Parse(DMY[2]).ToString();
                                    if (day.Length == 1)
                                        day = "0" + day;
                                    else if (day.Length == 4)
                                    {
                                        year = int.Parse(DMY[0]).ToString();
                                        day = int.Parse(DMY[2]).ToString();
                                        if (day.Length == 1)
                                            day = "0" + day;
                                    }

                                    month = int.Parse(DMY[1]).ToString();
                                    if (month.Length == 1)
                                        month = "0" + month;

                                    warr = year + "-" + month + "-" + day;


                                }
                                if (fecha != "")
                                {

                                    int month = int.Parse(this.month);
                                    if ((month > 12))
                                    {
                                        response = response + "Mes del Star Date no es valido, entre a la solicitud y modifique" + "<br>";
                                        continue;
                                    }
                                }
                                if (soldto != "")
                                {
                                    if ((soldto.Substring(0, 2) != "00"))
                                        soldto = ("00" + soldto);
                                }
                                if (shipto == "")
                                    shipto = soldto;
                                else
                                {
                                    if ((shipto.Substring(0, 2) != "00"))
                                        shipto = ("00" + shipto);
                                }

                                descripcion = val.RemoveSpecialChars(descripcion, 1);
                                descripcion = descripcion.ToUpper();


                                #endregion
                                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>
                                    {
                                        ["DESCRIPTION"] = descripcion.ToUpper(),
                                        ["SOLD_TO_PARTY"] = soldto,
                                        ["SHIP_TO_PARTY"] = shipto,
                                        ["MATERIAL"] = material.ToUpper(),
                                        ["SERIAL"] = serie.ToUpper(),
                                        ["ASSET"] = serie.ToUpper(),
                                        ["PLACA"] = placa.ToUpper(),
                                        ["PAIS"] = pais.ToUpper(),

                                        ["FECHA_INSTALACION"] = fecha,
                                        ["FIN_GARANTIA"] = warr
                                    };

                                    IRfcFunction func;
                                    try
                                    {
                                        func = sap.ExecuteRFC(erpMand, "ZDM_RPA_CE_001", parameters);
                                    }
                                    catch (Exception ex)
                                    {
                                        responseFailure = responseFailure + "<br><br>" + "Linea: " + i + ": " + ex.Message;
                                        validateData = false;
                                        continue;
                                    }

                                    #region Procesar Salidas del FM
                                    response = response + func.GetValue("RESPONSE").ToString() + "<br>";
                                    if (response.Contains("El material no existe en SAP"))
                                        validateData = false;
                                    console.WriteLine(func.GetValue("RESPONSE").ToString());
                                     log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Equipo", serie + " - " + material + ":" + func.GetValue("RESPONSE").ToString(), root.Subject);
                                    respFinal = respFinal + "\\n" + "Crear Equipo" + serie + " - " + material + ":" + func.GetValue("RESPONSE").ToString();

                                    if (response.ToLower().Contains("error"))
                                        validateData = false;
                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    responseFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                                    console.WriteLine("Finishing process " + responseFailure);
                                    response = response + material + " - " + serie + ": " + ex.ToString() + "<br>" + ex.StackTrace + "<br>";
                                    responseFailure = ex.ToString();
                                    validateData = false;
                                }
                            }
                        }
                        console.WriteLine("Finalizando solicitud");
                        if (!validateData)
                        {
                            if (response.Contains("El material no existe en SAP"))
                            {
                                wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + response);
                                response = response.Replace("<br>", " - ");
                                DM.ChangeStateDM(root.IdGestionDM, response, "5"); //RECHAZADO
                            }
                            else
                            {
                                //enviar email de repuesta de error a datos maestros
                                DM.ChangeStateDM(root.IdGestionDM, "ERROR", "4"); //ERROR
                                string[] cc = { "smarin@gbm.net" };
                                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + response + "<br>" + responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
                            }
                        }
                        else
                        {
                            //finalizar solicitud
                            DM.ChangeStateDM(root.IdGestionDM, response, "3"); //FINALIZADO
                            wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + response);
                        }

                    }
                    else
                    {
                        wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificacion de gestion de Equipos:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + response);
                        //mail.EnviarCorreo(response, root.Solicitante, root.Subject, 1, resp_type: 2);
                        response = response.Replace("<br>", " - ");
                        //El archivo no cumple con la info entonces lo devuelve
                        DM.ChangeStateDM(root.IdGestionDM, response, "5"); //RECHAZADO
                    }



                }



                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                string[] cc = { "smarin@gbm.net" };
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //ERROR
                mail.SendHTMLMail("Gestion: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }
    }
}
