using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.Projects.Materials;
using DataBotV5.Data.Projects.MasterData;
using DataBotV5.Logical.Processes;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Webex;
using DataBotV5.Data.Process;
using DataBotV5.Logical.Mail;

using DataBotV5.Logical.Web;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System;
using DataBotV5.Logical.MicrosoftTools;
using System.Data;

namespace DataBotV5.Automation.DM.Materials
{
    /// <summary>
    /// Clase DM Automation encargada de creación de materiales en datos maestros.
    /// </summary>
    class MaterialsSS
    {
        ProcessInteraction proc = new ProcessInteraction();
        
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        WebInteraction web = new WebInteraction();
        ProcessAdmin padmin = new ProcessAdmin();
        MasterDataSqlSS dm = new MasterDataSqlSS();
        MaterialsSel selMat = new MaterialsSel();
        Credentials cred = new Credentials();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        WebexTeams wt = new WebexTeams();
        Log log = new Log();
        Stats stats = new Stats();
        Rooting root = new Rooting();
        MsExcel ms = new MsExcel();

        string matType = "", material = "", desc = "", materialGroup = "", itemCat = "", gm1 = "", gm2 = "", serial = "", price = "", warr = "";

        string erpMand = "ERP";
        string crmMand = "CRM";

        string respFinal = "";

        /// <summary>
        /// metodo de DM web page
        /// </summary>
        public void Main()
        {
            string respuesta = dm.GetManagement("1"); //MATERIALES
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessMaterials();
                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void ProcessMaterials()
        {
            try
            {
                bool retRequest = false, validateData = true;

                int lenght, rows;

                string
                       retMsg = "",
                       resFailure = "",
                       resLog = "",
                       response1 = "",
                       response2 = "",
                       valData,
                       baw = "",
                       idBaw = "";


                #region extraer datos generales (cada clase ya que es data muy personal de la solicitud)
                console.WriteLine("Extraer datos generales...");
                JArray DG = JArray.Parse(root.datagDM);
                string materialGroup = "";
                for (int i = 0; i < DG.Count; i++)
                {
                    JObject fila = JObject.Parse(DG[i].ToString());
                    materialGroup = fila["materialGroupCode"].Value<string>();
                    baw = fila["bawCode"].Value<string>(); //si o no
                    idBaw = fila["bawManagement"].Value<string>(); //id de BAW
                }
                #endregion

                // string materialGroup = root.factorDM;

                if (root.metodoDM == "1") //LINEAL
                {
                    JArray gestiones = JArray.Parse(root.requestDetails);
                    for (int i = 0; i < gestiones.Count; i++)
                    {
                        JObject fila = JObject.Parse(gestiones[i].ToString());

                        string matId = fila["idMaterial"].Value<string>().Trim();
                        string matDesc = fila["description"].Value<string>().Trim().ToUpper();
                        string matType = fila["materialTypeCode"].Value<string>().Trim();
                        string matItemCat = fila["positionGroupCode"].Value<string>().Trim();
                        string matGM1 = fila["materialGroup1Code"].Value<string>().Trim();
                        string matGM2 = fila["materialGroup2Code"].Value<string>().Trim();
                        string matPrice = fila["price"].Value<string>().Trim();
                        string matWarr = fila["warrantyTypeCode"].Value<string>().Trim();
                        string matSerial = fila["serializableCode"].Value<string>().Trim();

                        #region validación de datos
                        console.WriteLine("Validando...");

                        #region validaciones Id
                        matId = matId.Replace("á", "a");
                        matId = matId.Replace("é", "e");
                        matId = matId.Replace("í", "i");
                        matId = matId.Replace("ó", "o");
                        matId = matId.Replace("ú", "u");
                        matId = matId.Replace("ñ", "n");
                        matId = matId.ToUpper();

                        if (matId == "")
                        {
                            retMsg = "Ingrese el material";
                            response2 = response2 + retMsg + "<br>";
                            retRequest = true;
                            continue;
                        }
                        if (matId.Length > 18)
                        {
                            retMsg = "El código del material debe de ser menor a 18 caracteres: " + matId;
                            retRequest = true;
                            response2 = response2 + retMsg + "<br>";
                            continue;
                        }

                        #endregion

                        #region validaciones Price
                        if (matPrice == "Vacio")
                            matPrice = "";

                        if (matPrice != "")
                        {
                            matPrice = matPrice.Replace("$", "");
                            if (matPrice.Length > 3)
                            {
                                if (matPrice.Substring(0, 3).Substring(matPrice.Substring(0, 3).Length - 1, 1) == "," && matPrice.Substring((matPrice.Length - 3)).Substring(0, 1) != "," || matPrice.Substring(0, 2).Substring(matPrice.Substring(0, 2).Length - 1, 1) == "," && matPrice.Substring(matPrice.Length - 3, 1).Substring(0, 1) != ",")
                                    matPrice = matPrice.Replace(",", "");

                                if ((matPrice.Substring(matPrice.Length - 3).Substring(0, 1) == ".") || (matPrice.Substring(matPrice.Length - 2).Substring(0, 1) == "."))
                                {
                                    // ejemplo 100,000.34 ---- 100000,34
                                    matPrice = matPrice.Replace(",", "");
                                    matPrice = matPrice.Replace(".", ",");
                                }
                            }

                            if (matPrice == "9999999" || matPrice == "999999.99" || matPrice == "999,999")
                                matPrice = "999999";

                            if (matPrice == "0" || matPrice == "0,00" || matPrice == "0.00")
                                matPrice = "";
                        }

                        #endregion

                        #region validaciones GM2
                        if (matGM2 == "Vacio")
                            matGM2 = "";
                        #endregion

                        #region validaciones item cat
                        if (matItemCat == "FEATURE")
                            matItemCat = "FEAT";
                        #endregion

                        #region validaciones Warr
                        if ((matWarr == "N/A") || (matWarr == "N/A - No Aplica"))
                            matWarr = "";
                        if (matWarr == "Vacio")
                            matWarr = "";
                        #endregion

                        #region validaciones GM1
                        if (matGM1.Length < 2 && matGM1 != "")
                            matGM1 = "0" + matGM1;
                        #endregion

                        #region validaciones Serial
                        if (matSerial == "S")
                            matSerial = "SI";
                        if (matSerial == "N")
                            matSerial = "NO";
                        #endregion

                        #region validaciones Mat Type
                        if (matType == "ZREP")
                            matType = "ZHRW";
                        #endregion

                        #region validaciones Desc
                        matDesc = val.RemoveSpecialChars(matDesc, 1);
                        #endregion


                        if (matId.Substring(0, 3) == "800" && materialGroup.Substring(0, 2) == "40")
                        {
                            retMsg = "El material: " + matId + " es un contrato por favor hacer la solicitud en el formulario de Servicios";
                            response2 = response2 + retMsg + "<br>";
                            retRequest = true;
                            continue;
                        }

                        if (materialGroup.Substring(0, 3) == "103" || materialGroup == "201010120")
                            matWarr = "";

                        if (matSerial == "SI" && matWarr == "N/A" && materialGroup.Substring(0, 3) != "103" && materialGroup != "201010120" || matSerial == "SI" && matWarr == "" && materialGroup.Substring(0, 3) != "103" && materialGroup != "201010120")
                        {
                            retMsg = "El material: " + matId + " es serializable y no tiene garantia";
                            response2 = response2 + retMsg + "<br>";
                            retRequest = true;
                            continue;
                        }

                        if (matId.Substring(matId.Length - 3, 3) == "_NI" && materialGroup.Substring(0, 3) == "402")
                        {
                            retMsg = "Por favor enviar esta solicitud directamente a datos maestros, ya que son materiales de PS";
                            response2 = response2 + retMsg + "<br>";
                            retRequest = true;
                            continue;
                        }

                        if (materialGroup == "1040109")
                            matType = "ZSFW";

                        #endregion validación de datos

                        #region SAP

                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {
                            Dictionary<string, string> zdmCreateMatParameters = new Dictionary<string, string>
                            {
                                ["TIPO_MAT"] = matType,
                                ["MATERIAL"] = matId,
                                ["GM"] = materialGroup,
                                ["ITEMCAT"] = matItemCat,
                                ["GM1"] = matGM1,
                                ["DESCRIPCION"] = matDesc,
                                ["SERIALIZABLE"] = matSerial,
                                ["PRECIO"] = matPrice,
                                ["GARANTIA"] = matWarr,
                                ["GM2"] = matGM2
                            };

                            IRfcFunction zdmCreateMat = sap.ExecuteRFC(erpMand, "ZDM_CREATE_MAT", zdmCreateMatParameters);
                            #region Procesar Salidas del FM

                            response1 = response1 + matId + ": " + zdmCreateMat.GetValue("RESULTADO").ToString() + "<br>";

                            if (zdmCreateMat.GetValue("RESULTADO").ToString() == "Material Bloqueado")
                            {
                                console.WriteLine(matId + ": " + zdmCreateMat.GetValue("RESULTADO").ToString());
                                response2 = response2 + matId + ": " + response1 + "<br>";
                                resLog = matId + ": " + response1;
                            }
                            else if (zdmCreateMat.GetValue("RESULTADO").ToString() == "Material ya existe")
                            {
                                Dictionary<string, string> zdmChangeMaterialParameters = new Dictionary<string, string>
                                {
                                    ["MATERIAL"] = matId,
                                    ["MG"] = materialGroup,
                                    ["GM1"] = matGM1,
                                    ["GM2"] = matGM2,
                                    ["DESCRIPCION"] = matDesc,
                                    ["ITEM"] = matItemCat,
                                    ["SERIALIZABLE"] = matSerial,
                                    ["PRECIO"] = matPrice
                                };
                                if (matWarr != "")
                                    zdmChangeMaterialParameters["GARANTIA"] = matWarr;

                                IRfcFunction zdmChangeMaterial = sap.ExecuteRFC(erpMand, "ZDM_CHANGE_MATERIAL", zdmChangeMaterialParameters);

                                if (zdmChangeMaterial.GetValue("RESULTADO_PRECIO").ToString() != "Se cambio el precio" && zdmChangeMaterial.GetValue("RESULTADO_PRECIO").ToString() != "")
                                    response1 = "Error al actualizar el material";
                                else
                                    response1 = "Material ha sido actualizado";

                                console.WriteLine(matId + ": " + response1);
                                response2 = response2 + matId + ": " + response1 + "<br>";
                                resLog = matId + ": " + response1;
                            }
                            else
                            {
                                Dictionary<string, string> zdmCreateExtraParameters = new Dictionary<string, string>
                                {
                                    ["MATERIAL"] = matId
                                };

                                IRfcFunction zdmCreateExtra = sap.ExecuteRFC(erpMand, "ZDM_CREATE_EXTRA", zdmCreateExtraParameters);

                                if (matWarr == "" || matWarr == "N/A" || matWarr == "Vacio")
                                {
                                    response2 = response2 + matId + ": " + zdmCreateExtra.GetValue("RESULTADO").ToString() + "<br>";
                                    resLog = matId + ": " + zdmCreateExtra.GetValue("RESULTADO").ToString();
                                    console.WriteLine(matId + ": " + zdmCreateExtra.GetValue("RESULTADO").ToString());
                                }
                                else
                                {
                                    string startDate = "", month = "";

                                    if (DateTime.Now.Month < 10)
                                        month = "0" + DateTime.Now.Month;
                                    else
                                        month = DateTime.Now.Month.ToString();

                                    startDate = "01." + month + "." + DateTime.Now.Year;

                                    Dictionary<string, string> zdmCreatePrdwtyParameters = new Dictionary<string, string>
                                    {
                                        ["MATERIAL"] = matId,
                                        ["START_DATE"] = startDate,
                                        ["END_DATE"] = "31.12.9999",
                                        ["GARANTIA"] = matWarr
                                    };

                                    IRfcFunction zdmCreatePrdwty = new SapVariants().ExecuteRFC(crmMand, "ZDM_CREATE_PRDWTY", zdmCreatePrdwtyParameters);

                                    if (zdmCreatePrdwty.GetValue("RESULTADO").ToString() == "Garantía relacionada con el material")
                                    {
                                        response2 = response2 + matId + ": " + "Material Creado con Exito" + "<br>";
                                        resLog = matId + ": " + "Material Creado con Exito";
                                        console.WriteLine(matId + ": " + "Material Creado con Exito");
                                    }
                                    else
                                    {
                                        response2 = response2 + matId + ": " + "Favor contactar a Datos Maestros: error al crear garantía. " + zdmCreatePrdwty.GetValue("RESULTADO").ToString() + " " + matWarr + "<br>";
                                        resLog = matId + ": " + "Favor contactar a Datos Maestros: error al crear garantía. " + zdmCreatePrdwty.GetValue("RESULTADO").ToString() + " " + matWarr + "<br>";
                                        console.WriteLine(matId + ": " + "Favor contactar a Datos Maestros: error al crear garantía. " + zdmCreatePrdwty.GetValue("RESULTADO").ToString() + " " + matWarr + "<br>");
                                    }
                                }
                            }

                            //log de cambios base de datos
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Material", resLog, root.Subject);
                            respFinal = respFinal + "\\n" + "Crear Material: " + resLog;

                            if (response2.Contains("Favor contactar a Datos Maestros:"))
                                validateData = false;
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            resFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + resFailure);
                            response2 = response2 + matId + ": " + ex.ToString() + "<br>";
                            resFailure = ex.ToString();
                            validateData = false;
                        }

                        #endregion
                    }
                }
                else //MASIVO
                {
                    string adjunto = root.ExcelFile;

                    #region abrir excel
                    console.WriteLine("Abriendo excel y validando");
                    DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + adjunto);
                    rows = xlWorkSheet.Rows.Count;

                    #endregion

                    if (rows > 50)
                    {
                        retMsg = "Para la creación masiva de datos, favor enviar la gestión directamente a Datos Maestros";

                        console.WriteLine("Devolviendo solicitud");
                        dm.ChangeStateDM(root.IdGestionDM, retMsg, "5"); //RECHAZADO
                        mail.SendHTMLMail(retMsg, new string[] { root.BDUserCreatedBy }, root.Subject, new string[] { "hlherrera@gbm.net" });
                        return;
                    }

                    foreach (DataRow row in xlWorkSheet.Rows)
                    {
                        material = row["Cod. Material"].ToString().Trim();
                        if (material != "")
                        {
                            matType = row["Tipo material"].ToString().Trim();
                            materialGroup = row["Grupo Articulo"].ToString().Trim();
                            itemCat = row["Grup.Tipo Posición"].ToString().Trim();
                            gm1 = row["Grupo de Material 1"].ToString().Trim();
                            desc = row["Texto comercial"].ToString().Trim();
                            serial = row["Perfil Numero Serie"].ToString().Trim();
                            price = row["Costo "].ToString().Trim();
                            warr = row["Garantia "].ToString().Trim();
                            gm2 = row["Grupo de Material 2"].ToString().Trim();

                            #region validación de datos

                            if (desc == "")
                            {
                                retMsg = "El material: " + material + " Por favor ingresar la descripción";
                                response2 = response2 + retMsg + "<br>";
                                retRequest = true;
                                continue;
                            }
                            if (materialGroup == "")
                            {
                                retMsg = "El material: " + material + " Por favor ingresar el material group";
                                response2 = response2 + retMsg + "<br>";
                                retRequest = true;
                                continue;
                            }
                            if ((itemCat.ToUpper() == "FEATURE"))
                                itemCat = "FEAT";
                            if ((itemCat != "FEAT") && (itemCat != "NORM"))
                            {
                                retMsg = "El material: " + material + " Por favor ingresar un item category group correcto";
                                response2 = response2 + retMsg + "<br>";
                                retRequest = true;
                                continue;
                            }
                            if ((warr == "N/A"))
                                warr = "";

                            if (serial == "" && warr == "" || serial == "" && warr == "N/A")
                                serial = "NO";

                            else if ((serial == "") && (warr != ""))
                                serial = "SI";
                            else if ((serial == "X") || (serial == "x"))
                                serial = "SI";

                            if (warr != "")
                            {
                                if (warr.Substring(0, 4) != "WAR-")
                                {
                                    retMsg = "El material: " + material + " La garantia no existe en SAP";
                                    response2 = response2 + retMsg + "<br>";
                                    retRequest = true;
                                    continue;
                                }

                                lenght = (warr.IndexOf(" - ") + 2);
                                if (lenght == 1)
                                { lenght = warr.Length + 2; }
                                warr = warr.Substring(0, lenght - 2);
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

                            lenght = (gm2.IndexOf("-") + 1);
                            if (lenght == 0)
                                lenght = gm2.Length + 2;

                            gm2 = gm2.Substring(0, lenght - 2);

                            matType = matType.Substring(0, 4);


                            material = material.ToUpper();

                            if (matType == "ZREP")
                                matType = "ZHRW";

                            if (material.Length > 18)
                            {
                                retMsg = "El codigo del material debe de ser menor a 18 caracteres: " + material;
                                retRequest = true;
                                response2 = response2 + retMsg + "<br>";
                                continue;
                            }

                            desc = val.RemoveSpecialChars(desc, 1);
                            desc = desc.ToUpper();

                            material = material.Replace("á", "a"); material = material.Replace("é", "e"); material = material.Replace("í", "i"); material = material.Replace("ó", "o"); material = material.Replace("ú", "u"); material = material.Replace("ñ", "n");

                            if (material.Substring(0, 3) == "800" && materialGroup.Substring(0, 2) == "40")
                            {
                                retMsg = "El material: " + material + " es un contrato por favor hacer la solicitud en el formulario de Servicios";
                                response2 = response2 + retMsg + "<br>";
                                retRequest = true;
                                continue;
                            }

                            if ((materialGroup.Substring(0, 3) == "103" || materialGroup == "201010120"))
                            {
                                warr = "";
                                gm2 = "";
                            }

                            if (serial == "SI" && warr == "N/A" && materialGroup.Substring(0, 3) != "103" && materialGroup != "201010120"
                                || serial == "SI" && warr == "" && materialGroup.Substring(0, 3) != "103" && materialGroup != "201010120")
                            {
                                retMsg = "El material: " + material + " es serializable y no tiene garantia";
                                response2 = response2 + retMsg + "<br>";
                                retRequest = true;
                                continue;
                            }

                            if (material.Substring(material.Length - 3, 3) == "_NI" && materialGroup.Substring(0, 3) == "402")
                            {
                                retMsg = "Por favor enviar esta solicitud directamente a datos maestros, ya que son materiales de PS";
                                response2 = response2 + retMsg + "<br>";
                                retRequest = true;
                                continue;
                            }

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
                                        // ejemplo 100,000.34 ---- 100000,34
                                        price = price.Replace(",", "");
                                        price = price.Replace(".", ",");
                                    }

                                }

                                if (price == "9999999" || price == "999999.99" || price == "999,999")
                                    price = "999999";

                                if (price == "0" || price == "0,00" || price == "0.00")
                                    price = "";
                            }


                            if (materialGroup == "1040109")
                                matType = "ZSFW";

                            #endregion validación de datos

                            #region SAP

                            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                            try
                            {

                                Dictionary<string, string> parameters = new Dictionary<string, string>
                                {
                                    ["TIPO_MAT"] = matType,
                                    ["MATERIAL"] = material,
                                    ["GM"] = materialGroup,
                                    ["ITEMCAT"] = itemCat,
                                    ["GM1"] = gm1,
                                    ["DESCRIPCION"] = desc,
                                    ["SERIALIZABLE"] = serial,
                                    ["PRECIO"] = price,
                                    ["GARANTIA"] = warr,
                                    ["GM2"] = gm2
                                };

                                IRfcFunction func = new SapVariants().ExecuteRFC(erpMand, "ZDM_CREATE_MAT", parameters);

                                #region Procesar Salidas del FM

                                response1 = response1 + material + ": " + func.GetValue("RESULTADO").ToString() + "<br>";

                                if (func.GetValue("RESULTADO").ToString() == "Material Bloqueado")
                                {
                                    console.WriteLine(material + ": " + func.GetValue("RESULTADO").ToString());
                                    response2 = response2 + material + ": " + response1 + "<br>";
                                    resLog = material + ": " + response1;
                                }
                                else if (func.GetValue("RESULTADO").ToString() == "Material ya existe")
                                {
                                    Dictionary<string, string> parameters2 = new Dictionary<string, string>
                                    {
                                        ["MATERIAL"] = material,
                                        ["MG"] = materialGroup,
                                        ["GM1"] = gm1,
                                        ["GM2"] = gm2,
                                        ["DESCRIPCION"] = desc,
                                        ["ITEM"] = itemCat,
                                        ["SERIALIZABLE"] = serial,
                                        ["PRECIO"] = price
                                    };
                                    if (warr != "")
                                        parameters2["GARANTIA"] = warr;

                                    IRfcFunction func_change = sap.ExecuteRFC(erpMand, "ZDM_CHANGE_MATERIAL", parameters2);

                                    if (func_change.GetValue("RESULTADO_PRECIO").ToString() != "Se cambio el precio" && func_change.GetValue("RESULTADO_PRECIO").ToString() != "")
                                        response1 = "Error al actualizar el material";
                                    else
                                        response1 = "Material ha sido actualizado";

                                    console.WriteLine(material + ": " + response1);
                                    response2 = response2 + material + ": " + response1 + "<br>";
                                    resLog = material + ": " + response1;
                                }
                                else
                                {
                                    Dictionary<string, string> parameters3 = new Dictionary<string, string>
                                    {
                                        ["MATERIAL"] = material
                                    };

                                    IRfcFunction func2 = sap.ExecuteRFC(erpMand, "ZDM_CREATE_EXTRA", parameters3);

                                    if (warr == "" || warr == "N/A" || warr == "Vacio")
                                    {
                                        response2 = response2 + material + ": " + func2.GetValue("RESULTADO").ToString() + "<br>";
                                        resLog = material + ": " + func2.GetValue("RESULTADO").ToString();
                                        console.WriteLine(material + ": " + func2.GetValue("RESULTADO").ToString());
                                    }
                                    else
                                    {
                                        string start_date = "", mes = "";

                                        if (DateTime.Now.Month < 10)
                                        { mes = "0" + DateTime.Now.Month; }
                                        else
                                        { mes = DateTime.Now.Month.ToString(); }

                                        start_date = "01." + mes + "." + DateTime.Now.Year;

                                        Dictionary<string, string> parameters4 = new Dictionary<string, string>
                                        {
                                            ["MATERIAL"] = material,
                                            ["START_DATE"] = start_date,
                                            ["END_DATE"] = "31.12.9999",
                                            ["GARANTIA"] = warr
                                        };

                                        IRfcFunction func_garantia = sap.ExecuteRFC(crmMand, "ZDM_CREATE_PRDWTY", parameters4);

                                        if (func_garantia.GetValue("RESULTADO").ToString() == "Garantía relacionada con el material")
                                        {
                                            response2 = response2 + material + ": " + "Material Creado con Éxito" + "<br>";
                                            resLog = material + ": " + "Material Creado con Éxito";
                                            console.WriteLine(material + ": " + "Material Creado con Éxito");
                                        }
                                        else
                                        {
                                            response2 = response2 + material + ": " + "Favor contactar a Datos Maestros: error al crear garantía <br>" + func_garantia.GetValue("RESULTADO").ToString();
                                            resLog = material + ": " + "Favor contactar a Datos Maestros: error al crear garantía <br>" + func_garantia.GetValue("RESULTADO").ToString();
                                            console.WriteLine(material + ": " + "Favor contactar a Datos Maestros: error al crear garantía. <br>" + func_garantia.GetValue("RESULTADO").ToString());
                                        }
                                    }

                                }

                                //log de cambios base de datos
                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Material", resLog, root.Subject);
                                respFinal = respFinal + "\\n" + "Crear Material: " + resLog;


                                if (response2.Contains("Favor contactar a Internal Customer Services:"))
                                    validateData = false;
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                resFailure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                                console.WriteLine("Finishing process " + resFailure);
                                response2 = response2 + material + ": " + ex.ToString() + "<br>";
                                resFailure = ex.ToString();
                                validateData = false;
                            }

                            #endregion

                        } //IF si el material esta en blanco

                    } //for para cada fila del excel
                    if (rows == 1)
                    {
                        retMsg = "El material: " + material + " Por favor ingresar la descripción";
                        response2 = response2 + retMsg + "<br>";
                        retRequest = true;
                    }
                }

                if (validateData == false)
                {
                    console.WriteLine("enviando error de solicitud");
                    dm.ChangeStateDM(root.IdGestionDM, response2 + "<br>" + resFailure, "4"); //ERROR

                    //enviar email de repuesta de error a ICS
                    mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + "<br>" + response2 + "<br>" + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, new string[] { "hlherrera@gbm.net" });
                }
                else if (retRequest == true)
                {
                    console.WriteLine("Devolviendo solicitud");
                    dm.ChangeStateDM(root.IdGestionDM, response2, "5"); //RECHAZADO
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Materiales:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + response2);
                }
                else
                {
                    //el material se creo con éxito
                    //if (baw == "S")
                    //{
                    //    string bawSelenium = selMat.ExecuteBaw(idBaw);
                    //    if (bawSelenium != "true")
                    //        mail.SendHTMLMail("Hubo un problema al ejecutar el BAW en BPM, por favor ingrese para ejecutar el id: " + idBaw + "<br><br>" + bawSelenium, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, new string[] { "jearaya@gbm.net" });
                    //}

                    console.WriteLine("Finalizando solicitud");
                    //finalizar solicitud
                    dm.ChangeStateDM(root.IdGestionDM, response2, "3"); //FINALIZADO
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Materiales:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + response2);
                }

                root.requestDetails = respFinal;

            }
            catch (Exception ex) //catch del Proccess
            {
                dm.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //ERROR
                mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, new string[] { "hlherrera@gbm.net" });
            }
        }
    }
}


