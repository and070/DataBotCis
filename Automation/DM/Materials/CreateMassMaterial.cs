using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.DM.Materials
{
    /// <summary>
    /// Clase DM Automation encargada de la creación masiva de materiales por e-mail en datos maestros.
    /// </summary>
    class CreateMassMaterial
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        readonly string[] cc = { "hlherrera@gbm.net" };
        const string erpMand = "ERP";
        const string crmMand = "CRM";

        public void Main()
        {
            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Solicitudes Materiales Masivo", "Procesados", "Procesados Materiales Masivo"))
            {
                console.WriteLine("Procesando...");
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessMassMaterials(excelDt);
                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void ProcessMassMaterials(DataTable excelDt)
        {
            bool returnRequest = false;
            bool validateData = true;
            
            int lenght;

            string resFailure = "";
            string respFinal = "";
            string returnMsg;
            string resLog;

            excelDt.Columns.Add("Respuesta");

            try
            {
                string res1 = "";
                string res2 = "";
                string validate;

                string attachFile = root.FilesDownloadPath + "\\" + "Resultado.xlsx";
                int rows = excelDt.Rows.Count;

                #region validación de cantidad de filas

                if (rows > 500)
                {
                    returnMsg = "Para la creación mayor a 500 registros, favor enviar la gestión directamente a Internal Customer Services";
                    console.WriteLine(returnMsg);

                    console.WriteLine("Devolviendo solicitud");
                    mail.SendHTMLMail(returnMsg, new string[] { root.BDUserCreatedBy }, root.Subject, cc);
                    returnRequest = true;
                    return;
                }

                #endregion

                #region Validar si el archivo es correcto
                validate = excelDt.Columns[10].ColumnName.Trim();

                if (validate.Substring(0, 1) != "x")
                {
                    console.WriteLine("Devolviendo Solicitud");
                    returnMsg = "Favor usar la plantilla oficial de Internal Customer Services";
                    console.WriteLine(returnMsg);

                    mail.SendHTMLMail(returnMsg, new string[] { root.BDUserCreatedBy }, root.Subject, cc);
                    returnRequest = true;
                    return;
                }

                #endregion

                foreach (DataRow row in excelDt.Rows)
                {
                    string material = row[1].ToString().Trim();
                    if (material == "")
                    {
                        //si la ultima celda registrada en rows esta vacia significa que esta tomando solamente el formato
                        material = row[1].ToString().Trim();
                        if (material == "")
                        {
                            break;
                        }
                        continue;
                    }
                    else
                    {
                        string materialType = row[0].ToString().Trim();
                        string materialGroup = row[2].ToString().Trim();
                        string itemCat = row[3].ToString().Trim();
                        string gm1 = row[4].ToString().Trim();
                        string description = row[5].ToString().Trim();
                        string serial = row[6].ToString().Trim();
                        string price = row[7].ToString().Trim();
                        string warr = row[8].ToString().Trim();
                        string gm2 = row[9].ToString().Trim();

                        #region Validación de datos

                        if (description == "")
                        {
                            returnMsg = "El material: " + material + " Por favor ingresar la descripción";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
                            continue;
                        }

                        if (materialGroup == "")
                        {
                            returnMsg = "El material: " + material + " Por favor ingresar el material group";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
                            continue;
                        }

                        if (itemCat.ToUpper() == "FEATURE")
                            itemCat = "FEAT";

                        if ((itemCat != "FEAT") && (itemCat != "NORM"))
                        {
                            returnMsg = "El material: " + material + " Por favor ingresar un item category group correcto";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
                            continue;
                        }

                        if ((warr == "N/A") || warr == "N/A - No aplica")
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
                                returnMsg = "El material: " + material + " La garantía no existe en SAP";
                                res2 = res2 + returnMsg + "<br>";
                                returnRequest = true;
                                continue;
                            }
                        }

                        lenght = materialGroup.IndexOf("-") + 1;
                        if (lenght == 0)
                            lenght = materialGroup.Length + 2;

                        materialGroup = materialGroup.Substring(0, lenght - 2);
                        materialGroup = materialGroup.Replace("#", "");

                        lenght = gm1.IndexOf("-") + 1;
                        if (lenght == 0)
                            lenght = gm1.Length + 2;

                        gm1 = gm1.Substring(0, lenght - 2);

                        lenght = gm2.IndexOf("-") + 1;
                        if (lenght == 0)
                            lenght = gm2.Length + 2;

                        gm2 = gm2.Substring(0, lenght - 2);

                        materialType = materialType.Substring(0, 4);

                        material = material.ToUpper();

                        if (materialType == "ZREP")
                            materialType = "ZHRW";

                        if (material.Length > 18)
                        {
                            returnMsg = "El código del material debe de ser menor a 18 caracteres: " + material;
                            returnRequest = true;
                            res2 = res2 + returnMsg + "<br>";
                            continue;
                        }

                        description = val.RemoveSpecialChars(description, 1);
                        description = description.ToUpper();

                        material = material.Replace("á", "a"); material = material.Replace("é", "e"); material = material.Replace("í", "i"); material = material.Replace("ó", "o"); material = material.Replace("ú", "u"); material = material.Replace("ñ", "n");

                        if (material.Substring(0, 3) == "800" && materialGroup.Substring(0, 2) == "40")
                        {
                            returnMsg = "El material: " + material + " es un contrato por favor hacer la solicitud en el formulario de Servicios";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
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
                            returnMsg = "El material: " + material + " es serializable y no tiene garantía";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
                            continue;
                        }

                        if (material.Substring(material.Length - 3, 3) == "_NI" && materialGroup.Substring(0, 3) == "402")
                        {
                            returnMsg = "Por favor enviar esta solicitud directamente a datos maestros, ya que son materiales de PS";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
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
                            materialType = "ZSFW";

                        #endregion Validación de datos

                        #region SAP

                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                        try
                        {
                            Dictionary<string, string> parametros = new Dictionary<string, string>
                            {
                                ["TIPO_MAT"] = materialType,
                                ["MATERIAL"] = material,
                                ["GM"] = materialGroup,
                                ["ITEMCAT"] = itemCat,
                                ["GM1"] = gm1,
                                ["DESCRIPCION"] = description,
                                ["SERIALIZABLE"] = serial,
                                ["PRECIO"] = price,
                                ["GARANTIA"] = warr,
                                ["GM2"] = gm2
                            };

                            IRfcFunction zdmCreateMat = sap.ExecuteRFC(erpMand, "ZDM_CREATE_MAT", parametros);

                            #region Procesar Salidas del FM

                            res1 = res1 + material + ": " + zdmCreateMat.GetValue("RESULTADO").ToString() + "<br>";

                            if (zdmCreateMat.GetValue("RESULTADO").ToString() == "Material Bloqueado")
                            {
                                console.WriteLine(material + ": " + zdmCreateMat.GetValue("RESULTADO").ToString());
                                res2 = res2 + material + ": " + res1 + "<br>";
                                resLog = material + ": " + res1;
                            }
                            else if (zdmCreateMat.GetValue("RESULTADO").ToString() == "Material ya existe")
                            {
                                price = price.Replace(",", ".");

                                Dictionary<string, string> parametros2 = new Dictionary<string, string>
                                {
                                    ["MATERIAL"] = material,
                                    ["MG"] = materialGroup,
                                    ["GM1"] = gm1,
                                    ["GM2"] = gm2,
                                    ["DESCRIPCION"] = description,
                                    ["ITEM"] = itemCat,
                                    ["SERIALIZABLE"] = serial,
                                    ["PRECIO"] = price
                                };

                                if (warr != "")
                                    parametros["GARANTIA"] = warr;

                                IRfcFunction zdmChangeMaterial = sap.ExecuteRFC(erpMand, "ZDM_CHANGE_MATERIAL", parametros2);

                                if (zdmChangeMaterial.GetValue("RESULTADO_PRECIO").ToString() != "Se cambio el precio" && zdmChangeMaterial.GetValue("RESULTADO_PRECIO").ToString() != "")
                                    res1 = "Error al actualizar el material";
                                else
                                    res1 = "Material ha sido actualizado";

                                console.WriteLine(material + ": " + res1);
                                res2 = res2 + material + ": " + res1 + "<br>";
                                resLog = material + ": " + res1;
                            }
                            else
                            {
                                Dictionary<string, string> parametros3 = new Dictionary<string, string>
                                {
                                    ["MATERIAL"] = material
                                };

                                IRfcFunction zdmCreateExtra = sap.ExecuteRFC(erpMand, "ZDM_CREATE_EXTRA", parametros3);

                                if (warr == "" || warr == "N/A" || warr == "Vacio" || warr == "N/A - No aplica")
                                {
                                    res2 = res2 + material + ": " + zdmCreateExtra.GetValue("RESULTADO").ToString() + "<br>";
                                    resLog = material + ": " + zdmCreateExtra.GetValue("RESULTADO").ToString();
                                    console.WriteLine(material + ": " + zdmCreateExtra.GetValue("RESULTADO").ToString());
                                }
                                else
                                {
                                    string start_date = "", mes = "";

                                    if (DateTime.Now.Month < 10)
                                        mes = "0" + DateTime.Now.Month;
                                    else
                                        mes = DateTime.Now.Month.ToString();

                                    start_date = "01." + mes + "." + DateTime.Now.Year;

                                    Dictionary<string, string> parametros4 = new Dictionary<string, string>
                                    {
                                        ["MATERIAL"] = material,
                                        ["START_DATE"] = start_date,
                                        ["END_DATE"] = "31.12.9999",
                                        ["GARANTIA"] = warr
                                    };

                                    IRfcFunction zdmCreatePrdwty = sap.ExecuteRFC(crmMand, "ZDM_CREATE_PRDWTY", parametros4);

                                    if (zdmCreatePrdwty.GetValue("RESULTADO").ToString() == "Garantía relacionada con el material")
                                    {
                                        res2 = res2 + material + ": " + "Material Creado con Exito" + "<br>";
                                        resLog = material + ": " + "Material Creado con Exito";
                                        console.WriteLine(material + ": " + "Material Creado con Exito");
                                    }
                                    else
                                    {
                                        res2 = res2 + material + ": " + "Favor contactar a Datos Maestros: error al crear garantía" + "<br>";
                                        resLog = material + ": " + "Favor contactar a Datos Maestros: error al crear garantía";
                                        console.WriteLine(material + ": " + "Favor contactar a Datos Maestros: error al crear garantía");
                                    }
                                }

                            }
                            row["Respuesta"] = resLog;

                            //log de cambios base de datos
                            log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Material", resLog, root.Subject);
                            respFinal = respFinal + "\\n" + "Crear Material " + resLog;

                            if (res2.Contains("Favor contactar a Internal Customer Services:"))
                                validateData = false;

                            #endregion
                        }
                        catch (Exception ex)
                        {
                            res2 = res2 + material + ": " + ex.ToString() + "<br>";
                            resFailure = ex.ToString();
                            validateData = false;
                            console.WriteLine("Finishing process: " + resFailure);
                        }

                        #endregion

                    }
                }

                excel.CreateExcel(excelDt, "Sheet1", attachFile);

                if (validateData == false)
                {
                    //enviar email de repuesta de error
                    mail.SendHTMLMail(res2 + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, cc, new string[] { attachFile });
                }
                else if (returnRequest == true)
                {
                    console.WriteLine("Devolviendo solicitud");
                    mail.SendHTMLMail(res2 + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, new string[] { root.BDUserCreatedBy });
                }
                else
                {
                    //enviar email de repuesta de éxito
                    mail.SendHTMLMail("Los resultados están en el excel", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC, new string[] { attachFile });
                }
            }
            catch (Exception ex)
            {
                mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }

            root.requestDetails = respFinal;
        }
    }
}