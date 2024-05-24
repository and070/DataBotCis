using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.App.Global.Interfaces;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Data.Process;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.MASS.Materials
{
    /// <summary>
    /// Clase MASS Automation encargada de la modificación masiva de materiales por e-mail.
    /// </summary>
    class ModMaterial : IProceso
    {
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly ValidateData val = new ValidateData();
        readonly MsExcel excel = new MsExcel();
        readonly Rooting root = new Rooting();
        readonly Log log = new Log();

        string respFinal = "";

        public void Main()
        {
            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Solicitudes Mod Mat", "Procesados", "Procesados Mod Mat"))
            {
                string filePath = root.FilesDownloadPath + "\\" + root.ExcelFile;
                DataTable excelDt = excel.GetExcel(filePath);
                //bool correctVersion = DM.CheckVersion("MODIFICACION_MATERIALES", filePath);

                ProcessingToModifyMaterials(excelDt, true);

                using (Stats stats = new Stats()) { stats.CreateStat(); }
            }
        }
        public void ProcessingToModifyMaterials(DataTable excelDt, bool correctFileVersion)
        {
            bool validateLines = true;

            int rowCount = excelDt.Rows.Count;

            bool validateFile = ValidateFile(excelDt);

            if (validateFile && correctFileVersion)
            {
                int cont = 0;
                foreach (DataRow item in excelDt.Rows)
                {
                    cont++;
                    try { Console.WriteLine("Progreso " + (cont * 100) / rowCount + "%"); } catch (Exception) { }

                    string itemRes = "";
                    string material = item[1].ToString().Trim();

                    if (material != "")
                    {
                        string matType = item[0].ToString().Trim();
                        string matGroup = item[2].ToString().Trim();
                        string itemCat = item[3].ToString().Trim();
                        string gm1 = item[4].ToString().Trim();
                        string matDesc = item[5].ToString().Trim();
                        string matSerial = item[6].ToString().Trim();
                        string price = item[7].ToString().Trim();
                        string matWarr = item[8].ToString().Trim();
                        string gm2 = item[9].ToString().Trim();
                        string flag = item[10].ToString().Trim();

                        #region Validación de datos

                        if (itemCat.ToUpper() == "FEATURE")
                            itemCat = "FEAT";

                        if ((itemCat != "FEAT") && (itemCat != "NORM") && (itemCat != ""))
                        {
                            itemRes = "El material: " + material + " Por favor ingresar un item category group correcto";
                            item["xRespuesta del databot"] = itemRes;
                            continue;
                        }

                        if ((matSerial == "X") || (matSerial == "x"))
                            matSerial = "SI";

                        if (matSerial == "S")
                            matSerial = "SI";

                        if (matSerial == "N")
                            matSerial = "NO";

                        int lenght;
                        if (matGroup != "")
                        {
                            lenght = (matGroup.IndexOf("-") + 1);
                            if (lenght == 0)
                                lenght = matGroup.Length + 2;
                            matGroup = matGroup.Substring(0, lenght - 2);
                            matGroup = matGroup.Replace("#", "");
                        }

                        if (gm1 != "")
                        {
                            lenght = (gm1.IndexOf("-") + 1);
                            if (lenght == 0)
                                lenght = gm1.Length + 2;
                            gm1 = gm1.Substring(0, lenght - 2);
                        }

                        if (gm2 != "")
                        {
                            lenght = (gm2.IndexOf("-") + 1);
                            if (lenght == 0)
                                lenght = gm2.Length + 2;
                            gm2 = gm2.Substring(0, lenght - 2);
                        }

                        if (matType.Length > 4)
                            matType = matType.Substring(0, 4);

                        if (material.Length > 18)
                        {
                            itemRes = "El material: " + material + " La longitud del material supera los 18 caracteres";
                            item["xRespuesta del databot"] = itemRes;
                            continue;
                        }

                        if (price == "Vacio")
                            price = "";

                        if (gm2 == "Vacio")
                            gm2 = "";

                        if ((matWarr.ToLower() == "n/a - no aplica") || (matWarr == "N/A"))
                            matWarr = "";

                        if (matWarr == "Vacio")
                            matWarr = "";

                        if (matWarr != "")
                        {
                            if (matWarr.Length > 4)
                            {
                                if (matWarr.Substring(0, 4) != "WAR-")
                                {
                                    itemRes = "El material: " + material + " La garantía no existe en SAP";
                                    item["xRespuesta del databot"] = itemRes;
                                    continue;
                                }
                            }

                        }

                        if (gm1.Length < 2 && gm1 != "")
                            gm1 = "0" + gm1;

                        material = material.ToUpper();

                        if (matDesc != "")
                        {
                            matDesc = val.RemoveSpecialChars(matDesc, 1);
                            matDesc = matDesc.ToUpper();
                        }

                        material = material.Replace("á", "a");
                        material = material.Replace("é", "e");
                        material = material.Replace("í", "i");
                        material = material.Replace("ó", "o");
                        material = material.Replace("ú", "u");
                        material = material.Replace("ñ", "n");

                        if (material.Length >= 3)
                        {
                            if (matGroup.Length >= 2)
                            {
                                if (material.Substring(0, 3) == "800" && matGroup.Substring(0, 2) == "40")
                                {
                                    itemRes = "El material: " + material + " es un contrato por favor hacer la solicitud en el formulario de Servicios";
                                    item["xRespuesta del databot"] = itemRes;
                                    continue;
                                }
                            }

                        }

                        if (matGroup != "")
                        {
                            if (matGroup.Length > 3)
                            {
                                if (matGroup.Substring(0, 3) == "103" || matGroup == "201010120")
                                    matWarr = "";
                            }

                        }

                        if (price != "")
                        {
                            price = price.Replace("$", "");
                            if (price.Length > 3)
                            {
                                if (price.Substring(0, 3).Substring(price.Substring(0, 3).Length - 1, 1) == "," && price.Substring((price.Length - 3)).Substring(0, 1) != "," || price.Substring(0, 2).Substring(price.Substring(0, 2).Length - 1, 1) == "," && price.Substring(price.Length - 3, 1).Substring(0, 1) != ",")
                                    price = price.Replace(",", "");

                                if ((price.Substring(price.Length - 3).Substring(0, 1) == ".") || (price.Substring(price.Length - 2).Substring(0, 1) == "."))
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

                        #endregion Validación de datos

                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                        try
                        {

                            Dictionary<string, string> parameters = new Dictionary<string, string>
                            {
                                ["MATERIAL"] = material,
                                ["MG"] = matGroup,
                                ["GM1"] = gm1,
                                ["GM2"] = gm2,
                                ["DESCRIPCION"] = matDesc,
                                ["ITEM"] = itemCat,
                                ["FLAG"] = flag,
                                ["SERIALIZABLE"] = matSerial,
                                ["PRECIO"] = price
                            };

                            if (matWarr != "")
                                parameters["GARANTIA"] = matWarr;

                            IRfcFunction zdmChangeMaterial = new SapVariants().ExecuteRFC("ERP", "ZDM_CHANGE_MATERIAL", parameters);

                            if (zdmChangeMaterial.GetValue("RESULTADO").ToString() == "Material Bloqueado")
                                itemRes = "Material Bloqueado";
                            else if (zdmChangeMaterial.GetValue("RESULTADO").ToString() == "Material No Existe")
                                itemRes = "Material No Existe";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_PRECIO").ToString() != "Se cambio el precio" && zdmChangeMaterial.GetValue("RESULTADO_PRECIO").ToString() != "")
                                itemRes = "Error al actualizar el precio del material";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_CAT").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error:" + zdmChangeMaterial.GetValue("RESULTADO_CAT").ToString();
                            else if (zdmChangeMaterial.GetValue("RESULTADO_MAT_TYPE").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error al actualizar el tipo de material";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_TEXT").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error al actualizar la descripción del material";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_WAR").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error al actualizar la garantía del material";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_ITEM").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error al actualizar el item category group del material";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_SERIAL").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error al actualizar la serialización del material";
                            else if (zdmChangeMaterial.GetValue("RESULTADO_FLAG").ToString().Contains("Favor contactar a datos maestros"))
                                itemRes = "Error al actualizar flag for deletion del material";
                            else
                                itemRes = "Material ha sido actualizado";

                            console.WriteLine(material + ": " + itemRes);

                            log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, material + ": " + itemRes, "", "");
                            respFinal = respFinal + "\\n" + material + ": " + itemRes;

                        }
                        catch (Exception ex)
                        {
                            using (RPAScheduler orquestador = new RPAScheduler())
                            {
                                new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                            }

                            console.WriteLine("Finishing process " + ex.Message);
                            itemRes = ex.Message;
                            validateLines = false;
                        }

                        #endregion
                    }

                    item["xRespuesta del databot"] = itemRes;
                }
            }
            else
            {
                console.WriteLine("Devolviendo Solicitud");
                mail.SendHTMLMail("Utilizar la plantilla oficial para modificar materiales.<br>", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }

            console.WriteLine("Respondiendo solicitud");

            string emailMsg = "Se muestra el resultado a continuación<br><br>" + val.ConvertDataTableToHTML(excelDt);

            if (validateLines == false)//enviar email de repuesta de error
                mail.SendHTMLMail(emailMsg, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, new string[] { root.BDUserCreatedBy });
            else// enviar email de repuesta de éxito
                mail.SendHTMLMail(emailMsg, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

            root.requestDetails = respFinal;

        }

        private bool ValidateFile(DataTable excelDt)
        {
            try
            {
                string validateFile = excelDt.Columns[11].ColumnName;
                if (validateFile.Substring(0, 1).ToLower() == "x")
                    return true;
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
