using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System.Linq;
using System;

namespace DataBotV5.Automation.ICS.BusinessPartners

{
    /// <summary>
    /// Clase ICS Automation encargada de actualizar las retenciones de Business Partner.
    /// </summary>
    class BPRetentions
    {
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();

        string mand = "ERP";
        string respFinal = "";


        public void Main()
        {
            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Solicitudes Retenciones", "Procesados", "Procesados Retenciones"))
            {
                console.WriteLine("Procesando...");
                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessRet(excelDt);
                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessRet(DataTable excelDt)
        {
            bool validateLines = true;
            string responseFailure = "";

            string validation = excelDt.Columns[4].ColumnName;

            if (validation.Substring(0, 1) != "x")
            {
                console.WriteLine("Devolviendo Solicitud");
                mail.SendHTMLMail("Utilizar la plantilla oficial de datos maestros", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }
            else
            {
                string fmrep, whType, whTaxCod, action;
                string bpType = excelDt.Columns[0].ColumnName;

                excelDt.Columns.Add("Resultado");

                if (bpType == "Proveedor")
                {
                    #region Validaciones
                    foreach (DataRow item in excelDt.Rows)
                    {
                        string vendor = item["Proveedor"].ToString().Trim();
                        if (vendor != "")
                        {
                            bool isNumeric = int.TryParse(vendor, out int n);
                            if (!isNumeric)
                            {
                                item["Resultado"] = vendor + ": " + "el Proveedor no es un ID valido" + "<br>";
                                validateLines = false;
                                continue;
                            }

                            string coCode = item["Company_Code"].ToString().Trim();

                            whType = item["Withholding tax type"].ToString().Trim();
                            whType = string.Concat(whType.Take(2));

                            whTaxCod = item["Withholding tax Code"].ToString().Trim();
                            whTaxCod = string.Concat(whTaxCod.Take(2));

                            action = item["xAcción"].ToString().Trim().ToUpper();

                            if (string.Concat(vendor.Take(2)) != "00")
                                vendor = "00" + vendor;

                            #region SAP
                            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                            try
                            {
                                Dictionary<string, string> parameters = new Dictionary<string, string>
                                {
                                    ["VENDOR"] = vendor,
                                    ["COMPANY"] = coCode,
                                    ["ACCION"] = action,
                                    ["TIPO_RET"] = whType,
                                    ["IND_RET"] = whTaxCod
                                };

                                IRfcFunction func = sap.ExecuteRFC(mand, "ZDM_RET_PROVD", parameters);

                                #region Procesar Salidas del FM
                                if (func.GetValue("RESPUESTA").ToString() == "ACCION INVALIDA")
                                    fmrep = "Error: No se pudo actualizar el Proveedor, acción seleccionada invalida";
                                else
                                    fmrep = func.GetValue("RESPUESTA").ToString();

                                console.WriteLine(vendor + ": " + fmrep);
                                item["Resultado"] = fmrep;

                                //log de cambios base de datos
                                log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Crear Retencion", vendor + ": " + fmrep, root.Subject);
                                respFinal = respFinal + "\\n" + vendor + ": " + fmrep;

                                if (fmrep.Contains("Error"))
                                    validateLines = false;
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                item["Resultado"] = vendor + ": " + ex.ToString() + "<br>";

                                console.WriteLine(" Finishing process " + val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0));
                                responseFailure = ex.ToString();
                                validateLines = false;
                            }
                            #endregion
                        }
                    }
                    #endregion
                }
                else
                {
                    //realiza el proceso de clientes
                    //Plantilla correcta, continúe las validaciones

                    foreach (DataRow item in excelDt.Rows)
                    {
                        string customer = item[1].ToString().Trim();
                        if (customer != "")
                        {
                            bool isNumeric = int.TryParse(customer, out int n);
                            if (!isNumeric)
                            {
                                item["Resultado"] = customer + ": " + "el Cliente no es el ID" + "<br>";
                                continue;
                            }
                            string country = item[0].ToString().Trim();

                            whType = item[2].ToString().Trim();
                            whType = string.Concat(whType.Take(2));

                            whTaxCod = item[3].ToString().Trim();
                            whTaxCod = string.Concat(whTaxCod.Take(2));

                            action = item[4].ToString().Trim().ToUpper();

                            if (string.Concat(customer.Take(2)) != "00")
                                customer = "00" + customer;

                            #region SAP
                            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                            try
                            {
                                Dictionary<string, string> parameters = new Dictionary<string, string>
                                {
                                    ["BP"] = customer,
                                    ["COMPANY"] = country,
                                    ["ACCION"] = action,
                                    ["TIPORET"] = whType,
                                    ["INDRET"] = whTaxCod
                                };

                                IRfcFunction func = sap.ExecuteRFC(mand, "ZRPA_BP_RET", parameters);

                                #region Procesar Salidas del FM

                                if (func.GetValue("RESPUESTA").ToString() == "ACCION INVALIDA")
                                    fmrep = "Error: No se pudo actualizar el BP, accion seleccionada invalida";
                                else
                                    fmrep = func.GetValue("RESPUESTA").ToString();

                                console.WriteLine(customer + ": " + fmrep);
                                item["Resultado"] = fmrep;

                                //log de cambios base de datos
                                log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Crear Retencion", customer + ": " + fmrep, root.Subject);
                                respFinal = respFinal + "\\n" + customer + ": " + fmrep;

                                if (fmrep.Contains("Error"))
                                    validateLines = false;
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                responseFailure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                                console.WriteLine(" Finishing process " + responseFailure);
                                item["Resultado"] = customer + ": " + ex.ToString() + "<br>";
                                responseFailure = ex.ToString();
                                validateLines = false;
                            }

                            #endregion
                        }
                    }
                }

                string htmlTable = val.ConvertDataTableToHTML(excelDt);

                console.WriteLine("Respondiendo solicitud");
                if (validateLines == false)
                    //enviar email de repuesta de error
                    mail.SendHTMLMail(htmlTable + "<br>" + responseFailure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
                else
                    //enviar email de repuesta de éxito
                    mail.SendHTMLMail(htmlTable, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

                root.requestDetails = respFinal;

            }
        }
    }
}
