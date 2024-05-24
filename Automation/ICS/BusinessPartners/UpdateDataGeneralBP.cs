using DataBotV5.Logical.MicrosoftTools;
using System.Text.RegularExpressions;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Stats;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;

namespace DataBotV5.Automation.ICS.BusinessPartners
{
    /// <summary>
    /// Clase ICS Automation encargada de modificar la data general de Business Partner.
    /// </summary>
    class UpdateDataGeneralBP
    {
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();

        string mandErp = "ERP";
        string respFinal = "";


        public void Main()
        {
            console.WriteLine("Descargando archivo");
            //leer correo y descargar archivo
            if (mail.GetAttachmentEmail("Modificaciones BP", "Procesados", "Procesados Mod BP"))
            {
                console.WriteLine("Procesando...");

                DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                ProcessBpGen(excelDt);

                console.WriteLine("Creando Estadísticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        public void ProcessBpGen(DataTable excelDt)
        {
            bool returnReq = false, validateLines = true;
            try { excelDt.Columns.Add("Resultado"); } catch (DuplicateNameException) { }

            string validate = excelDt.Columns[6].ColumnName;
            int rows = excelDt.Rows.Count;

            try { validate = validate.Substring(0, 1); } catch (Exception) { }

            if (validate == "x")
            {
                if (rows > 50)
                    mail.SendHTMLMail("Para la creación masiva de datos, favor enviar la gestión directamente a Internal Customer Services", new string[] { root.BDUserCreatedBy }, root.Subject);
                else
                {
                    //Plantilla correcta, continúe las validaciones
                    foreach (DataRow item in excelDt.Rows)
                    {
                        string fmRes = "";
                        string itemRes = "";
                        string customerID = item[0].ToString().Trim();

                        if (customerID != "")
                        {
                            string customerName = item[1].ToString().Trim().ToUpper();
                            string nif = item[2].ToString().Trim();
                            string address = item[3].ToString().Trim().ToUpper();
                            string phone = item[4].ToString().Trim().ToUpper();
                            string email = item[5].ToString().Trim();
                            string nif2 = item[6].ToString().Trim();

                            string firstName = "";
                            string lastName = "";
                            try { firstName = item[8].ToString().Trim().ToUpper(); } catch (Exception) { }
                            try { lastName = item[9].ToString().Trim().ToUpper(); } catch (Exception) { }

                            #region Validación de datos

                            #region Cliente

                            customerID = val.Clean(customerID);

                            bool Numeric = int.TryParse(customerID, out int num);
                            if (Numeric == false)
                            {
                                itemRes = "ID del cliente inválido";
                                continue;
                            }
                            if (customerID.Substring(0, 2) != "00")
                                customerID = "00" + customerID;

                            #endregion
                            #region Razón social

                            if (customerName != "")
                            {
                                customerName = val.Clean(customerName);
                                customerName = val.RemoveChars(customerName);
                                if (customerName.Length > 80)
                                    customerName = customerName.Substring(0, 80);
                            }

                            #endregion
                            #region Dirección

                            if (address != "")
                            {
                                address = val.Clean(address);
                                address = val.RemoveSpecialChars(address, 1);
                                if (address.Length > 140)
                                    address = address.Substring(0, 140);
                                address = val.RemoveChars(address);
                            }
                            #endregion
                            #region Nif

                            if (nif != "" || nif2 != "")
                            {
                                if (nif2 == "N/A" || nif2 == "NA")
                                    nif2 = "";
                                if (nif == nif2)
                                    nif2 = "";
                                if (nif2 == "00" || nif2 == "000")
                                    nif2 = "";
                                if (nif.Contains("EIN Number"))
                                    nif = nif.Replace("EIN Number", "");

                                string swap;
                                if (nif.Length < nif2.Length)
                                {
                                    swap = nif; nif = nif2; nif2 = swap;
                                }

                                if (nif == "N/A" || nif == "NA")
                                {
                                    returnReq = true;
                                    itemRes = "nif del cliente (" + customerID + " - " + nif + ") inválido";
                                }

                                #region Extraer país de cliente
                                string customerCountry = "";
                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>
                                    {
                                        ["BP"] = customerID
                                    };

                                    IRfcFunction zdmReadBp = new SapVariants().ExecuteRFC(mandErp, "ZDM_READ_BP", parameters);
                                    customerCountry = zdmReadBp.GetValue("PAIS").ToString();
                                }
                                catch (Exception) { }

                                #endregion

                                if (customerCountry != "")
                                {
                                    switch (customerCountry)
                                    {
                                        case "SV":
                                            if (nif.Length >= 17)
                                                nif = nif.Substring(0, nif.Length - 2) + nif.Substring(nif.Length - 1, 1);
                                            nif2 = Regex.Replace(nif2, "[^0-9-]", "");
                                            break;
                                        case "PA":
                                            int dv = nif.IndexOf("DV");
                                            if (dv > 0)
                                            {
                                                nif2 = nif.Substring(dv, nif.Length - dv);
                                                nif = nif.Substring(0, dv - 1);
                                            }
                                            nif2 = Regex.Replace(nif2, @"[^\d]", "");
                                            if (nif2.Length > 3)
                                            {
                                                returnReq = true;
                                                itemRes = "nif2 (" + customerID + " - " + nif2 + ") inválido";
                                            }
                                            break;
                                        case "CR":
                                            break;
                                        case "DO":
                                            break;
                                        case "HN":
                                            break;
                                        case "NI":
                                            break;
                                        case "CO":
                                            nif = val.RemoveSpecialChars(nif, 2);
                                            nif2 = "";
                                            break;
                                        case "VE":
                                            if (nif != "")
                                            {
                                                if (!(nif.Contains("-")))
                                                {
                                                    string mid;
                                                    string ini = nif.Substring(0, 1).ToString();

                                                    if (ini != "J")
                                                    {
                                                        ini = "J";
                                                        mid = nif.Substring(0, nif.Length - 1);
                                                    }
                                                    else
                                                        mid = nif.Substring(1, nif.Length - 2);

                                                    string fin = nif.Substring(nif.Length - 1, 1);
                                                    nif = ini + "-" + mid + "-" + fin;
                                                }
                                            }
                                            break;
                                        case "US":
                                        default:
                                            break;
                                    }
                                }
                            }
                            #endregion
                            #region Teléfono 

                            if (phone != "")
                            {
                                phone = val.EditPhone(phone);
                                if (phone == "err")
                                {
                                    returnReq = true;
                                    itemRes = "teléfono del cliente (" + customerID + " - " + phone + ") inválido";
                                }
                            }

                            #endregion

                            #endregion

                            if (returnReq)
                                returnReq = false;
                            else
                            {
                                #region SAP
                                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                                try
                                {
                                    Dictionary<string, string> parameters = new Dictionary<string, string>
                                    {
                                        ["BPFINAL"] = customerID,
                                        ["ADDRFINAL"] = address,
                                        ["EMAILFINAL"] = email,
                                        ["NAMEFINAL"] = customerName,
                                        ["TAXNFINAL"] = nif,
                                        ["TAXNFINAL2"] = nif2,
                                        ["PHONEFINAL"] = phone,
                                        ["FIRSTNAMEFINAL"] = firstName,
                                        ["LASTNAMEFINAL"] = lastName
                                    };

                                    IRfcFunction zdmChangeDatag = new SapVariants().ExecuteRFC(mandErp, "ZDM_CHANGE_DATAG", parameters);

                                    #region Procesar Salidas del FM

                                    if (zdmChangeDatag.GetValue("RESULTADO_ADDRESS").ToString() != "")
                                        fmRes = zdmChangeDatag.GetValue("RESULTADO_ADDRESS").ToString() + " - ";

                                    if (zdmChangeDatag.GetValue("RESULTADO_EMAIL").ToString() != "")
                                        fmRes = fmRes + zdmChangeDatag.GetValue("RESULTADO_EMAIL").ToString() + " - ";

                                    if (zdmChangeDatag.GetValue("RESULTADO_NAME").ToString() != "")
                                        fmRes = fmRes + zdmChangeDatag.GetValue("RESULTADO_NAME").ToString() + " - ";

                                    if (zdmChangeDatag.GetValue("RESULTADO_TAX1").ToString() != "")
                                        fmRes = fmRes + zdmChangeDatag.GetValue("RESULTADO_TAX1").ToString() + " - ";

                                    if (zdmChangeDatag.GetValue("RESULTADO_PHONE").ToString() != "")
                                        fmRes += zdmChangeDatag.GetValue("RESULTADO_PHONE").ToString();

                                    itemRes = fmRes;
                                    console.WriteLine(itemRes);

                                    //log de base de datos
                                    log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Modificar Data General BP", itemRes, root.Subject);
                                    respFinal = respFinal + "\\n" + "Modificar Data General BP: " + itemRes;

                                    if (itemRes.Contains("Error"))
                                        validateLines = false;

                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    validateLines = false;
                                    itemRes = ex.Message;
                                    console.WriteLine("Finishing process " + itemRes);
                                }
                                #endregion
                            }

                            item["Resultado"] = itemRes;
                        }
                    }
                }

                console.WriteLine("Respondiendo solicitud");

                string htmlTable = val.ConvertDataTableToHTML(excelDt);

                if (validateLines == false)
                    mail.SendHTMLMail("Uno o varios dieron error debido a:<br>" + htmlTable, new string[] { root.BDUserCreatedBy }, root.Subject, new string[] { "internalcustomersrvs@gbm.net" });
                else
                    mail.SendHTMLMail(htmlTable, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
                root.requestDetails = respFinal;

            }
            else
            {
                console.WriteLine("Devolviendo Solicitud");
                mail.SendHTMLMail("Por favor utilizar la plantilla oficial de datos maestros, sin modificar.<br>", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }


        }
    }
}
