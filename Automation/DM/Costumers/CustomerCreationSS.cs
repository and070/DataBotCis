using DataBotV5.Data.Projects.MasterData;
using Newtonsoft.Json.Linq;
using SAP.Middleware.Connector;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Exceptions;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.Logical.Webex;

using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Logical.MicrosoftTools;
using System.Data;
using DataBotV5.Data.Database;

namespace DataBotV5.Automation.DM.Costumers

{
    /// <summary>
    /// Clase DM Automation encargada de la creación de clientes en Datos Maestros.
    /// </summary>
    class CustomerCreationSS
    {
        
        ProcessInteraction proc = new ProcessInteraction();
        Stats stats = new Stats();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        WebInteraction web = new WebInteraction();
        MasterDataSqlSS DM = new MasterDataSqlSS();
        ValidateData vl = new ValidateData();
        Credentials cred = new Credentials();
        Rooting root = new Rooting();
        Log log = new Log();
        WebexTeams wt = new WebexTeams();
        MsExcel ms = new MsExcel();

        int[] campos = { 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 21, 23, 25, 26, 27, 28, 29, 30, 31, 32, 33, 35 };
        int[] campos2 = { 2, 3, 4, 5, 7, 35, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 22, 24, 25, 26, 27, 28, 29, 30, 31, 6, 11 };
        bool errorFm = false, returnRequest = false;
        string swap, bp = "", returnMsg = "", res1 = "", res2 = "";

        string erpMand = "ERP";
        string crmMand = "CRM";



        string customerTitle;
        string customerName;
        string salesRep;
        string nif;
        string nit2;
        string customerAddress;
        string customerCity;
        string customerName2;
        string postalCode;
        string customerCountry;
        string region;
        string subregion;
        string customerPhone;
        string customerEmail;
        string industry;
        string giroNegocio;
        string idFactElect;
        string gc5;
        string salesOrg;
        string contactTitle;
        string contactName;
        string contactLastName;
        string contactCountry;
        string contactAdress;
        string contactEmail;
        string contactPhone;
        string contactLang;
        string gc1 = "", channel = "", iva = "";
        string economicActivity = "";

        string respFinal = "";



        public void Main()
        {
            string respuesta = DM.GetManagement("2"); //CLIENTES
            if (!String.IsNullOrEmpty(respuesta) && respuesta != "ERROR")
            {
                console.WriteLine("Procesando...");
                ProcessClients();
                console.WriteLine("Creando Estadisticas");
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void ProcessClients()
        {
            try
            {
                #region Extraer datos generales (cada clase ya que es data muy personal de la solicitud)
                JArray DG = JArray.Parse(root.datagDM);
                for (int i = 0; i < DG.Count; i++)
                {
                    JObject fila = JObject.Parse(DG[i].ToString());
                    string pais_solicitud = fila["sendingCountryCode"].Value<string>(); //PAIS
                    gc1 = fila["valueTeamCode"].Value<string>();
                    channel = fila["channelCode"].Value<string>();
                    iva = fila["subjectVatCode"].Value<string>();
                }
                #endregion

                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                if (root.metodoDM == "1") //LINEAL
                {
                    #region PROCESAR LINEAL
                    JArray gestiones = JArray.Parse(root.requestDetails);
                    for (int i = 0; i < gestiones.Count; i++)
                    {
                        JObject fila = JObject.Parse(gestiones[i].ToString());

                        customerTitle = fila["generalTreatmentCode"].Value<string>().Trim();
                        customerName = fila["businessName"].Value<string>().Trim().ToUpper();
                        salesRep = "AA" + fila["salesRepresentativeCode"].Value<string>().Trim().PadLeft(8, '0');
                        nif = fila["identificationCard"].Value<string>().ToUpper().Trim();
                        nit2 = fila["nit"].Value<string>().Trim();
                        customerAddress = fila["address"].Value<string>().Trim().ToUpper();
                        customerCity = fila["additionalAddress"].Value<string>().Trim();
                        customerName2 = fila["additionalAddress"].Value<string>().Trim().ToUpper();
                        postalCode = "";
                        customerCountry = fila["countryCode"].Value<string>().Trim();

                        region = (fila["regionCode"].Value<string>().Trim() == "") ? fila["otherRegion"].Value<string>().Trim() : fila["regionCode"].Value<string>().Trim();


                        subregion = fila["subRegionCode"].Value<string>().Trim();
                        customerPhone = fila["phone"].Value<string>().Trim();
                        customerEmail = fila["email"].Value<string>().Trim();
                        industry = fila["branchCode"].Value<string>().Trim();
                        giroNegocio = fila["businessLine"].Value<string>().Trim();
                        idFactElect = fila["clientTypeCode"].Value<string>().Trim();
                        salesOrg = fila["salesOrganizationsCode"].Value<string>().Trim();
                        gc5 = fila["customerGroupCode"].Value<string>().Trim();


                        economicActivity = fila["economicActivity"].Value<string>().Trim();

                        contactTitle = fila["contactTreatmentCode"].Value<string>().Trim();
                        contactName = fila["name"].Value<string>().Trim().ToUpper();
                        contactLastName = fila["lastName"].Value<string>().Trim().ToUpper();
                        contactCountry = fila["countryContactCode"].Value<string>().Trim();
                        contactAdress = fila["addressContact"].Value<string>().Trim().ToUpper();
                        contactEmail = fila["emailContact"].Value<string>().Trim();
                        contactPhone = fila["phoneContact"].Value<string>().Trim();
                        contactLang = fila["languageCode"].Value<string>().Trim();

                        VerifyAndCreate();
                    }
                    #endregion
                }
                else
                {
                    #region PROCESAR MASIVO

                    string adjunto = root.ExcelFile;

                    if (!String.IsNullOrEmpty(adjunto))
                    {
                        console.WriteLine("Abriendo excel y validando");

                        #region abrir excel

                        DataTable xlWorkSheet = ms.GetExcel(root.FilesDownloadPath + "\\" + adjunto);

                        #endregion

                        //campos = campos2;
                        foreach (DataRow row in xlWorkSheet.Rows)
                        {
                            if (row["Nombre Razón Social"].ToString().Trim() != "")
                            {
                                #region Llenar variables

                                customerTitle = row["Tratamiento"].ToString().Trim();
                                customerName = row["Nombre Razón Social"].ToString().Trim();
                                salesRep = row["Representante de Ventas"].ToString().Trim();
                                nif = row["NIF / NRC"].ToString().ToUpper().Trim();
                                customerAddress = row["Direccion"].ToString().Trim();
                                customerName2 = row["Direccion"].ToString().Trim();
                                postalCode = row["Codigo Postal"].ToString().Trim();
                                customerCountry = row["Pais"].ToString().Trim();
                                region = row["Region"].ToString().Trim();
                                customerPhone = row["Telefono"].ToString().Trim();
                                customerEmail = row["Email"].ToString().Trim();
                                industry = row["Ramo"].ToString().Trim();
                                giroNegocio = row["Giro de Negocio"].ToString().Trim();
                                iva = row["Sujeto a IVA"].ToString().Trim();
                                salesOrg = row["Organizacion de ventas"].ToString().Trim();
                                gc1 = row["Grupo Ctel"].ToString().Trim();
                                gc5 = row["Customer Group 5"].ToString().Trim();
                                contactTitle = row["Tratamiento Contacto"].ToString().Trim();
                                contactName = row["Nombre"].ToString().Trim();
                                contactLastName = row["Apellido"].ToString().Trim();
                                contactCountry = row["Pais Contacto"].ToString().Trim();
                                contactAdress = row["Dirección Contacto"].ToString().Trim();
                                contactEmail = row["Email Contacto"].ToString().Trim();
                                contactPhone = row["Telefono Contacto"].ToString().Trim();
                                contactLang = row["Idioma"].ToString().Trim();
                                nit2 = row["NIT 2"].ToString().ToUpper().Trim();
                                customerCity = row["Ciudad"].ToString().Trim();

                                #endregion

                                #region Validación de datos


                                if (customerTitle == "" || customerName == "" || nif == "")
                                {
                                    returnMsg = "Por favor ingresar los campos obligatorios";
                                    res2 = res2 + returnMsg + "<br>";
                                    returnRequest = true;
                                    continue;
                                }

                                industry = vl.ExtractFieldWithSeparator("-", industry);
                                customerTitle = vl.ExtractFieldWithSeparator("-", customerTitle);
                                salesRep = vl.ExtractFieldWithSeparator("-", salesRep);
                                customerCountry = vl.ExtractFieldWithSeparator("_", customerCountry);
                                region = vl.ExtractFieldWithSeparator("-", region);
                                iva = vl.ExtractFieldWithSeparator("-", iva);
                                salesOrg = vl.ExtractFieldWithSeparator("-", salesOrg);
                                gc1 = vl.ExtractFieldWithSeparator("-", gc1);
                                gc5 = vl.ExtractFieldWithSeparator("-", gc5);
                                contactTitle = vl.ExtractFieldWithSeparator("-", contactTitle);
                                contactCountry = vl.ExtractFieldWithSeparator("_", contactCountry);
                                contactLang = vl.ExtractFieldWithSeparator("-", contactLang);

                                #endregion

                                VerifyAndCreate();
                            }
                        }

                    }
                    else
                    {
                        returnRequest = true;
                        res2 = "Error en la plantilla";
                    }


                    #endregion
                }

                console.WriteLine("Finalizando solicitud");

                #region Enviar Notificaciones
                if (errorFm == true && !res2.Contains("cliente existe"))
                {
                    string[] cc = { "lrrojas@gbm.net" };
                    //enviar email de repuesta de error a datos maestros
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Clientes:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazado, con el siguiente resultado: <br><br> " + res1 + "<br>" + res2);
                    mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + res1 + "<br>" + res2, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
                    DM.ChangeStateDM(root.IdGestionDM, res1 + "<br>" + res2, "5"); //RECHAZADO 
                }
                else if (res1.Contains("invalido"))
                {
                    console.WriteLine("Devolviendo solicitud");
                    DM.ChangeStateDM(root.IdGestionDM, res1 + "<br>" + res2, "5"); //RECHAZADO 
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Clientes:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazado, con el siguiente resultado: <br><br> " + res1 + "<br>" + res2);
                }
                else
                {
                    //finalizar solicitud
                    DM.ChangeStateDM(root.IdGestionDM, res1 + "<br>" + res2, "3"); //FINALIZADO
                    if (res1 == "")
                    {
                        mail.SendHTMLMail(res1 + "<br>" + res2, new string[] { "smarin@gbm.net" }, root.Subject + "***Se finalizo sin respuesta***");
                    }
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Clientes:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + res1 + "<br>" + res2);

                }
                #endregion

                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4");
                mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, new string[] { "smarin@gbm.net" });
            }
        }

        private bool ActivateCrmGui(string bp)
        {
            SapVariants sap = new SapVariants();
            bool correcto = true;

            //revisa si el usuario RPAUSER esta abierto
            bool check_login = sap.CheckLogin(crmMand);
            if (!check_login)
            {
                sap.BlockUser(crmMand, 1);
                try
                {
                    proc.KillProcess("saplogon", false);
                    sap.LogSAP(crmMand.ToString());
                    SapVariants.frame.Iconify();
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nY_DM_BP_001";
                    SapVariants.frame.SendVKey(0);
                    ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkP_TEST")).Selected = false;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtBP_NO-LOW")).Text = bp;
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                    string status2 = ((SAPFEWSELib.GuiLabel)SapVariants.session.FindById("wnd[0]/usr/lbl[51,2]")).Text.ToString();
                    status2 = status2.Substring(status2.Length - 1, 1);
                    ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/n";
                    SapVariants.frame.SendVKey(0);
                    correcto = status2 != "1"; //correcto = status2 == "1" ? false : true;
                }
                catch (Exception) { }

                sap.BlockUser(crmMand, 0);
                sap.KillSAP();
            }

            return correcto;
        }
        private string EmployeeToCustomer(string employee)
        {
            employee = employee.ToUpper();

            string response;
            try
            {
                Dictionary<string, string> zdmEmployeeToCustomerParameters = new Dictionary<string, string>();
                Dictionary<string, string> zhrSyncPersonFmParameters = new Dictionary<string, string>();

                zdmEmployeeToCustomerParameters["BP"] = employee;
                zhrSyncPersonFmParameters["ID"] = employee.Replace("AA", "");

                IRfcFunction zhrSyncPersonFm = new SapVariants().ExecuteRFC(erpMand, "ZHR_SYNC_PERSON_FM", zhrSyncPersonFmParameters);
                IRfcFunction zdmEmployeeToCustomer = new SapVariants().ExecuteRFC(erpMand, "ZDM_EMPLOYEE_TO_CUSTOMER", zdmEmployeeToCustomerParameters);

                string bp = zdmEmployeeToCustomer.GetValue("BP_OUT").ToString();
                string RESULTADO = zdmEmployeeToCustomer.GetValue("RESULTADO").ToString();
                string MENSAJE1 = zdmEmployeeToCustomer.GetValue("MENSAJE1").ToString();
                string TAX_WARR = zdmEmployeeToCustomer.GetValue("TAX_WARR").ToString();

                if (MENSAJE1 == "" && RESULTADO == "")
                    response = bp.ToString();
                else
                    response = "Error " + bp + ": " + RESULTADO + MENSAJE1;
            }
            catch (Exception ex)
            {
                response = "Exception: " + ex.Message;
            }

            return response;
        }
        public void VerifyAndCreate()
        {
            #region Validación de datos

            #region Ramo
            if (customerTitle == "0003" && industry == "PE01")
            {
                returnRequest = true;
                returnMsg += "El Ramo de una empresa no puede ser PE01, favor revisar<br>";
            }
            #endregion

            #region Razón social
            customerName = customerName.ToUpper();
            if (customerName.Length > 80)
            {
                customerName = customerName.Substring(0, 80);
            }
            customerName = vl.RemoveSpecialChars(customerName, 1);
            customerName = vl.RemoveChars(customerName);
            #endregion

            #region Dirección

            if (customerAddress == customerName2)
            {
                customerName2 = "";
            }

            if (customerName2 == "Vacio" || customerName2 == ".")
            {
                customerName2 = "";
            }
            else
            {
                customerName2 = customerName2.ToUpper();
                customerName2 = vl.RemoveSpecialChars(customerName2, 1);
                customerName2 = vl.RemoveChars(customerName2);
            }


            customerAddress = vl.RemoveSpecialChars(customerAddress, 1);
            customerAddress = customerAddress.ToUpper();

            customerAddress = customerAddress + " " + customerName2;

            if (customerAddress.Length > 140)
            {
                customerAddress = customerAddress.Substring(0, 140);
            }
            customerAddress = vl.RemoveChars(customerAddress);
            #endregion

            #region NIF
            if (nit2 == "N/A" || nit2 == "NA") { nit2 = ""; }
            if (nif == nit2) { nit2 = ""; }
            if (nit2 == "00" || nit2 == "000" || nit2 == "-") { nit2 = ""; }
            if (nif.Contains("EIN Number"))
            {
                nif = nif.Replace("EIN Number", "");
            }
            if (nif.Length < nit2.Length)
            {
                swap = nif; nif = nit2; nit2 = swap;
            }

            if (nif == "N/A" || nif == "NA" || nif == "")
            {
                returnRequest = true;
                returnMsg = returnMsg + "NIF del cliente (" + customerName + " - " + nif + ") inválido//" + "<br>";
            }


            switch (customerCountry)
            {
                case "SV":
                    if (nif.Length >= 17)
                    {
                        nif = nif.Substring(0, nif.Length - 2) + nif.Substring(nif.Length - 1, 1);
                    }
                    nit2 = Regex.Replace(nit2, "[^0-9-]", "");
                    postalCode = "";
                    break;
                case "PA":
                    int dv;
                    dv = nif.IndexOf("DV");
                    if (dv > 0)
                    {
                        nit2 = nif.Substring(dv, nif.Length - dv);
                        nif = nif.Substring(0, dv - 1);
                    }
                    nit2 = Regex.Replace(nit2, @"[^\d]", "");
                    if (nit2.Length > 3)
                    {
                        returnRequest = true;
                        returnMsg = returnMsg + "NIF2 (" + customerName + " - " + nit2 + ") invalido//" + "<br>";
                    }
                    postalCode = "";
                    break;
                case "CR":
                    nif = nif.Replace("-", "");
                    if (vl.IsNum(nif) == false)
                    {
                        returnRequest = true;
                        returnMsg = returnMsg + "NIF (" + customerName + " - " + nif + ") invalido//" + "<br>";
                    }
                    postalCode = "";
                    break;
                case "DO":
                    nif = nif.Replace("-", "");
                    if (vl.IsNum(nif) == false)
                    {
                        returnRequest = true;
                        returnMsg = returnMsg + "NIF (" + customerName + " - " + nif + ") invalido//" + "<br>";
                    }
                    postalCode = "";
                    break;
                case "HN":
                    postalCode = "";
                    break;
                case "NI":
                    nif = nif.Replace("-", "");
                    postalCode = "";
                    break;
                case "CO":
                    postalCode = "";
                    nif = vl.RemoveSpecialChars(nif, 2);
                    nit2 = "";
                    break;
                case "VE":
                    postalCode = "";
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
                            {
                                mid = nif.Substring(1, nif.Length - 2);
                            }
                            string fin = nif.Substring(nif.Length - 1, 1);

                            nif = ini + "-" + mid + "-" + fin;
                        }
                    }
                    break;
                case "US":
                default:
                    if (postalCode == "")
                    {
                        postalCode = "11111";
                    }
                    contactCountry = salesOrg.Substring(salesOrg.Length - 2, 2);
                    break;
            }
            #endregion

            #region Teléfono

            customerPhone = vl.EditPhone(customerPhone);
            if (customerPhone == "err")
            {
                returnRequest = true;
                returnMsg = returnMsg + "Teléfono del cliente (" + customerName + " - " + customerPhone + ") inválido//" + "<br>";
            }
            #endregion

            #region giro
            giroNegocio = giroNegocio.ToUpper();
            if (giroNegocio == "N/A")
            {
                giroNegocio = "";
            }
            giroNegocio = vl.RemoveSpecialChars(giroNegocio, 1);
            giroNegocio = vl.RemoveChars(giroNegocio);
            if (giroNegocio.Length > 132)
            {
                giroNegocio = giroNegocio.Substring(0, 132);
            }
            #endregion

            #region customer group
            if (gc1 == "002" && gc5 == "PRE")
            {
                gc5 = "A";
            }
            #endregion

            #region Contacto
            contactName = contactName.ToUpper();
            contactName = vl.RemoveSpecialChars(contactName, 2);

            contactLastName = contactLastName.ToUpper();
            contactLastName = vl.RemoveSpecialChars(contactLastName, 2);

            contactAdress = contactAdress.ToUpper();
            contactAdress = vl.RemoveSpecialChars(contactAdress, 1);
            if (contactAdress.Length > 60)
            { contactAdress = contactAdress.Substring(0, 60); }

            contactEmail = contactEmail.ToLower().Trim();
            if (vl.ValidateEmail(contactEmail) != true)
            {
                returnRequest = true;
                returnMsg = returnMsg + "Email del contacto (" + customerName + " - " + contactEmail + ") inválido//" + "<br>";
            }
            contactEmail = vl.RemoveEnne(contactEmail);

            contactPhone = vl.EditPhone(contactPhone);
            if (contactPhone == "err")
            {
                returnRequest = true;
                returnMsg = returnMsg + "Teléfono del contacto (" + customerName + " - " + contactPhone + ") inválido//";
            }
            #region pais_contacto

            string[] main_ctr = { "CR", "DO", "GT", "NI", "HN", "PA", "SV", "US", "VE" };
            if (!main_ctr.Contains(contactCountry))
            {
                switch (salesOrg)
                {
                    case "GBCR":
                        contactCountry = "CR";
                        break;
                    case "GBDR":
                        contactCountry = "DO";
                        break;
                    case "GBGT":
                        contactCountry = "GT";
                        break;
                    case "GBHN":
                        contactCountry = "HN";
                        break;
                    case "GBMD":
                        contactCountry = "US";
                        break;
                    case "GBNI":
                        contactCountry = "NI";
                        break;
                    case "GBPA":
                        contactCountry = "PA";
                        break;
                    case "GBSV":
                        contactCountry = "SV";
                        break;
                    case "ITC0":
                        contactCountry = "US";
                        break;
                    case "WTC0":
                        contactCountry = "US";
                        break;
                    case "GBCO":
                        contactCountry = "CO";
                        break;
                    case "BV01":
                        contactCountry = "US";
                        break;
                    case "LCFL":
                        contactCountry = "US";
                        break;
                    case "LCVE":
                        contactCountry = "VE";
                        break;
                }
            }
            #endregion

            #endregion

            #region Región
            if (region.Length > 3)
            {
                region = "";
            }
            #endregion

            #region Representante

            if (salesRep == "AA00000000")
            {
                salesRep = "";
            }
            if (salesOrg == "GBNI" && industry == "PE01" && salesRep != "AA30000608")
            {
                salesRep = "";
            }

            #endregion

            #region ciudad
            customerCity = customerCity.ToUpper();
            if (customerCity == "VACIO")
            {
                customerCity = "";
            }

            if (customerCity.Length > 40)
            {
                customerCity = customerCity.Substring(0, 40);
            }
            customerCity = vl.RemoveSpecialChars(customerCity, 1);


            #endregion

            #region email
            customerEmail = customerEmail.ToLower().Trim();
            customerEmail = vl.RemoveChars(customerEmail);
            if (vl.ValidateEmail(customerEmail) != true)
            {
                returnRequest = true;
                returnMsg = returnMsg + "Email del cliente (" + customerName + " - " + customerEmail + ") invalido//" + "<br>";
            }
            customerEmail = vl.RemoveEnne(customerEmail);
            #endregion

            if (returnRequest == true)
            {
                res1 = res1 + returnMsg + "<br>";
                returnRequest = false;
                return;
            }

            #endregion

            #region SAP
            console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
            try
            {
                RfcDestination destination = new SapVariants().GetDestRFC(erpMand);
                IRfcFunction func = destination.Repository.CreateFunction("ZDM_CREATE_CUSTOMER");
                IRfcTable telephones = func.GetTable("TELEFONO_");
                telephones.Clear();

                #region Parámetros SAP
                func.SetValue("TRATAMIENTO", customerTitle);
                func.SetValue("RAZON", customerName);
                func.SetValue("REPR", salesRep);
                func.SetValue("NIF", nif);
                func.SetValue("NIF2", nit2);
                func.SetValue("ADDRESS", customerAddress);
                func.SetValue("POSTAL", postalCode);
                func.SetValue("PAIS", customerCountry);
                if (region != "" && region != "Vacio")
                    func.SetValue("REGION", region);
                func.SetValue("SUBREGION", subregion);
                func.SetValue("CONTRIBUYENTE", idFactElect);
                func.SetValue("CIUDAD", customerCity);
                func.SetValue("TEL", customerPhone);
                func.SetValue("EMAIL", customerEmail);
                func.SetValue("RAMO", industry);
                func.SetValue("GIRO", giroNegocio);
                func.SetValue("IVA", iva);
                func.SetValue("SALES_ORG", salesOrg);
                func.SetValue("GROUP1", gc1);
                func.SetValue("GROUP5", gc5);
                func.SetValue("TRATAMIENTO_", contactTitle);
                func.SetValue("NOMBRE_", contactName);
                func.SetValue("APELLIDO_", contactLastName);
                func.SetValue("PAIS_", contactCountry);
                func.SetValue("DIRECCION_", contactAdress);
                func.SetValue("CORREO_", contactEmail);
                telephones.Append();
                telephones.SetValue("TELEPHONE", contactPhone);
                func.SetValue("IDIOMA_", contactLang);
                #endregion

                #region Invocar FM
                func.Invoke(destination);

                #endregion

                #region Procesar Salidas del FM
                bp = func.GetValue("BP").ToString();
                string msg1 = func.GetValue("MENSAJE1").ToString();
                string msg = func.GetValue("MENSAJE").ToString();

                if (msg1 == "cliente existe")
                {
                    if (bp != "")
                    {
                        if (bp.Substring(0, 2) == "AA")
                        {
                            console.WriteLine("El cliente: " + customerName + " ya existe: " + bp + " Se ampliará");
                            string employeeRes = EmployeeToCustomer(bp);
                            if (employeeRes.Contains("Exception") || employeeRes.Contains("Error"))
                            {
                                errorFm = true;
                                msg1 = employeeRes;
                            }
                            else { msg1 = ""; }
                        }
                    }
                }

                if (msg1 == "")
                {
                    console.WriteLine("El cliente: " + customerName + " ha sido creado con el ID: " + bp);
                    //log de cambios base de datos

                    res1 = res1 + "El cliente: " + customerName + " ha sido creado con el ID: " + bp + "<br>";
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear cliente", " El cliente: " + customerName + " ha sido creado con el ID: " + bp, root.Subject);
                    respFinal = respFinal + "\\n" + "El cliente: " + customerName + " ha sido creado con el ID: " + bp;

                    //crea el cliente en la bd del databot
                    try
                    {
                        log.RegisterNeewClient(int.Parse(bp), customerName.Replace("'", ""), customerCountry, gc1, salesRep, salesOrg, customerAddress, customerPhone, customerEmail);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            Errores errs = new Errores();
                            Exceptions exep = new Exceptions();
                            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                            exep.ExceptionsFormat(trace, errs, ex.ToString());
                        }
                        catch (Exception)
                        {
                            wt.SendNotification("dmeza@gbm.net", "", $"Error al ingresar cliente dentro la DB <br><br> {ex}");
                        }

                    }

                }
                else if (msg1 == "crm_err")
                {
                    errorFm = ActivateCrmGui(bp);
                }
                else
                {
                    errorFm = true;
                    res1 = res1 + bp + ": " + msg;
                    res2 = res2 + bp + ": " + msg1;
                }
                #endregion
            }
            catch (Exception ex)
            {
                res2 = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0);
                console.WriteLine("Finishing process " + res2);
                res1 = res1 + "Cliente: " + customerName + ": " + ex.ToString() + "<br>";
                res2 = ex.ToString();
                errorFm = true;
            }

            #endregion

        }

        private bool updateClientsSs()
        {
            try
            {
                DataTable resp = new DataTable();
                resp.Columns.Add("cliente");
                resp.Columns.Add("sql");
                resp.Columns.Add("resp");
                CRUD crud = new CRUD();
                SapVariants sap = new SapVariants();
                DataTable dt = crud.Select("SELECT * FROM clients", "databot_db", "QAS");
                foreach (DataRow item in dt.Rows)
                {
                    DataRow rRow = resp.Rows.Add();
                    try
                    {
                        Dictionary<string, string> parameters = new Dictionary<string, string>
                        {
                            ["BP"] = item["idClient"].ToString(),
                        };
                        IRfcFunction fm = sap.ExecuteRFC(erpMand, "ZDM_READ_BP", parameters);

                        string name = fm.GetValue("NOMBRE").ToString();
                        string add = fm.GetValue("ADDRESS").ToString() + " " + fm.GetValue("COMPLADDRESS").ToString();
                        string salesRep = fm.GetValue("SALESREP").ToString().Replace("AA", "");
                        string phone = fm.GetValue("PHONE").ToString();
                        string mail = fm.GetValue("EMAIL").ToString();

                        string upQuery = $@"UPDATE clients SET
                                            name = '{name}',
                                            accountManagerId = '{salesRep}', 
                                            accountManagerUser = (SELECT MIS.digital_sign.user from MIS.digital_sign WHERE MIS.digital_sign.UserID = '{salesRep}'), 
                                            address = '{add}', 
                                            telephone = '{phone}', 
                                            email = '{mail}', 
                                            updatedAt = CURRENT_TIMESTAMP,
                                            updatedBy = 'DMEZA'
                                            WHERE id = {item["id"]}";
                        bool upda = crud.Update(upQuery, "databot_db", "QAS");
                        rRow["cliente"] = item["idClient"].ToString();
                        rRow["sql"] = upQuery;
                        rRow["resp"] = upda;

                    }
                    catch (Exception)
                    {

                    }
                }
                resp.AcceptChanges();
                ms.CreateExcel(resp, "sheet1", root.FilesDownloadPath + "\\" + "respClients.xlsx");
            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message);
            }

            return false;
        }
    }
}

