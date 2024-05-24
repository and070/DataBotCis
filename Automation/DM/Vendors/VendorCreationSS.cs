using DataBotV5.Data.Projects.MasterData;
using System.Text.RegularExpressions;
using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using SAP.Middleware.Connector;
using DataBotV5.Logical.Webex;
using DataBotV5.Data.Database;
using DataBotV5.Logical.Mail;
using DataBotV5.App.Global;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Stats;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Data;
using System;
using DataBotV5.App.ConsoleApp;

namespace DataBotV5.Automation.DM.Vendors
{
    class VendorCreationSS
    {
        #region Variables Globales
        readonly MailInteraction mail = new MailInteraction();
        readonly ConsoleFormat console = new ConsoleFormat();
        readonly MasterDataSqlSS DM = new MasterDataSqlSS();
        readonly ValidateData val = new ValidateData();
        readonly SapVariants sap = new SapVariants();
        readonly WebexTeams wt = new WebexTeams();
        readonly Rooting root = new Rooting();
        readonly CRUD crud = new CRUD();
        readonly Log log = new Log();

        const string ssMandante = "QAS";
        const string erpMand = "ERP";
        readonly string[] cc = { "hlherrera@gbm.net" };

        string resFailure = "";
        string resDms = "";
        string respFinal = "";

        #endregion

        [STAThread]
        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
            string res1;
            if (!sap.CheckLogin(erpMand))
            {
                //procesar las solicitudes "EN REVISION"
                res1 = DM.GetManagement("9", "12");
                if (!String.IsNullOrEmpty(res1) && res1 != "ERROR")
                {

                    console.WriteLine("Procesando...");
                    ProcessVendor("check");

                    console.WriteLine("Creando Estadísticas");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }

                //Procesar solicitudes "EN PROCESO"
                res1 = DM.GetManagement("9"); //PROVEEDORES
                if (!String.IsNullOrEmpty(res1) && res1 != "ERROR")
                {
                    console.WriteLine("Procesando...");
                    sap.BlockUser(erpMand, 1);
                    ProcessVendor("create");
                    sap.BlockUser(erpMand, 0);

                    console.WriteLine("Creando Estadísticas");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }
        [STAThread]

        public void ProcessVendor(string function)
        {
            try
            {
                #region Variables Privadas

                resFailure = resDms = "";

                string vendorTitle, vendorName = "", nif, nif2, vendorAddress = "", postalCode, vendorCountry, vendorRegion,
                 vendorPhone = "", vendorEmail, vendorActivity, currency, bankCode, bankAccount, bankAccountUs = "", localAccount,
                 dolarAccount, checkAccount, controlCode, companyCode = "", paymentMethod = "", salesOrg = "", taxCode = "",
                 vendorType, vendorGroup = "", vendorDesc, vendorCat, fmRes, fmMsg1, fmMsg, generalFolderId, vendorId = "", iban = "",
                 docNumP = "", vendorName2 = "", dbAdd = "", msgSap = "", clipData = "", swap, msg = "";

                string res2 = "", returnMsg = "", response = "", resAuth = "", resDmsFiles = "";

                bool returnRequest = false, revision = false, createWithError = false, addAccount = false;
                bool validateData = true;

                JArray vendorRequests = new JArray();
                #endregion

                #region extraer datos generales (cada clase ya que es data muy personal de la solicitud 

                JArray generalData = JArray.Parse(root.datagDM);
                for (int i = 0; i < generalData.Count; i++)
                {
                    JObject row = JObject.Parse(generalData[i].ToString());
                    companyCode = row["companyCodeCode"].Value<string>().Trim();
                    vendorGroup = row["vendorGroupCode"].Value<string>().Trim();
                }
                #endregion

                root.requestDetails = root.requestDetails.Replace("\u00A0", " "); //eliminar non breaks spaces (char 160)
                root.requestDetails = root.requestDetails.Replace(@"[^\u0000-\u007F]+", ""); //eliminar caracteres no ASCII

                if (root.metodoDM == "1") // Proveedores solo tiene Lineal
                    vendorRequests = JArray.Parse(root.requestDetails);

                for (int i = 0; i < vendorRequests.Count; i++)
                {
                    response = "";
                    resAuth = "";
                    resDmsFiles = "";
                    resDms = "";

                    JObject row = JObject.Parse(vendorRequests[i].ToString());

                    vendorName = row["socialReason"].Value<string>().Trim().ToUpper();
                    if (vendorName != "")
                    {
                        console.WriteLine("Solicitud: " + row["requestId"].Value<string>().Trim());

                        vendorTitle = row["generalTreatmentCode"].Value<string>().Trim();
                        nif = row["nif"].Value<string>().Trim();
                        nif2 = row["nif2"].Value<string>().Trim();
                        vendorAddress = row["address"].Value<string>().Trim().ToUpper();
                        vendorName2 = row["additionalAddress"].Value<string>().Trim().ToUpper();
                        postalCode = "";
                        vendorCountry = row["countryCode"].Value<string>().Trim();

                        vendorRegion = (row["regionCode"].Value<string>().Trim() == null) ? row["otherRegion"].Value<string>().Trim() : row["regionCode"].Value<string>().Trim();

                        vendorPhone = row["phone"].Value<string>().Trim();
                        vendorEmail = row["email"].Value<string>().Trim();
                        vendorActivity = (row["lineOfBusiness"].Value<string>().ToString().Trim() != "") ? row["lineOfBusiness"].Value<string>().Trim() : "";
                        vendorCat = row["vendorCategoryCode"].Value<string>().Trim();
                        vendorDesc = DM.GetVendorDescription(vendorCat);
                        vendorType = row["vendorTypeCode"].Value<string>().Trim();
                        bankCode = row["bankNameCode"].Value<string>().Trim();
                        localAccount = row["localCurrencyAccount"].Value<string>().Trim();
                        dolarAccount = (row["usdCurrencyAccount"].Value<string>().ToString().Trim() != "") ? row["usdCurrencyAccount"].Value<string>().Trim() : "";
                        dbAdd = (row["additionalDetails"].Value<string>().ToString().Trim() != "") ? row["additionalDetails"].Value<string>().Trim() : "";

                        #region validación de datos

                        vendorName = vendorName.Replace("á", "a"); vendorName = vendorName.Replace("é", "e"); vendorName = vendorName.Replace("í", "i"); vendorName = vendorName.Replace("ó", "o"); vendorName = vendorName.Replace("ú", "u"); vendorName = vendorName.Replace("ñ", "n");
                        vendorAddress = val.RemoveSpecialChars(vendorAddress, 1);
                        vendorAddress = vendorAddress.ToUpper();
                        if (vendorAddress.Length > 140)
                            vendorAddress = vendorAddress.Substring(0, 140);
                        if (vendorTitle.Length > 4)
                            vendorTitle = vendorTitle.Substring(0, 4);
                        if (nif2 == "N/A")
                            nif2 = "";
                        if (nif2 == nif)
                            nif2 = "";
                        if (vendorCountry.Length > 2)
                            vendorCountry = vendorCountry.Substring(0, 2);
                        if (vendorRegion.Length > 3)
                            vendorRegion = "";
                        else if (vendorRegion.Length == 1)
                            vendorRegion = "0" + vendorRegion;

                        if (vendorAddress == vendorName2)
                            vendorName2 = "";

                        vendorAddress = vendorAddress + " " + vendorName2;

                        if (vendorAddress.Length > 140)
                            vendorAddress = vendorAddress.Substring(0, 140);

                        switch (vendorCountry)
                        {
                            case "SV":
                                if (nif.Length >= 17)
                                {
                                    nif = nif.Substring(0, nif.Length - 2) + nif.Substring(nif.Length - 1, 1);
                                }
                                nif2 = Regex.Replace(nif2, "[^0-9-]", "");
                                break;
                            case "PA":
                                int dv;
                                dv = nif.IndexOf("DV") + 1;
                                if (dv > 0)
                                {
                                    nif2 = nif.Substring(dv, nif.Length - dv);
                                    nif = nif.Substring(0, dv - 1);
                                }
                                nif2 = Regex.Replace(nif2, @"[^\d]", "");
                                if (nif.Length > 3)
                                {
                                    //nif2 invalido
                                }

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
                                        {
                                            mid = nif.Substring(1, nif.Length - 2);
                                        }
                                        string fin = nif.Substring(nif.Length - 1, 1);

                                        nif = ini + "-" + mid + "-" + fin;
                                    }
                                }
                                break;
                        }

                        if (nif2 == "N/A" || nif2 == "NA")
                            nif2 = "";
                        if (nif == nif2)
                            nif2 = "";

                        if (nif.Length < nif2.Length)
                        {
                            swap = nif;
                            nif = nif2;
                            nif2 = swap;
                        }

                        if (vendorPhone != "")
                        {
                            if (vendorPhone.Substring(0, 1) == "(")
                                vendorPhone = vendorPhone.Substring(5, vendorPhone.Length - 5);
                            vendorPhone.ToLower();
                            if (vendorPhone.Substring(0, 1) == "(" && vendorPhone.Substring(0, 2) == "+")
                                vendorPhone = vendorPhone.Substring(6, vendorPhone.Length - 6);
                            vendorPhone.ToLower();
                            if ((vendorPhone.IndexOf("ext") + 1) > 0)
                                vendorPhone = vendorPhone.Substring(0, vendorPhone.IndexOf("ext") - 1);

                            vendorPhone = vendorPhone.Replace("-", "");
                            vendorPhone = vendorPhone.Replace("+", "");
                            if (vendorPhone.Length > 30 | vendorPhone.Contains("tel") | vendorPhone.Contains("/") | vendorPhone.Contains(","))
                            {
                                returnMsg = vendorName + ": " + "Por favor ingresar el teléfono";
                                res2 = res2 + returnMsg + "<br>";
                                returnRequest = true;
                                continue;
                            }
                        }
                        else
                        {
                            returnMsg = vendorName + ": " + "Por favor ingresar el teléfono";
                            res2 = res2 + returnMsg + "<br>";
                            returnRequest = true;
                            continue;
                        }

                        if (localAccount == dolarAccount)
                            dolarAccount = "";

                        if (localAccount != "")
                        {
                            checkAccount = localAccount.Replace("0", "");

                            if (checkAccount == "")
                            {
                                localAccount = "";

                            }


                            if (localAccount.ToLower() == "n/a" || localAccount.ToLower() == "na")
                            {
                                localAccount = "";
                                if (dolarAccount != "")
                                {
                                    addAccount = true;
                                }
                            }
                        }

                        if (dolarAccount != "")
                        {
                            checkAccount = dolarAccount.Replace("0", "");

                            if (checkAccount == "")
                            {
                                dolarAccount = "";
                            }

                            if (dolarAccount.ToLower() == "n/a" || dolarAccount.ToLower() == "na")
                            {
                                dolarAccount = "";
                            }
                        }

                        if (localAccount != "" && dolarAccount != "")
                        {
                            addAccount = true;
                        }

                        if (localAccount != "" && addAccount == false)
                        {
                            currency = "ML";
                            controlCode = "ML";
                            bankAccount = localAccount;
                        }
                        else if (dolarAccount != "" && addAccount == false)
                        {
                            currency = "US";
                            controlCode = "US";
                            bankAccount = dolarAccount;
                        }
                        else if (addAccount == true)
                        {
                            currency = "ML";
                            controlCode = "ML";
                            bankAccount = localAccount;

                            bankAccountUs = dolarAccount;
                        }
                        else
                            currency = controlCode = bankAccount = bankCode = "";

                        vendorActivity = vendorActivity.ToUpper();
                        if (vendorActivity == "N/A" || vendorActivity == "NA")
                            vendorActivity = "";

                        vendorActivity = val.RemoveSpecialChars(vendorActivity, 1);
                        if (vendorActivity.Length > 132)
                            vendorActivity = vendorActivity.Substring(0, 132);

                        if (vendorRegion.Length > 3)
                            vendorRegion = "";

                        string expression = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";

                        if (Regex.IsMatch(vendorEmail, expression))
                        {
                            if (Regex.Replace(vendorEmail, expression, string.Empty).Length != 0)
                            {
                                returnMsg = vendorName + ": " + "Por favor ingresar un email válido";
                                res2 = res2 + returnMsg + "<br>";
                                returnRequest = true;
                                continue;
                            }
                        }

                        //if (companyCode != "GBMD")
                        //    vendorCity = "";
                        if (companyCode != "GBSV")
                            vendorActivity = "";

                        string account1 = "";
                        switch (companyCode)
                        {
                            case "GBCR":
                                salesOrg = "CR01";
                                account1 = bankAccount;

                                //cuenta sinpe o normal rechaza la solicitud
                                if (bankAccount.Length < 20 && bankAccount.Length > 0)
                                {
                                    returnMsg = vendorName + ": " + "Por favor ingrese las cuentas IBAN de 20 caracteres";
                                    res2 = res2 + returnMsg + "<br>";
                                    returnRequest = true;
                                    continue;
                                }

                                //IBAN sin CR
                                else if (bankAccount.Length == 20)
                                {
                                    if (bankAccount.Substring(0, 2) != "CR")
                                        iban = "CR" + bankAccount;
                                    else
                                        iban = bankAccount;
                                    bankAccount = account1.Substring(0, 17);
                                }
                                //IBAN normal
                                else if (bankAccount.Length == 22)
                                {
                                    if (bankAccount.Substring(0, 2) != "CR")
                                        iban = "CR" + bankAccount.Substring(0, bankAccount.Length - 2);
                                    else
                                        iban = bankAccount;
                                    bankAccount = account1.Substring(0, 17);
                                }
                                //no agrega cuenta
                                else
                                    bankAccount = iban = "";

                                break;

                            case "GBDR":
                                salesOrg = "DR01";
                                break;
                            case "GBGT":
                                salesOrg = "GT01";
                                break;
                            case "GBHN":
                                salesOrg = "HN01";
                                break;
                            case "GBHQ":
                                salesOrg = "CR01";
                                break;
                            case "GBMD":
                                salesOrg = "MD01";
                                break;
                            case "GBNI":
                                salesOrg = "NI01";
                                break;
                            case "GBPA":
                                salesOrg = "PA01";
                                break;
                            case "GBSV":
                                salesOrg = "SV01";
                                break;
                            case "ITC0":
                                salesOrg = "ITC1";
                                break;
                            case "WTC0":
                                salesOrg = "WTC1";
                                break;
                            case "BV01":
                                salesOrg = "BV01";
                                break;
                            case "LCFL":
                                salesOrg = "LCFL";
                                break;
                            case "LCVE":
                                salesOrg = "LCVE";
                                break;
                            case "SAC0":
                                salesOrg = "SA01";
                                account1 = bankAccount;
                                //cuenta sinpe o normal rechaza la solicitud
                                if (bankAccount.Length <= 17 && bankAccount.Length > 0)
                                {
                                    returnMsg = vendorName + ": " + "Por favor ingrese las cuentas IBAN de 20 caracteres";
                                    res2 = res2 + returnMsg + "<br>";
                                    returnRequest = true;
                                    continue;

                                }
                                //IBAN sin CR
                                else if (bankAccount.Length == 20)
                                {
                                    if (bankAccount.Substring(0, 2) != "CR")
                                        iban = "CR" + bankAccount;
                                    else
                                        iban = bankAccount;
                                    bankAccount = account1.Substring(0, 17);
                                }
                                //IBAN normal
                                else if (bankAccount.Length == 22)
                                {
                                    if (bankAccount.Substring(0, 2) != "CR")
                                        iban = "CR" + bankAccount.Substring(0, bankAccount.Length - 2);
                                    else
                                        iban = bankAccount;
                                    bankAccount = account1.Substring(0, 17);
                                }
                                //no agrega cuenta
                                else
                                {
                                    bankAccount = "";
                                    iban = "";
                                }
                                break;
                            case "GBCO":
                                salesOrg = "CO01";
                                break;
                        }


                        #endregion termina validaciones

                        if (returnMsg == "")
                        {
                            #region SAP

                            try
                            {
                                switch (Start.enviroment)
                                {
                                    case "DEV":
                                        docNumP = "100000002";
                                        break;

                                    case "QAS":
                                        docNumP = "100000949";
                                        break;

                                    case "PRD":
                                        docNumP = "100000949";
                                        break;
                                }

                                #region Parámetros de SAP
                                Dictionary<string, string> parameters = new Dictionary<string, string>
                                {
                                    ["TRATAMIENTO"] = vendorTitle,
                                    ["RAZON_SOC"] = vendorName,
                                    ["ID_NIF1"] = nif,
                                    ["ID_NIF2"] = nif2,
                                    ["DIRECCION"] = vendorAddress,
                                    ["COD_POSTAL"] = postalCode,
                                    ["CIUDAD"] = "", //vendorCity,
                                    ["PAIS"] = vendorCountry,
                                    ["REGION"] = vendorRegion,
                                    ["TELEFONO"] = vendorPhone,
                                    ["EMAIL"] = vendorEmail,
                                    ["RAMO"] = "MD01",
                                    ["GIRO"] = vendorActivity,
                                    ["MONEDA"] = currency,
                                    ["COD_BANCO"] = bankCode,
                                    ["CUENTA_BCO"] = bankAccount,
                                    ["COD_CONTROL"] = controlCode,
                                    ["REF_BANCO"] = "",
                                    ["IBAN"] = iban,
                                    ["CTA_HOLDER"] = "",
                                    ["COMPANY_COD"] = companyCode,
                                    ["MET_PAGO"] = paymentMethod,
                                    ["ORG_COMPRAS"] = salesOrg,
                                    ["ACCION"] = "ADICIONAR",
                                    ["TIPO_RET"] = "", //taxType,
                                    ["IND_RET"] = taxCode
                                };

                                if (function == "check")
                                    parameters["TEST"] = "X";

                                #endregion

                                #region Invocar FM
                                console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                                IRfcFunction zdmCrearProv = sap.ExecuteRFC(erpMand, "ZDM_CREAR_PROV", parameters);
                                #endregion

                                #region Procesar Salidas del FM
                                vendorId = zdmCrearProv.GetValue("BP").ToString();
                                fmRes = zdmCrearProv.GetValue("RESPUESTA").ToString();  //Mensaje OK cuando se crea correctamente, en otros casos es igual a blanco.
                                fmMsg = zdmCrearProv.GetValue("MENSAJE").ToString();
                                fmMsg1 = zdmCrearProv.GetValue("MENSAJE1").ToString();

                                if (bankAccount != "")
                                {
                                    console.WriteLine("Agregar cuenta bancaria");
                                    AddAccount(vendorId, bankCode, bankAccount, "ML", companyCode);

                                    if (addAccount == true && vendorId != "" && fmMsg1 != "proveedor existe")
                                        AddAccount(vendorId, bankCode, bankAccountUs, "US", companyCode); //, IBAN no se necesita pasar el parámetro para CR en el método se concatena
                                }

                                console.WriteLine(vendorId + "-> MENSAJE: " + fmMsg + "-> MENSAJE1: " + fmMsg1 + "-> RESPUESTA:" + fmRes);

                                if (fmRes != "OK" && fmRes == "" && fmMsg1 != "proveedor existe" && fmMsg1 != "cliente existe" && fmMsg1 != "Se amplio el cliente a proveedor" && function != "check")
                                {
                                    response = vendorName + ": " + fmMsg1 + " - " + fmMsg + "<br>";
                                    msg = "error";
                                    validateData = false;
                                }
                                else if (fmMsg1 != "" && fmMsg1 != "Se amplio el cliente a proveedor" && fmMsg1 != "proveedor existe" && fmMsg1 != "cliente existe")
                                {
                                    if (fmMsg1 == "class_error")
                                        response = vendorId + " - " + vendorName + ": " + fmMsg1 + " - " + fmMsg + " - " + "No se pudo ampliar a Company Code y Purchasing (pero ya se amplió a DMS)" + "<br>";
                                    else
                                    {
                                        if (function == "check")
                                            response = vendorName + ": " + fmMsg1 + " - " + fmMsg + " revisar error";
                                        else
                                            response = vendorId + " - " + vendorName + ": " + fmMsg1 + " - " + fmMsg + " - " + "Proveedor creado, revisar error (pero ya se amplió a DMS)" + "<br>";
                                    }

                                    msg = "OK";
                                    validateData = false;
                                }
                                else if (fmMsg1 == "proveedor existe")
                                {
                                    msg = "proveedor existe";
                                    response = vendorId + " - " + vendorName + ": " + "NIF ya está asociado a un BP por favor verificar (ya se creó la carpeta DMS)" + "<br>";
                                    returnRequest = true;
                                }
                                else if (fmMsg1 == "Se amplio el cliente a proveedor" || fmMsg1 == "cliente existe")
                                {
                                    response = vendorId + " - " + vendorName + ": " + "NIF ya esta asociado a un cliente, se amplió a proveedor " + "<br>";
                                    msg = "OK";
                                    revision = true;
                                }
                                else
                                {
                                    response = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito" + "<br>";
                                    msg = "OK";
                                    revision = true;
                                }

                                if (function == "create")
                                {
                                    if (false)  //eliminar IF cuando se arregle DMS
                                    {
                                        #region DMS
                                        console.WriteLine("Procesar DMS");

                                        if (msg == "proveedor existe" || msg == "OK")
                                        {
                                            if (vendorGroup != "G_INT")
                                            {
                                                //crea carpeta DMS
                                                //crea la carpeta y devuelve el id de la carpeta de documentos generales
                                                console.WriteLine("Crear Carpeta");
                                                generalFolderId = CreateFolder(vendorId, vendorName, vendorCountry, vendorCat, vendorDesc, vendorGroup, vendorType, vendorPhone, companyCode, docNumP);

                                                if (!string.IsNullOrEmpty(generalFolderId))
                                                {
                                                    //linea de código para dar autorizaciones
                                                    generalFolderId = generalFolderId.Substring(generalFolderId.Length - 15, 15);
                                                    sap.LogSAP(erpMand.ToString());

                                                    #region aprobación
                                                    console.WriteLine("Aprobación de carpetas");
                                                    try
                                                    {
                                                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/ncv02n";
                                                        SapVariants.frame.SendVKey(0);
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKNR")).Text = vendorId;
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKAR")).Text = "ZFV";
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKTL")).Text = "000";
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKVR")).Text = "00";
                                                        SapVariants.frame.SendVKey(0);
                                                        ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSACL")).Select();
                                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();
                                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                                                        ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/ncv02n";
                                                        SapVariants.frame.SendVKey(0);
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKNR")).Text = generalFolderId;
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKAR")).Text = "ZFG";
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKTL")).Text = "000";
                                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDRAW-DOKVR")).Text = "00";
                                                        SapVariants.frame.SendVKey(0);
                                                        ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSACL")).Select();
                                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();
                                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                                                    }
                                                    catch (Exception exs)
                                                    {
                                                        resAuth = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito, error al autorizar la carpeta" + "<br>" + exs + "<br>";
                                                        validateData = false;
                                                    }
                                                    #endregion

                                                    #region cargar archivos
                                                    //linea de código para subir los archivos a la carpeta de general
                                                    if (root.doc_aprob != null && root.doc_aprob[0] != null)
                                                    {
                                                        try
                                                        {
                                                            for (int w = 0; w < root.doc_aprob.Count; w++)
                                                            {
                                                                string adjunto = root.doc_aprob[w].ToString();
                                                                Newtonsoft.Json.Linq.JObject jsonFile = Newtonsoft.Json.Linq.JObject.Parse(adjunto);
                                                                string nameFile = jsonFile["name"].ToString();
                                                                string pathFile = jsonFile["path"].ToString();
                                                                string local_ruta = root.FilesDownloadPath + "\\" + nameFile;

                                                                #region descargar archivo del FTP   
                                                                MasterDataSqlSS md = new MasterDataSqlSS();

                                                                bool result = md.DownloadFile(pathFile);

                                                                if (result)
                                                                {
                                                                    clipData = clipData + nameFile + "\r\n";
                                                                }
                                                                //TransferOperationResult transferResult;
                                                                //TransferOptions transferOptions = new TransferOptions();

                                                                //SessionOptions sessionOptions = db2.ConnectFTP(1, "databot.gbm.net", 21, "gbmadmin", cred.password_server_web, false, "");

                                                                //sessionOptions.AddRawSettings("ProxyPort", "0");

                                                                //using (Session session = new Session())
                                                                //{
                                                                //    console.WriteLine(DateTime.Now + " > > > " + " Estableciendo conexión");
                                                                //    session.Open(sessionOptions);
                                                                //    console.WriteLine(DateTime.Now + " > > > " + " Descargando archivo");
                                                                //    transferOptions.TransferMode = TransferMode.Binary;
                                                                //    string ftp_ruta = "/dm_gestiones_mass/" + root.IdGestionDM + "/" + adjunto;

                                                                //    transferResult = session.GetFiles(ftp_ruta, local_ruta, false, transferOptions);
                                                                //    transferResult.Check();
                                                                //    session.Dispose();
                                                                //}
                                                                #endregion

                                                            }

                                                            console.WriteLine(DateTime.Now + " > > > " + "Cargando Archivos");
                                                            Clipboard.SetText(clipData.ToString(), TextDataFormat.Text);

                                                            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
                                                            SapVariants.frame.SendVKey(0);
                                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "ZDM_DMS_ADD_FILES";
                                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();

                                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtDOC_NUM-LOW")).Text = generalFolderId;
                                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtPATH-LOW")).Text = root.FilesDownloadPath + "\\";
                                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btn%_FILE_N_%_APP_%-VALU_PUSH")).Press();
                                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[24]")).Press();
                                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[8]")).Press();
                                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                                                            try
                                                            {
                                                                msgSap = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                                                            }
                                                            catch (Exception) { }

                                                            if (msgSap == "")
                                                            {
                                                                resDmsFiles = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito, error al cargar archivos" + "<br>";
                                                                validateData = false;
                                                                createWithError = true;
                                                            }
                                                            else if (msgSap.Contains("Error:"))
                                                            {
                                                                resDmsFiles = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito, error al cargar archivos" + "<br>";
                                                                validateData = false;
                                                                createWithError = true;
                                                            }
                                                            else
                                                            {
                                                                response = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito" + "<br>";
                                                            }
                                                        }
                                                        catch (Exception exe)
                                                        {
                                                            resDmsFiles = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito, error al cargar archivos" + "<br>" + exe + "<br>";
                                                            validateData = false;
                                                            createWithError = true;
                                                        }

                                                    }
                                                    #endregion

                                                    sap.KillSAP();
                                                    Clipboard.Clear();
                                                }
                                                else if (resDms.Contains("Error al agregar caracteristicas al documento")) //error al agregar categorias
                                                {
                                                    response = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito, error agregando características" + "<br>";
                                                    validateData = false;
                                                    createWithError = true;
                                                }
                                                else
                                                {
                                                    response = vendorId + " - " + vendorName + ": " + "proveedor creado con éxito, error al crear carpeta DMS" + "<br>";
                                                    validateData = false;
                                                    createWithError = true;
                                                }
                                            }
                                        }
                                        #endregion

                                        if (resDmsFiles != "" && resAuth != "")
                                            response = resAuth + " - " + resDmsFiles;
                                        else if (resDmsFiles != "")
                                            response = resDmsFiles;
                                        else if (resAuth != "")
                                            response = resAuth;
                                    }
                                    else
                                    {
                                        mail.SendHTMLMail("por favor crear DMS al proveedor: " + vendorId + " Solicitud: " + row["requestId"].Value<string>().Trim(), new string[] { "hlherrera@gbm.net" }, "crear DMS al proveedor: " + vendorId, cc);
                                    }

                                    //log de cambios base de datos
                                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear Proveedor", response, root.Subject);
                                    respFinal = respFinal + "\\n" + "Crear Proveedor: " + response;

                                    string username = root.BDUserCreatedBy.Substring(0, root.BDUserCreatedBy.IndexOf('@'));
                                    DM.AddVendor(vendorId, username.ToUpper(), root.IdGestionDM);
                                }
                                res2 = res2 + response + "<br>";
                                #endregion

                            }
                            catch (Exception mensaje_error)
                            {
                                console.WriteLine(DateTime.Now + " > > > " + " Finishing process " + mensaje_error.Message);
                                res2 = res2 + vendorName + ": " + mensaje_error.ToString() + "<br>";
                                resFailure = mensaje_error.ToString();
                                validateData = false;
                            }
                            #endregion
                        }
                    }
                }


                #region Enviar notificaciones

                if (validateData == false)
                {
                    console.WriteLine(DateTime.Now + " > > > " + "enviando error de solicitud");
                    DM.ChangeStateDM(root.IdGestionDM, res2 + "<br>" + resFailure, "4"); //error
                    //enviar email de repuesta de error a datos maestros
                    mail.SendHTMLMail("Error al intentar crear proveedor<br><br>Gestión: " + root.IdGestionDM + "<br>" + res2 + "<br>" + resFailure, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);

                    if (createWithError)
                    {
                        NotiGestor(companyCode, res2);

                        if ((vendorGroup == "G_FRE" || vendorGroup == "G_SUB" || vendorGroup == "G_INT") && companyCode == "BV01")
                        {
                            string frelanceBody = "<b>Notificación de gestión de Proveedores:</b><br>" +
                               "Estimado(a) se le notifica que la solicitud <b>#" + root.IdGestionDM + "</b> ha finalizado, con el siguiente resultado:<br><br>" +
                               "Proveedor creado con éxito<br><br>" +
                               vendorName + " Id: " + vendorId + "<br>" +
                               "<b>Dirección:</b> " + vendorAddress + " " + vendorName2 + "<br>" +
                                "<b>Teléfono:</b> " + vendorPhone + "<br>" +
                               "<b>Datos bancarios:</b> " + dbAdd;
                            string[] cc = { root.BDUserCreatedBy };
                            mail.SendHTMLMail(frelanceBody, new string[] { "VCHAVES@gbm.net" }, root.Subject, cc);
                        }
                    }
                }
                else if (returnRequest == true)
                {
                    console.WriteLine(DateTime.Now + " > > > " + "Devolviendo solicitud");
                    DM.ChangeStateDM(root.IdGestionDM, res2, "5"); //RECHAZADO
                    //mail.EnviarCorreo(respuesta2, root.Solicitante, root.Subject, resp_type: 2);
                    wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Proveedores:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + res2);
                    //mdl.sendNotification(root.IdGestionDM, new string[] { root.BDUserCreatedBy }, "Rechazado", "**Notificación de gestión de Proveedores:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha sido rechazada, con el siguiente resultado: <br><br> " + res2, root.formDm, root.typeOfManagementText, root.factorType, root.factorDM );

                }
                else
                {
                    if (function == "create")
                    {
                        //finalizar solicitud
                        DM.ChangeStateDM(root.IdGestionDM, res2, "3"); //FINALIZADO
                        //agregar el id de proveedor, solicitante y id de gestion a la tabla de reporteria de Business System.

                        #region enviar confirmación al gestor del pais
                        companyCode = (companyCode.Length > 4) ? companyCode.Substring(0, 4) : companyCode;
                        NotiGestor(companyCode, res2);
                        #endregion

                        #region enviar confirmación al encargado freelance
                        if ((vendorGroup == "G_FRE" || vendorGroup == "G_SUB" || vendorGroup == "G_INT") && companyCode == "BV01")
                        {
                            string frelanceBody = "<b>Notificación de gestión de Proveedores:</b><br>" +
                               "Estimado(a) se le notifica que la solicitud <b>#" + root.IdGestionDM + "</b> ha finalizado, con el siguiente resultado:<br><br>" +
                               "Proveedor creado con éxito<br><br>" +
                               vendorName + " Id: " + vendorId + "<br>" +
                               "<b>Dirección:</b> " + vendorAddress + " " + vendorName2 + "<br>" +
                                "<b>Teléfono:</b> " + vendorPhone + "<br>" +
                               "<b>Datos bancarios:</b> " + dbAdd;
                            string[] cc = { root.BDUserCreatedBy };
                            mail.SendHTMLMail(frelanceBody, new string[] { "VCHAVES@gbm.net" }, root.Subject, cc);
                        }
                        #endregion

                        mail.SendHTMLMail("<b>Notificación de gestión de Proveedores:</b> Estimado(a) se le notifica que la solicitud <b>#" + root.IdGestionDM + "</b> ha finalizado, con el siguiente resultado: <br><br>Proveedor creado con éxito", new string[] { root.BDUserCreatedBy }, root.Subject);
                        wt.SendNotification(root.BDUserCreatedBy, "Notificación Datos Maestros, solicitud: " + root.IdGestionDM, "**Notificación de gestión de Proveedores:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + res2);
                        //mdl.sendNotification(root.IdGestionDM, new string[] { root.BDUserCreatedBy }, "Finalizado", "**Notificación de gestión de Proveedores:** Estimado(a) se le notifica que la solicitud **#" + root.IdGestionDM + "** ha finalizado, con el siguiente resultado: <br><br> " + res2, root.formDm, root.typeOfManagementText, root.factorType, root.factorDM);
                    }
                    else if (function == "check" && revision == true)
                    {
                        DM.ChangeStateDM(root.IdGestionDM, "", "6"); //APROBACION CONTADORES
                        SendWebexNotification(root.BDUserCreatedBy, root.IdGestionDM);
                    }
                }
                #endregion

                root.requestDetails = respFinal;

            }
            catch (Exception ex)
            {
                DM.ChangeStateDM(root.IdGestionDM, ex.Message, "4"); //ERROR
                mail.SendHTMLMail("Gestión: " + root.IdGestionDM + "<br>" + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, "Error: " + root.Subject, cc);
            }
        }

        private string CreateFolder(string docNum, string description, string countyr, string vendorCat, string vendorDesc, string vendorGroup, string vendorType, string phone, string cCode, string docNumP)
        {
            string generalDocId = "";
            try
            {
                #region crea carpeta vendor
                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["DOC_NUM"] = docNum,
                    ["PART"] = "000",
                    ["VERSION"] = "00",
                    ["DOC_TYPE"] = "ZFV",
                    ["DESCRIPTION"] = description,
                    ["ORIGIN"] = countyr,
                    ["DOC_NUM_P"] = docNumP,
                    ["DOC_TYPE_P"] = "ZFG",
                    ["PART_P"] = "000",
                    ["VERSION_P"] = "00"
                };
                console.WriteLine("Ejecutando ZDM_CREATE_DMS");
                IRfcFunction func = sap.ExecuteRFC(erpMand, "ZDM_CREATE_DMS", parameters);
                #endregion

                string resp_dms_vendor = func.GetValue("RESPUESTA").ToString();

                if (resp_dms_vendor.Substring(0, 5) == "Error")
                {
                    resDms = resp_dms_vendor;
                }
                else
                {
                    #region crea carpeta doc generales

                    Dictionary<string, string> parameters2 = new Dictionary<string, string>
                    {
                        ["DOC_NUM"] = "",
                        ["PART"] = "",
                        ["VERSION"] = "",
                        ["DOC_TYPE"] = "ZFG",
                        ["DESCRIPTION"] = "Documentos Generales",
                        ["ORIGIN"] = "",
                        ["DOC_NUM_P"] = docNum,
                        ["DOC_TYPE_P"] = "ZFV",
                        ["PART_P"] = "000",
                        ["VERSION_P"] = "00"
                    };
                    console.WriteLine("Ejecutando ZDM_CREATE_DMS");
                    IRfcFunction func2 = sap.ExecuteRFC(erpMand, "ZDM_CREATE_DMS", parameters2);

                    string resp_dms_general = func2.GetValue("RESPUESTA").ToString();
                    if (resp_dms_general.Substring(0, 5) == "Error")
                        resDms = resp_dms_general;
                    else
                        generalDocId = func2.GetValue("DOCUMENT_NUMBER").ToString();

                    #endregion

                    #region agregar clases

                    Dictionary<string, string> parameters4 = new Dictionary<string, string>
                    {
                        ["DOC_NUM"] = docNum,
                        ["DOC_TYPE"] = "ZFV",
                        ["PART"] = "000",
                        ["VERSION"] = "00",
                        ["CLAS_TYPE"] = "017",
                        ["CLASE"] = "PRV",
                        ["VEN_TYPE"] = vendorType,
                        ["VEN_GROUP"] = vendorGroup,
                        ["VEN_CATEGORY"] = vendorCat,
                        ["VEN_DESC"] = vendorDesc,
                        ["VEN_TEL"] = phone,
                        ["VEN_COCODE"] = cCode
                    };
                    console.WriteLine("Ejecutando ZDM_ADD_CARACT_DMS");
                    IRfcFunction func3 = sap.ExecuteRFC(erpMand, "ZDM_ADD_CARACT_DMS", parameters4);

                    string resp_dms_caract = func3.GetValue("RESPUESTA").ToString();
                    if (resp_dms_caract.Substring(0, 5) == "Error")
                        resDms = resp_dms_caract;

                    #endregion
                }
            }
            catch (Exception ex)
            {
                resDms = ex.ToString();
            }
            return generalDocId;
        }
        private void NotiGestor(string coCodeRequest, string message)
        {
            root.requestDetails = "";
            List<string> gestores = new List<string>();
            try
            {
                DataTable gestoresDb = crud.Select($@"SELECT 
masterData.approvers.*, 
MIS.digital_sign.*,
masterData.factors.factor as factorName
FROM masterData.approvers
INNER JOIN MIS.digital_sign ON MIS.digital_sign.id = masterData.approvers.employee
INNER JOIN masterData.factors ON masterData.factors.id = masterData.approvers.factor
WHERE masterData.approvers.permission = 167 AND masterData.factors.factor = '{coCodeRequest}' ", "automation");

                if (gestoresDb.Rows.Count > 0)
                {
                    foreach (DataRow row in gestoresDb.Rows)
                    {
                        //JArray factores = JArray.Parse(row["FACTORES"].ToString());

                        //foreach (JToken factor in factores)
                        //{
                        //    string coCode = JObject.Parse(factor.ToString())["FACTOR"].ToString();
                        //    if (coCode == coCodeRequest)
                        //        gestores.Add(row["USUARIO"].ToString().ToLower() + "@gbm.net");
                        //}

                        gestores.Add(row["email"].ToString().ToLower());
                    }

                    if (gestores.Count > 0)
                    {
                        mail.SendHTMLMail("<b>Notificación de gestión de Proveedores:</b>" +
                                      " Estimado(a) se le notifica que la solicitud <b>#" + root.IdGestionDM +
                                      "</b> ha finalizado, con el siguiente resultado: <br><br> " + message +
                                      "<br><br> Solicitado por el usuario: " + root.BDUserCreatedBy + " <br> Sociedad del Proveedor: " + coCodeRequest,
                                      gestores.ToArray(), root.Subject);

                        mail.SendHTMLMail("<b>Notificación de gestión de Proveedores:</b>" +
                                      " Estimado(a) se le notifica que la solicitud <b>#" + root.IdGestionDM +
                                      "</b> ha finalizado, con el siguiente resultado: <br><br> " + message +
                                      "<br><br> Solicitado por el usuario: " + root.BDUserCreatedBy + " <br> Sociedad del Proveedor: " + coCodeRequest,
                                      new string[] { "hlherrera@gbm.net" }, root.Subject);
                    }
                    else
                        mail.SendHTMLMail("Se le informa que no se encontraron Gestores para el Company code: " + coCodeRequest + "de la gestión: " + root.IdGestionDM, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);

                }
            }
            catch (Exception ex)
            {
                mail.SendHTMLMail("Se le informa que no se envió el correo a Gestores de la gestión: " + root.IdGestionDM + "error: " + ex.Message, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject);
            }

        }
        private void SendWebexNotification(string sender, string requestId)
        {
            //sender = sender.Replace("@gbm.net", "").ToUpper();
            try
            {
                DataTable factorTable = crud.Select($"SELECT factor FROM masterDataRequests WHERE id = {requestId}", "masterData");

                //extraer requestApprovers para estado de aprobacion contadores
                string requestApprovers2 = $@"(SELECT MIS.digital_sign.user
    FROM approvers appr
    INNER JOIN MIS.digital_sign ON appr.employee = MIS.digital_sign.id
    WHERE
    appr.factor = {factorTable.Rows[0]["factor"]}
                AND appr.permission in 
    (
        SELECT fkPermission
        FROM approversPermissions appPerm
        WHERE appPerm.fkMotherTable = 9
                AND appPerm.fkStatus = 6
    )
    AND appr.active = 1)";
                DataTable aprobadores = crud.Select(requestApprovers2, "masterData");
                string approvers = "";
                foreach (DataRow item in aprobadores.Rows)
                {
                    approvers += item["user"] + ", ";
                    wt.SendNotification(item["user"] + "@gbm.net", "Pendiente Aprobación Proveedores", $"Estimado(a) se le informa que la gestión {requestId} se encuentra en espera de su aprobación.");
                }
                approvers = approvers.Substring(0, approvers.Length - 2);
                crud.Update($"UPDATE masterDataRequests SET requestApprovers = '{approvers}' WHERE id = {requestId}", "masterData");
                wt.SendNotification(sender, "Pendiente Aprobación Proveedores", $"Estimado(a) se le informa que la gestión {requestId} se encuentra en espera de aprobación por parte del Contador país: {approvers}");
                //crud.Delete("Databot", "DELETE FROM `programar_notificaciones` WHERE ESTADO = 'EN REVISION'", "automation");
                //crud.Insert("Databot", "INSERT INTO `programar_notificaciones`(`ID_GESTION`, `SOLICITANTE`, `TIPO`, `ESTADO`, `ACTIVO`) VALUES (" + requestId + ",'" + sender + "','PROVEEDORES','APROBACION CONTADORES','X')", "automation");
                //mdl.sendNotification(root.IdGestionDM, new string[] { root.BDUserCreatedBy }, "Aprobacion de Contadores", $"Estimado(a) se le informa que la gestión {requestId} se encuentra en espera de aprobación por parte del Contador país: {approvers}", root.formDm, root.typeOfManagementText, root.factorType, root.factorDM);
            }
            catch (Exception) { }
        }
        private void AddAccount(string bp, string bankCode, string account, string currency, string companyCode, string iban = "")
        {
            string country;

            #region Bancos_países
            if (bankCode == "OTHER")
            {
                switch (companyCode)
                {
                    case "GBDR":
                        country = "DO";
                        break;
                    case "GBGT":
                        country = "GT";
                        break;
                    case "GBHN":
                        country = "HN";
                        break;
                    case "GBNI":
                        country = "NI";
                        break;
                    case "GBPA":
                        country = "PA";
                        break;
                    case "GBSV":
                        country = "SV";
                        break;
                    case "GBMD":
                        country = "US";
                        break;
                    case "GBCR":
                        country = "CR";
                        break;
                    case "GBCO":
                        country = "CO";
                        break;
                    default:
                        country = "VG";
                        break;
                }
            }
            else if (bankCode == "PBBAC")
            {
                if (companyCode == "GBCR")
                    country = "CR";
                else
                    country = "VG";
            }
            else
            {
                DataTable dtCountry = crud.Select("SELECT `country` FROM `banksCountry` WHERE bankCode = '" + bankCode + "'", "banks_db");
                country = dtCountry.Rows[0].ItemArray[0].ToString();
            }
            #endregion

            if (country == "CR")
                iban = "CR" + account;

            #region Agregar cuenta adicional

            #region Parametros de SAP
            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                ["VENDOR"] = bp,
                ["PAIS"] = country,
                ["COD_BANCO"] = bankCode,
                ["CUENTA_BCO"] = account,
                ["MONEDA"] = currency,
                ["IBAN"] = iban,
                ["CTA_HOLDER"] = "",
                ["ACCION"] = "AGREGAR"
            };
            #endregion

            #region Invocar FM
            /*IRfcFunction zicsChangeCtaBanco =*/
            sap.ExecuteRFC(erpMand, "ZICS_CHANGE_CTA_BANCO", parameters);
            #endregion

            // res = zicsChangeCtaBanco.GetValue("MENSAJE").ToString();
            #endregion
        }
    }
}
