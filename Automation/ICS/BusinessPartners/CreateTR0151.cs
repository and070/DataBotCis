using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
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
    /// Clase ICS Automation encargada de agregar counter party.
    /// </summary>
    class CreateTR0151
    {
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        MsExcel excel = new MsExcel();
        Rooting root = new Rooting();
        Stats stats = new Stats();
        Log log = new Log();
        string respFinal = "";


        string mandErp = "ERP";

        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
            bool checkLogin = sap.CheckLogin(mandErp);
            if (!checkLogin)
            {
                //Leer correo y descargar archivo
                console.WriteLine("Descargando archivo");
                if (mail.GetAttachmentEmail("Solicitudes TR0151", "Procesados", "Procesados TR0151"))
                {
                    console.WriteLine("Procesando...");
                    sap.BlockUser(mandErp, 1);

                    DataTable excelDt = excel.GetExcel(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    ProcessTR0151(excelDt);

                    sap.BlockUser(mandErp, 0);
                    console.WriteLine("Creando Estadísticas");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }
        }
        public void ProcessTR0151(DataTable excelDt)
        {
            string customer, localAccount, dolarAccount, bankKey, currency;
            bool validateLines = true;

            excelDt.Columns.Add("Respuesta");

            string validation = excelDt.Columns[1].ColumnName;

            if (validation.Substring(0, 1) != "x")
            {
                console.WriteLine("Devolviendo solicitud");
                mail.SendHTMLMail("Utilizar la plantilla oficial de cambios", new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }
            else
            {
                proc.KillProcess("saplogon", false);

                foreach (DataRow item in excelDt.Rows)
                {
                    string result = "";

                    customer = item[0].ToString().Trim();
                    string coCode = item[1].ToString().Trim().ToUpper();

                    switch (coCode)
                    {
                        case "GBCR":
                            currency = "CRC"; localAccount = "07751"; dolarAccount = "69872"; bankKey = "CRBAC";
                            break;
                        case "GBDR":
                            currency = "DOP"; localAccount = "93925"; dolarAccount = "06077"; bankKey = "DPOPU";
                            break;
                        case "GBGT":
                            currency = "GTQ"; localAccount = "05358"; dolarAccount = "71909"; bankKey = "GTBAC";
                            break;
                        case "GBHN":
                            currency = "HNL"; localAccount = "00920"; dolarAccount = "00006"; bankKey = "HNBAC";
                            break;
                        case "GBHQ":
                            currency = "CRC"; localAccount = "07751"; dolarAccount = "69872"; bankKey = "CRBAC";
                            break;
                        case "GBMD":
                            currency = "USD"; localAccount = "08331"; dolarAccount = "08331"; bankKey = "USBAC";
                            break;
                        case "GBNI":
                            currency = "NIO"; localAccount = "05150"; dolarAccount = "03804"; bankKey = "NIBAC";
                            break;
                        case "GBPA":
                            currency = "PAB"; localAccount = "00045"; dolarAccount = "00045"; bankKey = "PBBAC";
                            break;
                        case "GBSV":
                            currency = "USD"; localAccount = "05502"; dolarAccount = ""; bankKey = "SVBAC";
                            break;
                        case "BV01":
                            currency = "USD"; localAccount = "60088"; dolarAccount = ""; bankKey = "USCIT";
                            break;
                        case "LCVE":
                            currency = "VEF"; localAccount = "26747"; dolarAccount = ""; bankKey = "VBANE";
                            break;
                        case "LCFL":
                            currency = "USD"; localAccount = "56895"; dolarAccount = ""; bankKey = "USCIT";
                            break;
                        default:
                            currency = ""; localAccount = ""; dolarAccount = ""; bankKey = "";
                            break;
                    }

                    if (customer != "")
                    {
                        console.WriteLine("Corriendo SAP GUI: " + root.BDProcess);
                        try
                        {
                            sap.LogSAP(mandErp.ToString());

                            #region SAP script

                            SapVariants.frame.Maximize();
                            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/NBP"; //Cod del cliente
                            SapVariants.frame.SendVKey(0); //enter
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press(); //Carpeta open
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtBUS_JOEL_MAIN-OPEN_NUMBER")).Text = customer; //Código del cliente
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();   //Enter

                            if (ExistBP() == true)
                                result = "El cliente no existe";
                            else
                            {
                                if (SapVariants.session.ActiveWindow.Name == "wnd[1]")
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();  //Do you want to apply this change?

                                if (Exist() == false)
                                {
                                    //se amplia 
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[6]")).Press();

                                    string errorStatus = IsErrorFree();

                                    if (errorStatus != "OK")
                                    {
                                        validateLines = false;
                                        result = errorStatus;
                                    }
                                    else
                                    {

                                        ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/subSCREEN_1100_ROLE_AND_TIME_AREA:SAPLBUPA_DIALOG_JOEL:1110/cmbBUS_JOEL_MAIN-PARTNER_ROLE")).Key = "TR0151"; //Seleccione el roll
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press(); //save

                                        if (SapVariants.session.ActiveWindow.Name == "wnd[1]")
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();  //Do you want to apply this change?

                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[26]")).Press(); //company code
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[6]")).Press(); //change 'esta fila puede que no se necesite


                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/subSCREEN_1100_SUB_HEADER_AREA:SAPLFS_BP_ECC_DIALOGUE:0001/btnPUSH_FSBP_CC_SWITCH")).Press();
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/subSCREEN_1100_SUB_HEADER_AREA:SAPLFS_BP_ECC_DIALOGUE:0001/ctxtBS001-BUKRS")).Text = coCode;
                                        ((SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[0]")).SendVKey(0); //enter

                                        //Inicia la segunda parte con moneda dolar

                                        string status2 = ((SAPFEWSELib.GuiTitlebar)SapVariants.session.FindById("wnd[0]/titl")).Text.ToString().Substring(0, 7);
                                        if (status2 == "Display")
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[6]")).Press();

                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-WAERS[0,0]")).Text = currency; //ej HNL
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-ZAHLVID[1,0]")).Text = (currency + "1"); //ej HNL1
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-HBKID[3,0]")).Text = bankKey; //ej  HNBAC
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-HKTID[4,0]")).Text = localAccount; //EJ 00920
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-RPZAHL[6,0]")).Text = customer; //BP Columna A
                                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/chkVTB_STC1-SZART[5,0]")).Selected = true;

                                        if (coCode != "GBSV" && coCode != "BV01" && coCode != "GBMD" && coCode != "LCFL")
                                        {
                                            ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/chkVTB_STC1-SZART[5,1]")).Selected = true;
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-WAERS[0,1]")).Text = "USD";
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-ZAHLVID[1,1]")).Text = "USD1";
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-HBKID[3,1]")).Text = bankKey; //ej CRBAC
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-HKTID[4,1]")).Text = dolarAccount;
                                            ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/chkVTB_STC1-SZART[5,0]")).Selected = true;
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1/ctxtVTB_STC1-RPZAHL[6,1]")).Text = customer;
                                        }

                                        //AQUI SE VA AL TAB INDICADO Y SELECCIONA EL CAMPO IPP LOAN GIVEN
                                        string tab = "";

                                        for (int j = 1; j < 5; j++)
                                        {
                                            tab = ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + j)).Text;
                                            if (tab == "SI: Authorizations")
                                            {
                                                tab = j.ToString();
                                                break;
                                            }
                                        }

                                        ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab)).Select();

                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab + "/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:0030/cntlTREE3/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("         99");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab + "/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:0030/cntlTREE3/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("        120");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab + "/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:0030/cntlTREE3/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("        142");

                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab + "/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:0030/cntlTREE3/shellcont/shell/shellcont[1]/shell[1]")).SelectItem("        143", "C          2");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab + "/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:0030/cntlTREE3/shellcont/shell/shellcont[1]/shell[1]")).EnsureVisibleHorizontalItem("        143", "C          2");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_0" + tab + "/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:0030/cntlTREE3/shellcont/shell/shellcont[1]/shell[1]")).ChangeCheckbox("        143", "C          2", true);

                                        ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01")).Select(); //va a la pestana sy payment
                                        ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1")).GetAbsoluteRow(0).Selected = true; //selecciona la fila en naranja
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/btnPUSH_FTB_ALL")).Press();  //assing

                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("         99");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("        120");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("        142");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).SelectItem("        143", "&Hierarchy");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).EnsureVisibleHorizontalItem("        143", "&Hierarchy");
                                        ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]")).PressButton("MARK1I");

                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).SelectItem("        143", "&Hierarchy");
                                        ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).EnsureVisibleHorizontalItem("        143", "&Hierarchy");
                                        ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]")).PressButton("MARK1O");
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[5]")).Press();

                                        if (coCode != "GBSV" && coCode != "BV01" && coCode != "GBMD" && coCode != "LCFL")
                                        {
                                            ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/tblSAPLFTBPC_VTB_STC1")).GetAbsoluteRow(1).Selected = true;
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/subA02P01:SAPLFTBP:1050/btnPUSH_FTB_ALL")).Press();  //assing

                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("         99");
                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("        120");
                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).ExpandNode("        142");
                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).SelectItem("        143", "&Hierarchy");
                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).EnsureVisibleHorizontalItem("        143", "&Hierarchy");
                                            ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]")).PressButton("MARK1I");
                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).SelectItem("        143", "&Hierarchy");
                                            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]")).EnsureVisibleHorizontalItem("        143", "&Hierarchy");
                                            ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7002/subA03P01:SAPLFTBP:0010/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]")).PressButton("MARK1O");
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[5]")).Press();
                                        }

                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press(); //save


                                        // manejar el error de CVI_API


                                        if (SapVariants.session.ActiveWindow.Name == "wnd[1]")
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();  //Do you want to apply this change?

                                        result = "El Counter party ha sido actualizado";

                                        //Log de base de datos
                                        console.WriteLine(result);
                                        log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Cambio Counter party", result, root.Subject);
                                        respFinal = respFinal + "\\n" + "Cambio Counter party: " + result;

                                    }

                                }
                                else
                                {
                                    //Log de base de datos
                                    result = "Cliente ya ampliado";
                                    console.WriteLine(customer + "Cliente ya ampliado");
                                    log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Cambio Counter party", result, root.Subject);
                                    respFinal = respFinal + "\\n" + "Cambio Counter party: " + result;

                                }
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            validateLines = false;
                            console.WriteLine(" Error " + val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, 0));
                            result = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString() + ": " + ex.ToString();
                        }
                    }

                    item["Respuesta"] = result;
                }

                console.WriteLine("Respondiendo solicitud");

                //Enviar email de repuesta
                string sender = root.BDUserCreatedBy;
                if (validateLines == false)
                    sender = "internalcustomersrvs@gbm.net";

                string msg = val.ConvertDataTableToHTML(excelDt);

                mail.SendHTMLMail(msg, new string[] { sender }, root.Subject);
                sap.KillSAP();

                root.requestDetails = respFinal;

            }
        }

        public bool Exist()
        {
            bool exi;
            try
            {
                ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/subSCREEN_1100_ROLE_AND_TIME_AREA:SAPLBUPA_DIALOG_JOEL:1110/cmbBUS_JOEL_MAIN-PARTNER_ROLE")).Key = "TR0151"; //Rol de Counterparty
                exi = true;
            }
            catch (Exception) { exi = false; }

            return exi;
        }
        public bool ExistBP()
        {
            //verifica si existe BP
            bool exi = false;
            string status = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
            if (status != "")
            {
                if (status.Substring(status.Length - 14) == "does not exist")
                    exi = true;
                else
                    exi = false;
            }
            return exi;
        }
        public string IsErrorFree()
        {
            //verifica si existe BP
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
            string status = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
            if (status == "Data of business partner is error-free")
                status = "OK";

            return status;
        }
    }
}

