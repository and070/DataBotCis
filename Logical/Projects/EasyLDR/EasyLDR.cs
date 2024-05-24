using System;
using System.Collections.Generic;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;

namespace DataBotV5.Logical.Projects.EasyLDR
{
    /// <summary>
    /// Clase Logical encargada de EasyLDR.
    /// </summary>
    class EasyLDR

    {

        ConsoleFormat console = new ConsoleFormat();
        public bool ConnectSAP(Dictionary<string, string> filesToUpload, String opportunity)
        {
            try
            {
                //SAP_Variants.frame.Iconify(); //minimizar el frame, ventana

                //Abrir la oportunidad mediante CRMD_ORDER 
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nCRMD_ORDER";
                SapVariants.frame.SendVKey(0);
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtGV_OBJECT_ID")).Text = opportunity;
                SapVariants.frame.SendVKey(0);

                //Click a la primer opp
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0130/subSUBSCREEN_1O_NAVIG:SAPLCRM_1O_LOCATOR:0110/ssubCRM_BUS_LOCATOR:SAPLBUS_LOCATOR:3101/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell")).SelectedRows = "0";//.SetCurrentCell(0, "OBJECT_ID");//.SelectItem("01", "OBJECT_ID");
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0130/subSUBSCREEN_1O_NAVIG:SAPLCRM_1O_LOCATOR:0110/ssubCRM_BUS_LOCATOR:SAPLBUS_LOCATOR:3101/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell")).DoubleClickCurrentCell();

                //Localizar el tab de attachments 
                ((SAPFEWSELib.GuiTab)SapVariants.session.FindById(@"wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\TOPP_HD10")).Select();

                //Boton de editar en SAP
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_COMMON:SAPLCRM_1O_UI:3150/subSCR_1O_TT:SAPLCRM_1O_UI:2600/btnGV_TOGGTRANS")).Press();
                foreach (KeyValuePair<string, string> file in filesToUpload)
                {

                    //Botón de importar documento 
                    ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById(@"wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\TOPP_HD10/ssubTABSTRIP_SUBSCREEN:SAPLCRM_1O_GEN_UI:6000/cntlCM_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell")).PressButton("KWUIFUNC_DOC_IMPORT");

                    //Ingresar la información de las rutas y nombres del file
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtDY_PATH")).Text = file.Value;
                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtDY_FILENAME")).Text = file.Key;
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();

                }

                 //Guardar la opp en SAP 
                 ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                return true;
            }
            catch (Exception e)
            {
                console.WriteLine(e.ToString());
                return false;
            }
        }

        public bool ConnectSAP2(List<String> fileLDRRoute, String opportunity)
        {
            try
            {
                //SAP_Variants.frame.Iconify(); //minimizar el frame, ventana

                //Abrir la oportunidad mediante CRMD_ORDER 
                ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nCRMD_ORDER";
                SapVariants.frame.SendVKey(0);
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press();
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtGV_OBJECT_ID")).Text = opportunity;
                SapVariants.frame.SendVKey(0);

                //Click a la primer opp
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0130/subSUBSCREEN_1O_NAVIG:SAPLCRM_1O_LOCATOR:0110/ssubCRM_BUS_LOCATOR:SAPLBUS_LOCATOR:3101/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell")).SelectedRows = "0";//.SetCurrentCell(0, "OBJECT_ID");//.SelectItem("01", "OBJECT_ID");
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0130/subSUBSCREEN_1O_NAVIG:SAPLCRM_1O_LOCATOR:0110/ssubCRM_BUS_LOCATOR:SAPLBUS_LOCATOR:3101/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3213/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell")).DoubleClickCurrentCell();

                //Localizar el tab de attachments 
                ((SAPFEWSELib.GuiTab)SapVariants.session.FindById(@"wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\TOPP_HD10")).Select();

                //Boton de editar en SAP
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_COMMON:SAPLCRM_1O_UI:3150/subSCR_1O_TT:SAPLCRM_1O_UI:2600/btnGV_TOGGTRANS")).Press();

                //Botón de importar documento 
                ((SAPFEWSELib.GuiToolbarControl)SapVariants.session.FindById(@"wnd[0]/usr/ssubSUBSCREEN_1O_MAIN:SAPLCRM_1O_MANAG_UI:0120/subSUBSCREEN_1O_WORKA:SAPLCRM_1O_WORKA_UI:2100/subSCR_1O_MAINTAIN:SAPLCRM_1O_UI:1100/subSCR_1O_MAINTAIN:SAPLCRM_OPPORT_UI:0101/subSCRAREA0:SAPLCRM_OPPORT_UI:3041/subSCRAREA1:SAPLCRM_OPPORT_UI:3100/tabsTABSTRIP_HEADER/tabpT\TOPP_HD10/ssubTABSTRIP_SUBSCREEN:SAPLCRM_1O_GEN_UI:6000/cntlCM_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell")).PressButton("KWUIFUNC_DOC_IMPORT");

                //Ingresar la información de las rutas y nombres del file
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtDY_PATH")).Text = fileLDRRoute[0];
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/txtDY_FILENAME")).Text = fileLDRRoute[1];
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();

                //Guardar la opp en SAP 
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                return true;
            }
            catch (Exception e)
            {
                console.WriteLine(e.ToString());
                return false;
            }
        }


        public void MrsJob()
        {
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nse38";
            SapVariants.frame.SendVKey(0);

            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtRS38M-PROGRAMM")).Text = "/MRSS/HCM_RPTWFMIF";
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press();
            ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")).CurrentCellRow = 5;
            ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")).SelectedRows = "5";
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[2]")).Press(); //Es una solucion alternativa que sustituye el DoubleClickItem en VB
            ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/mbar/menu[0]/menu[2]")).Select();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[13]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnSOFORT_PUSH")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();

        }


        public void LHOJA(String hoja, String fecha)
        {
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nml81n";
            SapVariants.frame.SendVKey(0);
            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]")).TopNode = "          4";
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[17]")).Press();
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtRM11R-LBLNI")).Text = hoja;
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[5]")).Press();
            ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/tabsTAB_HEADER/tabpREGA")).Select();
            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTAB_HEADER/tabpREGA/ssubSUB_ACCEPTANCE:SAPLMLSR:0420/ctxtESSR-BUDAT")).Text = fecha;
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[25]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnSPOP-OPTION1")).Press();
        }


        public void HOJA(String po, String item)
        {
            bool Listo = false;

            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/ncatm";

            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtCATSEKKO-SEBELN")).Text = po;
            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();

            for (int i = 4; i < 200; i++)
            {
                try
                {


                    String sapstrpo = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/lbl[6,' & i & ']")).Text;  //VERIFICAR si las comillas simples funcionan
                    String sapstritem = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/lbl[18,' & i & ']")).Text;  //VERIFICAR si las comillas simples funcionan

                    if ((sapstrpo == po) && (sapstritem == item))
                    {
                        ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[1]/usr/chk[2,' & i & ']")).Selected = true;
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[7]")).Press();
                        SapVariants.frame.SendVKey(0);
                        Listo = true;
                        break;

                    }

                }
                catch (Exception)
                {
                    console.WriteLine("Falló el método HOJA en la clase EasyLDR.");
                };
            }

            console.WriteLine("Hoja generada.");
        }

    }
}
