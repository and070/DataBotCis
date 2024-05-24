using DataBotV5.Logical.Processes;
using System.Collections.Generic;
using DataBotV5.App.Global;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using System.Globalization;
using System.Data;
using DataBotV5.Logical.MicrosoftTools;
using System;

namespace DataBotV5.Logical.Projects.CriticalTransactions
{
    internal class CriticalTransactionsLogical
    {
        readonly ProcessInteraction proc = new ProcessInteraction();
        readonly Rooting root = new Rooting();
        readonly MsExcel excel = new MsExcel();
        //ScreenCapture screen = new ScreenCapture();
   

        /// <summary>
        /// 
        /// </summary>
        /// <param name="period">"1: mes actual, 2: mes anterior, 3: mes trasanterior, 4: mes tras trasanterior, 5: hoy</param>
        public void IniST03(string period = "2")
        {
            SapVariants.frame.Maximize();
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nst03";
            SapVariants.frame.SendVKey(0);
            if (SapVariants.session.ActiveWindow.Name == "wnd[1]")
                ((SAPFEWSELib.GuiModalWindow)SapVariants.session.FindById("wnd[1]")).Close();

            if (period == "5")
            {
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).ExpandNode("B.001"); //ecc
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).TopNode = "B";
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).ExpandNode("B.001.1");//day
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).SelectedNode = "B.001.1.001";
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).TopNode = "B";
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).DoubleClickNode("B.001.1.001");
            }
            else
            {
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).ExpandNode("B.999"); //total
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).TopNode = "B";
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).ExpandNode("B.999.3");//month
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).SelectedNode = "B.999.3.00" + period;
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).TopNode = "B";
                ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell")).DoubleClickNode("B.999.3.00" + period);
            }

            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[2]/shell")).ExpandNode("C");
            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[2]/shell")).SelectedNode = "C.1";//transaction profile
            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[2]/shell")).TopNode = "ARoot";
            ((SAPFEWSELib.GuiTree)SapVariants.session.FindById("wnd[0]/shellcont/shell/shellcont[2]/shell")).DoubleClickNode("C.1");//standar
        }
        public TransactionData GetTransaction(string transaction, bool image = true, bool table = true)
        {
            TransactionData res = new TransactionData();
            byte[] resImage = null;
            DataTable resDt = new DataTable();

            string status2 = ((SAPFEWSELib.GuiTitlebar)SapVariants.session.FindById("wnd[0]/titl")).Text.ToString();
            if (status2.Contains("User to Transaction"))
            {
                SapVariants.frame.Maximize();
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                ((SAPFEWSELib.GuiDialogShell)SapVariants.session.FindById("wnd[0]/shellcont[1]")).Close();
            }

            try
            {
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_0:SAPWL_ST03N:1100/ssubWL_SUBSCREEN_1:SAPWL_ST03N:1110/tabsG_TABSTRIP/tabpTA00/ssubWL_SUBSCREEN_2:SAPWL_ST03N:1130/cntlALVCONTAINER/shellcont/shell")).CurrentCellRow = -1;
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_0:SAPWL_ST03N:1100/ssubWL_SUBSCREEN_1:SAPWL_ST03N:1110/tabsG_TABSTRIP/tabpTA00/ssubWL_SUBSCREEN_2:SAPWL_ST03N:1130/cntlALVCONTAINER/shellcont/shell")).SelectColumn("TCODE");
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_0:SAPWL_ST03N:1100/ssubWL_SUBSCREEN_1:SAPWL_ST03N:1110/tabsG_TABSTRIP/tabpTA00/ssubWL_SUBSCREEN_2:SAPWL_ST03N:1130/cntlALVCONTAINER/shellcont/shell")).PressToolbarButton("&MB_FILTER");
                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW")).Text = transaction;
                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_0:SAPWL_ST03N:1100/ssubWL_SUBSCREEN_1:SAPWL_ST03N:1110/tabsG_TABSTRIP/tabpTA00/ssubWL_SUBSCREEN_2:SAPWL_ST03N:1130/cntlALVCONTAINER/shellcont/shell")).SelectedRows = "0";
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/usr/ssubSUBSCREEN_0:SAPWL_ST03N:1100/ssubWL_SUBSCREEN_1:SAPWL_ST03N:1110/tabsG_TABSTRIP/tabpTA00/ssubWL_SUBSCREEN_2:SAPWL_ST03N:1130/cntlALVCONTAINER/shellcont/shell")).DoubleClickCurrentCell();
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/shellcont[1]/shell")).PressToolbarContextButton("&MB_VIEW");
                ((SAPFEWSELib.GuiGridView)SapVariants.session.FindById("wnd[0]/shellcont[1]/shell")).SelectContextMenuItem("&PRINT_BACK_PREVIEW");
                if (image)
                    resImage = ((SAPFEWSELib.GuiMainWindow)SapVariants.session.FindById("wnd[0]")).HardCopyToMemory(1);//1 es JPEG
                if (table)
                {
                    //guardar la tabla
                    try
                    {
                        string fileName = "tc.XLSX";
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[43]")).Press();//boton export
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();//boton ok del formato
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = root.FilesDownloadPath;//path del file
                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = fileName;//nombre del file
                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();//boton replace
                        System.Threading.Thread.Sleep(3000);
                        proc.KillProcess("EXCEL", true);
                        resDt = excel.GetExcel(root.FilesDownloadPath + "\\" + fileName);
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }//No hay registros

            res.ImageResult = resImage;
            res.DtResult = resDt;

            return res;
        }
        public string ByteToHtmlTag(byte[] imgBytes)
        {
            string imgString = Convert.ToBase64String(imgBytes);
            string html = string.Format("<img src=\"data:image/Jpeg;base64,{0}\">", imgString);
            return html;
        }
        public void CloseGui()
        {
            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nex";
            SapVariants.frame.SendVKey(0);
        }
    }
    public class TransactionData
    {
        public byte[] ImageResult { get; set; }
        public DataTable DtResult { get; set; }
    }
}
