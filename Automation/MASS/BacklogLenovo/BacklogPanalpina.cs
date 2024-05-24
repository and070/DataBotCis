using System;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Projects.BusinessSystem;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.MASS.BacklogLenovo
{
    /// <summary>
    /// Clase MASS Automation encargada de la actualización backlog panalpina lenovo de guías aéreas. 
    /// </summary>
    class BacklogPanalpina 
    {


        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        Rooting root = new Rooting();
        SapVariants sap = new SapVariants();
        ProcessInteraction proc = new ProcessInteraction();
        MailInteraction mail = new MailInteraction();
        BsSQL bsql = new BsSQL();
        Log log = new Log();
        Stats estadisticas = new Stats();
        object[] columnas_duplicate;
        object[] columnas_duplicate2;
        string respFinal = "";
        string mandante = "ERP";




        public void Main() 
        {
            //revisa si el usuario RPAUSER esta abierto
            if (!sap.CheckLogin(mandante))
            {
                //leer correo y descargar archivo
                if (mail.GetAttachmentEmail("Solicitudes BL Aereas", "Procesados", "Procesados BL Aereas"))
                {
                    console.WriteLine("Procesando...");
                    sap.BlockUser(mandante, 1);
                    ProcessBacklogPanalpina(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    sap.BlockUser(mandante, 0);

                    root.requestDetails = respFinal;
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }
            }

        }

        public void ProcessBacklogPanalpina(string route)
        {

            #region variables privadas
            long contador;
            string sheetname = "";
            string trad_inv = "";
            string ship_date = "";
            string AWB = "";
            string carrier = "";
            string quant = "";
            string PO = "";
            string item_num = "";
            string item_num_sap = "";
            string nivel_sla = "";
            string po_quant = "";
            int dia = 0;
            string sdia = "";
            int mes = 0;
            string smes = "";
            int ano = 0;
            string sMessage = "";
            int scroll = 0;
            string order_status = "";
            string mensaje_sap = "";
            string sap_quant = "";
            string respuesta = "";
            string validacion = "";
            int filas = 0;
            int sap_quant_i = 0;
            int po_quant_i = 0;
            bool devolver = false;
            #endregion
            console.WriteLine("Abrir Excel y modificando");
            #region abrir excel
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange1 = null;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            sheetname = xlWorkSheet.Name;
            #endregion

            respuesta = "";
            mensaje_sap = "";

            #region validacion
            validacion = xlWorkSheet.Cells[1, 8].text;
            if (validacion != "")
            { validacion = validacion.ToString().Trim(); }
            if (validacion != "Service Level")
            {
                respuesta = "Hola, no se detecto que el archivo adjunto sea la plantilla correcta, por favor revisar el excel enviado";
                devolver = true;

            }
            else
            {
                filas = xlWorkSheet.UsedRange.Rows.Count;
                #region eliminar duplicados
                //for para agregar todas las columnas respectivas al array para eliminar duplicados

                for (int i = 0; i <= 7; i++)
                {
                    //ReDim Preserve columnas_duplicate(x);
                    Array.Resize(ref columnas_duplicate2, i + 1);
                    switch (i)
                    {
                        case 0:
                            columnas_duplicate2[i] = 1;
                            break;
                        case 1:
                            columnas_duplicate2[i] = 2;
                            break;
                        case 2:
                            columnas_duplicate2[i] = 3;
                            break;
                        case 3:
                            columnas_duplicate2[i] = 4;
                            break;
                        case 4:
                            columnas_duplicate2[i] = 5;
                            break;
                        case 5:
                            columnas_duplicate2[i] = 6;
                            break;
                        case 6:
                            columnas_duplicate2[i] = 7;
                            break;
                        case 7:
                            columnas_duplicate2[i] = 8;
                            break;
                    }
                }
                Array.Resize(ref columnas_duplicate2, 8);
                //eliminar duplicados
                xlWorkSheet.Range["A1:H" + filas].RemoveDuplicates(columnas_duplicate2, Excel.XlYesNoGuess.xlYes);
                #endregion eliminar duplicados

                filas = xlWorkSheet.UsedRange.Rows.Count;
                sap.LogSAP(mandante);
                console.WriteLine("Cargando a SAP");
                for (int x = 2; x <= filas; x++)
                {
                    mensaje_sap = "";
                    scroll = 12;
                    trad_inv = xlWorkSheet.Cells[x, 1].text.ToString().Trim();
                    ship_date = xlWorkSheet.Cells[x, 2].text.ToString().Trim();
                    if (!(ship_date.Contains(".")))
                    {
                        respuesta = "La fecha no esta en formato correcto MM.DD.YYYY";
                        xlWorkSheet.Cells[x, 9].value = respuesta;
                        continue;
                    }
                    var DMY = ship_date.Split(new char[1] { '.' });
                    dia = int.Parse(DMY[1]);
                    mes = int.Parse(DMY[0]);
                    ano = int.Parse(DMY[2]);

                    if (dia < 10)
                    { sdia = "0" + dia.ToString(); }
                    else { sdia = dia.ToString(); }
                    if (mes < 10)
                    { smes = "0" + mes.ToString(); }
                    else { smes = mes.ToString(); }

                    ship_date = sdia + "." + smes + "." + ano.ToString();

                    AWB = xlWorkSheet.Cells[x, 3].Text.ToString().Trim();
                    carrier = xlWorkSheet.Cells[x, 4].text.ToString().Trim();
                    quant = xlWorkSheet.Cells[x, 5].text.ToString().Trim();
                    PO = xlWorkSheet.Cells[x, 6].text.ToString().Trim();
                    item_num = xlWorkSheet.Cells[x, 7].text.ToString().Trim();
                    nivel_sla = xlWorkSheet.Cells[x, 8].text.ToString().Trim();
                    nivel_sla = nivel_sla.ToUpper();
                    xlWorkSheet.Cells[1, 9].value = "Resultado";

                    #region Validacion de datos

                    if (nivel_sla == "" || ship_date == "" || AWB == "" || carrier == "" || quant == "" || PO == "" || item_num == "" || nivel_sla == "")
                    {
                        respuesta = "La informacion no esta completa";
                        xlWorkSheet.Cells[x, 9].value = respuesta;
                        continue;
                    }
                    else
                    {
                        try
                        {
                            SapVariants.frame.Iconify();
                            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nYMMBL";
                            SapVariants.frame.SendVKey(0);
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtYBL_DATA-EBELN_9004")).Text = PO;
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btnP1")).Press();
                            ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6")).Select();
                            for (int i = 0; i <= 1000; i++) //for al infinito para buscar en cada linea el item del excel
                            {
                                if (i == 12)
                                {
                                    ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6")).VerticalScrollbar.Position = scroll;
                                    i = 0;
                                    scroll = scroll + 12;
                                }

                                //[columna, fila]
                                item_num_sap = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-EBELP_9004[1," + i + "]")).Text.ToString();
                                if (item_num_sap == "_____")
                                {
                                    mensaje_sap = "No se encontro el item en SAP";
                                    xlWorkSheet.Cells[x, 9].value = mensaje_sap;
                                    break;
                                }

                                if (item_num == item_num_sap)
                                {
                                    po_quant = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-MENGE[3," + i + "]")).Text.ToString();
                                    po_quant = po_quant.Replace(",000", "");
                                    sap_quant = ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-CANTI[7," + i + "]")).Text.ToString();
                                    if (sap_quant == "")
                                    {
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-CANTI[7," + i + "]")).Text = quant;
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-TRINV[8," + i + "]")).Text = AWB;
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-NAMCP[9," + i + "]")).Text = carrier;
                                        ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/ctxtWA_YBL_SHDT-PRSDT[10," + i + "]")).Text = ship_date;
                                        ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/cmbWA_YBL_SHDT-SLFLD[11," + i + "]")).Key = nivel_sla;
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[28]")).Press();
                                        SAPFEWSELib.GuiFrameWindow frame1 = (SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[1]");
                                        frame1.Iconify();
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();

                                        try
                                        {
                                            sMessage = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text;
                                        }
                                        catch (Exception)
                                        { }
                                        if (sMessage != "")
                                        {
                                            mensaje_sap = "Error al guardar";
                                        }
                                        else
                                        {
                                            mensaje_sap = "Se cargo con exito";
                                        }
                                        xlWorkSheet.Cells[x, 9].value = mensaje_sap;
                                        break;

                                    }
                                    else
                                    {
                                        sap_quant_i = Int32.Parse(sap_quant);
                                        po_quant_i = Int32.Parse(po_quant);
                                        if (sap_quant_i <= po_quant_i)
                                        {
                                            mensaje_sap = "No se guardo ningun cambio, ya estaba cargado";
                                            xlWorkSheet.Cells[x, 9].value = mensaje_sap;
                                            break;
                                        }
                                        else
                                        {
                                            ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6")).GetAbsoluteRow(i).Selected = true;

                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/btnBTN1")).Press();
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-CANTI[7," + (i + 1) + "]")).Text = "";
                                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tabsTABSTRIP/tabpTAB6/ssubSUB6:Y_MM_BACKLOG:8006/tblY_MM_BACKLOGTABCONTROL6/txtWA_YBL_SHDT-CANTI[7," + (i + 1) + "]")).Text = quant;
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                                            SAPFEWSELib.GuiFrameWindow frame1 = (SAPFEWSELib.GuiFrameWindow)SapVariants.session.FindById("wnd[1]");
                                            frame1.Iconify();
                                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();
                                            try
                                            {
                                                sMessage = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text;
                                            }
                                            catch (Exception)
                                            { }
                                            if (sMessage != "")
                                            {
                                                mensaje_sap = "Error al guardar";
                                            }
                                            else
                                            {
                                                mensaje_sap = "Se cargo con exito, con split";
                                            }
                                            xlWorkSheet.Cells[x, 9].value = mensaje_sap;
                                            break;
                                        }
                                    }



                                }
                                else
                                {
                                    mensaje_sap = "No se encontro el item en SAP";
                                    xlWorkSheet.Cells[x, 9].value = mensaje_sap;
                                }


                            } //for de lineas de sap

                        }
                        catch (Exception ex)
                        {


                        }
                    } //if si toda la info esta completa
                    #endregion Validacion de datos

                    if (xlWorkSheet.Cells[x, 9].text == "")
                    {
                        xlWorkSheet.Cells[x, 9].value = "Error al cargar en SAP";
                    }

                    console.WriteLine(xlWorkSheet.Cells[x, 9].value);
                    log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Crear backLog Guias Aereas", PO, xlWorkSheet.Cells[x, 9].value);
                    respFinal = respFinal + "\\n" + "Crear backLog Guias Aereas : " + PO + xlWorkSheet.Cells[x, 9].value;


                } //for de cada fila del excel
                sap.KillSAP();
            } //if de validacion si el archivo es el correcto
            #endregion validacion


            console.WriteLine("Respondiendo solicitud");

            xlWorkBook.SaveAs(root.FilesDownloadPath + "\\" + root.ExcelFile);
            xlWorkBook.Close();


            xlApp.DisplayAlerts = false;
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);

            string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };
            string[] cc = bsql.EmailAddress(4);


            if (devolver == true)
            {
                mail.SendHTMLMail(respuesta, new string[] { root.f_sender }, root.Subject, root.CopyCC, adjunto);
            }
            else
            {
                respuesta = "Los resultados estan en el excel";
                mail.SendHTMLMail(respuesta, new string[] { root.f_sender }, root.Subject, root.CopyCC, adjunto);
            }
        }
    }
}
