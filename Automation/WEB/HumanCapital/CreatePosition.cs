using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;

namespace DataBotV5.Automation.WEB.HumanCapital
{
    /// <summary>
    /// Clase WEB Automation encargada de crear posiciones no planificadas de Human Capital a partir de la página de AM.
    /// </summary>
    class CreatePosition
    {
        #region variables globales
        ConsoleFormat console = new ConsoleFormat();
        public string response = "";
        public string response_failure = "";
        Credentials cred = new Credentials();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        SapVariants sap = new SapVariants();
        Stats estadisticas = new Stats();
        string mandante = "ERP";
        string respFinal = "";


        #endregion

        public void Main()  //Create_Pos_NoPlan_Main
        {
            //revisa si el usuario RPAUSER esta abierto
            if (!sap.CheckLogin(mandante))
            {

                if (mail.GetAttachmentEmail("Solicitudes Posicion No Planificada", "Procesados", "Procesados Posicion No Planificada"))
                {
                    for (int w = 0; w <= root.filesList.Length - 1; w++)
                    {
                        if (root.filesList[w].Length >= 21)
                        {
                            if (root.filesList[w].Substring(0, 21).ToString() == "PosicionNoPlanificada")
                            {
                                root.ExcelFile = root.filesList[w].ToString();
                                break;
                            }
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(root.ExcelFile))
                    {
                        sap.BlockUser(mandante, 1);
                        console.WriteLine("Procesando...");
                        ProcessPositionNoPlan(root.FilesDownloadPath + "\\" + root.ExcelFile);
                        response = "";
                        sap.BlockUser(mandante, 0);
                        console.WriteLine("Creando Estadisticas");
                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                    }
                }
             
            }
        }
        public void ProcessPositionNoPlan(string route)
        {
            #region Variables Privadas


            string SUBT9050 = "", SUBT9070 = "", SUBT9080 = "", SUBT9100 = "", SUBT9110 = "", SUBT9120 = "", SUBT9135 = "", SUBT9140 = "";
            string SUBT9150 = "", SUBT9170 = "", SUBT9190 = "", SUBT9200 = "", SUBT9210 = "", SUBT9220 = "", SUBT9230 = "", TipoSolicitud = "";
            string Desde = "", ObjetoDesc = "", PoscQueAprobaraRecurso = "", UnidadOrganiza = "", PosicionPadre = "", Funcion = "", Sociedad = "";
            string DivisionPers = "", SubdivPers = "", CentroCoste = "", GrupoPers = "", AreaPers = "", Comentarios = "";
            string usuario = "", Function_id = "", company_code = "", localRegionalSalaryScale = ""; string local_regional = "";

            int scroll = 12;
            int columna = 14;
            string SUBT = "", SUBT_value = "";
            string respuesta;

            int rows;
            string mensaje_devolucion = "";
            bool validar_lineas = true;
            respuesta = "";
            string validacion = "";
            #endregion

            #region abrir excel
            console.WriteLine("Abriendo Excel y Validando");

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            rows = xlWorkSheet.UsedRange.Rows.Count;

            #endregion

            validacion = xlWorkSheet.Cells[1, 1].text.ToString().Trim();

            if (validacion != "TipoSolicitud")
            {
                mensaje_devolucion = "Utilizar la plantilla oficial de la pagina de AM";
                validar_lineas = false;
            }
            else
            {
                for (int i = 2; i <= rows; i++)
                {
                    usuario = xlWorkSheet.Cells[i, 31].text.ToString().Trim();

                    if (usuario == "")
                    {
                        respuesta = "Ingrese toda la informacion";
                        continue;
                    }
                    else // si hay data
                    {
                        #region extraer data y validacion

                        SUBT9050 = xlWorkSheet.Cells[i, 14].text.ToString().Trim(); SUBT9070 = xlWorkSheet.Cells[i, 15].text.ToString().Trim();
                        SUBT9080 = xlWorkSheet.Cells[i, 16].text.ToString().Trim(); SUBT9100 = xlWorkSheet.Cells[i, 17].text.ToString().Trim();
                        SUBT9110 = xlWorkSheet.Cells[i, 18].text.ToString().Trim(); SUBT9120 = xlWorkSheet.Cells[i, 19].text.ToString().Trim();
                        SUBT9135 = xlWorkSheet.Cells[i, 20].text.ToString().Trim(); SUBT9140 = xlWorkSheet.Cells[i, 21].text.ToString().Trim();
                        SUBT9150 = xlWorkSheet.Cells[i, 22].text.ToString().Trim(); SUBT9170 = xlWorkSheet.Cells[i, 23].text.ToString().Trim();
                        SUBT9190 = xlWorkSheet.Cells[i, 24].text.ToString().Trim(); SUBT9200 = xlWorkSheet.Cells[i, 25].text.ToString().Trim();
                        SUBT9210 = xlWorkSheet.Cells[i, 26].text.ToString().Trim(); SUBT9220 = xlWorkSheet.Cells[i, 27].text.ToString().Trim();
                        SUBT9230 = xlWorkSheet.Cells[i, 28].text.ToString().Trim(); TipoSolicitud = xlWorkSheet.Cells[i, 1].text.ToString().Trim();
                        Desde = xlWorkSheet.Cells[i, 2].text.ToString().Trim(); ObjetoDesc = xlWorkSheet.Cells[i, 3].text.ToString().Trim();
                        PoscQueAprobaraRecurso = xlWorkSheet.Cells[i, 4].text.ToString().Trim(); UnidadOrganiza = xlWorkSheet.Cells[i, 5].text.ToString().Trim();
                        PosicionPadre = xlWorkSheet.Cells[i, 6].text.ToString().Trim(); Funcion = xlWorkSheet.Cells[i, 7].text.ToString().Trim();
                        Sociedad = xlWorkSheet.Cells[i, 8].text.ToString().Trim();
                        DivisionPers = xlWorkSheet.Cells[i, 9].text.ToString().Trim(); SubdivPers = xlWorkSheet.Cells[i, 10].text.ToString().Trim();
                        CentroCoste = xlWorkSheet.Cells[i, 11].text.ToString().Trim(); GrupoPers = xlWorkSheet.Cells[i, 12].text.ToString().Trim();
                        AreaPers = xlWorkSheet.Cells[i, 13].text.ToString().Trim(); Comentarios = xlWorkSheet.Cells[i, 30].text.ToString().Trim();
                        localRegionalSalaryScale = xlWorkSheet.Cells[i, 29].text.ToString().Trim();

                        Comentarios = Comentarios.Replace("\n", "");
                        Comentarios = Comentarios.Replace("\r", "");

                        #endregion

                        #region Extraer ID de la function de SAP
                        console.WriteLine("Extraer Funcion ID de SAP: " + root.BDProcess);

                        try
                        {
                            company_code = (Sociedad == "GBCR") ? Sociedad + "_" + SubdivPers : Sociedad;
                            //localRegionalSalaryScale = (Sociedad == "GBCO") ? "001" : localRegionalSalaryScale;

                            Dictionary<string, string> parametros = new Dictionary<string, string>();
                            parametros["FUNCTION"] = Funcion;
                            parametros["COMPANY_CODE"] = company_code;
                            parametros["LOCAL_REGIONAL"] = localRegionalSalaryScale;
                            IRfcFunction func = sap.ExecuteRFC(mandante, "ZHR_GET_JOB", parametros);



                            Function_id = func.GetValue("FUNCTION_ID").ToString();
                            console.WriteLine(Funcion + " : " + Function_id);
                            if (Function_id == "")
                            {

                                local_regional = (localRegionalSalaryScale == "001") ? "Local" : "Regional";
                                respuesta = "No se encontro la Function: " + Funcion + ". A nivel: " + local_regional + ". En el pais: " + company_code;
                                validar_lineas = false;
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            response_failure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + response_failure);
                            respuesta = ObjetoDesc + ": " + ex.ToString();
                            response_failure = ex.ToString();
                            validar_lineas = false;
                            continue;
                        }

                        #endregion

                        #region cargar la posicion en SAP
                        console.WriteLine(" Cargar la posicion en SAP");

                        sap.LogSAP(mandante.ToString());
                        try
                        {
                            // SAP_Variants.frame.Iconify();
                            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nzhr_wf14";
                            SapVariants.frame.SendVKey(0);
                            ((SAPFEWSELib.GuiCheckBox)SapVariants.session.FindById("wnd[0]/usr/chkZHRCP014-VACAN")).Selected = true;
                            ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/cmbZHRCP014-TPO_SOL")).Key = TipoSolicitud;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtZHRCP014-NUPOCR")).Text = "1";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-BEGDA")).Text = Desde;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-ENDDA")).Text = "31.12.9999";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtZHRCP014-SHORT")).Text = "Spendiente";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/txtZHRCP014-STEXT")).Text = ObjetoDesc;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-GERENTE")).Text = PoscQueAprobaraRecurso;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-BUKRS")).Text = Sociedad;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-PERSG")).Text = GrupoPers;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-WERKS")).Text = DivisionPers;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-PERSK")).Text = AreaPers;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-BTRTL")).Text = SubdivPers;
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP014-KOSTL")).Text = CentroCoste;
                            string sapMsj = "";
                            SapVariants.frame.SendVKey(0);
                            try
                            {
                                sapMsj = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                            }
                            catch (Exception)
                            { }

                            if (!string.IsNullOrWhiteSpace(sapMsj))
                            {
                                validar_lineas = false;
                                respuesta = sapMsj;
                                sap.KillSAP();
                                continue;
                            }

                            try
                            {
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btn%#AUTOTEXT029")).Press();
                                //   frame1.Iconify();
                                try
                                {
                                    //cambia el modo de busqueda a search tearm (en caso de que no este)
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[26]")).Press();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/subSUBSCR_SELONE:SAPLSDH4:0130/sub:SAPLSDH4:0130/btnG_SELONE_STATE-BUTTON_TEXT[0,0]")).Press();
                                }
                                catch (Exception)
                                { }
                                ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001")).Select();
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]")).Text = UnidadOrganiza;
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                respuesta = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                                if (respuesta.Contains("No values"))
                                {
                                    validar_lineas = false;
                                    respuesta = "La Posicion Encargada/Aprobadora no existe en SAP";
                                    sap.KillSAP();
                                    continue;
                                }
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                            }
                            catch (Exception)
                            {
                                validar_lineas = false;
                                respuesta = "La Unidad Organizativa no existe en SAP";
                                sap.KillSAP();
                                continue;
                            }

                            try
                            {
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btn%#AUTOTEXT030")).Press();
                                // frame1.Iconify();
                                try
                                {
                                    //cambia el modo de busqueda a search tearm (en caso de que no este)
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[26]")).Press();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/subSUBSCR_SELONE:SAPLSDH4:0130/sub:SAPLSDH4:0130/btnG_SELONE_STATE-BUTTON_TEXT[0,0]")).Press();
                                }
                                catch (Exception)
                                { }
                                ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001")).Select();
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]")).Text = PosicionPadre;
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                respuesta = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                                if(respuesta.Contains("No values"))
                                {
                                    validar_lineas = false;
                                    respuesta = "La Posicion Encargada/Aprobadora no existe en SAP";
                                    sap.KillSAP();
                                    continue;
                                }
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                            }
                            catch (Exception)
                            {
                                validar_lineas = false;
                                respuesta = "La Posicion Encargada/Aprobadora no existe en SAP";
                                sap.KillSAP();
                                continue;
                            }

                            try
                            {
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/usr/btn%#AUTOTEXT028")).Press();
                                // frame1.Iconify();
                                try
                                {
                                    //cambia el modo de busqueda a search tearm (en caso de que no este)
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[26]")).Press();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/usr/subSUBSCR_SELONE:SAPLSDH4:0130/sub:SAPLSDH4:0130/btnG_SELONE_STATE-BUTTON_TEXT[0,0]")).Press();
                                }
                                catch (Exception)
                                { }
                                ((SAPFEWSELib.GuiTab)SapVariants.session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001")).Select();
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]")).Text = Function_id;
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                respuesta = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                                if (respuesta.Contains("No values"))
                                {
                                    validar_lineas = false;
                                    respuesta = "La Posicion Encargada/Aprobadora no existe en SAP";
                                    sap.KillSAP();
                                    continue;
                                }
                                ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                            }
                            catch (Exception)
                            {
                                validar_lineas = false;
                                respuesta = "El Job/Function no existe en SAP";
                                sap.KillSAP();
                                continue;
                            }


                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tblZHRPG_NUEVAS_POSICIONESTC_ZHRCP014P/ctxtZHRCP014P-SUBTY[0,0]")).Text = "9055";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tblZHRPG_NUEVAS_POSICIONESTC_ZHRCP014P/ctxtZHRCP014P-HILFM[2,0]")).Text = "002";

                            for (int e = 1; e < 15; e++)
                            {
                                if (e == 12)
                                {
                                    ((SAPFEWSELib.GuiTableControl)SapVariants.session.FindById("wnd[0]/usr/tblZHRPG_NUEVAS_POSICIONESTC_ZHRCP014P")).VerticalScrollbar.Position = scroll;
                                    e = 0;
                                    scroll = scroll + 12;
                                }
                                SUBT = xlWorkSheet.Cells[1, columna].text.ToString().Trim();
                                if (SUBT.Substring(0, 4) != "SUBT")
                                {
                                    break;
                                }
                                SUBT = SUBT.Substring(4, 4);
                                SUBT_value = xlWorkSheet.Cells[i, columna].text.ToString().Trim();
                                if (SUBT == "9170" && SUBT_value == "000" || SUBT == "9100" && SUBT_value == "000" || SUBT == "9110" && SUBT_value == "000"
                                    || SUBT == "9120" && SUBT_value == "0" || SUBT == "9135" && SUBT_value == "0")
                                {
                                    e = e - 1;
                                    columna += 1;
                                    continue;
                                }

                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tblZHRPG_NUEVAS_POSICIONESTC_ZHRCP014P/ctxtZHRCP014P-SUBTY[0," + e + "]")).Text = SUBT;
                                ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/tblZHRPG_NUEVAS_POSICIONESTC_ZHRCP014P/ctxtZHRCP014P-HILFM[2," + e + "]")).Text = SUBT_value;
                                columna += 1;

                            }

                            ((SAPFEWSELib.GuiTextedit)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_NUEVAS_POSICIONES:1001/cntlCTRL_TEXT/shellcont/shell")).SetSelectionIndexes(17, 17);

                            if (Comentarios == "NA")
                            {
                                Comentarios = "Usuario Solicitante: " + usuario + ". No hay mas comentarios";
                            }
                            else
                            {
                                Comentarios = Comentarios + ". Usuario Solicitante: " + usuario;
                            }


                            ((SAPFEWSELib.GuiTextedit)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_NUEVAS_POSICIONES:1001/cntlCTRL_TEXT/shellcont/shell")).Text = Comentarios;

                            if (root.filesList != null && root.filesList[0] != null)
                            {
                                for (int w = 0; w <= root.filesList.Length - 1; w++)
                                {
                                    if (root.filesList[w].Length >= 21)
                                    {
                                        if (root.filesList[w].Substring(0, 21).ToString() == "PosicionNoPlanificada")
                                        {
                                            continue;
                                        }
                                    }
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_NUEVAS_POSICIONES:1001/ctxtFILE")).Text = "";
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_NUEVAS_POSICIONES:1001/ctxtFILE")).SetFocus();
                                    SapVariants.frame.SendVKey(4);
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = root.FilesDownloadPath + "\\";
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = root.filesList[w].ToString();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                }
                            }
                            ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[0]/tbar[1]/btn[20]")).Press();
                            try
                            {
                                respuesta = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString();
                            }
                            catch (Exception)
                            { }

                        }
                        catch (Exception ex)
                        {
                            try
                            { mensaje_devolucion = ((SAPFEWSELib.GuiStatusbar)SapVariants.session.FindById("wnd[0]/sbar")).Text.ToString(); }
                            catch (Exception) { }
                            response_failure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + response_failure);
                            respuesta = ObjetoDesc + ": " + mensaje_devolucion + "<br>" + root.ExcelFile + "<br>" + "<br>" + ex.ToString();

                            response_failure = ex.ToString();
                            validar_lineas = false;
                            sap.KillSAP();
                            continue;
                        }
                        sap.KillSAP();

                        #endregion


                        //log de base de datos
                        log.LogDeCambios("Creacion", root.BDProcess, usuario, "Crear Posicion no planificada", ObjetoDesc + " : " + respuesta, root.Subject);
                        respFinal= respFinal + "\\n" + $"Crear Posicion no planificada {usuario}" + ObjetoDesc + " : " + respuesta;
                    }
                } //for

            } //else de validation
            xlApp.DisplayAlerts = false;
            xlWorkBook.Close();
            xlApp.Workbooks.Close();
            xlApp.Quit();
            proc.KillProcess("EXCEL", true);
            console.WriteLine("Respondiendo solicitud");
            root.requestDetails = respFinal;

            if (respuesta == "")
            {
                respuesta = "Hubo un error al crear la posicion, por favor verifique la data";
                validar_lineas = false;
            }

            if (validar_lineas == false)
            {
                //enviar email de repuesta de error
                string[] cc = { usuario, "gvillalobos@gbm.net" };
                mail.SendHTMLMail(respuesta, new string[] {"appmanagement@gbm.net"}, root.Subject + " - Error", cc);
            }
            else
            {
                string[] cc = { "gvillalobos@gbm.net" };
                //enviar email de repuesta de exito
                mail.SendHTMLMail(respuesta, new string[] { usuario }, root.Subject, cc);
            }
        }

    }
}
