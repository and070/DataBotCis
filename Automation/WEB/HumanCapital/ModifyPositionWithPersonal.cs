using System;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;

namespace DataBotV5.Automation.WEB.HumanCapital
{
    /// <summary>
    /// Clase WEB Automation encargada de modificar la posición 
    /// </summary>
    class ModifyPositionWithPersonal
    {
        #region variables globales
        public string response = "";
        public string response_failure = "";
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
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
        public void Main()
        {
            //revisa si el usuario RPAUSER esta abierto
            if (!sap.CheckLogin(mandante))
            {
                if (mail.GetAttachmentEmail("Solicitudes Posicion con Personal", "Procesados", "Procesados Posicion con Personal"))

                {
                    for (int w = 0; w <= root.filesList.Length - 1; w++)
                    {
                        if (root.filesList[w].Length >= 19)
                        {
                            if (root.filesList[w].ToString().Contains("PosicionConPersonal"))
                            {
                                root.ExcelFile = root.filesList[w].ToString();
                                break;
                            }
                        }
                    }
                    if (root.ExcelFile != null && root.ExcelFile != "")
                    {
                        sap.BlockUser(mandante, 1);
                        console.WriteLine("Procesando...");
                        ProcessPositionWithPersonal(root.FilesDownloadPath + "\\" + root.ExcelFile);
                        response = "";
                        sap.BlockUser(mandante, 0);

                        using (Stats stats = new Stats())
                        {
                            stats.CreateStat();
                        }
                    }
                }
           
            }
        }
        public void ProcessPositionWithPersonal(string route)
        {
            #region Variables Privadas


            string id_colaborador = "", fecha_solicitud = "", Comentarios = "";
            string usuario = "";
            string titulo = "", valor = "";
            string comentario_wf = "";
            string respuesta;

            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;
            bool devolver = false;
            respuesta = "";
            string validacion = "";

            #endregion

            #region abrir excel
            console.WriteLine("Abriendo Excel y Validando");
            Excel.Range xlRango;
            Excel.Range xlRangoDuplicate;

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

            if (validacion != "ID Colaborador")
            {
                mensaje_devolucion = "Utilizar la plantilla oficial de la pagina de AM";
                validar_lineas = false;
            }
            else
            {
                string respFinal = "";
                for (int i = 2; i <= rows; i++)
                {
                    usuario = xlWorkSheet.Cells[i, 18].text.ToString().Trim();

                    if (usuario == "")
                    {
                        respuesta = "Ingrese toda la informacion";
                        continue;
                    }
                    else // si hay data
                    {
                        #region extraer data y crear comentario

                        Comentarios = xlWorkSheet.Cells[i, 17].text.ToString().Trim();
                        id_colaborador = xlWorkSheet.Cells[i, 1].text.ToString().Trim();
                        fecha_solicitud = xlWorkSheet.Cells[i, 4].text.ToString().Trim();

                        Comentarios = Comentarios.Replace("\n", "");
                        Comentarios = Comentarios.Replace("\r", "");

                        comentario_wf = "Campos a modificar:" + "\r\n";
                        for (int e = 2; e <= 16; e++)
                        {
                            if (e == 4)
                            {
                                //fecha de solicitud
                                continue;
                            }
                            titulo = xlWorkSheet.Cells[1, e].text.ToString().Trim();
                            valor = xlWorkSheet.Cells[i, e].text.ToString().Trim();
                            if (valor != "")
                            {
                                comentario_wf = comentario_wf + titulo + ": " + valor + "\r\n";
                            }
                        }
                        comentario_wf = comentario_wf + "\r\n";
                        comentario_wf = comentario_wf + "Fecha de vigencia:" + "\r\n";
                        comentario_wf = comentario_wf + fecha_solicitud + "\r\n";
                        comentario_wf = comentario_wf + "\r\n";
                        comentario_wf = comentario_wf + "Comentarios de la solicitud:" + "\r\n";
                        comentario_wf = comentario_wf + Comentarios + "\r\n";
                        comentario_wf = comentario_wf + "\r\n";
                        comentario_wf = comentario_wf + "Usuario Solicitante: " + "\r\n";
                        comentario_wf = comentario_wf + usuario;

                        #endregion


                        #region cargar la posicion en SAP
                        console.WriteLine(" Cargar la posicion en SAP");

                        sap.LogSAP(mandante.ToString());
                        try
                        {
                            // SAP_Variants.frame.Iconify();
                            ((SAPFEWSELib.GuiOkCodeField)SapVariants.session.FindById("wnd[0]/tbar[0]/okcd")).Text = "/nZHR_WF19";
                            SapVariants.frame.SendVKey(0);
                            ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/cmbZHRCP019-TPO_SOL")).Key = "12";
                            ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/ctxtZHRCP019-EMPLEADO")).Text = id_colaborador;
                            SapVariants.frame.SendVKey(0);
                            ((SAPFEWSELib.GuiTextedit)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_WF_OTROS:1001/cntlCTRL_TEXT/shellcont/shell")).Text = comentario_wf;

                            if (root.filesList != null && root.filesList[0] != null)
                            {
                                for (int w = 0; w <= root.filesList.Length - 1; w++)
                                {
                                    if (root.filesList[w].Length >= 19)
                                    {
                                        if (root.filesList[w].ToString().Contains("PosicionConPersonal"))
                                        {
                                            continue;
                                        }
                                    }
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_WF_OTROS:1001/ctxtFILE")).Text = "";
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[0]/usr/subSUB_SCREEN:ZHRPG_WF_OTROS:1001/ctxtFILE")).SetFocus();
                                    SapVariants.frame.SendVKey(4);
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = root.FilesDownloadPath + "\\";
                                    ((SAPFEWSELib.GuiTextField)SapVariants.session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = root.filesList[w].ToString();
                                    ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                    try
                                    {
                                        //EN CASO DE QUE SALGA UN POP DE REPETIDO
                                        ((SAPFEWSELib.GuiButton)SapVariants.session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                                    }
                                    catch (Exception)
                                    {

                                    }


                                }
                            }
                            //nuevo campo motivo
                            //session.findById("wnd[0]/usr/cmbZHRCP019-MOTIVO").key = "04"
                            ((SAPFEWSELib.GuiComboBox)SapVariants.session.FindById("wnd[0]/usr/cmbZHRCP019-MOTIVO")).Key = "04";
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
                            response_failure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + response_failure);
                            respuesta = id_colaborador + ": " + mensaje_devolucion + "<br>" + "<br>" + ex.ToString();
                            response_failure = ex.ToString();
                            validar_lineas = false;
                            sap.KillSAP();
                            continue;
                        }
                        sap.KillSAP();

                        #endregion


                        //log de base de datos
                        log.LogDeCambios("Modificacion", root.BDProcess, usuario, "Modificar Posicion con Personal", id_colaborador + " : " + respuesta, root.Subject);
                        respFinal = respFinal + "\\n" + $"Modificar Posicion con Personal {usuario}" + id_colaborador + " : " + respuesta;

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
