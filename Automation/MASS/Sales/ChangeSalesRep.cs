using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;
using DataBotV5.Data.SAP;
using DataBotV5.Logical.Projects.UpdateClients;
using ClosedXML.Excel;

namespace DataBotV5.Automation.MASS.Sales
{
    /// <summary>
    ///Clase MASS Automation encargada del cambio de representante de ventas.
    /// </summary>
    class ChangeSalesRep
    {
        public string response = "";
        public string response_failure = "";
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        UpdateClients updateClients = new UpdateClients();
        Stats estadisticas = new Stats();
        string Cliente = "";
        string Representante = "";
        string respuesta = "";
        string fmrep = "";
        string validacion = "";
        string dia_mes = "";
        string mandante = "ERP";
        string enviroment = "QAS";
        string respFinal = "";


        public void Main()
        {
            dia_mes = DateTime.Now.Day.ToString();
            if (dia_mes == "13" || dia_mes == "14" || dia_mes == "15")
            {//no se cambia representantes estos dias
                console.WriteLine("Dia:" + dia_mes + ", no se actualizan empleados responsables");
            }
            else
            {
                console.WriteLine("Descargando archivo");
                //leer correo y descargar archivo
                if (mail.GetAttachmentEmail("Solicitudes Representantes", "Procesados", "Procesados Representantes"))
                {
                    console.WriteLine("Procesando...");
                    ProcessSalesRep(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    response = "";
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }


            }


        }

        public void ProcessSalesRep(string route)
        {

            #region Variables Privadas
            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;
            respuesta = "";
            #endregion

            console.WriteLine("Abriendo Excel y Validando");
            #region abrir excel
            Excel.Range xlRango;
            Excel.Range xlRangoDuplicate;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            rows = xlWorkSheet.UsedRange.Rows.Count;
            #endregion

            validacion = xlWorkSheet.Cells[1, 2].text.ToString().Trim();

            if (validacion.Substring(0, 1) != "x")
            {
                console.WriteLine("Devolviendo Solicitud");
                respuesta = "Utilizar la plantilla oficial de datos maestros";
                xlApp.DisplayAlerts = false;
                xlApp.Workbooks.Close();
                xlApp.Quit();
                respuesta = "";
                proc.KillProcess("EXCEL",true);
                mail.SendHTMLMail(respuesta, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);

            }
            else
            {
                xlWorkBook.Unprotect("dominogbm12%");
                xlWorkSheet.Unprotect("dominogbm12%");
                xlWorkSheet.Cells[1, 3].value = "Resultado";
                xlWorkSheet.Range["A1"].Copy();
                Microsoft.Office.Interop.Excel.XlPasteType paste = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats;
                Microsoft.Office.Interop.Excel.XlPasteSpecialOperation pasteop = Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;
                xlWorkSheet.Range["C1"].PasteSpecial(paste, pasteop, false, false);

                //Planntilla correcta, continue las validaciones
                for (int i = 2; i <= rows; i++)
                {

                    Cliente = xlWorkSheet.Cells[i, 1].text.ToString().Trim();

                    if (Cliente == "")
                    {
                        continue;
                    }
                    else
                    {

                        Cliente = Cliente.Replace("\n", "");
                        Cliente = Cliente.Replace("\r", "");
                        bool Numeric = int.TryParse(Cliente, out int num);
                        if (Numeric == false)
                        {
                            respuesta = respuesta + Cliente + " - " + Representante + ": " + "el Cliente no es el ID" + "<br>";
                            continue;
                        }
                        if (Cliente.Substring(0, 2) != "00")
                        { Cliente = "00" + Cliente; }

                        Representante = xlWorkSheet.Cells[i, 2].text.ToString().Trim().ToUpper();
                        Representante = Representante.Replace("\n", "");
                        Representante = Representante.Replace("\r", "");
                        Representante = Representante.Replace("AA", "");
                        bool isNumeric = int.TryParse(Representante, out int n);
                        if (isNumeric == false)
                        {
                            respuesta = respuesta + Cliente + " - " + Representante + ": " + "el representante no es el ID del colaborador" + "<br>";
                            continue;
                        }

                        #region SAP
                        console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);

                        try
                        {
                            Dictionary<string, string> parametros = new Dictionary<string, string>(); 
                            parametros["BP"] = Cliente;
                            parametros["COLABORADOR"] = Representante;

                            IRfcFunction func = sap.ExecuteRFC(mandante, "ZDM_RPA_VTS_001", parametros);
                            #region Procesar Salidas del FM

                            if (func.GetValue("RESPUESTA").ToString() == "OK")
                            {
                                fmrep = "El representante ha sido actualizado";
                                //actualizar la BD entidades
                                updateClients.UpdateEntyties(Cliente, Representante, enviroment);
                                updateClients.UpdateClient(Cliente, Representante);
                            }
                            else if (func.GetValue("RESPUESTA").ToString() == "ID invalido")
                            { fmrep = "El representante no existe en SAP"; }
                            else if (func.GetValue("RESPUESTA").ToString() == "ERROR")
                            { fmrep = "Error en el momento de actualizar el representante"; }
                            else
                            { fmrep = "Error inesperado"; }

                            respuesta = respuesta + Cliente + " - " + Representante + ": " + fmrep + "<br>";
                            console.WriteLine(Cliente + " - " + Representante + ": " + fmrep);
                            //log de base de datos
                            log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Modificar Sales Rep", Cliente + " - " + Representante + ": " + fmrep, root.Subject);
                            respFinal = respFinal + "\\n" + Cliente + " - " + Representante + ": " + fmrep;


                            if (respuesta.Contains("Error"))
                            { validar_lineas = false; }

                            xlWorkSheet.Cells[i, 3].value = fmrep;
                            xlWorkSheet.Range["A" + i].Copy();
                            xlWorkSheet.Range["C" + i].PasteSpecial(paste, pasteop, false, false);
                            //log de cambios
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            response_failure = new ValidateData().LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                            console.WriteLine(" Finishing process " + response_failure);
                            respuesta = respuesta + Cliente + " - " + Representante + ": " + ex.ToString() + "<br>";
                            response_failure = ex.ToString();
                            validar_lineas = false;

                        }

                        #endregion

                    }
                } //for de cada linea del excel



                if (rows > 30)
                {
                    try { xlWorkBook.SaveAs(route); } catch(Exception e) { xlWorkBook.Save(); }                    
                    xlWorkSheet.Protect("dominogbm12%");
                    xlWorkBook.Protect("dominogbm12%");
                    xlWorkBook.Close();
                }

                xlApp.DisplayAlerts = false;
                xlApp.Workbooks.Close();
                xlApp.Quit();
                proc.KillProcess("EXCEL",true);
                console.WriteLine("Respondiendo solicitud");
                if (validar_lineas == false)
                {
                    string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };
                    //enviar email de repuesta de error
                    mail.SendHTMLMail(respuesta + "<br>" + response_failure, new string[] { "internalcustomersrvs@gbm.net" }, root.Subject, attachments: adjunto);
                }
                else
                {
                    string[] adjunto = { root.FilesDownloadPath + "\\" + root.ExcelFile };
                    //enviar email de repuesta de exito
                    mail.SendHTMLMail(respuesta, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC, adjunto);
                }
                root.requestDetails = respFinal;


            }



        }


    }
}
