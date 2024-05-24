using System;
using System.IO;
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

namespace DataBotV5.Automation.RPA2.MRS
{
    /// <summary>
    /// Clase RPA Automation encargada de agregar colaboradores a MRS.
    /// </summary>
    class AddCollaboratorToMRS
    {
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        SapVariants sap = new SapVariants();
        Log log = new Log();
        Stats estadisticas = new Stats();
        string accion = "", posicion = "", colaborador = "";
        string mes = "";
        string fecha = "";
        public string response = "";
        string mandante = "ERP";
        public string response_failure = "";

        public void Main() //MRS_Main
        {
            //leer correo y descargar archivo
            console.WriteLine("Descargando archivo");
            if (mail.GetAttachmentEmail("Solicitudes MRS", "Procesados", "Procesados MRS"))
            { 
                string extArchivo = Path.GetExtension(root.FilesDownloadPath + "\\" + root.ExcelFile);
                if (extArchivo == ".xlsx")
                {
                    console.WriteLine("Procesando...");
                    ProcessMRS(root.FilesDownloadPath + "\\" + root.ExcelFile);
                    console.WriteLine("Creando Estadisticas");
                    
                }
                else
                {
                    console.WriteLine("Devolviendo Solicitud");
                    string mensaje_devolucion = "";
                    mensaje_devolucion = "Favor adjuntar unicamente archivos de Excel (.xlsx), el archivo: " +
                        root.ExcelFile + " no corresponde al tipo indicado.";
                    string[] cc = { };
                    mail.SendHTMLMail(mensaje_devolucion, new string[] { root.BDUserCreatedBy }, "Error en su gestion de MRS: " + root.Subject, cc);
                }
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }

            }

        }

        public void ProcessMRS(string route)
        {
            #region Variables Privadas ProcesarMRS
            int rows;
            string mensaje_devolucion = "";
            string validar_strc;
            bool validar_lineas = true;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkBook.Unprotect("wmuk2eer");
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            xlWorkSheet.Unprotect("wmuk2eer");
            rows = xlWorkSheet.UsedRange.Rows.Count;
            #endregion

            for (int i = 2; i <= rows; i++)
            {
                accion = xlWorkSheet.Cells[i, 1].text.ToString().Trim();
                if (accion == "" || accion == "X")
                {
                    continue;
                }
                else
                {
                    posicion = xlWorkSheet.Cells[i, 3].text.ToString().Trim();
                    if (posicion.Length > 8)
                    {
                        posicion = posicion.Substring(posicion.Length - 8, 8);
                    }
                    colaborador = xlWorkSheet.Cells[i, 2].text.ToString().Trim();
                    switch (colaborador.Length)
                    {
                        case 4:
                            colaborador = "0000" + colaborador;
                            break;
                        case 5:
                            colaborador = "000" + colaborador;
                            break;
                        case 6:
                            colaborador = "00" + colaborador;
                            break;
                        case 7:
                            colaborador = "0" + colaborador;
                            break;
                    }
                    if (DateTime.Now.Month < 10)
                    {
                        mes = "0" + DateTime.Now.Month.ToString();
                    }
                    else
                    {
                        mes = DateTime.Now.Month.ToString();
                    }

                    fecha = DateTime.Now.Year + "-" + mes + "-01";

                    #region SAP RFC
                    console.WriteLine("Corriendo RFC de SAP: " + root.BDProcess);
                    try
                    {
                        switch (accion)
                        {
                            case "A":
                                #region Parametros de SAP

                                Dictionary<string, string> parameters = new Dictionary<string, string>();
                                parameters["COLAB"] = colaborador;
                                parameters["POSITION"] = posicion;
                                parameters["FECHA_INICIO"] = fecha;
                                IRfcFunction funcA = sap.ExecuteRFC(mandante, "ZHR_PP01_REL_CREATE", parameters);


                                #endregion

                                #region Procesar Salidas del FM
                                response = response + colaborador + ": " + funcA.GetValue("RESPONSE").ToString() + "<br>";
                                //log de base de datos
                                console.WriteLine(colaborador + ": " + funcA.GetValue("RESPONSE").ToString());
                                log.LogDeCambios("Creacion", root.BDProcess, root.BDUserCreatedBy, "Agregar Colaborador a Posicion", colaborador + ": " + posicion, root.Subject);
                                if (response.Contains("error"))
                                { validar_lineas = false; }
                                #endregion
                                break;

                            case "E":
                                #region Parametros de SAP

                                Dictionary<string, string> parameters2 = new Dictionary<string, string>();
                                parameters2["COLAB"] = colaborador;
                                parameters2["POSITION"] = posicion;
                                parameters2["FECHA_INICIO"] = fecha;
                                IRfcFunction funcD = sap.ExecuteRFC(mandante, "ZHR_PP01_REL_DELETE", parameters2); 

                                #endregion
                                #region Procesar Salidas del FM
                                response = response + colaborador + ": " + funcD.GetValue("RESPONSE").ToString() + "<br>";
                                //log de base de datos
                                console.WriteLine(colaborador + ": " + funcD.GetValue("RESPONSE").ToString());
                                log.LogDeCambios("Modificacion", root.BDProcess, root.BDUserCreatedBy, "Eliminar Colaborador a Posicion", colaborador, root.Subject);
                                if (response.Contains("error"))
                                { validar_lineas = false; }
                                #endregion
                                break;
                            default:
                                response = response + colaborador + ", " + "Accion: " + accion.ToString() + " no existe o no se encuentra configurada" + "<br>";
                                break;
                        }

                      
                    }
                    #endregion
                    catch (Exception ex)
                    {
                        response_failure = val.LogErrors(ex.Message, ex.ToString(), root.ExcelFile, root.BDProcess, i);
                        console.WriteLine(" Finishing process " + response_failure);
                        response = response + colaborador + ": " + ex.ToString() + "<br>";
                        response_failure = ex.ToString();
                        validar_lineas = false;

                    }


                }

            } //for x cada fila del excel

            if (validar_lineas == false)
            {
                //enviar email de repuesta de error
                mail.SendHTMLMail(response + "<br>" + response_failure, new string[] {"appmanagement@gbm.net"}, root.Subject);
            }
            else
            {
                //enviar email de repuesta de exito
                mail.SendHTMLMail(response, new string[] { root.BDUserCreatedBy }, root.Subject, root.CopyCC);
            }

            xlApp.DisplayAlerts = false;
            xlApp.Workbooks.Close();
            xlApp.Quit();
            response = "";
            response_failure = "";
            proc.KillProcess("EXCEL",true);

        }

     
    }
}
