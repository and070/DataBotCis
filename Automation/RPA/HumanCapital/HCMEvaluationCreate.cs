using ClosedXML.Excel;
using SAP.Middleware.Connector;
using System;
using System.Data;
using System.IO;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.MicrosoftTools;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using System.Collections.Generic;

namespace DataBotV5.Automation.RPA.HumanCapital
{
    /// <summary>
    /// Clase RPA Automation encargada de creación de la productividad de un colaborador (Infotipo 9003 y 19G1).
    /// </summary>
    class HCMEvaluationCreate
    {
        
        Rooting root = new Rooting();
        MsExcel MsExcel = new MsExcel();
        MailInteraction mail = new MailInteraction();
        ConsoleFormat console = new ConsoleFormat();
        Credentials cred = new Credentials();
        Log log = new Log();
        Database DB = new Database();
        ValidateData val = new ValidateData();
        SapVariants sap = new SapVariants();
        string respFinal = "";

        string mandante = "ERP";
        public void Main()
        {
            if (mail.GetAttachmentEmail("Solicitudes Evaluacion HCM", "Procesados", "Procesados Evaluacion HCM"))
            {
                console.WriteLine("Procesando...");
                EvaCreate(root.FilesDownloadPath + "\\" + root.ExcelFile);
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        private void EvaCreate(string route)
        {
            bool valLines = true;
            DataSet excelBook = MsExcel.GetExcelBook(route);
            DataTable excel = excelBook.Tables["Resumen_2"];
            DataTable excelResult = new DataTable();
            excelResult.Columns.Add("País");
            excelResult.Columns.Add("Manager");
            excelResult.Columns.Add("ID Colaborador");
            excelResult.Columns.Add("Nombre Completo");
            excelResult.Columns.Add("Fecha de Contratacion");
            excelResult.Columns.Add("Sub Area de Personal");
            excelResult.Columns.Add("Posición");
            excelResult.Columns.Add("Resultado Performance");
            excelResult.Columns.Add("Resultado");
            if (excel == null)
            {
                mail.SendHTMLMail("Error al leer la plantilla de Productividad", new string[] { root.BDUserCreatedBy }, "Error al leer la plantilla de productividad, verifique el nombre de la hoja sea Resumen_2 o bien el título de las columnas",  new string[] { "appmanagement@gbm.net" });
                return;
            }

            foreach (DataRow row in excel.Rows)
            {
                string response = "";
                string idColaborador = "";
                string notaF = "";
                try
                {
                    idColaborador = row["ID Colaborador"].ToString();
                    if (idColaborador != "")
                    {
                        notaF = (float.Parse(row["Resultado Performance"].ToString()) * 100).ToString();
                        evaInfo info = new evaInfo();
                        info.EmployeeNumber = idColaborador;
                        info.NotaF = notaF;
                        info.Evaluator = row["User(ID) Supervisor"].ToString();
                        response = CreateEvaluation(info);
                        if (response.Contains("Error"))
                        {
                            valLines = false;
                        }
                    }



                    DataRow rRow = excelResult.Rows.Add();
                    rRow["País"] = row["Id Compannia"].ToString();
                    rRow["Manager"] = row["User(ID) Supervisor"].ToString();
                    rRow["ID Colaborador"] = idColaborador;
                    rRow["Nombre Completo"] = row["Nombre Completo"].ToString();
                    rRow["Fecha de Contratacion"] = row["Fecha de Contratacion"].ToString();
                    rRow["Sub Area de Personal"] = row["Sub Area de Personal"].ToString();
                    rRow["Posición"] = row["Posición"].ToString();
                    rRow["Resultado Performance"] = notaF;
                    rRow["Resultado"] = response;
                    excelResult.AcceptChanges();
                    log.LogDeCambios("Creacion", root.BDProcess, "HCM", "Carga de Evaluación de Productividad", idColaborador + " : " + notaF, response);
                    respFinal = respFinal + "\\n" + idColaborador + " : " + notaF;
                }
                catch (Exception ex)
                {
                    valLines = false;
                    response = ex.Message;
                }

            }

            console.WriteLine("Save Excel...");
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(excelResult, "Resultados");
            route = root.FilesDownloadPath + $"\\Resultados Creacion de Evaluaciones {DateTime.Now.ToString("yyyyMMddHHmmssffff")}.xlsx";
            if (File.Exists(route))
            {
                File.Delete(route);
            }
            wb.SaveAs(route);
            string msj = "Estimado(a) se le adjunta el Excel con los resultados de la creación de las evaluaciones de productividad al día de hoy.";
            string html = Properties.Resources.emailtemplate1;
            html = html.Replace("{subject}", "Creación de Infotipo 19G1 y 9003");
            html = html.Replace("{cuerpo}", msj);
            html = html.Replace("{contenido}", "");
            console.WriteLine("Send Email...");
            try
            {
                mail.SendHTMLMail(html, new string[] { root.BDUserCreatedBy }, $"Notificacion solicitud carga de evaluaciones de productividad", root.CopyCC, new string[] { route });
            }
            catch (Exception ex)
            {
                mail.SendHTMLMail("Error al responder email de productividad " + ex.Message, new string[] { "dmeza@gbm.net" }, "Error solicitud carga de evaluaciones de productividad",  null, null);
            }
            if (!valLines)
            {
                mail.SendHTMLMail(html, new string[] {"appmanagement@gbm.net"}, $"Error: Notificacion solicitud carga de evaluaciones de productividad", new string[] { "dmeza@gbm.net", root.BDUserCreatedBy }, new string[] { route });
            }

            root.requestDetails = respFinal;


        }
        private string CreateEvaluation(evaInfo info)
        {

            try
            {
                info.NotaF = info.NotaF.Replace(",", ".");
                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    ["EMPLOYEENUMBER"] = info.EmployeeNumber,
                    ["NOTAF"] = info.NotaF,
                    ["STARTD"] = "0001-01-01",
                    ["ENDAT"] = "0001-01-01",
                    ["EVALUADOR"] = info.Evaluator
                };

                IRfcFunction createEvaluation = sap.ExecuteRFC(mandante, "ZHR_CREATE_PA9003", parameters);


                console.WriteLine(createEvaluation.GetValue("RESPONSE").ToString());
                return createEvaluation.GetValue("RESPONSE").ToString();
            }
            catch (Exception ex)
            {
                return $"Error: {ex}";
            }

        }
    }
    public class evaInfo
    {
        public string EmployeeNumber { get; set; }
        public string NotaF { get; set; }
        public string Evaluator { get; set; }
    }
}
