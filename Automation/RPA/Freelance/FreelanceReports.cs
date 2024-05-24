using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Database;
using DataBotV5.Data.Root;
using DataBotV5.Logical.Processes;
using DataBotV5.App.Global;
using DataBotV5.Data.Projects.Freelance;

namespace DataBotV5.Automation.RPA.Freelance
{
    /// <summary>
    /// Clase RPA Automation encargada de la generación de reportes freelance.
    /// </summary>
    class FreelanceReports 
    {
        Rooting root = new Rooting();
        ProcessInteraction proc = new ProcessInteraction();
        ConsoleFormat console = new ConsoleFormat();
        ValidateData val = new ValidateData();
        CRUD crud = new CRUD();
        FreelanceReportsSQL frsql = new FreelanceReportsSQL();

        public void Main()
        {
            console.WriteLine(" Verificando reporteria Freelance");
            //GenerateReport();
            console.WriteLine(" Reportes enviados o no hay reportes que enviar");
        }
    
        //private string CreateExcelPO(List<ReportPO> listado)
        //{
        //    int narchivo_final = 0;
        //    string ruta_plantilla = root.reportes_freelance + @"\plantillas\po.xlsx";
        //    string ruta_final = "";
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;

        //    xlApp = new Excel.Application();
        //    xlApp.Visible = false;
        //    xlWorkBook = xlApp.Workbooks.Open(ruta_plantilla);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
        //    narchivo_final = val.RandomNumber(10000000, 20000000);

        //    for (int i = 0; i < listado.Count; i++)
        //    {
        //        xlWorkSheet.Cells[(i + 4), 1].value = listado[i].Date;
        //        xlWorkSheet.Cells[(i + 4), 2].value = listado[i].PO;
        //        xlWorkSheet.Cells[(i + 4), 3].value = listado[i].Item;
        //        xlWorkSheet.Cells[(i + 4), 4].value = listado[i].Description;
        //        xlWorkSheet.Cells[(i + 4), 5].value = listado[i].Responsable;
        //        xlWorkSheet.Cells[(i + 4), 6].value = listado[i].Consultor;
        //        xlWorkSheet.Cells[(i + 4), 7].value = listado[i].Area;
        //    }

        //    ruta_final = root.reportes_freelance + @"\PO_REPORT_" + narchivo_final.ToString() + ".xlsx";
        //    xlWorkBook.SaveAs(ruta_final);
        //    xlWorkBook.Close();
        //    proc.KillProcess("EXCEL", true);
        //    return ruta_final;
        //}
        //private string CreateExcelHours(List<HoursReport> listado)
        //{
        //    int narchivo_final = 0;
        //    string ruta_plantilla = root.reportes_freelance + @"\plantillas\hours.xlsx";
        //    string ruta_final = "";
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;

        //    xlApp = new Excel.Application();
        //    xlApp.Visible = false;
        //    xlWorkBook = xlApp.Workbooks.Open(ruta_plantilla);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
        //    narchivo_final = val.RandomNumber(10000000, 20000000);

        //    for (int i = 0; i < listado.Count; i++)
        //    {
        //        xlWorkSheet.Cells[(i + 5), 1].value = listado[i].BusinessArea;
        //        xlWorkSheet.Cells[(i + 5), 2].value = listado[i].Employee;
        //        xlWorkSheet.Cells[(i + 5), 3].value = listado[i].Coordinator;
        //        xlWorkSheet.Cells[(i + 5), 4].value = listado[i].Client;
        //        xlWorkSheet.Cells[(i + 5), 5].value = listado[i].PO;
        //        xlWorkSheet.Cells[(i + 5), 6].value = listado[i].Item;
        //        xlWorkSheet.Cells[(i + 5), 7].value = listado[i].Date;
        //        xlWorkSheet.Cells[(i + 5), 8].value = listado[i].Ticket;
        //        xlWorkSheet.Cells[(i + 5), 9].value = listado[i].Hours;
        //        xlWorkSheet.Cells[(i + 5), 10].value = listado[i].Cats;
        //        xlWorkSheet.Cells[(i + 5), 11].value = listado[i].State;
        //        xlWorkSheet.Cells[(i + 5), 12].value = listado[i].Sheet;
        //        xlWorkSheet.Cells[(i + 5), 13].value = listado[i].ReasonSAPDesc;
        //        xlWorkSheet.Cells[(i + 5), 14].value = listado[i].ReasonEmployee;
        //        xlWorkSheet.Cells[(i + 5), 15].value = listado[i].SAPErr;
        //        xlWorkSheet.Cells[(i + 5), 16].value = listado[i].Detail;
        //    }

        //    ruta_final = root.reportes_freelance + @"\HRS_REPORT_" + narchivo_final.ToString() + ".xlsx";
        //    xlWorkBook.SaveAs(ruta_final);
        //    xlWorkBook.Close();
        //    proc.KillProcess("EXCEL", true);
        //    return ruta_final;
        //}
        //private string CrearExcelSheets(List<ReporteHoja> listado)
        //{
        //    int narchivo_final = 0;
        //    string ruta_plantilla = root.reportes_freelance + @"\plantillas\entry.xlsx";
        //    string ruta_final = "";
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;

        //    xlApp = new Excel.Application();
        //    xlApp.Visible = false;
        //    xlWorkBook = xlApp.Workbooks.Open(ruta_plantilla);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
        //    narchivo_final = val.RandomNumber(10000000, 20000000);

        //    for (int i = 0; i < listado.Count; i++)
        //    {
        //        xlWorkSheet.Cells[(i + 3), 1].value = listado[i].BusinessArea;
        //        xlWorkSheet.Cells[(i + 3), 2].value = listado[i].PO;
        //        xlWorkSheet.Cells[(i + 3), 3].value = listado[i].Item;
        //        xlWorkSheet.Cells[(i + 3), 4].value = listado[i].Hours;
        //        xlWorkSheet.Cells[(i + 3), 5].value = listado[i].Sheet;
        //        xlWorkSheet.Cells[(i + 3), 6].value = listado[i].Date;
        //        xlWorkSheet.Cells[(i + 3), 7].value = listado[i].Client;
        //        xlWorkSheet.Cells[(i + 3), 8].value = listado[i].Coordinator;
        //        xlWorkSheet.Cells[(i + 3), 9].value = listado[i].TS;
        //    }

        //    ruta_final = root.reportes_freelance + @"\HES_REPORT_" + narchivo_final.ToString() + ".xlsx";
        //    xlWorkBook.SaveAs(ruta_final);
        //    xlWorkBook.Close();
        //    proc.KillProcess("EXCEL", true);
        //    return ruta_final;
        //}
        //private string GenerateReport()
        //{
        //    string ruta = root.reportes_freelance;
        //    string ruta_archivo = "";
        //    string sql_select = "";
        //    DataTable solicitudes = frsql.Requests();
        //    if (solicitudes != null)
        //    {
        //        if (solicitudes.Rows.Count > 0)
        //        {
        //            List<LSAP> lsap = new List<LSAP>();
        //            for (int t = 0; t < 3; t++)
        //            {
        //                LSAP l1 = new LSAP();
        //                lsap.Add(l1);
        //            }

        //            lsap[0].Reason = "YB10";
        //            lsap[0].Description = "Item - Cliente incorrecto";
        //            lsap[1].Reason = "YB11";
        //            lsap[1].Description = "PO incorrecta";
        //            lsap[2].Reason = "YB12";
        //            lsap[2].Description = "Horas excedidas";


        //            for (int i = 0; i < solicitudes.Rows.Count; i++)
        //            {
        //                string iden = solicitudes.Rows[i][0].ToString();
        //                string info = solicitudes.Rows[i][1].ToString();
        //                string solicitante = solicitudes.Rows[i][2].ToString();
        //                string coordinadores = solicitudes.Rows[i][4].ToString();
        //                RequestFormat formato = JsonConvert.DeserializeObject<RequestFormat>(info);
        //                List<FreelanceCoordinators> lcord = JsonConvert.DeserializeObject<List<FreelanceCoordinators>>(coordinadores);


        //                switch (formato.Type)
        //                {
        //                    case "PO":
        //                        switch (formato.Category)
        //                        {
        //                            case "ALL":
        //                                sql_select = "SELECT * FROM freelance_po WHERE ACTIVO = 'X'";
        //                                break;
        //                            case "SMA":
        //                                sql_select = "SELECT * FROM freelance_po WHERE ACTIVO = 'X' AND AREA = 'SMA'";
        //                                break;
        //                            case "SIL":
        //                                sql_select = "SELECT * FROM freelance_po WHERE ACTIVO = 'X' AND AREA = 'SIL'";
        //                                break;
        //                            case "IBU":
        //                                sql_select = "SELECT * FROM freelance_po WHERE ACTIVO = 'X' AND AREA = 'IBU'";
        //                                break;
        //                            case "ES":
        //                                sql_select = "SELECT * FROM freelance_po WHERE ACTIVO = 'X' AND AREA = 'ES'";
        //                                break;
        //                        }

        //                        DataTable rPO = frsql.Extractor(sql_select);
        //                        if (rPO != null)
        //                        {
        //                            if (rPO.Rows.Count > 0)
        //                            {
        //                                List<ReportPO> listaPO = new List<ReportPO>();
        //                                for (int x = 0; x < rPO.Rows.Count; x++)
        //                                {
        //                                    ReportPO repo = new ReportPO();
        //                                    string ts_f = rPO.Rows[x][1].ToString();
        //                                    long ts = long.Parse(ts_f);
        //                                    System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
        //                                    dtDateTime = dtDateTime.AddMilliseconds(ts).ToLocalTime();
        //                                    repo.Date = dtDateTime.ToString();
        //                                    repo.Consultor = rPO.Rows[x][2].ToString();
        //                                    repo.PO = rPO.Rows[x][3].ToString();
        //                                    repo.Item = rPO.Rows[x][4].ToString();
        //                                    repo.Description = rPO.Rows[x][5].ToString();
        //                                    repo.Responsable = rPO.Rows[x][6].ToString();
        //                                    repo.Area = rPO.Rows[x][7].ToString();

        //                                    listaPO.Add(repo);
        //                                }
        //                                ruta_archivo = CreateExcelPO(listaPO);
        //                                frsql.SendReport(solicitante, ruta_archivo, iden, "DM & Automation Analytics: Reporte de Purchase Orders");
        //                            }
        //                        }

        //                        break;
        //                    case "REP":


        //                        if (formato.StartDate != null && formato.EndDate != null)
        //                        {
        //                            //cree el reporte
        //                            DateTime fi = DateTime.Parse(formato.StartDate);
        //                            DateTime ff = DateTime.Parse(formato.EndDate);
        //                            string fecha_i = fi.Date.ToString("yyyy-MM-dd");
        //                            string fecha_f = ff.Date.ToString("yyyy-MM-dd");
        //                            List<HoursReport> reporteHoras = new List<HoursReport>();
        //                            sql_select = "SELECT *  FROM `freelance_g` WHERE `FECHA` BETWEEN '" + fecha_i + "' AND '" + fecha_f + "' ORDER BY `ID` ASC";
        //                            DataTable rhrs = frsql.Extractor(sql_select);
        //                            if (rhrs != null)
        //                            {
        //                                if (rhrs.Rows.Count > 0)
        //                                {

        //                                    #region Crear Lista con los reusltados y filtros
        //                                    for (int x = 0; x < rhrs.Rows.Count; x++)
        //                                    {
        //                                        HoursReport rep = new HoursReport();
        //                                        rep.Coordinator = rhrs.Rows[x][2].ToString();
        //                                        int indx = lcord.FindIndex(y => y.Coordinator == rep.Coordinator);
        //                                        rep.BusinessArea = lcord[indx].Area;
        //                                        if (formato.Category == "ALL")
        //                                        {
        //                                            rep.Employee = rhrs.Rows[x][1].ToString();
        //                                            rep.Client = rhrs.Rows[x][3].ToString();
        //                                            rep.PO = rhrs.Rows[x][4].ToString();
        //                                            rep.Item = rhrs.Rows[x][5].ToString();
        //                                            rep.Date = rhrs.Rows[x][6].ToString();
        //                                            rep.Ticket = rhrs.Rows[x][7].ToString();
        //                                            rep.Detail = rhrs.Rows[x][8].ToString();
        //                                            try
        //                                            {
        //                                                byte[] data = System.Convert.FromBase64String(rep.Detail);
        //                                                string token = System.Text.ASCIIEncoding.ASCII.GetString(data);
        //                                                rep.Detail = token;
        //                                            }
        //                                            catch (Exception)
        //                                            {

        //                                                rep.Detail = "NA";

        //                                            }
        //                                            rep.Hours = rhrs.Rows[x][9].ToString();
        //                                            rep.State = rhrs.Rows[x][10].ToString();
        //                                            rep.TSCreation = rhrs.Rows[x][11].ToString();
        //                                            rep.TSAprobation = rhrs.Rows[x][12].ToString();
        //                                            rep.TSSAP = rhrs.Rows[x][13].ToString();
        //                                            rep.Cats = rhrs.Rows[x][14].ToString();
        //                                            rep.Sheet = rhrs.Rows[x][16].ToString();
        //                                            rep.ReasonSAP = rhrs.Rows[x][17].ToString();
        //                                            rep.ReasonEmployee = rhrs.Rows[x][18].ToString();
        //                                            if (rep.ReasonEmployee != "")
        //                                            {
        //                                                try
        //                                                {
        //                                                    byte[] data2 = System.Convert.FromBase64String(rep.Detail);
        //                                                    string token2 = System.Text.ASCIIEncoding.ASCII.GetString(data2);
        //                                                    rep.ReasonEmployee = token2;
        //                                                }
        //                                                catch (Exception)
        //                                                {
        //                                                    rep.ReasonEmployee = "";

        //                                                }

        //                                            }
        //                                            if (rep.ReasonSAP != "")
        //                                            {
        //                                                int indxy = lsap.FindIndex(r => r.Reason == rep.ReasonSAP);
        //                                                rep.ReasonSAPDesc = lsap[indxy].Description;
        //                                            }
        //                                            else
        //                                            {
        //                                                rep.ReasonSAPDesc = "";
        //                                            }
        //                                            rep.SAPErr = rhrs.Rows[x][19].ToString();


        //                                            reporteHoras.Add(rep);
        //                                        }
        //                                        else if (formato.Category == rep.BusinessArea)
        //                                        {
        //                                            rep.Employee = rhrs.Rows[x][1].ToString();
        //                                            rep.Client = rhrs.Rows[x][3].ToString();
        //                                            rep.PO = rhrs.Rows[x][4].ToString();
        //                                            rep.Item = rhrs.Rows[x][5].ToString();
        //                                            rep.Date = rhrs.Rows[x][6].ToString();
        //                                            rep.Ticket = rhrs.Rows[x][7].ToString();
        //                                            rep.Detail = rhrs.Rows[x][8].ToString();
        //                                            try
        //                                            {
        //                                                byte[] data = System.Convert.FromBase64String(rep.Detail);
        //                                                string token = System.Text.ASCIIEncoding.ASCII.GetString(data);
        //                                                rep.Detail = token;
        //                                            }
        //                                            catch (Exception)
        //                                            {

        //                                                rep.Detail = "NA";

        //                                            }

        //                                            rep.Hours = rhrs.Rows[x][9].ToString();
        //                                            rep.State = rhrs.Rows[x][10].ToString();
        //                                            rep.TSCreation = rhrs.Rows[x][11].ToString();
        //                                            rep.TSAprobation = rhrs.Rows[x][12].ToString();
        //                                            rep.TSSAP = rhrs.Rows[x][13].ToString();
        //                                            rep.Cats = rhrs.Rows[x][14].ToString();
        //                                            rep.Sheet = rhrs.Rows[x][16].ToString();
        //                                            rep.ReasonSAP = rhrs.Rows[x][17].ToString();
        //                                            rep.ReasonEmployee = rhrs.Rows[x][18].ToString();
        //                                            if (rep.ReasonEmployee != "")
        //                                            {
        //                                                try
        //                                                {
        //                                                    byte[] data2 = System.Convert.FromBase64String(rep.Detail);
        //                                                    string token2 = System.Text.ASCIIEncoding.ASCII.GetString(data2);
        //                                                    rep.ReasonEmployee = token2;
        //                                                }
        //                                                catch (Exception)
        //                                                {
        //                                                    rep.ReasonEmployee = "";

        //                                                }
        //                                            }
        //                                            if (rep.ReasonSAP != "")
        //                                            {
        //                                                int indxy = lsap.FindIndex(r => r.Reason == rep.ReasonSAP);
        //                                                rep.ReasonSAPDesc = lsap[indxy].Description;
        //                                            }
        //                                            else
        //                                            {
        //                                                rep.ReasonSAPDesc = "";
        //                                            }
        //                                            rep.SAPErr = rhrs.Rows[x][19].ToString();


        //                                            reporteHoras.Add(rep);
        //                                        }



        //                                    }
        //                                    #endregion

        //                                    if (reporteHoras.Count > 0)
        //                                    {
        //                                        ruta_archivo = CreateExcelHours(reporteHoras);
        //                                        frsql.SendReport(solicitante, ruta_archivo, iden, "DM & Automation Analytics: Reporte Horas Freelance");
        //                                    }
        //                                    else
        //                                    {
        //                                        string sql_update = "UPDATE reportes_freelance SET ESTADO = 'COMPLETADO' WHERE ID = '" + iden + "'";
        //                                        CRUD cr = new CRUD();
        //                                        //cr.Update("Databot", sql_update, "automation");
        //                                    }
        //                                }
        //                            }

        //                        }
        //                        else
        //                        {
        //                            string sql_update = "UPDATE reportes_freelance SET ESTADO = 'COMPLETADO' WHERE ID = '" + iden + "'";
        //                            CRUD cr = new CRUD();
        //                            //cr.Update("Databot", sql_update, "automation");
        //                        }

        //                        break;
        //                    case "HES":
        //                        sql_select = "SELECT * FROM freelance_h";
        //                        DataTable rhoja = frsql.Extractor(sql_select);
        //                        List<ReporteHoja> reporteHojas = new List<ReporteHoja>();
        //                        if (rhoja != null)
        //                        {
        //                            if (rhoja.Rows.Count > 0)
        //                            {
        //                                for (int w = 0; w < rhoja.Rows.Count; w++)
        //                                {
        //                                    ReporteHoja rh = new ReporteHoja();
        //                                    rh.Coordinator = rhoja.Rows[w][7].ToString();
        //                                    int indx = lcord.FindIndex(y => y.Coordinator == rh.Coordinator);
        //                                    rh.BusinessArea = lcord[indx].Area;
        //                                    if (formato.Category == "ALL")
        //                                    {
        //                                        rh.PO = rhoja.Rows[w][1].ToString();
        //                                        rh.Item = rhoja.Rows[w][2].ToString();
        //                                        rh.Hours = rhoja.Rows[w][3].ToString();
        //                                        rh.Sheet = rhoja.Rows[w][4].ToString();
        //                                        rh.Date = rhoja.Rows[w][5].ToString();
        //                                        rh.Client = rhoja.Rows[w][6].ToString();
        //                                        rh.TS = rhoja.Rows[w][8].ToString();
        //                                        reporteHojas.Add(rh);
        //                                    }
        //                                    else if (formato.Category == rh.BusinessArea)
        //                                    {
        //                                        rh.PO = rhoja.Rows[w][1].ToString();
        //                                        rh.Item = rhoja.Rows[w][2].ToString();
        //                                        rh.Hours = rhoja.Rows[w][3].ToString();
        //                                        rh.Sheet = rhoja.Rows[w][4].ToString();
        //                                        rh.Date = rhoja.Rows[w][5].ToString();
        //                                        rh.Client = rhoja.Rows[w][6].ToString();
        //                                        rh.TS = rhoja.Rows[w][8].ToString();
        //                                        reporteHojas.Add(rh);
        //                                    }
        //                                }
        //                                if (reporteHojas.Count > 0)
        //                                {
        //                                    ruta_archivo = CrearExcelSheets(reporteHojas);
        //                                    frsql.SendReport(solicitante, ruta_archivo, iden, "DM & Automation Analytics: Reporte HES");
        //                                }
        //                            }

        //                        }
        //                        break;

        //                }
        //            }
        //        }
        //    }




        //    return "";
        //}
    }

    public class FreelanceFiles
    {
        public string Id { get; set; }
        public string IdManagement { get; set; }
        public string Name { get; set; }
        public byte[] File { get; set; }
    }
    public class LSAP
    {
        public string Reason { get; set; }
        public string Description { get; set; }
    }
    public class HoursReport
    {
        public string BusinessArea { get; set; }
        public string Employee { get; set; }
        public string Coordinator { get; set; }
        public string Client { get; set; }
        public string PO { get; set; }
        public string Item { get; set; }
        public string Date { get; set; }
        public string Ticket { get; set; }
        public string Detail { get; set; }
        public string Hours { get; set; }
        public string State { get; set; }
        public string TSCreation { get; set; }
        public string TSAprobation { get; set; }
        public string TSSAP { get; set; }
        public string Cats { get; set; }
        public string Sheet { get; set; }
        public string ReasonSAP { get; set; }
        public string ReasonSAPDesc { get; set; }
        public string ReasonEmployee { get; set; }
        public string SAPErr { get; set; }
    }
    public class Details
    {
        public string PO { get; set; }
        public string Item { get; set; }
        public string Hours { get; set; }
    }

    public class ReportPO
    {
        public string Date { get; set; }
        public string Consultor { get; set; }
        public string PO { get; set; }
        public string Item { get; set; }
        public string Description { get; set; }
        public string Responsable { get; set; }
        public string Area { get; set; }
    }

    public class ReporteHoja
    {
        public string Date { get; set; }
        public string Client { get; set; }
        public string PO { get; set; }
        public string Item { get; set; }
        public string Hours { get; set; }
        public string Sheet { get; set; }
        public string TS { get; set; }
        public string State { get; set; }
        public string Coordinator { get; set; }
        public string BusinessArea { get; set; }
    }


    #region Analytics Freelance Classes

    public class PFreelance
    {
        /// <summary>
        /// Pos por area
        /// </summary>
        public List<POArea> Areas { get; set; }
        public List<ReportState> StatesArea { get; set; }
        public double CHRP { get; set; }
        public double CHAPRDS { get; set; }
        public double CHAPROB { get; set; }
        public int CFREE { get; set; }
    }
    /// <summary>
    /// Pos por area
    /// </summary>
    public class POArea
    {
        public string Area { get; set; }
        public double Quantity { get; set; }
    }
    public class ReportState
    {
        public string Area { get; set; }
        public int Approval { get; set; }
        public int Passed { get; set; }
        public int Rejected { get; set; }
        public int Devolution { get; set; }
        public int Error { get; set; }
        public int Completed { get; set; }
    }
    public class FreelanceCoordinators
    {
        public string Coordinator { get; set; }
        public string Area { get; set; }
    }
    public class RequestFormat
    {
        public string Type { get; set; }
        public string Category { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string Token { get; set; }
    }
    public class FreelanceVendors
    {
        public string User { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public List<string> Consultors { get; set; }
    }
    public class DeterminateVendor
    {
        public string EmailVendor { get; set; }
        public string ResponsableEmail { get; set; }
        public string NameVendor { get; set; }

        DeterminateVendor(List<FreelanceVendors> vendors, string consultor, string responsable)
        {
            int indx = vendors.FindIndex(x =>
            {
                return x.Consultors.Contains(consultor);
            });
            if (indx == -1)
            {
                //No existe una relacion del freelance con algun proveedor
            }

        }

    }
    #endregion

}
