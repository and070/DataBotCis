using System;
using Excel = Microsoft.Office.Interop.Excel;
using SAP.Middleware.Connector;
using DataBotV5.Data.Credentials;
using DataBotV5.Logical.Mail;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;
using DataBotV5.Logical.Processes;
using DataBotV5.Logical.Web;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;

namespace DataBotV5.Automation.ICS.TransportRequest

{
    /// <summary>
    /// Clase ICS Automation encargada de la administración de transportes.
    /// </summary>
    class TransportRequests 
    {
        #region Vars Globales
        Credentials cred = new Credentials();
        ConsoleFormat console = new ConsoleFormat();
        MailInteraction mail = new MailInteraction();
        Rooting root = new Rooting();
        ValidateData val = new ValidateData();
        ProcessInteraction proc = new ProcessInteraction();
        Log log = new Log();
        WebInteraction web = new WebInteraction();
        Stats estadisticas = new Stats();
        int rows = 0;
        string respuesta = "";
        string dominio_1; string sistema_1; string opcion_1_1; string opcion_2_1; string opcion_3_1;
        string opcion_4_1; string opcion_5_1; string ignorar_error_1; string s_m_1; string ambiente_1;
        string dominio_2; string sistema_2; string opcion_1_2; string opcion_2_2; string opcion_3_2;
        string opcion_4_2; string opcion_5_2; string ignorar_error_2; string s_m_2; string ambiente_2;
        #endregion

        public void Main()
        {
            console.WriteLine("Descargando archivo");
            root.ExcelFile = "Formato Transportes.xlsx";
            if (root.ExcelFile != null && root.ExcelFile != "")
            {
                console.WriteLine("Procesando...");
                ProcessTransports(root.FilesDownloadPath + "\\" + root.ExcelFile);

                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }
        public void ProcessTransports(string route)
        {
            console.WriteLine("Abriendo excel y validando");
            #region Variables Privadas
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(route);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
            rows = xlWorkSheet.UsedRange.Rows.Count;

            string mandante_1 = xlWorkSheet.Cells[2, 1].text.ToString().Trim();
            string mandante_2 = xlWorkSheet.Cells[3, 1].text.ToString().Trim();
            #endregion
            #region Ambientes-opciones de transportes
            console.WriteLine("Procesando hoja de Opciones");
            int mandante = 0;
            if (mandante_1 != "" || mandante_1 != null)
            {
                switch (mandante_1)
                {
                    case "260":
                        dominio_1 = "DOMAIN_DEV";
                        sistema_1 = "QAS";
                        ambiente_1 = "ERP";
                        s_m_1 = xlWorkSheet.Cells[2, 8].text.ToString().Trim();
                        if (s_m_1 == "" || s_m_1 == null) { s_m_1 = "S"; }
                        opcion_1_1 = xlWorkSheet.Cells[2, 2].text.ToString().Trim();
                        opcion_2_1 = xlWorkSheet.Cells[2, 3].text.ToString().Trim();
                        opcion_3_1 = xlWorkSheet.Cells[2, 4].text.ToString().Trim();
                        opcion_4_1 = xlWorkSheet.Cells[2, 5].text.ToString().Trim();
                        opcion_5_1 = xlWorkSheet.Cells[2, 6].text.ToString().Trim();
                        ignorar_error_1 = xlWorkSheet.Cells[2, 7].text.ToString().Trim();
                        mandante = 260;
                        break;
                    case "460":
                        dominio_1 = "DOMAIN_CAD";
                        sistema_1 = "CRQ";
                        ambiente_1 = "CRM";
                        s_m_1 = xlWorkSheet.Cells[2, 8].text.ToString().Trim();
                        if (s_m_1 == "" || s_m_1 == null) { s_m_1 = "S"; }
                        opcion_1_1 = xlWorkSheet.Cells[2, 2].text.ToString().Trim();
                        opcion_2_1 = xlWorkSheet.Cells[2, 3].text.ToString().Trim();
                        opcion_3_1 = xlWorkSheet.Cells[2, 4].text.ToString().Trim();
                        opcion_4_1 = xlWorkSheet.Cells[2, 5].text.ToString().Trim();
                        opcion_5_1 = xlWorkSheet.Cells[2, 6].text.ToString().Trim();
                        ignorar_error_1 = xlWorkSheet.Cells[2, 7].text.ToString().Trim();
                        mandante = 260;
                        break;
                }
            }

            if (mandante_2 != "" || mandante_2 != null)
            {
                switch (mandante_2)
                {
                    case "260":
                        dominio_2 = "DOMAIN_DEV";
                        sistema_2 = "QAS";
                        ambiente_2 = "ERP";
                        s_m_2 = xlWorkSheet.Cells[3, 8].text.ToString().Trim();
                        if (s_m_2 == "" || s_m_2 == null) { s_m_2 = "S"; }
                        opcion_1_2 = xlWorkSheet.Cells[3, 2].text.ToString().Trim();
                        opcion_2_2 = xlWorkSheet.Cells[3, 3].text.ToString().Trim();
                        opcion_3_2 = xlWorkSheet.Cells[3, 4].text.ToString().Trim();
                        opcion_4_2 = xlWorkSheet.Cells[3, 5].text.ToString().Trim();
                        opcion_5_2 = xlWorkSheet.Cells[3, 6].text.ToString().Trim();
                        ignorar_error_2 = xlWorkSheet.Cells[3, 7].text.ToString().Trim();
                        mandante = 260;
                        break;
                    case "460":
                        dominio_2 = "DOMAIN_CAD";
                        sistema_2 = "CRQ";
                        ambiente_2 = "CRM";
                        s_m_2 = xlWorkSheet.Cells[3, 8].text.ToString().Trim();
                        if (s_m_2 == "" || s_m_2 == null) { s_m_2 = "S"; }
                        opcion_1_2 = xlWorkSheet.Cells[3, 2].text.ToString().Trim();
                        opcion_2_2 = xlWorkSheet.Cells[3, 3].text.ToString().Trim();
                        opcion_3_2 = xlWorkSheet.Cells[3, 4].text.ToString().Trim();
                        opcion_4_2 = xlWorkSheet.Cells[3, 5].text.ToString().Trim();
                        opcion_5_2 = xlWorkSheet.Cells[3, 6].text.ToString().Trim();
                        ignorar_error_2 = xlWorkSheet.Cells[3, 7].text.ToString().Trim();
                        mandante = 260;
                        break;
                }
            }
            #endregion
            #region Cambio a hoja 2 para procesar transportes
            console.WriteLine("Cambio a hoja de transportes");
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[3];
            rows = xlWorkSheet.UsedRange.Rows.Count;
            #endregion
            #region Conectando al SAP RFC
            console.WriteLine("Definiendo parametros de coneccion RFC");
            RfcDestination destination = new SapVariants().GetDestRFC("", mandante);
            IRfcFunction func = destination.Repository.CreateFunction("ZSTMS_TRANSPORT");
            #endregion
            #region Llenar tabla interna de opciones
            console.WriteLine("Llenando tabla de opciones");
            IRfcTable options = func["OPTIONS"].GetTable();

            options.Append();
            options.SetValue("SISTEMA", sistema_1);
            options.SetValue("DOMINIO", dominio_1);
            options.SetValue("LATER_IMPORT", opcion_1_1);
            options.SetValue("REQUEST_AGAIN", opcion_2_1);
            options.SetValue("OW_ORIGINALS", opcion_3_1);
            options.SetValue("OW_U_REPAIRS", opcion_4_1);
            options.SetValue("IGNORE_VERSION", opcion_5_1);
            options.SetValue("IGNORE_ERROR", ignorar_error_1);
            options.SetValue("MASS_OR_SINGLE", s_m_1);
            options.SetValue("ERP_OR_CRM", ambiente_1);
            options.SetValue("CLIENTE", mandante_1);

            if (mandante_2 != "" || mandante_2 != null)
            {
                options.Append();
                options.SetValue("SISTEMA", sistema_2);
                options.SetValue("DOMINIO", dominio_2);
                options.SetValue("LATER_IMPORT", opcion_1_2);
                options.SetValue("REQUEST_AGAIN", opcion_2_2);
                options.SetValue("OW_ORIGINALS", opcion_3_2);
                options.SetValue("OW_U_REPAIRS", opcion_4_2);
                options.SetValue("IGNORE_VERSION", opcion_5_2);
                options.SetValue("IGNORE_ERROR", ignorar_error_2);
                options.SetValue("MASS_OR_SINGLE", s_m_2);
                options.SetValue("ERP_OR_CRM", ambiente_2);
                options.SetValue("CLIENTE", mandante_2);
            }
            #endregion
            #region Llenar tabla de requests(transportes)
            console.WriteLine("Llenando tabla de transportes");
            IRfcTable requests = func["REQUESTS"].GetTable();
            for (int i = 2; i <= rows; i++)
            {
                string request_id = xlWorkSheet.Cells[i, 1].text.ToString().Trim();
                string origen = xlWorkSheet.Cells[i, 2].text.ToString().Trim();
                string id_ambiente = xlWorkSheet.Cells[i, 3].text.ToString().Trim();               
                if (request_id != "" || request_id != null)
                {
                    requests.Append();
                    requests.SetValue("REQUEST_ID", request_id);
                    requests.SetValue("MANDATE_ORIGEN", origen );
                    requests.SetValue("ID_AMBIENTE", id_ambiente );
                    switch (origen)
                    {
                        case "110":
                            requests.SetValue("AMBIENTE", "ERP");
                            break;
                        case "410":
                            requests.SetValue("AMBIENTE", "CRM");
                            break;
                    }
                }
            }
            #endregion
            #region Invocar FM
            console.WriteLine("Invocando funcion RFC - Transportando Requests");
            try
            {
                  func.Invoke(destination);
                #region Procesar Respuesta
                IRfcTable response = func["RESPONSE"].GetTable();
                for (int i = 0; i < response.Count; i++)
                {
                    respuesta = respuesta + response[i].GetValue("REQUEST_ID") + "\t" + response[i].GetValue("STATUS")
                        + "\t" + response[i].GetValue("MESSAGE") + "<br>";
                }
                #endregion
            }
            catch (Exception)
            {
                respuesta += "Se produjo un error, favor verificar que los transportes se encuentren liberados.";
            }           
            #endregion          
            if (respuesta != "" || respuesta != null)
            {
                //mail.EnviarCorreo(respuesta, root.Solicitante, root.Subject, 1);
                console.WriteLine(respuesta);
            }
            xlApp.Quit();
            proc.KillProcess("EXCEL",true);
        }

    }
}
