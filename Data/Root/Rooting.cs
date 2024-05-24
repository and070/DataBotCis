using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json.Linq;


namespace DataBotV5.Data.Root
{
    /// <summary>
    /// Clase Data que contiene todas las rutas del proyecto.
    /// </summary>
    class Rooting : IDisposable
    {

        #region Variables Locales

        private static string facturas_freelance = @"\\Rpaweb\fc\";
        private static string databotPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot";
        private static string h2hFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\H2H";
        private static string h2hFolderArchive = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\H2H\H2H_ARCHIVE";
        private static string CDriver = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\chromedriver";
        private static string MDriver = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\firefoxdriver";
        private static string chromeOptions = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\LICDR\chrome-options";
        private static string chromeDriverServices = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\chromedriver";
        private static string CDownloads = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\downloads";
        private static string LReport = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\LenovoReports";
        private static string FError = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\email_format\FormatoErrorOrquestador.htm";
        private static string FFirma = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\email_format\FormatoFirma.htm";
        private static string fListo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\email_format\FormatoListo.htm";
        private static string RTxtCoe = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\text_format\Coe_response.txt";
        private static string CoeXlsx = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\excel_format\CoE_Format.xlsx";
        private static string LPAXlsx = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\excel_format\Reporte_Cotizaciones_Rapidas_Licitaciones_GBPA.xlsx";
        private static string CDMS_BS = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\DMS_FILES";
        private static string BUP = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\backup_files";
        private static string FSB = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\SpecialBid_Files";
        private static string Reportes = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\reportes";
        private static string borrarArchivos = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\LICDR\DOWNLOADS";
        private static string[] ExcelLicitacionesDr = { Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\LICDR\" };
        private static string chromeDriverLog = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Databot\chromedriver\chromedriver.log";
        private static string NombreArchivoXlsX;
        private static string IdGestion;
        private static string url_dm;
        private static string user_log;
        private static string solicitante;
        private static string subject;
        private static Microsoft.Office.Interop.Outlook.MailItem mail;
        private static EmailMessage email;
        private static DateTime receiveTime;
        private static string body;
        private static string detalle_gestion;
        private static string metodo_DM;
        private static string formDmText;
        private static string fecha_DM;
        private static string data_general_DM;
        private static string factor_dm;
        private static string factorTextDm;
        private static string id_sb;
        private static string correo = "databot@gbm.net"; //"databot8061@gbmcorp.onmicrosoft.com"
        private static string[] cc;
        private static string[] rec;
        private static string type_gestion;
        private static string typeOfManagement;
        private static string type_dato;
        private static string aprob;
        private static string bs_poe_pdf;
        private static string bs_sql_sender;
        private static string bs_sql_cc1;
        private static string bs_sql_cc2;
        private static string bs_sql_cc3;
        private static string bs_sql_cc4;
        private static string bs_sql_cc5;
        private static string bs_sql_cc6;
        private static string bs_sql_cc7;
        private static string bs_sql_cc8;
        private static string bs_sql_cc9;
        private static string bs_sql_usuario;
        private static string bs_sql_po_pais;
        private static string bs_sql_po_sv;
        private static string bs_sql_customer;
        private static string bs_sql_vendor;
        private static string bs_sql_id_gestion;
        private static string bs_sql_tipo_gestion;
        private static string bs_sql_comentario;
        private static string bs_sql_status;
        private static string clase;
        private static string idClase;
        private static string metodo;
        private static bool activo;
        private static bool ejecutar;
        private static bool activate;
        private static string proceso_actual;
        private static string area;
        private static string url_cd;
        private static string urlBaw;
        private static string tokenBaw;
        // private static string s_contract;
        private static DateTime tiempo_final;
        private static string[] cfglist;
        private static string[] filelist;
        //private static string[] DM_Files; // = new ArrayList();
        private static JArray DM_Files; // = new ArrayList();
        private static string[] mq_mensajes = { "0", "", "0" };
        private static string cambio_plani;
        //private static DateTime plani_minute;
        private static Dictionary<string, DateTime> plani_minute = new Dictionary<string, DateTime>();
        private static string[] ccDr;
        private static Dictionary<string, Dictionary<string, string[]>> sc_active = new Dictionary<string, Dictionary<string, string[]>>(); //proceso, info(dias, semana, mes, tipo, contador)

        #endregion

        #region Properties de Clases
        public string DatabotPath { set { databotPath = value; } get { return databotPath; } }
        public string[] ExcelDr { set { ccDr = value; } get { return ExcelLicitacionesDr; } }
        public string borrarArchivo { set { borrarArchivos = value; } get { return borrarArchivos; } }
        public string chromeLog { set { chromeDriverLog = value; } get { return chromeDriverLog; } }
        public string ReferenciaCoe => CoeXlsx;
        public string quickQuoteReport => LPAXlsx;
        public string txtCoe => RTxtCoe;
        //public string Direccion_email => correo;
        public string Direccion_email { set { correo = value; } get { return correo; } }
        public string requestDetails { set { detalle_gestion = value; } get { return detalle_gestion; } }
        public string Chrome_Options { set { chromeOptions = value; } get { return chromeOptions; } }
        public string Chrome_Driver_Services { set { chromeDriverServices = value; } get { return chromeDriverServices; } }
        public string metodoDM { set { metodo_DM = value; } get { return metodo_DM; } }
        public string formDm { set { formDmText = value; } get { return formDmText; } }
        public string fechaDM { set { fecha_DM = value; } get { return fecha_DM; } }
        public string reportes_freelance { set { Reportes = value; } get { return Reportes; } }
        public string datagDM { set { data_general_DM = value; } get { return data_general_DM; } }
        public string factorDM { set { factor_dm = value; } get { return factor_dm; } }
        public string factorType { set { factorTextDm = value; } get { return factorTextDm; } }
        public string tipo_gestion { set { type_gestion = value; } get { return type_gestion; } }
        public string typeOfManagementText { set { typeOfManagement = value; } get { return typeOfManagement; } }
        public string aprobadorDM { set { aprob = value; } get { return aprob; } }
        public string[] CopyCC { set { cc = value; } get { return cc; } }
        public string[] recipientes { set { rec = value; } get { return rec; } }
        public string BDUserCreatedBy { set { solicitante = value; } get { return solicitante; } }
        public string Subject { set { subject = value; } get { return subject; } }
        public Microsoft.Office.Interop.Outlook.MailItem Email { set { mail = value; } get { return mail; } }
        public EmailMessage EmailObject { set { email = value; } get { return email; } }
        public DateTime ReceivedTime { set { receiveTime = value; } get { return receiveTime; } }
        public string Email_Body { set { body = value; } get { return body; } }
        public string Current_User { set { user_log = value; } get { return user_log; } }
        public string Google_Driver { set { CDriver = value; } get { return CDriver; } }
        public string Facturas_freelance { set { facturas_freelance = value; } get { return facturas_freelance; } }
        public string Mozilla_Driver { set { MDriver = value; } get { return MDriver; } }
        public string ExcelFile { set { NombreArchivoXlsX = value; } get { return NombreArchivoXlsX; } }
        public string IdGestionDM { set { IdGestion = value; } get { return IdGestion; } }
        public string FilesDownloadPath { set { CDownloads = value; } get { return CDownloads; } }
        public string LenovoReports { set { LReport = value; } get { return LReport; } }
        public string h2hDownload { set { h2hFolder = value; } get { return h2hFolder; } }
        public string h2hDownloadArchive { set { h2hFolderArchive = value; } get { return h2hFolderArchive; } }
        public string DMS_BS_Download { set { CDMS_BS = value; } get { return CDMS_BS; } }
        public string backup_root { set { BUP = value; } get { return BUP; } }
        public string SB_FILE_Download { set { FSB = value; } get { return FSB; } }
        public string Formato_Error { set { FError = value; } get { return FError; } }
        public string Formato_Firma { set { FFirma = value; } get { return FFirma; } }
        public string Formato_Listo { set { fListo = value; } get { return fListo; } }
        public string URL_DATOS_MAESTROS { set { url_dm = value; } get { return url_dm; } }
        public string id_special_bid { set { id_sb = value; } get { return id_sb; } }
        public string[] cfr_list { set { cfglist = value; } get { return cfglist; } }
        public string[] filesList { set { filelist = value; } get { return filelist; } }
        public string f_sender { set { bs_sql_sender = value; } get { return bs_sql_sender; } }
        public string f_copy1 { set { bs_sql_cc1 = value; } get { return bs_sql_cc1; } }
        public string f_copy2 { set { bs_sql_cc2 = value; } get { return bs_sql_cc2; } }
        public string f_copy3 { set { bs_sql_cc3 = value; } get { return bs_sql_cc3; } }
        public string f_copy4 { set { bs_sql_cc4 = value; } get { return bs_sql_cc4; } }
        public string f_copy5 { set { bs_sql_cc5 = value; } get { return bs_sql_cc5; } }
        public string f_copy6 { set { bs_sql_cc6 = value; } get { return bs_sql_cc6; } }
        public string f_copy7 { set { bs_sql_cc7 = value; } get { return bs_sql_cc7; } }
        public string f_copy8 { set { bs_sql_cc8 = value; } get { return bs_sql_cc8; } }
        public string f_copy9 { set { bs_sql_cc9 = value; } get { return bs_sql_cc9; } }
        public string bs_usuario { set { bs_sql_usuario = value; } get { return bs_sql_usuario; } }
        public string bs_po_pais { set { bs_sql_po_pais = value; } get { return bs_sql_po_pais; } }
        public string bs_po_sv { set { bs_sql_po_sv = value; } get { return bs_sql_po_sv; } }
        public string bs_customer { set { bs_sql_customer = value; } get { return bs_sql_customer; } }
        public string bs_vendor { set { bs_sql_vendor = value; } get { return bs_sql_vendor; } }
        public string bs_id_gestion { set { bs_sql_id_gestion = value; } get { return bs_sql_id_gestion; } }
        public string bs_tipo_gestion { set { bs_sql_tipo_gestion = value; } get { return bs_sql_tipo_gestion; } }
        public string bs_comentario { set { bs_sql_comentario = value; } get { return bs_sql_comentario; } }
        public string bs_status { set { bs_sql_comentario = value; } get { return bs_sql_comentario; } }
        public string ibm_pdf2 { set { bs_poe_pdf = value; } get { return bs_poe_pdf; } }
        public string BDClass { set { clase = value; } get { return clase; } }
        public string BDIdClass { set { idClase = value; } get { return idClase; } }
        public string BDMethod { set { metodo = value; } get { return metodo; } }
        public bool BDActive { set { activo = value; } get { return activo; } }
        public bool BDExecute { set { ejecutar = value; } get { return ejecutar; } }
        public string BDProcess { set { proceso_actual = value; } get { return proceso_actual; } }
        public bool BDActivate { set { activate = value; } get { return activate; } }
        public string BDArea { set { area = value; } get { return area; } }
        public string UrlCd { set { url_cd = value; } get { return url_cd; } }
        public string UrlBaw { set { urlBaw = value; } get { return urlBaw; } }
        public string TokenBaw { set { tokenBaw = value; } get { return tokenBaw; } }
        public JArray doc_aprob { set { DM_Files = value; } get { return DM_Files; } }
        public DateTime BDStartDate { set { tiempo_final = value; } get { return tiempo_final; } }
        public string[] Mq_mensaje { set { mq_mensajes = value; } get { return mq_mensajes; } }

        #region CrBids Variab
        public string downloadfolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\LICCR\\DOWNLOADS";
        public string optionsfolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\LICCR\\chrome-options";

        #endregion

        #region Properties para la gestión del planificador.
        public string planiChange { set { cambio_plani = value; } get { return cambio_plani; } }
        public Dictionary<string, DateTime> planiMinute { set { plani_minute = value; } get { return plani_minute; } }
        public Dictionary<string, Dictionary<string, string[]>> planner { set { sc_active = value; } get { return sc_active; } } //proceso, info(dias, semana, mes, tipo, contador)
        #endregion

        #endregion
        public void Dispose()
        {
            NombreArchivoXlsX = "";
            IdGestion = "";
            url_dm = "";
            user_log = "";
            solicitante = "";
            subject = "";
            detalle_gestion = "";
            metodo_DM = "";
            fecha_DM = "";
            data_general_DM = "";
            factor_dm = "";
            aprob = "";
            type_dato = "";
            id_sb = "";
            body = "";
            type_gestion = "";
            bs_sql_sender = "";
            bs_sql_cc1 = "";
            bs_sql_cc2 = "";
            bs_sql_cc3 = "";
            bs_sql_cc4 = "";
            bs_sql_cc5 = "";
            bs_sql_cc6 = "";
            bs_sql_cc7 = "";
            bs_sql_cc8 = "";
            bs_sql_cc9 = "";
            bs_poe_pdf = "";
            bs_sql_usuario = "";
            bs_sql_po_pais = "";
            bs_sql_vendor = "";
            bs_sql_id_gestion = "";
            bs_sql_tipo_gestion = "";
            bs_sql_comentario = "";
            bs_sql_status = "";
            clase = "";
            metodo = "";
            //s_contract = "";
            proceso_actual = "";
            area = "";
            activo = false;
            url_cd = "";
            bs_sql_po_sv = "";
            bs_sql_customer = "";
            try
            {
                if (CopyCC != null && CopyCC[0] != null)
                { Array.Clear(CopyCC, 0, CopyCC.Length); }
                doc_aprob.Clear();
                if (cfglist != null && cfglist[0] != null)
                { Array.Clear(cfglist, 0, cfglist.Length); }
                if (filelist != null && filelist[0] != null)
                { Array.Clear(filelist, 0, filelist.Length); }
                DM_Files.Clear();
            }
            catch (Exception) { }
        }

    }
}
