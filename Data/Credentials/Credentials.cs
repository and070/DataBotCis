using System;
using DataBotV5.Security;
using SAP.Middleware.Connector;
using DataBotV5.Data.Root;

namespace DataBotV5.Data.Credentials
{
    /// <summary>
    /// Clase Data encargada de administración de credenciales.
    /// </summary>
    public class Credentials : IDisposable
    {
        public System.Net.NetworkCredential credenciales1 = new System.Net.NetworkCredential();
        public RfcConfigParameters parametros = new RfcConfigParameters();
        private SecureAccess login = new SecureAccess();
        Rooting root = new Rooting();

        private bool disposedValue;
        #region hide


        #region usuarios
        private static string PRD_USER = " ";
        private static string uscd = "rpausers@gbm.net"; //user CD 20240506
        private static string ussam = " @gbm.net";
        private static string ussap = " ";
        private static string ssUsername = " ";
        private static string ssUsernameQa = " ";
        private static string usdm = " @gbm.net";
        private static string ush2h = " ";
        private static string us1 = " ";
        private static string userBaw = " ";
        private static string userBawQa = " ";
        private static string _userAdminActiveDirectory = " ";
        #endregion

        #region contraseñas encriptadas
        private static readonly string pass_sam = " ";
        private static readonly string pdb = " ==";
        private static readonly string pdb1 = " ==";
        private static readonly string pdb2 = " ==";
        private static readonly string pdb3 = " ==";
        private static string contrasena_sap = " =";
        private static readonly string pdm = " ==";
        private static readonly string pss = " ==";
        private static readonly string psprd = " =";
        private static readonly string pcd = "Indic4$$23$$";
        private static readonly string pass_exchange = " =";
        private static readonly string passh2h = " =";
        private static readonly string pserverweb = " ";
        private static readonly string ppdfsb = " ==";
        private static readonly string pass_qaerp = " =";
        private static readonly string pass_qacrm = " =";
        private static readonly string pass_dev = " ";
        private static readonly string pass_devr = " ";
        private static readonly string userbpm = " ";
        private static readonly string pass_bpm = " =";
        private static readonly string pass_baw = " =";
        private static readonly string pass_baw_qa = " ";
        private static readonly string baw_admin = " =";
        private static readonly string _passAdminActiveDirectory = " ";
        private static readonly string pass_outlook = " ==";
        private static readonly string passOutlookQa = " ==";
        private static readonly string passRpauser = " =";
        private static readonly string faApiKey = " =";

        private static string passOut = "";

        #endregion

        #region variables set para pass
        private static string passsam;
        private static string db;
        private static string pscd;
        private static string pssap;
        private static string psdm;
        private static string psss;
        private static string usuario_sap;
        private static string test120;
        private static string test120pass;
        private static string PRD_PASS;
        private static string pass_exc;
        private static string pass_h2h;
        private static string pass_web;
        private static string passqaerp;
        private static string passqacrm;
        private static string passdev;
        private static string passdevr;
        private static string pass_pdf_sb;
        private static string passbpm;
        private static string userbpms;
        private static string passinvoices;
        private static string passoutlook;
        private static string passoutlookQa;
        private static string passRpauserDominio;
        private static string passbaw;
        private static string passbawqa;
        private static string bawadmin;
        #endregion

        #endregion

        public string pass_db1 { set { db = value; } get { return login.DecodePass(pdb1); } }
        public string pass_db2 { set { db = value; } get { return login.DecodePass(pdb2); } }
        public string pass_db3 { set { db = value; } get { return login.DecodePass(pdb3); } }
        public string pass_db_GN { set { db = value; } get { return db; } }
        public string username_db { set { us1 = value; } get { return us1; } }
        public string username_ss { set { ssUsername = value; } get { return ssUsername; } }
        public string username_ss_qa { set { ssUsernameQa = value; } get { return ssUsernameQa; } }
        public string username_dm { set { usdm = value; } get { return usdm; } }
        public string password_dm { set { psdm = value; } get { return login.DecodePass(pdm); } }
        public string password_ss { set { psss = value; } get { return login.DecodePass(pss); } }
        public string usernameSAP { set { ussap = value; } get { return ussap; } }
        public string passwordSAP { set { pssap = value; } get { return login.DecodePass(contrasena_sap); } }
        public string username_CD { set { uscd = value; } get { return uscd; } }
        public string password_CD { set { pscd = value; } get { return pcd; } }
        public string username_SAP { set { usuario_sap = value; } get { return usuario_sap; } }
        public string password_SAP { set { contrasena_sap = value; } get { return login.DecodePass(contrasena_sap); } }
        public string password_QA_ERP { set { passqaerp = value; } get { return login.DecodePass(pass_qaerp); } }
        public string password_QA_CRM { set { passqacrm = value; } get { return login.DecodePass(pass_qacrm); } }
        public string password_DEVr { set { passdev = value; } get { return login.DecodePass(pass_dev); } }
        public string password_DEV { set { passdevr = value; } get { return login.DecodePass(pass_devr); } }
        public string username_SAPPRD { set { PRD_USER = value; } get { return PRD_USER; } }
        public string password_SAPPRD { set { PRD_PASS = value; } get { return login.DecodePass(psprd); } }
        public string fusionUser { set { ussam = value; } get { return ussam; } }
        public string fusionPassword { set { passsam = value; } get { return login.DecodePass(pass_sam); } }
        public string password_exchange { set { pass_exc = value; } get { return login.DecodePass(pass_exchange); } }
        public string usuario_h2h { set { ush2h = value; } get { return ush2h; } }
        public string password_h2h { set { pass_h2h = value; } get { return login.DecodePass(passh2h); } }
        public string password_server_web { set { pass_web = value; } get { return login.DecodePass(pserverweb); } }
        public string password_pdf_sb { set { pass_pdf_sb = value; } get { return login.DecodePass(ppdfsb); } }
        public string password_bpm { set { passbpm = value; } get { return login.DecodePass(pass_bpm); } }
        public string user_bpm { set { userbpms = value; } get { return login.DecodePass(userbpm); } }
        public string password_outlook { set { passoutlook = value; } get { return login.DecodePass(pass_outlook); } }
        public string password_outlook_qa { set { passoutlookQa = value; } get { return login.DecodePass(passOutlookQa); } }
        public string password_rpauser_dominio { set { passRpauserDominio = value; } get { return login.DecodePass(passRpauser); } }
        public string passOutlook { set { passOut = value; } get { return passOut; } }
        public string userAdminActiveDirectory { set { _userAdminActiveDirectory = value; } get { return _userAdminActiveDirectory; } }
        public string passAdminActiveDirectory { set { } get { return login.DecodePass(_passAdminActiveDirectory); } }
        public string fusionAuthApiKey { set { } get { return login.DecodePass(faApiKey); } }

        #region Baw 
        public string user_baw { set { userBaw = value; } get { return userBaw; } }
        public string password_baw { set { passbaw = value; } get { return login.DecodePass(pass_baw); } }
        public string user_baw_qa { set { userBawQa = value; } get { return userBawQa; } }
        public string password_baw_qa { set { passbawqa = value; } get { return login.DecodePass(pass_baw_qa); } }
        #endregion
        public string BawAdmin { set { bawadmin = value; } get { return login.DecodePass(baw_admin); } }


        #region Claves del servidor de Smart and Simple

        private static string passSsQa;
        private static string passSsPrd;
        private static string pass_SS_QA = " ";
        private static string pass_SS_PRD = " ";
        public string passSmartSimpleServerQA { set { passSsQa = value; } get { return login.DecodePass(pass_SS_QA); } }
        public string passSmartSimpleServerPRD { set { passSsPrd = value; } get { return login.DecodePass(pass_SS_PRD); } }
        #endregion

        #region nueva conexion
        public string PRD_DATA_BASE_SERVER { get { return PRD_DATABASE_IP; } set { PRD_DATABASE_IP = value; } }
        public string QA_DATA_BASE_SERVER { get { return QA_DATABASE_IP; } set { QA_DATABASE_IP = value; } }
        public string DEV_DATA_BASE_SERVER { get { return DEV_DATABASE_IP; } set { DEV_DATABASE_IP = value; } }

        private static string PRD_DATABASE_IP = "10.7.60.72";
        private static string QA_DATABASE_IP = "localhost";
        private static string DEV_DATABASE_IP = "localhost";

        private static string PRD_DATABASE_USER = "databot";
        private static string QA_DATABASE_USER = "databot";
        private static string DEV_DATABASE_USER = "databot";

        private static string PRD_DATABASE_PASS = "";
        private static string QA_DATABASE_PASS = "";
        private static string DEV_DATABASE_PASS = "UqJkkoxRVkIXSJYf";

        public string PRD_DATA_BASE_SERVER_USER { get { return PRD_DATABASE_USER; } set { PRD_DATABASE_USER = value; } }
        public string QA_DATA_BASE_SERVER_USER { get { return QA_DATABASE_USER; } set { QA_DATABASE_USER = value; } }
        public string DEV_DATA_BASE_SERVER_USER { get { return DEV_DATABASE_USER; } set { DEV_DATABASE_USER = value; } }
        public string PRD_DATA_BASE_SERVER_PASS { get { return PRD_DATABASE_PASS; } set { PRD_DATABASE_PASS = value; } }
        public string QA_DATA_BASE_SERVER_PASS { get { return QA_DATABASE_PASS; } set { QA_DATABASE_PASS = value; } }
        public string DEV_DATA_BASE_SERVER_PASS { get { return DEV_DATABASE_PASS; } set { DEV_DATABASE_PASS = value; } }

        #endregion

        #region nueva conexion Smart and Simple
        public string PRD_SS_BASE_SERVER { get { return PRD_SS_DATABASE_IP; } set { PRD_SS_DATABASE_IP = value; } }
        public string PRD_SS_APP_SERVER { get { return PRD_SS_APP_IP; } set { PRD_SS_APP_IP = value; } }
        public string QA_SS_BASE_SERVER { get { return QA_SS_DATABASE_IP; } set { QA_SS_DATABASE_IP = value; } }
        public string QA_SS_APP_SERVER_USER { get { return QA_SS_APP_USER; } set { QA_SS_APP_USER = value; } }
        public string PRD_SS_APP_SERVER_USER { get { return PRD_SS_APP_USER; } set { PRD_SS_APP_USER = value; } }

        private static string PRD_SS_DATABASE_IP = "10.7.60.138";
        private static string PRD_SS_APP_IP = "10.7.60.122";
        private static string QA_SS_DATABASE_IP = "10.7.60.151";
        private static string QA_SS_APP_USER = " ";
        private static string PRD_SS_APP_USER = " ";

        #endregion

        #region APIs MSECOH S&S

        private static string SS_URL_MSECOH = "https://smartsimple.gbm.net:43888";
        private static string IN_MSECOH_UPDATE_DATE_API = $"{SS_URL_MSECOH}/secoh/update-target-start-date-request-applied";
        private static string IN_MSECOH_GET_DATE_API = $"{SS_URL_MSECOH}/secoh/find-all-target-start-date-request-by-apply";
        private static string IN_MSECOH_CREATE_CONTRACT_API = $"{SS_URL_MSECOH}/secoh/create-contract-on-hold";
        private static string IN_MSECOH_UPDATE_CONTRACT_API = $"{SS_URL_MSECOH}/secoh/update-contract-on-hold-by-id/";
        private static string IN_MSECOH_FIND_CONTRACT_ONHOLD_API = $"{SS_URL_MSECOH}/secoh/find-contracts-on-hold/";
        public string MESCOH_UPDATE_DATE { get { return IN_MSECOH_UPDATE_DATE_API; } set { MESCOH_UPDATE_DATE = value; } }
        public string MESCOH_GET_DATE { get { return IN_MSECOH_GET_DATE_API; } set { MESCOH_GET_DATE = value; } }
        public string MESCOH_CREATE_CONTRACT { get { return IN_MSECOH_CREATE_CONTRACT_API; } set { MESCOH_CREATE_CONTRACT = value; } }
        public string MESCOH_UPDATE_CONTRACT { get { return IN_MSECOH_UPDATE_CONTRACT_API; } set { MESCOH_UPDATE_CONTRACT = value; } }
        public string MESCOH_GET_CONTRACT_ONHOLD { get { return IN_MSECOH_FIND_CONTRACT_ONHOLD_API; } set { IN_MSECOH_FIND_CONTRACT_ONHOLD_API = value; } }

        private static string SS_URL_MSECOH_QA = "https://smartsimple-qa2.gbm.net:23888";
        private static string IN_MSECOH_UPDATE_DATE_API_QA = $"{SS_URL_MSECOH_QA}/secoh/update-target-start-date-request-applied";
        private static string IN_MSECOH_GET_DATE_API_QA = $"{SS_URL_MSECOH_QA}/secoh/find-all-target-start-date-request-by-apply";
        private static string IN_MSECOH_CREATE_CONTRACT_API_QA = $"{SS_URL_MSECOH_QA}/secoh/create-contract-on-hold";
        private static string IN_MSECOH_UPDATE_CONTRACT_API_QA = $"{SS_URL_MSECOH_QA}/secoh/update-contract-on-hold-by-id/";
        private static string IN_MSECOH_FIND_CONTRACT_ONHOLD_API_QA = $"{SS_URL_MSECOH_QA}/secoh/find-contracts-on-hold/";
        public string MESCOH_UPDATE_DATE_QA { get { return IN_MSECOH_UPDATE_DATE_API_QA; } set { MESCOH_UPDATE_DATE_QA = value; } }
        public string MESCOH_GET_DATE_QA { get { return IN_MSECOH_GET_DATE_API_QA; } set { MESCOH_GET_DATE_QA = value; } }
        public string MESCOH_CREATE_CONTRACT_QA { get { return IN_MSECOH_CREATE_CONTRACT_API_QA; } set { MESCOH_CREATE_CONTRACT_QA = value; } }
        public string MESCOH_UPDATE_CONTRACT_QA { get { return IN_MSECOH_UPDATE_CONTRACT_API_QA; } set { MESCOH_UPDATE_CONTRACT_QA = value; } }
        public string MESCOH_GET_CONTRACT_ONHOLD_QA { get { return IN_MSECOH_FIND_CONTRACT_ONHOLD_API_QA; } set { IN_MSECOH_FIND_CONTRACT_ONHOLD_API_QA = value; } }


        #endregion

        #region Credenciales Microsoft API (azure app: databotPOC)

        private static readonly string cid = " "; //" 
        private static readonly string cs = "eWpXOFF+ =="; //" 
        private static readonly string tid = " "; //"9 
        private static string cliId;
        private static string cliSec;
        private static string tenId;
        public string clientId { set { cliId = value; } get { return login.DecodePass(cid); } }
        public string clientSecret { set { cliSec = value; } get { return login.DecodePass(cs); } }
        public string tenantId { set { tenId = value; } get { return login.DecodePass(tid); } }
        #endregion
        public void EstablecerCredenciales(string user, string pass)
        {
            credenciales1.UserName = user;
            credenciales1.Password = pass;
        }
        /// <summary>
        /// Destructor
        /// </summary>
        ~Credentials()
        {
            parametros.Clear();
        }

        public void IngresarAmbiente(int ambiente)
        {
            parametros.Clear();

            switch (ambiente)
            {
                case 110:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.11.111");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_DEV);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "DEV");
                    parametros.Add(RfcConfigParameters.Name, "NSP110");
                    //parametros.Add(RfcConfigParameters.UseSAPGui, "1");
                    break;

                case 120:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.11.111");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_DEV);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "DEV");
                    parametros.Add(RfcConfigParameters.Name, "NSP120");
                    break;

                case 420:
                case 410:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.11.110");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_DEV);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "DEV");
                    parametros.Add(RfcConfigParameters.Name, "NSP410");
                    break;
                case 260:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.11.117");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_QA_ERP);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "QAS");
                    parametros.Add(RfcConfigParameters.Name, "NSP260");
                    break;
                case 460:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.11.113");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_QA_CRM);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "QAS");
                    parametros.Add(RfcConfigParameters.Name, "NSP460");
                    break;
                case 300:

                    parametros.Add(RfcConfigParameters.AppServerHost, "ecc-prod-app.gbm.net");
                    parametros.Add(RfcConfigParameters.SystemNumber, "01");
                    parametros.Add(RfcConfigParameters.User, username_SAPPRD);
                    parametros.Add(RfcConfigParameters.Password, password_SAPPRD);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "ECC");
                    parametros.Add(RfcConfigParameters.Name, "NSP300");
                    break;
                case 500:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.11.29");
                    parametros.Add(RfcConfigParameters.SystemNumber, "01");
                    parametros.Add(RfcConfigParameters.User, username_SAPPRD);
                    parametros.Add(RfcConfigParameters.Password, password_SAPPRD);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "CRP");
                    parametros.Add(RfcConfigParameters.Name, "NSP500");
                    break;

                case 100:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.60.41");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_DEVr);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "FID");
                    parametros.Add(RfcConfigParameters.Name, "NSP100");
                    break;
                case 400:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.60.43");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_DEVr);
                    parametros.Add(RfcConfigParameters.Client, ambiente.ToString());
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "FIP");
                    parametros.Add(RfcConfigParameters.Name, "NSP400");
                    break;
                case 4002:
                    parametros.Add(RfcConfigParameters.AppServerHost, "10.7.60.49");
                    parametros.Add(RfcConfigParameters.SystemNumber, "00");
                    parametros.Add(RfcConfigParameters.User, "RPAUSER");
                    parametros.Add(RfcConfigParameters.Password, password_DEVr);
                    parametros.Add(RfcConfigParameters.Client, "400");
                    parametros.Add(RfcConfigParameters.Language, "EN");
                    parametros.Add(RfcConfigParameters.SystemID, "FIQ");
                    parametros.Add(RfcConfigParameters.Name, "NSP400");
                    break;
            }
        }
        public bool ConnectDB()
        {
            bool status = false;
            if (login.DecodePass(pdb) == pass_db1)
            {
                status = true;
            }
            return status;
        }
        public bool ConnectDM()
        {
            bool status = false;
            if (login.DecodePass(pdm) == password_dm)
            {
                status = true;
            }
            return status;
        }
        public bool ConnectSAP()
        {
            bool status = false;
            if (login.DecodePass(psprd) == password_SAPPRD)
            {
                status = true;
            }
            return status;
        }
        public bool ConnectCD()
        {
            bool status = false;
            if (login.DecodePass(pcd) == password_CD)
            {
                status = true;
            }
            return status;
        }

        public void SelectCdMand(string id)
        {
            switch (id)
            {
                case "PRD2":
                    root.UrlCd = "http://10.7.60.20";
                    break;
                case "PRD":
                    root.UrlCd = "https://controldesk.gbm.net";
                    break;
                case "DEV":
                    root.UrlCd = "http://controldesk-dev.gbm.net";
                    break;
                case "QAS":
                    root.UrlCd = "http://controldesk-qa.gbm.net";
                    break;
            }
        }
        /// <summary>
        /// DEV, QAS o PRD
        /// </summary>
        /// <param name="id"></param>
        public void SelectBawMand(string id)
        {
            switch (id)
            {
                case "PRD":
                    root.UrlBaw = "https://prod-ihs-03.gbm.net";
                    break;
                case "DEV":
                    root.UrlBaw = "https://des-pc-01-clone.gbm.net:9443";
                    break;
                case "QAS":
                    root.UrlBaw = "https://test-ihsbaw-01.gbm.net";
                    break;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ConsoleFormat()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            //Dispose(disposing: true);
            //GC.SuppressFinalize(this);
        }
    }
}

