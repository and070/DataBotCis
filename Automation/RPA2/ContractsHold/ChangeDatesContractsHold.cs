using Newtonsoft.Json;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using StackExchange.Redis;
using DataBotV5.Data.Credentials;
using DataBotV5.Data.Database;
using DataBotV5.Data.Process;
using DataBotV5.App.Global;
using DataBotV5.Data.SAP;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.RPA2.ContractsHold
{
    /// <summary>
    /// Clase RPA Automation encargada de procesar los request de cambio de fechas de S & S de _ContractsHold.
    /// </summary>
    class ChangeDatesContractsHold
    {
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        string mand = "ERP";
        bool executeStats = false;
        public void Main()
        {
           

            //cada 5 min            
                console.WriteLine(" Starting process: ProcessOnHoldDates");
                ProcessOnHoldDates();
            

        }

        #region Métodos del API
        /// <summary>
        /// Método que se utiliza para actualizar los en S & S las fechas de contrato una vez éstas sean actualizadas en SAP-CRM.
        /// </summary>
        /// <param name="ids">Una lista de los id_Contract de tipo entero.</param>
        private void UpdateDateAPI(List<int> ids)
        {
            try
            {
                Credentials cred = new Credentials();
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                string url = cred.MESCOH_UPDATE_DATE;
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "PUT";
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    ContractDate cds = new ContractDate
                    {
                        contractsIDs = ids
                    };
                    string json = JsonConvert.SerializeObject(cds);

                    streamWriter.Write(json);
                }
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    console.WriteLine(result);
                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message.ToString());
            }
        }
        /// <summary>
        /// Método que obtiene las solicitudes de cambio de fecha generados por S & S.
        /// </summary>
        /// <returns></returns>
        private PostingRequest GetDateRequests()
        {
            PostingRequest pr = new PostingRequest();

            try
            {
                Credentials cred = new Credentials();
                ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);
                string url = cred.MESCOH_GET_DATE;

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "GET";
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    pr.requestData = JsonConvert.DeserializeObject<RequestResponse>(result);
                    if (pr.requestData.status == 200)
                    {
                        pr.processOk = true;
                    }
                    else
                    {
                        pr.processOk = false;
                    }
                }

            }
            catch (Exception ex)
            {
                pr.processOk = false;
                console.WriteLine(ex.Message.ToString());
            }


            return pr;
        }

        #endregion
        #region Metodos Main
        /// <summary>
        /// Método que se encarga de procesar los request de cambio de fechas de S & S.
        /// </summary>
        /// <param name="mandante">Mandante de SAP a modificar</param>
        private void ProcessOnHoldDates()
        {
            //Obtiene los requests de las fechas a cambiar
            PostingRequest pr = GetDateRequests();
            //Verifica que estado del request sea 200 ( True ), que el payload no sea nulo y que existan requests
            if (pr.processOk && pr.requestData.payload.requests != null && pr.requestData.payload.requests.Count > 0)
            {
                executeStats = true;
                List<int> ids = new List<int>();
                //Para cada request del objeto, realice el cambio de fecha en SAP
                (pr.requestData.payload.requests).ForEach(x =>
                {
                    //Metodo RFC que cambia la fecha con base en el tipo deseado (CONTSIGNON)
                    ProcessDateChangeRFC(x.contractNumber, "CONTSIGNON", x.newTargetDate);
                    //Agrega el id del contrato (API numero interno de S&S) y lo agrega a una lista de enteros
                    ids.Add(x.id_Contract);
                });
                //Actualiza los request de los contratos en S&S
                UpdateDateAPI(ids);
                using (Stats stats = new Stats())
                {
                    stats.CreateStat();
                }
            }
        }

        #endregion
        #region Metodos de validación y utilidad
        public bool AcceptAllCertifications(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        /// <summary>
        /// Formato de fecha RFC de SAP
        /// </summary>
        /// <param name="current">Fecha extraída del RFC</param>
        /// <returns>Retorna formato de fecha YYYY-MM-DD</returns>
        private string FormatDate(string current)
        {
            return $"{current.Substring(0, 4)}-{current.Substring(4, 2)}-{current.Substring(6, 2)}";
        }
        /// <summary>
        /// Parsea la información de una estructura de SAP al formato requerido para uso del API de S & S.
        /// </summary>
        /// <param name="valores">Estructura RFC</param>
        /// <returns>Retorna un objeto con los valores en el formato deseado.</returns>
        private ContratosHoldParseOut ParseOut(IRfcStructure valores)
        {
            ContratosHoldParseOut parseOut = new ContratosHoldParseOut();
            List<ItemsTBOut> listado = new List<ItemsTBOut>();

            //Campos Integer

            string idc = valores.GetString("ID_CONTRACT");

            long id_contrato = long.Parse(idc);
            int id_customer = (int)long.Parse(valores.GetString("ID_CUSTOMER"));

            //Campos fechas

            string start_date = FormatDate(valores.GetString("START_DATE"));
            string posting_date = FormatDate(valores.GetString("POSTING_DATE"));
            string end_date = FormatDate(valores.GetString("END_DATE"));
            string onhold_date = FormatDate(valores.GetString("ONHOLD_DATE"));

            //Campos Float

            double gross_value = double.Parse(((valores.GetString("GROSS_VALUE").Trim()).Replace('.', ',')));
            double net_value = double.Parse(((valores.GetString("NET_VALUE").Trim()).Replace('.', ',')));
            double tax_value = double.Parse(((valores.GetString("TAX_VALUE").Trim()).Replace('.', ',')));
            double shipment_value = double.Parse(((valores.GetString("SHIPMENT_VALUE").Trim()).Replace('.', ',')));

            parseOut.ID_CONTRACT = id_contrato;
            parseOut.DESC_CONTRACT = valores.GetString("DESC_CONTRACT");
            parseOut.CTY_CONTRACT = valores.GetString("CTY_CONTRACT");
            parseOut.ID_CUSTOMER = id_customer;
            parseOut.DESC_CUSTOMER = valores.GetString("DESC_CUSTOMER");
            parseOut.EXTERNAL_REF = valores.GetString("EXTERNAL_REF");
            parseOut.START_DATE = start_date;
            parseOut.GROSS_VALUE = gross_value;
            parseOut.NET_VALUE = net_value;
            parseOut.TAX_VALUE = tax_value;
            parseOut.SHIPMENT_VALUE = shipment_value;
            parseOut.POSTING_DATE = posting_date;
            parseOut.END_DATE = end_date;
            parseOut.ONHOLD_DATE = onhold_date;
            parseOut.ONHOLD_USER = valores.GetString("ONHOLD_USER");
            parseOut.CURRENT_STATUS = valores.GetString("CURRENT_STATUS");
            parseOut.SERVICE = valores.GetString("SERVICE");
            parseOut.OUTSOURCING = valores.GetString("OUTSOURCING");
            parseOut.RENEWAL_TYPE = valores.GetString("RENEWAL_TYPE");
            parseOut.VARIABLE_CONTRACT = valores.GetString("VARIABLE_CONTRACT");
            parseOut.SALES_ORG = valores.GetString("SALES_ORG");
            parseOut.SERVICE_ORG = valores.GetString("SERVICE_ORG");
            parseOut.EMPLOYEE = valores.GetString("EMPLOYEE");
            parseOut.EMPLOYEE_ID = valores.GetString("EMPLOYEE_ID");

            IRfcTable items = valores.GetTable("ITEMS_TB");

            foreach (IRfcStructure item in items)
            {
                double service_qty = 0;
                if (item.GetValue("SERVICE_QTY") != null && item.GetValue("SERVICE_QTY").ToString() != "")
                {
                    service_qty = double.Parse(((item.GetString("SERVICE_QTY").Trim()).Replace('.', ',')));
                }
                ItemsTBOut ito = new ItemsTBOut
                {
                    SERVICE_MG = item.GetString("SERVICE_MG"),
                    SERVICE_NAME = item.GetString("SERVICE_NAME"),
                    SERVICE_PRODUCT = item.GetString("SERVICE_PRODUCT"),
                    SERVICE_QTY = service_qty,
                };
                listado.Add(ito);
            }

            parseOut.ITEMS_TB = listado;

            return parseOut;
        }
        #endregion
        #region Metodos SAP RFC
        /// <summary>
        /// Método que cambia la fecha de un contrato en SAP CRM.
        /// </summary>
        /// <param name="id">Id del contrato</param>
        /// <param name="type">Tipo de fecha a cambiar</param>
        /// <param name="date">Fecha en formato YYYY-MM-DD</param>
        /// <param name="mand">Mandante de SAP a modificar</param>
        private void ProcessDateChangeRFC(string id, string type, string date)
        {
            Credentials cred = new Credentials();

            RfcDestination destination = sap.GetDestRFC(mand);
            RfcRepository repo = destination.Repository;
            IRfcFunction updateContractDates = repo.CreateFunction("ZPUT_UPDATE_DATES");
            updateContractDates.SetValue("ID", id);
            updateContractDates.SetValue("DATE", date); 
            updateContractDates.SetValue("TYPE_REL", type);
            updateContractDates.Invoke(destination); //Actualiza fecha con UTC
            updateContractDates.Invoke(destination); //Actualiza fecha real
            console.WriteLine(updateContractDates.GetValue("RESPONSE").ToString());
        }

        #endregion
    }
    #region Clases de parseo JSON
    public class ContratosHoldParseIn
    {
        public string ID_CONTRACT { get; set; }
        public string DESC_CONTRACT { get; set; }
        public string CTY_CONTRACT { get; set; }
        public string ID_CUSTOMER { get; set; }
        public string DESC_CUSTOMER { get; set; }
        public string EXTERNAL_REF { get; set; }
        public string START_DATE { get; set; }
        public string GROSS_VALUE { get; set; }
        public string NET_VALUE { get; set; }
        public string TAX_VALUE { get; set; }
        public string SHIPMENT_VALUE { get; set; }
        public string POSTING_DATE { get; set; }
        public string END_DATE { get; set; }
        public string ONHOLD_DATE { get; set; }
        public string ONHOLD_USER { get; set; }
        public string CURRENT_STATUS { get; set; }
        public string SERVICE { get; set; }
        public string OUTSOURCING { get; set; }
        public string RENEWAL_TYPE { get; set; }
        public string VARIABLE_CONTRACT { get; set; }
        public string SALES_ORG { get; set; }
        public string SERVICE_ORG { get; set; }
        public string EMPLOYEE { get; set; }
        public string EMPLOYEE_ID { get; set; }
        public List<ItemsTBIn> ITEMS_TB { get; set; }
    }
    public class ContratosHoldParseOut
    {
        public long ID_CONTRACT { get; set; }
        public string DESC_CONTRACT { get; set; }
        public string CTY_CONTRACT { get; set; }
        public int ID_CUSTOMER { get; set; }
        public string DESC_CUSTOMER { get; set; }
        public string EXTERNAL_REF { get; set; }
        public string START_DATE { get; set; }
        public double GROSS_VALUE { get; set; }
        public double NET_VALUE { get; set; }
        public double TAX_VALUE { get; set; }
        public double SHIPMENT_VALUE { get; set; }
        public string POSTING_DATE { get; set; }
        public string END_DATE { get; set; }
        public string ONHOLD_DATE { get; set; }
        public string ONHOLD_USER { get; set; }
        public string CURRENT_STATUS { get; set; }
        public string SERVICE { get; set; }
        public string OUTSOURCING { get; set; }
        public string RENEWAL_TYPE { get; set; }
        public string VARIABLE_CONTRACT { get; set; }
        public string SALES_ORG { get; set; }
        public string SERVICE_ORG { get; set; }
        public string EMPLOYEE { get; set; }
        public string EMPLOYEE_ID { get; set; }
        public string SETTLEMENT_PERIOD { get; set; }
        public string BILLING_DATE { get; set; }
        public string BILL_CREATION_DATE { get; set; }
        public List<ItemsTBOut> ITEMS_TB { get; set; }
        public RequestResponse requestData { get; set; }
    }
    public class ItemsTBIn
    {
        public string SERVICE_NAME { get; set; }
        public string SERVICE_MG { get; set; }
        public string SERVICE_PRODUCT { get; set; }
        public string SERVICE_QTY { get; set; }
    }
    public class ItemsTBOut
    {
        public string SERVICE_NAME { get; set; }
        public string SERVICE_MG { get; set; }
        public string SERVICE_PRODUCT { get; set; }
        public double SERVICE_QTY { get; set; }
    }
    public class PostingRequest
    {
        public bool processOk { get; set; }
        public RequestResponse requestData { get; set; }
    }
    public class RequestResponse
    {
        public int status { get; set; }
        public bool success { get; set; }
        public RequestResponsePayload payload { get; set; }
    }
    public class RequestResponsePayload
    {
        public string message { get; set; }
        public List<RequestContract> requests { get; set; }
    }
    public class RequestContract
    {
        public int id { get; set; }
        public int id_Contract { get; set; }
        public string contractNumber { get; set; }
        public string newTargetDate { get; set; }
    }
    public class ContractDate
    {
        public List<int> contractsIDs { get; set; }
    }
    #endregion

}
