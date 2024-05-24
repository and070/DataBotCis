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
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Root;
using DataBotV5.Data.Stats;

namespace DataBotV5.Automation.RPA2.ContractsHold
{
    /// <summary>
    ///  Clase RPA Automation encargada de procesar los nuevos contratos on hold, actualizarlos o eliminarlos segun corresponda el criterio del estado e información almacenada
    /// Cada vez que se crea un contrato en S & S, una copia de la informacion es almacenada en 
    /// el Redis del servidor .72 con una llave en la nomenclatura gbm:msecoh:{mandante del main}:{id del contrato de SAP CRM}
    /// Esta llave se usa para comparar (utilizando el metodo compare) ciertos valores de los objetos para observar si cambian o no
    /// Si estos cambian, ejecuta una actualización. Si se detecta que un contrato que esta en Redis no aparece más en el lista de contratos extraidos en el RFC
    /// se asume que ya no se encuentra on Hold, por lo que este se procede a eliminar de Redis y a cambiar el estado en S & S, ejecutando una actualización
    /// </summary>
    class ContractsOnHold
    {
        ConsoleFormat console = new ConsoleFormat();
        SapVariants sap = new SapVariants();
        Rooting root = new Rooting();
        string respFinal = "";
        Log log = new Log();
        string mand = "CRM";
        public void Main()
        {
           
            ProcessAdmin padmin = new ProcessAdmin();

            //Se debe de establecer la frecuencia en el planificador (No se encuentra implementado en esta clase) dado que el proceso suele durar 1 - 3 minutos

            //todos los días a las 6:30 am

            console.WriteLine(" Starting process: ProcessOnHold");
            ProcessOnHold();


        }

        /// <summary>
        /// Método que se encarga de procesar los nuevos contratos on hold, actualizarlos o eliminarlos segun corresponda el criterio del estado e información almacenada
        /// Cada vez que se crea un contrato en S & S, una copia de la informacion es almacenada en 
        /// el Redis del servidor .72 con una llave en la nomenclatura gbm:msecoh:{mandante del main}:{id del contrato de SAP CRM}
        /// Esta llave se usa para comparar (utilizando el metodo compare) ciertos valores de los objetos para observar si cambian o no
        /// Si estos cambian, ejecuta una actualización. Si se detecta que un contrato que esta en Redis no aparece más en el lista de contratos extraidos en el RFC
        /// se asume que ya no se encuentra on Hold, por lo que este se procede a eliminar de Redis y a cambiar el estado en S & S, ejecutando una actualización
        /// </summary>
        /// <param name="mand">Mandante de SAP a modificar</param>
        private void ProcessOnHold()
        {
            //Lista de las llaves a extraer en Redis
            List<string> contractKeys = new List<string>();
            //Lista de los contratos extraidos en SAP
            List<ContratosHoldParseOut> msecohData = new List<ContratosHoldParseOut>();
            //Lista de todas las llaves relacionadas a contratos
            List<ContratosHoldParseOut> SSData = new List<ContratosHoldParseOut>();

            //Extrae todas las llaves que contengan el patron definido de Redis

            SSData = GetContractOnHold();


            //Verifica si existen llaves con el patron dentro de la lista
            if (SSData.Count > 0)
            {
                //Si existen llaves en Redis, saque primero los contratos on hold de SAP
                msecohData = GetContracts();
                SSData.ForEach(x =>
                {
                    //Busca dentro de la lista de contratos onhold, si el ID de Contrato de Redis existe
                    int indx = msecohData.FindIndex(y => y.ID_CONTRACT == x.ID_CONTRACT);
                    if (indx == -1)
                    {
                        //Si no existe es porque el contrato ya no esta On Hold, por lo tanto elimine la llave del Redis y cambie el estado dentro de S & S
                        DeleteContractAPI(x);
                        log.LogDeCambios("Eliminar", root.BDProcess, "Contratos", "Eliminar contrato que ya no esta On Hold", x.ID_CONTRACT.ToString(), "");
                        respFinal = respFinal + "\\n Eliminar contrato que ya no esta On Hold: " + x.ID_CONTRACT.ToString();

                    }
                    else
                    {
                        //Si existe el contrato de Redis dentro de los contratos extraidos de SAP on hold, entonces compare los campos claves para determinar si existen cambios
                        if (Compare(x, msecohData[indx]))
                        {
                            //Si existen cambios actualice la informacion en S & S y la llave de Redis
                            UpdateContractAPI(msecohData[indx]);
                            log.LogDeCambios("Actualización", root.BDProcess, "Contratos", "Actualizar contrato", x.ID_CONTRACT.ToString(), "");
                            respFinal = respFinal + "\\n Actualizar contrato: " + x.ID_CONTRACT.ToString();

                        }
                    }
                });

                //Una vez comparada la data y actualizada, en caso de que aplicase, busca si existen nuevos contratos on hold, para ello
                //Es necesario eliminar todos aquellos contratos extraidos en SAP que ya se encuentran en Redis
                SSData.ForEach(x =>
                {
                    int indx = msecohData.FindIndex(y => y.ID_CONTRACT == x.ID_CONTRACT);
                    if (indx != -1)
                    {
                        msecohData.RemoveAt(indx);
                    }
                });
                //Si aun existen contratos, entonces es porque son contratos nuevos
                if (msecohData.Count > 0)
                {
                    //Por cada nuevo contrato, se crea una llave de Redis e inserta la información en S & S
                    msecohData.ForEach(x =>
                    {
                        string con = x.ID_CONTRACT.ToString();

                        CreateContractAPI(x);
                        //}
                        log.LogDeCambios("Creación", root.BDProcess, "Contratos", "Crear contrato en S&S", x.ID_CONTRACT.ToString(), "");
                        respFinal = respFinal + "\\n Crear contrato en S&S: " + x.ID_CONTRACT.ToString();

                    });
                }


                if (msecohData.Count > 0)
                {
                    root.BDUserCreatedBy = "RSAGUER";
                    root.requestDetails = respFinal;
                    console.WriteLine("Creando estadísticas...");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }

            }
            else
            {
                //Sino existen llaves del todo, entonces obtenga los contraos on Hold de SAP, cree las llaves y de paso, inserte la informacion en S & S

                //Obtiene los contratos on hold, del mandate deseado
                msecohData = GetContracts();
                //Para cada contrato, crea la llave en Redis e inserta la información en S & S
                msecohData.ForEach(x =>
                {
                    CreateContractAPI(x);
                    log.LogDeCambios("Creación", root.BDProcess, "Contratos", "Crear contrato en S&S", x.ID_CONTRACT.ToString(), "");
                    respFinal = respFinal + "\\n Crear contrato en S&S: " + x.ID_CONTRACT.ToString();

                });

                if (msecohData.Count > 0)
                {
                    root.BDUserCreatedBy = "RSAGUER";
                    root.requestDetails = respFinal;

                    console.WriteLine("Creando estadísticas...");
                    using (Stats stats = new Stats())
                    {
                        stats.CreateStat();
                    }
                }


            }


        }


        #region Métodos del API

        /// <summary>
        /// Extrae los contratos on Hold de la DB de Smart&Simple mediante una API
        /// </summary>
        /// <returns>una lista de contratos on hold en S&S</returns>
        private List<ContratosHoldParseOut> GetContractOnHold()
        {
            List<ContratosHoldParseOut> infoList = new List<ContratosHoldParseOut>();

            try
            {
                Credentials cred = new Credentials();
                //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);
                string url = cred.MESCOH_GET_CONTRACT_ONHOLD;

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "GET";
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    dynamic onHoldInfo = JObject.Parse(result);
                    dynamic contracts = onHoldInfo.payload.contracts;
                    foreach (var item in contracts)
                    {
                        ContratosHoldParseOut info = new ContratosHoldParseOut();
                        if (item.id == 32)
                        {
                            string s = "";
                        }
                        info.requestData = JsonConvert.DeserializeObject<RequestResponse>(result);
                        info.ID_CONTRACT = item.contractNumber;
                        info.DESC_CONTRACT = item.description;
                        info.CTY_CONTRACT = item.country;
                        info.ID_CUSTOMER = item.customerID;
                        info.DESC_CUSTOMER = item.customerName;
                        info.EXTERNAL_REF = item.externalReference;
                        DateTime sDate = item.startDate;
                        info.START_DATE = sDate.ToString("yyyy-MM-dd");
                        info.GROSS_VALUE = item.grossValue;
                        info.NET_VALUE = item.netValue;
                        info.TAX_VALUE = item.taxValue;
                        info.SHIPMENT_VALUE = item.shipmentValue;
                        try
                        {
                            DateTime pDate = item.postingDate;
                            info.POSTING_DATE = pDate.ToString("yyyy-MM-dd");
                        }
                        catch (Exception)
                        {
                            info.POSTING_DATE = "";
                        }

                        try
                        {
                            DateTime eDate = item.endDate;
                            info.END_DATE = eDate.ToString("yyyy-MM-dd");
                        }
                        catch (Exception)
                        {
                            info.END_DATE = "";
                        }

                        try
                        {
                            DateTime oDate = item.onHoldDate;
                            info.ONHOLD_DATE = oDate.ToString("yyyy-MM-dd");
                        }
                        catch (Exception)
                        {
                            info.ONHOLD_DATE = "";
                        }



                        info.ONHOLD_USER = item.onHoldUser;
                        info.CURRENT_STATUS = item.currentStatus;
                        info.SERVICE = item.service;
                        bool outSourcing = Convert.ToBoolean(item.outsourcing);
                        info.OUTSOURCING = outSourcing.ToString().ToLower();
                        info.RENEWAL_TYPE = item.renewalType;
                        bool vContract = Convert.ToBoolean(item.variableContract);
                        info.VARIABLE_CONTRACT = vContract.ToString().ToLower();
                        info.SALES_ORG = item.salesOrganization;
                        info.SERVICE_ORG = item.servicesOrganization;
                        info.EMPLOYEE = item.employeeName;
                        info.EMPLOYEE_ID = item.employeeID;
                        //Nuevo campos--------------------------------------
                        info.SETTLEMENT_PERIOD = item.settlementPeriod;
                        info.BILLING_DATE = item.billingDate;
                        info.BILL_CREATION_DATE = item.billingCreationDate;
                        //--------------------------------------------------
                        List<ItemsTBOut> listado = new List<ItemsTBOut>();
                        foreach (var service in item.items)
                        {
                            ItemsTBOut ser = new ItemsTBOut();
                            ser.SERVICE_NAME = service.name;
                            ser.SERVICE_MG = "";
                            ser.SERVICE_PRODUCT = "";
                            ser.SERVICE_QTY = 1;
                            listado.Add(ser);
                        }
                        info.ITEMS_TB = listado;
                        infoList.Add(info);
                    }


                }

            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message.ToString());
            }


            return infoList;
        }

        /// <summary>
        /// Método para la creación de un nuevo contrato on Hold en las BDs de S & S.
        /// </summary>
        /// <param name="info">Objeto con los datos parseados provenientes de SAP</param>
        private void CreateContractAPI(ContratosHoldParseOut info)
        {
            try
            {
                Credentials cred = new Credentials();
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                string url = cred.MESCOH_CREATE_CONTRACT;
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    string json = JsonConvert.SerializeObject(info);

                    streamWriter.Write(json);
                }
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    console.WriteLine(result);
                    //valid = bool.Parse(result);
                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message.ToString());
            }
        }
        /// <summary>
        /// Método para actualizacion de un contrato existente en las bases de datos de S & S.
        /// </summary>
        /// <param name="info">Objeto con los datos parseados provenientes de SAP</param>
        private void UpdateContractAPI(ContratosHoldParseOut info)
        {
            try
            {
                Credentials cred = new Credentials();
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                string url = $"{cred.MESCOH_UPDATE_CONTRACT}{info.ID_CONTRACT}";
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "PUT";
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    string json = JsonConvert.SerializeObject(info);

                    streamWriter.Write(json);
                }
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    console.WriteLine(result);
                    //valid = bool.Parse(result);
                }

            }
            catch (Exception ex)
            {
                console.WriteLine(ex.Message.ToString());
            }
        }
        /// <summary>
        /// Método de actualización y borrado lógico al cambiar el estado de un contrato en S&S cuando este se detecta que el estado ya no se encuentra On Hold (Aplica como un update).
        /// </summary>
        /// <param name="info">Objeto con los datos parseados provenientes de SAP</param>
        private void DeleteContractAPI(ContratosHoldParseOut info)
        {
            info.CURRENT_STATUS = "In Progress";
            try
            {
                Credentials cred = new Credentials();
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                string url = $"{cred.MESCOH_UPDATE_CONTRACT}{info.ID_CONTRACT}";
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "PUT";
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    string json = JsonConvert.SerializeObject(info);

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
        #endregion

        #region Metodos de validación y utilidad
        public bool AcceptAllCertifications(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }
        /// <summary>
        /// Compara dos objetos para determinar si existen cambios.
        /// </summary>
        /// <param name="one">Objeto #1</param>
        /// <param name="two">Objeto #2</param>
        /// <returns>Retorna true, si existen cambios, false si no existen cambios.</returns>
        private bool Compare(ContratosHoldParseOut one, ContratosHoldParseOut two)
        {
            bool result = false;
            if (one.DESC_CONTRACT != two.DESC_CONTRACT)
            {
                result = true;
            }
            if (one.EMPLOYEE_ID != two.EMPLOYEE_ID)
            {
                result = true;
            }
            if (one.EXTERNAL_REF != two.EXTERNAL_REF)
            {
                result = true;
            }
            if (one.ITEMS_TB.Count != two.ITEMS_TB.Count)
            {
                result = true;
            }
            if (one.NET_VALUE != two.NET_VALUE)
            {
                result = true;
            }
            if (one.ONHOLD_DATE != two.ONHOLD_DATE)
            {
                result = true;
            }
            if (one.ONHOLD_USER != two.ONHOLD_USER)
            {
                result = true;
            }
            if (one.OUTSOURCING != two.OUTSOURCING)
            {
                result = true;
            }
            if (one.RENEWAL_TYPE != two.RENEWAL_TYPE)
            {
                result = true;
            }
            if (one.START_DATE != two.START_DATE)
            {
                result = true;
            }
            if (one.POSTING_DATE != two.POSTING_DATE)
            {
                result = true;
            }
            if (one.END_DATE != two.END_DATE)
            {
                result = true;
            }
            //Nuevo campos--------------------------------------
            if (one.SETTLEMENT_PERIOD != two.SETTLEMENT_PERIOD)
            {
                result = true;
            }
            if (one.BILLING_DATE != two.BILLING_DATE)
            {
                result = true;
            }
            if (one.BILL_CREATION_DATE != two.BILL_CREATION_DATE)
            {
                result = true;
            }
            //--------------------------------------------------
            return result;
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

            //Nuevo campos--------------------------------------
            parseOut.SETTLEMENT_PERIOD = valores.GetString("SETTLEMENT_PERIOD");
            parseOut.BILLING_DATE = valores.GetString("BILLING_DATE");
            parseOut.BILL_CREATION_DATE = valores.GetString("BILL_CREATION_DATE");
            //--------------------------------------------------

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
        private void ProcessDateChangeRFC(string id, string type, string date, string mand)
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
        /// <summary>
        /// Método que extrae los contratos on Hold, por estatus ( personalizable )
        /// </summary>
        /// <param name="mand">Mandante de SAP a modificar</param>
        /// <returns>Retorna una lista de objetos parseados para su uso en el bot y S&S</returns>
        private List<ContratosHoldParseOut> GetContracts()
        {
            List<ContratosHoldParseOut> contratos = new List<ContratosHoldParseOut>();
            Credentials cred = new Credentials();

            Dictionary<string, string> parametros = new Dictionary<string, string>();
            parametros["CONTRACT_TYPES"] = "803";
            parametros["ONHOLD_STATUS"] = "E0011";

            IRfcFunction getContractsCRM = sap.ExecuteRFC(mand, "ZGET_ALL_ONHOLD", parametros);

            IRfcTable resultado = getContractsCRM.GetTable("RESPONSE");
            foreach (IRfcStructure item in resultado)
            {

                ContratosHoldParseOut con = ParseOut(item);
                contratos.Add(con);

            }

            return contratos;
        }
        #endregion
    }

}
