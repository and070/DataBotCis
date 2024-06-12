using System.Collections.Specialized;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using DataBotV5.Data.Root;
using System.Data;
using System.Linq;
using System.Text;
using System.Net;
using System.Xml;
using System.IO;
using System;
using DataBotV5.Automation.ICS.ControlDesk;
using Microsoft.VisualStudio.OLE.Interop;

namespace DataBotV5.Logical.Projects.ControlDesk
{
    /// <summary>
    /// Clase Logical encargada de interacción con Control Desk.
    /// </summary>
    class ControlDeskInteraction
    {
        readonly Credentials cred = new Credentials();
        readonly Rooting root = new Rooting();

        const string apiRest = "/maxrest/rest/os/";
        const string osService = "/meaweb/os/";

        private bool ValidateXml(string xml)
        {
            XmlDocument xmlDoc = new XmlDocument();

            try
            {
                xmlDoc.LoadXml(xml);
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }
        public string PostCD(string destinationUrl, string objStructure, string requestXml)
        {
            string user = cred.username_CD;
            string pass = cred.password_CD;

            destinationUrl = destinationUrl + osService + objStructure;
            byte[] bytes = Encoding.UTF8.GetBytes(requestXml);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationUrl);
            request.ContentType = "text/xml; encoding='utf-8'";

            request.ContentLength = bytes.Length;
            request.Method = "POST";

            NetworkCredential myCred = new NetworkCredential(user, pass);
            CredentialCache myCache = new CredentialCache();
            myCache.Add(new Uri(destinationUrl), "Basic", myCred);
            WebRequest wr = WebRequest.Create(destinationUrl);
            wr.Credentials = myCache;
            request.Credentials = wr.Credentials;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = response.GetResponseStream();
                    string responseStr = new StreamReader(responseStream).ReadToEnd();
                    return responseStr;
                }
                return null;
            }
            catch (WebException ex)
            {
                try
                {
                    Stream data = ex.Response.GetResponseStream();
                    StreamReader reader = new StreamReader(data);
                    string errorMessage = reader.ReadToEnd();
                    if (errorMessage == "")
                        return ex.Message;
                    else
                        return errorMessage;
                }
                catch (Exception e)
                {
                    return e.Message;
                }
            }

        }
        private string CallMaxRestPost(string destinationUrl, string objStructure, string uniqueId, Dictionary<string, string> objectFields)
        {
            string user = cred.username_CD;
            string pass = cred.password_CD;
            string ret = "ERROR";

            destinationUrl = destinationUrl + apiRest + objStructure + "/" + uniqueId; ;

            if (objectFields.Count > 0)
            {
                try
                {
                    NameValueCollection postData = new NameValueCollection();

                    foreach (KeyValuePair<string, string> objectField in objectFields)
                        postData.Add(objectField.Key.ToString(), objectField.Value.ToString());

                    WebClient client = new WebClient();
                    string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(user + ":" + pass));

                    client.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                    client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";

                    byte[] responseBytes = client.UploadValues(destinationUrl, "POST", postData);
                    string responseText = Encoding.UTF8.GetString(responseBytes);

                    XmlDocument outXml = new XmlDocument();
                    outXml.LoadXml(responseText);

                    foreach (KeyValuePair<string, string> objectField in objectFields)
                    {
                        string tagName = objectField.Key.ToString();
                        XmlNode currentField = outXml.GetElementsByTagName(tagName)[0];
                        if (objectField.Value.ToString() == currentField.InnerText)
                            ret = "OK";
                        else
                        {
                            ret = "ERROR: no se pudo cambiar el campo: " + tagName;
                            break;
                        }
                    }

                }
                catch (WebException ex)
                {
                    // Leer la respuesta del servidor en caso de error
                    if (ex.Response != null)
                    {
                        using (HttpWebResponse errorResponse = (HttpWebResponse)ex.Response)
                        using (StreamReader reader = new StreamReader(errorResponse.GetResponseStream()))
                        {
                            ret = reader.ReadToEnd();
                        }
                    }
                    else
                        ret = "ERROR: " + ex.Message;
                }
            }
            return ret;
        }
        public string CallMaxRestGet(string destinationUrl, string user, string pass)
        {
            string ret = "";

            try
            {
                WebClient client = new WebClient();

                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(user + ":" + pass));
                client.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;

                byte[] responseData = client.DownloadData(destinationUrl);
                string response = Encoding.UTF8.GetString(responseData);

                ret = response;
            }
            catch (WebException ex)
            {
                // Leer la respuesta del servidor en caso de error
                if (ex.Response != null)
                {
                    using (HttpWebResponse errorResponse = (HttpWebResponse)ex.Response)
                    using (StreamReader reader = new StreamReader(errorResponse.GetResponseStream()))
                    {
                        ret = reader.ReadToEnd();
                    }
                }
                else
                {
                    ret = "Error: " + ex.Message;
                }
            }
            return ret;
        }
        private string ChangeCdStatus(string objStruct, string uniqueId, string newStatus)
        {

            string user = cred.username_CD;
            string pass = cred.password_CD;
            string ret;
            string url = "http://controldesk-dev.gbm.net/maxrest/rest/os/" + objStruct + "/" + uniqueId;

            try
            {
                WebClient client = new WebClient();
                NameValueCollection postData = new NameValueCollection() { { "STATUS", newStatus } };
                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(user + ":" + pass));

                client.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";

                byte[] responseBytes = client.UploadValues(url, "POST", postData);
                string responseText = Encoding.UTF8.GetString(responseBytes);

                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);
                XmlNode currentStatus = outXml.GetElementsByTagName("STATUS")[0];
                if (newStatus == currentStatus.InnerText)
                    ret = "OK";
                else
                    ret = "ERROR";
            }
            catch (WebException ex)
            {
                // Leer la respuesta del servidor en caso de error
                if (ex.Response != null)
                {
                    using (HttpWebResponse errorResponse = (HttpWebResponse)ex.Response)
                    using (StreamReader reader = new StreamReader(errorResponse.GetResponseStream()))
                    {
                        ret = reader.ReadToEnd();
                    }
                }
                else
                    ret = "ERROR: " + ex.Message;
            }

            return ret;
        }
        public List<CdContractData> GetAllConRevisions(string user, string pass, string contract)
        {

            string url = "http://controldesk-dev.gbm.net/maxrest/rest/os/smarin1?AGREEMENT=" + contract;

            string response = CallMaxRestGet(url, user, pass);
            return ParseContractXml(response);
        }

        #region Métodos de consulta

        #region Collections
        public bool CheckCollectionExistence(string colNum)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXCOLLECTIONICS xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXCOLLECTIONICSQuery>" +
            "<WHERE>COLLECTIONNUM = '" + colNum.ToUpper() + "'</WHERE>" +
            "</MXCOLLECTIONICSQuery>" +
            "</QueryMXCOLLECTIONICS>";

            string idRes = "";

            #region Process Response
            string responseText = PostCD(root.UrlCd, "MXCOLLECTIONICS", xml);

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                try { idRes = outXml.GetElementsByTagName("COLLECTIONNUM")[0].InnerText; } catch (Exception) { }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error  
            #endregion

            if (idRes == colNum.ToUpper())
                return true;
            else
                return false;
        }
        public CdCollectionData GetCollectionData(string colNum)
        {
            CdCollectionData collectionInfo = new CdCollectionData();
            collectionInfo.CollectionParties = new List<CdCollectionPartyData>();

            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXCOLLECTIONICS xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXCOLLECTIONICSQuery>" +
            "<WHERE>COLLECTIONNUM = '" + colNum.ToUpper() + "'</WHERE>" +
            "</MXCOLLECTIONICSQuery>" +
            "</QueryMXCOLLECTIONICS>";

            string resXml = PostCD(root.UrlCd, "MXCOLLECTIONICS", xml);

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(resXml);
                XmlNodeList collections = outXml.GetElementsByTagName("COLLECTION");
                foreach (XmlNode collection in collections)
                {
                    foreach (XmlElement collectionFields in collection)
                    {
                        if (collectionFields.Name == "COLLECTIONNUM") { collectionInfo.CollectionNum = collectionFields.InnerText; }
                        if (collectionFields.Name == "COLLSUPERVISOR") { collectionInfo.Supervisor = collectionFields.InnerText; }
                        if (collectionFields.Name == "DESCRIPTION") { collectionInfo.Description = collectionFields.InnerText; }
                        if (collectionFields.Name == "GBMCOLLECTIONINTPARTY")
                        {
                            CdCollectionPartyData partyInfo = new CdCollectionPartyData();
                            XmlNode collectionIntParty = collectionFields;
                            foreach (XmlElement party in collectionIntParty)
                            {
                                if (party.Name == "DESCRIPTION") { partyInfo.Description = party.InnerText; }
                                if (party.Name == "GBMCOLLECTIONINTPARTYID") { partyInfo.Id = party.InnerText; }
                                if (party.Name == "PERSONGROUP") { partyInfo.PersonGroup = party.InnerText; }
                                if (party.Name == "PERSONID") { partyInfo.PersonId = party.InnerText; }
                                if (party.Name == "TYPE") { partyInfo.Type = party.InnerText; }
                            }
                            collectionInfo.CollectionParties.Add(partyInfo);
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine(resXml);
            }

            return collectionInfo;
        }
        #endregion
        #region People
        public string CheckPersonExistence(string email)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXPERSON xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXPERSONQuery>" +
            "<WHERE>exists (select 1 from maximo.email where ((upper(emailaddress) = &apos;" + email.ToUpper() + "&apos;)) and (personid=person.personid))</WHERE>" +
            "</MXPERSONQuery>" +
            "</QueryMXPERSON>";

            string responseText = PostCD(root.UrlCd, "MXPERSON", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (responseText.Contains("PERSONID"))
                {
                    try
                    {
                        responseText = outXml.GetElementsByTagName("PERSONID")[0].InnerText;
                    }
                    catch (Exception ex)
                    {
                        responseText = ex.Message;
                    }
                }
                else
                {
                    responseText = "NE";
                }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return responseText;
        }
        #endregion
        #region Person Groups
        public string CheckPersonGroupExistence(string personGroup)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXPERSONGROUP xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXPERSONGROUPQuery>" +
            "<WHERE>PERSONGROUP ='" + personGroup.ToUpper() + "'</WHERE>" +
            "</MXPERSONGROUPQuery>" +
            "</QueryMXPERSONGROUP>";

            string responseText = PostCD(root.UrlCd, "MXPERSONGROUP", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (responseText.Contains("<PERSONGROUP"))
                {
                    try
                    {
                        //responseText = outXml.GetElementsByTagName("PERSONGROUP")[0].InnerText;
                        responseText = "OK";

                    }
                    catch (Exception ex)
                    {
                        responseText = ex.Message;
                    }
                }
                else
                {
                    responseText = "NE";
                }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return responseText;
        }
        public string[] GetPersonGroupPeople(string personGroup)
        {
            string[] ret = null;
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXPERSONGROUP xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"900\">" +
            "<MXPERSONGROUPQuery>" +
            "<WHERE>PERSONGROUP ='" + personGroup.ToUpper() + "'</WHERE>" +
            "</MXPERSONGROUPQuery>" +
            "</QueryMXPERSONGROUP>";

            string responseText = PostCD(root.UrlCd, "MXPERSONGROUP", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (responseText.Contains("<PERSONGROUP"))
                {
                    try
                    {
                        XmlNodeList respParties = outXml.GetElementsByTagName("RESPPARTY");
                        List<string> mail = new List<string>();
                        foreach (XmlNode respParty in respParties)
                            mail.Add(respParty.InnerText.Trim());
                        mail = mail.Distinct().ToList();
                        ret = mail.ToArray();
                    }
                    catch (Exception) { }
                }
            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return ret;
        }
        #endregion
        #region Users
        public bool CheckUserExistence(string email)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXPERUSER xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXPERUSERQuery>" +
            "<WHERE>PERSONID = '" + email.ToUpper() + "'</WHERE>" +
            "</MXPERUSERQuery>" +
            "</QueryMXPERUSER>";

            string resId = "";

            #region Process Response
            string responseText = PostCD(root.UrlCd, "MXPERUSER", xml);

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                try { resId = outXml.GetElementsByTagName("PERSONID")[0].InnerText; } catch (Exception) { }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error  
            #endregion

            if (resId == email.ToUpper())
                return true;
            else
                return false;
        }
        #endregion
        #region Configuration Items
        public string GetCisXml(List<string> cIs)
        {
            string joinedCIs = string.Join(",", cIs.Select(v => $"'{v}'"));

            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<QueryMXAUTHCI xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""900"">";
            xml += "<MXAUTHCIQuery><WHERE>CINUM IN (" + joinedCIs + ")</WHERE>";
            xml += "</MXAUTHCIQuery>";
            xml += "</QueryMXAUTHCI>";

            string infoXml = PostCD(root.UrlCd, "MXAUTHCI", xml);

            return infoXml;
        }
        public List<CdConfigurationItemData> ParseCiXml(string cIXml)
        {
            List<CdConfigurationItemData> ret = new List<CdConfigurationItemData>();

            XmlDocument outXml = new XmlDocument();
            outXml.LoadXml(cIXml);
            XmlNodeList cis = outXml.GetElementsByTagName("CI");
            foreach (XmlNode ci in cis)
            {
                CdConfigurationItemData ciInfo = new CdConfigurationItemData();

                foreach (XmlElement ciFields in ci)
                {
                    if (ciFields.Name == "CINAME") { ciInfo.CiName = ciFields.InnerText; }
                    if (ciFields.Name == "CINUM") { ciInfo.CiNum = ciFields.InnerText; }
                    if (ciFields.Name == "PERSONID") { ciInfo.PersonId = ciFields.InnerText; }

                    //Por ahora solo me interesan estos datos, podrían ser mas
                }
                ret.Add(ciInfo);
            }

            return ret;
        }
        #endregion
        #region Customer Agreements
        public List<CdContractData> ParseContractXml(string contractsXml)
        {
            List<CdContractData> ret = new List<CdContractData>();

            if (ValidateXml(contractsXml))
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(contractsXml);
                XmlNodeList contracts = outXml.GetElementsByTagName("PLUSPAGREEMENT");
                foreach (XmlNode contract in contracts)
                {
                    CdContractData conInfo = new CdContractData();
                    conInfo.MaterialArray = new List<string>();
                    conInfo.ManualServiceArray = new List<string>();
                    conInfo.EquipArray = new List<string>();
                    conInfo.CisArray = new List<string>();
                    conInfo.PriceSchedules = new List<CdPriceScheduleData>();

                    foreach (XmlElement contractFields in contract)
                    {
                        if (contractFields.Name == "STATUS") { conInfo.Status = contractFields.InnerText; }
                        if (contractFields.Name == "REVISIONNUM") { conInfo.Revision = contractFields.InnerText; }
                        if (contractFields.Name == "CUSTOMER") { conInfo.IdCustomer = contractFields.InnerText; }
                        if (contractFields.Name == "DESCRIPTION") { conInfo.Description = contractFields.InnerText; }
                        if (contractFields.Name == "STARTDATE") { conInfo.StartDate = contractFields.InnerText; }
                        if (contractFields.Name == "ENDDATE") { conInfo.EndDate = contractFields.InnerText; }
                        if (contractFields.Name == "AGREEMENT") { conInfo.IdContract = contractFields.InnerText; }
                        if (contractFields.Name == "PLUSPAGREEMENTID") { conInfo.PluspAgreementId = contractFields.InnerText; }
                        if (contractFields.Name == "PLUSPPRICESCHED")
                        {
                            CdPriceScheduleData cdPriceScheduleData = new CdPriceScheduleData();

                            XmlNode priceSchedules = contractFields;
                            foreach (XmlElement priceSchedulesFields in priceSchedules)
                            {
                                if (priceSchedulesFields.Name == "SANUM")
                                {
                                    cdPriceScheduleData.SaNum = priceSchedulesFields.InnerText;
                                }
                                if (priceSchedulesFields.Name == "PRICESCHEDULE")
                                {
                                    cdPriceScheduleData.PriceSchedule = priceSchedulesFields.InnerText;
                                }
                                if (priceSchedulesFields.Name == "PLUSPAPPLSERV")
                                {
                                    XmlNode services = priceSchedulesFields;
                                    string commodity = "";

                                    foreach (XmlElement servicesFields in services)
                                    {
                                        if (servicesFields.Name == "COMMODITY")
                                        {
                                            commodity = servicesFields.InnerText;
                                            conInfo.MaterialArray.Add(commodity);
                                        }
                                    }
                                    foreach (XmlElement servicesFields in services)
                                    {
                                        if (servicesFields.Name == "MANUAL")
                                        {
                                            if (servicesFields.InnerText == "1")
                                            {
                                                conInfo.ManualServiceArray.Add(commodity);
                                            }
                                        }
                                    }

                                }
                                if (priceSchedulesFields.Name == "PLUSPAPPLASSET")
                                {
                                    XmlNode assets = priceSchedulesFields;
                                    foreach (XmlElement assetsFields in assets)
                                    {
                                        if (assetsFields.Name == "ASSETNUM")
                                        {
                                            conInfo.EquipArray.Add(assetsFields.InnerText);
                                        }
                                    }
                                }
                                if (priceSchedulesFields.Name == "PLUSPAPPLCI")
                                {
                                    XmlNode Cis = priceSchedulesFields;
                                    foreach (XmlElement CisFields in Cis)
                                    {
                                        if (CisFields.Name == "CINUM")
                                        {
                                            conInfo.CisArray.Add(CisFields.InnerText);
                                        }
                                    }
                                }
                            }
                            conInfo.PriceSchedules.Add(cdPriceScheduleData);
                        }
                    }

                    conInfo.MaterialArray = conInfo.MaterialArray.Distinct().ToList();
                    conInfo.EquipArray = conInfo.EquipArray.Distinct().ToList();
                    conInfo.CisArray = conInfo.CisArray.Distinct().ToList();

                    ret.Add(conInfo);
                }

            }
            return ret;
        }
        public string GetContractRevision(string contract)
        {
            #region XML(MXCUSTAGREEMENT)
            //string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            //    @"<QueryMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""901"">" +
            //    "<MXCUSTAGREEMENTQuery>" +
            //    "<WHERE>agreement = '" + con + "' and status IN ('APPR','WSTART')</WHERE>" +
            //    "</MXCUSTAGREEMENTQuery>" +
            //    "</QueryMXCUSTAGREEMENT>"; 
            #endregion

            #region XML(MXCUSTAGREEMENTEXIST) custom object
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                @"<QueryMXCUSTAGREEMENTEXIST xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""901"">" +
                "<MXCUSTAGREEMENTEXISTQuery>" +
                "<WHERE>agreement = '" + contract + "' and status IN ('APPR','WSTART')</WHERE>" +
                "</MXCUSTAGREEMENTEXISTQuery>" +
                "</QueryMXCUSTAGREEMENTEXIST>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXCUSTAGREEMENTEXIST", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (responseText.Contains("REVISIONNUM"))
                {
                    try
                    {
                        responseText = outXml.GetElementsByTagName("REVISIONNUM")[0].InnerText;
                    }
                    catch (Exception ex)
                    {
                        responseText = ex.Message;
                    }
                }
                else
                {
                    responseText = "-1";
                }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return responseText;
        }
        public string GetContractStatus(string idContract)
        {
            IDictionary<string, string> list = new Dictionary<string, string>();
            string responseText, status = "", rev = "";

            #region XML(MXCUSTAGREEMENT)
            //string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            //                 @"<QueryMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""900"">" +
            //                 "<MXCUSTAGREEMENTQuery>" +
            //                 "<WHERE>agreement = '" + idcontrato + "' </WHERE>" +
            //                 "</MXCUSTAGREEMENTQuery>" +
            //                 "</QueryMXCUSTAGREEMENT>";
            #endregion

            #region XML(MXCUSTAGREEMENTEXIST) custom object
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                             @"<QueryMXCUSTAGREEMENTEXIST xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""900"">" +
                             "<MXCUSTAGREEMENTEXISTQuery>" +
                             "<WHERE>agreement like ('%" + idContract + "%') </WHERE>" +
                             "</MXCUSTAGREEMENTEXISTQuery>" +
                             "</QueryMXCUSTAGREEMENTEXIST>";
            #endregion

            responseText = PostCD(root.UrlCd, "MXCUSTAGREEMENTEXIST", xml);

            #region existe o no
            string[] stringSeparators0 = new string[] { @"rsTotal=""" };
            string[] sp = responseText.Split(stringSeparators0, StringSplitOptions.None);
            if (sp.Length > 1)
            {
                int? count = int.Parse(sp[1]?.ToString().Substring(0, 1));
                if (count != null)
                {
                    if (count > 0)
                    {
                        #region saca el status
                        try
                        {
                            XmlDocument outXml = new XmlDocument();
                            outXml.LoadXml(responseText);
                            XmlNodeList cons = outXml.GetElementsByTagName("PLUSPAGREEMENT");//todos los contratos
                            foreach (XmlNode contrato in cons)
                            {
                                foreach (XmlElement campos in contrato)
                                {
                                    if (campos.Name == "STATUS") { status = campos.InnerText; }
                                    if (campos.Name == "REVISIONNUM") { rev = campos.InnerText; }
                                }
                                list.Add(rev, status);
                            }
                            if (list.Count > 0)
                                status = list[(list.Count - 1).ToString()];

                        }
                        catch (Exception)
                        {
                            status = "NE";
                        }
                        #endregion
                    }
                    else
                        status = "NE";
                }
            }
            else
                status = "NE";
            #endregion

            return status;
        }
        public string GetContractsXml(List<string> contracts)
        {
            string infoXml = "";
            if (contracts.Count > 0)
            {
                string joinedCons = string.Join(",", contracts.Select(v => $"'{v}'"));

                string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
                xml += @"<QueryMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""900"">";
                xml += "<MXCUSTAGREEMENTQuery>";
                xml += "<WHERE>agreement IN (" + joinedCons + ") and status = 'APPR'</WHERE>";
                xml += "</MXCUSTAGREEMENTQuery>";
                xml += "</QueryMXCUSTAGREEMENT>";

                infoXml = PostCD(root.UrlCd, "MXCUSTAGREEMENT", xml);
            }

            return infoXml;
        }

        public string GetContractRevXml(string contract, string revision)
        {
            string infoXml = "";

            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<QueryMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""900"">";
            xml += "<MXCUSTAGREEMENTQuery>";
            xml += "<WHERE>agreement = " + contract + " and revisionnum = " + revision + "</WHERE>";
            xml += "</MXCUSTAGREEMENTQuery>";
            xml += "</QueryMXCUSTAGREEMENT>";

            infoXml = PostCD(root.UrlCd, "MXCUSTAGREEMENT", xml);

            return infoXml;
        }

        #endregion
        #region Response Plans
        public string GetResponsePlanStatus(string idRp)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                @"<QueryMXRESPONSEPLANICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""1"">" +
                "<MXRESPONSEPLANICSQuery>" +
                "<WHERE>SANUM = '" + idRp + "'</WHERE>" +
                "</MXRESPONSEPLANICSQuery>" +
                "</QueryMXRESPONSEPLANICS>";

            string responseText = PostCD(root.UrlCd, "MXRESPONSEPLANICS", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (responseText.Contains("STATUS"))
                {
                    try
                    {
                        responseText = outXml.GetElementsByTagName("STATUS")[0].InnerText;
                    }
                    catch (Exception ex)
                    {
                        responseText = ex.Message;
                    }
                }
                else
                {
                    responseText = "NE";
                }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return responseText;
        }
        public bool CheckResponsePlansExistence(string saNum)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXRESPONSEPLANICS xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXRESPONSEPLANICSQuery>" +
            "<WHERE>SANUM = '" + saNum + "'</WHERE>" +
            "</MXRESPONSEPLANICSQuery>" +
            "</QueryMXRESPONSEPLANICS>";

            string resId = "";

            #region Process Response
            string responseText = PostCD(root.UrlCd, "MXRESPONSEPLANICS", xml);

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                try { resId = outXml.GetElementsByTagName("SANUM")[0].InnerText; } catch (Exception) { }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error  
            #endregion

            if (resId == saNum.ToUpper())
                return true;
            else
                return false;
        }
        public string GetResponsePlanXml(string singleRpId)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
            "<QueryMXRESPONSEPLANICS xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\" maxItems=\"1\">" +
            "<MXRESPONSEPLANICSQuery>" +
            "<WHERE>SANUM = '" + singleRpId + "'</WHERE>" +
            "</MXRESPONSEPLANICSQuery>" +
            "</QueryMXRESPONSEPLANICS>";

            string responseText = PostCD(root.UrlCd, "MXRESPONSEPLANICS", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (!responseText.Contains("PLUSPRESPPLAN"))
                    responseText = "NE";

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return responseText;
        }
        public CdResponsePlanData ParseXmlToResponsePlanData(string singleConXml)
        {
            CdResponsePlanData rpInfo = new CdResponsePlanData();
            List<CdServicesData> services = new List<CdServicesData>();
            CdClassStructureData classStructure = new CdClassStructureData();

            rpInfo.ConfigurationItems = new List<string>();

            XmlDocument outXml = new XmlDocument();
            outXml.LoadXml(singleConXml);
            XmlNodeList responsePlans = outXml.GetElementsByTagName("PLUSPRESPPLAN");
            foreach (XmlNode responsePlan in responsePlans)
            {
                foreach (XmlElement responsePlanFields in responsePlan)
                {
                    if (responsePlanFields.Name == "ASSIGNOWNERGROUP") { rpInfo.OwnerGroup = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "CALENDAR") { rpInfo.Calendar = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "CLASSSTRUCTUREID") { rpInfo.ClassStructureId = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "CONDITION") { rpInfo.Condition = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "DESCRIPTION") { rpInfo.Description = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "ESCALATION") { rpInfo.Escalation = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "OBJECTNAME") { rpInfo.ObjectName = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "PLUSPSERVAGREEID") { rpInfo.PluspServAgreeId = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "RANKING") { rpInfo.Ranking = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "SANUM") { rpInfo.Sanum = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "SHIFT") { rpInfo.Shift = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "STATUS") { rpInfo.Status = responsePlanFields.InnerText; }
                    if (responsePlanFields.Name == "PLUSPAPPLSERV")
                    {
                        XmlNode pluspAppServs = responsePlanFields;
                        CdServicesData service = new CdServicesData();
                        foreach (XmlElement ppluspAppServFields in pluspAppServs)
                        {

                            if (ppluspAppServFields.Name == "COMMODITY") { service.Commodity = ppluspAppServFields.InnerText; }
                            if (ppluspAppServFields.Name == "COMMODITYGROUP") { service.CommodityGroup = ppluspAppServFields.InnerText; }
                            if (ppluspAppServFields.Name == "PLUSPAPPLSERVID") { service.PluspApplServId = ppluspAppServFields.InnerText; }


                        }
                        services.Add(service);
                    }
                    if (responsePlanFields.Name == "PLUSPAPPLCI")
                    {
                        XmlNode pluspAppCis = responsePlanFields;
                        foreach (XmlElement pluspAppCi in pluspAppCis)
                        {
                            if (pluspAppCi.Name == "CINUM") { rpInfo.ConfigurationItems.Add(pluspAppCi.InnerText); }

                        }
                    }
                    if (responsePlanFields.Name == "CLASSSTRUCTURE")
                    {
                        XmlNode classStructures = responsePlanFields;
                        foreach (XmlElement classStructureField in classStructures)
                        {
                            if (classStructureField.Name == "CLASSIFICATIONDESC") { classStructure.ClassificationDesc = classStructureField.InnerText; }
                            if (classStructureField.Name == "DESCRIPTION") { classStructure.Description = classStructureField.InnerText; }
                            if (classStructureField.Name == "HIERARCHYPATH") { classStructure.HierarchyPath = classStructureField.InnerText; }

                        }
                    }
                }

                rpInfo.Services = services.ToArray();
                rpInfo.ClassStructure = classStructure;
            }

            return rpInfo;
        }
        public CdResponsePlanData GetResponsePlanData(string singleRpId)
        {
            CdResponsePlanData ret = new CdResponsePlanData();
            if (!string.IsNullOrEmpty(singleRpId))
            {
                string rpXml = GetResponsePlanXml(singleRpId);
                if (rpXml != "NE")
                    ret = ParseXmlToResponsePlanData(rpXml);
            }
            return ret;
        }
        #endregion
        #region Communication Templates
        public Dictionary<string, string> GetCommTemplatesStatus(List<string> commTemplates)
        {
            Dictionary<string, string> res = new Dictionary<string, string>();

            string query = string.Join("','", commTemplates);

            #region XML(MXCUSTAGREEMENTEXIST) custom object
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                             @"<QueryMXL_COMMTEMPLATEICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""900"">" +
                             "<MXL_COMMTEMPLATEICSQuery>" +
                             "<WHERE>TEMPLATEID IN('" + query + "')</WHERE>" +
                             "</MXL_COMMTEMPLATEICSQuery>" +
                             "</QueryMXL_COMMTEMPLATEICS>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXL_COMMTEMPLATEICS", xml);


            XmlDocument outXml = new XmlDocument();
            outXml.LoadXml(responseText);
            XmlNodeList commTemplatesNodes = outXml.GetElementsByTagName("COMMTEMPLATE");
            foreach (XmlNode commTemplateNode in commTemplatesNodes)
            {
                string id = "", status = "";
                foreach (XmlElement element in commTemplateNode)
                {
                    if (element.Name == "COMMTEMPLATEID") { id = element.InnerText; }
                    if (element.Name == "STATUS") { status = element.InnerText; }
                }
                try { res.Add(id, status); } catch (Exception) { }
            }

            return res;
        }
        #endregion
        #region Customers
        public string GetCustomerName(string customerId)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                @"<QueryGBMMXPLUSPCUSTOMER xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""1"">" +
                "<GBMMXPLUSPCUSTOMERQuery>" +
                "<WHERE>CUSTOMER = '" + customerId + "'</WHERE>" +
                "</GBMMXPLUSPCUSTOMERQuery>" +
                "</QueryGBMMXPLUSPCUSTOMER>";

            string responseText = PostCD(root.UrlCd, "GBMMXPLUSPCUSTOMER", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                if (responseText.Contains("NAME"))
                {
                    try
                    {
                        responseText = outXml.GetElementsByTagName("NAME")[0].InnerText;
                    }
                    catch (Exception ex)
                    {
                        responseText = ex.Message;
                    }
                }
                else
                {
                    responseText = "NE";
                }

            }
            catch (XmlException) { } //la respuesta no es un XML, probablemente error

            #endregion

            return responseText;
        }
        #endregion
        #region Clasifficartions
        public DataTable GetInternalClassificationId(string classStructureId)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("CLASSSTRUCTUREID");
            dt.Columns.Add("CLASSIFICATIONID");
            dt.Columns.Add("APPLICATION");
            dt.Columns.Add("DESCRIPTION");

            classStructureId = classStructureId.Replace("&", "&amp;");

            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += "<QueryMXCLASSIFICATION xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" maxItems=\"900\">";
            xml += "<MXCLASSIFICATIONQuery><WHERE>CLASSSTRUCTUREID = '" + classStructureId + "'</WHERE>";
            xml += "</MXCLASSIFICATIONQuery>";
            xml += "</QueryMXCLASSIFICATION>";

            #region Process Response
            string responseText = PostCD(root.UrlCd, "MXCLASSIFICATION", xml);

            XmlDocument outXml = new XmlDocument();
            outXml.LoadXml(responseText);

            XmlNodeList classStructures = outXml.GetElementsByTagName("CLASSSTRUCTURE");

            foreach (XmlElement classStructure in classStructures)
            {
                string application = "";
                XmlNodeList classUseWiths = classStructure.GetElementsByTagName("CLASSUSEWITH");

                foreach (XmlElement classUseWith in classUseWiths)
                {
                    string objectName = classUseWith.GetElementsByTagName("OBJECTNAME")[0].InnerText;
                    if (objectName == "SR" || objectName == "INCIDENT")
                        application = objectName;
                }

                DataRow row = dt.NewRow();
                row[0] = classStructure.GetElementsByTagName("CLASSSTRUCTUREID")[0].InnerText;
                row[1] = classStructure.GetElementsByTagName("CLASSIFICATIONID")[0].InnerText;
                row[3] = classStructure.GetElementsByTagName("DESCRIPTION")[0].InnerText;
                row[2] = application;
                dt.Rows.Add(row);
            }

            #endregion

            return dt;
        }
        #endregion
        #region Service Groups
        public string GetCommodityGroup(string commodity)
        {
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                @"<QueryMXL_COMMODITIES xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" maxItems=""1"">" +
                "<MXL_COMMODITIESQuery>" +
                "<WHERE>COMMODITY = '" + commodity + "' </WHERE>" +
                "</MXL_COMMODITIESQuery>" +
                "</QueryMXL_COMMODITIES>";

            string responseText = "";
            if (commodity != "")
            {

                responseText = PostCD(root.UrlCd, "MXL_COMMODITIES", xml);

                #region Process Response
                try
                {
                    XmlDocument outXml = new XmlDocument();
                    outXml.LoadXml(responseText);

                    if (responseText.Contains("PARENT"))
                    {
                        try
                        {
                            responseText = outXml.GetElementsByTagName("PARENT")[0].InnerText;
                        }
                        catch (Exception ex)
                        {
                            responseText = ex.Message;
                        }
                    }
                    else
                    {
                        responseText = "NE";
                    }

                }
                catch (XmlException) { } //la respuesta no es un XML, probablemente error

                #endregion

            }
            return responseText;
        }
        #endregion
        #region Releases
        public List<string> GetContractReleases(string customerAgreement)
        {
            List<string> releases = new List<string>();

            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                        "<QueryMXRELEASE xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" maxItems=\"999\">" +
                        "<MXRELEASEQuery><WHERE>WORELEASE.PLUSPAGREEMENT = '" + customerAgreement + "'</WHERE>" +
                        "</MXRELEASEQuery>" +
                        "</QueryMXRELEASE>";

            string resXml = PostCD(root.UrlCd, "MXRELEASE", xml);

            XmlDocument outXml = new XmlDocument();
            outXml.LoadXml(resXml);
            XmlNodeList contracts = outXml.GetElementsByTagName("WORELEASE");
            foreach (XmlNode contract in contracts)
                foreach (XmlElement contractFields in contract)
                    if (contractFields.Name == "WONUM") { releases.Add(contractFields.InnerText); }

            return releases;
        }
        public List<CdReleaseData> GetContractReleaseData(string customerAgreement)
        {
            List<CdReleaseData> releasesData = new List<CdReleaseData>();


            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                        "<QueryMXRELEASE xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" maxItems=\"999\">" +
                        "<MXRELEASEQuery><WHERE>WORELEASE.PLUSPAGREEMENT = '" + customerAgreement + "'</WHERE>" +
                        "</MXRELEASEQuery>" +
                        "</QueryMXRELEASE>";

            string resXml = PostCD(root.UrlCd, "MXRELEASE", xml);

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(resXml);
                XmlNodeList releases = outXml.GetElementsByTagName("WORELEASE");
                foreach (XmlNode release in releases)
                {
                    CdReleaseData releaseData = new CdReleaseData();
                    foreach (XmlElement releaseFields in release)
                    {
                        if (releaseFields.Name == "DESCRIPTION") { releaseData.Description = releaseFields.InnerText; }
                        if (releaseFields.Name == "PLUSPCUSTOMER") { releaseData.PluspCustomer = releaseFields.InnerText; }
                        if (releaseFields.Name == "CLASSSTRUCTUREID") { releaseData.Classification = releaseFields.InnerText; }
                        if (releaseFields.Name == "COMMODITY") { releaseData.Commodity = releaseFields.InnerText; }
                        if (releaseFields.Name == "COMMODITYGROUP") { releaseData.CommodityGroup = releaseFields.InnerText; }
                        if (releaseFields.Name == "ENVIRONMENT") { releaseData.Environment = releaseFields.InnerText; }
                        if (releaseFields.Name == "PMRELEMERGENCY") { releaseData.PmRelEmergency = releaseFields.InnerText; }
                        if (releaseFields.Name == "PMRELIMPACT") { releaseData.PmRelImpact = releaseFields.InnerText; }
                        if (releaseFields.Name == "PMRELURGENCY") { releaseData.PmRelUrgency = releaseFields.InnerText; }
                        if (releaseFields.Name == "WOPRIORITY") { releaseData.WoPriority = releaseFields.InnerText; }
                        if (releaseFields.Name == "TARGSTARTDATE") { releaseData.TargStartDate = releaseFields.InnerText; }
                        if (releaseFields.Name == "TARGCOMPDATE") { releaseData.TargCompDate = releaseFields.InnerText; }
                        if (releaseFields.Name == "ONBEHALFOF") { releaseData.Employee = releaseFields.InnerText; }
                        if (releaseFields.Name == "WHOMISCHANGEFOR") { releaseData.Contact = releaseFields.InnerText; }
                        if (releaseFields.Name == "PLUSPAGREEMENT") { releaseData.Contract = releaseFields.InnerText; }
                        if (releaseFields.Name == "PLUSPREVNUM") { releaseData.ConRev = releaseFields.InnerText; }
                        if (releaseFields.Name == "OWNERGROUP") { releaseData.OwnerGroup = releaseFields.InnerText; }
                        if (releaseFields.Name == "OWNER") { releaseData.Owner = releaseFields.InnerText; }
                        if (releaseFields.Name == "EXTERNALREFERENCE") { releaseData.ExtRef = releaseFields.InnerText; }
                        if (releaseFields.Name == "WONUM") { releaseData.relId = releaseFields.InnerText; }
                    }
                    releasesData.Add(releaseData);
                }
            }
            catch (Exception)
            {
                Console.WriteLine(resXml);
            }

            return releasesData;
        }
        #endregion

        #endregion

        #region Métodos de Creación

        #region Releases
        public string[] CreateRelease(CdReleaseData rel)
        {
            string[] responseText = { "", "", "" };

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                           "<SyncMXRELEASE xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\">" +
                           "<MXRELEASESet action=\"AddChange\">" +
                           "<WORELEASE action=\"AddChange\">" +
                           "<SITEID>GBMHQ</SITEID>" +
                           "<DESCRIPTION>" + rel.Description.Replace("&", "&amp;") + "</DESCRIPTION>" +
                           "<PLUSPCUSTOMER>" + rel.PluspCustomer + "</PLUSPCUSTOMER>" +

            #region Campos de los valores obligatorios en la aplicación
                            "<CLASSSTRUCTUREID>" + rel.Classification + "</CLASSSTRUCTUREID>" +
                           "<COMMODITY>" + rel.Commodity + "</COMMODITY>" +
                           "<COMMODITYGROUP>" + rel.CommodityGroup + "</COMMODITYGROUP>" +
                           "<ENVIRONMENT>" + rel.Environment + "</ENVIRONMENT>" +
                           "<PMRELEMERGENCY>" + rel.PmRelEmergency + "</PMRELEMERGENCY>" +
                           "<PMRELIMPACT>" + rel.PmRelImpact + "</PMRELIMPACT>" +
                           "<PMRELURGENCY>" + rel.PmRelUrgency + "</PMRELURGENCY>" +
                           "<WOPRIORITY>" + rel.WoPriority + "</WOPRIORITY>" +
            #endregion

                            "<TARGSTARTDATE>" + rel.TargStartDate + "T08:00:00-06:00</TARGSTARTDATE> " +
                           "<TARGCOMPDATE>" + rel.TargCompDate + "T08:00:00-06:00</TARGCOMPDATE>" +
                           "<ONBEHALFOF>" + rel.Employee + "</ONBEHALFOF>" +
                           "<WHOMISCHANGEFOR>" + rel.Contact + "</WHOMISCHANGEFOR>" +
                           "<PLUSPAGREEMENT>" + rel.Contract + "</PLUSPAGREEMENT>" +
                           "<PLUSPREVNUM>" + rel.ConRev + "</PLUSPREVNUM>" +
                           "<OWNERGROUP>" + rel.OwnerGroup + "</OWNERGROUP>" +
                           "<EXTERNALREFERENCE>" + rel.ExtRef.Replace("&", "&amp;") + "</EXTERNALREFERENCE>" +
                           "</WORELEASE>" +
                           "</MXRELEASESet>" +
                           "</SyncMXRELEASE>";
            #endregion

            string res = PostCD(root.UrlCd, "MXRELEASE", xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(res);

                try
                {
                    responseText[0] = outXml.GetElementsByTagName("WONUM")[0].InnerText;
                    responseText[1] = outXml.GetElementsByTagName("WORKORDERID")[0].InnerText;
                }
                catch (Exception ex)
                {
                    responseText[2] = ex.Message;
                }

            }
            catch (XmlException)
            {
                //la respuesta no es un XML, probablemente error
                responseText[2] = res;
            }
            #endregion

            return responseText;
        }
        #endregion
        #region Service Level Agreements
        public string CreateSla(CdSlaData sla)
        {
            #region xmlServicePart
            string xmlServicePart = "";
            foreach (string commodity in sla.PluspApplServCommodity)
            {
                xmlServicePart += "<PLUSPAPPLSERV>";
                xmlServicePart += "<COMMODITYGROUP>" + GetCommodityGroup(commodity) + "</COMMODITYGROUP>";
                xmlServicePart += "<COMMODITY>" + commodity + "</COMMODITY>";
                xmlServicePart += "<ITEMSETID>ITEMSET1</ITEMSETID>";
                xmlServicePart += "<OWNERID>" + sla.Sanum + "</OWNERID>";
                xmlServicePart += "<OWNERTABLE>SLA</OWNERTABLE>";
                xmlServicePart += "</PLUSPAPPLSERV>";
            }
            #endregion

            #region xmlCommitmentsPart
            string xmlCommitmentPart = "";
            foreach (CdSlaCommitments commitment in sla.CdSlaCommitments)
            {
                xmlCommitmentPart += "<SLACOMMITMENTS>";
                xmlCommitmentPart += "<DESCRIPTION>" + commitment.Description + "</DESCRIPTION>";
                xmlCommitmentPart += "<TYPE>" + commitment.Type + "</TYPE>";
                xmlCommitmentPart += "<COMMITMENTID>" + commitment.Type + "</COMMITMENTID>";
                xmlCommitmentPart += "<VALUE>" + commitment.Value.Replace(',', '.') + "</VALUE>";
                xmlCommitmentPart += "<UNITOFMEASURE>" + commitment.UnitOfMeasure + "</UNITOFMEASURE>";
                xmlCommitmentPart += "</SLACOMMITMENTS>";
            }

            //#region xmlResponsePart

            //if (sla.ResponseDescription != "")
            //{
            //    xmlResponsePart += "<SLACOMMITMENTS>";
            //    xmlResponsePart += "<DESCRIPTION>" + sla.ResponseDescription + "</DESCRIPTION>";
            //    xmlResponsePart += "<TYPE>" + sla.ResponseType + "</TYPE>";
            //    xmlResponsePart += "<COMMITMENTID>" + sla.ResponseCommitmentId + "</COMMITMENTID>";
            //    xmlResponsePart += "<VALUE>" + sla.ResponseValue + "</VALUE>";
            //    xmlResponsePart += "<UNITOFMEASURE>HOURS</UNITOFMEASURE>";
            //    xmlResponsePart += "</SLACOMMITMENTS>";
            //}
            //#endregion

            //#region xmlResolutionPart
            //string xmlResolutionPart = "";
            //if (sla.ResolutionDescription != "" && sla.ResolutionType != "" && sla.ResolutionCommitmentId != "" && sla.ResolutionValue != "")
            //{
            //    xmlResolutionPart += "<SLACOMMITMENTS>";
            //    xmlResolutionPart += "<DESCRIPTION>" + sla.ResolutionDescription + "</DESCRIPTION>";
            //    xmlResolutionPart += "<TYPE>" + sla.ResolutionType + "</TYPE>";
            //    xmlResolutionPart += "<COMMITMENTID>" + sla.ResolutionCommitmentId + "</COMMITMENTID>";
            //    xmlResolutionPart += "<VALUE>" + sla.ResolutionValue + "</VALUE>";
            //    xmlResolutionPart += "<UNITOFMEASURE>HOURS</UNITOFMEASURE>";
            //    xmlResolutionPart += "</SLACOMMITMENTS>";
            //}
            //#endregion

            #endregion

            #region xmlCalcPart
            string xmlCalcPart = "";
            if (sla.CalcOrgId != "" && sla.CalcCalendar != "" && sla.CalcShift != "")
            {
                xmlCalcPart += "<CALCORGID>" + sla.CalcOrgId + "</CALCORGID>";
                xmlCalcPart += "<CALCCALENDAR>" + sla.CalcCalendar + "</CALCCALENDAR>";
                xmlCalcPart += "<CALCSHIFT>" + sla.CalcShift + "</CALCSHIFT>";
            }
            #endregion

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMX-SLA xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MX-SLASet>";
            xml += @"<SLA action=""AddChange"">";
            xml += "<SANUM>" + sla.Sanum + "</SANUM>";
            xml += "<OBJECTNAME>" + sla.ObjectName + "</OBJECTNAME>";
            xml += "<RANKING>" + sla.Ranking + "</RANKING>";
            xml += "<INTPRIORITYEVAL>" + sla.IntPriorityEval + "</INTPRIORITYEVAL>";
            xml += "<INTPRIORITYVALUE>" + sla.IntPriorityValue + "</INTPRIORITYVALUE>";
            xml += xmlCalcPart;
            xml += "<DESCRIPTION>" + sla.Description + "</DESCRIPTION>";
            xml += "<CONDITION>" + sla.Condition.Replace("\"", "&apos;") + "</CONDITION>";
            xml += xmlCommitmentPart;
            xml += xmlServicePart;
            xml += "</SLA>";
            xml += "</MX-SLASet>";
            xml += "</SyncMX-SLA>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MX-SLA", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Tickets(SR/IN)
        public string[] CreateTicket(CdTicketData ticket)
        {
            string[] ret = { "", "", "" };

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += "<SyncMX" + ticket.TicketType + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\">";
            xml += "<MX" + ticket.TicketType + "Set action=\"AddChange\">";
            xml += "<" + ticket.TicketType + " action=\"AddChange\">";
            xml += "<REPORTEDBY>" + ticket.ReportedEmail + "</REPORTEDBY>";
            xml += "<COUNTRY>" + ticket.Country + "</COUNTRY>";
            xml += "<CLASSSTRUCTUREID>" + ticket.ClassStructureId + "</CLASSSTRUCTUREID>";
            if (!string.IsNullOrWhiteSpace(ticket.CommodityGroup))
                xml += "<COMMODITYGROUP>" + ticket.CommodityGroup + "</COMMODITYGROUP>";
            xml += "<COMMODITY>" + ticket.Commodity + "</COMMODITY>";
            xml += "<PLUSPCUSTOMER>" + ticket.PluspCustomer + "</PLUSPCUSTOMER>";
            xml += "<DESCRIPTION>" + ticket.Description + "</DESCRIPTION>";
            xml += "<DESCRIPTION_LONGDESCRIPTION>" + ticket.LongDescription + "</DESCRIPTION_LONGDESCRIPTION>";
            xml += "<IMPACT>" + ticket.Impact + "</IMPACT>";
            xml += "<URGENCY>" + ticket.Urgency + "</URGENCY>";
            xml += "<EXTERNALSYSTEM>" + ticket.ExternalSystem + "</EXTERNALSYSTEM>";
            xml += "<GBMPLUSPAGREEMENT>" + ticket.GbmPluspAgreement + "</GBMPLUSPAGREEMENT>";
            xml += "</" + ticket.TicketType + ">";
            xml += "</MX" + ticket.TicketType + "Set>";
            xml += "</SyncMX" + ticket.TicketType + ">";
            #endregion

            string responseText = PostCD(root.UrlCd + "MX", ticket.TicketType, xml);

            #region Process Response
            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);

                try
                {
                    ret[0] = outXml.GetElementsByTagName("TICKETID")[0].InnerText;
                    ret[1] = outXml.GetElementsByTagName("TICKETUID")[0].InnerText;
                }
                catch (Exception ex)
                {
                    ret[2] = ex.Message;
                }
            }
            catch (XmlException)
            {
                //la respuesta no es un XML, probablemente error
                ret[2] = responseText;
            }
            #endregion
            return ret;
        }
        #endregion
        #region Users
        public string CreateUser(CdUserData cd)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += "<SyncMXL_USERGRP xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\">";
            xml += "<MXL_USERGRPSet action=\"AddChange\">";
            xml += "<MAXUSER action=\"Change\">";
            xml += "<USERID>" + cd.User.Trim() + "</USERID>";

            foreach (string rol in cd.Roles)
            {
                xml += "<GROUPUSER action=\"AddChange\">";
                xml += "<GROUPNAME>" + rol.Trim() + "</GROUPNAME>";
                xml += "</GROUPUSER>";
            }
            xml += "</MAXUSER>";
            xml += "</MXL_USERGRPSet>";
            xml += "</SyncMXL_USERGRP>";
            #endregion

            string Response_Text = PostCD(root.UrlCd, "MXL_USERGRP", xml);

            #region true o false

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(Response_Text);
                XmlNodeList tempList = outXml.GetElementsByTagName("USERID");

                string change = "";

                try
                {
                    change = tempList[0].Attributes.GetNamedItem("changed").Value;
                }
                catch (Exception)
                {
                    change = tempList[0].InnerText == cd.User.ToUpper() ? "2" : "";
                }

                if (change == "1")
                {
                    //todo bien
                    return "OK";
                }
                else if (change == "2")
                {
                    return "No se aplicaron cambios";
                }
                else
                {
                    //return "Error al actualizar los roles, la consulta fue correcta pero el resultado esperado no";
                    return Response_Text;
                }
            }
            catch (XmlException)
            {
                //la respuesta no es un XML, probablemente error
                return Response_Text;
            }

            #endregion
        }
        #endregion
        #region Customer Agreements
        public string CreateContract(CdContractData con)
        {
            string ret;

            if (string.IsNullOrEmpty(con.Revision))
                con.Revision = "0";

            #region Parte del XML de los MG (PLUSPAPPLSERV)
            string xmlMgPart = "";
            foreach (string item in con.MaterialArray)
            {
                xmlMgPart += @"<PLUSPAPPLSERV action=""AddChange"">";
                xmlMgPart += "<COMMODITY>" + item + "</COMMODITY>";
                xmlMgPart += "<ITEMSETID>ITEMSET1</ITEMSETID>";
                if (con.ManualServiceArray != null)
                    if (con.ManualServiceArray.Contains(item))
                        xmlMgPart += "<MANUAL>1</MANUAL>";
                xmlMgPart += "<OWNERTABLE>PLUSPPRICESCHED</OWNERTABLE>";
                xmlMgPart += "</PLUSPAPPLSERV>";
            }
            #endregion

            #region Parte del XML de equipos (PLUSPAPPLSERV)
            string xmlEquiPart = "";
            if (con.EquipArray.Count > 0)
            {
                foreach (string equipment in con.EquipArray)
                {
                    xmlEquiPart += @"<PLUSPAPPLASSET action=""AddChange"">";
                    xmlEquiPart += "<SITEID>GBMHQ</SITEID>";
                    xmlEquiPart += "<ASSETNUM>" + equipment + "</ASSETNUM>";
                    xmlEquiPart += "<INCLUDECHILDREN>0</INCLUDECHILDREN>";
                    xmlEquiPart += "</PLUSPAPPLASSET>";
                }
            }
            #endregion

            #region  Service XML Request

            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" >";
            xml += @"<MXCUSTAGREEMENTSet action=""AddChange"">";
            xml += @"<PLUSPAGREEMENT action=""AddChange"">";
            xml += "<AGREEMENT>" + con.IdContract + "</AGREEMENT>";
            xml += "<DESCRIPTION>" + con.Description.Replace("&", "&amp;") + "</DESCRIPTION>";
            xml += "<ORGID>GBM</ORGID>";
            xml += "<REVISIONNUM>" + con.Revision + "</REVISIONNUM>";
            xml += "<CUSTOMER>" + con.IdCustomer + "</CUSTOMER>";
            xml += "<STARTDATE>" + con.StartDate + "</STARTDATE>";
            xml += "<ENDDATE>" + con.EndDate + "</ENDDATE>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_I</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Incident</DESCRIPTION>";
            xml += "<OBJECTNAME>INCIDENT</OBJECTNAME>";
            xml += "<RANKING>1</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_R</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Request</DESCRIPTION>";
            xml += "<OBJECTNAME>SR</OBJECTNAME>";
            xml += "<RANKING>2</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_P</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Problem</DESCRIPTION>";
            xml += "<OBJECTNAME>PROBLEM</OBJECTNAME>";
            xml += "<RANKING>3</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_C</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Change</DESCRIPTION>";
            xml += "<OBJECTNAME>WOCHANGE</OBJECTNAME>";
            xml += "<RANKING>4</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_RE</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Release</DESCRIPTION>";
            xml += "<OBJECTNAME>WORELEASE</OBJECTNAME>";
            xml += "<RANKING>5</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_A</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Activity</DESCRIPTION>";
            xml += "<OBJECTNAME>WOACTIVITY</OBJECTNAME>";
            xml += "<RANKING>6</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += @"<PLUSPPRICESCHED action=""AddChange"">";
            xml += "<PRICESCHEDULE>PS_W</PRICESCHEDULE>";
            xml += "<DESCRIPTION>Workorder</DESCRIPTION>";
            xml += "<OBJECTNAME>WORKORDER</OBJECTNAME>";
            xml += "<RANKING>7</RANKING>";
            xml += xmlMgPart;
            xml += xmlEquiPart;
            xml += "</PLUSPPRICESCHED>";

            xml += "</PLUSPAGREEMENT>";
            xml += "</MXCUSTAGREEMENTSet>";
            xml += "</SyncMXCUSTAGREEMENT>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXCUSTAGREEMENT", xml);

            if (ValidateXml(responseText))
                ret = "OK";
            else
                ret = responseText;

            return ret;
        }
        #endregion
        #region Contacts
        public string CreateContact(CdContactData cont)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXPERSON xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += @"<MXPERSONSet action=""AddChange"">";
            xml += @"<PERSON action=""AddChange"">";
            xml += "<PERSONID>" + cont.PersonId + "</PERSONID>";
            xml += "<FIRSTNAME>" + cont.FirstName + "</FIRSTNAME>";
            xml += "<LASTNAME>" + cont.LastName + "</LASTNAME>";
            xml += "<GBMSAPID>" + cont.SapId + "</GBMSAPID>";
            xml += "<PRIMARYSMS>" + cont.Telephone + "</PRIMARYSMS>";
            xml += @"<EMAIL action=""AddChange"">";
            xml += "<EMAILADDRESS>" + cont.Email + "</EMAILADDRESS>";
            xml += "<ISPRIMARY>1</ISPRIMARY>";
            xml += "</EMAIL>";
            xml += "</PERSON>";
            xml += "</MXPERSONSet>";
            xml += "</SyncMXPERSON>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXPERSON", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;

        }
        #endregion
        #region Customers
        public string CreateCustomer(CdCustomerData cust)
        {
            #region XML
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncGBMMXPLUSPCUSTOMER xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"" >";
            xml += @"<GBMMXPLUSPCUSTOMERSet action=""AddChange"">";
            xml += @"<PLUSPCUSTOMER action=""AddChange"">";
            xml += "<CUSTOMER>" + cust.IdCustomer + "</CUSTOMER>";
            xml += "<NAME>" + string.Concat(cust.NameCustomer.Take(50)).Replace("&", "&amp;") + "</NAME>";
            xml += "<COUNTRY>" + cust.Country + "</COUNTRY>";
            xml += "<CURRENCYCODE>USD</CURRENCYCODE>";
            xml += "<STATUS>ACTIVE</STATUS>";
            xml += @"<PLUSPCUSTCONTACT action=""AddChange"">";
            xml += "<PERSONID>" + cust.PersonId + "</PERSONID>";
            xml += "<TYPE>CONTRACT</TYPE>";
            xml += "</PLUSPCUSTCONTACT>";
            xml += "</PLUSPCUSTOMER>";
            xml += "</GBMMXPLUSPCUSTOMERSet>";
            xml += "</SyncGBMMXPLUSPCUSTOMER>";
            #endregion

            string responseText = PostCD(root.UrlCd, "GBMMXPLUSPCUSTOMER", xml);
            if (ValidateXml(responseText))
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Locations
        public string CreateLocation(CdLocationData loc)
        {
            #region Service XML Request
            string xml =
                @"<?xml version=""1.0"" encoding=""utf-8""?>" +
                @"<SyncMXOPERLOC xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">" +
                @"<MXOPERLOCSet action=""AddChange"">" +
                @"<LOCATIONS action=""AddChange"">" +
                "<LOCATION>" + loc.Location.Replace("&", "&amp;") + "</LOCATION>" +
                "<SITEID>GBMHQ</SITEID>" +
                "<DESCRIPTION>" + loc.Description.Replace("&", "&amp;") + "</DESCRIPTION>" +
                "<STATUS>OPERATING</STATUS>" +
                "<TYPE>OPERATING</TYPE>" +
                "</LOCATIONS>" +
                "</MXOPERLOCSet>" +
                "</SyncMXOPERLOC>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXOPERLOC", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Assets
        public string CreateAsset(CdAssetData asset)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXASSET xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += @"  <MXASSETSet action=""AddChange"">";
            xml += @"  <ASSET action=""AddChange"">";
            xml += "    <ASSETNUM>" + asset.AssetNum + "</ASSETNUM>";
            xml += "   <SITEID>GBMHQ</SITEID>";
            xml += " <DESCRIPTION>" + asset.AssetText.Replace("&", "&amp;") + "</DESCRIPTION>";
            xml += "  <STATUS>OPERATING</STATUS>";
            xml += "  <LOCATION>" + asset.Location + "</LOCATION>";
            xml += "  <SERIALNUM>" + asset.SerialNum + "</SERIALNUM>";
            xml += "  <GBMITEMMATERIAL>" + asset.MaterialNum + "</GBMITEMMATERIAL>";
            if (asset.Warranty != "")
            {
                xml += "  <INSTALLDATE>" + asset.InstallDate + "T08:00:00-06:00</INSTALLDATE>";
                xml += "  <EQ23>" + asset.EndDate + "T08:00:00-06:00</EQ23>";
                xml += "  <GBMWARRANTY>" + asset.Warranty + "</GBMWARRANTY>";
                xml += "  <GBMWARRDESCRIPTION>" + asset.WarrantyText + "</GBMWARRDESCRIPTION>";
            }
            xml = xml + "  <COMMODITY>" + asset.MaterialGroup + "</COMMODITY>";
            if (asset.Placa != "")
            {
                xml += "  <EQ1>" + asset.Placa + "</EQ1>";
            }
            xml += " </ASSET>";
            xml += "</MXASSETSet>";
            xml += "</SyncMXASSET>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXASSET", xml);
            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Response Plans
        public CdResponsePlanData CreateOrChangeResponsePlans(CdResponsePlanData rp)
        {
            #region xmlCiPart
            string xmlCiPart = "";
            if (rp.ConfigurationItems != null)
            {
                foreach (string ci in rp.ConfigurationItems)
                {
                    xmlCiPart += "<PLUSPAPPLCI>";
                    xmlCiPart += "<CINUM>" + ci + "</CINUM>";
                    xmlCiPart += String.IsNullOrEmpty(rp.Sanum) ? "" : "<OWNERID>" + rp.Sanum + "</OWNERID>";
                    xmlCiPart += "<OWNERTABLE>PLUSPRESPPLAN</OWNERTABLE>";
                    xmlCiPart += "</PLUSPAPPLCI>";
                }
            }
            #endregion

            #region xmlCommodityPart
            string xmlCommodityPart = "";
            if (rp.Services != null)
            {
                foreach (CdServicesData commodity in rp.Services)
                {
                    if (!String.IsNullOrEmpty(commodity.Commodity))
                    {
                        xmlCommodityPart += "<PLUSPAPPLSERV>";
                        xmlCommodityPart += "<COMMODITYGROUP>" + GetCommodityGroup(commodity.Commodity) + "</COMMODITYGROUP>";
                        xmlCommodityPart += "<COMMODITY>" + commodity.Commodity + "</COMMODITY>";
                        xmlCommodityPart += "<ITEMSETID>ITEMSET1</ITEMSETID>";
                        xmlCommodityPart += String.IsNullOrEmpty(rp.Sanum) ? "" : "<OWNERID>" + rp.Sanum + "</OWNERID>";
                        xmlCommodityPart += "</PLUSPAPPLSERV>";
                    }
                }
            }
            #endregion

            #region xmlAction
            string xmlEscRefPoint = "";

            if (!String.IsNullOrEmpty(rp.Action))

            {
                xmlEscRefPoint += "<ESCREFPOINT>";
                xmlEscRefPoint += "<ACTION>" + rp.Action + "</ACTION>";
                xmlEscRefPoint += "<REFPOINTNUM>1</REFPOINTNUM>";
                xmlEscRefPoint += "</ESCREFPOINT>";
            }

            #endregion

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXRESPONSEPLANICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXRESPONSEPLANICSSet>";
            xml += @"<PLUSPRESPPLAN action=""AddChange"">";
            xml += String.IsNullOrEmpty(rp.Sanum) ? "" : "<SANUM>" + rp.Sanum + "</SANUM>";
            xml += "<DESCRIPTION>" + rp.Description.Replace("&", "&amp;") + "</DESCRIPTION>";
            xml += "<GBMAUTOASSIGNMENT>" + rp.GbmAutoAssignment + "</GBMAUTOASSIGNMENT>";
            xml += "<RANKING>" + rp.Ranking + "</RANKING>";
            xml += "<OBJECTNAME>" + rp.ObjectName + "</OBJECTNAME>";
            xml += "<CALENDARORGID>" + rp.CalendarOrgId + "</CALENDARORGID>";
            xml += "<CALENDAR>" + rp.Calendar + "</CALENDAR>";
            xml += "<SHIFT>" + rp.Shift + "</SHIFT>";
            xml += "<ASSIGNOWNERGROUP>" + rp.AssignOwnerGroup + "</ASSIGNOWNERGROUP>";
            xml += "<STATUS>DRAFT</STATUS>";
            xml += "<CLASSSTRUCTUREID>" + rp.ClassStructureId + "</CLASSSTRUCTUREID>";
            xml += "<CONDITION>" + rp.Condition.Replace("&", "&amp;") + "</CONDITION>";
            xml += xmlCiPart;
            xml += xmlCommodityPart;
            xml += xmlEscRefPoint;
            xml += "</PLUSPRESPPLAN>";
            xml += "</MXRESPONSEPLANICSSet>";
            xml += "</SyncMXRESPONSEPLANICS>";
            #endregion

            rp.ResponseMessage = PostCD(root.UrlCd, "MXRESPONSEPLANICS", xml);

            if (rp.Sanum == "" && !rp.ResponseMessage.Contains("Error"))
            {
                try
                {
                    XmlDocument outXml = new XmlDocument();
                    outXml.LoadXml(rp.ResponseMessage);

                    try
                    {
                        rp.Sanum = outXml.GetElementsByTagName("SANUM")[0].InnerText;
                    }
                    catch (Exception ex)
                    {
                        rp.ResponseMessage = "Error: " + ex.Message;
                    }
                }
                catch (XmlException ex)
                {
                    rp.ResponseMessage = "Error: " + ex.Message;
                }

            }

            if (ValidateXml(rp.ResponseMessage))
                rp.ResponseMessage = ChangeResponsePlanStatus(rp.Sanum, rp.Status);

            return rp;
        }
        #endregion
        #region Person Groups
        public string CreatePersonGroup(CdPersonGroupData pg)
        {
            #region xmlPersonGroupTeamPart
            string xmlPersonGroupTeamPart = "";
            string groupDefault = "1";
            foreach (string respPartyGroup in pg.Emails)
            {
                int respPartyGroupSeq = 10;
                xmlPersonGroupTeamPart += "<PERSONGROUPTEAM>";
                xmlPersonGroupTeamPart += "<RESPPARTY>" + respPartyGroup + "</RESPPARTY>";
                xmlPersonGroupTeamPart += "<RESPPARTYGROUP>" + respPartyGroup + "</RESPPARTYGROUP>";
                xmlPersonGroupTeamPart += "<RESPPARTYGROUPSEQ>" + respPartyGroupSeq + "</RESPPARTYGROUPSEQ>";
                xmlPersonGroupTeamPart += "<GROUPDEFAULT>" + groupDefault + "</GROUPDEFAULT>";
                xmlPersonGroupTeamPart += "</PERSONGROUPTEAM>";
                groupDefault = "0";
            }

            #endregion


            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXL_PERGRP xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXL_PERGRPSet>";
            xml += "<PERSONGROUP action=\"AddChange\">";
            xml += "<PERSONGROUP>" + pg.PersonGroupId + "</PERSONGROUP>";
            xml += xmlPersonGroupTeamPart;
            xml += "</PERSONGROUP>";
            xml += "</MXL_PERGRPSet>";
            xml += "</SyncMXL_PERGRP>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXL_PERGRP", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Communication Templates
        public string CreateCommunicationTemplates(CdCommTemplate ct)
        {
            string error = "";
            string templateId = "";

            foreach (string pg in ct.CommTmpltSendToValue)
            {
                string type = "GROUP";
                if (pg.Contains("@"))
                    type = "EMAIL";
                else if (pg == "GBMINCOWNP" || pg == "GBMSRCOWNP")
                    type = "ROLE";

                #region Service XML Request
                string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
                xml += @"<SyncMXL_COMMTEMPLATEICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
                xml += "<MXL_COMMTEMPLATEICSSet>";
                xml += "<COMMTEMPLATE action=\"AddChange\">";
                if (templateId != "")
                    xml += "<TEMPLATEID>" + templateId + "</TEMPLATEID>";
                xml += "<DESCRIPTION>" + ct.Description + "</DESCRIPTION>";
                xml += "<OBJECTNAME>" + ct.ObjectName + "</OBJECTNAME>";
                xml += "<USEWITH>ESCALATION</USEWITH>";
                xml += "<SENDFROM>controldesk@gbm.net</SENDFROM>";
                xml += "<TRACKFAILEDMESSAGES>0</TRACKFAILEDMESSAGES>";
                xml += "<LOGFLAG>1</LOGFLAG>";
                xml += "<SUBJECT>" + ct.Subject + "</SUBJECT>";
                xml += "<MESSAGE>" + ct.Message + "</MESSAGE>";
                xml += "<COMMTMPLTSENDTO>";
                xml += "<TYPE>" + type + "</TYPE>";
                xml += "<SENDTOVALUE>" + pg.ToUpper() + "</SENDTOVALUE>";
                xml += "<SENDTO>1</SENDTO>";
                xml += "<CC>0</CC>";
                if (type == "GROUP")
                    xml += "<ISBROADCAST>1</ISBROADCAST>";
                xml += "<BCC>0</BCC>";
                xml += "</COMMTMPLTSENDTO>";
                xml += "</COMMTEMPLATE>";
                xml += "</MXL_COMMTEMPLATEICSSet>";
                xml += "</SyncMXL_COMMTEMPLATEICS>";
                #endregion

                string responseText = PostCD(root.UrlCd, "MXL_COMMTEMPLATEICS", xml);

                #region Process Response
                try
                {
                    XmlDocument outXml = new XmlDocument();
                    outXml.LoadXml(responseText);

                    try
                    {
                        templateId = outXml.GetElementsByTagName("TEMPLATEID")[0].InnerText;
                        //ret[1] = outXml.GetElementsByTagName("COMMTEMPLATEID")[0].InnerText;
                    }
                    catch (Exception ex)
                    {
                        error = ex.Message;
                    }
                }
                catch (XmlException)
                {
                    //la respuesta no es un XML, probablemente error
                    error = responseText;
                }
                #endregion

                if (error != "")
                    return "Error: " + error;
            }
            return templateId;
        }
        #endregion
        #region Collections
        public string CreateCollection(CdCollectionData coll)
        {
            #region xmlCiPart
            string xmlParties = "";
            foreach (CdCollectionPartyData party in coll.CollectionParties)
            {
                xmlParties += "<GBMCOLLECTIONINTPARTY>";
                if (party.Id != null)
                    xmlParties += "<GBMCOLLECTIONINTPARTYID>" + party.Id + "</GBMCOLLECTIONINTPARTYID>";
                xmlParties += "<PERSONID>" + party.PersonId + "</PERSONID>";
                xmlParties += "<PERSONGROUP>" + party.PersonGroup + "</PERSONGROUP>";
                xmlParties += "<DESCRIPTION>" + party.Description + "</DESCRIPTION>";
                xmlParties += "<TYPE>" + party.Type + "</TYPE>";
                xmlParties += "</GBMCOLLECTIONINTPARTY>";
            }
            #endregion

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXCOLLECTIONICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += @"<MXCOLLECTIONICSSet action=""AddChange"">";
            xml += @"<COLLECTION action=""AddChange"">";
            xml += "<COLLECTIONNUM>" + coll.CollectionNum.Trim() + "</COLLECTIONNUM>";
            xml += "<DESCRIPTION>" + coll.Description + "</DESCRIPTION>";
            xml += "<COLLSUPERVISOR>" + coll.Supervisor + "</COLLSUPERVISOR>";
            xml += "<ISACTIVE>1</ISACTIVE>";
            xml += xmlParties;
            xml += "</COLLECTION>";
            xml += "</MXCOLLECTIONICSSet>";
            xml += "</SyncMXCOLLECTIONICS>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXCOLLECTIONICS", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Escalations
        public string CreateSlaEscalation(CdSlaData sla, string sch)
        {
            #region xmlEscalaationsPart
            string xmlEscalationPart = "";
            for (int i = 0; i < sla.CdSlaEscalations.Count; i++)
            {
                string refPointNum = (i + 1).ToString();
                CdSlaEscalation escalation = sla.CdSlaEscalations[i];
                CdCommTemplate commTemplate = escalation.Notifications[0]; //solo toma una notification

                xmlEscalationPart += "<ESCREFPOINT>";
                xmlEscalationPart += "<ELAPSEDINTERVAL>" + escalation.TimeInterval.Replace(',', '.') + "</ELAPSEDINTERVAL>";
                xmlEscalationPart += "<EVENTATTRIBUTE>" + escalation.TimeAttribute + "</EVENTATTRIBUTE>";
                xmlEscalationPart += "<EVENTCONDITION>" + escalation.Condition + "</EVENTCONDITION>";
                xmlEscalationPart += "<INTERVALUOM>HOURS</INTERVALUOM>";
                xmlEscalationPart += "<REFPOINTNUM>" + refPointNum + "</REFPOINTNUM>";
                xmlEscalationPart += "<CALCORGID>" + sla.CalcOrgId + "</CALCORGID>";
                xmlEscalationPart += "<CALCCALENDAR>" + sla.CalcCalendar + "</CALCCALENDAR>";
                xmlEscalationPart += "<CALCSHIFT>" + sla.CalcShift + "</CALCSHIFT>";
                xmlEscalationPart += "<ESCNOTIFICATION>";
                xmlEscalationPart += "<TEMPLATEID>" + commTemplate.TemplateId + "</TEMPLATEID>";
                xmlEscalationPart += "</ESCNOTIFICATION>";
                xmlEscalationPart += "</ESCREFPOINT>";
            }
            #endregion

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXL_ESCALATIONICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";

            xml += "<MXL_ESCALATIONICSSet>";
            xml += @"<ESCALATION action=""AddChange"">";
            xml += "<ESCALATION>" + sla.Escalation + "</ESCALATION>";
            xml += "<ACTIVE>0</ACTIVE>";
            xml += "<SCHEDULE>" + sch + "</SCHEDULE>";
            xml += xmlEscalationPart;
            xml += "</ESCALATION>";
            xml += "</MXL_ESCALATIONICSSet>";
            xml += "</SyncMXL_ESCALATIONICS>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXL_ESCALATIONICS", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;

        }
        #endregion
        #region Configuration Items
        public string CreateCi(CdConfigurationItemData ci)
        {
            #region collectDetailsPart
            string xmlCispecPart = "";
            foreach (CdCiSpecData ciSpec in ci.CiSpecs)
            {
                xmlCispecPart += @"<CISPEC action=""Add"">";
                xmlCispecPart += "<ASSETATTRID>" + ciSpec.AssetAttrId + "</ASSETATTRID>";
                xmlCispecPart += "<CCISUMSPECVALUE>" + ciSpec.CCiSumSpecValue + "</CCISUMSPECVALUE>";
                xmlCispecPart += "<SECTION> </SECTION>";
                xmlCispecPart += "</CISPEC>";
            }
            #endregion


            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXAUTHCI xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += @"<MXAUTHCISet action=""Add"">";
            xml += @"<CI action=""Add"">";
            xml += "<CINUM>" + ci.CiNum + "</CINUM>";
            xml += "<CINAME>" + ci.CiName + "</CINAME>";
            xml += "<DESCRIPTION>" + ci.Description + "</DESCRIPTION>";
            xml += "<STATUS>" + ci.Status + "</STATUS>";
            xml += "<PMCCIIMPACT>" + ci.PmcCiImpact + "</PMCCIIMPACT>";
            xml += "<CILOCATION>" + ci.CiLocation + "</CILOCATION>";
            xml += "<CLASSSTRUCTUREID>" + ci.ClassStructureId + "</CLASSSTRUCTUREID>";
            xml += "<SERVICEGROUP>" + ci.ServiceGroup + "</SERVICEGROUP>";
            xml += "<SERVICE>" + ci.Service + "</SERVICE>";
            xml += "<PLUSPCUSTOMER>" + ci.PluspCustomer + "</PLUSPCUSTOMER>";
            xml += "<PERSONID>" + ci.PersonId + "</PERSONID>";
            xml += "<GBM_ADMINISTRATOR>" + ci.GbmAdministrator + "</GBM_ADMINISTRATOR>";
            xml += xmlCispecPart;
            xml += "</CI>";
            xml += "</MXAUTHCISet>";
            xml += "</SyncMXAUTHCI>";


            #endregion

            string responseText = PostCD(root.UrlCd, "MXAUTHCI", xml);

            if (ValidateXml(responseText))
                return "OK";
            else
                return responseText;
        }
        #endregion

        #endregion

        #region Métodos de Modificación

        internal string AddCiContract(CdContractData contract)

        {
            #region collectDetailsPart
            string xmlCiPart = "";
            foreach (string ci in contract.CisArray)
            {
                foreach (CdPriceScheduleData pSData in contract.PriceSchedules)
                {
                    xmlCiPart += "<PLUSPPRICESCHED action=\"AddChange\">";
                    xmlCiPart += "<PRICESCHEDULE>" + pSData.PriceSchedule + "</PRICESCHEDULE>";
                    xmlCiPart += "<PLUSPAPPLCI action=\"AddChange\">";
                    xmlCiPart += "<OWNERID>" + pSData.SaNum + "</OWNERID>";
                    xmlCiPart += "<CINUM>" + ci + "</CINUM>";
                    xmlCiPart += "<OWNERTABLE>PLUSPPRICESCHED</OWNERTABLE>";
                    xmlCiPart += "</PLUSPAPPLCI>";
                    xmlCiPart += "</PLUSPPRICESCHED>";
                }

            }
            #endregion


            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXCUSTAGREEMENT xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXCUSTAGREEMENTSet action=\"AddChange\">";
            xml += "<PLUSPAGREEMENT action=\"AddChange\">";
            xml += "<AGREEMENT>" + contract.IdContract + "</AGREEMENT>";
            //xml += "<DESCRIPTION>"+contract.Description+"</DESCRIPTION>";
            xml += "<ORGID>GBM</ORGID>";
            xml += "<REVISIONNUM>" + contract.Revision + "</REVISIONNUM>";
            xml += xmlCiPart;
            xml += "</PLUSPAGREEMENT>";
            xml += "</MXCUSTAGREEMENTSet>";
            xml += "</SyncMXCUSTAGREEMENT>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXCUSTAGREEMENT", xml);

            if (ValidateXml(responseText))
                return "OK";
            else
                return responseText;
        }

        #region Response Plans
        public string DeleteRpCommodity(string responsePlan, string pluspApplServId)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXRPDELETESERVICESICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXRPDELETESERVICESICSSet>";
            xml += @"<PLUSPAPPLSERV action=""Delete"">";
            xml += "<PLUSPAPPLSERVID>" + pluspApplServId + "</PLUSPAPPLSERVID>";
            xml += "<OWNERID>" + responsePlan + "</OWNERID>";
            xml += "<OWNERTABLE>PLUSPRESPPLAN</OWNERTABLE>";
            xml += "<ITEMSETID>ITEMSET1</ITEMSETID>";
            xml += "</PLUSPAPPLSERV>";
            xml += "</MXRPDELETESERVICESICSSet>";
            xml += "</SyncMXRPDELETESERVICESICS>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXRPDELETESERVICESICS", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Users
        public string InactivateUser(string user)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXPERSON xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"">";
            xml += "<MXPERSONSet>";
            xml += @"<PERSON action=""AddChange"">";
            xml += "<PERSONID>" + user.ToUpper() + "@GBM.NET</PERSONID>";
            xml += "<STATUS>INACTIVE</STATUS>";
            xml += "</PERSON>";
            xml += "</MXPERSONSet>";
            xml += "</SyncMXPERSON>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXPERSON", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        public string UpdateUserSupervisor(string user, string supervisor)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += "<SyncMXPERUSER xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\">";
            xml += "<MXPERUSERSet action=\"Change\">";
            xml += "<PERSON action=\"Change\">";
            xml += "<PERSONID>" + user.ToUpper() + "</PERSONID>";
            xml += "<SUPERVISOR>" + supervisor.ToUpper() + "</SUPERVISOR>";
            xml += "</PERSON></MXPERUSERSet></SyncMXPERUSER>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXPERUSER", xml);

            try
            {
                XmlDocument outXml = new XmlDocument();
                outXml.LoadXml(responseText);
                XmlNodeList tempList = outXml.GetElementsByTagName("PERSONID");

                string cambio = "";

                try
                {
                    cambio = tempList[0].Attributes.GetNamedItem("changed").Value;
                }
                catch (Exception)
                {
                    cambio = tempList[0].InnerText == user.ToUpper() ? "2" : "";
                }

                if (cambio == "1")
                {
                    //todo bien
                    return "OK";
                }
                else if (cambio == "2")
                {
                    return "SAME";
                }
                else
                {
                    return responseText;
                }
            }
            catch (XmlException)
            {
                //la respuesta no es un XML, probablemente error
                return responseText;
            }

        }
        #endregion
        #region Releases
        public string UpdateReleaseTargCompDate(string releaseID, string startDate)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += "<SyncMXRELEASE xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.ibm.com/maximo\" baseLanguage=\"EN\" transLanguage=\"EN\">";
            xml += "<MXRELEASESet>";
            xml += "<WORELEASE action=\"AddChange\">";
            xml += "<TARGCOMPDATE>" + startDate + "T08:00:00-06:00</TARGCOMPDATE>";
            xml += "<WONUM>" + releaseID + "</WONUM>";
            xml += "<SITEID>GBMHQ</SITEID>";
            xml += "</WORELEASE>";
            xml += "</MXRELEASESet>";
            xml += "</SyncMXRELEASE>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXRELEASE", xml);

            if (ValidateXml(responseText))
                return "OK";
            else
                return responseText;

        }
        #endregion
        #region Collection
        public string DeleteCollectionParty(string partyId)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXCHANGECOLLECTION xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXCHANGECOLLECTIONSet>";
            xml += "<GBMCOLLECTIONINTPARTY action=\"Delete\">";
            xml += "<GBMCOLLECTIONINTPARTYID>" + partyId + "</GBMCOLLECTIONINTPARTYID>";
            xml += "</GBMCOLLECTIONINTPARTY>";
            xml += "</MXCHANGECOLLECTIONSet>";
            xml += "</SyncMXCHANGECOLLECTION>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXCHANGECOLLECTION", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        public string AddCollectionParty(CdCollectionPartyData colParty, string collectionId)
        {

            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXCHANGECOLLECTION xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXCHANGECOLLECTIONSet>";
            xml += "<GBMCOLLECTIONINTPARTY action=\"Add\">";
            xml += "<COLLECTIONNUM>" + collectionId + "</COLLECTIONNUM>";
            xml += "<PERSONID>" + colParty.PersonId + "</PERSONID>";
            xml += "<PERSONGROUP>" + colParty.PersonGroup + "</PERSONGROUP>";
            xml += "<DESCRIPTION>" + colParty.Description + "</DESCRIPTION>";
            xml += "<TYPE>" + colParty.Type + "</TYPE>";
            xml += "</GBMCOLLECTIONINTPARTY>";
            xml += "</MXCHANGECOLLECTIONSet>";
            xml += "</SyncMXCHANGECOLLECTION>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXCHANGECOLLECTION", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        public string AddCollectionCis(CdCollectionData col)
        {
            #region collectDetailsPart
            string xmlCollectDetailsPart = "";
            foreach (string ci in col.Cis)
            {
                xmlCollectDetailsPart += "<COLLECTDETAILS>";
                xmlCollectDetailsPart += "<CINUM>" + ci + "</CINUM>";
                xmlCollectDetailsPart += "</COLLECTDETAILS>";
            }
            #endregion


            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXCOLLECTIONICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXCOLLECTIONICSSet>";
            xml += @"<COLLECTION action=""AddChange"">";
            xml += "<COLLECTIONNUM>" + col.CollectionNum + "</COLLECTIONNUM>";
            xml += xmlCollectDetailsPart;
            xml += "</COLLECTION>";
            xml += "</MXCOLLECTIONICSSet>";
            xml += "</SyncMXCOLLECTIONICS>";
            #endregion

            string responseText = PostCD(root.UrlCd, "MXCOLLECTIONICS", xml);

            if (ValidateXml(responseText))
                return "OK";
            else
                return responseText;
        }
        #endregion
        #region Response Plans
        public string DeleteRpCis(string responsePlan, string ciNum)
        {
            #region Service XML Request
            string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
            xml += @"<SyncMXRESPONSEPLANCISICS xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.ibm.com/maximo"" baseLanguage=""EN"" transLanguage=""EN"">";
            xml += "<MXRESPONSEPLANCISICSSet>";
            xml += @"<PLUSPAPPLCI action=""Delete"">";
            xml += "<CINUM>" + ciNum + "</CINUM>";
            xml += "<OWNERID>" + responsePlan + "</OWNERID>";
            xml += "<OWNERTABLE>PLUSPRESPPLAN</OWNERTABLE>";
            xml += "</PLUSPAPPLCI>";
            xml += "</MXRESPONSEPLANCISICSSet>";
            xml += "</SyncMXRESPONSEPLANCISICS>";

            #endregion

            string responseText = PostCD(root.UrlCd, "MXRESPONSEPLANCISICS", xml);

            if (ValidateXml(responseText) == true)
                return "OK";
            else
                return responseText;
        }
        public string ChangeResponsePlanStatus(string sanum, string status)
        {
            Dictionary<string, string> rpFields = new Dictionary<string, string> { { "STATUS", status } };

            CdResponsePlanData rpData = GetResponsePlanData(sanum);
            string pluspServAgreeId = rpData.PluspServAgreeId;

            string responseText = CallMaxRestPost(root.UrlCd, "MXRESPONSEPLANICS", pluspServAgreeId, rpFields);

            return responseText;
        }
        public string ChangeCommunicationTemplatesStatus(List<string> commTemplatesIds, string status)
        {
            Dictionary<string, string> rpFields = new Dictionary<string, string> { { "STATUS", status } };

            string responseText = "";

            foreach (string commTemplateId in commTemplatesIds)
            {
                responseText = CallMaxRestPost(root.UrlCd, "MXL_COMMTEMPLATEICS", commTemplateId, rpFields);

                if (responseText.ToUpper().Contains("ERROR"))
                    return responseText;
            }

            return responseText;
        }

        #endregion

        #endregion
    }
    public class CdTicketData
    {
        public string TicketType { get; set; }
        public string ReportedEmail { get; set; }
        public string Country { get; set; }
        public string ClassStructureId { get; set; }
        public string CommodityGroup { get; set; }
        public string Commodity { get; set; }
        public string PluspCustomer { get; set; }
        public string Description { get; set; }
        public string LongDescription { get; set; }
        public string Impact { get; set; }
        public string Urgency { get; set; }
        public string GbmPluspAgreement { get; set; }
        public string ExternalSystem { get; set; }
    }
    public class CdCollectionData
    {
        public string CollectionNum { get; set; }
        public string Description { get; set; }
        public string Supervisor { get; set; }
        public List<CdCollectionPartyData> CollectionParties { get; set; }
        public List<string> Cis { get; set; }
    }
    public class CdCollectionPartyData
    {
        public string Id { get; set; }
        public string PersonId { get; set; }
        public string Type { get; set; }
        public string PersonGroup { get; set; }
        public string Description { get; set; }
    }
    public class CdContractData
    {
        public string Status { get; set; }
        public string Revision { get; set; }
        public string IdContract { get; set; }
        public string IdCustomer { get; set; }
        public string Description { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string PluspAgreementId { get; set; }
        public List<string> MaterialArray { get; set; }
        public List<string> ManualServiceArray { get; set; }
        public List<string> EquipArray { get; set; }
        public List<string> CisArray { get; set; }
        public List<CdPriceScheduleData> PriceSchedules { get; set; }
    }


    public class CdPriceScheduleData
    {
        public string PriceSchedule { get; set; }
        public string SaNum { get; set; }
    }
    public class CdServicesData
    {
        public string Commodity { get; set; }
        public string PluspApplServId { get; set; }
        public string CommodityGroup { get; set; }
    }
    public class CdClassStructureData
    {
        public string ClassificationDesc { get; set; }
        public string Description { get; set; }
        public string HierarchyPath { get; set; }
    }
    public class CdConfigurationItemData
    {
        public string CiName { get; set; }
        public string CiNum { get; set; }
        public string PersonId { get; set; }
        public string Description { get; set; }
        public string Status { get; set; }
        public string PmcCiImpact { get; set; }
        public string CiLocation { get; set; }
        public string ClassStructureId { get; set; }
        public string ServiceGroup { get; set; }
        public string Service { get; set; }
        public string PluspCustomer { get; set; }
        public string GbmAdministrator { get; set; }
        public List<CdCiSpecData> CiSpecs { get; set; }
    }
    public class CdCiSpecData
    {
        public string AssetAttrId { get; set; }
        public string CCiSumSpecValue { get; set; }
    }
    public class CdContactData
    {
        public string PersonId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string SapId { get; set; }
        public string Email { get; set; }
        public string Telephone { get; set; }
    }
    public class CdCustomerData
    {
        public string PersonId { get; set; }
        public string IdCustomer { get; set; }
        public string NameCustomer { get; set; }
        public string Country { get; set; }
    }
    public class CdLocationData
    {
        public string Location { get; set; }
        public string Description { get; set; }
    }
    public class CdAssetData
    {
        public string AssetNum { get; set; }
        public string AssetText { get; set; }
        public string Location { get; set; }
        public string SerialNum { get; set; }
        public string MaterialNum { get; set; }
        public string InstallDate { get; set; }
        public string EndDate { get; set; }
        public string Warranty { get; set; }
        public string WarrantyText { get; set; }
        public string MaterialGroup { get; set; }
        public string Placa { get; set; }

    }
    public class CdReleaseData
    {
        public string Description { get; set; }
        public string PluspCustomer { get; set; }
        public string Classification { get; set; }
        public string Commodity { get; set; }
        public string CommodityGroup { get; set; }
        public string Environment { get; set; }
        public string PmRelEmergency { get; set; }
        public string PmRelImpact { get; set; }
        public string PmRelUrgency { get; set; }
        public string WoPriority { get; set; }
        public string TargStartDate { get; set; }
        public string TargCompDate { get; set; }
        public string Employee { get; set; }
        public string Contact { get; set; }
        public string Contract { get; set; }
        public string ConRev { get; set; }
        public string OwnerGroup { get; set; }
        public string Owner { get; set; }
        public string ExtRef { get; set; }
        public string relId { get; set; }
    }
    public class CdUserData
    {
        public string User { get; set; }
        public string[] Roles { get; set; }
    }
    public class CdResponsePlanData
    {
        public string Sanum { get; set; }
        public string Description { get; set; }
        public string Ranking { get; set; }
        public string ObjectName { get; set; }
        public string OwnerGroup { get; set; }
        public string Calendar { get; set; }
        public string Shift { get; set; }
        public string Status { get; set; }
        public string GbmAutoAssignment { get; set; }
        public string CalendarOrgId { get; set; }
        public string AssignOwnerGroup { get; set; }
        public string CustomerId { get; set; }
        public string Action { get; set; }
        public string ClassStructureId { get; set; }
        public string Condition { get; set; }
        public string Escalation { get; set; }
        public string PluspServAgreeId { get; set; }
        public List<string> ConfigurationItems { get; set; }
        public CdServicesData[] Services { get; set; }
        public CdClassStructureData ClassStructure { get; set; }
        public string ResponseMessage { get; set; }
    }
    public class CdCommTemplate
    {
        public string TemplateId { get; set; }
        public string Description { get; set; }
        public string ObjectName { get; set; }
        public string Subject { get; set; }
        public string Message { get; set; }
        public List<string> CommTmpltSendToValue { get; set; }
    }
    public class CdPersonGroupData
    {
        public List<string> Emails { get; set; }
        public string PersonGroupId { get; set; }
    }
    public class CdSlaData
    {
        public string Sanum { get; set; }
        public string ObjectName { get; set; }
        public string Ranking { get; set; }
        public string Description { get; set; }
        public string IntPriorityEval { get; internal set; }
        public string IntPriorityValue { get; internal set; }
        public string CalcOrgId { get; set; }
        public string CalcCalendar { get; set; }
        public string CalcShift { get; set; }
        public string Condition { get; set; }
        public string Escalation { get; set; }
        public List<string> PluspApplServCommodity { get; set; }
        public List<CdSlaEscalation> CdSlaEscalations { get; set; }
        public List<CdSlaCommitments> CdSlaCommitments { get; set; }
    }
    public class CdSlaCommitments
    {
        public string Description { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
        public string UnitOfMeasure { get; set; }
    }
    public class CdSlaEscalation
    {
        public string TimeAttribute { get; set; }
        public string TimeInterval { get; set; }
        public string IntervalUnit { get; set; }
        public string Condition { get; set; }
        public List<CdCommTemplate> Notifications { get; set; }
    }
    public class CdCommunicationTemplatesData
    {
        public string Description { get; set; }
        public string Subject { get; set; }
        public string Message { get; set; }
        public List<string> Recipients { get; set; }
        public string ObjectName { get; set; }
    }
}