using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Root;
using System.Data;
using System.Linq;
using System.Text;
using System.Net;
using System.Xml;
using System.IO;
using System;
using Newtonsoft.Json;

namespace DataBotV5.Logical.Projects.UserSAM
{
    /// <summary>
    /// Clase Logical encargada de interacción con Fusion Auth.
    /// </summary>
    class FusionAuthInteraction
    {
        readonly Credentials cred = new Credentials();
        readonly string urlFusionAuth = "https://gbm.fusionauth.io";

        private bool ValidateJson(string json)
        {
            try
            {
                JToken.Parse(json);
                return true;
            }
            catch (JsonReaderException)
            {
                return false;
            }
        }
        private string CallFaApi(string destinationUrl, string requestJson, string apiKey, string method)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(requestJson);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationUrl);
            request.Method = method;
            request.Headers.Add("Authorization", apiKey);

            if (method != "GET")
            {
                request.ContentType = "application/json ";
                request.ContentLength = bytes.Length;

                Stream requestStream = request.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();
            }


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

        #region Métodos creación de datos en Fusion Auth
        public string CreateUser(FaUserData userData)
        {
            string joinedRoles = string.Join(",", userData.Registration.Roles.Select(r => $"\"{r}\""));
            #region JSON

            string json = "{";
            json += "\"registration\": {";
            json += "\"applicationId\": \"" + userData.Registration.ApplicationId + "\",";
            json += "\"roles\": [ " + joinedRoles + " ],";
            json += "\"username\": \"" + userData.Registration.UserName + "\"";
            json += "},";
            json += "\"user\": {";
            json += "\"email\": \"" + userData.Email + "\",";
            json += "\"firstName\": \"" + userData.FirstName + "\",";
            json += "\"fullName\": \"" + userData.FullName + "\",";
            json += "\"lastName\": \"" + userData.LastName + "\",";
            json += "\"password\": \"" + userData.Password + "\",";
            json += "\"username\": \"" + userData.UserName + "\"";
            json += "}";
            json += "}";

            #endregion

            string responseText = CallFaApi(urlFusionAuth + "/api/user/registration/", json, cred.fusionAuthApiKey, "POST");

            #region Process Response
            try
            {
                JToken outJson = JToken.Parse(responseText);

                try
                {
                    responseText = outJson["user"]["id"].ToString();
                    if (responseText != "")
                        responseText = "OK";
                }
                catch (Exception) { }

            }
            catch (JsonReaderException) { }  //la respuesta no es un Json, probablemente error   
            #endregion

            return responseText;
        }
        #endregion

        #region Métodos de Modificación
        public string ChangeUserPass(FaUserData userData)
        {
            #region JSON

            string json = "{";
            json += "\"user\": {";
            json += "\"password\": \"" + userData.Password + "\"";
            json += "}";
            json += "}";

            #endregion

            string responseText = CallFaApi(urlFusionAuth + "/api/user/" + userData.Id, json, cred.fusionAuthApiKey, "PATCH");

            #region Process Response
            try
            {
                JToken outJson = JToken.Parse(responseText);

                try
                {
                    responseText = outJson["user"]["id"].ToString();
                    if (responseText != "")
                        responseText = "OK";
                }
                catch (Exception) { }

            }
            catch (JsonReaderException) { }  //la respuesta no es un Json, probablemente error   
            #endregion

            return responseText;
        }
        #endregion

        #region Métodos de consulta
        public string GetUserId(FaUserData userData)
        {
            string responseText = CallFaApi(urlFusionAuth + "/api/user?username=" + userData.UserName.ToLower(), "", cred.fusionAuthApiKey, "GET");

            #region Process Response
            try
            {
                JToken outJson = JToken.Parse(responseText);

                try
                {
                    responseText = outJson["user"]["id"].ToString();
                }
                catch (Exception) { }

            }
            catch (JsonReaderException) { }  //la respuesta no es un Json, probablemente error   
            #endregion

            return responseText;
        }
        #endregion
    }
}

public class FaUserData
{
    public string Id { get; set; }
    public string Email { get; set; }
    public string FirstName { get; set; }
    public string FullName { get; set; }
    public string LastName { get; set; }
    public string Password { get; set; }
    public string UserName { get; set; }
    public FaRegistrationData Registration { get; set; }
}
public class FaRegistrationData
{
    public string ApplicationId { get; set; }
    public string UserName { get; set; }
    public List<string> Roles { get; set; }
}