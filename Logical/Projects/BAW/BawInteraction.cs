using DataBotV5.Automation.ICS.BawUsers;
using DataBotV5.Data.Credentials;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using DataBotV5.Data.Root;
using Newtonsoft.Json;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System;

namespace DataBotV5.Logical.Projects.BAW
{
    /// <summary>
    /// Clase Logical encargada de interacción con Fusion Auth.
    /// </summary>
    class BawInteraction
    {
        readonly Rooting root = new Rooting();
        readonly Credentials cred = new Credentials();
        const string bawAdminUser = "deadmin";

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
        private string CallBawApi(string destinationUrl, string requestJson, string csrfToken = "")
        {
            string ret = "ERROR";

            try
            {
                WebClient client = new WebClient();
                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(bawAdminUser + ":" + cred.BawAdmin));

                client.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                client.Headers[HttpRequestHeader.ContentType] = "application/json";
                client.Headers[HttpRequestHeader.Accept] = "application/json";

                if (csrfToken != "")
                    client.Headers["BPMCSRFToken"] = csrfToken;

                if (requestJson == "")
                    ret = client.DownloadString(destinationUrl);
                else
                    ret = client.UploadString(destinationUrl, "POST", requestJson);
            }
            catch (WebException ex)
            {
                // Leer la respuesta del servidor en caso de error
                if (ex.Response != null)
                {
                    using (HttpWebResponse errorResponse = (HttpWebResponse)ex.Response)
                    using (StreamReader reader = new StreamReader(errorResponse.GetResponseStream()))
                    {
                        string errS = reader.ReadToEnd();
                        if (errS != "")
                            ret = errS;
                        else
                            ret = "ERROR: " + ex.Message;
                    }
                }
                else
                    ret = "ERROR: " + ex.Message;
            }
            return ret;
        }

        #region Métodos creación de datos en BAW
        /// <summary>
        /// copia los roles de un refUser a un newUser de un proceso
        /// </summary>
        /// <param name="newUser"></param>
        /// <param name="refUser"></param>
        /// <param name="processData"></param>
        /// <returns></returns>
        internal List<string> AddUserToGroup(string newUser, string refUser, GetProcessData processData)
        {
            List<string> ret = new List<string>();
            string[] userGroups = GetProcessUsers(refUser, processData);

            foreach (string userGroup in userGroups)
            {
                string url = root.UrlBaw + "/ops/std/bpm/containers/" + processData.Container + "/versions/" + processData.Version + "/team_bindings/" + userGroup;
                string requestJson = "{\"add_users\":[\"" + newUser + "\"]}";
                string apiResponse = CallBawApi(url, requestJson, root.TokenBaw);

                #region Process Response

                JObject parsedData = JObject.Parse(apiResponse);

                foreach (JToken team in parsedData["team_bindings"])
                {
                    if (team["name"].ToString() == userGroup)
                    {
                        foreach (JToken member in team["user_members"].ToArray())
                        {
                            if (member.ToString() == newUser)
                            {
                                ret.Add(userGroup);
                                break;
                            }
                        }
                    }
                }

                #endregion
            }

            return ret;
        }
        #endregion

        #region Métodos de consulta
        /// <summary>
        /// obtiene un Token para usar el API de BAW
        /// </summary>
        /// <returns></returns>
        internal string GetBawApiToken()
        {
            string url = root.UrlBaw + "/ops/system/login";
            string requestJson = "{\"refresh_groups\":true,\"requested_lifetime\":60}";
            string csrfToken = CallBawApi(url, requestJson);

            #region Process Response
            try
            {
                JToken outJson = JToken.Parse(csrfToken);
                try { csrfToken = outJson["csrf_token"].ToString(); } catch (Exception) { }
            }
            catch (JsonReaderException) { }  //la respuesta no es un Json, probablemente error   
            #endregion

            return csrfToken;
        }
        /// <summary>
        /// establece la variable root.TokenApi para usar el API de BAW
        /// </summary>
        /// <returns></returns>
        internal bool SetBawApiToken()
        {
            bool ret = false;
            string url = root.UrlBaw + "/ops/system/login";
            string requestJson = "{\"refresh_groups\":true,\"requested_lifetime\":180}";
            string csrfToken = CallBawApi(url, requestJson);

            #region Process Response
            if (!csrfToken.ToUpper().Contains("ERROR"))
            {
                try
                {
                    JToken outJson = JToken.Parse(csrfToken);
                    root.TokenBaw = outJson["csrf_token"].ToString();
                    ret = true;
                }
                catch (Exception) { }  //la respuesta no es un Json, probablemente error   
            }
            #endregion

            return ret;
        }
        /// <summary>
        /// Obtiene los roles que tiene un usuario en un proceso 
        /// </summary>
        /// <param name="refUser"></param>
        /// <param name="container"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        internal string[] GetProcessUsers(string refUser, GetProcessData processData)
        {
            List<string> userGroups = new List<string>();

            string url = root.UrlBaw + "/ops/std/bpm/containers/" + processData.Container + "/versions/" + processData.Version + "/team_bindings";

            string apiResponse = CallBawApi(url, "", root.TokenBaw);

            #region Process Response
            JToken responseJson = JToken.Parse(apiResponse);
            JToken groups = responseJson["team_bindings"];

            if (groups != null)
            {
                foreach (JToken group in groups)
                {
                    string groupName = group["name"].ToString();
                    JToken groupMembers = group["group_members"];
                    JToken usersMembers = group["user_members"];
                    foreach (JToken groupMember in groupMembers)
                    {
                        string member = groupMember.ToString();
                        if (member == refUser)
                            userGroups.Add(groupName);
                    }
                    foreach (JToken userMembers in usersMembers)
                    {
                        string member = userMembers.ToString();
                        if (member == refUser)
                            userGroups.Add(groupName);
                    }
                }
            }
            else
            {
                string error = responseJson["error_message"].ToString();
                throw new Exception(error);
            }
            #endregion

            return userGroups.ToArray();
        }
        #endregion

        #region Métodos de Modificación
        //vacío
        #endregion

    }
}

