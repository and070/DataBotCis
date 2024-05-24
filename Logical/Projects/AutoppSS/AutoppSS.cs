
using DataBotV5.App.Global;
using DataBotV5.Logical.Mail;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web;

namespace DataBotV5.Logical.Projects.AutoppSS
{
    /// <summary>
    /// En esta sección lógica de la Fábrica de Propuestas,
    /// se administra la parte de notificaciones con card através de Webex Teams.
    /// Coded by: Eduardo Piedra Sanabria - Application Management Analyst
    /// /// </summary>
    class AutoppSS : IDisposable
    {

        MailInteraction mail = new MailInteraction();


        #region Cards
        //Card template para notificaciones.
        public static string salesTeamCard = @"

{
    ""toPersonEmail"": ""CORREO"",
    ""text"": ""TEXTONOTIFICACION"",
    ""attachments"": [
     {
          ""contentType"": ""application/vnd.microsoft.card.adaptive"",
          ""content"": {
            ""type"": ""AdaptiveCard"",
            ""body"": [

              {
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""items"": [
                      {
                        ""type"": ""Image"",
                        ""url"": ""https://databot.ngrok.io/ext/Automation.png"",
                        ""spacing"": ""Medium"",
                        ""size"": ""Medium"",
                        ""height"": ""50px""
                      }
                    ],
                    ""width"": ""auto""
                  },
                  {
                    ""type"": ""Column"",
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Management Information Systems"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Accent"",
                        ""horizontalAlignment"": ""Start""

                      },
                      {
                        ""type"": ""TextBlock"",
                        ""weight"": ""Bolder"",
                        ""text"": ""TITLENOTIFICATION"",
                        ""wrap"": true,
                        ""color"": ""Light"",
                        ""size"": ""Large"",
                        ""spacing"": ""Small""
                      }
                    ],
                    ""width"": ""stretch""
                  }
                ]
              },

            {
                ""type"": ""TextBlock"",
                ""text"": ""MENSAJE"",
                ""wrap"": true,
                ""color"": ""Light""

            },
              {
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""width"": 35,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Portal:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Cliente:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Oportunidad:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Tipo de oportunidad:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Rol asignado:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Empleado responsable:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      }
                    ]
                  },
                  {
                    ""type"": ""Column"",
                    ""width"": 65,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""PORTAL"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""CLIENT"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""OPP"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""TYPEOPPORTUNITY"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""TYPEROLE"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""EMPLOYEERESPONSIBLE"",
                        ""color"": ""Light""
                      }
                    ]
                  }
                ],
                ""spacing"": ""Padding"",
                ""horizontalAlignment"": ""Center""
              },

            {
                ""type"": ""ActionSet"",
                ""actions"": [
                    {
                        ""type"": ""Action.OpenUrl"",
                        ""title"": ""Ir al portal de S&S"",
                        ""url"": ""LINK""
                    }
                            ],
                ""horizontalAlignment"": ""Center""
            }
            ],
            ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
            ""version"": ""1.2""
          }
        }
    ]
  }

";

        public static string successCard = @"

{
    ""toPersonEmail"": ""CORREO"",
    ""text"": ""TEXTONOTIFICACION"",
    ""attachments"": [
     {
          ""contentType"": ""application/vnd.microsoft.card.adaptive"",
          ""content"": {
            ""type"": ""AdaptiveCard"",
            ""body"": [

              {
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""items"": [
                      {
                        ""type"": ""Image"",
                        ""url"": ""https://databot.ngrok.io/ext/Automation.png"",
                        ""spacing"": ""Medium"",
                        ""size"": ""Medium"",
                        ""height"": ""50px""
                      }
                    ],
                    ""width"": ""auto""
                  },
                  {
                    ""type"": ""Column"",
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Management Information Systems (MIS)"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Accent"",
                        ""horizontalAlignment"": ""Start""

                      },
                      {
                        ""type"": ""TextBlock"",
                        ""weight"": ""Bolder"",
                        ""text"": ""TITLENOTIFICATION"",
                        ""wrap"": true,
                        ""color"": ""Light"",
                        ""size"": ""Large"",
                        ""spacing"": ""Small""
                      }
                    ],
                    ""width"": ""stretch""
                  }
                ]
              },

            {
                ""type"": ""TextBlock"",
                ""text"": ""MENSAJE"",
                ""wrap"": true,
                ""color"": ""Light""

            },
              {
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""width"": 35,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Portal:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Cliente:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Oportunidad:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Tipo de oportunidad:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Empleado responsable:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      }
                    ]
                  },
                  {
                    ""type"": ""Column"",
                    ""width"": 65,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""PORTAL"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""CLIENT"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""OPP"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""TYPEOPPORTUNITY"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""EMPLOYEERESPONSIBLE"",
                        ""color"": ""Light""
                      }
                    ]
                  }
                ],
                ""spacing"": ""Padding"",
                ""horizontalAlignment"": ""Center""
              },

            {
                ""type"": ""ActionSet"",
                ""actions"": [
                    {
                        ""type"": ""Action.OpenUrl"",
                        ""title"": ""Ir al portal de S&S"",
                        ""url"": ""LINK""
                    }
                            ],
                ""horizontalAlignment"": ""Center""
            }
            ],
            ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
            ""version"": ""1.2""
          }
        }
    ]
  }

";

        #endregion

        private bool disposedValue;



        //Link de Smart & Simple
        string linkSS = "https://smartsimple.gbm.net/admin/AutoppLdrs/autoppstart";

        //Diccionario de mensajes prestablecidos
        Dictionary<string, string> words = new Dictionary<string, string>(){
            {"salesTeamsNotification", "Estimado(a) USER, se le informa que ha sido agregado al Sales Team de la siguiente oportunidad: "},
            {"successNotification", "Estimado(a) USER, se le informa que ha sido generada exitosamente la siguiente oportunidad de venta: "},
        };


        /// <summary>
        /// Realiza una notificación al user de los parámetros, en base al tipo de notificación preestablecida, además se le 
        /// aporta un diccionario toReplace que indican las palabras que se le deben sustituir al card.
        /// </summary>
        /// <returns>Void</returns>
        public void AutoppNotifications(string typeNotification, string user, Dictionary<string, string> toReplace, string notificationsConfig)
        {
            string card = "";

            if (typeNotification == "salesTeamsNotification")
            {
                card = salesTeamCard;
            }
            else if (typeNotification == "successNotification")
            {
                card = successCard;
            }


            card = card.Replace("TEXTONOTIFICACION", "Fábrica de propuestas");
            card = card.Replace("CORREO", notificationsConfig == "admin" ? "EPIEDRA@GBM.NET" : (user + "@GBM.NET"));
            card = card.Replace("PORTAL", "Fábrica de Propuestas - Smart & Simple");
            card = card.Replace("AREA", "Ventas");
            //card = card.Replace("USUARIO", "user");
            card = card.Replace("MENSAJE", words[typeNotification]);  //GetMessage(typeNotification, toReplace)) ; // words[typeNotification]);
            card = card.Replace("LINK", linkSS);

            card = ReplaceWordsCard(card, toReplace);



            string url = "https://webexapis.com/v1/messages";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(card);
            }
            try
            {

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                }

            }
            catch (Exception e)
            {

                mail.SendHTMLMail(card, new string[] { "epiedra@gbm.net" }, $"No se pudo notificar por webex");

            }

        }


        /// <summary>
        /// Replaza las palabras clave por ejemplo OPP del card, con cada uno de los ítems de diccionario. Ej: ("OPP", 000001632);
        /// </summary>
        /// <returns>Retorna un string con el card con palabras reemplazadas.</returns>
        public string ReplaceWordsCard(string card, Dictionary<string, string> toReplace)
        {

            foreach (KeyValuePair<string, string> replace in toReplace)
            {
                card = card.Replace(replace.Key, replace.Value);
            }

            return card;
        }

        private bool ValidateWebexMail(string email)
        {
            bool validate = false;
            string url = $"https://webexapis.com/v1/people?email={HttpUtility.UrlEncode(email)}";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "GET";
            httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                //var definition = new { Items = "" };
                var result = streamReader.ReadToEnd();
                ValidateMail items_json = JsonConvert.DeserializeObject<ValidateMail>(result);
                if (items_json.items.Count > 0)
                {
                    validate = true;
                }

            }
            return validate;
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
        // ~WebexTeams()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
    public class MensajeNormal
    {
        public string roomId { get; set; }
        public string text { get; set; }
        public string markdown { get; set; }
    }
    public class MensajeNormalCC
    {
        public string toPersonEmail { get; set; }
        public string text { get; set; }
        public string markdown { get; set; }
    }
    public class MensajeNormalEmail
    {
        public string toPersonEmail { get; set; }
        public string text { get; set; }
        public string markdown { get; set; }
    }
    public class CardMessageWebex
    {
        public string toPersonEmail { get; set; }
        public string text { get; set; }
        public string attachments { get; set; }
        public string markdown { get; set; }
    }
    public class ValidateMail
    {
        public string notFoundIds { get; set; }
        public List<ValidateMailItems> items { get; set; }
    }
    public class ValidateMailItems
    {
        public string id { get; set; }
    }
}

