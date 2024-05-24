using DataBotV5.App.Global;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web;

namespace DataBotV5.Logical.Webex
{
    class WebexTeams:IDisposable
    {
        public static string card_carta_condiciones = @"

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
                        ""text"": ""Data Management & Automation"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Accent""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""weight"": ""Bolder"",
                        ""text"": ""Notificación carta de posición"",
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
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""width"": 35,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Portal:"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Area:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Solicitante:"",
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
                        ""text"": ""AREA"",
                        ""color"": ""Light"",
                        ""weight"": ""Lighter"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""USUARIO"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      }
                    ]
                  }
                ],
                ""spacing"": ""Padding"",
                ""horizontalAlignment"": ""Center""
              },
            {
                ""type"": ""TextBlock"",
                ""text"": ""MENSAJE"",
                ""wrap"": true
            },
            {
                ""type"": ""ActionSet"",
                ""actions"": [
                    {
                        ""type"": ""Action.OpenUrl"",
                        ""title"": ""Descargar archivo"",
                        ""url"": ""LINK""
                    }
                            ]
            }
            ],
            ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
            ""version"": ""1.2""
          }
        }
    ]
  }

";
        public static string card_errores_databot = @"

{
    ""roomId"": ""ROOM"",
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
                        ""text"": ""Data Management & Automation"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Accent""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""weight"": ""Bolder"",
                        ""text"": ""Notificación errores del Databot"",
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
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""width"": 35,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Clase:"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Linea:"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Columna:"",
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
                        ""text"": ""CLASE"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""LINEA"",
                        ""color"": ""Light"",
                        ""weight"": ""Lighter"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""COLUMNA"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      }
                    ]
                  }
                ],
                ""spacing"": ""Padding"",
                ""horizontalAlignment"": ""Center""
              },
            {
                ""type"": ""TextBlock"",
                ""text"": ""MENSAJE"",
                ""wrap"": true
            }            
            ],
            ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
            ""version"": ""1.2""
          }
        }
    ]
  }

";
        public static string card_notificaciones_fabrica = @"

{
    ""toPersonEmail"": ""EMAIL"",
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
                        ""text"": ""Data Management & Automation"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Accent""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""weight"": ""Bolder"",
                        ""text"": ""TITULO"",
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
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""width"": 35,
                    ""items"": [
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Portal:"",
                        ""color"": ""Light""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""Proceso:"",
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
                        ""text"": ""PROCESO"",
                        ""color"": ""Light"",
                        ""weight"": ""Lighter"",
                        ""spacing"": ""Medium""
                      },
                      {
                        ""type"": ""TextBlock"",
                        ""text"": ""OPORTUNIDAD"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      }
                    ]
                  }
                ],
                ""spacing"": ""Padding"",
                ""horizontalAlignment"": ""Center""
              },
            {
                ""type"": ""TextBlock"",
                ""text"": ""MENSAJE"",
                ""wrap"": true
            }            
            ],
            ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
            ""version"": ""1.2""
          }
        }
    ]
  }

";
        private bool disposedValue;

        public void SendLetterHCM(string user, string link)
        {
            string card = card_carta_condiciones;
            card = card.Replace("TEXTONOTIFICACION", "Nueva carta");
            card = card.Replace("CORREO", user + "@GBM.NET");
            card = card.Replace("PORTAL", "Carta de posiciones");
            card = card.Replace("AREA", "HCM");
            card = card.Replace("USUARIO", user);
            card = card.Replace("MENSAJE", "Estimado(a) " + user + " se ha generado una nueva carta de posiciones desde el portal DM & Automation, haz click en el botón para descargar el archivo, recuerda tener el VPN activado en caso de no estar en la red de GBM.");
            card = card.Replace("LINK", link);
            //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            //Credenciales cred = new Credenciales();
            string url = "https://webexapis.com/v1/messages";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(card);
            }
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
            }
        }
        public void SendCCNotification(string email, string title, string process, string opportunity, string message)
        {

            MensajeNormalCC mens = new MensajeNormalCC();
            mens.toPersonEmail = email;
            mens.text = title;
            mens.markdown = message;

            string json = JsonConvert.SerializeObject(mens);
            //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            //Credenciales cred = new Credenciales();
            string url = "https://webexapis.com/v1/messages";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
            }
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
            }
        }
        public void SendNotification(string email, string title, string message)
        {
            try
            {
                if (ValidateWebexMail(email))
                {
                    MensajeNormalEmail mens = new MensajeNormalEmail();
                    mens.toPersonEmail = email;
                    mens.text = title;
                    mens.markdown = message;
                   

                    string json = JsonConvert.SerializeObject(mens);
                    //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    //Credenciales cred = new Credenciales();
                    string url = "https://webexapis.com/v1/messages";
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);
                    }
                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                    }
                }         
            }
            catch (Exception ex)
            {
                try
                {

                    new ConsoleFormat().WriteLine(ex.ToString());
                    NotificationErrors(ex + "<br><br>" + email + "<br><br>" + message, 0, 0, "", "Y2lzY29zcGFyazovL3VzL1JPT00vNDk4MjM4NzAtMzFiMy0xMWViLTk3ZjAtYzVjODdmZTg4ZjE3", "Error al enviar notificación");

                }
                catch (Exception)
                {

                }
            }
        }
        public void SendNotificationCard(string email, string title, string message)
        {
            try
            {
                if (ValidateWebexMail(email))
                {
                    CardMessageWebex mens = new CardMessageWebex();
                    mens.toPersonEmail = email;
                    mens.text = title;
                    mens.attachments = message;
                    mens.markdown = "";


                    string json = JsonConvert.SerializeObject(mens);
                    //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    //Credenciales cred = new Credenciales();
                    string url = "https://webexapis.com/v1/messages";
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);
                    }
                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {

                    new ConsoleFormat().WriteLine(ex.ToString());
                    NotificationErrors(ex + "<br><br>" + email + "<br><br>" + message, 0, 0, "", "Y2lzY29zcGFyazovL3VzL1JPT00vNDk4MjM4NzAtMzFiMy0xMWViLTk3ZjAtYzVjODdmZTg4ZjE3", "Error al enviar notificación");

                }
                catch (Exception)
                {

                }
            }
        }

        public void NotificationErrors(string message, int line, int column, string @class, string sala, string notificationText)
        {
            MensajeNormal mens = new MensajeNormal();
            mens.roomId = sala;
            mens.text = notificationText;
            mens.markdown = message;

            string json = JsonConvert.SerializeObject(mens);
            string url = "https://webexapis.com/v1/messages";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.Headers.Add("Authorization", "Bearer NzZhNDE5MzQtMTNkMS00M2Q4LThjNWMtNzg5MTBjOTU1YTM1NjU4OTNmOWQtOGE0_PF84_a91d9855-d761-4b0a-a9e5-7d2345410a6d");
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
            }
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
            }
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
