using DataBotV5.Logical.Webex;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.Logical.Projects.MasterData
{
    
    public class MasterDataLogical
    {
        readonly public string[] erroresDatosMaestros = { "smarin@gbm.net", "hlherrera@gbm.net" };
        public string json_string_solicitante = @"

            {
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
                            ""style"": ""Person"",
                            ""url"": ""https://scontent.fsjo7-1.fna.fbcdn.net/v/t1.0-9/99109328_10207263323777217_132984462500691968_n.jpg?_nc_cat=100&_nc_sid=0debeb&_nc_ohc=1kv2Cv5_hP4AX9BqyHo&_nc_ht=scontent.fsjo7-1.fna&oh=66bdebc1745b91f9c0e9c445f034c475&oe=5EEDED63"",
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
                            ""text"": ""Notificacion de gestión #XXX"",
                            ""horizontalAlignment"": ""Left"",
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
                            ""text"": ""Fecha:"",
                            ""color"": ""Light""
                        },
                        {
                            ""type"": ""TextBlock"",
                            ""text"": ""Dato maestro:"",
                            ""weight"": ""Lighter"",
                            ""color"": ""Light"",
                            ""spacing"": ""Medium""
                        },
                        {
                            ""type"": ""TextBlock"",
                            ""text"": ""Tipo gestión:"",
                            ""weight"": ""Lighter"",
                            ""color"": ""Light"",
                           ""spacing"": ""Medium""
                        },
{
                            ""type"": ""TextBlock"",
                            ""text"": ""TIPOFACTOR:"",
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
                            ""text"": ""FECHACTUAL"",
                            ""color"": ""Light""
                        },
                        {
                            ""type"": ""TextBlock"",
                            ""text"": ""DATOMAESTRO"",
                            ""color"": ""Light"",
                            ""weight"": ""Lighter"",
                           ""spacing"": ""Medium""
                        },
                        {
                            ""type"": ""TextBlock"",
                            ""text"": ""TIPOGESTION"",
                            ""weight"": ""Lighter"",
                            ""color"": ""Light"",
                            ""spacing"": ""Medium""
                        },
 {
                            ""type"": ""TextBlock"",
                            ""text"": ""FACTOR"",
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
            ""type"": ""RichTextBlock"",
            ""inlines"": [
                {
                    ""type"": ""TextRun"",
                    ""text"": ""APROBADORES""
                }
                
            ]
            
        },
        {
            ""type"": ""ActionSet"",
            ""actions"": [
                {
                    ""type"": ""Action.OpenUrl"",
                    ""title"": ""Ir a mis gestiones"",
                    ""url"": ""https://databot.gbm.net/automation/dm/gestiones""
                }
            ],
            ""horizontalAlignment"": ""Left"",
            ""spacing"": ""None""
        }
    ],
    ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
    ""version"": ""1.2""
}

";
        public string jsonCardStyle = @"[
        {
          ""contentType"": ""application/vnd.microsoft.card.adaptive"",
          ""content"": {
            ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
            ""type"": ""AdaptiveCard"",
            ""version"": ""1.0"",
            ""body"": [
              {
                ""type"": ""ColumnSet"",
                ""columns"": [
                  {
                    ""type"": ""Column"",
                    ""items"": [
                      {
                        ""type"": ""Image"",
                        ""style"": ""Person"",
                        ""url"": ""https://scontent.fsjo7-1.fna.fbcdn.net/v/t1.0-9/99109328_10207263323777217_132984462500691968_n.jpg?_nc_cat=100&_nc_sid=0debeb&_nc_ohc=1kv2Cv5_hP4AX9BqyHo&_nc_ht=scontent.fsjo7-1.fna&oh=66bdebc1745b91f9c0e9c445f034c475&oe=5EEDED63"",
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
                        ""text"": ""Application Management"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Accent""
                      },
                      {
        ""type"": ""TextBlock"",
                        ""weight"": ""Bolder"",
                        ""text"": ""Notificacion de gestión {GESTION}"",
                        ""horizontalAlignment"": ""Left"",
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
                        ""text"": ""Fecha: "",
                        ""color"": ""Light""
                      },
                      {
            ""type"": ""TextBlock"",
                        ""text"": ""Dato maestro: "",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
            ""type"": ""TextBlock"",
                        ""text"": ""Tipo gestión: "",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
            ""type"": ""TextBlock"",
                        ""text"": ""{FACTOR}"":,
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
                        ""text"": ""{FECHA}"",
                        ""color"": ""Light""
                      },
                      {
            ""type"": ""TextBlock"",
                        ""text"": ""{FORMULARIO}"",
                        ""color"": ""Light"",
                        ""weight"": ""Lighter"",
                        ""spacing"": ""Medium""
                      },
                      {
            ""type"": ""TextBlock"",
                        ""text"": ""{TIPOGESTION}"",
                        ""weight"": ""Lighter"",
                        ""color"": ""Light"",
                        ""spacing"": ""Medium""
                      },
                      {
            ""type"": ""TextBlock"",
                        ""text"": ""{FACTOR}"",
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
                ""text"": ""{MSJ}"",
                ""wrap"": true
              },

              {
    ""type"": ""TextBlock"",
                ""text"": ""{SENDER}"",
                ""wrap"": true
              },
            ],
            ""actions "": [
              {
                ""type"": ""Action.OpenUrl"",
                ""title"": ""Ir a mis gestiones"",
                ""url"": ""https://smartsimple.gbm.net/"",

              }
            ],
            ""horizontalAlignment "": ""Left "",
            ""spacing "": ""None ""
          }
        }
      ]";
public bool sendNotification(string idRequest, string user, string status, string message, string form, string typeOfManagement, string factorType, string factorValue)
        {
            jsonCardStyle = jsonCardStyle.Replace("{GESTION}", idRequest);
            jsonCardStyle = jsonCardStyle.Replace("{FECHA}", DateTime.Now.ToString() + " GMT -6");
            jsonCardStyle = jsonCardStyle.Replace("{FORMULARIO}", form);
            jsonCardStyle = jsonCardStyle.Replace("{TIPOGESTION}", typeOfManagement);
            //jsonCardStyle = jsonCardStyle.Replace("TIPOFACTOR", factorType);
            jsonCardStyle = jsonCardStyle.Replace("{FACTOR}", factorValue);
            jsonCardStyle = jsonCardStyle.Replace("{MSJ}", message);
            jsonCardStyle = jsonCardStyle.Replace("{SENDER}", user);
            using (WebexTeams wb = new WebexTeams())
            {
                wb.SendNotificationCard(user, "", jsonCardStyle);
            }
            return true;
        }
    }
}
