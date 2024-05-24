using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Microsoft.Graph.Core;
using Microsoft.Identity.Client;
//using Microsoft.Graph.Models;
using Azure.Identity;
using Azure.Core;
using DataBotV5.Data.Credentials;

namespace DataBotV5.Logical.MicrosoftTools
{
    /// <summary>
    /// 
    /// </summary>
    class MicrosoftTeams : IDisposable
    {
        private bool disposedValue;
        private object scopes;
        Credentials cred = new Credentials();

        private GraphServiceClient InitializeGraphClient()
        {
            string authority = $"https://login.microsoftonline.com/{cred.tenantId}";
            string[] scopes = { "https://graph.microsoft.com/.default" };

            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(cred.clientId)
                .WithClientSecret(cred.clientSecret)
                .WithAuthority(new Uri(authority))
                .Build();

            var authResult = app.AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();

            GraphServiceClient _graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider((requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                    return System.Threading.Tasks.Task.CompletedTask;
                }));

            return _graphClient;
        }

        private async Task<GraphServiceClient> GetAuthenticatedGraphClient()
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var clientSecretCredential = new ClientSecretCredential(cred.tenantId, cred.clientId, cred.clientSecret);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            // You might want to force the authentication here to ensure it's completed
            // before returning the client
            await clientSecretCredential.GetTokenAsync(new TokenRequestContext(scopes));

            return graphClient;
        }
        public async void EnviarChatMS()
        {
            GraphServiceClient graphClient = await GetAuthenticatedGraphClient();


            //buscar el Azure Id del usuario al que se le va a enviar el chat
            //https://graph.microsoft.com/v1.0/users('epiedra@gbm.net')
            //crea el chat con el userId
            //https://graph.microsoft.com/v1.0/chats
//            {
//                "chatType": "oneOnOne",
//    "members": [
//        {
//                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
//            "roles": [
//                "owner"
//            ],
//            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('{your-user-id}')"
//        },
//        {
//                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
//            "roles": [
//                "owner"
//            ],
//            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('{user-id}')"
//        }
//    ]
//}
            //devuelve un chat Id

            ChatMessage requestBody = new ChatMessage
            {
                Body = new ItemBody
                {
                    Content = "hola desde databot",
                },
            };

            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            //var result = await graphClient.Teams["{team-id}"].Channels["{channel-id}"].Messages.PostAsync(requestBody);
            var result = await graphClient.Chats["19:9a9521e3-7ded-4330-bd39-45497065b079_ae05a931-c8e8-41fb-a553-1bcc48f2fd28@unq.gbl.spaces"].Messages
                            .Request()
                            .AddResponseAsync(requestBody);

           

        }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
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
    }

}
