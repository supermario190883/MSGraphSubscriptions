using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlueNordSPOIntegration
{
    public static class MSGraphHelper
    {
        public static GraphServiceClient InitializeGraphClientAsync()
        {

            // TODO: These values should come from Keyvault
            var clientId = "<ClientId>";
            var clientSecret = "<SecretId>";
            var tenantId = "<TenantId>";
            /////

            var scopes = new[] { ".default" };

            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            return graphClient;
        }

        public static Subscription CreateNewSubscription(GraphServiceClient graphClient, string siteId, string listId)
        {

            var subscription = new Subscription
            {
                ChangeType = "updated",
                //TODO: replace the notificationURL with the your Azure function URL that will receive the event from Graph
                NotificationUrl = "https://subscriptionsoperations.azurewebsites.net/api/ReceiveEventSample?code=57RBd5u4PXogNeCLVGN4wrkOItnufNUXjpLWBkqnuka_AzFuiGHoeA==",
                //NotificationUrl = "<AZURE_FUNCTION_URL>",
                Resource = $"/sites/{siteId}/lists/{listId}",
                ExpirationDateTime = DateTime.UtcNow.AddDays(30)

            };

            var subscriptionCreationResult = graphClient.Subscriptions.PostAsync(subscription).Result;

            return subscriptionCreationResult;

        }

        public static bool RenewSubscription(GraphServiceClient graphClient, string subscriptionId)
        {


            try
            {
                var graphSubscription = new Subscription
                {
                    ExpirationDateTime = DateTime.UtcNow.AddDays(30)
                };

                var graphSubscriptionRenewal = graphClient.Subscriptions[subscriptionId].PatchAsync(graphSubscription);

                return true;

            }catch(Exception exp)
            {
                return false;
            }


        }

    }
}
