using System.Net;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Azure.Data.Tables;

namespace BlueNordSPOIntegration
{
    public class SubscriptionsGovernance
    {
        private readonly ILogger _logger;

        public SubscriptionsGovernance(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<SubscriptionsGovernance>();
        }

        [Function("SubscriptionGovernance")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");


            try
            {
                //Step 1: Create MS Graph client

                var msGraphClient = MSGraphHelper.InitializeGraphClientAsync();

                //Step 2: Get list of subscriptions about to expire

                var expiringSubscriptions = AzureTableOperations.GetExpiringSubscriptions();

                foreach(var subscriptionExpiring in expiringSubscriptions)
                {
                    //Step 3: Renew subscription in MS Graph
                    var renewedSubscription = MSGraphHelper.RenewSubscription(msGraphClient, subscriptionExpiring.SubscriptionId);

                    //Step 4: Update the subscription in the Azure Table

                    if (renewedSubscription)
                    {
                        AzureTableOperations.UpdateSubscriptionExpirationDate(subscriptionExpiring.SubscriptionId);
                    }
                }


            }
            catch (Exception exp)
            {
                
                var responseError = req.CreateResponse(HttpStatusCode.InternalServerError);
                responseError.Headers.Add("Content-Type", "text/plain; charset=utf-8");

                responseError.WriteString($"An error as occurred");

                return responseError;
            }


            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");

            response.WriteString($"Subscriptions renewed");

            return response;
        }



  
    }
}
