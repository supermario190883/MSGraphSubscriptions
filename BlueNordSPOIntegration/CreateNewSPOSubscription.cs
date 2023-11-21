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
    public class CreateNewSPOSubscription
    {
        private readonly ILogger _logger;

        public CreateNewSPOSubscription(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<CreateNewSPOSubscription>();
        }

        [Function("CreateNewSPOSubscription")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");


            string siteId = req.Query["siteId"]; //example: 782dfe75-806e-4780-b239-126c422d7059
            string listId = req.Query["listId"]; //example: 84afe751-d292-4f9a-a918-db35557852ec

            if (string.IsNullOrEmpty(siteId) || string.IsNullOrEmpty(listId))
            {
                return req.CreateResponse(HttpStatusCode.BadRequest);
            }

            string responseMessage = string.Empty;

            try
            {
                //Step 1: Create MS Graph client

                var msGraphClient = MSGraphHelper.InitializeGraphClientAsync();


                //Step 2: Create a new MS Graph subcription for the specific siteId + ListId

                var newSubscription = MSGraphHelper.CreateNewSubscription(msGraphClient, siteId, listId);


                //Step 3: Create a new row in the Azure table with the new subscription

                var azureTableSubscriptionLog = AzureTableOperations.AddSubscriptionToAzureTable(newSubscription, siteId, listId);

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

            response.WriteString($"Subscription created for site {siteId} and list {listId}.");

            return response;
        }







    }
}
