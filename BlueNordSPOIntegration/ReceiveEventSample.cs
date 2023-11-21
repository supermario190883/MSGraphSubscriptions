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
    public class ReceiveEventSample
    {
        private readonly ILogger _logger;

        public ReceiveEventSample(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<SubscriptionsGovernance>();
        }

        [Function("ReceiveEventSample")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");


            //This part is needed for when MSGraph tests a new subscription webhook
            var validationToken = req.Query["validationToken"];
            if (!string.IsNullOrEmpty(validationToken))
            {
                HttpResponseData graphResponse = req.CreateResponse(HttpStatusCode.OK);
                graphResponse.Headers.Add("Content-Type", "text/plain; charset=utf-8");

                graphResponse.WriteString(validationToken);

                return graphResponse;
            }
            //////////


            //Handle the event code goes here

            //TODO: PA Consulting

            //


            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");

            response.WriteString($"Event received");

            return response;
        }



  
    }
}
