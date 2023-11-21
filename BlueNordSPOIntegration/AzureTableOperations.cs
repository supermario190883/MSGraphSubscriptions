using Azure;
using Azure.Data.Tables;
using Azure.Data.Tables.Models;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using Azure.Data.Tables;
using Google.Protobuf.Collections;

namespace BlueNordSPOIntegration
{
    public static class AzureTableOperations
    {
        //TODO: Replace the connectionString with your Azure Table ConnectionString
        public static string tableName = "subscriptions";
        public static string connectionString = "<Connection String>";

        public static bool AddSubscriptionToAzureTable(Subscription subscription, string siteId, string listId)
        {

            TableClient tableClient = new TableClient(connectionString, tableName);


            // Create new item using composite key constructor
            var newSPOSubscription = new BlueNordSubscription()
            {
                PartitionKey = "BlueNordSPOSubscriptions",
                RowKey = subscription.Id,
                SiteId = siteId,
                ListId = listId,
                SubscriptionId = subscription.Id,
                SubscriptionCreatedDate = DateTimeOffset.UtcNow,
                SubscriptionExpirationDate = DateTimeOffset.UtcNow.AddDays(30),

            };

            tableClient.AddEntity(newSPOSubscription);

            return true;
        }

        public static List<BlueNordSubscription>? GetExpiringSubscriptions()
        {

            var subscriptionsExpiring = new List<BlueNordSubscription>();

            TableClient tableClient = new TableClient(connectionString, tableName);

            var subscriptions = tableClient.Query<BlueNordSubscription>(x => x.PartitionKey == "BlueNordSPOSubscriptions");

            Console.WriteLine("Found these subscriptions about to expire:");
            foreach (var item in subscriptions)
            {
                if (item.SubscriptionExpirationDate < DateTime.UtcNow.AddHours(48))
                {
                    Console.WriteLine(item.SubscriptionExpirationDate);
                    subscriptionsExpiring.Add(item);
                }
            }

            return subscriptionsExpiring;
        }

        public static bool UpdateSubscriptionExpirationDate(string subscriptionId)
        {
            try
            {
                // Create new item using composite key constructor
                var newSPOSubscription = new BlueNordSubscription()
                {
                    PartitionKey = "BlueNordSPOSubscriptions",
                    RowKey = subscriptionId,
                    SubscriptionId = subscriptionId,
                    SubscriptionExpirationDate = DateTimeOffset.UtcNow.AddDays(30)

                };

                TableClient tableClient = new TableClient(connectionString, tableName);

                tableClient.UpdateEntityAsync<BlueNordSubscription>(newSPOSubscription, ETag.All, TableUpdateMode.Merge);

                return true;

            }
            catch(Exception exp)
            {
                return false;
            }

            
        }
        public record BlueNordSubscription : ITableEntity
        {
            public required string RowKey { get; set; }
            public required string PartitionKey { get; set; }
            public required string SubscriptionId { get; init; }
            public string? SiteId { get; init; }
            public string? ListId { get; init; }
            public DateTimeOffset? SubscriptionCreatedDate { get; set; }
            public required DateTimeOffset SubscriptionExpirationDate { get; set; }
            public DateTimeOffset? Timestamp { get; set; }
            public ETag ETag { get; set; }
        }

    }
}
