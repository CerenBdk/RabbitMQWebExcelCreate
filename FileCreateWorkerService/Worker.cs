using ClosedXML.Excel;
using Common;
using FileCreateWorkerService.Models;
using FileCreateWorkerService.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace FileCreateWorkerService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly RabbitMQClientService _rabbitMQClientService;
        private readonly IServiceProvider _serviceProvider;

        private IModel _channel;

        public Worker(ILogger<Worker> logger, RabbitMQClientService rabbitMQClientService, IServiceProvider serviceProvider)
        {
            _logger = logger;
            _rabbitMQClientService = rabbitMQClientService;
            _serviceProvider = serviceProvider;
        }

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            _channel = _rabbitMQClientService.Connect();
            _channel.BasicQos(0, 1, false);

            return base.StartAsync(cancellationToken);
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            var consumer = new AsyncEventingBasicConsumer(_channel);
            _channel.BasicConsume(RabbitMQClientService.QueueName, false, consumer);
            consumer.Received += Consumer_Received;
            return Task.CompletedTask;
        }

        private async Task Consumer_Received(object sender, BasicDeliverEventArgs @event)
        {
             await Task.Delay(5000);
            var createExcelMessage = JsonSerializer.Deserialize<CreateExcelMessage>(Encoding.UTF8.GetString(@event.Body.ToArray()));

            MemoryStream memoryStream = new MemoryStream();

            var workBook = new XLWorkbook();
            var dataSet = new DataSet();

            dataSet.Tables.Add(GetTable("products"));

            workBook.Worksheets.Add(dataSet);
            workBook.SaveAs(memoryStream);

            //MultipartFormDataContent multipartFormDataContent = new();
            //multipartFormDataContent.Add(new ByteArrayContent(memoryStream.ToArray()), "file", Guid.NewGuid().ToString() + ".xlsx");

            //var baseUrl = "https://localhost:44324/api/files";

            var client = new RestClient("https://localhost:44324/api/files/upload?fileId=" + createExcelMessage.FileId);
            client.Timeout = -1;
            ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
            client.RemoteCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            var request = new RestRequest(Method.POST);
            request.AddFile("file", memoryStream.ToArray(), Guid.NewGuid().ToString() + ".xlsx");
            IRestResponse response = client.Execute(request);

            if (response.IsSuccessful)
            {
                _logger.LogInformation($"File ( Id : {createExcelMessage.FileId}) was created by successful");
                _channel.BasicAck(@event.DeliveryTag, false);
            }
            //using (var httpClient = new HttpClient())
            //{

            //    var response = await httpClient.PostAsync($"{baseUrl}?fileId={createExcelMessage.FileId}", multipartFormDataContent);

            //    if (response.IsSuccessStatusCode)
            //    {

            //        _logger.LogInformation($"File ( Id : {createExcelMessage.FileId}) was created by successful");
            //        _channel.BasicAck(@event.DeliveryTag, false);
            //    }
            //}

        }

        private DataTable GetTable(string tableName)
        {
            List<FileCreateWorkerService.Models.Product> products;
            using (var scope = _serviceProvider.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<NorthwindContext>();
                products = context.Products.ToList();
            }

            DataTable table = new DataTable { TableName = tableName };
            table.Columns.Add("ProductID", typeof(int));
            table.Columns.Add("ProductName", typeof(String));
            table.Columns.Add("SupplierID", typeof(int));
            table.Columns.Add("CategoryID", typeof(int));
            table.Columns.Add("QuantityPerUnit", typeof(String));
            table.Columns.Add("UnitPrice", typeof(decimal));
            table.Columns.Add("UnitsInStock", typeof(short));
            table.Columns.Add("UnitsOnOrder", typeof(short));
            table.Columns.Add("ReorderLevel", typeof(short));
            table.Columns.Add("Discontinued", typeof(bool));

            products.ForEach(x =>
            {

                table.Rows.Add(x.ProductId, x.ProductName, x.SupplierId, x.CategoryId, x.QuantityPerUnit,
                    x.UnitPrice, x.UnitsInStock, x.UnitsOnOrder, x.ReorderLevel, x.Discontinued);

            });

            return table;
        }
    }
}