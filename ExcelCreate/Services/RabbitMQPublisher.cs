using Common;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelCreate.Services
{
    public class RabbitMQPublisher
    {
        private readonly RabbitMQClientService _rabbitMQClientService;

        public RabbitMQPublisher(RabbitMQClientService rabbitMQClientService)
        {
            this._rabbitMQClientService = rabbitMQClientService;
        }

        public void Publish(CreateExcelMessage createExcelMessage)
        {

            var channel = _rabbitMQClientService.Connect();

            var bodyStr = JsonSerializer.Serialize(createExcelMessage);

            var bodyByte = Encoding.UTF8.GetBytes(bodyStr);

            var property = channel.CreateBasicProperties();
            property.Persistent = true;

            channel.BasicPublish(RabbitMQClientService.ExchangeName, RabbitMQClientService.RoutingExcel, property, bodyByte);
        }
    }
}
