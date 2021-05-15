using Microsoft.Extensions.Logging;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelCreate.Services
{
    public class RabbitMQClientService
    {
        private readonly ConnectionFactory _connectionFactory;
        private IConnection _connection;
        private IModel _channel;
        public static string ExchangeName = "ExcelDirectExchange";
        public static string RoutingExcel = "route-excel-file";
        public static string QueueName = "queue-excel-file";

        private readonly ILogger<RabbitMQClientService> _logger;

        public RabbitMQClientService(ConnectionFactory connectionFactory, ILogger<RabbitMQClientService> logger)
        {
            this._connectionFactory = connectionFactory;
            this._logger = logger;
        }

        public IModel Connect()
        {

            _connection = _connectionFactory.CreateConnection();
            if (_channel is { IsOpen: true })
            {
                return _channel;
            }

            _channel = _connection.CreateModel();
            _channel.ExchangeDeclare(ExchangeName, ExchangeType.Direct, true, false);
            _channel.QueueDeclare(QueueName, true, false, false, null);
            _channel.QueueBind(QueueName, ExchangeName, RoutingExcel);

            _logger.LogInformation("Connection with RabbitMQ has been established.");
            return _channel;
        }

        public void Dispose()
        {
            _channel?.Close();
            _channel?.Dispose();
            _connection?.Close();
            _connection?.Dispose();

            _logger.LogInformation("Connection with RabbitMQ has been lost.");
        }
    }
}
