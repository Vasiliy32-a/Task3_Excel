using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelApp
{
    public class ExcelHandler
    {
        private string _filePath;
        private IXLWorkbook _workbook;

        public ExcelHandler(string filePath)
        {
            _filePath = filePath;
            _workbook = new XLWorkbook(filePath);
        }

        public List<OrderInfo> GetOrdersByProductName(string productName)
        {
            // Получаем рабочие листы
            var productSheet = _workbook.Worksheet("Товары");
            var orderSheet = _workbook.Worksheet("Заявки");
            var clientSheet = _workbook.Worksheet("Клиенты");

            // Ищем строку с нужным наименованием товара
            var productRow = productSheet.RowsUsed()
                .FirstOrDefault(row => string.Equals(row.Cell(2).GetValue<string>()?.Trim(), productName?.Trim(), StringComparison.OrdinalIgnoreCase));

            // Если товар не найден, возвращаем пустой список
            if (productRow == null)
            {
                return new List<OrderInfo>();
            }

            var productCode = productRow.Cell(1).GetValue<string>();

            // Ищем заказы по коду товара
            var orders = orderSheet.RowsUsed()
                .Where(row => string.Equals(row.Cell(2).GetValue<string>()?.Trim(), productCode?.Trim(), StringComparison.OrdinalIgnoreCase));

            var orderInfos = new List<OrderInfo>();

            foreach (var order in orders)
            {
                var clientCode = order.Cell(3).GetValue<string>();
                var clientRow = clientSheet.RowsUsed()
                    .FirstOrDefault(row => string.Equals(row.Cell(1).GetValue<string>()?.Trim(), clientCode?.Trim(), StringComparison.OrdinalIgnoreCase));

                // Если клиент не найден, пропускаем этот заказ
                if (clientRow == null)
                {
                    continue;
                }

                var clientInfo = new ClientInfo
                {
                    OrganizationName = clientRow.Cell(2).GetValue<string>(),
                    ContactPerson = clientRow.Cell(4).GetValue<string>()
                };

                // Добавляем информацию о заказе
                orderInfos.Add(new OrderInfo
                {
                    Client = clientInfo,
                    Quantity = order.Cell(5).GetValue<int>(),
                    Price = productRow.Cell(4).GetValue<decimal>(),
                    OrderDate = order.Cell(6).GetValue<DateTime>()
                });
            }

            return orderInfos;
        }



        public bool UpdateClientContactPerson(string organizationName, string newContactPerson)
        {
            var clientSheet = _workbook.Worksheet("Клиенты");
            // Находим строку, где наименование организации совпадает с указанным
            var clientRow = clientSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(2).GetValue<string>() == organizationName);
            // Если строка не найдена, возвращаем false
            if (clientRow == null)
                return false;
            // Обновляем контактное лицо в найденной строке
            clientRow.Cell(4).SetValue(newContactPerson);
            _workbook.Save();
            return true;
        }

        public string GetGoldenClient(int year, int month)
        {
            var orderSheet = _workbook.Worksheet("Заявки");

            // Получаем все заказы за указанный период
            var orders = orderSheet.RowsUsed()
                .Skip(1) // Пропускаем заголовок
                .Where(row =>
                {
                    try
                    {
                        var orderDate = row.Cell(6).GetDateTime();
                        return orderDate.Year == year && orderDate.Month == month;
                    }
                    catch
                    {
                        return false;
                    }
                });
            // Подсчитываем количество заявок и общее количество требуемого товара
            var clientOrderCounts = orders
                .GroupBy(row => row.Cell(3).GetValue<string>()) // Группируем по коду клиента
                .Select(group => new
                {
                    ClientCode = group.Key,
                    OrderCount = group.Count(),
                    TotalQuantity = group.Sum(row => row.Cell(5).GetValue<int>())
                })
                .OrderByDescending(x => x.OrderCount) // Сортируем по количеству заявок
                .FirstOrDefault();

            if (clientOrderCounts == null)
                return "Нет данных за указанный период";

            // Ищем клиента на листе "Клиенты" по коду
            var clientSheet = _workbook.Worksheet("Клиенты");
            var clientRow = clientSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(1).GetValue<string>() == clientOrderCounts.ClientCode);

            if (clientRow == null)
                return $"Клиент с кодом {clientOrderCounts.ClientCode} не найден";

            // Возвращаем код клиента, наименование организации и общее количество требуемого товара
            var clientName = clientRow.Cell(2).GetValue<string>();
            return $"Код клиента: {clientOrderCounts.ClientCode}, Наименование организации: {clientName}, Общее количество требуемого товара: {clientOrderCounts.TotalQuantity}";
        }

    }
}
