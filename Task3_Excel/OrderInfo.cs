using System;

namespace ExcelApp
{
    public class OrderInfo
    {
        public ClientInfo Client { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public DateTime OrderDate { get; set; }

        public override string ToString()
        {
            return $"Клиент: {Client.OrganizationName}, Контактное лицо: {Client.ContactPerson}, Количество: {Quantity}, Цена: {Price}, Дата заказа: {OrderDate.ToShortDateString()}";
        }
    }
}
