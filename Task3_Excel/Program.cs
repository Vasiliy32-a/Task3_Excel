using System;
using System.Collections.Generic;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите путь до файла Excel:");
            string filePath = Console.ReadLine();

            ExcelHandler excelHandler = new ExcelHandler(filePath);

            while (true)
            {
                Console.WriteLine("Выберите команду:");
                Console.WriteLine("1: Найти информацию о клиентах по наименованию товара");
                Console.WriteLine("2: Изменить контактное лицо клиента");
                Console.WriteLine("3: Определить золотого клиента");
                Console.WriteLine("0: Выход");

                string command = Console.ReadLine();
                switch (command)
                {
                    case "1":
                        Console.WriteLine("Введите наименование товара:");
                        string productName = Console.ReadLine();
                        var orders = excelHandler.GetOrdersByProductName(productName);
                        foreach (var order in orders)
                        {
                            Console.WriteLine(order);
                        }
                        break;
                    case "2":
                        Console.WriteLine("Введите название организации:");
                        string organizationName = Console.ReadLine();
                        Console.WriteLine("Введите новое ФИО контактного лица:");
                        string newContactPerson = Console.ReadLine();
                        bool result = excelHandler.UpdateClientContactPerson(organizationName, newContactPerson);
                        Console.WriteLine(result ? "Контактное лицо обновлено" : "Ошибка при обновлении контактного лица");
                        break;
                    case "3":
                        Console.WriteLine("Введите год:");
                        int year = int.Parse(Console.ReadLine());
                        Console.WriteLine("Введите месяц:");
                        int month = int.Parse(Console.ReadLine());
                        var goldenClient = excelHandler.GetGoldenClient(year, month);
                        Console.WriteLine($"Золотой клиент: {goldenClient}");
                        break;
                    case "0":
                        return;
                    default:
                        Console.WriteLine("Неверная команда");
                        break;
                }
            }
        }
    }
}