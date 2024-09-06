using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Task_3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // "C:\Users\magom\Downloads\Практическое задание для кандидата.xlsx"
            EnteringPathToFile(out string filePath);

            var products = LoadProducts(filePath);
            var clients = LoadClients(filePath);
            var orders = LoadOrders(filePath);

            int numberCommand = int.MinValue;

            while (numberCommand != 0)
            {
                DisplayCommandList();

                if (int.TryParse(Console.ReadLine(), out numberCommand))
                {
                    ProcessCommand(numberCommand, products, clients, orders, filePath);
                }
                else
                {
                    Console.Clear();
                    Console.WriteLine("Некорректный ввод. Введите номер команды от 1 до 3.");
                }
            }

            DisplayCompletion();
        }


        /// <summary>
        /// Загружает список продуктов из файла Excel
        /// </summary>
        /// <param name="filePath">Путь к файлу Excel</param>
        /// <returns>Список продуктов</returns>
        private static List<Product> LoadProducts(string filePath)
        {
            var products = new List<Product>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(document, "Товары");
                if (worksheetPart == null) throw new Exception("Лист 'Товары' не найден.");

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var rows = sheetData.Elements<Row>().Skip(1); // Пропускаем заголовка

                foreach (var row in rows)
                {
                    var cells = row.Elements<Cell>().ToList();
                    products.Add(new Product
                    {
                        Code = GetCellValue(cells[0], document),
                        Name = GetCellValue(cells[1], document),
                        Unit = GetCellValue(cells[2], document),
                        Price = decimal.Parse(GetCellValue(cells[3], document))
                    });
                }
            }
            return products;
        }

        private static void DisplayCommandList()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nСписок команд: ");
            Console.WriteLine("0. Завершить работу программы");
            Console.WriteLine("1. Узнать кто заказывал товар");
            Console.WriteLine("2. Изменить информацию о клиенте");
            Console.WriteLine("3. Золотой клиент\n");
            Console.ResetColor();
            Console.Write("Выберите номер команды: ");
        }

        private static void DisplayCompletion()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Clear();
            Console.WriteLine("...программа завершила свою работу");
            Console.ReadLine();
        }

        static void ProcessCommand(int numberCommand, List<Product> products, List<Client> clients, List<Order> orders, string filePath)
        {
            switch (numberCommand)
            {
                case 1:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Clear();
                    Console.WriteLine("Вы выбрали команду №1 | Узнать кто заказывал товар\n");
                    Console.ResetColor();
                    Console.Write("Введите название продукта: ");
                    string productName = Console.ReadLine();
                    DisplayOrderInfoByProductName(productName, products, clients, orders);
                    break;

                case 2:
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Вы выбрали команду №2 | Изменить информацию о клиенте\n");
                    Console.ResetColor();
                    Console.Write("Введите название организации для смены контактного лица: ");
                    string organizationName = Console.ReadLine();

                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
                    {
                        UpdateClientContact(organizationName, document, clients);
                    }

                    Console.ResetColor();
                    break;

                case 3:
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Вы выбрали команду №3 | Золотой клиент\n");
                    Console.ResetColor();
                    Console.Write("Укажите год: ");
                    int year = Convert.ToInt32(Console.ReadLine());
                    if (year > DateTime.Now.Year || year <= 2000)
                    {
                        Console.WriteLine("Ошибка! Введите год корректно!");
                        break;
                    }
                    Console.Write("Укажите месяц: ");
                    int month = Convert.ToInt32(Console.ReadLine());
                    if (month > 12 || month <= 0)
                    {
                        Console.WriteLine("Некорректный ввод. Введите номер месяца от 1 до 12");
                        break;
                    }
                    FindGoldClient(year, month, orders, clients);
                    break;

                default:
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Некорректный ввод. Введите номер команды от 1 до 3.");
                    Console.ResetColor();

                    break;
            }
        }

        /// <summary>
        /// Загружает список клиентов из файла Excel
        /// </summary>
        /// <param name="filePath">Путь к файлу Excel</param>
        /// <returns>Список клиентов</returns>
        private static List<Client> LoadClients(string filePath)
        {
            var clients = new List<Client>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(document, "Клиенты");
                if (worksheetPart == null) throw new Exception("Лист 'Клиенты' не найден.");

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var rows = sheetData.Elements<Row>().Skip(1); // Пропуск заголовка

                foreach (var row in rows)
                {
                    var cells = row.Elements<Cell>().ToList();
                    clients.Add(new Client
                    {
                        Code = GetCellValue(cells[0], document),
                        Name = GetCellValue(cells[1], document),
                        Address = GetCellValue(cells[2], document),
                        ContactName = GetCellValue(cells[3], document)
                    });
                }
            }
            return clients;
        }

        /// <summary>
        /// Загружает список заявок из файла Excel
        /// </summary>
        /// <param name="filePath">Путь к файлу Excel.</param>
        /// <returns>Список заявок</returns>
        private static List<Order> LoadOrders(string filePath)
        {
            var orders = new List<Order>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(document, "Заявки");
                if (worksheetPart == null) throw new Exception("Лист 'Заявки' не найден.");

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var rows = sheetData.Elements<Row>().Skip(1); // Пропуск заголовка

                foreach (var row in rows)
                {
                    var cells = row.Elements<Cell>().ToList();

                    orders.Add(new Order
                    {
                        OrderCode = GetCellValue(cells[0], document),
                        ProductCode = GetCellValue(cells[1], document),
                        ClientCode = GetCellValue(cells[2], document),
                        RequestNumber = GetCellValue(cells[3], document),
                        Quantity = (GetCellValue(cells[4], document)),
                        Date = ParseDate(GetCellValue(cells[5], document))
                    });
                }
            }
            return orders;
        }

        /// <summary>
        /// Получает элемент <see cref="WorksheetPart"/> для листа с указанным именем
        /// </summary>
        /// <param name="document">Документ Excel, из которого необходимо извлечь лист</param>
        /// <param name="sheetName">Имя листа, для которого нужно получить <see cref="WorksheetPart"/></param>
        /// <returns>
        /// Возвращает <see cref="WorksheetPart"/> для листа с указанным именем или null если такого листа нет
        /// </returns>
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            var sheet = document.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                .FirstOrDefault(s => s.Name == sheetName);

            if (sheet == null) return null;

            return (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
        }


        /// <summary>
        /// Преобразует строку, представляющую дату, в <see cref="DateTime"/>. 
        /// Если строка не распознается, пытается преобразовать строку по заданным форматам.
        /// </summary>
        /// <param name="dateString">Строка, представляющая дату.</param>
        /// <returns>Возвращает <see cref="DateTime"/>, если преобразование прошло успешно.</returns>
        private static DateTime ParseDate(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
            {
                //Чтобы не вылетала ошибка при обработке пустых строк (а они будут), но работать с этими датами не будут.
                return DateTime.MinValue;

            }

            // Так как дата поступает как числовое значение, надо преобразовать его в дату
            if (double.TryParse(dateString, out double oaDate))
            {
                DateTime date = DateTime.FromOADate(oaDate);

                return date;
            }

            throw new FormatException($"Не удалось распознать дату: '{dateString}'");
        }


        /// <summary>
        /// По наименованию товара выводит информацию о клиентах, заказавших этот товар, с указанием информации по количеству товара, цене и дате заказа.
        /// </summary>
        /// <param name="productName"></param>
        /// <param name="products"></param>
        /// <param name="clients"></param>
        /// <param name="orders"></param>
        static public void DisplayOrderInfoByProductName(string productName, List<Product> products, List<Client> clients, List<Order> orders)
        {
            var product = products.FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));
            if (product == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Товар {productName} не найден.");
                Console.ResetColor();
                return;
            }

            var orderDetails = from order in orders
                               join client in clients on order.ClientCode equals client.Code
                               where order.ProductCode == product.Code
                               select new
                               {
                                   client.Name,
                                   order.Quantity,
                                   order.Date,
                                   product.Price
                               };
            Console.Clear();
            Console.WriteLine($"Информация о клиентах, заказавших товар '{productName}':");
            if (!orderDetails.Any())
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Никто не приобретал товар '{productName}'");
            }

            foreach (var detail in orderDetails)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Организация: {detail.Name}, Количество: {detail.Quantity}, Цена: {detail.Price}, Дата заказа: {detail.Date:yyyy-MM-dd}");
            }
            Console.ResetColor();
        }

        /// <summary>
        /// Обновляет контактное лицо в Exel-документе
        /// </summary>
        /// <param name="organizationName">Название организации, для которой нужно изменить контактное лицо</param>
        /// <param name="document">объект типа SpreadsheetDocument, представляющий Excel-документ, где хранится информация о клиентах</param>
        /// <param name="clients">список клиентов</param>
        static public void UpdateClientContact(string organizationName, SpreadsheetDocument document, List<Client> clients)
        {
            // Находиим клиента по названию организации
            var client = clients.FirstOrDefault(c => c.Name.Equals(organizationName, StringComparison.OrdinalIgnoreCase));
            if (client == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Клиент не найден.");
                Console.ResetColor();
                return;
            }

            // Обновляем контактное лицо в памяти

            Console.Write("Введите новое ФИО нового контактного лица: ");
            string newContactName = Console.ReadLine();
            client.ContactName = newContactName;

            // Обновляем контактное лицо в Excel
            var clientsSheetPart = GetWorksheetPartByName(document, "Клиенты");
            if (clientsSheetPart == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Лист 'Клиенты' не найден.");
                Console.ResetColor();
                return;
            }

            var sheetData = clientsSheetPart.Worksheet.GetFirstChild<SheetData>();
            foreach (var row in sheetData.Elements<Row>())
            {
                var cellValues = row.Elements<Cell>().Select(c => GetCellValue(c, document)).ToArray();
                if (cellValues.Length > 0 && cellValues[0] == client.Code) // соответствует ли первая ячейка (которая содержит код клиента) коду клиента, найденному в памяти
                {
                    var contactCell = row.Elements<Cell>().ElementAt(3);
                    contactCell.CellValue = new CellValue(newContactName);
                    break;
                }
            }

            // Сохранить изменения
            clientsSheetPart.Worksheet.Save();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Контактное лицо клиента '{organizationName}' было обновлено на '{newContactName}'.");
            Console.ResetColor();
        }

        static public void FindGoldClient(int year, int month, List<Order> orders, List<Client> clients)
        {
            // Фильтруем заказы по году и месяцу
            var filteredOrders = orders.Where(o => o.Date.Year == year && o.Date.Month == month);

            // Подсчитываем количество заказов по каждому клиенту
            var clientOrderCounts = filteredOrders
                .GroupBy(o => o.ClientCode)
                .Select(g => new
                {
                    ClientCode = g.Key,
                    OrderCount = g.Count()
                })
                .OrderByDescending(c => c.OrderCount)
                .FirstOrDefault();

            if (clientOrderCounts == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Нет заказов за указанный период.");
                Console.ResetColor();
                return;
            }

            // Найти информацию о клиенте
            var goldClient = clients.FirstOrDefault(c => c.Code == clientOrderCounts.ClientCode);
            if (goldClient == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Клиент не найден.");
                Console.ResetColor();
                return;
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\nЗолотой клиент за {month} месяц {year} года: ");
            Console.WriteLine($"Организация: {goldClient.Name}, Количество заказов: {clientOrderCounts.OrderCount}");
            Console.ResetColor();
        }


        /// <summary>
        /// Получает значение ячейки как строку, либо напрямую либо по индексу
        /// </summary>
        /// <param name="cell">Ячейка, значение которой нужно получить</param>
        /// <param name="document">Документ Excel, содержащий таблицу</param>
        /// <returns>Возвращает значение ячейки как строку</returns>
        private static string GetCellValue(Cell cell, SpreadsheetDocument document)
        {
            if (cell == null || cell.CellValue == null)
            {
                return string.Empty;
            }

            string value = cell.CellValue.Text;

            // Извлечение строки из таблицы общих строк
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringTablePart != null)
                {
                    // извлекаем строку по индексу (value)
                    value = sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value;
        }


        private static void EnteringPathToFile(out string path)
        {
            path = string.Empty;

            while (string.IsNullOrWhiteSpace(path))
            {
                Console.Write("Введите путь до файла с данными: ");
                path = "C:\\Users\\magom\\Downloads\\Практическое задание для кандидата.xlsx";
                //path = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(path))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Путь не может быть пустым. Попробуйте снова.");
                    Console.ResetColor();
                }
                else if (File.Exists(path))
                {
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Файл успешно загружен!");
                    Console.ResetColor();

                }
                else if (!File.Exists(path))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Файл по указанному пути не существует. Попробуйте снова.");
                    Console.ResetColor();
                    path = string.Empty;
                }

            }

        }
    }
}
