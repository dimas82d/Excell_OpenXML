using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrderManagement
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Введите путь до файла с данными Excel: ");
            string filePath = Console.ReadLine();

            if (!string.IsNullOrEmpty(filePath))
            {
                var products = LoadProducts(filePath);
                var customers = LoadCustomers(filePath);
                var orders = LoadOrders(filePath);

                while (true)
                {
                    Console.WriteLine("\nВыберите команду:");
                    Console.WriteLine("1. Получить информацию о клиенте по товару");
                    Console.WriteLine("2. Изменить контактное лицо клиента");
                    Console.WriteLine("3. Определить золотого клиента");
                    Console.WriteLine("4. Выход");

                    var choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            GetCustomersByProduct(products, customers, orders);
                            break;
                        case "2":
                            UpdateContactPerson(customers, filePath); // Передаем путь к файлу для записи изменений
                            break;
                        case "3":
                            GetGoldenClient(customers, orders);
                            break;
                        case "4":
                            return;
                        default:
                            Console.WriteLine("Неверный ввод. Пожалуйста, попробуйте снова.");
                            break;
                    }
                }
            }
        }

  

        static List<Product> LoadProducts(string filePath)
        {
            var products = new List<Product>();

            try
            {
                using (var doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var sheet = doc.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name == "Товары");
                    var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                    foreach (var row in rows.Skip(1)) // пропустить заголовки
                    {
                        var cells = row.Elements<Cell>().ToList();
                        if (cells.Count < 4) continue; // удостовериться, что достаточно колонок

                        products.Add(new Product
                        {
                            ProductCode = GetCellValue(doc, cells[0]),
                            Name = GetCellValue(doc, cells[1]),
                            Unit = GetCellValue(doc, cells[2]),
                            Price = decimal.Parse(GetCellValue(doc, cells[3]))
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке товаров: {ex.Message}");
            }

            return products;
        }

        static List<Customer> LoadCustomers(string filePath)
        {
            var customers = new List<Customer>();

            try
            {
                using (var doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var sheet = doc.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name == "Клиенты");
                    var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                    foreach (var row in rows.Skip(1)) // пропустить заголовки
                    {
                        var cells = row.Elements<Cell>().ToList();
                        if (cells.Count < 4) continue; // удостовериться, что достаточно колонок

                        customers.Add(new Customer
                        {
                            CustomerCode = GetCellValue(doc, cells[0]),
                            OrganizationName = GetCellValue(doc, cells[1]),
                            Address = GetCellValue(doc, cells[2]),
                            ContactPerson = GetCellValue(doc, cells[3])
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке клиентов: {ex.Message}");
            }

            return customers;
        }

        static List<Order> LoadOrders(string filePath)
        {
            var orders = new List<Order>();

            try
            {
                using (var doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var sheet = doc.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name == "Заявки");
                    var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                    foreach (var row in rows.Skip(1)) // пропустить заголовки
                    {
                        var cells = row.Elements<Cell>().ToList();
                        if (cells.Count < 6) continue; // удостовериться, что достаточно колонок
                        
                        orders.Add(new Order
                        {
                            OrderCode = GetCellValue(doc, cells[0]),
                            ProductCode = GetCellValue(doc, cells[1]),
                            CustomerCode = GetCellValue(doc, cells[2]),
                            ApplicationNumber = GetCellValue(doc, cells[3]),
                            Quantity = int.Parse(GetCellValue(doc, cells[4])),
                            OrderDate = DateTime.FromOADate(double.Parse(GetCellValue(doc, cells[5]))),
                        });
   
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке заказов: {ex.Message}");
            }

            return orders;
        }

        static string GetCellValue(SpreadsheetDocument doc, Cell cell)
         {
            // Получить значение ячейки
            if (cell.DataType == null || cell.DataType.Value != CellValues.SharedString)
                 return cell.InnerText;

            var stringTablePart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTablePart != null)
            {
                return stringTablePart.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
            }
            
            

            return null;
         }
      

        static void GetCustomersByProduct(List<Product> products, List<Customer> customers, List<Order> orders)
        {
            Console.WriteLine("Введите наименование товара:");
            string productName = Console.ReadLine();

            var product = products.FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));
            if (product == null)
            {
                Console.WriteLine("Товар не найден.");
                return;
            }
  
            foreach (var orde in orders) {

                Console.WriteLine("Заказ: " + orde.ProductCode + " " + orde.CustomerCode + " " + orde.OrderCode + " " + orde.OrderDate.ToString("dd.MM.yyyy"));
            }
            var productOrders = orders.Where(o => o.ProductCode == product.ProductCode).ToList();

            if (!productOrders.Any())
            {
                Console.WriteLine("Нет заказов на этот товар.");
                return;
            }

            Console.WriteLine($"Клиенты, заказавшие товар \"{productName}\":");
            foreach (var order in productOrders)
            {
                var customer = customers.FirstOrDefault(c => c.CustomerCode == order.CustomerCode);
                if (customer != null)
                {
                    Console.WriteLine($"Организация: {customer.OrganizationName}, " +
                                      $"Количество: {order.Quantity}, " +
                                      $"Цена: {product.Price * order.Quantity}, " +
                                      $"Дата заказа: {order.OrderDate.ToString("dd.MM.yyyy")}");
                }
            }
        }

        static void UpdateContactPerson(List<Customer> customers, string filePath)
        {
            Console.WriteLine("Введите название организации:");
            string organizationName = Console.ReadLine();

            var customer = customers.FirstOrDefault(c => c.OrganizationName.Equals(organizationName, StringComparison.OrdinalIgnoreCase));
            if (customer == null)
            {
                Console.WriteLine("Клиент не найден.");
                return;
            }

            Console.WriteLine("Введите ФИО нового контактного лица:");
            string newContactPerson = Console.ReadLine();

            customer.ContactPerson = newContactPerson;

            // Запись изменений обратно в Excel
            SaveCustomerChanges(customers, filePath);

            Console.WriteLine($"Контактное лицо для организации \"{organizationName}\" успешно обновлено на \"{newContactPerson}\".");
        }

        static void SaveCustomerChanges(List<Customer> customers, string filePath)
        {
            try
            {
                using (var doc = SpreadsheetDocument.Open(filePath, true)) // открываем для записи
                {
                    var sheet = doc.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name == "Клиенты");
                    var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().ToList();

                    // Начинаем с первой строки после заголовков
                    for (int i = 1; i < rows.Count; i++)
                    {
                        var cells = rows[i].Elements<Cell>().ToList();
                        if (cells.Count < 4) continue; // удостовериться, что достаточно колонок

                        var customerCode = GetCellValue(doc, cells[0]);
                        var customer = customers.FirstOrDefault(c => c.CustomerCode == customerCode);

                        if (customer != null)
                        {
                            cells[3].CellValue = new CellValue(customer.ContactPerson); // обновляем ячейку с контактным лицом
                        }
                    }

                    worksheetPart.Worksheet.Save(); // сохраняем изменения
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении изменений: {ex.Message}");
            }
        }

        static void GetGoldenClient(List<Customer> customers, List<Order> orders)
        {
            Console.WriteLine("Введите год:");
            int year;
            while (!int.TryParse(Console.ReadLine(), out year))
            {
                Console.WriteLine("Неверный ввод года. Пожалуйста, введите корректное значение.");
            }

            Console.WriteLine("Введите месяц:");
            int month;
            while (!int.TryParse(Console.ReadLine(), out month) || month < 1 || month > 12)
            {
                Console.WriteLine("Неверный ввод месяца. Пожалуйста, введите корректное значение от 1 до 12.");
            }

            var groupedOrders = orders.Where(o => o.OrderDate.Year == year && o.OrderDate.Month == month)
                .GroupBy(o => o.CustomerCode)
                .Select(g => new
                {
                    CustomerCode = g.Key,
                    OrderCount = g.Count()
                }).ToList();

            var goldenClientCode = groupedOrders.OrderByDescending(g => g.OrderCount).FirstOrDefault()?.CustomerCode;

            if (goldenClientCode == null)
            {
                Console.WriteLine("Нет заказов за указанный период.");
                return;
            }

            var goldenCustomer = customers.FirstOrDefault(c => c.CustomerCode == goldenClientCode);
            Console.WriteLine($"Золотой клиент за {month}/{year}: {goldenCustomer.OrganizationName} с {groupedOrders.FirstOrDefault(g => g.CustomerCode == goldenClientCode)?.OrderCount} заказами.");
        }
    }

    public class Product
    {
        public string ProductCode { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }
    }

    public class Customer
    {
        public string CustomerCode { get; set; }
        public string OrganizationName { get; set; }
        public string Address { get; set; }
        public string ContactPerson { get; set; }
    }

    public class Order
    {
        public string OrderCode { get; set; }
        public string ProductCode { get; set; }
        public string CustomerCode { get; set; }
        public string ApplicationNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }
}
