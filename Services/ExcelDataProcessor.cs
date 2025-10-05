using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ExcelDataReader;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Класс для обработки данных из файла Excel
    /// </summary>
    public class ExcelDataProcessor : IDataProcessor
    {
        // Константы для обработки данных
        private const double SIMILARITY_THRESHOLD = 0.8; // Порог схожести для группировки товаров
        
        /// <summary>
        /// Загрузка данных из файла Excel
        /// </summary>
        /// <param name="filePath">Путь к файлу Excel</param>
        /// <returns>Список элементов заявок</returns>
        public List<OrderItem> LoadData(string filePath)
        {
            var orderItems = new List<OrderItem>();
            
            try
            {
                // Регистрируем кодировку для корректной работы с кириллицей
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Автоматически определяем формат файла (xls или xlsx)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Преобразуем Excel в DataSet
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true // Используем первую строку как заголовки
                            }
                        });
                        
                        // Предполагаем, что данные находятся на первом листе
                        DataTable dataTable = result.Tables[0];
                        
                        // Обрабатываем каждую строку данных
                        foreach (DataRow row in dataTable.Rows)
                        {
                            // Проверяем, что строка не пустая
                            if (row[0] == DBNull.Value || string.IsNullOrWhiteSpace(row[0].ToString()))
                                continue;
                            
                            var orderItem = new OrderItem
                            {
                                // Преобразуем значения ячеек в соответствующие типы данных
                                OrderDate = ParseDate(GetCellValue(row, 0)),
                                OrderNumber = GetCellValue(row, 1),
                                PositionNumber = GetCellValue(row, 2),
                                ProductName = GetCellValue(row, 3),
                                ArticleNumber = GetCellValue(row, 4),
                                OrderedQuantity = ParseDouble(GetCellValue(row, 5)),
                                DeliveredQuantity = ParseDeliveredQuantity(GetCellValue(row, 6)),
                                DeliveryDate = ParseNullableDate(GetCellValue(row, 7)),
                                Notes = GetCellValue(row, 8)
                            };
                            
                            // Если артикул не указан, пытаемся извлечь его из наименования
                            if (string.IsNullOrWhiteSpace(orderItem.ArticleNumber) && !string.IsNullOrWhiteSpace(orderItem.ProductName))
                            {
                                orderItem.ArticleNumber = ExtractArticleFromName(orderItem.ProductName);
                            }
                            
                            orderItems.Add(orderItem);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // В реальном приложении здесь должна быть обработка ошибок
                // Логирование ошибки вместо вывода в консоль
                throw new Exception($"Ошибка при загрузке данных: {ex.Message}", ex);
            }
            
            return orderItems;
        }
        
        /// <summary>
        /// Унификация наименований и артикулов товаров
        /// </summary>
        /// <param name="orderItems">Список элементов заявок</param>
        /// <returns>Список унифицированных товаров</returns>
        public List<UnifiedProduct> UnifyProducts(List<OrderItem> orderItems)
        {
            var unifiedProducts = new List<UnifiedProduct>();
            
            // Группируем заказы по артикулам
            var groupsByArticle = orderItems
                .Where(item => !string.IsNullOrWhiteSpace(item.ArticleNumber))
                .GroupBy(item => item.ArticleNumber!.Trim().ToUpper());
            
            // Создаем унифицированные товары для каждой группы
            foreach (var group in groupsByArticle)
            {
                var unifiedProduct = new UnifiedProduct
                {
                    UnifiedArticle = group.Key,
                    PrimaryName = GetMostFrequentName(group),
                    OrderHistory = group.ToList(),
                    NameVariations = new List<string>(),
                    ArticleVariations = new List<string>()
                };
                
                // Добавляем все вариации наименований
                unifiedProduct.NameVariations = group
                    .Select(item => item.ProductName ?? string.Empty)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Distinct()
                    .ToList();
                
                // Добавляем артикул в список вариаций
                unifiedProduct.ArticleVariations.Add(group.Key);
                
                unifiedProducts.Add(unifiedProduct);
            }
            
            // Обрабатываем товары без артикулов, группируя их по схожести наименований
            var itemsWithoutArticle = orderItems
                .Where(item => string.IsNullOrWhiteSpace(item.ArticleNumber))
                .ToList();
            
            foreach (var item in itemsWithoutArticle)
            {
                bool added = false;
                
                // Пропускаем товары без наименования
                if (string.IsNullOrWhiteSpace(item.ProductName))
                    continue;

                // Пытаемся найти подходящий унифицированный товар по схожести наименования
                foreach (var unifiedProduct in unifiedProducts)
                {
                    if (unifiedProduct.NameVariations != null &&
                        unifiedProduct.NameVariations.Any(name =>
                            !string.IsNullOrEmpty(name) &&
                            CalculateSimilarity(name, item.ProductName) > SIMILARITY_THRESHOLD))
                    {
                        // Добавляем товар к существующему унифицированному товару
                        unifiedProduct.OrderHistory.Add(item);

                        // Добавляем наименование, если его еще нет в списке вариаций
                        if (!unifiedProduct.NameVariations.Contains(item.ProductName))
                        {
                            unifiedProduct.NameVariations.Add(item.ProductName);
                        }

                        added = true;
                        break;
                    }
                }
                
                // Если не нашли подходящий товар, создаем новый
                if (!added)
                {
                    var newUnifiedProduct = new UnifiedProduct
                    {
                        UnifiedArticle = $"AUTO_{unifiedProducts.Count + 1}",
                        PrimaryName = item.ProductName ?? string.Empty,
                        OrderHistory = new List<OrderItem> { item }
                    };
                    
                    if (!string.IsNullOrEmpty(item.ProductName))
                    {
                        newUnifiedProduct.NameVariations.Add(item.ProductName);
                    }
                    
                    unifiedProducts.Add(newUnifiedProduct);
                }
            }
            
            // Рассчитываем статистические показатели для каждого унифицированного товара
            foreach (var product in unifiedProducts)
            {
                CalculateStatistics(product);
            }
            
            return unifiedProducts;
        }
        
        /// <summary>
        /// Сохранение базы соответствий наименований и артикулов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        public void SaveItemMapping(List<UnifiedProduct> unifiedProducts, string filePath)
        {
            try
            {
                // Создаем упрощенную структуру для сохранения
                var mappings = unifiedProducts.Select(p => new
                {
                    p.UnifiedArticle,
                    p.PrimaryName,
                    p.NameVariations,
                    p.ArticleVariations
                }).ToList();
                
                // Сериализуем в JSON и сохраняем в файл
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(mappings, options);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                // Логирование ошибки вместо вывода в консоль
                throw new Exception($"Ошибка при сохранении базы соответствий: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// Загрузка базы соответствий наименований и артикулов
        /// </summary>
        /// <param name="filePath">Путь к файлу с базой соответствий</param>
        /// <returns>Список унифицированных товаров</returns>
        public List<UnifiedProduct> LoadItemMapping(string filePath)
        {
            var unifiedProducts = new List<UnifiedProduct>();
            
            try
            {
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    var mappings = JsonSerializer.Deserialize<List<MappingItem>>(json) ?? new List<MappingItem>();
                    
                    foreach (var mapping in mappings)
                    {
                        unifiedProducts.Add(new UnifiedProduct
                        {
                            UnifiedArticle = mapping.UnifiedArticle ?? string.Empty,
                            PrimaryName = mapping.PrimaryName ?? string.Empty,
                            NameVariations = mapping.NameVariations ?? new List<string>(),
                            ArticleVariations = mapping.ArticleVariations ?? new List<string>()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                // Логирование ошибки вместо вывода в консоль
                throw new Exception($"Ошибка при загрузке базы соответствий: {ex.Message}", ex);
            }
            
            return unifiedProducts;
        }
        
        /// <summary>
        /// Получение значения ячейки из DataRow
        /// </summary>
        private string GetCellValue(DataRow row, int columnIndex)
        {
            if (columnIndex >= row.ItemArray.Length || row[columnIndex] == DBNull.Value)
                return string.Empty;
                
            return row[columnIndex].ToString() ?? string.Empty;
        }
        
        #region Вспомогательные методы
        
        /// <summary>
        /// Преобразование строки в дату
        /// </summary>
        private DateTime ParseDate(string dateString)
        {
            DateTime result;
            if (DateTime.TryParse(dateString, out result))
                return result;
            
            return DateTime.MinValue;
        }
        
        /// <summary>
        /// Преобразование строки в nullable дату
        /// </summary>
        private DateTime? ParseNullableDate(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return null;
            
            DateTime result;
            if (DateTime.TryParse(dateString, out result))
                return result;
            
            return null;
        }
        
        /// <summary>
        /// Преобразование строки в число с плавающей точкой
        /// </summary>
        private double ParseDouble(string numberString)
        {
            double result;
            if (double.TryParse(numberString.Replace(',', '.'), out result))
                return result;
            
            return 0;
        }
        
        /// <summary>
        /// Обработка поля "Поставленное количество", которое может содержать формулы через "+"
        /// </summary>
        private double ParseDeliveredQuantity(string quantityString)
        {
            if (string.IsNullOrWhiteSpace(quantityString))
                return 0;
            
            // Если строка содержит "+", разбиваем ее и суммируем части
            if (quantityString.Contains("+"))
            {
                double sum = 0;
                var parts = quantityString.Split('+');
                
                foreach (var part in parts)
                {
                    sum += ParseDouble(part.Trim());
                }
                
                return sum;
            }
            
            return ParseDouble(quantityString);
        }
        
        /// <summary>
        /// Извлечение артикула из наименования товара с использованием регулярных выражений
        /// </summary>
        private string ExtractArticleFromName(string productName)
        {
            if (string.IsNullOrWhiteSpace(productName))
                return string.Empty;
            
            // Пытаемся найти артикул в формате "арт. XXXXX" или "артикул XXXXX"
            var artMatch = Regex.Match(productName, @"(?:арт(?:икул)?\.?\s*)([A-Za-z0-9\-]+)");
            if (artMatch.Success)
                return artMatch.Groups[1].Value;
            
            // Пытаемся найти артикул в формате "XXXXX" (только цифры и буквы)
            var codeMatch = Regex.Match(productName, @"([A-Za-z0-9]{5,})");
            if (codeMatch.Success)
                return codeMatch.Groups[1].Value;
            
            return string.Empty;
        }
        
        /// <summary>
        /// Получение наиболее часто встречающегося наименования товара в группе
        /// </summary>
        private string GetMostFrequentName(IEnumerable<OrderItem> items)
        {
            if (items == null || !items.Any())
                return string.Empty;

            var mostFrequent = items
                .Where(item => !string.IsNullOrEmpty(item.ProductName))
                .GroupBy(item => item.ProductName)
                .OrderByDescending(group => group.Count())
                .FirstOrDefault();

            return mostFrequent?.Key ?? string.Empty;
        }

        /// <summary>
        /// Расчет степени схожести двух строк (упрощенная реализация расстояния Левенштейна)
        /// </summary>
        private double CalculateSimilarity(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
                return 0;
            
            s1 = s1.ToLower();
            s2 = s2.ToLower();
            
            int distance = LevenshteinDistance(s1, s2);
            int maxLength = Math.Max(s1.Length, s2.Length);
            
            return 1.0 - (double)distance / maxLength;
        }

        /// <summary>
        /// Расчет расстояния Левенштейна между двумя строками
        /// </summary>
        private int LevenshteinDistance(string s1, string s2)
        {
            int[,] distance = new int[s1.Length + 1, s2.Length + 1];
            
            for (int i = 0; i <= s1.Length; i++)
                distance[i, 0] = i;
            
            for (int j = 0; j <= s2.Length; j++)
                distance[0, j] = j;
            
            for (int i = 1; i <= s1.Length; i++)
            {
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    
                    distance[i, j] = Math.Min(
                        Math.Min(
                            distance[i - 1, j] + 1,     // удаление
                            distance[i, j - 1] + 1),    // вставка
                        distance[i - 1, j - 1] + cost); // замена
                }
            }
            
            return distance[s1.Length, s2.Length];
        }
        
        /// <summary>
        /// Расчет статистических показателей для унифицированного товара
        /// </summary>
        private void CalculateStatistics(UnifiedProduct product)
        {
            var orderHistory = product.OrderHistory;
            
            if (orderHistory.Count == 0)
                return;
            
            // Сортируем историю заказов по дате
            var sortedHistory = orderHistory.OrderBy(item => item.OrderDate).ToList();
            
            // Рассчитываем средний интервал между заказами
            if (sortedHistory.Count > 1)
            {
                double totalDays = 0;
                for (int i = 1; i < sortedHistory.Count; i++)
                {
                    totalDays += (sortedHistory[i].OrderDate - sortedHistory[i - 1].OrderDate).TotalDays;
                }
                
                product.AverageOrderInterval = totalDays / (sortedHistory.Count - 1);
            }
            
            // Рассчитываем среднее количество в заказе
            product.AverageOrderQuantity = orderHistory.Average(item => item.OrderedQuantity);
            
            // Рассчитываем средний срок поставки
            var itemsWithDelivery = orderHistory
                .Where(item => item.DeliveryDate.HasValue)
                .ToList();
            
            if (itemsWithDelivery.Count > 0)
            {
                double totalDeliveryDays = 0;
                foreach (var item in itemsWithDelivery)
                {
                    if (item.DeliveryDate.HasValue)
                    {
                        totalDeliveryDays += (item.DeliveryDate.Value - item.OrderDate).TotalDays;
                    }
                }
                
                product.AverageDeliveryTime = totalDeliveryDays / itemsWithDelivery.Count;
            }
            
            // Устанавливаем дату последнего заказа
            product.LastOrderDate = sortedHistory.Last().OrderDate;
            
            // Рассчитываем коэффициенты сезонности по месяцам
            CalculateSeasonalityCoefficients(product);
            
            // Прогнозируем дату следующего заказа
            if (product.LastOrderDate.HasValue && product.AverageOrderInterval > 0)
            {
                int lastMonth = product.LastOrderDate.Value.Month - 1; // Индекс от 0 до 11
                double seasonalCoefficient = product.SeasonalityCoefficients[lastMonth];
                
                // Если коэффициент сезонности равен 0, используем 1 (нет данных о сезонности)
                if (seasonalCoefficient == 0)
                    seasonalCoefficient = 1;
                
                double daysToAdd = product.AverageOrderInterval * seasonalCoefficient;
                product.NextPredictedOrderDate = product.LastOrderDate.Value.AddDays(daysToAdd);
                
                // Рассчитываем рекомендуемое количество
                double trendFactor = CalculateTrendFactor(sortedHistory);
                int nextMonth = product.NextPredictedOrderDate.Value.Month - 1;
                double nextSeasonalCoefficient = product.SeasonalityCoefficients[nextMonth];
                
                // Если коэффициент сезонности равен 0, используем 1 (нет данных о сезонности)
                if (nextSeasonalCoefficient == 0)
                    nextSeasonalCoefficient = 1;
                
                product.RecommendedQuantity = product.AverageOrderQuantity * (1 + trendFactor) * nextSeasonalCoefficient;
                
                // Рассчитываем оптимальную дату размещения заказа
                if (product.AverageDeliveryTime > 0)
                {
                    double safetyBuffer = product.AverageDeliveryTime * 0.2; // 20% запаса на непредвиденные задержки
                    product.OptimalOrderPlacementDate = product.NextPredictedOrderDate.Value
                        .AddDays(-product.AverageDeliveryTime - safetyBuffer);
                }
            }
        }
        
        /// <summary>
        /// Расчет коэффициентов сезонности по месяцам
        /// </summary>
        private void CalculateSeasonalityCoefficients(UnifiedProduct product)
        {
            var orderHistory = product.OrderHistory;
            
            if (orderHistory.Count < 12) // Недостаточно данных для надежного расчета сезонности
            {
                // Заполняем коэффициенты значением 1 (нет сезонности)
                for (int i = 0; i < 12; i++)
                {
                    product.SeasonalityCoefficients[i] = 1;
                }
                
                return;
            }
            
            // Группируем заказы по месяцам и рассчитываем среднее количество для каждого месяца
            var monthlyAverages = new double[12];
            var monthCounts = new int[12];
            
            foreach (var order in orderHistory)
            {
                int month = order.OrderDate.Month - 1; // Индекс от 0 до 11
                monthlyAverages[month] += order.OrderedQuantity;
                monthCounts[month]++;
            }
            
            // Рассчитываем среднее количество для каждого месяца
            for (int i = 0; i < 12; i++)
            {
                if (monthCounts[i] > 0)
                {
                    monthlyAverages[i] /= monthCounts[i];
                }
            }
            
            // Рассчитываем общее среднее количество
            double overallAverage = orderHistory.Average(item => item.OrderedQuantity);
            
            // Рассчитываем коэффициенты сезонности
            for (int i = 0; i < 12; i++)
            {
                if (monthCounts[i] > 0 && overallAverage > 0)
                {
                    product.SeasonalityCoefficients[i] = monthlyAverages[i] / overallAverage;
                }
                else
                {
                    product.SeasonalityCoefficients[i] = 1; // Нет данных о сезонности
                }
            }
        }
        
        /// <summary>
        /// Расчет коэффициента тренда на основе истории заказов
        /// </summary>
        private double CalculateTrendFactor(List<OrderItem> sortedHistory)
        {
            if (sortedHistory.Count < 2)
                return 0;
            
            double firstQuantity = sortedHistory.First().OrderedQuantity;
            double lastQuantity = sortedHistory.Last().OrderedQuantity;
            
            // Если первое количество равно 0, используем 1 для избежания деления на 0
            if (firstQuantity == 0)
                firstQuantity = 1;
            
            return (lastQuantity - firstQuantity) / firstQuantity / (sortedHistory.Count - 1);
        }
        
        #endregion
    }
    
    /// <summary>
    /// Вспомогательный класс для десериализации базы соответствий
    /// </summary>
    internal class MappingItem
    {
        public string? UnifiedArticle { get; set; }
        public string? PrimaryName { get; set; }
        public List<string>? NameVariations { get; set; }
        public List<string>? ArticleVariations { get; set; }

        public MappingItem()
        {
            NameVariations = new List<string>();
            ArticleVariations = new List<string>();
        }
    }
}
