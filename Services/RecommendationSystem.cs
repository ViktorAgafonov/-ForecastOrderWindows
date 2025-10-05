using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Класс для формирования рекомендаций по заказам
    /// </summary>
    public class RecommendationSystem : IRecommendationSystem
    {
        /// <summary>
        /// Генерация списка товаров для заказа на указанный период
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="startDate">Дата начала периода</param>
        /// <param name="endDate">Дата окончания периода</param>
        /// <returns>Список товаров для заказа с рекомендациями</returns>
        public List<ForecastResult> GenerateOrderRecommendations(List<UnifiedProduct> unifiedProducts, DateTime startDate, DateTime endDate)
        {
            var recommendations = new List<ForecastResult>();
            
            foreach (var product in unifiedProducts)
            {
                // Пропускаем товары без прогноза даты следующего заказа
                if (!product.NextPredictedOrderDate.HasValue)
                    continue;
                
                // Пропускаем товары, у которых оптимальная дата размещения заказа выходит за пределы указанного периода
                if (product.OptimalOrderPlacementDate.HasValue && 
                    (product.OptimalOrderPlacementDate.Value < startDate || product.OptimalOrderPlacementDate.Value > endDate))
                    continue;
                
                var forecast = new ForecastResult
                {
                    UnifiedArticle = product.UnifiedArticle,
                    ProductName = product.PrimaryName,
                    NextOrderDate = product.NextPredictedOrderDate.Value,
                    RecommendedQuantity = product.RecommendedQuantity,
                    OptimalOrderPlacementDate = product.OptimalOrderPlacementDate ?? DateTime.Now,
                    // Расчет приоритета на основе близости даты размещения заказа к текущей дате
                    Priority = CalculatePriority(product.OptimalOrderPlacementDate ?? DateTime.Now, startDate),
                    // Расчет уверенности в прогнозе на основе количества исторических данных и их разброса
                    Confidence = CalculateConfidence(product)
                };
                
                // Формируем подробные комментарии для рекомендации
                forecast.Notes = GenerateDetailedNotes(product);
                
                recommendations.Add(forecast);
            }
            
            // Сортируем рекомендации по приоритету (от высокого к низкому) и дате размещения
            return recommendations
                .OrderBy(r => r.Priority)
                .ThenBy(r => r.OptimalOrderPlacementDate)
                .ToList();
        }
        
        /// <summary>
        /// Формирование календаря будущих заказов
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <returns>Календарь заказов по датам</returns>
        public Dictionary<DateTime, List<ForecastResult>> CreateOrderCalendar(List<ForecastResult> forecasts)
        {
            var calendar = new Dictionary<DateTime, List<ForecastResult>>();
            
            // Группируем прогнозы по дате размещения заказа
            var groupedByDate = forecasts
                .GroupBy(f => f.OptimalOrderPlacementDate.Date);
            
            foreach (var group in groupedByDate)
            {
                calendar[group.Key] = group.ToList();
            }
            
            return calendar;
        }
        
        /// <summary>
        /// Расчет приоритетов заказов
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <returns>Список прогнозов с рассчитанными приоритетами</returns>
        public List<ForecastResult> CalculateOrderPriorities(List<ForecastResult> forecasts)
        {
            // Рассчитываем приоритеты на основе нескольких факторов
            foreach (var forecast in forecasts)
            {
                // Базовый приоритет на основе близости даты размещения к текущей дате
                int basePriority = CalculatePriority(forecast.OptimalOrderPlacementDate, DateTime.Now);
                
                // Корректируем приоритет на основе уверенности в прогнозе
                if (forecast.Confidence < 50)
                {
                    // Снижаем приоритет для прогнозов с низкой уверенностью
                    basePriority = Math.Min(5, basePriority + 1);
                }
                
                forecast.Priority = basePriority;
            }
            
            return forecasts;
        }
        
        /// <summary>
        /// Группировка заказов по оптимальным партиям
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <returns>Сгруппированные заказы</returns>
        public List<List<ForecastResult>> GroupOrdersByBatches(List<ForecastResult> forecasts)
        {
            var batches = new List<List<ForecastResult>>();
            
            // Группируем заказы по близким датам (в пределах 3 дней)
            var sortedForecasts = forecasts
                .OrderBy(f => f.OptimalOrderPlacementDate)
                .ToList();
            
            if (sortedForecasts.Count == 0)
                return batches;
            
            var currentBatch = new List<ForecastResult> { sortedForecasts[0] };
            var currentDate = sortedForecasts[0].OptimalOrderPlacementDate;
            
            for (int i = 1; i < sortedForecasts.Count; i++)
            {
                var forecast = sortedForecasts[i];
                
                // Если дата размещения близка к текущей дате группы, добавляем в текущую группу
                if ((forecast.OptimalOrderPlacementDate - currentDate).TotalDays <= 3)
                {
                    currentBatch.Add(forecast);
                }
                else
                {
                    // Иначе создаем новую группу
                    batches.Add(currentBatch);
                    currentBatch = new List<ForecastResult> { forecast };
                    currentDate = forecast.OptimalOrderPlacementDate;
                }
            }
            
            // Добавляем последнюю группу
            if (currentBatch.Count > 0)
            {
                batches.Add(currentBatch);
            }
            
            return batches;
        }
        
        /// <summary>
        /// Сохранение прогнозов и рекомендаций в файл
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        public void SaveRecommendations(List<ForecastResult> forecasts, string filePath)
        {
            try
            {
                // Сериализуем прогнозы в JSON и сохраняем в файл
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(forecasts, options);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении рекомендаций: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Загрузка сохраненных рекомендаций из файла
        /// </summary>
        /// <param name="filePath">Путь к файлу с сохраненными прогнозами</param>
        /// <returns>Список прогнозов</returns>
        public List<ForecastResult> LoadRecommendations(string filePath)
        {
            var forecasts = new List<ForecastResult>();
            
            try
            {
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    forecasts = JsonSerializer.Deserialize<List<ForecastResult>>(json) ?? new List<ForecastResult>();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке рекомендаций: {ex.Message}");
            }
            
            return forecasts;
        }
        
        #region Вспомогательные методы
        
        /// <summary>
        /// Расчет приоритета заказа на основе близости даты размещения к указанной дате
        /// </summary>
        private int CalculatePriority(DateTime optimalOrderPlacementDate, DateTime referenceDate)
        {
            // Рассчитываем количество дней до оптимальной даты размещения
            double daysUntilPlacement = (optimalOrderPlacementDate - referenceDate).TotalDays;
            
            // Определяем приоритет на основе количества дней
            if (daysUntilPlacement < 0) // Дата размещения уже прошла
                return 1; // Наивысший приоритет
            else if (daysUntilPlacement < 7) // До 7 дней
                return 2; // Высокий приоритет
            else if (daysUntilPlacement < 14) // До 14 дней
                return 3; // Средний приоритет
            else if (daysUntilPlacement < 21) // До 21 дня
                return 4; // Низкий приоритет
            else // Более 21 дня
                return 5; // Самый низкий приоритет
        }
        
        /// <summary>
        /// Расчет уверенности в прогнозе на основе количества исторических данных и их разброса
        /// </summary>
        private double CalculateConfidence(UnifiedProduct product)
        {
            // Базовая уверенность на основе количества исторических данных
            double baseConfidence = CalculateBaseConfidence(product.OrderHistory.Count);
            
            // Корректировка уверенности на основе разброса данных
            double variationFactor = CalculateVariationFactor(product);
            
            // Итоговая уверенность с учетом разброса (не более 95%)
            return Math.Min(95, baseConfidence * (1 - variationFactor));
        }
        
        /// <summary>
        /// Расчет базовой уверенности на основе количества исторических данных
        /// </summary>
        private double CalculateBaseConfidence(int historyCount)
        {
            // Определяем уверенность на основе количества исторических данных
            if (historyCount < 3)
                return 30; // Низкая уверенность
            else if (historyCount < 5)
                return 50; // Средняя уверенность
            else if (historyCount < 10)
                return 70; // Хорошая уверенность
            else if (historyCount < 20)
                return 85; // Высокая уверенность
            else
                return 95; // Очень высокая уверенность
        }
        
        /// <summary>
        /// Расчет фактора вариации данных
        /// </summary>
        private double CalculateVariationFactor(UnifiedProduct product)
        {
            var orderHistory = product.OrderHistory;
            
            if (orderHistory.Count < 2)
                return 0.5; // Средний фактор вариации при недостатке данных
            
            // Рассчитываем коэффициент вариации для интервалов между заказами
            var sortedHistory = orderHistory.OrderBy(item => item.OrderDate).ToList();
            var intervals = new List<double>();
            
            for (int i = 1; i < sortedHistory.Count; i++)
            {
                double days = (sortedHistory[i].OrderDate - sortedHistory[i - 1].OrderDate).TotalDays;
                intervals.Add(days);
            }
            
            if (intervals.Count == 0)
                return 0.5;
            
            // Рассчитываем среднее и стандартное отклонение
            double mean = intervals.Average();
            
            if (mean == 0)
                return 0.5;
            
            double sumOfSquaresOfDifferences = intervals.Sum(val => Math.Pow(val - mean, 2));
            double standardDeviation = Math.Sqrt(sumOfSquaresOfDifferences / intervals.Count);
            
            // Коэффициент вариации (отношение стандартного отклонения к среднему)
            double coefficientOfVariation = standardDeviation / mean;
            
            // Нормализуем коэффициент вариации в диапазон [0, 0.5]
            return Math.Min(0.5, coefficientOfVariation / 2);
        }
        
        /// <summary>
        /// Генерация подробных комментариев для рекомендации
        /// </summary>
        private string GenerateDetailedNotes(UnifiedProduct product)
        {
            var notes = new List<string>();
            
            // Добавляем информацию о тренде объемов заказов
            notes.Add(GenerateVolumeTrendNote(product));
            
            // Добавляем информацию о сроках поставки
            if (product.AverageDeliveryTime > 0)
            {
                notes.Add(GenerateDeliveryTimeNote(product));
            }
            
            // Добавляем информацию о сезонности
            if (product.NextPredictedOrderDate.HasValue)
            {
                int nextMonth = product.NextPredictedOrderDate.Value.Month - 1;
                double seasonalCoefficient = product.SeasonalityCoefficients[nextMonth];
                
                if (Math.Abs(seasonalCoefficient - 1) > 0.1) // Если есть значимая сезонность
                {
                    string seasonalityDirection = seasonalCoefficient > 1 ? "повышение" : "снижение";
                    notes.Add($"Сезонный фактор: {seasonalityDirection} на {Math.Round(Math.Abs(seasonalCoefficient - 1) * 100)}% в {product.NextPredictedOrderDate.Value.ToString("MMMM")}."); 
                }
            }
            
            // Добавляем информацию о частоте заказов
            if (product.AverageOrderInterval > 0)
            {
                notes.Add($"Средний интервал между заказами: {Math.Round(product.AverageOrderInterval)} дней.");
            }
            
            return string.Join(" ", notes);
        }
        
        /// <summary>
        /// Генерация комментария о тренде объемов заказов
        /// </summary>
        private string GenerateVolumeTrendNote(UnifiedProduct product)
        {
            // Сортируем историю заказов по дате
            var sortedHistory = product.OrderHistory
                .OrderBy(item => item.OrderDate)
                .ToList();
            
            if (sortedHistory.Count < 2)
                return "Недостаточно данных для анализа тренда объемов заказов.";
            
            double firstQuantity = sortedHistory.First().OrderedQuantity;
            double lastQuantity = sortedHistory.Last().OrderedQuantity;
            
            // Если первое количество равно 0, используем 1 для избежания деления на 0
            if (firstQuantity == 0)
                firstQuantity = 1;
            
            // Рассчитываем относительное изменение
            double change = (lastQuantity - firstQuantity) / firstQuantity;
            
            if (Math.Abs(change) < 0.1) // Изменение менее 10%
                return "Объемы заказов стабильны.";
            else if (change > 0)
                return $"Наблюдается рост объемов заказов на {Math.Round(change * 100)}%.";
            else
                return $"Наблюдается снижение объемов заказов на {Math.Round(Math.Abs(change) * 100)}%.";
        }
        
        /// <summary>
        /// Генерация комментария о сроках поставки
        /// </summary>
        private string GenerateDeliveryTimeNote(UnifiedProduct product)
        {
            return $"Средний срок поставки: {Math.Round(product.AverageDeliveryTime)} дней. " +
                   $"Запас времени: {Math.Round(product.AverageDeliveryTime * 0.2)} дней.";
        }
        
        #endregion
    }
}
