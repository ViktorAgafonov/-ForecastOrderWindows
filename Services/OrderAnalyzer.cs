using System;
using System.Collections.Generic;
using System.Linq;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Класс для анализа заказов
    /// </summary>
    public class OrderAnalyzer : IOrderAnalyzer
    {
        /// <summary>
        /// Анализ частоты заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        public void AnalyzeOrderFrequency(List<UnifiedProduct> unifiedProducts)
        {
            foreach (var product in unifiedProducts)
            {
                // Сортируем историю заказов по дате
                var sortedHistory = product.OrderHistory
                    .OrderBy(item => item.OrderDate)
                    .ToList();
                
                if (sortedHistory.Count < 2)
                    continue;
                
                // Рассчитываем интервалы между заказами
                var intervals = new List<double>();
                for (int i = 1; i < sortedHistory.Count; i++)
                {
                    double days = (sortedHistory[i].OrderDate - sortedHistory[i - 1].OrderDate).TotalDays;
                    intervals.Add(days);
                }
                
                // Удаляем выбросы (значения, сильно отличающиеся от медианы)
                var filteredIntervals = FilterOutliers(intervals);
                
                // Рассчитываем средний интервал между заказами
                if (filteredIntervals.Count > 0)
                {
                    product.AverageOrderInterval = filteredIntervals.Average();
                }
                else
                {
                    product.AverageOrderInterval = intervals.Average();
                }
                
                // Устанавливаем дату последнего заказа
                product.LastOrderDate = sortedHistory.Last().OrderDate;
                
                // Прогнозируем дату следующего заказа
                if (product.LastOrderDate.HasValue && product.AverageOrderInterval > 0)
                {
                    // Получаем коэффициент сезонности для текущего месяца
                    int currentMonth = DateTime.Now.Month - 1; // Индекс от 0 до 11
                    double seasonalCoefficient = product.SeasonalityCoefficients[currentMonth];
                    
                    // Если коэффициент сезонности равен 0, используем 1 (нет данных о сезонности)
                    if (seasonalCoefficient == 0)
                        seasonalCoefficient = 1;
                    
                    // Рассчитываем дату следующего заказа с учетом сезонности
                    double daysToAdd = product.AverageOrderInterval * seasonalCoefficient;
                    product.NextPredictedOrderDate = product.LastOrderDate.Value.AddDays(daysToAdd);
                }
            }
        }
        
        /// <summary>
        /// Анализ объемов заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        public void AnalyzeOrderVolumes(List<UnifiedProduct> unifiedProducts)
        {
            foreach (var product in unifiedProducts)
            {
                var orderHistory = product.OrderHistory;
                
                if (orderHistory.Count == 0)
                    continue;
                
                // Сортируем историю заказов по дате
                var sortedHistory = orderHistory
                    .OrderBy(item => item.OrderDate)
                    .ToList();
                
                // Рассчитываем среднее количество в заказе
                product.AverageOrderQuantity = orderHistory.Average(item => item.OrderedQuantity);
                
                // Рассчитываем тренд изменения объемов заказов
                double trendFactor = CalculateTrendFactor(sortedHistory);
                
                // Если есть прогноз даты следующего заказа, рассчитываем рекомендуемое количество
                if (product.NextPredictedOrderDate.HasValue)
                {
                    int nextMonth = product.NextPredictedOrderDate.Value.Month - 1;
                    double nextSeasonalCoefficient = product.SeasonalityCoefficients[nextMonth];
                    
                    // Если коэффициент сезонности равен 0, используем 1 (нет данных о сезонности)
                    if (nextSeasonalCoefficient == 0)
                        nextSeasonalCoefficient = 1;
                    
                    // Рассчитываем рекомендуемое количество с учетом тренда и сезонности
                    product.RecommendedQuantity = product.AverageOrderQuantity * (1 + trendFactor) * nextSeasonalCoefficient;
                    
                    // Округляем до целого числа
                    product.RecommendedQuantity = Math.Ceiling(product.RecommendedQuantity);
                }
            }
        }
        
        /// <summary>
        /// Анализ сезонности заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        public void AnalyzeSeasonality(List<UnifiedProduct> unifiedProducts)
        {
            foreach (var product in unifiedProducts)
            {
                var orderHistory = product.OrderHistory;
                
                if (orderHistory.Count < 12) // Недостаточно данных для надежного расчета сезонности
                {
                    // Заполняем коэффициенты значением 1 (нет сезонности)
                    for (int i = 0; i < 12; i++)
                    {
                        product.SeasonalityCoefficients[i] = 1;
                    }
                    
                    continue;
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
        }
        
        /// <summary>
        /// Анализ сроков поставки
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        public void AnalyzeDeliveryTimes(List<UnifiedProduct> unifiedProducts)
        {
            foreach (var product in unifiedProducts)
            {
                var orderHistory = product.OrderHistory;
                
                // Отбираем только заказы с указанной датой поставки
                var itemsWithDelivery = orderHistory
                    .Where(item => item.DeliveryDate.HasValue)
                    .ToList();
                
                if (itemsWithDelivery.Count == 0)
                {
                    // Если нет данных о сроках поставки, используем стандартное значение (например, 14 дней)
                    product.AverageDeliveryTime = 14;
                    continue;
                }
                
                // Рассчитываем сроки поставки для каждого заказа
                var deliveryTimes = new List<double>();
                foreach (var item in itemsWithDelivery)
                {
                    if (item.DeliveryDate.HasValue)
                    {
                        double days = (item.DeliveryDate.Value - item.OrderDate).TotalDays;
                        deliveryTimes.Add(days);
                    }
                }
                
                // Удаляем выбросы
                var filteredDeliveryTimes = FilterOutliers(deliveryTimes);
                
                // Рассчитываем средний срок поставки
                if (filteredDeliveryTimes.Count > 0)
                {
                    product.AverageDeliveryTime = filteredDeliveryTimes.Average();
                }
                else
                {
                    product.AverageDeliveryTime = deliveryTimes.Average();
                }
                
                // Рассчитываем оптимальную дату размещения заказа
                if (product.NextPredictedOrderDate.HasValue && product.AverageDeliveryTime > 0)
                {
                    double safetyBuffer = product.AverageDeliveryTime * 0.2; // 20% запаса на непредвиденные задержки
                    product.OptimalOrderPlacementDate = product.NextPredictedOrderDate.Value
                        .AddDays(-product.AverageDeliveryTime - safetyBuffer);
                }
            }
        }
        
        #region Вспомогательные методы
        
        /// <summary>
        /// Удаление выбросов из списка значений с использованием метода IQR
        /// </summary>
        private List<double> FilterOutliers(List<double> values)
        {
            if (values.Count < 4) // Недостаточно данных для надежного определения выбросов
                return values;

            // Сортируем значения
            var sortedValues = values.OrderBy(v => v).ToList();

            // Правильный расчет квартилей
            int n = sortedValues.Count;

            // Q1 - первый квартиль (25-й перцентиль)
            double q1Index = (n - 1) * 0.25;
            int q1Lower = (int)Math.Floor(q1Index);
            int q1Upper = (int)Math.Ceiling(q1Index);
            double q1 = q1Lower == q1Upper ? sortedValues[q1Lower] :
                        sortedValues[q1Lower] + (sortedValues[q1Upper] - sortedValues[q1Lower]) * (q1Index - q1Lower);

            // Q3 - третий квартиль (75-й перцентиль)
            double q3Index = (n - 1) * 0.75;
            int q3Lower = (int)Math.Floor(q3Index);
            int q3Upper = (int)Math.Ceiling(q3Index);
            double q3 = q3Lower == q3Upper ? sortedValues[q3Lower] :
                        sortedValues[q3Lower] + (sortedValues[q3Upper] - sortedValues[q3Lower]) * (q3Index - q3Lower);

            // Рассчитываем межквартильный размах
            double iqr = q3 - q1;

            // Определяем границы для выбросов (метод Тьюки)
            double lowerBound = q1 - 1.5 * iqr;
            double upperBound = q3 + 1.5 * iqr;

            // Отфильтровываем выбросы
            return sortedValues.Where(v => v >= lowerBound && v <= upperBound).ToList();
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
            
            // Рассчитываем относительное изменение на один заказ
            return (lastQuantity - firstQuantity) / firstQuantity / (sortedHistory.Count - 1);
        }
        
        #endregion
    }
}
