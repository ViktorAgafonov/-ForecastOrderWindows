using System;
using System.Collections.Generic;
using System.Linq;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Класс для прогнозирования заявок
    /// </summary>
    public class ForecastEngine : IForecastEngine
    {
        /// <summary>
        /// Прогнозирование дат заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="daysAhead">Количество дней для прогноза вперед</param>
        /// <returns>Список прогнозов по датам</returns>
        public List<ForecastResult> ForecastOrderDates(List<UnifiedProduct> unifiedProducts, int daysAhead = 30)
        {
            var forecasts = new List<ForecastResult>();
            DateTime endDate = DateTime.Now.AddDays(daysAhead);
            
            foreach (var product in unifiedProducts)
            {
                // Пропускаем товары без прогноза даты следующего заказа
                if (!product.NextPredictedOrderDate.HasValue)
                    continue;
                
                // Пропускаем товары, у которых прогнозируемая дата выходит за пределы указанного периода
                if (product.NextPredictedOrderDate.Value > endDate)
                    continue;
                
                var forecast = new ForecastResult
                {
                    UnifiedArticle = product.UnifiedArticle,
                    ProductName = product.PrimaryName,
                    NextOrderDate = product.NextPredictedOrderDate.Value,
                    RecommendedQuantity = product.RecommendedQuantity,
                    OptimalOrderPlacementDate = product.OptimalOrderPlacementDate ?? DateTime.Now,
                    // Расчет приоритета на основе близости даты размещения заказа к текущей дате
                    Priority = CalculatePriority(product.OptimalOrderPlacementDate ?? DateTime.Now),
                    // Расчет уверенности в прогнозе на основе количества исторических данных
                    Confidence = CalculateConfidence(product.OrderHistory.Count)
                };
                
                forecasts.Add(forecast);
            }
            
            // Сортируем прогнозы по приоритету (от высокого к низкому)
            return forecasts.OrderBy(f => f.Priority).ToList();
        }
        
        /// <summary>
        /// Прогнозирование объемов заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <returns>Список прогнозов с объемами</returns>
        public List<ForecastResult> ForecastOrderVolumes(List<UnifiedProduct> unifiedProducts)
        {
            var forecasts = new List<ForecastResult>();
            
            foreach (var product in unifiedProducts)
            {
                // Пропускаем товары без прогноза даты следующего заказа
                if (!product.NextPredictedOrderDate.HasValue)
                    continue;
                
                var forecast = new ForecastResult
                {
                    UnifiedArticle = product.UnifiedArticle,
                    ProductName = product.PrimaryName,
                    NextOrderDate = product.NextPredictedOrderDate.Value,
                    RecommendedQuantity = product.RecommendedQuantity,
                    OptimalOrderPlacementDate = product.OptimalOrderPlacementDate ?? DateTime.Now,
                    Priority = CalculatePriority(product.OptimalOrderPlacementDate ?? DateTime.Now),
                    Confidence = CalculateConfidence(product.OrderHistory.Count)
                };
                
                // Добавляем комментарий о тренде изменения объемов заказов
                forecast.Notes = GenerateVolumeNotes(product);
                
                forecasts.Add(forecast);
            }
            
            // Сортируем прогнозы по рекомендуемому количеству (от большего к меньшему)
            return forecasts.OrderByDescending(f => f.RecommendedQuantity).ToList();
        }
        
        /// <summary>
        /// Определение оптимальных сроков размещения заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <returns>Список прогнозов с оптимальными сроками размещения</returns>
        public List<ForecastResult> DetermineOptimalOrderPlacementDates(List<UnifiedProduct> unifiedProducts)
        {
            var forecasts = new List<ForecastResult>();
            
            foreach (var product in unifiedProducts)
            {
                // Пропускаем товары без прогноза даты следующего заказа или оптимальной даты размещения
                if (!product.NextPredictedOrderDate.HasValue || !product.OptimalOrderPlacementDate.HasValue)
                    continue;
                
                var forecast = new ForecastResult
                {
                    UnifiedArticle = product.UnifiedArticle,
                    ProductName = product.PrimaryName,
                    NextOrderDate = product.NextPredictedOrderDate.Value,
                    RecommendedQuantity = product.RecommendedQuantity,
                    OptimalOrderPlacementDate = product.OptimalOrderPlacementDate.Value,
                    Priority = CalculatePriority(product.OptimalOrderPlacementDate.Value),
                    Confidence = CalculateConfidence(product.OrderHistory.Count)
                };
                
                // Добавляем комментарий о сроках поставки
                forecast.Notes = GenerateDeliveryNotes(product);
                
                forecasts.Add(forecast);
            }
            
            // Сортируем прогнозы по оптимальной дате размещения (от ранней к поздней)
            return forecasts.OrderBy(f => f.OptimalOrderPlacementDate).ToList();
        }
        
        /// <summary>
        /// Формирование полных прогнозов на указанный период
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="startDate">Дата начала периода прогнозирования</param>
        /// <param name="endDate">Дата окончания периода прогнозирования</param>
        /// <returns>Список полных прогнозов</returns>
        public List<ForecastResult> GenerateFullForecasts(List<UnifiedProduct> unifiedProducts, DateTime startDate, DateTime endDate)
        {
            var forecasts = new List<ForecastResult>();
            int daysAhead = (int)(endDate - startDate).TotalDays;
            
            // Сначала прогнозируем даты заказов для всех товаров
            foreach (var product in unifiedProducts)
            {
                // Пропускаем товары без истории заказов
                if (product.OrderHistory == null || product.OrderHistory.Count == 0)
                    continue;
                
                // Обновляем прогнозируемую дату следующего заказа для периода
                // Если дата уже установлена и находится в будущем, используем её
                if (!product.NextPredictedOrderDate.HasValue || product.NextPredictedOrderDate.Value < startDate)
                {
                    // Вычисляем среднее время между заказами
                    if (product.OrderHistory.Count >= 2)
                    {
                        var sortedHistory = product.OrderHistory.OrderBy(h => h.OrderDate).ToList();
                        double totalDays = (sortedHistory.Last().OrderDate - sortedHistory.First().OrderDate).TotalDays;
                        double averageInterval = totalDays / (sortedHistory.Count - 1);
                        
                        // Прогнозируем дату следующего заказа от последнего известного
                        var lastOrderDate = sortedHistory.Last().OrderDate;
                        product.NextPredictedOrderDate = lastOrderDate.AddDays(averageInterval);
                        
                        // Если прогнозируемая дата в прошлом, корректируем на текущую дату
                        if (product.NextPredictedOrderDate < startDate)
                        {
                            // Вычисляем, сколько интервалов нужно добавить, чтобы дата была в будущем
                            double daysSinceLastOrder = (startDate - lastOrderDate).TotalDays;
                            int intervalsToAdd = (int)Math.Ceiling(daysSinceLastOrder / averageInterval);
                            product.NextPredictedOrderDate = lastOrderDate.AddDays(intervalsToAdd * averageInterval);
                        }
                        
                        // Рассчитываем рекомендуемый объем заказа
                        product.RecommendedQuantity = product.OrderHistory.Average(h => h.OrderedQuantity);
                        
                        // Рассчитываем оптимальную дату размещения заказа
                        if (product.AverageDeliveryTime > 0)
                        {
                            double safetyFactor = 0.2; // 20% запас по времени
                            double leadTime = product.AverageDeliveryTime * (1 + safetyFactor);
                            product.OptimalOrderPlacementDate = product.NextPredictedOrderDate.Value.AddDays(-leadTime);
                        }
                        else
                        {
                            product.OptimalOrderPlacementDate = product.NextPredictedOrderDate;
                        }
                    }
                }
                
                // Добавляем прогноз, если дата следующего заказа в пределах периода
                if (product.NextPredictedOrderDate.HasValue && 
                    product.NextPredictedOrderDate.Value >= startDate && 
                    product.NextPredictedOrderDate.Value <= endDate)
                {
                    var forecast = new ForecastResult
                    {
                        UnifiedArticle = product.UnifiedArticle,
                        ProductName = product.PrimaryName,
                        NextOrderDate = product.NextPredictedOrderDate.Value,
                        RecommendedQuantity = product.RecommendedQuantity,
                        OptimalOrderPlacementDate = product.OptimalOrderPlacementDate ?? DateTime.Now,
                        Priority = CalculatePriority(product.OptimalOrderPlacementDate ?? DateTime.Now),
                        Confidence = CalculateConfidence(product.OrderHistory.Count)
                    };
                    
                    // Формируем подробные комментарии для прогноза
                    forecast.Notes = GenerateDetailedNotes(product);
                    
                    forecasts.Add(forecast);
                    
                    // Прогнозируем следующие заказы в пределах периода
                    DateTime nextDate = product.NextPredictedOrderDate.Value;
                    double interval = product.AverageOrderInterval;
                    
                    // Если интервал не определен, используем стандартное значение
                    if (interval <= 0) interval = 30;
                    
                    // Добавляем прогнозы на следующие даты в пределах периода
                    for (int i = 1; i <= 5; i++) // Ограничиваем количество прогнозов для каждого товара
                    {
                        nextDate = nextDate.AddDays(interval);
                        if (nextDate > endDate) break;
                        
                        var nextForecast = new ForecastResult
                        {
                            UnifiedArticle = product.UnifiedArticle,
                            ProductName = product.PrimaryName,
                            NextOrderDate = nextDate,
                            RecommendedQuantity = product.RecommendedQuantity,
                            OptimalOrderPlacementDate = nextDate.AddDays(-(product.AverageDeliveryTime > 0 ? 
                                product.AverageDeliveryTime * 1.2 : 0)),
                            Priority = CalculatePriority(nextDate.AddDays(-(product.AverageDeliveryTime > 0 ? 
                                product.AverageDeliveryTime * 1.2 : 0))),
                            Confidence = Math.Max(50, CalculateConfidence(product.OrderHistory.Count) - i * 10) // Снижаем уверенность для более отдаленных прогнозов
                        };
                        
                        nextForecast.Notes = $"Прогноз #{i+1} для товара. " + GenerateDetailedNotes(product) + 
                                          $" Уверенность снижена из-за отдаленности прогноза.";
                        
                        forecasts.Add(nextForecast);
                    }
                }
            }
            
            // Сортируем прогнозы по приоритету и оптимальной дате размещения
            return forecasts
                .OrderBy(f => f.Priority)
                .ThenBy(f => f.OptimalOrderPlacementDate)
                .ToList();
        }
        
        #region Вспомогательные методы
        
        /// <summary>
        /// Расчет приоритета заказа на основе близости даты размещения к текущей дате
        /// </summary>
        private int CalculatePriority(DateTime optimalOrderPlacementDate)
        {
            // Рассчитываем количество дней до оптимальной даты размещения
            double daysUntilPlacement = (optimalOrderPlacementDate - DateTime.Now).TotalDays;
            
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
        /// Расчет уверенности в прогнозе на основе количества исторических данных
        /// </summary>
        private double CalculateConfidence(int historyCount)
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
        /// Генерация комментария о тренде изменения объемов заказов
        /// </summary>
        private string GenerateVolumeNotes(UnifiedProduct product)
        {
            // Сортируем историю заказов по дате
            var sortedHistory = product.OrderHistory
                .OrderBy(item => item.OrderDate)
                .ToList();
            
            if (sortedHistory.Count < 2)
                return "Недостаточно данных для анализа тренда объемов заказов.";
            
            double firstQuantity = sortedHistory.First().OrderedQuantity;
            double lastQuantity = sortedHistory.Last().OrderedQuantity;
            
            // Рассчитываем относительное изменение
            double change = (lastQuantity - firstQuantity) / Math.Max(1, firstQuantity);
            
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
        private string GenerateDeliveryNotes(UnifiedProduct product)
        {
            if (product.AverageDeliveryTime <= 0)
                return "Нет данных о сроках поставки.";
            
            return $"Средний срок поставки: {Math.Round(product.AverageDeliveryTime)} дней. " +
                   $"Запас времени: {Math.Round(product.AverageDeliveryTime * 0.2)} дней.";
        }
        
        /// <summary>
        /// Генерация подробных комментариев для прогноза
        /// </summary>
        private string GenerateDetailedNotes(UnifiedProduct product)
        {
            var notes = new List<string>();
            
            // Добавляем информацию о тренде объемов заказов
            notes.Add(GenerateVolumeNotes(product));
            
            // Добавляем информацию о сроках поставки
            if (product.AverageDeliveryTime > 0)
            {
                notes.Add(GenerateDeliveryNotes(product));
            }
            
            // Добавляем информацию о сезонности
            int nextMonth = product.NextPredictedOrderDate.Value.Month - 1;
            double seasonalCoefficient = product.SeasonalityCoefficients[nextMonth];
            
            if (seasonalCoefficient != 1) // Если есть сезонность
            {
                string seasonalityDirection = seasonalCoefficient > 1 ? "повышение" : "снижение";
                notes.Add($"Сезонный фактор: {seasonalityDirection} на {Math.Round(Math.Abs(seasonalCoefficient - 1) * 100)}% в {product.NextPredictedOrderDate.Value.ToString("MMMM")}."); 
            }
            
            // Добавляем информацию о частоте заказов
            if (product.AverageOrderInterval > 0)
            {
                notes.Add($"Средний интервал между заказами: {Math.Round(product.AverageOrderInterval)} дней.");
            }
            
            return string.Join(" ", notes);
        }
        
        #endregion
    }
}
