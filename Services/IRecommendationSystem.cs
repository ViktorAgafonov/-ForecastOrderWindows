using System;
using System.Collections.Generic;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Интерфейс для системы формирования рекомендаций по заказам
    /// </summary>
    public interface IRecommendationSystem
    {
        /// <summary>
        /// Генерация списка товаров для заказа на указанный период
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="startDate">Дата начала периода</param>
        /// <param name="endDate">Дата окончания периода</param>
        /// <returns>Список товаров для заказа с рекомендациями</returns>
        List<ForecastResult> GenerateOrderRecommendations(List<UnifiedProduct> unifiedProducts, DateTime startDate, DateTime endDate);
        
        /// <summary>
        /// Формирование календаря будущих заказов
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <returns>Календарь заказов по датам</returns>
        Dictionary<DateTime, List<ForecastResult>> CreateOrderCalendar(List<ForecastResult> forecasts);
        
        /// <summary>
        /// Расчет приоритетов заказов
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <returns>Список прогнозов с рассчитанными приоритетами</returns>
        List<ForecastResult> CalculateOrderPriorities(List<ForecastResult> forecasts);
        
        /// <summary>
        /// Группировка заказов по оптимальным партиям
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <returns>Сгруппированные заказы</returns>
        List<List<ForecastResult>> GroupOrdersByBatches(List<ForecastResult> forecasts);
        
        /// <summary>
        /// Сохранение прогнозов и рекомендаций в файл
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        void SaveRecommendations(List<ForecastResult> forecasts, string filePath);
        
        /// <summary>
        /// Загрузка прогнозов и рекомендаций из файла
        /// </summary>
        /// <param name="filePath">Путь к файлу с сохраненными прогнозами</param>
        /// <returns>Список прогнозов</returns>
        List<ForecastResult> LoadRecommendations(string filePath);
    }
}
