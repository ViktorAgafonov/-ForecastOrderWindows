using System.Collections.Generic;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Интерфейс для анализатора заказов
    /// </summary>
    public interface IOrderAnalyzer
    {
        /// <summary>
        /// Анализ частоты заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        void AnalyzeOrderFrequency(List<UnifiedProduct> unifiedProducts);
        
        /// <summary>
        /// Анализ объемов заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        void AnalyzeOrderVolumes(List<UnifiedProduct> unifiedProducts);
        
        /// <summary>
        /// Анализ сезонности заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        void AnalyzeSeasonality(List<UnifiedProduct> unifiedProducts);
        
        /// <summary>
        /// Анализ сроков поставки
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        void AnalyzeDeliveryTimes(List<UnifiedProduct> unifiedProducts);
    }
}
