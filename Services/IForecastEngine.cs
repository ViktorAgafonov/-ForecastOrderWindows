using System;
using System.Collections.Generic;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Интерфейс для движка прогнозирования
    /// </summary>
    public interface IForecastEngine
    {
        /// <summary>
        /// Прогнозирование дат заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="daysAhead">Количество дней для прогноза вперед</param>
        /// <returns>Список прогнозов по датам</returns>
        List<ForecastResult> ForecastOrderDates(List<UnifiedProduct> unifiedProducts, int daysAhead = 30);
        
        /// <summary>
        /// Прогнозирование объемов заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <returns>Список прогнозов с объемами</returns>
        List<ForecastResult> ForecastOrderVolumes(List<UnifiedProduct> unifiedProducts);
        
        /// <summary>
        /// Определение оптимальных сроков размещения заказов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <returns>Список прогнозов с оптимальными сроками размещения</returns>
        List<ForecastResult> DetermineOptimalOrderPlacementDates(List<UnifiedProduct> unifiedProducts);
        
        /// <summary>
        /// Формирование полных прогнозов на указанный период
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="startDate">Дата начала периода прогнозирования</param>
        /// <param name="endDate">Дата окончания периода прогнозирования</param>
        /// <returns>Список полных прогнозов</returns>
        List<ForecastResult> GenerateFullForecasts(List<UnifiedProduct> unifiedProducts, DateTime startDate, DateTime endDate);
    }
}
