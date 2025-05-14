using System;
using System.Collections.Generic;

namespace Forecast.Models
{
    /// <summary>
    /// Класс, представляющий результат прогнозирования заявок
    /// </summary>
    public class ForecastResult
    {
        /// <summary>
        /// Унифицированный артикул товара
        /// </summary>
        public string? UnifiedArticle { get; set; }
        
        /// <summary>
        /// Наименование товара
        /// </summary>
        public string? ProductName { get; set; }
        
        /// <summary>
        /// Прогнозируемая дата следующего заказа
        /// </summary>
        public DateTime NextOrderDate { get; set; }
        
        /// <summary>
        /// Рекомендуемое количество для заказа
        /// </summary>
        public double RecommendedQuantity { get; set; }
        
        /// <summary>
        /// Оптимальная дата размещения заказа
        /// </summary>
        public DateTime OptimalOrderPlacementDate { get; set; }
        
        /// <summary>
        /// Уровень приоритета заказа (1-5, где 1 - наивысший)
        /// </summary>
        public int Priority { get; set; }
        
        /// <summary>
        /// Уверенность в прогнозе (0-100%)
        /// </summary>
        public double Confidence { get; set; }
        
        /// <summary>
        /// Дополнительные комментарии к прогнозу
        /// </summary>
        public string Notes { get; set; }
    }
}
