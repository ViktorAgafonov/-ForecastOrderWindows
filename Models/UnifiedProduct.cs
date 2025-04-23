using System;
using System.Collections.Generic;

namespace Forecast.Models
{
    /// <summary>
    /// Класс, представляющий унифицированный товар с группировкой всех вариаций наименований и артикулов
    /// </summary>
    public class UnifiedProduct
    {
        /// <summary>
        /// Унифицированный артикул товара
        /// </summary>
        public string UnifiedArticle { get; set; }
        
        /// <summary>
        /// Основное наименование товара
        /// </summary>
        public string PrimaryName { get; set; }
        
        /// <summary>
        /// Список всех вариаций наименований товара
        /// </summary>
        public List<string> NameVariations { get; set; } = new List<string>();
        
        /// <summary>
        /// Список всех вариаций артикулов товара
        /// </summary>
        public List<string> ArticleVariations { get; set; } = new List<string>();
        
        /// <summary>
        /// История заказов данного товара
        /// </summary>
        public List<OrderItem> OrderHistory { get; set; } = new List<OrderItem>();
        
        /// <summary>
        /// Средний интервал между заказами (в днях)
        /// </summary>
        public double AverageOrderInterval { get; set; }
        
        /// <summary>
        /// Среднее количество товара в заказе
        /// </summary>
        public double AverageOrderQuantity { get; set; }
        
        /// <summary>
        /// Средний срок поставки (в днях)
        /// </summary>
        public double AverageDeliveryTime { get; set; }
        
        /// <summary>
        /// Коэффициенты сезонности по месяцам (индекс 0-11 соответствует месяцам с января по декабрь)
        /// </summary>
        public double[] SeasonalityCoefficients { get; set; } = new double[12];
        
        /// <summary>
        /// Дата последнего заказа
        /// </summary>
        public DateTime? LastOrderDate { get; set; }
        
        /// <summary>
        /// Прогнозируемая дата следующего заказа
        /// </summary>
        public DateTime? NextPredictedOrderDate { get; set; }
        
        /// <summary>
        /// Рекомендуемое количество для следующего заказа
        /// </summary>
        public double RecommendedQuantity { get; set; }
        
        /// <summary>
        /// Оптимальная дата размещения следующего заказа
        /// </summary>
        public DateTime? OptimalOrderPlacementDate { get; set; }
    }
}
