using System;

namespace Forecast.Models
{
    /// <summary>
    /// Класс, представляющий элемент заявки из файла Excel
    /// </summary>
    public class OrderItem
    {
        /// <summary>
        /// Дата создания заявки
        /// </summary>
        public DateTime OrderDate { get; set; }
        
        /// <summary>
        /// Номер заявки
        /// </summary>
        public string OrderNumber { get; set; }
        
        /// <summary>
        /// Номер позиции в заявке
        /// </summary>
        public string PositionNumber { get; set; }
        
        /// <summary>
        /// Наименование товара
        /// </summary>
        public string ProductName { get; set; }
        
        /// <summary>
        /// Артикул товара
        /// </summary>
        public string ArticleNumber { get; set; }
        
        /// <summary>
        /// Заказанное количество товара
        /// </summary>
        public double OrderedQuantity { get; set; }
        
        /// <summary>
        /// Поставленное количество товара
        /// </summary>
        public double DeliveredQuantity { get; set; }
        
        /// <summary>
        /// Дата фактической поставки товара
        /// </summary>
        public DateTime? DeliveryDate { get; set; }
        
        /// <summary>
        /// Дополнительная информация о заявке
        /// </summary>
        public string Notes { get; set; }
        
        /// <summary>
        /// Унифицированный артикул для группировки схожих товаров
        /// </summary>
        public string UnifiedArticle { get; set; }
    }
}
