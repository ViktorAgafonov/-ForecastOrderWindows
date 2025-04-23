using System.Collections.Generic;
using Forecast.Models;

namespace Forecast.Services
{
    /// <summary>
    /// Интерфейс для обработчика данных из файла Excel
    /// </summary>
    public interface IDataProcessor
    {
        /// <summary>
        /// Загрузка данных из файла Excel
        /// </summary>
        /// <param name="filePath">Путь к файлу Excel</param>
        /// <returns>Список элементов заявок</returns>
        List<OrderItem> LoadData(string filePath);
        
        /// <summary>
        /// Унификация наименований и артикулов товаров
        /// </summary>
        /// <param name="orderItems">Список элементов заявок</param>
        /// <returns>Список унифицированных товаров</returns>
        List<UnifiedProduct> UnifyProducts(List<OrderItem> orderItems);
        
        /// <summary>
        /// Сохранение базы соответствий наименований и артикулов
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        void SaveItemMapping(List<UnifiedProduct> unifiedProducts, string filePath);
        
        /// <summary>
        /// Загрузка базы соответствий наименований и артикулов
        /// </summary>
        /// <param name="filePath">Путь к файлу с базой соответствий</param>
        /// <returns>Список унифицированных товаров</returns>
        List<UnifiedProduct> LoadItemMapping(string filePath);
    }
}
