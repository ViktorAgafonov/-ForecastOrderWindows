using System;
using System.Collections.Generic;

namespace Forecast.Models
{
    /// <summary>
    /// Класс, представляющий группу соответствий для унификации товаров
    /// </summary>
    public class MappingGroup
    {
        /// <summary>
        /// Уникальный идентификатор группы
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();
        
        /// <summary>
        /// Название группы соответствий
        /// </summary>
        public string Name { get; set; } = "";
        
        /// <summary>
        /// Унифицированный артикул товара
        /// </summary>
        public string UnifiedArticle { get; set; } = "";
        
        /// <summary>
        /// Основное наименование товара
        /// </summary>
        public string PrimaryName { get; set; } = "";
        
        /// <summary>
        /// Список всех вариаций наименований товара
        /// </summary>
        public List<string> NameVariations { get; set; } = new List<string>();
        
        /// <summary>
        /// Список всех вариаций артикулов товара
        /// </summary>
        public List<string> ArticleVariations { get; set; } = new List<string>();
    }
    
    /// <summary>
    /// Класс для хранения списка групп соответствий
    /// </summary>
    public class MappingDatabase
    {
        /// <summary>
        /// Список групп соответствий
        /// </summary>
        public List<MappingGroup> Groups { get; set; } = new List<MappingGroup>();
    }
}
