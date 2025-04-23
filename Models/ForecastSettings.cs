using System;
using System.IO;
using System.Text.Json;

namespace Forecast.Models
{
    /// <summary>
    /// Класс для хранения настроек прогнозирования
    /// </summary>
    public class ForecastSettings
    {
        /// <summary>
        /// Количество дней для прогноза вперед
        /// </summary>
        public int DaysAhead { get; set; } = 30;
        
        /// <summary>
        /// Минимальное количество исторических заказов для высокой уверенности
        /// </summary>
        public int MinOrderHistoryForHighConfidence { get; set; } = 5;
        
        /// <summary>
        /// Минимальное количество исторических заказов для средней уверенности
        /// </summary>
        public int MinOrderHistoryForMediumConfidence { get; set; } = 3;
        
        /// <summary>
        /// Коэффициент запаса для расчета оптимальной даты размещения заказа
        /// </summary>
        public double SafetyFactorForOrderPlacement { get; set; } = 0.2;
        
        /// <summary>
        /// Порог для определения стабильности объемов заказов (в процентах)
        /// </summary>
        public double StableVolumeThreshold { get; set; } = 10.0;
        
        /// <summary>
        /// Порог для определения высокого приоритета (в днях)
        /// </summary>
        public int HighPriorityThreshold { get; set; } = 7;
        
        /// <summary>
        /// Порог для определения среднего приоритета (в днях)
        /// </summary>
        public int MediumPriorityThreshold { get; set; } = 14;
        
        /// <summary>
        /// Порог для определения низкого приоритета (в днях)
        /// </summary>
        public int LowPriorityThreshold { get; set; } = 30;
        
        /// <summary>
        /// Коэффициент сезонности по умолчанию
        /// </summary>
        public double DefaultSeasonalityCoefficient { get; set; } = 1.0;
        
        /// <summary>
        /// Загрузка настроек из файла
        /// </summary>
        /// <param name="filePath">Путь к файлу настроек</param>
        /// <returns>Объект настроек</returns>
        public static ForecastSettings LoadFromFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    return JsonSerializer.Deserialize<ForecastSettings>(json) ?? new ForecastSettings();
                }
            }
            catch (Exception)
            {
                // В случае ошибки возвращаем настройки по умолчанию
            }
            
            return new ForecastSettings();
        }
        
        /// <summary>
        /// Сохранение настроек в файл
        /// </summary>
        /// <param name="filePath">Путь к файлу настроек</param>
        public void SaveToFile(string filePath)
        {
            try
            {
                string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json);
            }
            catch (Exception)
            {
                // Обработка ошибок сохранения
            }
        }
        
        /// <summary>
        /// Описание влияния параметра на прогноз
        /// </summary>
        /// <param name="parameterName">Имя параметра</param>
        /// <returns>Описание влияния</returns>
        public static string GetParameterDescription(string parameterName)
        {
            return parameterName switch
            {
                nameof(DaysAhead) => "Определяет период в днях, на который будет составлен прогноз. Увеличение этого значения позволит планировать заказы на более длительный срок, но может снизить точность прогноза для отдаленных дат.",
                
                nameof(MinOrderHistoryForHighConfidence) => "Минимальное количество исторических заказов для высокой уверенности в прогнозе (95%). Чем больше исторических данных, тем точнее прогноз.",
                
                nameof(MinOrderHistoryForMediumConfidence) => "Минимальное количество исторических заказов для средней уверенности в прогнозе (85%). При меньшем количестве заказов уверенность будет низкой (70%).",
                
                nameof(SafetyFactorForOrderPlacement) => "Коэффициент запаса времени для расчета оптимальной даты размещения заказа. Увеличение этого значения даст больший запас времени, но может привести к преждевременному заказу товаров.",
                
                nameof(StableVolumeThreshold) => "Порог изменения объемов заказов (в процентах), при котором объемы считаются стабильными. Если изменение меньше этого значения, система не будет сообщать о тренде.",
                
                nameof(HighPriorityThreshold) => "Количество дней до оптимальной даты размещения заказа, при котором заказу присваивается высокий приоритет (1-2). Чем меньше дней осталось, тем выше приоритет.",
                
                nameof(MediumPriorityThreshold) => "Количество дней до оптимальной даты размещения заказа, при котором заказу присваивается средний приоритет (3). Заказы с более отдаленной датой получат низкий приоритет.",
                
                nameof(LowPriorityThreshold) => "Количество дней до оптимальной даты размещения заказа, при котором заказу присваивается низкий приоритет (4). Заказы с более отдаленной датой получат самый низкий приоритет (5).",
                
                nameof(DefaultSeasonalityCoefficient) => "Коэффициент сезонности по умолчанию для месяцев без выраженной сезонности. Значение 1.0 означает отсутствие сезонного влияния.",
                
                _ => "Описание недоступно"
            };
        }
    }
}
