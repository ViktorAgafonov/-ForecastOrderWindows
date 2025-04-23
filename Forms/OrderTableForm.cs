using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Forecast.Models;
using Forecast.Services;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Forecast.Forms
{
    /// <summary>
    /// Форма для отображения таблицы заказов
    /// </summary>
    public partial class OrderTableForm : Form
    {
        private readonly List<ForecastResult> _forecasts;
        private readonly IRecommendationSystem _recommendationSystem;
        private readonly ForecastSettings _forecastSettings;
        
        /// <summary>
        /// Конструктор формы таблицы заказов
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <param name="recommendationSystem">Система рекомендаций</param>
        /// <param name="forecastSettings">Настройки прогнозирования</param>
        public OrderTableForm(List<ForecastResult> forecasts, IRecommendationSystem recommendationSystem, ForecastSettings forecastSettings)
        {
            InitializeComponent();
            
            _forecasts = forecasts;
            _recommendationSystem = recommendationSystem;
            _forecastSettings = forecastSettings;
            
            // Настройка формы
            this.Text = "Таблица заказов";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Size = new Size(1000, 600);
            
            // Инициализация компонентов формы
            InitializeCustomComponents();
            
            // Загрузка данных
            LoadData();
        }
        
        /// <summary>
        /// Инициализация компонентов формы
        /// </summary>
        private void InitializeCustomComponents()
        {
            // Создание панели инструментов
            var toolStrip = new ToolStrip();
            toolStrip.Dock = DockStyle.Top;
            
            var printButton = new ToolStripButton("Печать");
            printButton.Click += PrintButton_Click;
            toolStrip.Items.Add(printButton);
            
            var exportButton = new ToolStripButton("Экспорт");
            exportButton.Click += ExportButton_Click;
            toolStrip.Items.Add(exportButton);
            
            // Добавление разделителя
            toolStrip.Items.Add(new ToolStripSeparator());
            
            // Добавление выпадающего списка для сортировки
            var sortLabel = new ToolStripLabel("Сортировка:");
            toolStrip.Items.Add(sortLabel);
            
            var sortComboBox = new ToolStripComboBox("sortComboBox");
            sortComboBox.Items.AddRange(new object[] {
                "По дате заказа (возрастание)",
                "По дате заказа (убывание)",
                "По дате размещения (возрастание)",
                "По дате размещения (убывание)",
                "По приоритету"
            });
            sortComboBox.SelectedIndex = 1; // По умолчанию - по дате заказа (убывание)
            sortComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            sortComboBox.Width = 200;
            sortComboBox.SelectedIndexChanged += SortComboBox_SelectedIndexChanged;
            toolStrip.Items.Add(sortComboBox);
            
            // Добавление фильтра по приоритету
            var priorityLabel = new ToolStripLabel("Приоритет:");
            toolStrip.Items.Add(priorityLabel);
            
            var priorityComboBox = new ToolStripComboBox("priorityComboBox");
            priorityComboBox.Items.AddRange(new object[] {
                "Все",
                "Наивысший (1)",
                "Высокий (2)",
                "Средний (3)",
                "Низкий (4)",
                "Самый низкий (5)"
            });
            priorityComboBox.SelectedIndex = 0; // По умолчанию - все
            priorityComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            priorityComboBox.Width = 150;
            priorityComboBox.SelectedIndexChanged += PriorityComboBox_SelectedIndexChanged;
            toolStrip.Items.Add(priorityComboBox);
            
            // Создание строки состояния
            var statusStrip = new StatusStrip();
            statusStrip.Dock = DockStyle.Bottom;
            
            var totalItemsLabel = new ToolStripStatusLabel();
            totalItemsLabel.Name = "totalItemsLabel";
            statusStrip.Items.Add(totalItemsLabel);
            
            // Создание панели для таблицы
            var contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            
            // Создание таблицы заказов
            var ordersGridView = new DataGridView();
            ordersGridView.Dock = DockStyle.Fill;
            ordersGridView.AllowUserToAddRows = false;
            ordersGridView.ReadOnly = true;
            ordersGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            ordersGridView.Name = "ordersGridView";
            ordersGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            ordersGridView.RowHeadersVisible = false;
            ordersGridView.AllowUserToResizeRows = false;
            ordersGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            ordersGridView.CellDoubleClick += OrdersGridView_CellDoubleClick;
            
            // Настройка цветов для приоритетов
            ordersGridView.CellFormatting += (sender, e) =>
            {
                if (e.ColumnIndex == 0 && e.RowIndex >= 0) // Столбец "Приоритет"
                {
                    if (e.Value != null)
                    {
                        int priority;
                        if (int.TryParse(e.Value.ToString(), out priority))
                        {
                            switch (priority)
                            {
                                case 1: // Наивысший приоритет
                                    e.CellStyle.BackColor = Color.Red;
                                    e.CellStyle.ForeColor = Color.White;
                                    break;
                                case 2: // Высокий приоритет
                                    e.CellStyle.BackColor = Color.Orange;
                                    break;
                                case 3: // Средний приоритет
                                    e.CellStyle.BackColor = Color.Yellow;
                                    break;
                                case 4: // Низкий приоритет
                                    e.CellStyle.BackColor = Color.LightGreen;
                                    break;
                                case 5: // Самый низкий приоритет
                                    e.CellStyle.BackColor = Color.LightBlue;
                                    break;
                            }
                        }
                    }
                }
            };
            
            contentPanel.Controls.Add(ordersGridView);
            
            // Добавление компонентов на форму в правильном порядке
            this.Controls.Add(contentPanel);
            this.Controls.Add(statusStrip);
            this.Controls.Add(toolStrip);
        }
        
        /// <summary>
        /// Загрузка данных
        /// </summary>
        private void LoadData()
        {
            if (_forecasts == null || _forecasts.Count == 0)
                return;
            
            // Получение таблицы заказов
            var ordersGridView = this.Controls.Find("ordersGridView", true).FirstOrDefault() as DataGridView;
            
            if (ordersGridView != null)
            {
                // По умолчанию сортируем по дате заказа (убывание)
                var sortedForecasts = _forecasts.OrderByDescending(f => f.NextOrderDate).ToList();
                
                // Отображение данных
                DisplayOrders(sortedForecasts);
            }
        }
        
        /// <summary>
        /// Отображение заказов в таблице
        /// </summary>
        /// <param name="forecasts">Список прогнозов для отображения</param>
        private void DisplayOrders(List<ForecastResult> forecasts)
        {
            try
            {
                // Получаем ссылки на элементы управления
                DataGridView? ordersGridView = null;
                ToolStripStatusLabel? totalItemsLabel = null;
                
                // Ищем DataGridView
                foreach (Control control in this.Controls)
                {
                    if (control is Panel panel)
                    {
                        foreach (Control panelControl in panel.Controls)
                        {
                            if (panelControl is DataGridView dataGridView && dataGridView.Name == "ordersGridView")
                            {
                                ordersGridView = dataGridView;
                                break;
                            }
                        }
                    }
                }
                
                // Ищем StatusStrip и ToolStripStatusLabel
                foreach (Control control in this.Controls)
                {
                    if (control is StatusStrip statusStrip)
                    {
                        foreach (ToolStripItem item in statusStrip.Items)
                        {
                            if (item is ToolStripStatusLabel label && label.Name == "totalItemsLabel")
                            {
                                totalItemsLabel = label;
                                break;
                            }
                        }
                    }
                }
                
                // Проверка на null-ссылки
                if (ordersGridView == null || totalItemsLabel == null)
                {
                    MessageBox.Show("Не удалось найти элементы управления для отображения заказов. Возможно, они не были инициализированы.", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // Проверка на наличие данных
                if (forecasts == null || forecasts.Count == 0)
                {
                    ordersGridView.DataSource = null;
                    totalItemsLabel.Text = "Нет данных для отображения";
                    return;
                }
                
                // Фильтрация прогнозов по порогу уверенности
                var filteredForecasts = forecasts;
                if (_forecastSettings != null && _forecastSettings.MinConfidenceThreshold > 0)
                {
                    filteredForecasts = forecasts.Where(f => f.Confidence >= _forecastSettings.MinConfidenceThreshold).ToList();
                    
                    // Если после фильтрации нет данных
                    if (filteredForecasts.Count == 0)
                    {
                        ordersGridView.DataSource = null;
                        totalItemsLabel.Text = $"Нет данных, соответствующих порогу уверенности {_forecastSettings.MinConfidenceThreshold}%";
                        return;
                    }
                }
                
                // Очистка таблицы
                ordersGridView.DataSource = null;
                
                // Создание источника данных для таблицы
                var dataSource = filteredForecasts.Select(order => new
                {
                    Приоритет = order.Priority,
                    Артикул = order.UnifiedArticle,
                    Наименование = order.ProductName,
                    Дата_заказа = order.NextOrderDate.ToShortDateString(),
                    Количество = Math.Round(order.RecommendedQuantity, 0),
                    Дата_размещения = order.OptimalOrderPlacementDate.ToShortDateString(),
                    Уверенность = $"{order.Confidence:F1}%",
                    Примечание = order.Notes
                }).ToList();
                
                // Установка источника данных
                ordersGridView.DataSource = dataSource;
                
                // Настройка заголовков столбцов
                if (ordersGridView.Columns.Count >= 8) // Проверка, что все столбцы существуют
                {
                    // Установка заголовков столбцов
                    ordersGridView.Columns[0].HeaderText = "Приоритет";
                    ordersGridView.Columns[1].HeaderText = "Артикул";
                    ordersGridView.Columns[2].HeaderText = "Наименование товара";
                    ordersGridView.Columns[3].HeaderText = "Дата заказа";
                    ordersGridView.Columns[4].HeaderText = "Количество";
                    ordersGridView.Columns[5].HeaderText = "Дата размещения";
                    ordersGridView.Columns[6].HeaderText = "Уверенность, %";
                    ordersGridView.Columns[7].HeaderText = "Примечание";
                    
                    // Настройка ширины столбцов
                    ordersGridView.Columns[0].Width = 80; // Приоритет
                    ordersGridView.Columns[1].Width = 100; // Артикул
                    ordersGridView.Columns[2].Width = 200; // Наименование
                    ordersGridView.Columns[3].Width = 100; // Дата заказа
                    ordersGridView.Columns[4].Width = 80; // Количество
                    ordersGridView.Columns[5].Width = 120; // Дата размещения
                    ordersGridView.Columns[6].Width = 100; // Уверенность
                    ordersGridView.Columns[7].Width = 300; // Примечание
                    
                    // Выравнивание в столбцах
                    ordersGridView.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Приоритет
                    ordersGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Дата заказа
                    ordersGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // Количество
                    ordersGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Дата размещения
                    ordersGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Уверенность
                }
                
                // Обновление статусной строки
                totalItemsLabel.Text = $"Всего заказов: {forecasts.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отображении таблицы заказов: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Сортировка заказов
        /// </summary>
        /// <param name="sortIndex">Индекс выбранного варианта сортировки</param>
        /// <param name="priorityFilter">Фильтр по приоритету (0 - все, 1-5 - конкретный приоритет)</param>
        private void SortOrders(int sortIndex, int priorityFilter)
        {
            if (_forecasts == null || _forecasts.Count == 0)
                return;
            
            // Применение фильтра по приоритету
            var filteredForecasts = _forecasts;
            
            if (priorityFilter > 0)
            {
                filteredForecasts = _forecasts.Where(f => f.Priority == priorityFilter).ToList();
            }
            
            // Применение сортировки
            IEnumerable<ForecastResult> sortedForecasts;
            
            switch (sortIndex)
            {
                case 0: // По дате заказа (возрастание)
                    sortedForecasts = filteredForecasts.OrderBy(f => f.NextOrderDate);
                    break;
                case 1: // По дате заказа (убывание)
                    sortedForecasts = filteredForecasts.OrderByDescending(f => f.NextOrderDate);
                    break;
                case 2: // По дате размещения (возрастание)
                    sortedForecasts = filteredForecasts.OrderBy(f => f.OptimalOrderPlacementDate);
                    break;
                case 3: // По дате размещения (убывание)
                    sortedForecasts = filteredForecasts.OrderByDescending(f => f.OptimalOrderPlacementDate);
                    break;
                case 4: // По приоритету
                    sortedForecasts = filteredForecasts.OrderBy(f => f.Priority);
                    break;
                default:
                    sortedForecasts = filteredForecasts.OrderByDescending(f => f.NextOrderDate);
                    break;
            }
            
            // Отображение отсортированных данных
            DisplayOrders(sortedForecasts.ToList());
        }
        
        #region Обработчики событий
        
        /// <summary>
        /// Обработчик события изменения выбранного варианта сортировки
        /// </summary>
        private void SortComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var sortComboBox = sender as ToolStripComboBox;
            
            if (sortComboBox == null) return;
            
            var toolStrip = sortComboBox.GetCurrentParent();
            var priorityComboBox = toolStrip?.Items.OfType<ToolStripComboBox>().FirstOrDefault(x => x.Name == "priorityComboBox");
            
            if (priorityComboBox != null)
            {
                int sortIndex = sortComboBox.SelectedIndex;
                int priorityFilter = priorityComboBox.SelectedIndex;
                
                SortOrders(sortIndex, priorityFilter);
            }
        }
        
        /// <summary>
        /// Обработчик события изменения выбранного фильтра по приоритету
        /// </summary>
        private void PriorityComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var priorityComboBox = sender as ToolStripComboBox;
            
            if (priorityComboBox == null) return;
            
            var toolStrip = priorityComboBox.GetCurrentParent();
            var sortComboBox = toolStrip?.Items.OfType<ToolStripComboBox>().FirstOrDefault(x => x.Name == "sortComboBox");
            
            if (sortComboBox != null)
            {
                int priorityFilter = priorityComboBox.SelectedIndex;
                int sortIndex = sortComboBox.SelectedIndex;
                
                SortOrders(sortIndex, priorityFilter);
            }
        }
        
        /// <summary>
        /// Обработчик события двойного клика по ячейке таблицы
        /// </summary>
        private void OrdersGridView_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var ordersGridView = sender as DataGridView;
                
                if (ordersGridView != null)
                {
                    try
                    {
                        // Получение артикула выбранного заказа
                        string? article = ordersGridView.Rows[e.RowIndex].Cells["Артикул"].Value?.ToString();
                        
                        if (string.IsNullOrEmpty(article)) return;
                        
                        // Поиск соответствующего прогноза
                        var forecast = _forecasts.FirstOrDefault(f => f.UnifiedArticle == article);
                        
                        if (forecast != null)
                        {
                            // Отображение детальной информации о заказе
                            MessageBox.Show(
                                $"Артикул: {forecast.UnifiedArticle}\n" +
                                $"Наименование: {forecast.ProductName}\n" +
                                $"Дата заказа: {forecast.NextOrderDate.ToShortDateString()}\n" +
                                $"Количество: {Math.Round(forecast.RecommendedQuantity, 0)}\n" +
                                $"Дата размещения: {forecast.OptimalOrderPlacementDate.ToShortDateString()}\n" +
                                $"Приоритет: {forecast.Priority}\n" +
                                $"Уверенность: {forecast.Confidence:P0}\n" +
                                $"Примечание: {forecast.Notes}",
                                "Детальная информация о заказе",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при получении детальной информации: {ex.Message}",
                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        
        /// <summary>
        /// Обработчик события нажатия кнопки "Печать"
        /// </summary>
        private void PrintButton_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("Функция печати таблицы заказов будет реализована в следующей версии.", 
                "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// Обработчик события нажатия кнопки "Экспорт"
        /// </summary>
        private void ExportButton_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_forecasts == null || _forecasts.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта.", "Предупреждение", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Создание диалога сохранения файла
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel файлы (*.xlsx)|*.xlsx",
                    Title = "Экспорт таблицы заказов",
                    FileName = $"Таблица_заказов_{DateTime.Now:yyyy-MM-dd}"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Экспорт данных в Excel
                    ExportToExcel(saveFileDialog.FileName);
                    
                    // Сообщение об успешном экспорте
                    MessageBox.Show($"Данные успешно экспортированы в файл:\n{saveFileDialog.FileName}", 
                        "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    // Открытие файла после экспорта
                    if (MessageBox.Show("Открыть созданный файл?", "Вопрос", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", saveFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Экспорт данных в Excel-файл
        /// </summary>
        /// <param name="filePath">Путь к файлу для сохранения</param>
        private void ExportToExcel(string filePath)
        {
            // Установка лицензии EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Фильтрация прогнозов по порогу уверенности
            var filteredForecasts = _forecasts;
            if (_forecastSettings != null && _forecastSettings.MinConfidenceThreshold > 0)
            {
                filteredForecasts = _forecasts.Where(f => f.Confidence >= _forecastSettings.MinConfidenceThreshold).ToList();
            }
            
            using (var package = new ExcelPackage())
            {
                // Создание листа с таблицей заказов
                var worksheet = package.Workbook.Worksheets.Add("Таблица заказов");
                
                // Заголовки столбцов
                worksheet.Cells[1, 1].Value = "Приоритет";
                worksheet.Cells[1, 2].Value = "Артикул";
                worksheet.Cells[1, 3].Value = "Наименование";
                worksheet.Cells[1, 4].Value = "Дата заказа";
                worksheet.Cells[1, 5].Value = "Количество";
                worksheet.Cells[1, 6].Value = "Дата размещения";
                worksheet.Cells[1, 7].Value = "Уверенность";
                worksheet.Cells[1, 8].Value = "Примечание";
                
                // Форматирование заголовков
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                
                // Заполнение данными
                int row = 2;
                foreach (var forecast in filteredForecasts.OrderBy(f => f.NextOrderDate))
                {
                    worksheet.Cells[row, 1].Value = forecast.Priority;
                    worksheet.Cells[row, 2].Value = forecast.UnifiedArticle;
                    worksheet.Cells[row, 3].Value = forecast.ProductName;
                    worksheet.Cells[row, 4].Value = forecast.NextOrderDate;
                    worksheet.Cells[row, 5].Value = Math.Round(forecast.RecommendedQuantity, 0);
                    worksheet.Cells[row, 6].Value = forecast.OptimalOrderPlacementDate;
                    worksheet.Cells[row, 7].Value = forecast.Confidence / 100; // Для отображения в процентном формате
                    worksheet.Cells[row, 8].Value = forecast.Notes;
                    
                    // Форматирование ячеек с приоритетом
                    var priorityCell = worksheet.Cells[row, 1];
                    priorityCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    priorityCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    priorityCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    priorityCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    priorityCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    
                    // Цвет фона в зависимости от приоритета
                    switch (forecast.Priority)
                    {
                        case 1: // Наивысший приоритет
                            priorityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            priorityCell.Style.Fill.BackgroundColor.SetColor(Color.Red);
                            priorityCell.Style.Font.Color.SetColor(Color.White);
                            break;
                        case 2: // Высокий приоритет
                            priorityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            priorityCell.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                            break;
                        case 3: // Средний приоритет
                            priorityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            priorityCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                            break;
                        case 4: // Низкий приоритет
                            priorityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            priorityCell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                            break;
                        case 5: // Самый низкий приоритет
                            priorityCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            priorityCell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                            break;
                    }
                    
                    // Форматирование остальных ячеек в строке
                    using (var rowRange = worksheet.Cells[row, 2, row, 8])
                    {
                        rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }
                    
                    // Форматирование дат
                    worksheet.Cells[row, 4].Style.Numberformat.Format = "dd.MM.yyyy";
                    worksheet.Cells[row, 6].Style.Numberformat.Format = "dd.MM.yyyy";
                    
                    // Форматирование процентов
                    worksheet.Cells[row, 7].Style.Numberformat.Format = "0%";
                    
                    row++;
                }
                
                // Автоподбор ширины столбцов
                worksheet.Cells.AutoFitColumns();
                
                // Сохранение файла
                package.SaveAs(new FileInfo(filePath));
            }
        }
        
        #endregion
    }
}
