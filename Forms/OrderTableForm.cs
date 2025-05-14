using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Forecast.Models;
using Forecast.Services;
using ClosedXML.Excel;

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
        private DataGridView? _ordersGridView;
        
        /// <summary>
        /// Конструктор формы таблицы заказов
        /// </summary>
        /// <param name="forecasts">Список прогнозов</param>
        /// <param name="recommendationSystem">Система рекомендаций</param>
        /// <param name="forecastSettings">Настройки прогнозирования</param>
        public OrderTableForm(List<ForecastResult> forecasts, IRecommendationSystem recommendationSystem, ForecastSettings forecastSettings)
        {
            if (forecasts == null)
                throw new ArgumentNullException(nameof(forecasts));
            
            InitializeComponent();
            _forecasts = forecasts ?? throw new ArgumentNullException(nameof(forecasts));
            _recommendationSystem = recommendationSystem ?? throw new ArgumentNullException(nameof(recommendationSystem));
            _forecastSettings = forecastSettings ?? throw new ArgumentNullException(nameof(forecastSettings));
            InitializeCustomComponents();
            LoadData();
        }

        private void InitializeCustomComponents()
        {

            
            try
            {
                // Создание панели инструментов

                var toolStrip = new ToolStrip();
                if (toolStrip == null)
                {

                    return;
                }
                toolStrip.Dock = DockStyle.Top;
                
                // Создание кнопки печати

                var printButton = new ToolStripButton("Печать");
                if (printButton == null)
                {

                    return;
                }
                printButton.Click += PrintButton_Click;
                toolStrip.Items.Add(printButton);
                
                // Создание кнопки экспорта

                var exportButton = new ToolStripButton("Экспорт");
                if (exportButton == null)
                {

                    return;
                }
                exportButton.Click += ExportButton_Click;
                toolStrip.Items.Add(exportButton);
                
                // Добавление разделителя

                toolStrip.Items.Add(new ToolStripSeparator());
                
                // Добавление выпадающего списка для сортировки

                var sortLabel = new ToolStripLabel("Сортировка:");
                toolStrip.Items.Add(sortLabel);
                
                var sortComboBox = new ToolStripComboBox("sortComboBox");
                if (sortComboBox == null)
                {

                    return;
                }
                sortComboBox.Items.AddRange(new object[] {
                    "Дата заказа (возрастание)",
                    "Дата заказа (убывание)",
                    "Дата размещения (возрастание)",
                    "Дата размещения (убывание)",
                    "По приоритету"
                });
                sortComboBox.SelectedIndex = 1; // По умолчанию - по дате заказа (убывание)
                sortComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                sortComboBox.AutoSize = false;
                sortComboBox.Width = 200;
                sortComboBox.SelectedIndexChanged += SortComboBox_SelectedIndexChanged;
                toolStrip.Items.Add(sortComboBox);
                
                // Добавление фильтра по приоритету

                var priorityLabel = new ToolStripLabel("Приоритет:");
                toolStrip.Items.Add(priorityLabel);
                
                var priorityComboBox = new ToolStripComboBox("priorityComboBox");
                if (priorityComboBox == null)
                {

                    return;
                }
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
                sortComboBox.AutoSize = false;
                priorityComboBox.Width = 150;
                priorityComboBox.SelectedIndexChanged += PriorityComboBox_SelectedIndexChanged;
                toolStrip.Items.Add(priorityComboBox);
                
                // Создание строки состояния

                var statusStrip = new StatusStrip();
                if (statusStrip == null)
                {

                    return;
                }
                statusStrip.Dock = DockStyle.Bottom;
                
                var totalItemsLabel = new ToolStripStatusLabel();
                if (totalItemsLabel == null)
                {

                    return;
                }
                totalItemsLabel.Name = "totalItemsLabel";
                statusStrip.Items.Add(totalItemsLabel);
                
                // Создание панели для таблицы

                var contentPanel = new Panel();
                if (contentPanel == null)
                {

                    return;
                }
                contentPanel.Dock = DockStyle.Fill;
                
                // Создание таблицы заказов

                _ordersGridView = new DataGridView();
                if (_ordersGridView == null)
                {

                    return;
                }
                _ordersGridView.Dock = DockStyle.Fill;
                _ordersGridView.AllowUserToAddRows = false;
                _ordersGridView.ReadOnly = true;
                _ordersGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; // Изменено с Fill на None
                _ordersGridView.Name = "ordersGridView";
                _ordersGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                _ordersGridView.RowHeadersVisible = false;
                _ordersGridView.AllowUserToResizeRows = false;
                _ordersGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
                _ordersGridView.CellDoubleClick += OrdersGridView_CellDoubleClick;
                
                // Настройка цветов для приоритетов
                _ordersGridView.CellFormatting += (sender, e) =>
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
                
                contentPanel.Controls.Add(_ordersGridView);
                
                // Добавление компонентов на форму в правильном порядке

                this.Controls.Add(contentPanel);
                this.Controls.Add(statusStrip);
                this.Controls.Add(toolStrip);
                

            }
            catch (Exception ex)
            {
                // Логирование ошибки
                MessageBox.Show($"Ошибка при инициализации компонентов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void LoadData()
        {
            if (_forecasts == null || _forecasts.Count == 0)
            {

                return;
            }
            

            
            // Получение таблицы заказов
            DataGridView? ordersGridView = null;
            
            // Сначала ищем по имени
            ordersGridView = this.Controls.OfType<DataGridView>().FirstOrDefault(dg => dg.Name == "ordersGridView");
            
            if (ordersGridView == null)
            {
                // Если не найдено по имени, ищем в дочерних контролях
                foreach (var control in this.Controls)
                {
                    if (control is Panel panel)
                    {
                        ordersGridView = panel.Controls.OfType<DataGridView>().FirstOrDefault(dg => dg.Name == "ordersGridView");
                        if (ordersGridView != null)
                            break;
                    }
                }
            }
            
            if (ordersGridView == null)
            {

                return;
            }
            

            
            // По умолчанию сортируем по дате заказа (убывание)
            try
            {
                var sortedForecasts = _forecasts.OrderByDescending(f => f.NextOrderDate).ToList();
                
                if (sortedForecasts == null)
                {

                    return;
                }
                

                
                // Отображение данных
                DisplayOrders(sortedForecasts);
            }
            catch (Exception ex)
            {
                // Логирование ошибки
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
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
                
                // Применяем настройки таблицы после установки источника данных
                ordersGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                
                // Настройка названий, высоты и ширины заголовков столбцов
                if (ordersGridView.Columns.Count >= 8) // Проверка, что все столбцы существуют
                {
                    // Задать высоту заголовка как две строки
                    ordersGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                    ordersGridView.ColumnHeadersHeight = 40;
                    // Установка выравнивания заголовков столбцов по центру
                    ordersGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    // Возможность переноса
                    ordersGridView.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

                    // Установка названий заголовков столбцов
                    ordersGridView.Columns[0].HeaderText = "Приоритет";
                    ordersGridView.Columns[1].HeaderText = "Артикул";
                    ordersGridView.Columns[2].HeaderText = "Наименование товара";
                    ordersGridView.Columns[3].HeaderText = "Дата\nзаказа";
                    ordersGridView.Columns[4].HeaderText = "Количество";
                    ordersGridView.Columns[5].HeaderText = "Дата\nразмещения";
                    ordersGridView.Columns[6].HeaderText = "Уверенность,\n%";
                    ordersGridView.Columns[7].HeaderText = "Примечание";

                    // Настройка ширины столбцов - принудительно отключаем автоматическое изменение размера
                    
                    ordersGridView.Columns[0].Width = 70; // Приоритет 80
                    ordersGridView.Columns[1].Width = 100; // Артикул 100
                    ordersGridView.Columns[2].Width = 400; // Наименование 200
                    ordersGridView.Columns[3].Width = 80; // Дата заказа 100
                    ordersGridView.Columns[4].Width = 80; // Количество 80
                    ordersGridView.Columns[5].Width = 80; // Дата размещения 120
                    ordersGridView.Columns[6].Width = 100; // Уверенность 100
                    ordersGridView.Columns[7].Width = 400; // Примечание 300

                    // Выравнивание данных в столбцах
                    ordersGridView.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Приоритет
                    ordersGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter; // Артикул
                    ordersGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Дата заказа
                    ordersGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;  // Количество
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
            // Фильтрация прогнозов по порогу уверенности
            var filteredForecasts = _forecasts;
            if (_forecastSettings != null && _forecastSettings.MinConfidenceThreshold > 0)
            {
                filteredForecasts = _forecasts.Where(f => f.Confidence >= _forecastSettings.MinConfidenceThreshold).ToList();
            }
            
            using (var workbook = new XLWorkbook())
            {
                // Создание листа с таблицей заказов
                var worksheet = workbook.Worksheets.Add("Таблица заказов");
                
                // Заголовки столбцов
                worksheet.Cell(1, 1).Value = "Приоритет";
                worksheet.Cell(1, 2).Value = "Артикул";
                worksheet.Cell(1, 3).Value = "Наименование";
                worksheet.Cell(1, 4).Value = "Дата заказа";
                worksheet.Cell(1, 5).Value = "Количество";
                worksheet.Cell(1, 6).Value = "Дата размещения";
                worksheet.Cell(1, 7).Value = "Уверенность";
                worksheet.Cell(1, 8).Value = "Примечание";
                
                // Форматирование заголовков
                var headerRange = worksheet.Range(1, 1, 1, 8);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                
                // Заполнение данными
                int row = 2;
                foreach (var forecast in filteredForecasts.OrderBy(f => f.NextOrderDate))
                {
                    worksheet.Cell(row, 1).Value = forecast.Priority;
                    worksheet.Cell(row, 2).Value = forecast.UnifiedArticle;
                    worksheet.Cell(row, 3).Value = forecast.ProductName;
                    worksheet.Cell(row, 4).Value = forecast.NextOrderDate;
                    worksheet.Cell(row, 5).Value = Math.Round(forecast.RecommendedQuantity, 0);
                    worksheet.Cell(row, 6).Value = forecast.OptimalOrderPlacementDate;
                    worksheet.Cell(row, 7).Value = forecast.Confidence / 100.0; // Для отображения в процентном формате
                    worksheet.Cell(row, 8).Value = forecast.Notes;
                    
                    // Форматирование ячеек с приоритетом
                    var priorityCell = worksheet.Cell(row, 1);
                    priorityCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    priorityCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    
                    // Цвет фона в зависимости от приоритета
                    switch (forecast.Priority)
                    {
                        case 1: // Наивысший приоритет
                            priorityCell.Style.Fill.BackgroundColor = XLColor.Red;
                            priorityCell.Style.Font.FontColor = XLColor.White;
                            break;
                        case 2: // Высокий приоритет
                            priorityCell.Style.Fill.BackgroundColor = XLColor.Orange;
                            break;
                        case 3: // Средний приоритет
                            priorityCell.Style.Fill.BackgroundColor = XLColor.Yellow;
                            break;
                        case 4: // Низкий приоритет
                            priorityCell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                            break;
                        case 5: // Самый низкий приоритет
                            priorityCell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                            break;
                    }
                    
                    // Форматирование остальных ячеек в строке
                    var rowRange = worksheet.Range(row, 2, row, 8);
                    rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    
                    // Форматирование дат
                    worksheet.Cell(row, 4).Style.DateFormat.Format = "dd.MM.yyyy";
                    worksheet.Cell(row, 6).Style.DateFormat.Format = "dd.MM.yyyy";
                    
                    // Форматирование процентов
                    worksheet.Cell(row, 7).Style.NumberFormat.Format = "0%";
                    
                    row++;
                }
                
                // Автоподбор ширины столбцов
                worksheet.Columns().AdjustToContents();
                
                // Сохранение файла
                workbook.SaveAs(filePath);
            }
        }
        
        #endregion
    }
}
