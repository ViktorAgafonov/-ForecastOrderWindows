using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Forecast.Models;
using Forecast.Services;

namespace Forecast.Forms
{
    /// <summary>
    /// Главная форма приложения
    /// </summary>
    public partial class MainForm : Form
    {
        // Сервисы для работы с данными и прогнозами
        private readonly IDataProcessor _dataProcessor;
        private readonly IOrderAnalyzer _orderAnalyzer;
        private readonly IForecastEngine _forecastEngine;
        private readonly IRecommendationSystem _recommendationSystem;
        
        // Данные
        private List<OrderItem> _orderItems;
        private List<UnifiedProduct> _unifiedProducts;
        private List<ForecastResult> _forecasts;
        
        // Пути к файлам
        private string _excelFilePath = string.Empty;
        private string _mappingFilePath = string.Empty;
        private string _forecastsFilePath = string.Empty;
        private string _settingsFilePath = string.Empty;
        
        // Настройки прогноза
        private ForecastSettings _forecastSettings = new ForecastSettings();
        
        /// <summary>
        /// Конструктор главной формы
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
            
            // Инициализация сервисов
            _dataProcessor = new ExcelDataProcessor();
            _orderAnalyzer = new OrderAnalyzer();
            _forecastEngine = new ForecastEngine();
            _recommendationSystem = new RecommendationSystem();
            
            // Инициализация данных
            _orderItems = new List<OrderItem>();
            _unifiedProducts = new List<UnifiedProduct>();
            _forecasts = new List<ForecastResult>();
            
            // Установка путей к файлам
            string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Forecast");
            Directory.CreateDirectory(appDataPath); // Создаем директорию, если она не существует
            
            _mappingFilePath = Path.Combine(appDataPath, "item_mapping.json");
            _forecastsFilePath = Path.Combine(appDataPath, "forecasts.json");
            _settingsFilePath = Path.Combine(appDataPath, "forecast_settings.json");
            
            // Настройка формы
            this.Text = "Прогнозирование заявок";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(1000, 700);
            
            // Инициализация компонентов формы
            InitializeCustomComponents();
            
            // Загрузка сохраненных данных при запуске
            LoadSavedData();
        }
        
        /// <summary>
        /// Инициализация компонентов формы
        /// </summary>
        private void InitializeCustomComponents()
        {
            // Создание меню
            var mainMenu = new MenuStrip();
            this.MainMenuStrip = mainMenu;
            this.Controls.Add(mainMenu);
            
            // Пункт меню "Файл"
            var fileMenuItem = new ToolStripMenuItem("Файл");
            mainMenu.Items.Add(fileMenuItem);
            
            var openFileMenuItem = new ToolStripMenuItem("Открыть Excel файл...");
            openFileMenuItem.Click += OpenFileMenuItem_Click;
            openFileMenuItem.ShortcutKeys = Keys.Control | Keys.O;
            fileMenuItem.DropDownItems.Add(openFileMenuItem);
            
            fileMenuItem.DropDownItems.Add(new ToolStripSeparator());
            
            var exitMenuItem = new ToolStripMenuItem("Выход");
            exitMenuItem.Click += (s, e) => this.Close();
            exitMenuItem.ShortcutKeys = Keys.Alt | Keys.F4;
            fileMenuItem.DropDownItems.Add(exitMenuItem);
            
            // Пункт меню "Анализ"
            var analysisMenuItem = new ToolStripMenuItem("Анализ");
            mainMenu.Items.Add(analysisMenuItem);
            
            var processDataMenuItem = new ToolStripMenuItem("Обработать данные");
            processDataMenuItem.Click += ProcessDataMenuItem_Click;
            processDataMenuItem.ShortcutKeys = Keys.Control | Keys.P;
            analysisMenuItem.DropDownItems.Add(processDataMenuItem);
            
            var analyzeDataMenuItem = new ToolStripMenuItem("Анализировать данные");
            analyzeDataMenuItem.Click += AnalyzeDataMenuItem_Click;
            analyzeDataMenuItem.ShortcutKeys = Keys.Control | Keys.A;
            analysisMenuItem.DropDownItems.Add(analyzeDataMenuItem);
            
            var generateForecastsMenuItem = new ToolStripMenuItem("Сформировать прогнозы");
            generateForecastsMenuItem.Click += GenerateForecastsMenuItem_Click;
            generateForecastsMenuItem.ShortcutKeys = Keys.Control | Keys.F;
            analysisMenuItem.DropDownItems.Add(generateForecastsMenuItem);
            
            // Пункт меню "Настройки"
            var settingsMenuItem = new ToolStripMenuItem("Настройки");
            mainMenu.Items.Add(settingsMenuItem);
            
            var mappingMenuItem = new ToolStripMenuItem("Настройки соответствий");
            mappingMenuItem.Click += MappingMenuItem_Click;
            settingsMenuItem.DropDownItems.Add(mappingMenuItem);
            
            var forecastSettingsMenuItem = new ToolStripMenuItem("Настройки прогноза");
            forecastSettingsMenuItem.Click += ForecastSettingsMenuItem_Click;
            settingsMenuItem.DropDownItems.Add(forecastSettingsMenuItem);
            
            // Пункт меню "Отчеты"
            var reportsMenuItem = new ToolStripMenuItem("Отчеты");
            mainMenu.Items.Add(reportsMenuItem);
            
            var orderTableMenuItem = new ToolStripMenuItem("Таблица заказов");
            orderTableMenuItem.Click += OrderTableMenuItem_Click;
            orderTableMenuItem.ShortcutKeys = Keys.Control | Keys.T;
            reportsMenuItem.DropDownItems.Add(orderTableMenuItem);
            
            var productReportMenuItem = new ToolStripMenuItem("Отчет по товарам");
            productReportMenuItem.Click += ProductReportMenuItem_Click;
            productReportMenuItem.ShortcutKeys = Keys.Control | Keys.R;
            reportsMenuItem.DropDownItems.Add(productReportMenuItem);
            
            // Пункт меню "Справка"
            var helpMenuItem = new ToolStripMenuItem("Справка");
            mainMenu.Items.Add(helpMenuItem);
            
            var aboutMenuItem = new ToolStripMenuItem("О программе");
            aboutMenuItem.Click += AboutMenuItem_Click;
            helpMenuItem.DropDownItems.Add(aboutMenuItem);
            
            // Создание панели статуса
            var statusStrip = new StatusStrip();
            this.Controls.Add(statusStrip);
            
            var statusLabel = new ToolStripStatusLabel("Готово");
            statusStrip.Items.Add(statusLabel);
            
            // Создание вкладок
            var tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            this.Controls.Add(tabControl);
            
            // Вкладка "Данные"
            var dataTabPage = new TabPage("Данные");
            tabControl.TabPages.Add(dataTabPage);
            
            var dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.AllowUserToAddRows = false;
            dataGridView.ReadOnly = true;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView.Name = "dataGridView";
            dataTabPage.Controls.Add(dataGridView);
            
            // Вкладка "Унифицированные товары"
            var productsTabPage = new TabPage("Унифицированные товары");
            tabControl.TabPages.Add(productsTabPage);
            
            var productsGridView = new DataGridView();
            productsGridView.Dock = DockStyle.Fill;
            productsGridView.AllowUserToAddRows = false;
            productsGridView.ReadOnly = true;
            productsGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            productsGridView.Name = "productsGridView";
            productsTabPage.Controls.Add(productsGridView);
            
            // Вкладка "Прогнозы"
            var forecastsTabPage = new TabPage("Прогнозы");
            tabControl.TabPages.Add(forecastsTabPage);
            
            var forecastsGridView = new DataGridView();
            forecastsGridView.Dock = DockStyle.Fill;
            forecastsGridView.AllowUserToAddRows = false;
            forecastsGridView.ReadOnly = true;
            forecastsGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            forecastsGridView.Name = "forecastsGridView";
            forecastsTabPage.Controls.Add(forecastsGridView);
        }
        
        /// <summary>
        /// Загрузка сохраненных данных при запуске
        /// </summary>
        private void LoadSavedData()
        {
            try
            {
                // Загрузка настроек прогноза
                if (File.Exists(_settingsFilePath))
                {
                    _forecastSettings = ForecastSettings.LoadFromFile(_settingsFilePath);
                }
                else
                {
                    // Если файл настроек не существует, создаем его с настройками по умолчанию
                    _forecastSettings = new ForecastSettings();
                    _forecastSettings.SaveToFile(_settingsFilePath);
                }
                
                // Загрузка базы соответствий товаров
                if (File.Exists(_mappingFilePath))
                {
                    _unifiedProducts = _dataProcessor.LoadItemMapping(_mappingFilePath);
                    UpdateProductsGrid();
                }
                
                // Загрузка сохраненных прогнозов
                if (File.Exists(_forecastsFilePath))
                {
                    _forecasts = _recommendationSystem.LoadRecommendations(_forecastsFilePath);
                    UpdateForecastsGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке сохраненных данных: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        #region Обработчики событий меню
        
        /// <summary>
        /// Обработчик события открытия файла Excel
        /// </summary>
        private void OpenFileMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                Title = "Выберите файл с данными о заявках"
            };
            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                _excelFilePath = openFileDialog.FileName;
                LoadExcelData();
            }
        }
        
        /// <summary>
        /// Обработчик события обработки данных
        /// </summary>
        private void ProcessDataMenuItem_Click(object sender, EventArgs e)
        {
            if (_orderItems == null || _orderItems.Count == 0)
            {
                MessageBox.Show("Сначала необходимо загрузить данные из Excel файла.", "Предупреждение", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                Cursor = Cursors.WaitCursor;
                
                // Унификация наименований и артикулов
                _unifiedProducts = _dataProcessor.UnifyProducts(_orderItems);
                
                // Сохранение базы соответствий
                _dataProcessor.SaveItemMapping(_unifiedProducts, _mappingFilePath);
                
                // Обновление таблицы унифицированных товаров
                UpdateProductsGrid();
                
                MessageBox.Show($"Обработка данных завершена. Унифицировано {_unifiedProducts.Count} товаров.", 
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обработке данных: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        
        /// <summary>
        /// Обработчик события анализа данных
        /// </summary>
        private void AnalyzeDataMenuItem_Click(object sender, EventArgs e)
        {
            if (_unifiedProducts == null || _unifiedProducts.Count == 0)
            {
                MessageBox.Show("Сначала необходимо обработать данные.", "Предупреждение", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                Cursor = Cursors.WaitCursor;
                
                // Анализ частоты заказов
                _orderAnalyzer.AnalyzeOrderFrequency(_unifiedProducts);
                
                // Анализ сезонности
                _orderAnalyzer.AnalyzeSeasonality(_unifiedProducts);
                
                // Анализ объемов заказов
                _orderAnalyzer.AnalyzeOrderVolumes(_unifiedProducts);
                
                // Анализ сроков поставки
                _orderAnalyzer.AnalyzeDeliveryTimes(_unifiedProducts);
                
                // Обновление таблицы унифицированных товаров
                UpdateProductsGrid();
                
                MessageBox.Show("Анализ данных завершен.", "Информация", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при анализе данных: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        
        /// <summary>
        /// Обработчик события формирования прогнозов
        /// </summary>
        private void GenerateForecastsMenuItem_Click(object sender, EventArgs e)
        {
            if (_unifiedProducts == null || _unifiedProducts.Count == 0)
            {
                MessageBox.Show("Сначала необходимо обработать и проанализировать данные.", "Предупреждение", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                Cursor = Cursors.WaitCursor;
                
                // Формирование прогнозов с использованием настроек
                DateTime startDate = DateTime.Now;
                DateTime endDate = startDate.AddDays(_forecastSettings.DaysAhead);
                
                _forecasts = _forecastEngine.GenerateFullForecasts(_unifiedProducts, startDate, endDate);
                
                // Расчет приоритетов заказов
                _forecasts = _recommendationSystem.CalculateOrderPriorities(_forecasts);
                
                // Сохранение прогнозов
                _recommendationSystem.SaveRecommendations(_forecasts, _forecastsFilePath);
                
                // Обновление таблицы прогнозов
                UpdateForecastsGrid();
                
                MessageBox.Show($"Сформировано {_forecasts.Count} прогнозов на период с {startDate.ToShortDateString()} по {endDate.ToShortDateString()}.", 
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании прогнозов: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        
        /// <summary>
        /// Обработчик события открытия таблицы заказов
        /// </summary>
        private void OrderTableMenuItem_Click(object sender, EventArgs e)
        {
            if (_forecasts == null || _forecasts.Count == 0)
            {
                MessageBox.Show("Сначала необходимо сформировать прогнозы.", "Предупреждение", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                // Создание и отображение формы таблицы заказов
                var tableForm = new OrderTableForm(_forecasts, _recommendationSystem, _forecastSettings);
                tableForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии таблицы заказов: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Обработчик события открытия отчета по товарам
        /// </summary>
        private void ProductReportMenuItem_Click(object sender, EventArgs e)
        {
            if (_unifiedProducts == null || _unifiedProducts.Count == 0)
            {
                MessageBox.Show("Сначала необходимо обработать данные.", "Предупреждение", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                // Создание и отображение формы отчета по товарам
                var productReportForm = new ProductReportForm(_unifiedProducts);
                productReportForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии отчета по товарам: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Обработчик события открытия окна "О программе"
        /// </summary>
        private void AboutMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Прогнозирование заявок\n\n" +
                "Версия 1.0\n\n" +
                "Программа для прогнозирования будущих заявок на основе исторических данных.\n\n" +
                " 2025",
                "О программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// Обработчик события открытия окна настроек прогноза
        /// </summary>
        private void ForecastSettingsMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Создаем форму для редактирования настроек прогноза
                using (var form = new Form())
                {
                    form.Text = "Настройки прогноза";
                    form.Size = new Size(950, 650);
                    form.StartPosition = FormStartPosition.CenterParent;
                    form.MinimizeBox = false;
                    form.MaximizeBox = true;
                    form.FormBorderStyle = FormBorderStyle.Sizable;
                    
                    // Создаем главную панель с вертикальным расположением
                    var mainPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        RowCount = 2,
                        Padding = new Padding(10)
                    };
                    
                    mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                    mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
                    
                    // Создаем панель с параметрами
                    var settingsPanel = new Panel
                    {
                        Dock = DockStyle.Fill,
                        AutoScroll = true,
                        Padding = new Padding(5)
                    };
                    
                    // Добавляем заголовок
                    var headerLabel = new Label
                    {
                        Text = "Настройки параметров прогнозирования",
                        Font = new Font(form.Font.FontFamily, 12, FontStyle.Bold),
                        Dock = DockStyle.Top,
                        Height = 30,
                        TextAlign = ContentAlignment.MiddleCenter
                    };
                    
                    // Добавляем описание
                    var descriptionHeaderLabel = new Label
                    {
                        Text = "Эти параметры влияют на алгоритм прогнозирования заказов и определение их приоритетов.",
                        Dock = DockStyle.Top,
                        Height = 30,
                        TextAlign = ContentAlignment.MiddleCenter,
                        ForeColor = Color.DarkBlue
                    };
                    
                    settingsPanel.Controls.Add(descriptionHeaderLabel);
                    settingsPanel.Controls.Add(headerLabel);
                    
                    // Словарь для хранения элементов управления значениями
                    var valueControls = new Dictionary<string, Control>();
                    
                    // Создаем таблицу настроек
                    var settingsTable = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                        Margin = new Padding(10, 70, 10, 10)
                    };
                    
                    // Настраиваем колонки и строки
                    settingsTable.ColumnCount = 2;
                    settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Имя параметра
                    settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75)); // Значение и описание
                    
                    // Добавляем настройки в виде групп
                    AddSettingGroup(settingsTable, 0, "Количество дней для прогноза", _forecastSettings.DaysAhead.ToString(), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.DaysAhead)), valueControls, nameof(ForecastSettings.DaysAhead));
                    
                    AddSettingGroup(settingsTable, 1, "Мин. кол-во заказов для высокой уверенности", _forecastSettings.MinOrderHistoryForHighConfidence.ToString(), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.MinOrderHistoryForHighConfidence)), valueControls, nameof(ForecastSettings.MinOrderHistoryForHighConfidence));
                    
                    AddSettingGroup(settingsTable, 2, "Мин. кол-во заказов для средней уверенности", _forecastSettings.MinOrderHistoryForMediumConfidence.ToString(), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.MinOrderHistoryForMediumConfidence)), valueControls, nameof(ForecastSettings.MinOrderHistoryForMediumConfidence));
                    
                    AddSettingGroup(settingsTable, 3, "Коэффициент запаса для даты размещения", _forecastSettings.SafetyFactorForOrderPlacement.ToString("F2"), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.SafetyFactorForOrderPlacement)), valueControls, nameof(ForecastSettings.SafetyFactorForOrderPlacement));
                    
                    AddSettingGroup(settingsTable, 4, "Порог стабильности объемов (%)", _forecastSettings.StableVolumeThreshold.ToString("F1"), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.StableVolumeThreshold)), valueControls, nameof(ForecastSettings.StableVolumeThreshold));
                    
                    AddSettingGroup(settingsTable, 5, "Порог высокого приоритета (дни)", _forecastSettings.HighPriorityThreshold.ToString(), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.HighPriorityThreshold)), valueControls, nameof(ForecastSettings.HighPriorityThreshold));
                    
                    AddSettingGroup(settingsTable, 6, "Порог среднего приоритета (дни)", _forecastSettings.MediumPriorityThreshold.ToString(), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.MediumPriorityThreshold)), valueControls, nameof(ForecastSettings.MediumPriorityThreshold));
                    
                    AddSettingGroup(settingsTable, 7, "Порог низкого приоритета (дни)", _forecastSettings.LowPriorityThreshold.ToString(), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.LowPriorityThreshold)), valueControls, nameof(ForecastSettings.LowPriorityThreshold));
                    
                    AddSettingGroup(settingsTable, 8, "Коэффициент сезонности по умолчанию", _forecastSettings.DefaultSeasonalityCoefficient.ToString("F2"), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.DefaultSeasonalityCoefficient)), valueControls, nameof(ForecastSettings.DefaultSeasonalityCoefficient));
                    
                    AddSettingGroup(settingsTable, 9, "Минимальный порог уверенности (%)", _forecastSettings.MinConfidenceThreshold.ToString("F1"), 
                        ForecastSettings.GetParameterDescription(nameof(ForecastSettings.MinConfidenceThreshold)), valueControls, nameof(ForecastSettings.MinConfidenceThreshold));
                    
                    settingsPanel.Controls.Add(settingsTable);
                    
                    // Добавляем панель кнопок
                    var buttonPanel = new Panel();
                    buttonPanel.Dock = DockStyle.Fill;
                    
                    var saveButton = new Button();
                    saveButton.Text = "Сохранить";
                    saveButton.DialogResult = DialogResult.OK;
                    saveButton.Width = 120;
                    saveButton.Height = 30;
                    saveButton.Location = new Point(form.Width / 2 - 130, 10);
                    saveButton.BackColor = Color.LightGreen;
                    saveButton.FlatStyle = FlatStyle.Flat;
                    saveButton.Click += (s, args) => 
                    {
                        try
                        {
                            // Обновляем настройки из элементов управления
                            _forecastSettings.DaysAhead = int.Parse(((TextBox)valueControls[nameof(ForecastSettings.DaysAhead)]).Text);
                            _forecastSettings.MinOrderHistoryForHighConfidence = int.Parse(((TextBox)valueControls[nameof(ForecastSettings.MinOrderHistoryForHighConfidence)]).Text);
                            _forecastSettings.MinOrderHistoryForMediumConfidence = int.Parse(((TextBox)valueControls[nameof(ForecastSettings.MinOrderHistoryForMediumConfidence)]).Text);
                            _forecastSettings.SafetyFactorForOrderPlacement = double.Parse(((TextBox)valueControls[nameof(ForecastSettings.SafetyFactorForOrderPlacement)]).Text);
                            _forecastSettings.StableVolumeThreshold = double.Parse(((TextBox)valueControls[nameof(ForecastSettings.StableVolumeThreshold)]).Text);
                            _forecastSettings.HighPriorityThreshold = int.Parse(((TextBox)valueControls[nameof(ForecastSettings.HighPriorityThreshold)]).Text);
                            _forecastSettings.MediumPriorityThreshold = int.Parse(((TextBox)valueControls[nameof(ForecastSettings.MediumPriorityThreshold)]).Text);
                            _forecastSettings.LowPriorityThreshold = int.Parse(((TextBox)valueControls[nameof(ForecastSettings.LowPriorityThreshold)]).Text);
                            _forecastSettings.DefaultSeasonalityCoefficient = double.Parse(((TextBox)valueControls[nameof(ForecastSettings.DefaultSeasonalityCoefficient)]).Text);
                            _forecastSettings.MinConfidenceThreshold = double.Parse(((TextBox)valueControls[nameof(ForecastSettings.MinConfidenceThreshold)]).Text);
                            
                            // Сохраняем настройки в файл
                            _forecastSettings.SaveToFile(_settingsFilePath);
                            
                            form.DialogResult = DialogResult.OK;
                            form.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                $"Ошибка при сохранении настроек: {ex.Message}",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                    };
                    
                    var cancelButton = new Button();
                    cancelButton.Text = "Отмена";
                    cancelButton.DialogResult = DialogResult.Cancel;
                    cancelButton.Width = 120;
                    cancelButton.Height = 30;
                    cancelButton.Location = new Point(form.Width / 2 + 10, 10);
                    cancelButton.FlatStyle = FlatStyle.Flat;
                    
                    buttonPanel.Controls.Add(saveButton);
                    buttonPanel.Controls.Add(cancelButton);
                    
                    // Добавляем панели в главную панель
                    mainPanel.Controls.Add(settingsPanel, 0, 0);
                    mainPanel.Controls.Add(buttonPanel, 0, 1);
                    
                    // Добавляем главную панель в форму
                    form.Controls.Add(mainPanel);
                    form.AcceptButton = saveButton;
                    form.CancelButton = cancelButton;
                    
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Обновляем статус
                        var statusStrip = this.Controls.OfType<StatusStrip>().FirstOrDefault();
                        if (statusStrip != null)
                        {
                            var statusLabel = statusStrip.Items.OfType<ToolStripStatusLabel>().FirstOrDefault();
                            if (statusLabel != null)
                            {
                                statusLabel.Text = "Настройки прогноза обновлены";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при работе с настройками прогноза: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Добавление группы настроек в таблицу
        /// </summary>
        private void AddSettingGroup(TableLayoutPanel tableLayoutPanel, int rowIndex, string parameterName, string value, 
            string description, Dictionary<string, Control> valueControls, string parameterKey)
        {
            // Создаем метку с названием параметра
            var nameLabel = new Label
            {
                Text = parameterName,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Margin = new Padding(5),
                AutoSize = true
            };
            
            // Создаем панель для значения и описания
            var valuePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            
            valuePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Для значения
            valuePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Для описания
            
            // Создаем панель для текстового поля и подсказки
            var valueInputPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 30
            };
            
            // Создаем текстовое поле для значения
            var valueTextBox = new TextBox
            {
                Text = value,
                TextAlign = HorizontalAlignment.Center,
                Width = 100,
                Location = new Point(10, 5),
                BorderStyle = BorderStyle.FixedSingle
            };
            
            valueInputPanel.Controls.Add(valueTextBox);
            
            // Создаем метку с описанием
            var descriptionLabel = new Label
            {
                Text = description,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                Margin = new Padding(10, 5, 10, 10),
                AutoSize = false,
                MaximumSize = new Size(1000, 0),
                AutoEllipsis = false
            };
            
            // Устанавливаем перенос текста для описания
            descriptionLabel.AutoSize = true;
            
            // Добавляем элементы в панель значения
            valuePanel.Controls.Add(valueInputPanel, 0, 0);
            valuePanel.Controls.Add(descriptionLabel, 0, 1);
            
            // Добавляем элементы в таблицу настроек
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel.Controls.Add(nameLabel, 0, rowIndex);
            tableLayoutPanel.Controls.Add(valuePanel, 1, rowIndex);
            
            // Сохраняем ссылку на текстовое поле
            valueControls[parameterKey] = valueTextBox;
        }
        
        /// <summary>
        /// Обработчик события открытия окна настроек соответствий
        /// </summary>
        private void MappingMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем наличие файла соответствий
                bool fileExists = File.Exists(_mappingFilePath);
                
                if (!fileExists)
                {
                    // Создаем пустой файл соответствий, если он не существует
                    File.WriteAllText(_mappingFilePath, "[]"); // Пустой массив JSON
                    MessageBox.Show(
                        "Файл соответствий не найден. Создан новый пустой файл.",
                        "Информация",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                
                // Открываем редактор соответствий
                using (var mappingEditor = new MappingEditorForm(_mappingFilePath))
                {
                    if (mappingEditor.ShowDialog() == DialogResult.OK)
                    {
                        // Перезагрузить данные соответствий, если они были изменены
                        LoadSavedData();
                        
                        // Обновить интерфейс
                        UpdateDataGrid();
                        UpdateProductsGrid();
                        UpdateForecastsGrid();
                        
                        // Обновить статус
                        var statusStrip = this.Controls.OfType<StatusStrip>().FirstOrDefault();
                        if (statusStrip != null)
                        {
                            var statusLabel = statusStrip.Items.OfType<ToolStripStatusLabel>().FirstOrDefault();
                            if (statusLabel != null)
                            {
                                statusLabel.Text = "Настройки соответствий обновлены";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при работе с файлом соответствий: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        #endregion
        
        #region Вспомогательные методы
        
        /// <summary>
        /// Загрузка данных из Excel файла
        /// </summary>
        private void LoadExcelData()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                
                // Загрузка данных из Excel файла
                _orderItems = _dataProcessor.LoadData(_excelFilePath);
                
                // Обновление таблицы данных
                UpdateDataGrid();
                
                MessageBox.Show($"Загружено {_orderItems.Count} записей из файла.", "Информация", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        
        /// <summary>
        /// Обновление таблицы данных
        /// </summary>
        private void UpdateDataGrid()
        {
            var dataGridView = this.Controls.Find("dataGridView", true).FirstOrDefault() as DataGridView;
            
            if (dataGridView != null && _orderItems != null)
            {
                // Создание источника данных для таблицы
                var dataSource = _orderItems.Select(item => new
                {
                    Дата_заявки = item.OrderDate.ToShortDateString(),
                    Номер_заявки = item.OrderNumber,
                    Номер_позиции = item.PositionNumber,
                    Наименование = item.ProductName,
                    Артикул = item.ArticleNumber,
                    Заказано = item.OrderedQuantity,
                    Поставлено = item.DeliveredQuantity,
                    Дата_поставки = item.DeliveryDate.HasValue ? item.DeliveryDate.Value.ToShortDateString() : "",
                    Примечание = item.Notes
                }).ToList();
                
                dataGridView.DataSource = dataSource;
            }
        }
        
        /// <summary>
        /// Обновление таблицы унифицированных товаров
        /// </summary>
        private void UpdateProductsGrid()
        {
            var productsGridView = this.Controls.Find("productsGridView", true).FirstOrDefault() as DataGridView;
            
            if (productsGridView != null && _unifiedProducts != null)
            {
                // Создание источника данных для таблицы
                var dataSource = _unifiedProducts.Select(product => new
                {
                    Унифицированный_артикул = product.UnifiedArticle,
                    Наименование = product.PrimaryName,
                    Количество_заказов = product.OrderHistory.Count,
                    Средний_интервал = Math.Round(product.AverageOrderInterval, 1),
                    Среднее_количество = Math.Round(product.AverageOrderQuantity, 1),
                    Средний_срок_поставки = Math.Round(product.AverageDeliveryTime, 1),
                    Последний_заказ = product.LastOrderDate.HasValue ? product.LastOrderDate.Value.ToShortDateString() : "",
                    Следующий_заказ = product.NextPredictedOrderDate.HasValue ? product.NextPredictedOrderDate.Value.ToShortDateString() : "",
                    Рекомендуемое_количество = Math.Round(product.RecommendedQuantity, 0),
                    Дата_размещения = product.OptimalOrderPlacementDate.HasValue ? product.OptimalOrderPlacementDate.Value.ToShortDateString() : ""
                }).ToList();
                
                productsGridView.DataSource = dataSource;
            }
        }
        
        /// <summary>
        /// Обновление таблицы прогнозов
        /// </summary>
        private void UpdateForecastsGrid()
        {
            var forecastsGridView = this.Controls.Find("forecastsGridView", true).FirstOrDefault() as DataGridView;
            
            if (forecastsGridView != null && _forecasts != null)
            {
                // Создание источника данных для таблицы
                var dataSource = _forecasts.Select(forecast => new
                {
                    Приоритет = forecast.Priority,
                    Артикул = forecast.UnifiedArticle,
                    Наименование = forecast.ProductName,
                    Дата_заказа = forecast.NextOrderDate.ToShortDateString(),
                    Количество = Math.Round(forecast.RecommendedQuantity, 0),
                    Дата_размещения = forecast.OptimalOrderPlacementDate.ToShortDateString(),
                    Уверенность = $"{forecast.Confidence:F1}%",
                    Примечание = forecast.Notes
                }).ToList();
                
                forecastsGridView.DataSource = dataSource;
                
                // Настройка цветов для приоритетов
                forecastsGridView.CellFormatting += (sender, e) =>
                {
                    if (e.ColumnIndex == 0 && e.RowIndex >= 0) // Столбец "Приоритет"
                    {
                        int priority = Convert.ToInt32(e.Value);
                        
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
                };
            }
        }
        
        #endregion
    }
}
