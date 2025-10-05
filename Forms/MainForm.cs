using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Forecast.Models;
using Forecast.Services;
using System.Text.Json;

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
            
            // Инициализация данных - пустые коллекции по умолчанию
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

            // Добавляем обработчик события закрытия формы
            this.FormClosing += MainForm_FormClosing;

            // Инициализация компонентов формы
            InitializeCustomComponents();
            
            // Загрузка только настроек и соответствий товаров при запуске
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
            exitMenuItem.Click += ExitMenuItem_Click;
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
        /// Загрузка сохраненных настроек и соответствий товаров (без прогнозов)
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
                
                // При запуске не загружаем прогнозы автоматически
                // Пользователь должен сначала открыть файл с данными
                _forecasts = new List<ForecastResult>();
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
        private void OpenFileMenuItem_Click(object? sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel файлы (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*";
                openFileDialog.Title = "Выберите Excel файл с данными заказов";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _excelFilePath = openFileDialog.FileName;
                    LoadAndProcessExcelData();
                }
            }
        }
        
        /// <summary>
        /// Обработчик события обработки данных
        /// </summary>
        private void ProcessDataMenuItem_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_orderItems == null || _orderItems.Count == 0)
                {
                    MessageBox.Show("Сначала загрузите данные из Excel файла.", "Предупреждение", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                Cursor = Cursors.WaitCursor;
                
                // Обработка данных и создание унифицированных товаров
                _unifiedProducts = _dataProcessor.UnifyProducts(_orderItems);
                
                // Обновление таблицы унифицированных товаров
                UpdateProductsGrid();
                
                MessageBox.Show($"Обработано {_unifiedProducts.Count} уникальных товаров.", "Информация", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        private void AnalyzeDataMenuItem_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_unifiedProducts == null || _unifiedProducts.Count == 0)
                {
                    MessageBox.Show("Сначала обработайте данные.", "Предупреждение", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                Cursor = Cursors.WaitCursor;
                
                // Анализ данных и расчет статистик
                _orderAnalyzer.AnalyzeOrderFrequency(_unifiedProducts);
                _orderAnalyzer.AnalyzeOrderVolumes(_unifiedProducts);
                _orderAnalyzer.AnalyzeSeasonality(_unifiedProducts);
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
        private void GenerateForecastsMenuItem_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_unifiedProducts == null || _unifiedProducts.Count == 0)
                {
                    MessageBox.Show("Сначала обработайте и проанализируйте данные.", "Предупреждение", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                Cursor = Cursors.WaitCursor;
                
                // Генерация прогнозов
                _forecasts = _forecastEngine.GenerateFullForecasts(_unifiedProducts, DateTime.Now, DateTime.Now.AddDays(_forecastSettings.DaysAhead));
                
                // Обновление таблицы прогнозов
                UpdateForecastsGrid();
                
                MessageBox.Show($"Сгенерировано {_forecasts.Count} прогнозов.", "Информация", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при генерации прогнозов: {ex.Message}", "Ошибка", 
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
        private void OrderTableMenuItem_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_forecasts == null || _forecasts.Count == 0)
                {
                    MessageBox.Show("Сначала сгенерируйте прогнозы.", "Предупреждение", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Создание и отображение формы таблицы заказов
                var orderTableForm = new OrderTableForm(_forecasts, _recommendationSystem, _forecastSettings);
                orderTableForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии таблицы заказов: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Сохранение настроек прогноза
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                if (_forecastSettings != null && !string.IsNullOrEmpty(_settingsFilePath))
                {
                    string json = System.Text.Json.JsonSerializer.Serialize(_forecastSettings);
                    File.WriteAllText(_settingsFilePath, json);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении настроек: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Обработчик события закрытия формы
        /// </summary>
        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            try
            {
                // Сохраняем настройки перед закрытием
                SaveSettings();
                
                // Сохраняем прогнозы, если они есть
                if (_forecasts != null && _forecasts.Count > 0)
                {
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    string json = JsonSerializer.Serialize(_forecasts, options);
                    File.WriteAllText(_forecastsFilePath, json);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении данных: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Обработчик нажатия на пункт меню "Выход"
        /// </summary>
        private void ExitMenuItem_Click(object? sender, EventArgs e)
        {
            Close();
        }
        
        /// <summary>
        /// Обработчик события открытия отчета по товарам
        /// </summary>
        private void ProductReportMenuItem_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_unifiedProducts == null || _unifiedProducts.Count == 0)
                {
                    MessageBox.Show("Сначала обработайте данные.", "Предупреждение", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
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
        private void AboutMenuItem_Click(object? sender, EventArgs e)
        {
            MessageBox.Show(
                "Система прогнозирования заказов\n\n" +
                "Версия 1.0\n" +
                "Разработано для автоматизации процесса планирования заказов\n\n" +
                "Функции:\n" +
                "• Загрузка данных из Excel\n" +
                "• Анализ истории заказов\n" +
                "• Прогнозирование будущих заказов\n" +
                "• Управление соответствиями товаров\n" +
                "• Генерация отчетов",
                "О программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// Обработчик события открытия окна настроек прогноза
        /// </summary>
        private void ForecastSettingsMenuItem_Click(object? sender, EventArgs e)
        {
            // Создание и отображение формы настроек прогнозирования
            var settingsForm = new Form();
            settingsForm.Text = "Настройки прогнозирования";
            settingsForm.Size = new Size(600, 500);
            settingsForm.StartPosition = FormStartPosition.CenterParent;
            settingsForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            settingsForm.MaximizeBox = false;
            settingsForm.MinimizeBox = false;
            
            // Создание таблицы для размещения элементов управления
            var tableLayoutPanel = new TableLayoutPanel();
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.ColumnCount = 2;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
            tableLayoutPanel.Padding = new Padding(10);
            
            // Словарь для хранения элементов управления
            var valueControls = new Dictionary<string, Control>();
            
            int rowIndex = 0;
            
            // Добавление настроек
            AddSettingGroup(tableLayoutPanel, rowIndex++, "MinConfidenceThreshold", 
                _forecastSettings.MinConfidenceThreshold.ToString(), 
                "Минимальный порог уверенности для прогноза (%)", valueControls, "MinConfidenceThreshold");
            
            AddSettingGroup(tableLayoutPanel, rowIndex++, "DaysAhead", 
                _forecastSettings.DaysAhead.ToString(), 
                "Количество дней для прогноза вперед", valueControls, "DaysAhead");
            
            AddSettingGroup(tableLayoutPanel, rowIndex++, "SafetyFactorForOrderPlacement", 
                _forecastSettings.SafetyFactorForOrderPlacement.ToString(), 
                "Коэффициент запаса для расчета оптимальной даты размещения заказа", valueControls, "SafetyFactorForOrderPlacement");
            
            AddSettingGroup(tableLayoutPanel, rowIndex++, "DefaultSeasonalityCoefficient", 
                _forecastSettings.DefaultSeasonalityCoefficient.ToString(), 
                "Коэффициент сезонности по умолчанию", valueControls, "DefaultSeasonalityCoefficient");
            
            AddSettingGroup(tableLayoutPanel, rowIndex++, "StableVolumeThreshold", 
                _forecastSettings.StableVolumeThreshold.ToString(), 
                "Порог стабильности объемов заказов (%)", valueControls, "StableVolumeThreshold");
            
            AddSettingGroup(tableLayoutPanel, rowIndex++, "HighPriorityThreshold", 
                _forecastSettings.HighPriorityThreshold.ToString(), 
                "Порог высокого приоритета (дни)", valueControls, "HighPriorityThreshold");
            
            // Кнопки
            var buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Bottom;
            buttonPanel.Height = 50;
            buttonPanel.Padding = new Padding(10);
            
            var saveButton = new Button();
            saveButton.Text = "Сохранить";
            saveButton.DialogResult = DialogResult.OK;
            saveButton.Location = new Point(settingsForm.Width - 200, 10);
            saveButton.Click += (s, e) =>
            {
                try
                {
                    // Сохранение значений из элементов управления
                    if (valueControls.ContainsKey("MinConfidenceThreshold") && 
                        valueControls["MinConfidenceThreshold"] is TextBox minConfidenceTextBox)
                    {
                        if (double.TryParse(minConfidenceTextBox.Text, out double minConfidence))
                            _forecastSettings.MinConfidenceThreshold = minConfidence;
                    }
                    
                    if (valueControls.ContainsKey("DaysAhead") && 
                        valueControls["DaysAhead"] is TextBox daysAheadTextBox)
                    {
                        if (int.TryParse(daysAheadTextBox.Text, out int daysAhead))
                            _forecastSettings.DaysAhead = daysAhead;
                    }
                    
                    if (valueControls.ContainsKey("SafetyFactorForOrderPlacement") && 
                        valueControls["SafetyFactorForOrderPlacement"] is TextBox safetyFactorTextBox)
                    {
                        if (double.TryParse(safetyFactorTextBox.Text, out double safetyFactor))
                            _forecastSettings.SafetyFactorForOrderPlacement = safetyFactor;
                    }
                    
                    if (valueControls.ContainsKey("DefaultSeasonalityCoefficient") && 
                        valueControls["DefaultSeasonalityCoefficient"] is TextBox seasonalityWeightTextBox)
                    {
                        if (double.TryParse(seasonalityWeightTextBox.Text, out double seasonalityWeight))
                            _forecastSettings.DefaultSeasonalityCoefficient = seasonalityWeight;
                    }
                    
                    if (valueControls.ContainsKey("StableVolumeThreshold") && 
                        valueControls["StableVolumeThreshold"] is TextBox stableVolumeThresholdTextBox)
                    {
                        if (double.TryParse(stableVolumeThresholdTextBox.Text, out double stableVolumeThreshold))
                            _forecastSettings.StableVolumeThreshold = stableVolumeThreshold;
                    }
                    
                    if (valueControls.ContainsKey("HighPriorityThreshold") && 
                        valueControls["HighPriorityThreshold"] is TextBox highPriorityThresholdTextBox)
                    {
                        if (int.TryParse(highPriorityThresholdTextBox.Text, out int highPriorityThreshold))
                            _forecastSettings.HighPriorityThreshold = highPriorityThreshold;
                    }
                    
                    // Сохранение настроек
                    SaveSettings();
                    
                    MessageBox.Show("Настройки сохранены.", "Информация", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении настроек: {ex.Message}", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            
            var cancelButton = new Button();
            cancelButton.Text = "Отмена";
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Location = new Point(settingsForm.Width - 100, 10);
            
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(cancelButton);
            
            settingsForm.Controls.Add(tableLayoutPanel);
            settingsForm.Controls.Add(buttonPanel);
            
            settingsForm.ShowDialog();
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
        private void MappingMenuItem_Click(object? sender, EventArgs e)
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
        /// Загрузка и автоматическая обработка данных из Excel файла
        /// </summary>
        private void LoadAndProcessExcelData()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                UpdateStatus("Загрузка данных из Excel...");

                // Загрузка данных из Excel файла
                _orderItems = _dataProcessor.LoadData(_excelFilePath);
                UpdateDataGrid();
                UpdateStatus($"Загружено {_orderItems.Count} записей. Обработка данных...");

                // Автоматическая обработка данных
                _unifiedProducts = _dataProcessor.UnifyProducts(_orderItems);
                UpdateProductsGrid();
                UpdateStatus($"Обработано {_unifiedProducts.Count} товаров. Анализ данных...");

                // Автоматический анализ данных
                _orderAnalyzer.AnalyzeOrderFrequency(_unifiedProducts);
                _orderAnalyzer.AnalyzeOrderVolumes(_unifiedProducts);
                _orderAnalyzer.AnalyzeSeasonality(_unifiedProducts);
                _orderAnalyzer.AnalyzeDeliveryTimes(_unifiedProducts);
                UpdateProductsGrid();
                UpdateStatus("Анализ завершен. Формирование прогнозов...");

                // Автоматическая генерация прогнозов
                _forecasts = _forecastEngine.GenerateFullForecasts(_unifiedProducts, DateTime.Now,
                    DateTime.Now.AddDays(_forecastSettings.DaysAhead));
                UpdateForecastsGrid();
                UpdateStatus("Готово");

                // Автосохранение базы соответствий
                _dataProcessor.SaveItemMapping(_unifiedProducts, _mappingFilePath);

                MessageBox.Show(
                    $"Обработка завершена:\n\n" +
                    $"• Загружено записей: {_orderItems.Count}\n" +
                    $"• Унифицировано товаров: {_unifiedProducts.Count}\n" +
                    $"• Сгенерировано прогнозов: {_forecasts.Count}",
                    "Информация",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Ошибка");
                MessageBox.Show($"Ошибка при обработке данных: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// Обновление статуса в строке состояния
        /// </summary>
        private void UpdateStatus(string message)
        {
            var statusStrip = this.Controls.OfType<StatusStrip>().FirstOrDefault();
            if (statusStrip != null)
            {
                var statusLabel = statusStrip.Items.OfType<ToolStripStatusLabel>().FirstOrDefault();
                if (statusLabel != null)
                {
                    statusLabel.Text = message;
                    statusStrip.Refresh();
                    Application.DoEvents(); // Обновляем UI
                }
            }
        }

        /// <summary>
        /// Загрузка данных из Excel файла (устаревший метод, оставлен для совместимости)
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
                    Номер_заявки = item.OrderNumber ?? string.Empty,
                    Номер_позиции = item.PositionNumber ?? string.Empty,
                    Наименование = item.ProductName ?? string.Empty,
                    Артикул = item.ArticleNumber ?? string.Empty,
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
                var dataSource = _unifiedProducts.Where(product => product != null).Select(product => new
                {
                    Унифицированный_артикул = product.UnifiedArticle ?? string.Empty,
                    Наименование = product.PrimaryName ?? string.Empty,
                    Количество_заказов = product.OrderHistory?.Count ?? 0,
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
                var dataSource = _forecasts.Where(forecast => forecast != null).Select(forecast => new
                {
                    Приоритет = forecast.Priority,
                    Артикул = forecast.UnifiedArticle ?? string.Empty,
                    Наименование = forecast.ProductName ?? string.Empty,
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
                    if (e.Value != null && e.ColumnIndex == 0 && e.CellStyle != null) // Столбец "Приоритет"
                    {
                        var valueString = e.Value.ToString();
                        if (!string.IsNullOrEmpty(valueString) && int.TryParse(valueString, out int priority))
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
                };
            }
        }
        
        #endregion
    }
}
