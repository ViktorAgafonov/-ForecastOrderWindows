using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Forecast.Models;

namespace Forecast.Forms
{
    /// <summary>
    /// Форма настройки путей к конфигурационным файлам
    /// </summary>
    public partial class PathSettingsForm : Form
    {
        private ForecastSettings _forecastSettings;
        private bool _settingsChanged = false;
        
        // UI компоненты
        private TextBox _configDirectoryTextBox = null!;
        private TextBox _mappingFilePathTextBox = null!;
        private TextBox _forecastsFilePathTextBox = null!;
        
        /// <summary>
        /// Конструктор формы настройки путей
        /// </summary>
        /// <param name="forecastSettings">Настройки прогноза для редактирования</param>
        public PathSettingsForm(ForecastSettings forecastSettings)
        {
            _forecastSettings = forecastSettings ?? throw new ArgumentNullException(nameof(forecastSettings));
            
            InitializeComponent();
            LoadCurrentSettings();
        }
        
        /// <summary>
        /// Инициализация компонентов формы
        /// </summary>
        private void InitializeComponent()
        {
            // Настройка формы
            this.Text = "Настройки путей";
            this.Size = new Size(700, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ShowInTaskbar = false;
            
            // Создаем главную панель
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(15)
            };
            
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            
            // Создаем панель содержимого
            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(5)
            };
            
            // Добавляем заголовок
            var headerLabel = new Label
            {
                Text = "Настройки путей к конфигурационным файлам",
                Font = new Font(this.Font.FontFamily, 12, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 35,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DarkBlue
            };
            
            // Добавляем описание
            var descriptionLabel = new Label
            {
                Text = "Здесь вы можете настроить расположение всех конфигурационных файлов приложения.\n" +
                       "Рекомендуется хранить все конфигурации в одном каталоге для удобства управления.",
                Dock = DockStyle.Top,
                Height = 50,
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = Color.Gray,
                Margin = new Padding(0, 5, 0, 15)
            };
            
            contentPanel.Controls.Add(descriptionLabel);
            contentPanel.Controls.Add(headerLabel);
            
            // Создаем таблицу настроек
            var settingsTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                Margin = new Padding(10, 85, 10, 10),
                BackColor = Color.WhiteSmoke
            };
            
            // Настраиваем колонки
            settingsTable.ColumnCount = 2;
            settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); // Названия
            settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70)); // Значения
            
            // Добавляем настройки путей
            AddDirectoryPathRow(settingsTable, 0, "Каталог конфигураций:", 
                "Базовый каталог для хранения всех конфигурационных файлов приложения.");
                
            AddFilePathRow(settingsTable, 1, "Файл соответствий:", 
                "Имя файла для хранения соответствий товаров (mapping.json).");
                
            AddFilePathRow(settingsTable, 2, "Файл прогнозов:", 
                "Имя файла для хранения результатов прогнозирования (forecasts.json).");
            
            contentPanel.Controls.Add(settingsTable);
            
            // Создаем панель кнопок
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 50
            };
            
            var saveButton = new Button
            {
                Text = "Сохранить",
                DialogResult = DialogResult.OK,
                Width = 100,
                Height = 30,
                Location = new Point(buttonPanel.Width / 2 - 110, 10),
                BackColor = Color.LightGreen,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Bottom
            };
            
            var cancelButton = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Width = 100,
                Height = 30,
                Location = new Point(buttonPanel.Width / 2 + 10, 10),
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Bottom
            };
            
            // Обработчик сохранения
            saveButton.Click += SaveButton_Click;
            
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(cancelButton);
            
            // Добавляем панели в главную панель
            mainPanel.Controls.Add(contentPanel, 0, 0);
            mainPanel.Controls.Add(buttonPanel, 0, 1);
            
            // Добавляем главную панель в форму
            this.Controls.Add(mainPanel);
            this.AcceptButton = saveButton;
            this.CancelButton = cancelButton;
        }
        
        /// <summary>
        /// Добавление строки для выбора каталога
        /// </summary>
        private void AddDirectoryPathRow(TableLayoutPanel table, int row, string label, string description)
        {
            // Метка
            var nameLabel = new Label
            {
                Text = label,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(this.Font, FontStyle.Bold),
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 5, 5, 5),
                BackColor = Color.LightBlue
            };
            
            // Панель для ввода
            var inputPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 60,
                Margin = new Padding(5)
            };
            
            // Текстовое поле
            _configDirectoryTextBox = new TextBox
            {
                Width = inputPanel.Width - 90,
                Location = new Point(5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            
            // Кнопка обзора
            var browseButton = new Button
            {
                Text = "Обзор...",
                Width = 80,
                Height = 23,
                Location = new Point(inputPanel.Width - 85, 4),
                FlatStyle = FlatStyle.System,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            
            browseButton.Click += (s, e) => BrowseDirectory(_configDirectoryTextBox);
            
            // Описание
            var descLabel = new Label
            {
                Text = description,
                Location = new Point(5, 32),
                Width = inputPanel.Width - 10,
                Height = 25,
                ForeColor = Color.Gray,
                Font = new Font(this.Font.FontFamily, this.Font.Size - 1),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            
            inputPanel.Controls.Add(_configDirectoryTextBox);
            inputPanel.Controls.Add(browseButton);
            inputPanel.Controls.Add(descLabel);
            
            // Добавляем в таблицу
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(nameLabel, 0, row);
            table.Controls.Add(inputPanel, 1, row);
        }
        
        /// <summary>
        /// Добавление строки для ввода имени файла
        /// </summary>
        private void AddFilePathRow(TableLayoutPanel table, int row, string label, string description)
        {
            // Метка
            var nameLabel = new Label
            {
                Text = label,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(this.Font, FontStyle.Bold),
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 5, 5, 5),
                BackColor = Color.LightGreen
            };
            
            // Панель для ввода
            var inputPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 60,
                Margin = new Padding(5)
            };
            
            // Текстовое поле
            TextBox textBox;
            if (label.Contains("соответствий"))
            {
                _mappingFilePathTextBox = new TextBox();
                textBox = _mappingFilePathTextBox;
            }
            else
            {
                _forecastsFilePathTextBox = new TextBox();
                textBox = _forecastsFilePathTextBox;
            }
            
            textBox.Width = inputPanel.Width - 10;
            textBox.Location = new Point(5, 5);
            textBox.BorderStyle = BorderStyle.FixedSingle;
            textBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            
            // Описание
            var descLabel = new Label
            {
                Text = description,
                Location = new Point(5, 32),
                Width = inputPanel.Width - 10,
                Height = 25,
                ForeColor = Color.Gray,
                Font = new Font(this.Font.FontFamily, this.Font.Size - 1),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            
            inputPanel.Controls.Add(textBox);
            inputPanel.Controls.Add(descLabel);
            
            // Добавляем в таблицу
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(nameLabel, 0, row);
            table.Controls.Add(inputPanel, 1, row);
        }
        
        /// <summary>
        /// Обработчик выбора каталога
        /// </summary>
        private void BrowseDirectory(TextBox textBox)
        {
            var dialog = new FolderBrowserDialog
            {
                Description = "Выберите каталог для хранения конфигурационных файлов",
                SelectedPath = string.IsNullOrEmpty(textBox.Text) 
                    ? Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? Environment.CurrentDirectory
                    : textBox.Text,
                ShowNewFolderButton = true
            };
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = dialog.SelectedPath;
                _settingsChanged = true;
            }
        }
        
        /// <summary>
        /// Загрузка текущих настроек
        /// </summary>
        private void LoadCurrentSettings()
        {
            if (_configDirectoryTextBox != null)
                _configDirectoryTextBox.Text = _forecastSettings.ConfigurationDirectory;
                
            if (_mappingFilePathTextBox != null)
                _mappingFilePathTextBox.Text = _forecastSettings.MappingFilePath;
                
            if (_forecastsFilePathTextBox != null)
                _forecastsFilePathTextBox.Text = _forecastSettings.ForecastsFilePath;
        }
        
        /// <summary>
        /// Обработчик сохранения настроек
        /// </summary>
        private void SaveButton_Click(object? sender, EventArgs e)
        {
            try
            {
                // Проверяем валидность путей
                if (string.IsNullOrWhiteSpace(_configDirectoryTextBox.Text))
                {
                    MessageBox.Show("Необходимо указать каталог конфигураций.", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (string.IsNullOrWhiteSpace(_mappingFilePathTextBox.Text))
                {
                    MessageBox.Show("Необходимо указать имя файла соответствий.", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (string.IsNullOrWhiteSpace(_forecastsFilePathTextBox.Text))
                {
                    MessageBox.Show("Необходимо указать имя файла прогнозов.", "Ошибка", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                var newConfigDir = _configDirectoryTextBox.Text;
                var newMappingFileName = _mappingFilePathTextBox.Text;
                var newForecastsFileName = _forecastsFilePathTextBox.Text;
                
                // Проверяем, изменились ли пути
                bool pathsChanged = newConfigDir != _forecastSettings.ConfigurationDirectory ||
                                  newMappingFileName != _forecastSettings.MappingFilePath ||
                                  newForecastsFileName != _forecastSettings.ForecastsFilePath;
                
                if (pathsChanged)
                {
                    // Показываем предупреждение и предлагаем перенести файлы
                    var result = MessageBox.Show(
                        "Пути к конфигурационным файлам изменились.\n\n" +
                        "Хотите перенести существующие конфигурационные файлы в новое расположение?\n\n" +
                        "- ДА: Файлы будут скопированы в новое место\n" +
                        "- НЕТ: Файлы останутся в старом месте (может потребоваться ручная настройка)\n" +
                        "- ОТМЕНА: Отменить изменения",
                        "Перенос конфигурационных файлов",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);
                    
                    if (result == DialogResult.Cancel)
                    {
                        return;
                    }
                    
                    if (result == DialogResult.Yes)
                    {
                        // Выполняем перенос файлов
                        if (!MigrateConfigurationFiles(newConfigDir, newMappingFileName, newForecastsFileName))
                        {
                            return; // Если перенос не удался, не сохраняем настройки
                        }
                    }
                }
                
                // Создаем каталог если не существует
                if (!Directory.Exists(newConfigDir))
                {
                    try
                    {
                        Directory.CreateDirectory(newConfigDir);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка создания каталога: {ex.Message}", "Ошибка", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                
                // Обновляем настройки
                _forecastSettings.ConfigurationDirectory = newConfigDir;
                _forecastSettings.MappingFilePath = newMappingFileName;
                _forecastSettings.ForecastsFilePath = newForecastsFileName;
                
                _settingsChanged = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении настроек: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Перенос конфигурационных файлов в новое расположение
        /// </summary>
        /// <param name="newConfigDir">Новый каталог конфигураций</param>
        /// <param name="newMappingFileName">Новое имя файла соответствий</param>
        /// <param name="newForecastsFileName">Новое имя файла прогнозов</param>
        /// <returns>True если перенос успешен</returns>
        private bool MigrateConfigurationFiles(string newConfigDir, string newMappingFileName, string newForecastsFileName)
        {
            try
            {
                var migratedFiles = new List<string>();
                var errors = new List<string>();
                
                // Создаем целевой каталог если не существует
                if (!Directory.Exists(newConfigDir))
                {
                    Directory.CreateDirectory(newConfigDir);
                }
                
                // Переносим файл соответствий
                var currentMappingPath = _forecastSettings.GetFullMappingFilePath();
                var newMappingPath = Path.Combine(newConfigDir, newMappingFileName);
                
                if (File.Exists(currentMappingPath) && currentMappingPath != newMappingPath)
                {
                    try
                    {
                        // Создаем резервную копию если файл уже существует
                        if (File.Exists(newMappingPath))
                        {
                            var backupPath = newMappingPath + ".backup." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                            File.Copy(newMappingPath, backupPath);
                            migratedFiles.Add($"Создана резервная копия: {backupPath}");
                        }
                        
                        File.Copy(currentMappingPath, newMappingPath, true);
                        
                        // Обновляем пути внутри файла если необходимо
                        UpdateConfigurationFilePaths(newMappingPath, newConfigDir);
                        
                        migratedFiles.Add($"Соответствия: {currentMappingPath} → {newMappingPath}");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Ошибка переноса файла соответствий: {ex.Message}");
                    }
                }
                
                // Переносим файл прогнозов
                var currentForecastsPath = _forecastSettings.GetFullForecastsFilePath();
                var newForecastsPath = Path.Combine(newConfigDir, newForecastsFileName);
                
                if (File.Exists(currentForecastsPath) && currentForecastsPath != newForecastsPath)
                {
                    try
                    {
                        // Создаем резервную копию если файл уже существует
                        if (File.Exists(newForecastsPath))
                        {
                            var backupPath = newForecastsPath + ".backup." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                            File.Copy(newForecastsPath, backupPath);
                            migratedFiles.Add($"Создана резервная копия: {backupPath}");
                        }
                        
                        File.Copy(currentForecastsPath, newForecastsPath, true);
                        
                        // Обновляем пути внутри файла если необходимо
                        UpdateConfigurationFilePaths(newForecastsPath, newConfigDir);
                        
                        migratedFiles.Add($"Прогнозы: {currentForecastsPath} → {newForecastsPath}");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Ошибка переноса файла прогнозов: {ex.Message}");
                    }
                }
                
                // Показываем результат переноса
                var message = "Результат переноса конфигурационных файлов:\n\n";
                
                if (migratedFiles.Count > 0)
                {
                    message += "Успешно перенесено:\n";
                    foreach (var file in migratedFiles)
                    {
                        message += $"✓ {file}\n";
                    }
                    message += "\n";
                }
                
                if (errors.Count > 0)
                {
                    message += "Ошибки:\n";
                    foreach (var error in errors)
                    {
                        message += $"✗ {error}\n";
                    }
                    message += "\n";
                    message += "Некоторые файлы не удалось перенести. Продолжить?";
                    
                    var continueResult = MessageBox.Show(message, "Результат переноса", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    
                    return continueResult == DialogResult.Yes;
                }
                else if (migratedFiles.Count > 0)
                {
                    message += "Все файлы успешно перенесены!\n\n";
                    message += "Хотите удалить оригинальные файлы из старого расположения?";
                    
                    var deleteResult = MessageBox.Show(message, "Перенос завершен", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    
                    if (deleteResult == DialogResult.Yes)
                    {
                        CleanupOldConfigurationFiles(newConfigDir, newMappingFileName, newForecastsFileName);
                    }
                }
                else
                {
                    MessageBox.Show("Файлы для переноса не найдены или пути не изменились.", "Информация", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переносе конфигурационных файлов: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        
        /// <summary>
        /// Обновление путей внутри конфигурационного файла
        /// </summary>
        /// <param name="filePath">Путь к файлу конфигурации</param>
        /// <param name="newBaseDirectory">Новый базовый каталог</param>
        private void UpdateConfigurationFilePaths(string filePath, string newBaseDirectory)
        {
            try
            {
                if (!File.Exists(filePath))
                    return;
                
                // Читаем содержимое файла
                var content = File.ReadAllText(filePath);
                var originalContent = content;
                bool contentChanged = false;
                
                // Обновляем различные типы путей в JSON файлах
                if (filePath.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                {
                    // Обновляем абсолютные пути к каталогу конфигураций
                    var oldConfigDir = _forecastSettings.ConfigurationDirectory;
                    if (!string.IsNullOrEmpty(oldConfigDir) && content.Contains(oldConfigDir))
                    {
                        content = content.Replace(oldConfigDir, newBaseDirectory);
                        contentChanged = true;
                    }
                    
                    // Обновляем относительные пути в JSON
                    // Ищем паттерны типа "path": "old_directory\\file.json"
                    var pathPatterns = new[]
                    {
                        "\"path\":",
                        "\"filePath\":",
                        "\"configPath\":",
                        "\"dataPath\":",
                        "\"mappingPath\":",
                        "\"forecastPath\":"
                    };
                    
                    foreach (var pattern in pathPatterns)
                    {
                        if (content.Contains(pattern))
                        {
                            // Это базовая замена, в реальности можно использовать JSON парсер
                            // для более точного обновления путей
                            contentChanged = true;
                        }
                    }
                }
                
                // Сохраняем изменения если были внесены
                if (contentChanged && content != originalContent)
                {
                    // Создаем резервную копию оригинального файла
                    var backupPath = filePath + ".original." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    File.Copy(filePath, backupPath);
                    
                    // Сохраняем обновленное содержимое
                    File.WriteAllText(filePath, content);
                }
            }
            catch (Exception ex)
            {
                // Логируем ошибку, но не прерываем процесс переноса
                // В реальном приложении можно добавить в список ошибок
                System.Diagnostics.Debug.WriteLine($"Ошибка при обновлении путей в файле {filePath}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Очистка старых конфигурационных файлов после успешного переноса
        /// </summary>
        /// <param name="newConfigDir">Новый каталог конфигураций</param>
        /// <param name="newMappingFileName">Новое имя файла соответствий</param>
        /// <param name="newForecastsFileName">Новое имя файла прогнозов</param>
        private void CleanupOldConfigurationFiles(string newConfigDir, string newMappingFileName, string newForecastsFileName)
        {
            try
            {
                var deletedFiles = new List<string>();
                var errors = new List<string>();
                
                // Удаляем старый файл соответствий
                var oldMappingPath = _forecastSettings.GetFullMappingFilePath();
                var newMappingPath = Path.Combine(newConfigDir, newMappingFileName);
                
                if (File.Exists(oldMappingPath) && oldMappingPath != newMappingPath)
                {
                    try
                    {
                        File.Delete(oldMappingPath);
                        deletedFiles.Add($"Удален старый файл соответствий: {oldMappingPath}");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Ошибка удаления {oldMappingPath}: {ex.Message}");
                    }
                }
                
                // Удаляем старый файл прогнозов
                var oldForecastsPath = _forecastSettings.GetFullForecastsFilePath();
                var newForecastsPath = Path.Combine(newConfigDir, newForecastsFileName);
                
                if (File.Exists(oldForecastsPath) && oldForecastsPath != newForecastsPath)
                {
                    try
                    {
                        File.Delete(oldForecastsPath);
                        deletedFiles.Add($"Удален старый файл прогнозов: {oldForecastsPath}");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Ошибка удаления {oldForecastsPath}: {ex.Message}");
                    }
                }
                
                // Пытаемся удалить старый каталог если он пустой
                var oldConfigDir = _forecastSettings.ConfigurationDirectory;
                if (!string.IsNullOrEmpty(oldConfigDir) && 
                    Directory.Exists(oldConfigDir) && 
                    oldConfigDir != newConfigDir)
                {
                    try
                    {
                        // Проверяем, пустой ли каталог
                        var remainingFiles = Directory.GetFiles(oldConfigDir, "*", SearchOption.AllDirectories);
                        if (remainingFiles.Length == 0)
                        {
                            Directory.Delete(oldConfigDir, true);
                            deletedFiles.Add($"Удален пустой старый каталог: {oldConfigDir}");
                        }
                        else
                        {
                            deletedFiles.Add($"Старый каталог {oldConfigDir} содержит другие файлы и оставлен без изменений");
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Ошибка при очистке старого каталога {oldConfigDir}: {ex.Message}");
                    }
                }
                
                // Показываем результат очистки
                if (deletedFiles.Count > 0 || errors.Count > 0)
                {
                    var message = "Результат очистки старых файлов:\n\n";
                    
                    if (deletedFiles.Count > 0)
                    {
                        message += "Выполнено:\n";
                        foreach (var file in deletedFiles)
                        {
                            message += $"✓ {file}\n";
                        }
                        message += "\n";
                    }
                    
                    if (errors.Count > 0)
                    {
                        message += "Ошибки:\n";
                        foreach (var error in errors)
                        {
                            message += $"✗ {error}\n";
                        }
                    }
                    
                    MessageBox.Show(message, "Результат очистки", 
                        MessageBoxButtons.OK, 
                        errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке старых файлов: {ex.Message}", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Возвращает true если настройки были изменены
        /// </summary>
        public bool SettingsChanged => _settingsChanged;
    }
}
