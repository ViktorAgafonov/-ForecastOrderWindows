using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using Forecast.Models;

namespace Forecast.Forms
{
    /// <summary>
    /// Форма для редактирования групп соответствий товаров
    /// </summary>
    public class MappingEditorForm : Form
    {
        // Путь к файлу соответствий
        private readonly string _mappingFilePath;
        
        // База данных соответствий
        private MappingDatabase _mappingDatabase = new MappingDatabase();
        
        // Текущая выбранная группа
        private MappingGroup _selectedGroup;
        
        // Компоненты интерфейса
        private SplitContainer _mainSplitContainer;
        private ListBox _groupsListBox;
        private Button _addGroupButton;
        private Button _removeGroupButton;
        private Button _renameGroupButton;
        private TextBox _groupNameTextBox;
        
        private ListBox _variationsListBox;
        private Button _addVariationButton;
        private Button _removeVariationButton;
        private TextBox _nameVariationTextBox;
        private TextBox _articleVariationTextBox;
        private RadioButton _nameRadioButton;
        private RadioButton _articleRadioButton;
        
        private Label _unifiedArticleLabel;
        private TextBox _unifiedArticleTextBox;
        private Label _primaryNameLabel;
        private TextBox _primaryNameTextBox;
        private Button _applyUnifiedButton;
        
        private Button _saveButton;
        private Button _cancelButton;
        
        /// <summary>
        /// Конструктор формы редактора соответствий
        /// </summary>
        /// <param name="mappingFilePath">Путь к файлу соответствий</param>
        public MappingEditorForm(string mappingFilePath)
        {
            _mappingFilePath = mappingFilePath;
            
            // Настройка формы
            Text = "Редактор групп соответствий";
            Size = new Size(1000, 700);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = true;
            FormBorderStyle = FormBorderStyle.Sizable;
            
            // Инициализация компонентов
            InitializeComponents();
            
            // Загрузка данных
            LoadMappingData();
        }
        
        /// <summary>
        /// Инициализация компонентов формы
        /// </summary>
        private void InitializeComponents()
        {
            // Создаем главную панель с разделением на левую и правую части
            _mainSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal
            };
            
            // Создаем меню
            var menuStrip = new MenuStrip();
            
            // Пункт меню "Файл"
            var fileMenuItem = new ToolStripMenuItem("Файл");
            
            var saveMenuItem = new ToolStripMenuItem("Сохранить");
            saveMenuItem.ShortcutKeys = Keys.Control | Keys.S;
            saveMenuItem.Click += (s, e) => SaveMappingData();
            
            var closeMenuItem = new ToolStripMenuItem("Закрыть");
            closeMenuItem.ShortcutKeys = Keys.Alt | Keys.F4;
            closeMenuItem.Click += (s, e) => Close();
            
            fileMenuItem.DropDownItems.Add(saveMenuItem);
            fileMenuItem.DropDownItems.Add(new ToolStripSeparator());
            fileMenuItem.DropDownItems.Add(closeMenuItem);
            
            // Пункт меню "Группы"
            var groupsMenuItem = new ToolStripMenuItem("Группы");
            
            var addGroupMenuItem = new ToolStripMenuItem("Добавить группу");
            addGroupMenuItem.ShortcutKeys = Keys.Control | Keys.N;
            addGroupMenuItem.Click += (s, e) => AddNewGroup();
            
            var removeGroupMenuItem = new ToolStripMenuItem("Удалить группу");
            removeGroupMenuItem.ShortcutKeys = Keys.Control | Keys.Delete;
            removeGroupMenuItem.Click += (s, e) => RemoveSelectedGroup();
            
            var renameGroupMenuItem = new ToolStripMenuItem("Переименовать группу");
            renameGroupMenuItem.ShortcutKeys = Keys.F2;
            renameGroupMenuItem.Click += (s, e) => RenameSelectedGroup();
            
            groupsMenuItem.DropDownItems.Add(addGroupMenuItem);
            groupsMenuItem.DropDownItems.Add(removeGroupMenuItem);
            groupsMenuItem.DropDownItems.Add(renameGroupMenuItem);
            
            // Пункт меню "Вариации"
            var variationsMenuItem = new ToolStripMenuItem("Вариации");
            
            var addVariationMenuItem = new ToolStripMenuItem("Добавить вариацию");
            addVariationMenuItem.ShortcutKeys = Keys.Control | Keys.Insert;
            addVariationMenuItem.Click += (s, e) => AddVariation();
            
            var removeVariationMenuItem = new ToolStripMenuItem("Удалить вариацию");
            removeVariationMenuItem.ShortcutKeys = Keys.Delete;
            removeVariationMenuItem.Click += (s, e) => RemoveSelectedVariation();
            
            variationsMenuItem.DropDownItems.Add(addVariationMenuItem);
            variationsMenuItem.DropDownItems.Add(removeVariationMenuItem);
            
            // Добавляем пункты в меню
            menuStrip.Items.Add(fileMenuItem);
            menuStrip.Items.Add(groupsMenuItem);
            menuStrip.Items.Add(variationsMenuItem);
            
            // Добавляем меню на форму
            Controls.Add(menuStrip);
            
            // Настраиваем левую панель (список групп)
            var leftPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            
            var groupsLabel = new Label
            {
                Text = "Группы соответствий:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font(Font, FontStyle.Bold)
            };
            
            _groupsListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.One,
                Font = new Font("Segoe UI", 10)
            };
            _groupsListBox.SelectedIndexChanged += GroupsListBox_SelectedIndexChanged;
            
            var groupButtonsPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 90
            };
            
            _groupNameTextBox = new TextBox
            {
                Width = 200,
                Location = new Point(10, 10)
            };
            
            _addGroupButton = new Button
            {
                Text = "Добавить",
                Width = 80,
                Location = new Point(10, 40),
                BackColor = Color.LightGreen
            };
            _addGroupButton.Click += (s, e) => AddNewGroup();
            
            _removeGroupButton = new Button
            {
                Text = "Удалить",
                Width = 80,
                Location = new Point(100, 40),
                BackColor = Color.LightPink
            };
            _removeGroupButton.Click += (s, e) => RemoveSelectedGroup();
            
            _renameGroupButton = new Button
            {
                Text = "Переименовать",
                Width = 110,
                Location = new Point(10, 70)
            };
            _renameGroupButton.Click += (s, e) => RenameSelectedGroup();
            
            groupButtonsPanel.Controls.Add(_groupNameTextBox);
            groupButtonsPanel.Controls.Add(_addGroupButton);
            groupButtonsPanel.Controls.Add(_removeGroupButton);
            groupButtonsPanel.Controls.Add(_renameGroupButton);
            
            leftPanel.Controls.Add(_groupsListBox);
            leftPanel.Controls.Add(groupsLabel);
            leftPanel.Controls.Add(groupButtonsPanel);
            
            // Настраиваем правую панель (редактирование группы)
            var rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            
            var rightSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 400,
                Panel1MinSize = 200,
                Panel2MinSize = 150
            };
            
            // Верхняя часть правой панели (список вариаций)
            var variationsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };
            
            var variationsLabel = new Label
            {
                Text = "Вариации наименований и артикулов:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font(Font, FontStyle.Bold)
            };
            
            _variationsListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.One,
                Font = new Font("Segoe UI", 10)
            };
            
            // Добавляем обработчик события выбора вариации
            _variationsListBox.SelectedIndexChanged += (s, e) => UpdateControlsState();
            
            var variationButtonsPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 100
            };
            
            // Радиокнопки для выбора типа вариации
            _nameRadioButton = new RadioButton
            {
                Text = "Наименование",
                Location = new Point(10, 10),
                Checked = true
            };
            
            _articleRadioButton = new RadioButton
            {
                Text = "Артикул",
                Location = new Point(150, 10)
            };
            
            // Поля ввода для наименования и артикула
            _nameVariationTextBox = new TextBox
            {
                Width = 280,
                Location = new Point(10, 35),
                PlaceholderText = "Введите вариацию наименования"
            };
            
            _articleVariationTextBox = new TextBox
            {
                Width = 280,
                Location = new Point(10, 35),
                PlaceholderText = "Введите вариацию артикула",
                Visible = false
            };
            
            // Переключение между полями ввода
            _nameRadioButton.CheckedChanged += (s, e) => 
            {
                _nameVariationTextBox.Visible = _nameRadioButton.Checked;
                _articleVariationTextBox.Visible = !_nameRadioButton.Checked;
            };
            
            _articleRadioButton.CheckedChanged += (s, e) =>
            {
                _articleVariationTextBox.Visible = _articleRadioButton.Checked;
                _nameVariationTextBox.Visible = !_articleRadioButton.Checked;
            };
            
            _addVariationButton = new Button
            {
                Text = "Добавить вариацию",
                Width = 140,
                Location = new Point(10, 65),
                BackColor = Color.LightGreen
            };
            _addVariationButton.Click += (s, e) => AddVariation();
            
            _removeVariationButton = new Button
            {
                Text = "Удалить вариацию",
                Width = 140,
                Location = new Point(160, 65),
                BackColor = Color.LightPink
            };
            _removeVariationButton.Click += (s, e) => RemoveSelectedVariation();
            
            variationButtonsPanel.Controls.Add(_nameRadioButton);
            variationButtonsPanel.Controls.Add(_articleRadioButton);
            variationButtonsPanel.Controls.Add(_nameVariationTextBox);
            variationButtonsPanel.Controls.Add(_articleVariationTextBox);
            variationButtonsPanel.Controls.Add(_addVariationButton);
            variationButtonsPanel.Controls.Add(_removeVariationButton);
            
            variationsPanel.Controls.Add(_variationsListBox);
            variationsPanel.Controls.Add(variationsLabel);
            variationsPanel.Controls.Add(variationButtonsPanel);
            
            // Нижняя часть правой панели (унифицированные значения)
            var unifiedPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };
            
            var unifiedLabel = new Label
            {
                Text = "Унифицированные значения:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font(Font, FontStyle.Bold)
            };
            
            var unifiedTablePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(5)
            };
            
            unifiedTablePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            unifiedTablePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            unifiedTablePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            unifiedTablePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            
            _unifiedArticleLabel = new Label
            {
                Text = "Унифицированный артикул:",
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            
            _unifiedArticleTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10)
            };
            
            _primaryNameLabel = new Label
            {
                Text = "Основное наименование:",
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            
            _primaryNameTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10)
            };
            
            unifiedTablePanel.Controls.Add(_unifiedArticleLabel, 0, 0);
            unifiedTablePanel.Controls.Add(_unifiedArticleTextBox, 1, 0);
            unifiedTablePanel.Controls.Add(_primaryNameLabel, 0, 1);
            unifiedTablePanel.Controls.Add(_primaryNameTextBox, 1, 1);
            
            _applyUnifiedButton = new Button
            {
                Text = "Применить унифицированные значения",
                Dock = DockStyle.Bottom,
                Height = 30,
                BackColor = Color.LightBlue
            };
            _applyUnifiedButton.Click += (s, e) => ApplyUnifiedValues();
            
            unifiedPanel.Controls.Add(unifiedTablePanel);
            unifiedPanel.Controls.Add(unifiedLabel);
            unifiedPanel.Controls.Add(_applyUnifiedButton);
            
            // Добавляем панели в правый сплиттер
            rightSplitContainer.Panel1.Controls.Add(variationsPanel);
            rightSplitContainer.Panel2.Controls.Add(unifiedPanel);
            
            rightPanel.Controls.Add(rightSplitContainer);
            
            // Добавляем панели в главный сплиттер
            _mainSplitContainer.Panel1.Controls.Add(leftPanel);
            _mainSplitContainer.Panel2.Controls.Add(rightPanel);
            
            // Создаем панель кнопок внизу формы
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };
            
            _saveButton = new Button
            {
                Text = "Сохранить",
                DialogResult = DialogResult.OK,
                Width = 100,
                Height = 30,
                Location = new Point(Width - 230, 10),
                BackColor = Color.LightGreen
            };
            _saveButton.Click += (s, e) => SaveMappingData();
            
            _cancelButton = new Button
            {
                Text = "Отмена",
                DialogResult = DialogResult.Cancel,
                Width = 100,
                Height = 30,
                Location = new Point(Width - 120, 10)
            };
            
            bottomPanel.Controls.Add(_saveButton);
            bottomPanel.Controls.Add(_cancelButton);
            
            // Добавляем компоненты на форму
            Controls.Add(_mainSplitContainer);
            Controls.Add(bottomPanel);
            
            // При загрузке формы настроим разделитель
            this.Shown += (s, e) => 
            {
                try 
                {
                    // Изменяем ориентацию на вертикальную и устанавливаем расстояние
                    _mainSplitContainer.Orientation = Orientation.Vertical;
                    _mainSplitContainer.SplitterDistance = 250;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при настройке разделителя: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            
            // Устанавливаем кнопки по умолчанию
            AcceptButton = _saveButton;
            CancelButton = _cancelButton;
            
            // Настраиваем начальное состояние элементов управления
            UpdateControlsState();
        }
        
        /// <summary>
        /// Загрузка данных соответствий из файла
        /// </summary>
        private void LoadMappingData()
        {
            try
            {
                if (File.Exists(_mappingFilePath))
                {
                    string json = File.ReadAllText(_mappingFilePath);
                    
                    // Пробуем сначала загрузить в новом формате (MappingDatabase)
                    try
                    {
                        _mappingDatabase = JsonSerializer.Deserialize<MappingDatabase>(json) ?? new MappingDatabase();
                    }
                    catch
                    {
                        // Если не удалось, пробуем загрузить в старом формате (список UnifiedProduct)
                        var unifiedProducts = JsonSerializer.Deserialize<List<UnifiedProduct>>(json) ?? new List<UnifiedProduct>();
                        
                        // Конвертируем в новый формат
                        _mappingDatabase = new MappingDatabase();
                        foreach (var product in unifiedProducts)
                        {
                            var group = new MappingGroup
                            {
                                Name = product.PrimaryName,
                                UnifiedArticle = product.UnifiedArticle,
                                PrimaryName = product.PrimaryName,
                                NameVariations = product.NameVariations,
                                ArticleVariations = product.ArticleVariations
                            };
                            
                            _mappingDatabase.Groups.Add(group);
                        }
                    }
                }
                
                // Обновляем список групп
                UpdateGroupsList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при загрузке файла соответствий: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Сохранение данных соответствий в файл
        /// </summary>
        private void SaveMappingData()
        {
            try
            {
                // Сохраняем текущие изменения в выбранной группе
                SaveCurrentGroupChanges();
                
                // Сериализуем базу данных соответствий в JSON
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_mappingDatabase, options);
                
                // Записываем в файл
                File.WriteAllText(_mappingFilePath, json);
                
                MessageBox.Show(
                    "Настройки соответствий успешно сохранены.",
                    "Информация",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при сохранении файла соответствий: {ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Обновление списка групп в интерфейсе
        /// </summary>
        private void UpdateGroupsList()
        {
            _groupsListBox.Items.Clear();
            
            foreach (var group in _mappingDatabase.Groups)
            {
                _groupsListBox.Items.Add(group.Name);
            }
            
            // Выбираем первую группу, если она есть
            if (_groupsListBox.Items.Count > 0)
            {
                _groupsListBox.SelectedIndex = 0;
            }
            else
            {
                // Если групп нет, очищаем интерфейс
                _selectedGroup = null;
                UpdateVariationsList();
                UpdateUnifiedValues();
            }
            
            UpdateControlsState();
        }
        
        /// <summary>
        /// Обновление списка вариаций в интерфейсе
        /// </summary>
        private void UpdateVariationsList()
        {
            _variationsListBox.Items.Clear();
            
            if (_selectedGroup != null)
            {
                // Добавляем вариации наименований
                foreach (var name in _selectedGroup.NameVariations)
                {
                    _variationsListBox.Items.Add($"Наименование: {name}");
                }
                
                // Добавляем вариации артикулов
                foreach (var article in _selectedGroup.ArticleVariations)
                {
                    _variationsListBox.Items.Add($"Артикул: {article}");
                }
            }
            
            UpdateControlsState();
        }
        
        /// <summary>
        /// Обновление унифицированных значений в интерфейсе
        /// </summary>
        private void UpdateUnifiedValues()
        {
            if (_selectedGroup != null)
            {
                _unifiedArticleTextBox.Text = _selectedGroup.UnifiedArticle;
                _primaryNameTextBox.Text = _selectedGroup.PrimaryName;
            }
            else
            {
                _unifiedArticleTextBox.Text = string.Empty;
                _primaryNameTextBox.Text = string.Empty;
            }
            
            UpdateControlsState();
        }
        
        /// <summary>
        /// Обновление состояния элементов управления
        /// </summary>
        private void UpdateControlsState()
        {
            bool hasGroups = _groupsListBox.Items.Count > 0;
            bool hasSelectedGroup = _selectedGroup != null;
            bool hasSelectedVariation = _variationsListBox.SelectedIndex >= 0;
            
            // Кнопки управления группами
            _removeGroupButton.Enabled = hasSelectedGroup;
            _renameGroupButton.Enabled = hasSelectedGroup;
            
            // Кнопки управления вариациями
            _addVariationButton.Enabled = hasSelectedGroup;
            _removeVariationButton.Enabled = hasSelectedGroup && hasSelectedVariation;
            _nameVariationTextBox.Enabled = hasSelectedGroup;
            _articleVariationTextBox.Enabled = hasSelectedGroup;
            _nameRadioButton.Enabled = hasSelectedGroup;
            _articleRadioButton.Enabled = hasSelectedGroup;
            
            // Поля унифицированных значений
            _unifiedArticleTextBox.Enabled = hasSelectedGroup;
            _primaryNameTextBox.Enabled = hasSelectedGroup;
            _applyUnifiedButton.Enabled = hasSelectedGroup;
            
            // Правая панель в целом
            _mainSplitContainer.Panel2.Enabled = hasSelectedGroup;
        }
        
        /// <summary>
        /// Сохранение изменений текущей выбранной группы
        /// </summary>
        private void SaveCurrentGroupChanges()
        {
            if (_selectedGroup != null)
            {
                _selectedGroup.UnifiedArticle = _unifiedArticleTextBox.Text;
                _selectedGroup.PrimaryName = _primaryNameTextBox.Text;
            }
        }
        
        /// <summary>
        /// Добавление новой группы соответствий
        /// </summary>
        private void AddNewGroup()
        {
            string groupName = _groupNameTextBox.Text.Trim();
            
            if (string.IsNullOrEmpty(groupName))
            {
                MessageBox.Show(
                    "Введите название группы.",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            // Проверяем, что группа с таким названием не существует
            if (_mappingDatabase.Groups.Any(g => g.Name.Equals(groupName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    "Группа с таким названием уже существует.",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            // Сохраняем изменения текущей группы
            SaveCurrentGroupChanges();
            
            // Создаем новую группу
            var newGroup = new MappingGroup
            {
                Name = groupName,
                UnifiedArticle = "",
                PrimaryName = groupName
            };
            
            // Добавляем в базу данных
            _mappingDatabase.Groups.Add(newGroup);
            
            // Обновляем список групп
            UpdateGroupsList();
            
            // Выбираем новую группу
            int index = _groupsListBox.Items.IndexOf(groupName);
            if (index >= 0)
            {
                _groupsListBox.SelectedIndex = index;
            }
            
            // Очищаем поле ввода
            _groupNameTextBox.Text = "";
        }
        
        /// <summary>
        /// Удаление выбранной группы соответствий
        /// </summary>
        private void RemoveSelectedGroup()
        {
            if (_selectedGroup == null)
                return;
            
            // Запрашиваем подтверждение
            var result = MessageBox.Show(
                $"Вы действительно хотите удалить группу \"{_selectedGroup.Name}\"?",
                "Подтверждение удаления",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                // Удаляем группу из базы данных
                _mappingDatabase.Groups.Remove(_selectedGroup);
                
                // Обновляем список групп
                UpdateGroupsList();
            }
        }
        
        /// <summary>
        /// Переименование выбранной группы соответствий
        /// </summary>
        private void RenameSelectedGroup()
        {
            if (_selectedGroup == null)
                return;
            
            string newName = _groupNameTextBox.Text.Trim();
            
            if (string.IsNullOrEmpty(newName))
            {
                MessageBox.Show(
                    "Введите новое название группы.",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            // Проверяем, что группа с таким названием не существует
            if (_mappingDatabase.Groups.Any(g => g != _selectedGroup && g.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    "Группа с таким названием уже существует.",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            // Запоминаем индекс выбранной группы
            int selectedIndex = _groupsListBox.SelectedIndex;
            
            // Переименовываем группу
            _selectedGroup.Name = newName;
            
            // Обновляем список групп
            UpdateGroupsList();
            
            // Восстанавливаем выбор
            if (selectedIndex >= 0 && selectedIndex < _groupsListBox.Items.Count)
            {
                _groupsListBox.SelectedIndex = selectedIndex;
            }
            
            // Очищаем поле ввода
            _groupNameTextBox.Text = "";
        }
        
        /// <summary>
        /// Добавление новой вариации в выбранную группу
        /// </summary>
        private void AddVariation()
        {
            if (_selectedGroup == null)
                return;
            
            // Определяем, какой тип вариации добавляем
            bool isArticle = _articleRadioButton.Checked;
            string variation = isArticle ? _articleVariationTextBox.Text.Trim() : _nameVariationTextBox.Text.Trim();
            
            if (string.IsNullOrEmpty(variation))
            {
                MessageBox.Show(
                    $"Введите текст вариации {(isArticle ? "артикула" : "наименования")}.",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            // Проверяем, что такой вариации еще нет
            var existingList = isArticle ? _selectedGroup.ArticleVariations : _selectedGroup.NameVariations;
            
            if (existingList.Any(v => v.Equals(variation, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    $"Такая вариация {(isArticle ? "артикула" : "наименования")} уже существует в этой группе.",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            
            // Добавляем вариацию
            existingList.Add(variation);
            
            // Обновляем список вариаций
            UpdateVariationsList();
            
            // Очищаем поле ввода
            if (isArticle)
                _articleVariationTextBox.Text = "";
            else
                _nameVariationTextBox.Text = "";
        }
        
        /// <summary>
        /// Удаление выбранной вариации из группы
        /// </summary>
        private void RemoveSelectedVariation()
        {
            try
            {
                // Проверяем, что есть выбранная группа и выбранная вариация
                if (_selectedGroup == null)
                {
                    MessageBox.Show("Не выбрана группа соответствий", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (_variationsListBox.SelectedIndex < 0)
                {
                    MessageBox.Show("Не выбрана вариация для удаления", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Получаем выбранную вариацию
                object selectedObj = _variationsListBox.SelectedItem;
                if (selectedObj == null)
                {
                    MessageBox.Show("Ошибка при получении выбранной вариации", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                string selectedItem = selectedObj.ToString();
                
                // Определяем тип вариации и извлекаем значение
                bool isArticle = selectedItem.StartsWith("Артикул:");
                string value = selectedItem.Substring(isArticle ? 9 : 13).Trim();
                
                // Удаляем вариацию из соответствующего списка
                bool removed = false;
                if (isArticle)
                {
                    removed = _selectedGroup.ArticleVariations.Remove(value);
                }
                else
                {
                    removed = _selectedGroup.NameVariations.Remove(value);
                }
                
                // Обновляем список вариаций
                UpdateVariationsList();
                
                // Сообщаем об успешном удалении
                if (removed)
                {
                    MessageBox.Show($"Вариация {(isArticle ? "артикула" : "наименования")} \"{value}\" успешно удалена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении вариации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Применение унифицированных значений
        /// </summary>
        private void ApplyUnifiedValues()
        {
            if (_selectedGroup == null)
                return;
            
            _selectedGroup.UnifiedArticle = _unifiedArticleTextBox.Text.Trim();
            _selectedGroup.PrimaryName = _primaryNameTextBox.Text.Trim();
            
            MessageBox.Show(
                "Унифицированные значения применены.",
                "Информация",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// Обработчик события изменения выбранной группы
        /// </summary>
        private void GroupsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Сохраняем изменения текущей группы
            SaveCurrentGroupChanges();
            
            // Получаем выбранную группу
            int selectedIndex = _groupsListBox.SelectedIndex;
            
            if (selectedIndex >= 0 && selectedIndex < _mappingDatabase.Groups.Count)
            {
                _selectedGroup = _mappingDatabase.Groups[selectedIndex];
            }
            else
            {
                _selectedGroup = null;
            }
            
            // Обновляем интерфейс
            UpdateVariationsList();
            UpdateUnifiedValues();
        }
        
        /// <summary>
        /// Определение является ли строка артикулом
        /// </summary>
        private bool IsArticleNumber(string text)
        {
            // Простая эвристика: если строка содержит цифры и короче 20 символов, считаем её артикулом
            return text.Any(char.IsDigit) && text.Length < 20;
        }
    }
}
