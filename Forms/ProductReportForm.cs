using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Forecast.Models;

namespace Forecast.Forms
{
    /// <summary>
    /// Форма для отображения отчета по товарам
    /// </summary>
    public partial class ProductReportForm : Form
    {
        private readonly List<UnifiedProduct> _unifiedProducts;
        
        /// <summary>
        /// Конструктор формы отчета по товарам
        /// </summary>
        /// <param name="unifiedProducts">Список унифицированных товаров</param>
        public ProductReportForm(List<UnifiedProduct> unifiedProducts)
        {
            InitializeComponent();
            
            _unifiedProducts = unifiedProducts;
            
            // Настройка формы
            this.Text = "Отчет по товарам";
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
            this.Controls.Add(toolStrip);
            
            var printButton = new ToolStripButton("Печать");
            printButton.Click += PrintButton_Click;
            toolStrip.Items.Add(printButton);
            
            var exportButton = new ToolStripButton("Экспорт");
            exportButton.Click += ExportButton_Click;
            toolStrip.Items.Add(exportButton);
            
            // Создание разделенной панели
            var splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = Orientation.Vertical;
            splitContainer.SplitterDistance = 300;
            this.Controls.Add(splitContainer);
            
            // Верхняя панель - список товаров
            var productsListBox = new ListBox();
            productsListBox.Dock = DockStyle.Fill;
            productsListBox.SelectedIndexChanged += ProductsListBox_SelectedIndexChanged;
            productsListBox.Name = "productsListBox";
            splitContainer.Panel1.Controls.Add(productsListBox);
            
            // Нижняя панель - детальная информация о товаре
            var tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            tabControl.Name = "tabControl";
            splitContainer.Panel2.Controls.Add(tabControl);
            
            // Вкладка "Общая информация"
            var generalTabPage = new TabPage("Общая информация");
            tabControl.TabPages.Add(generalTabPage);
            
            var generalPanel = new Panel();
            generalPanel.Dock = DockStyle.Fill;
            generalPanel.AutoScroll = true;
            generalPanel.Name = "generalPanel";
            generalTabPage.Controls.Add(generalPanel);
            
            // Вкладка "История заказов"
            var historyTabPage = new TabPage("История заказов");
            tabControl.TabPages.Add(historyTabPage);
            
            var historyGridView = new DataGridView();
            historyGridView.Dock = DockStyle.Fill;
            historyGridView.AllowUserToAddRows = false;
            historyGridView.ReadOnly = true;
            historyGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            historyGridView.Name = "historyGridView";
            historyTabPage.Controls.Add(historyGridView);
            
            // Вкладка "Сезонность"
            var seasonalityTabPage = new TabPage("Сезонность");
            tabControl.TabPages.Add(seasonalityTabPage);
            
            var seasonalityGridView = new DataGridView();
            seasonalityGridView.Dock = DockStyle.Fill;
            seasonalityGridView.AllowUserToAddRows = false;
            seasonalityGridView.ReadOnly = true;
            seasonalityGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            seasonalityGridView.Name = "seasonalityGridView";
            seasonalityTabPage.Controls.Add(seasonalityGridView);
            
            // Вкладка "Прогноз"
            var forecastTabPage = new TabPage("Прогноз");
            tabControl.TabPages.Add(forecastTabPage);
            
            var forecastPanel = new Panel();
            forecastPanel.Dock = DockStyle.Fill;
            forecastPanel.AutoScroll = true;
            forecastPanel.Name = "forecastPanel";
            forecastTabPage.Controls.Add(forecastPanel);
        }
        
        /// <summary>
        /// Загрузка данных
        /// </summary>
        private void LoadData()
        {
            var productsListBox = this.Controls.Find("productsListBox", true).FirstOrDefault() as ListBox;
            
            if (productsListBox != null && _unifiedProducts != null)
            {
                // Заполнение списка товаров
                productsListBox.Items.Clear();
                
                foreach (var product in _unifiedProducts.OrderBy(p => p.PrimaryName))
                {
                    productsListBox.Items.Add($"{product.PrimaryName} ({product.UnifiedArticle})");
                }
                
                // Выбор первого товара в списке
                if (productsListBox.Items.Count > 0)
                {
                    productsListBox.SelectedIndex = 0;
                }
            }
        }
        
        /// <summary>
        /// Отображение информации о выбранном товаре
        /// </summary>
        /// <param name="product">Выбранный товар</param>
        private void ShowProductInfo(UnifiedProduct product)
        {
            // Отображение общей информации о товаре
            ShowGeneralInfo(product);
            
            // Отображение истории заказов
            ShowOrderHistory(product);
            
            // Отображение сезонности
            ShowSeasonality(product);
            
            // Отображение прогноза
            ShowForecast(product);
        }
        
        /// <summary>
        /// Отображение общей информации о товаре
        /// </summary>
        /// <param name="product">Выбранный товар</param>
        private void ShowGeneralInfo(UnifiedProduct product)
        {
            var generalPanel = this.Controls.Find("generalPanel", true).FirstOrDefault() as Panel;
            
            if (generalPanel != null)
            {
                // Очистка панели
                generalPanel.Controls.Clear();
                
                // Создание таблицы для отображения информации
                var tableLayoutPanel = new TableLayoutPanel();
                tableLayoutPanel.Dock = DockStyle.Top;
                tableLayoutPanel.AutoSize = true;
                tableLayoutPanel.ColumnCount = 2;
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
                tableLayoutPanel.RowCount = 10;
                
                for (int i = 0; i < 10; i++)
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                }
                
                // Заполнение таблицы информацией о товаре
                int row = 0;
                
                AddTableRow(tableLayoutPanel, "Унифицированный артикул:", product.UnifiedArticle, row++);
                AddTableRow(tableLayoutPanel, "Наименование:", product.PrimaryName, row++);
                AddTableRow(tableLayoutPanel, "Количество заказов:", product.OrderHistory.Count.ToString(), row++);
                AddTableRow(tableLayoutPanel, "Средний интервал между заказами:", $"{Math.Round(product.AverageOrderInterval, 1)} дней", row++);
                AddTableRow(tableLayoutPanel, "Среднее количество в заказе:", Math.Round(product.AverageOrderQuantity, 1).ToString(), row++);
                AddTableRow(tableLayoutPanel, "Средний срок поставки:", $"{Math.Round(product.AverageDeliveryTime, 1)} дней", row++);
                AddTableRow(tableLayoutPanel, "Дата последнего заказа:", product.LastOrderDate.HasValue ? product.LastOrderDate.Value.ToShortDateString() : "Нет данных", row++);
                
                // Добавление вариаций наименований
                var nameVariationsLabel = new Label();
                nameVariationsLabel.Text = "Вариации наименований:";
                nameVariationsLabel.Dock = DockStyle.Fill;
                nameVariationsLabel.TextAlign = ContentAlignment.MiddleLeft;
                nameVariationsLabel.Font = new Font(nameVariationsLabel.Font, FontStyle.Bold);
                tableLayoutPanel.Controls.Add(nameVariationsLabel, 0, row);
                
                var nameVariationsTextBox = new TextBox();
                nameVariationsTextBox.Multiline = true;
                nameVariationsTextBox.ReadOnly = true;
                nameVariationsTextBox.ScrollBars = ScrollBars.Vertical;
                nameVariationsTextBox.Dock = DockStyle.Fill;
                nameVariationsTextBox.Height = 60;
                nameVariationsTextBox.Text = string.Join(Environment.NewLine, product.NameVariations);
                tableLayoutPanel.Controls.Add(nameVariationsTextBox, 1, row++);
                
                // Добавление вариаций артикулов
                var articleVariationsLabel = new Label();
                articleVariationsLabel.Text = "Вариации артикулов:";
                articleVariationsLabel.Dock = DockStyle.Fill;
                articleVariationsLabel.TextAlign = ContentAlignment.MiddleLeft;
                articleVariationsLabel.Font = new Font(articleVariationsLabel.Font, FontStyle.Bold);
                tableLayoutPanel.Controls.Add(articleVariationsLabel, 0, row);
                
                var articleVariationsTextBox = new TextBox();
                articleVariationsTextBox.Multiline = true;
                articleVariationsTextBox.ReadOnly = true;
                articleVariationsTextBox.ScrollBars = ScrollBars.Vertical;
                articleVariationsTextBox.Dock = DockStyle.Fill;
                articleVariationsTextBox.Height = 60;
                articleVariationsTextBox.Text = string.Join(Environment.NewLine, product.ArticleVariations);
                tableLayoutPanel.Controls.Add(articleVariationsTextBox, 1, row++);
                
                generalPanel.Controls.Add(tableLayoutPanel);
            }
        }
        
        /// <summary>
        /// Добавление строки в таблицу
        /// </summary>
        private void AddTableRow(TableLayoutPanel table, string labelText, string valueText, int row)
        {
            var label = new Label();
            label.Text = labelText;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Font = new Font(label.Font, FontStyle.Bold);
            table.Controls.Add(label, 0, row);
            
            var valueLabel = new Label();
            valueLabel.Text = valueText;
            valueLabel.Dock = DockStyle.Fill;
            valueLabel.TextAlign = ContentAlignment.MiddleLeft;
            table.Controls.Add(valueLabel, 1, row);
        }
        
        /// <summary>
        /// Отображение истории заказов
        /// </summary>
        /// <param name="product">Выбранный товар</param>
        private void ShowOrderHistory(UnifiedProduct product)
        {
            var historyGridView = this.Controls.Find("historyGridView", true).FirstOrDefault() as DataGridView;
            
            if (historyGridView != null)
            {
                // Очистка таблицы
                historyGridView.DataSource = null;
                
                // Создание источника данных для таблицы
                var dataSource = product.OrderHistory
                    .OrderByDescending(item => item.OrderDate)
                    .Select(item => new
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
                
                historyGridView.DataSource = dataSource;
            }
        }
        
        /// <summary>
        /// Отображение сезонности
        /// </summary>
        /// <param name="product">Выбранный товар</param>
        private void ShowSeasonality(UnifiedProduct product)
        {
            var seasonalityGridView = this.Controls.Find("seasonalityGridView", true).FirstOrDefault() as DataGridView;
            
            if (seasonalityGridView != null)
            {
                // Очистка таблицы
                seasonalityGridView.DataSource = null;
                
                // Создание источника данных для таблицы
                var dataSource = new List<object>();
                
                string[] months = new string[] 
                { 
                    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", 
                    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" 
                };
                
                for (int i = 0; i < 12; i++)
                {
                    dataSource.Add(new
                    {
                        Месяц = months[i],
                        Коэффициент_сезонности = Math.Round(product.SeasonalityCoefficients[i], 2)
                    });
                }
                
                seasonalityGridView.DataSource = dataSource;
                
                // Настройка цветов для коэффициентов сезонности
                seasonalityGridView.CellFormatting += (sender, e) =>
                {
                    if (e.ColumnIndex == 1 && e.RowIndex >= 0) // Столбец "Коэффициент_сезонности"
                    {
                        double coefficient = Convert.ToDouble(e.Value);
                        
                        if (coefficient > 1.2) // Высокая сезонность (более 20% выше среднего)
                        {
                            e.CellStyle.BackColor = Color.LightGreen;
                        }
                        else if (coefficient < 0.8) // Низкая сезонность (более 20% ниже среднего)
                        {
                            e.CellStyle.BackColor = Color.LightSalmon;
                        }
                        else // Нормальная сезонность
                        {
                            e.CellStyle.BackColor = Color.LightYellow;
                        }
                    }
                };
            }
        }
        
        /// <summary>
        /// Отображение прогноза
        /// </summary>
        /// <param name="product">Выбранный товар</param>
        private void ShowForecast(UnifiedProduct product)
        {
            var forecastPanel = this.Controls.Find("forecastPanel", true).FirstOrDefault() as Panel;
            
            if (forecastPanel != null)
            {
                // Очистка панели
                forecastPanel.Controls.Clear();
                
                // Создание таблицы для отображения информации
                var tableLayoutPanel = new TableLayoutPanel();
                tableLayoutPanel.Dock = DockStyle.Top;
                tableLayoutPanel.AutoSize = true;
                tableLayoutPanel.ColumnCount = 2;
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
                tableLayoutPanel.RowCount = 5;
                
                for (int i = 0; i < 5; i++)
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                }
                
                // Заполнение таблицы информацией о прогнозе
                int row = 0;
                
                // Прогнозируемая дата следующего заказа
                string nextOrderDate = product.NextPredictedOrderDate.HasValue 
                    ? product.NextPredictedOrderDate.Value.ToShortDateString() 
                    : "Нет данных";
                
                AddTableRow(tableLayoutPanel, "Прогнозируемая дата следующего заказа:", nextOrderDate, row++);
                
                // Рекомендуемое количество
                string recommendedQuantity = product.RecommendedQuantity > 0 
                    ? Math.Round(product.RecommendedQuantity, 0).ToString() 
                    : "Нет данных";
                
                AddTableRow(tableLayoutPanel, "Рекомендуемое количество:", recommendedQuantity, row++);
                
                // Оптимальная дата размещения заказа
                string optimalOrderDate = product.OptimalOrderPlacementDate.HasValue 
                    ? product.OptimalOrderPlacementDate.Value.ToShortDateString() 
                    : "Нет данных";
                
                AddTableRow(tableLayoutPanel, "Оптимальная дата размещения заказа:", optimalOrderDate, row++);
                
                // Расчет приоритета заказа
                string priority = "Нет данных";
                Color priorityColor = Color.White;
                
                if (product.OptimalOrderPlacementDate.HasValue)
                {
                    double daysUntilPlacement = (product.OptimalOrderPlacementDate.Value - DateTime.Now).TotalDays;
                    
                    if (daysUntilPlacement < 0) // Дата размещения уже прошла
                    {
                        priority = "1 - Наивысший";
                        priorityColor = Color.Red;
                    }
                    else if (daysUntilPlacement < 7) // До 7 дней
                    {
                        priority = "2 - Высокий";
                        priorityColor = Color.Orange;
                    }
                    else if (daysUntilPlacement < 14) // До 14 дней
                    {
                        priority = "3 - Средний";
                        priorityColor = Color.Yellow;
                    }
                    else if (daysUntilPlacement < 21) // До 21 дня
                    {
                        priority = "4 - Низкий";
                        priorityColor = Color.LightGreen;
                    }
                    else // Более 21 дня
                    {
                        priority = "5 - Самый низкий";
                        priorityColor = Color.LightBlue;
                    }
                }
                
                var priorityLabel = new Label();
                priorityLabel.Text = "Приоритет заказа:";
                priorityLabel.Dock = DockStyle.Fill;
                priorityLabel.TextAlign = ContentAlignment.MiddleLeft;
                priorityLabel.Font = new Font(priorityLabel.Font, FontStyle.Bold);
                tableLayoutPanel.Controls.Add(priorityLabel, 0, row);
                
                var priorityValueLabel = new Label();
                priorityValueLabel.Text = priority;
                priorityValueLabel.Dock = DockStyle.Fill;
                priorityValueLabel.TextAlign = ContentAlignment.MiddleLeft;
                priorityValueLabel.BackColor = priorityColor;
                
                if (priorityColor == Color.Red)
                {
                    priorityValueLabel.ForeColor = Color.White;
                }
                
                tableLayoutPanel.Controls.Add(priorityValueLabel, 1, row++);
                
                // Комментарии к прогнозу
                var notesLabel = new Label();
                notesLabel.Text = "Комментарии к прогнозу:";
                notesLabel.Dock = DockStyle.Fill;
                notesLabel.TextAlign = ContentAlignment.MiddleLeft;
                notesLabel.Font = new Font(notesLabel.Font, FontStyle.Bold);
                tableLayoutPanel.Controls.Add(notesLabel, 0, row);
                
                var notesTextBox = new TextBox();
                notesTextBox.Multiline = true;
                notesTextBox.ReadOnly = true;
                notesTextBox.ScrollBars = ScrollBars.Vertical;
                notesTextBox.Dock = DockStyle.Fill;
                notesTextBox.Height = 100;
                
                // Формирование комментариев к прогнозу
                var notes = new List<string>();
                
                // Тренд объемов заказов
                if (product.OrderHistory.Count >= 2)
                {
                    var sortedHistory = product.OrderHistory.OrderBy(item => item.OrderDate).ToList();
                    double firstQuantity = sortedHistory.First().OrderedQuantity;
                    double lastQuantity = sortedHistory.Last().OrderedQuantity;
                    
                    // Если первое количество равно 0, используем 1 для избежания деления на 0
                    if (firstQuantity == 0)
                        firstQuantity = 1;
                    
                    double change = (lastQuantity - firstQuantity) / firstQuantity;
                    
                    if (Math.Abs(change) < 0.1) // Изменение менее 10%
                        notes.Add("Объемы заказов стабильны.");
                    else if (change > 0)
                        notes.Add($"Наблюдается рост объемов заказов на {Math.Round(change * 100)}%.");
                    else
                        notes.Add($"Наблюдается снижение объемов заказов на {Math.Round(Math.Abs(change) * 100)}%.");
                }
                else
                {
                    notes.Add("Недостаточно данных для анализа тренда объемов заказов.");
                }
                
                // Сроки поставки
                if (product.AverageDeliveryTime > 0)
                {
                    notes.Add($"Средний срок поставки: {Math.Round(product.AverageDeliveryTime)} дней. " +
                              $"Запас времени: {Math.Round(product.AverageDeliveryTime * 0.2)} дней.");
                }
                
                // Сезонность
                if (product.NextPredictedOrderDate.HasValue)
                {
                    int nextMonth = product.NextPredictedOrderDate.Value.Month - 1;
                    double seasonalCoefficient = product.SeasonalityCoefficients[nextMonth];
                    
                    if (Math.Abs(seasonalCoefficient - 1) > 0.1) // Если есть значимая сезонность
                    {
                        string seasonalityDirection = seasonalCoefficient > 1 ? "повышение" : "снижение";
                        notes.Add($"Сезонный фактор: {seasonalityDirection} на {Math.Round(Math.Abs(seasonalCoefficient - 1) * 100)}% в {product.NextPredictedOrderDate.Value.ToString("MMMM")}.");
                    }
                }
                
                // Частота заказов
                if (product.AverageOrderInterval > 0)
                {
                    notes.Add($"Средний интервал между заказами: {Math.Round(product.AverageOrderInterval)} дней.");
                }
                
                notesTextBox.Text = string.Join(Environment.NewLine, notes);
                tableLayoutPanel.Controls.Add(notesTextBox, 1, row++);
                
                forecastPanel.Controls.Add(tableLayoutPanel);
            }
        }
        
        #region Обработчики событий
        
        /// <summary>
        /// Обработчик события выбора товара в списке
        /// </summary>
        private void ProductsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var productsListBox = sender as ListBox;
            
            if (productsListBox != null && productsListBox.SelectedIndex >= 0)
            {
                // Получение выбранного товара
                string selectedItem = productsListBox.SelectedItem.ToString();
                string unifiedArticle = selectedItem.Substring(selectedItem.LastIndexOf('(') + 1, selectedItem.LastIndexOf(')') - selectedItem.LastIndexOf('(') - 1);
                
                var product = _unifiedProducts.FirstOrDefault(p => p.UnifiedArticle == unifiedArticle);
                
                if (product != null)
                {
                    // Отображение информации о выбранном товаре
                    ShowProductInfo(product);
                }
            }
        }
        
        /// <summary>
        /// Обработчик события нажатия кнопки "Печать"
        /// </summary>
        private void PrintButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Функция печати отчета по товарам будет реализована в следующей версии.", 
                "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// Обработчик события нажатия кнопки "Экспорт"
        /// </summary>
        private void ExportButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Функция экспорта отчета по товарам будет реализована в следующей версии.", 
                "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        #endregion
    }
}
