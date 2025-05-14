using System;
using System.Windows.Forms;
using Forecast.Forms;

namespace Forecast
{
    /// <summary>
    /// Главный класс программы
    /// </summary>
    static class Program
    {
        /// <summary>
        /// Главная точка входа в приложение
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Настройка приложения для высокого DPI и других параметров
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ApplicationConfiguration.Initialize();

            // Запуск главной формы приложения
            Application.Run(new MainForm());
        }
    }
}