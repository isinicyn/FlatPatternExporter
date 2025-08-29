using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using Clipboard = System.Windows.Clipboard;

namespace FlatPatternExporter
{
    public partial class AboutWindow : Window
    {
        private readonly FlatPatternExporterMainWindow _mainWindow; // Ссылка на главное окно

        // Конструктор, принимающий ссылку на главное окно
        public AboutWindow(FlatPatternExporterMainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow; // Сохраняем ссылку на главное окно
            SetVersion();
            SetLastUpdateDate();
        }

        // Установка версии программы из главного окна
        private void SetVersion()
        {
            VersionTextBlock.Text = "Версия программы: " + _mainWindow.GetVersion();
        }

        // Установка даты последнего обновления согласно коммиту
        private void SetLastUpdateDate()
        {
            LastUpdateTextBlock.Text = "Последнее обновление: " + GetLastCommitDate();
        }

        // Метод для получения даты последнего коммита
        private string GetLastCommitDate()
        {
            try
            {
                // Получаем путь к директории, где находится исполняемый файл
                string? executingDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                // Создаем новый процесс
                using (var process = new Process())
                {
                    // Настраиваем процесс для выполнения команды Git
                    process.StartInfo.FileName = "git";
                    process.StartInfo.Arguments = "log -1 --format=%cd --date=format:\"%d.%m.%Y %H:%M:%S\"";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.WorkingDirectory = executingDir ?? AppDomain.CurrentDomain.BaseDirectory;

                    // Запускаем процесс
                    process.Start();

                    // Читаем вывод команды
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    // Проверяем, что вывод не пустой
                    if (!string.IsNullOrWhiteSpace(output))
                    {
                        return output.Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                // Логируем исключение (можно заменить на ваш способ логирования)
                Debug.WriteLine($"Ошибка при получении даты последнего коммита: {ex.Message}");
            }

            // Возвращаем заглушку, если не удалось получить дату
            return "Дата не определена";
        }        

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CopyVersionToClipboard_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(VersionTextBlock.Text);
        }
    }
}