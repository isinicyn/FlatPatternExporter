using System.Diagnostics;
using System.Reflection;
using System.Windows;

namespace FPECORE
{
    public partial class AboutWindow
    {
        public AboutWindow()
        {
            InitializeComponent();
            SetVersion();
            SetLastUpdateDate();
        }

        // Установка версии программы из кода
        private void SetVersion()
        {
            VersionTextBlock.Text = "Версия программы: " + GetVersion();
        }

        // Установка даты последнего обновления согласно коммиту
        private void SetLastUpdateDate()
        {
            // Получение даты последнего коммита из сборки (если у вас есть метод, который достает информацию о коммите)
            LastUpdateTextBlock.Text = "Последнее обновление: " + GetLastCommitDate();
        }

        // Этот метод может быть уже у вас есть, для получения версии программы
        private string GetVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var informationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
            var assemblyVersion = assembly.GetName().Version?.ToString();

            if (informationalVersion != null && informationalVersion.Contains("+"))
            {
                var parts = informationalVersion.Split('+');
                var shortCommitHash = parts[1].Length > 7 ? parts[1].Substring(0, 7) : parts[1];
                informationalVersion = $"{parts[0]}+{shortCommitHash}";
            }

            return informationalVersion ?? fileVersion ?? assemblyVersion ?? "Версия неизвестна";
        }

        // Метод для получения даты последнего коммита
        private string GetLastCommitDate()
        {
            try
            {
                // Настройка процесса для выполнения команды git log
                var processInfo = new ProcessStartInfo("git", "log -1 --format=%cd")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = GetGitRepositoryPath() // Путь к директории с репозиторием
                };

                // Запуск процесса
                using (var process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        // Чтение вывода команды
                        string output = process.StandardOutput.ReadToEnd();
                        process.WaitForExit();

                        // Успешное получение даты
                        if (!string.IsNullOrEmpty(output))
                        {
                            return output.Trim(); // Возвращаем дату последнего коммита
                        }
                    }
                }

                return "Не удалось получить дату коммита";
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                Debug.WriteLine($"Ошибка при получении даты коммита: {ex.Message}");
                return "Ошибка получения даты коммита";
            }
        }

        // Метод для получения пути к директории с GIT-репозиторием
        // Убедитесь, что этот путь корректен, если репозиторий в другой папке
        private string GetGitRepositoryPath()
        {
            // Например, если GIT-репозиторий находится в корневой папке проекта
            return AppDomain.CurrentDomain.BaseDirectory;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
