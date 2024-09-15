using System.Diagnostics;
using System.Reflection;
using System.Windows;
using Clipboard = System.Windows.Clipboard;

namespace FPECORE
{
    public partial class AboutWindow : Window
    {
        private readonly MainWindow _mainWindow; // Ссылка на главное окно

        // Конструктор, принимающий ссылку на главное окно
        public AboutWindow(MainWindow mainWindow)
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
            // Ваш код для получения даты коммита
            return "Дата коммита";
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