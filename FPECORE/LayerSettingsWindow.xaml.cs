using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;

namespace FPECORE
{
    public partial class LayerSettingsWindow : Window, INotifyPropertyChanged
    {
        private ObservableCollection<string> availableColors;
        private ObservableCollection<string> lineTypes;
        private ObservableCollection<LayerSetting> layerSettings;

        public ObservableCollection<string> AvailableColors
        {
            get => availableColors;
            set
            {
                availableColors = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<string> LineTypes
        {
            get => lineTypes;
            set
            {
                lineTypes = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<LayerSetting> LayerSettings
        {
            get => layerSettings;
            set
            {
                layerSettings = value;
                OnPropertyChanged();
            }
        }

        public LayerSettingsWindow()
        {
            InitializeComponent();
            DataContext = this;
            InitializeLayerSettings();
        }

        private void InitializeLayerSettings()
        {
            // Инициализация доступных цветов и типов линий
            AvailableColors = LayerSettingsHelper.GetAvailableColors();
            LineTypes = LayerSettingsHelper.GetLineTypes();
            
            // Инициализация настроек слоев
            LayerSettings = LayerSettingsHelper.InitializeLayerSettings();
        }

        // Метод для обработки нажатия кнопки экспорта
        private void ExportLayerOptions(object sender, RoutedEventArgs e)
        {
            string exportString = LayerSettingsHelper.ExportLayerOptions(LayerSettings);
            ExportStringTextBox.Text = exportString;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}