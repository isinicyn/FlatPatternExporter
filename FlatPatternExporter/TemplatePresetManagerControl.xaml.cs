using System.Windows;
using System.Windows.Controls;
using UserControl = System.Windows.Controls.UserControl;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;

public partial class TemplatePresetManagerControl : UserControl
{
    public static readonly DependencyProperty PresetManagerProperty =
        DependencyProperty.Register(nameof(PresetManager), typeof(TemplatePresetManager), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(null, OnPresetManagerChanged));

    public static readonly DependencyProperty TokenServiceProperty =
        DependencyProperty.Register(nameof(TokenService), typeof(TokenService), typeof(TemplatePresetManagerControl));

    public TemplatePresetManager PresetManager
    {
        get => (TemplatePresetManager)GetValue(PresetManagerProperty);
        set => SetValue(PresetManagerProperty, value);
    }

    public TokenService TokenService
    {
        get => (TokenService)GetValue(TokenServiceProperty);
        set => SetValue(TokenServiceProperty, value);
    }

    public TemplatePresetManagerControl()
    {
        InitializeComponent();
    }

    private static void OnPresetManagerChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is TemplatePresetManagerControl control)
        {
            control.UpdateEmptyState();
        }
    }

    private void UpdateEmptyState()
    {
        if (PresetManager != null)
        {
            bool isEmpty = PresetManager.TemplatePresets.Count == 0;
            EmptyStatePanel.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
            TemplatePresetsComboBox.Visibility = isEmpty ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    private void TemplatePresetsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PresetManager?.SelectedTemplatePreset != null && 
            !string.IsNullOrEmpty(PresetManager.SelectedTemplatePreset.Template) &&
            TokenService != null)
        {
            TokenService.FileNameTemplate = PresetManager.SelectedTemplatePreset.Template;
        }
    }

    private void DeletePresetButton_Click(object sender, RoutedEventArgs e)
    {
        if (PresetManager?.DeleteSelectedPreset() == true)
        {
            UpdateEmptyState();
        }
    }

    private void SavePresetButton_Click(object sender, RoutedEventArgs e)
    {
        if (TokenService == null || PresetManager == null)
            return;

        var currentTemplate = TokenService.FileNameTemplate;
        if (string.IsNullOrEmpty(currentTemplate))
        {
            MessageBox.Show("Шаблон не может быть пустым.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string presetName = PresetNameTextBox.Text.Trim();
        if (PresetManager.SavePreset(presetName, currentTemplate))
        {
            PresetNameTextBox.Text = "";
            UpdateEmptyState();
        }
    }
}