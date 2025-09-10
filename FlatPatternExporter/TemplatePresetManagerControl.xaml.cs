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
            TemplatePresetsListBox.Visibility = isEmpty ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    private void TemplatePresetsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PresetManager?.SelectedTemplatePreset != null && 
            !string.IsNullOrEmpty(PresetManager.SelectedTemplatePreset.Template) &&
            TokenService != null)
        {
            TokenService.FileNameTemplate = PresetManager.SelectedTemplatePreset.Template;
            PresetNameTextBox.Text = PresetManager.SelectedTemplatePreset.Name;
            ClearNameValidation();
        }
    }

    private void DeletePresetButton_Click(object sender, RoutedEventArgs e)
    {
        if (PresetManager?.DeleteSelectedPreset() == true)
        {
            UpdateEmptyState();
        }
    }

    

    private void CreatePresetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!ValidateAndGetTemplate(out var template)) return;

        var presetName = GetTrimmedPresetName();
        if (!PresetManager.CreatePreset(presetName, template, out var error))
        {
            ShowNameValidation(error);
            return;
        }

        RefreshUI();
    }

    private void SaveChangesButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!ValidateAndGetTemplate(out var template)) return;
        if (!HasSelectedPreset()) return;

        if (PresetManager.SelectedTemplatePreset!.Template == template)
        {
            ShowMessage("Изменений не обнаружено.", MessageBoxImage.Information);
            return;
        }

        if (PresetManager.UpdateSelectedTemplate(template, out var error))
        {
            ClearNameValidation();
        }
        else if (!string.IsNullOrEmpty(error))
        {
            ShowMessage(error);
        }
    }

    private bool ValidateAndGetTemplate(out string template)
    {
        template = "";
        
        if (!IsServicesReady()) 
            return false;

        template = TokenService!.FileNameTemplate;
        if (string.IsNullOrWhiteSpace(template))
        {
            ShowMessage("Шаблон не может быть пустым.");
            return false;
        }

        if (!TokenService.IsFileNameTemplateValid)
        {
            ShowMessage("Текущий шаблон содержит ошибки. Исправьте шаблон перед операцией.");
            return false;
        }

        return true;
    }

    private void ShowMessage(string message, MessageBoxImage icon = MessageBoxImage.Warning)
    {
        MessageBox.Show(message, "Ошибка", MessageBoxButton.OK, icon);
    }

    private void RefreshUI()
    {
        ClearNameValidation();
        UpdateEmptyState();
    }

    private void RenamePresetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;

        var newName = GetTrimmedPresetName();
        if (!PresetManager!.RenameSelected(newName, out var error))
        {
            ShowNameValidation(error);
            return;
        }

        ClearNameValidation();
    }

    private void DuplicatePresetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;

        if (!PresetManager!.DuplicateSelected(out var error) && !string.IsNullOrEmpty(error))
        {
            ShowMessage(error);
            return;
        }

        PresetNameTextBox.Text = PresetManager.SelectedTemplatePreset?.Name ?? "";
        RefreshUI();
    }

    private bool HasSelectedPreset() => PresetManager?.SelectedTemplatePreset != null;
    
    private bool IsServicesReady() => TokenService != null && PresetManager != null;
    
    private string GetTrimmedPresetName() => PresetNameTextBox.Text.Trim();

    private void ShowNameValidation(string? message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            ClearNameValidation();
            return;
        }

        NameValidationTextBlock.Text = message;
        NameValidationTextBlock.Visibility = Visibility.Visible;
    }

    private void ClearNameValidation()
    {
        NameValidationTextBlock.Text = string.Empty;
        NameValidationTextBlock.Visibility = Visibility.Collapsed;
    }
}
