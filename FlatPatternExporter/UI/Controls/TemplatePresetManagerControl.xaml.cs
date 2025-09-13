using System.ComponentModel;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using UserControl = System.Windows.Controls.UserControl;

namespace FlatPatternExporter;

public partial class TemplatePresetManagerControl : UserControl
{
    public static readonly DependencyProperty PresetManagerProperty =
        DependencyProperty.Register(nameof(PresetManager), typeof(TemplatePresetManager), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(null, OnPresetManagerChanged));

    public static readonly DependencyProperty TokenServiceProperty =
        DependencyProperty.Register(nameof(TokenService), typeof(TokenService), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(null, OnTokenServiceChanged));

    public static readonly DependencyProperty IsRenameButtonEnabledProperty =
        DependencyProperty.Register(nameof(IsRenameButtonEnabled), typeof(bool), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(false));

    public static readonly DependencyProperty IsPresetSelectedProperty =
        DependencyProperty.Register(nameof(IsPresetSelected), typeof(bool), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(false));

    public static readonly DependencyProperty IsSaveButtonEnabledProperty =
        DependencyProperty.Register(nameof(IsSaveButtonEnabled), typeof(bool), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(false));

    public static readonly DependencyProperty IsCreateButtonEnabledProperty =
        DependencyProperty.Register(nameof(IsCreateButtonEnabled), typeof(bool), typeof(TemplatePresetManagerControl),
            new PropertyMetadata(false));

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

    public bool IsRenameButtonEnabled
    {
        get => (bool)GetValue(IsRenameButtonEnabledProperty);
        set => SetValue(IsRenameButtonEnabledProperty, value);
    }

    public bool IsPresetSelected
    {
        get => (bool)GetValue(IsPresetSelectedProperty);
        set => SetValue(IsPresetSelectedProperty, value);
    }

    public bool IsSaveButtonEnabled
    {
        get => (bool)GetValue(IsSaveButtonEnabledProperty);
        set => SetValue(IsSaveButtonEnabledProperty, value);
    }

    public bool IsCreateButtonEnabled
    {
        get => (bool)GetValue(IsCreateButtonEnabledProperty);
        set => SetValue(IsCreateButtonEnabledProperty, value);
    }

    public TemplatePresetManagerControl()
    {
        InitializeComponent();
        PresetNameTextBox.TextChanged += PresetNameTextBox_TextChanged;
    }

    private static void OnTokenServiceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not TemplatePresetManagerControl control) return;

        if (e.OldValue is TokenService oldService)
        {
            oldService.PropertyChanged -= control.TokenService_PropertyChanged;
        }

        if (e.NewValue is TokenService newService)
        {
            newService.PropertyChanged += control.TokenService_PropertyChanged;
        }

        control.UpdateButtonStates();
    }

    private void TokenService_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TokenService.FileNameTemplate))
        {
            UpdateButtonStates();
        }
    }

    private static void OnPresetManagerChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not TemplatePresetManagerControl control) return;

        if (e.OldValue is TemplatePresetManager oldManager)
        {
            oldManager.PropertyChanged -= control.PresetManager_PropertyChanged;
            oldManager.TemplatePresets.CollectionChanged -= control.TemplatePresets_CollectionChanged;
        }

        if (e.NewValue is TemplatePresetManager newManager)
        {
            newManager.PropertyChanged += control.PresetManager_PropertyChanged;
            newManager.TemplatePresets.CollectionChanged += control.TemplatePresets_CollectionChanged;
        }

        control.UpdateEmptyState();
        control.UpdatePresetNameTextBox();
    }

    private void UpdateEmptyState()
    {
        var isEmpty = PresetManager?.TemplatePresets.Count == 0;
        EmptyStatePanel.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
        TemplatePresetsListBox.Visibility = isEmpty ? Visibility.Collapsed : Visibility.Visible;
    }

    private void PresetManager_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TemplatePresetManager.SelectedTemplatePreset))
        {
            UpdatePresetNameTextBox();
            UpdateButtonStates();
        }
    }

    private void TemplatePresets_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        UpdateEmptyState();
    }

    private void UpdatePresetNameTextBox()
    {
        var selected = PresetManager?.SelectedTemplatePreset;
        PresetNameTextBox.Text = selected?.Name ?? "";

        if (selected?.Template is { Length: > 0 } template && 
            TokenService != null && 
            TokenService.FileNameTemplate != template)
        {
            TokenService.FileNameTemplate = template;
        }

    }

    private void PresetNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        var selected = PresetManager?.SelectedTemplatePreset;
        var trimmedName = GetTrimmedPresetName();

        IsPresetSelected = selected != null;

        IsRenameButtonEnabled = selected != null &&
                               !string.IsNullOrWhiteSpace(trimmedName) &&
                               trimmedName != selected.Name &&
                               (trimmedName == selected.Name || !PresetManager!.PresetNameExists(trimmedName));

        IsSaveButtonEnabled = selected != null &&
                             TokenService != null &&
                             selected.Template != TokenService.FileNameTemplate;

        IsCreateButtonEnabled = PresetManager != null &&
                               TokenService != null &&
                               !string.IsNullOrWhiteSpace(trimmedName) &&
                               !PresetManager.PresetNameExists(trimmedName) &&
                               !string.IsNullOrWhiteSpace(TokenService.FileNameTemplate) &&
                               TokenService.IsFileNameTemplateValid;
    }

    private void DeletePresetButton_Click(object sender, RoutedEventArgs e)
    {
        PresetManager?.DeleteSelectedPreset();
        UpdateEmptyState();
    }

    private void CreatePresetButton_Click(object? sender, RoutedEventArgs e)
    {
        PresetManager!.CreatePreset(GetTrimmedPresetName(), TokenService!.FileNameTemplate, out _);
        RefreshUI();
    }

    private void SaveChangesButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;

        PresetManager!.UpdateSelectedTemplate(TokenService!.FileNameTemplate);
        UpdateButtonStates();
    }

    private void RefreshUI()
    {
        UpdateEmptyState();
    }

    private void RenamePresetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;

        PresetManager!.RenameSelected(GetTrimmedPresetName(), out _);
        UpdateButtonStates();
    }

    private void DuplicatePresetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!HasSelectedPreset()) return;

        PresetManager!.DuplicateSelected();
        RefreshUI();
    }

    private bool HasSelectedPreset() => PresetManager?.SelectedTemplatePreset != null;
    
    private string GetTrimmedPresetName() => PresetNameTextBox.Text.Trim();

}
