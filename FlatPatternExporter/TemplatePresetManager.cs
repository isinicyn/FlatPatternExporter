using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter;

public class TemplatePreset
{
    public string Name { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
    
    public override string ToString() => Name;
}

public class TemplatePresetManager : INotifyPropertyChanged
{
    public ObservableCollection<TemplatePreset> TemplatePresets { get; } = [];

    private TemplatePreset? _selectedTemplatePreset;
    public TemplatePreset? SelectedTemplatePreset
    {
        get => _selectedTemplatePreset;
        set
        {
            _selectedTemplatePreset = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    
    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public bool SavePreset(string presetName, string template)
    {
        if (string.IsNullOrWhiteSpace(presetName))
        {
            MessageBox.Show("Введите имя пресета", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        var existingPreset = TemplatePresets.FirstOrDefault(p => p.Name == presetName);
        if (existingPreset != null)
        {
            var result = MessageBox.Show(
                $"Пресет с именем '{presetName}' уже существует. Перезаписать?",
                "Пресет уже существует",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                existingPreset.Template = template;
                SelectedTemplatePreset = existingPreset;
                return true;
            }
            return false;
        }

        var newPreset = new TemplatePreset
        {
            Name = presetName,
            Template = template
        };
        TemplatePresets.Add(newPreset);
        SelectedTemplatePreset = newPreset;
        return true;
    }

    public bool DeleteSelectedPreset()
    {
        if (SelectedTemplatePreset == null)
            return false;

        var result = MessageBox.Show(
            $"Удалить пресет '{SelectedTemplatePreset.Name}'?",
            "Подтверждение удаления",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            TemplatePresets.Remove(SelectedTemplatePreset);
            SelectedTemplatePreset = null;
            return true;
        }
        return false;
    }

    public void LoadPresets(IEnumerable<TemplatePresetData> presetData, int selectedPresetIndex = -1)
    {
        TemplatePresets.Clear();
        foreach (var data in presetData)
        {
            TemplatePresets.Add(new TemplatePreset
            {
                Name = data.Name,
                Template = data.Template
            });
        }

        if (selectedPresetIndex >= 0 && selectedPresetIndex < TemplatePresets.Count)
        {
            SelectedTemplatePreset = TemplatePresets[selectedPresetIndex];
        }
        else
        {
            SelectedTemplatePreset = null;
        }
    }

    public IEnumerable<TemplatePresetData> GetPresetData()
    {
        return TemplatePresets.Select(preset => new TemplatePresetData
        {
            Name = preset.Name,
            Template = preset.Template
        });
    }

    public int GetSelectedPresetIndex()
    {
        return SelectedTemplatePreset == null ? -1 : TemplatePresets.IndexOf(SelectedTemplatePreset);
    }
}