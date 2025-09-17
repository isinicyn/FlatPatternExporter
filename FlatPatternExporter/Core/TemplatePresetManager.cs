using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace FlatPatternExporter.Core;

public class TemplatePreset : INotifyPropertyChanged
{
    private string _name = string.Empty;
    public string Name
    {
        get => _name;
        set
        {
            if (_name != value)
            {
                _name = value;
                OnPropertyChanged();
            }
        }
    }

    private string _template = string.Empty;
    public string Template
    {
        get => _template;
        set
        {
            if (_template != value)
            {
                _template = value;
                OnPropertyChanged();
            }
        }
    }
    
    public override string ToString() => Name;

    public event PropertyChangedEventHandler? PropertyChanged;
    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
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


    public bool CreatePreset(string presetName, string template, out string? error)
    {
        error = null;
        presetName = presetName.Trim();

        var newPreset = new TemplatePreset { Name = presetName, Template = template };
        TemplatePresets.Add(newPreset);
        SelectedTemplatePreset = newPreset;
        return true;
    }

    public void UpdateSelectedTemplate(string template)
    {
        SelectedTemplatePreset!.Template = template;
    }

    public bool RenameSelected(string newName, out string? error)
    {
        error = null;
        if (SelectedTemplatePreset == null)
            return false;

        newName = newName.Trim();
        SelectedTemplatePreset.Name = newName;
        return true;
    }

    public void DuplicateSelected()
    {
        var baseName = SelectedTemplatePreset!.Name;
        var newName = GenerateUniqueName($"{baseName} (копия)");
        var duplicate = new TemplatePreset
        {
            Name = newName,
            Template = SelectedTemplatePreset.Template
        };
        TemplatePresets.Add(duplicate);
        SelectedTemplatePreset = duplicate;
    }


    private string GenerateUniqueName(string baseName)
    {
        if (!PresetNameExists(baseName))
            return baseName;

        int i = 2;
        string candidate;
        do
        {
            candidate = $"{baseName} {i++}";
        } while (PresetNameExists(candidate));

        return candidate;
    }

    public bool PresetNameExists(string name) =>
        TemplatePresets.Any(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));

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
