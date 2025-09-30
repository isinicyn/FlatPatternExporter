using System.Windows.Data;
using System.Windows.Markup;
using FlatPatternExporter.Services;
using FlatPatternExporter.Converters;

namespace FlatPatternExporter.Extensions;

[MarkupExtensionReturnType(typeof(BindingExpression))]
public class LocalizeExtension : MarkupExtension
{
    public string Key { get; set; }

    public LocalizeExtension(string key)
    {
        Key = key;
    }

    public override object ProvideValue(IServiceProvider serviceProvider)
    {
        var binding = new System.Windows.Data.Binding("CurrentCulture")
        {
            Source = LocalizationManager.Instance,
            Converter = new LocalizationKeyConverter(),
            ConverterParameter = Key,
            Mode = BindingMode.OneWay
        };

        return binding.ProvideValue(serviceProvider);
    }
}