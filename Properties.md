```csharp
// New class: PropertyManager
// This class provides centralized access to all properties, handling mapping, errors, and non-iProperties.
// It works with any Document type, falling back gracefully for non-applicable properties.

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Inventor;
using System.Linq;
using System.Diagnostics;

namespace FlatPatternExporter
{
    public class PropertyManager
    {
        private readonly Document _document;
        private static readonly string SheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";

        private static readonly Dictionary<string, (string SetName, string InventorName)> PropertyMapping = new Dictionary<string, (string, string)>
        {
            // Summary Information
            { "Author", ("Summary Information", "Author") },
            { "Revision", ("Summary Information", "Revision Number") },
            { "Title", ("Summary Information", "Title") },
            { "Subject", ("Summary Information", "Subject") },
            { "Keywords", ("Summary Information", "Keywords") },
            { "Comments", ("Summary Information", "Comments") },

            // Document Summary Information
            { "Category", ("Document Summary Information", "Category") },
            { "Manager", ("Document Summary Information", "Manager") },
            { "Company", ("Document Summary Information", "Company") },

            // Design Tracking Properties
            { "PartNumber", ("Design Tracking Properties", "Part Number") },
            { "Description", ("Design Tracking Properties", "Description") },
            { "Material", ("Design Tracking Properties", "Material") },
            { "Project", ("Design Tracking Properties", "Project") },
            { "StockNumber", ("Design Tracking Properties", "Stock Number") },
            { "CreationTime", ("Design Tracking Properties", "Creation Time") },
            { "CostCenter", ("Design Tracking Properties", "Cost Center") },
            { "CheckedBy", ("Design Tracking Properties", "Checked By") },
            { "EngApprovedBy", ("Design Tracking Properties", "Engr Approved By") },
            { "UserStatus", ("Design Tracking Properties", "User Status") },
            { "CatalogWebLink", ("Design Tracking Properties", "Catalog Web Link") },
            { "Vendor", ("Design Tracking Properties", "Vendor") },
            { "MfgApprovedBy", ("Design Tracking Properties", "Mfg Approved By") },
            { "DesignStatus", ("Design Tracking Properties", "Design Status") },
            { "Designer", ("Design Tracking Properties", "Designer") },
            { "Engineer", ("Design Tracking Properties", "Engineer") },
            { "Authority", ("Design Tracking Properties", "Authority") },
            { "Mass", ("Design Tracking Properties", "Mass") },
            { "SurfaceArea", ("Design Tracking Properties", "SurfaceArea") },
            { "Volume", ("Design Tracking Properties", "Volume") },
            { "SheetMetalRule", ("Design Tracking Properties", "Sheet Metal Rule") },
            { "FlatPatternWidth", ("Design Tracking Properties", "Flat Pattern Width") },
            { "FlatPatternLength", ("Design Tracking Properties", "Flat Pattern Length") },
            { "FlatPatternArea", ("Design Tracking Properties", "Flat Pattern Area") },
            { "Appearance", ("Design Tracking Properties", "Appearance") }
        };

        public PropertyManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public string GetProperty(string setName, string propName)
        {
            try
            {
                var propSet = _document.PropertySets[setName];
                var prop = propSet[propName];
                return prop.Value?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting property '{propName}' from set '{setName}': {ex.Message}");
                return string.Empty;
            }
        }

        public void SetProperty(string setName, string propName, object value)
        {
            try
            {
                var propSet = _document.PropertySets[setName];
                var prop = propSet[propName];
                prop.Value = value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting property '{propName}' in set '{setName}': {ex.Message}");
                MessageBox.Show($"Failed to update property '{propName}'.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public string GetMappedProperty(string ourName)
        {
            if (PropertyMapping.TryGetValue(ourName, out var mapping))
            {
                return GetProperty(mapping.SetName, mapping.InventorName);
            }

            // For custom properties, assume they are in "Inventor User Defined Properties" and use ourName directly
            return GetProperty("Inventor User Defined Properties", ourName);
        }

        public void SetMappedProperty(string ourName, object value)
        {
            if (PropertyMapping.TryGetValue(ourName, out var mapping))
            {
                SetProperty(mapping.SetName, mapping.InventorName, value);
                return;
            }

            // For custom properties
            SetProperty("Inventor User Defined Properties", ourName, value);
        }

        // Non-iProperty methods
        public string GetFileName()
        {
            try
            {
                return Path.GetFileNameWithoutExtension(_document.FullFileName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting FileName: {ex.Message}");
                return string.Empty;
            }
        }

        public string GetFullFileName()
        {
            try
            {
                return _document.FullFileName;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting FullFileName: {ex.Message}");
                return string.Empty;
            }
        }

        public string GetModelState()
        {
            try
            {
                return _document.ModelStateName ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting ModelState: {ex.Message}");
                return string.Empty;
            }
        }

        public double GetThickness()
        {
            try
            {
                if (_document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
                {
                    var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                    var thicknessParam = smCompDef.Thickness;
                    return Math.Round((double)thicknessParam.Value * 10, 1); // Convert to mm and round to one decimal
                }
                return 0.0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting Thickness: {ex.Message}");
                return 0.0;
            }
        }

        public bool HasFlatPattern()
        {
            try
            {
                if (_document is PartDocument partDoc && partDoc.SubType == SheetMetalSubType)
                {
                    var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                    return smCompDef.HasFlatPattern;
                }
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error checking HasFlatPattern: {ex.Message}");
                return false;
            }
        }
    }
}
```

### Интеграция в существующий код
Удалите все старые методы чтения свойств (`ReadDocumentProperties`, `ReadCoreIProperties`, `ReadExtendedIProperties`, `ReadSummaryInformation`, `ReadDocumentSummaryInformation`, `ReadExtendedDesignTrackingProperties`, `GetProperty`, `SetProperty` и т.д.), так как они теперь централизованы в `PropertyManager`.

В методе `ReadAllPropertiesFromPart(PartDocument partDoc, PartData partData)` замените на:
```csharp
private void ReadAllPropertiesFromPart(PartDocument partDoc, PartData partData)
{
    var mgr = new PropertyManager(partDoc);

    // Non-iProperties
    partData.FileName = mgr.GetFileName();
    partData.FullFileName = mgr.GetFullFileName();
    partData.ModelState = mgr.GetModelState();
    partData.Thickness = mgr.GetThickness();
    partData.HasFlatPattern = mgr.HasFlatPattern();

    // Core iProperties (always read)
    partData.PartNumber = mgr.GetMappedProperty("PartNumber");
    partData.Description = mgr.GetMappedProperty("Description");
    partData.Material = mgr.GetMappedProperty("Material");

    // Extended iProperties (conditional on column presence)
    if (IsColumnPresent("Author")) partData.Author = mgr.GetMappedProperty("Author");
    if (IsColumnPresent("Revision")) partData.Revision = mgr.GetMappedProperty("Revision");
    if (IsColumnPresent("Title")) partData.Title = mgr.GetMappedProperty("Title");
    if (IsColumnPresent("Subject")) partData.Subject = mgr.GetMappedProperty("Subject");
    if (IsColumnPresent("Keywords")) partData.Keywords = mgr.GetMappedProperty("Keywords");
    if (IsColumnPresent("Comments")) partData.Comments = mgr.GetMappedProperty("Comments");
    if (IsColumnPresent("Category")) partData.Category = mgr.GetMappedProperty("Category");
    if (IsColumnPresent("Manager")) partData.Manager = mgr.GetMappedProperty("Manager");
    if (IsColumnPresent("Company")) partData.Company = mgr.GetMappedProperty("Company");
    if (IsColumnPresent("Project")) partData.Project = mgr.GetMappedProperty("Project");
    if (IsColumnPresent("StockNumber")) partData.StockNumber = mgr.GetMappedProperty("StockNumber");
    if (IsColumnPresent("CreationTime")) partData.CreationTime = mgr.GetMappedProperty("CreationTime");
    if (IsColumnPresent("CostCenter")) partData.CostCenter = mgr.GetMappedProperty("CostCenter");
    if (IsColumnPresent("CheckedBy")) partData.CheckedBy = mgr.GetMappedProperty("CheckedBy");
    if (IsColumnPresent("EngApprovedBy")) partData.EngApprovedBy = mgr.GetMappedProperty("EngApprovedBy");
    if (IsColumnPresent("UserStatus")) partData.UserStatus = mgr.GetMappedProperty("UserStatus");
    if (IsColumnPresent("CatalogWebLink")) partData.CatalogWebLink = mgr.GetMappedProperty("CatalogWebLink");
    if (IsColumnPresent("Vendor")) partData.Vendor = mgr.GetMappedProperty("Vendor");
    if (IsColumnPresent("MfgApprovedBy")) partData.MfgApprovedBy = mgr.GetMappedProperty("MfgApprovedBy");
    if (IsColumnPresent("DesignStatus")) partData.DesignStatus = mgr.GetMappedProperty("DesignStatus");
    if (IsColumnPresent("Designer")) partData.Designer = mgr.GetMappedProperty("Designer");
    if (IsColumnPresent("Engineer")) partData.Engineer = mgr.GetMappedProperty("Engineer");
    if (IsColumnPresent("Authority")) partData.Authority = mgr.GetMappedProperty("Authority");
    if (IsColumnPresent("Mass")) partData.Mass = mgr.GetMappedProperty("Mass");
    if (IsColumnPresent("SurfaceArea")) partData.SurfaceArea = mgr.GetMappedProperty("SurfaceArea");
    if (IsColumnPresent("Volume")) partData.Volume = mgr.GetMappedProperty("Volume");
    if (IsColumnPresent("SheetMetalRule")) partData.SheetMetalRule = mgr.GetMappedProperty("SheetMetalRule");
    if (IsColumnPresent("FlatPatternWidth")) partData.FlatPatternWidth = mgr.GetMappedProperty("FlatPatternWidth");
    if (IsColumnPresent("FlatPatternLength")) partData.FlatPatternLength = mgr.GetMappedProperty("FlatPatternLength");
    if (IsColumnPresent("FlatPatternArea")) partData.FlatPatternArea = mgr.GetMappedProperty("FlatPatternArea");
    if (IsColumnPresent("Appearance")) partData.Appearance = mgr.GetMappedProperty("Appearance");
}
```

### Другие изменения
- В `GetCustomIPropertyValue`: Замените на `mgr.GetMappedProperty(propertyName)` (поскольку кастомные свойства обрабатываются автоматически).
- В `EditIProperty_Click`: Используйте `mgr.SetMappedProperty("PartNumber", editDialog.PartNumber);` и `mgr.SetMappedProperty("Description", editDialog.Description);`.
- В `ProcessBOMRow` и других местах: Замените прямые вызовы `GetProperty` на `mgr.GetMappedProperty("PartNumber")` (создайте экземпляр `PropertyManager` для соответствующего документа).
- В `GetDocumentProperties`: Замените на использование `PropertyManager` для получения `PartNumber` и `Description`.
- Удалите все дублирующиеся методы доступа к свойствам (например, в `FlatPatternExporterMainWindow.xaml.cs` и `.part2.cs`).
- Для любого другого доступа к свойствам (например, в `GetPartDataAsync`): Используйте `PropertyManager`.
- Это обеспечивает единый интерфейс, централизованную обработку ошибок и маппинг имен без дублирования. Менеджер универсален и работает с любым `Document`, возвращая defaults при ошибках или несовместимости.