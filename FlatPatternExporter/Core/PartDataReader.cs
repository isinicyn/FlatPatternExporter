using System.Collections.ObjectModel;
using System.Linq;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using Inventor;

namespace FlatPatternExporter.Core;

public class PartDataReader
{
    private readonly InventorManager _inventorManager;
    private readonly DocumentScanner _documentScanner;

    public PartDataReader(InventorManager inventorManager, DocumentScanner documentScanner)
    {
        _inventorManager = inventorManager;
        _documentScanner = documentScanner;
    }

    /// <param name="displayedColumns">Properties shown as table columns; null when the data is not displayed
    /// (export only), then the thumbnail and the fx states are not read.</param>
    public PartData GetPartData(ScannedPart part, IReadOnlyDictionary<string, string> rootProperties, int quantity, int itemNumber, IReadOnlySet<string>? displayedColumns = null)
    {
        var partDoc = _documentScanner.DocumentCache.GetCachedPartDocument(part.FullDocumentName) ?? _inventorManager.FindPartDocument(part.FullDocumentName);
        if (partDoc == null) return null!;

        var mgr = new PropertyManager((Document)partDoc);
        var partData = new PartData
        {
            Item = itemNumber,
            OriginalQuantity = quantity,
            FileName = System.IO.Path.GetFileNameWithoutExtension(part.FullFileName),
            FullFileName = part.FullFileName,
            FullDocumentName = part.FullDocumentName,
            ModelState = part.ModelState,
            HasFlatPattern = mgr.HasFlatPattern(),
            Thickness = mgr.GetThickness(),
            BomItem = part.BomItem,
            RootProperties = rootProperties
        };

        ReadPropertyValues(partData, mgr);

        if (displayedColumns != null)
        {
            // Runs after the values: properties missing from the part are skipped without COM calls
            SetExpressionStates(partData, mgr, displayedColumns);
            partData.Preview = ThumbnailGenerator.GetThumbnail(part.FullFileName);
        }

        partData.SetQuantityInternal(quantity);

        return partData;
    }

    private static void ReadPropertyValues(PartData partData, PropertyManager mgr)
    {
        var values = mgr.GetMappedProperties(PropertyMetadataRegistry.Properties.Values
            .Where(p => p.Type == PropertyMetadataRegistry.PropertyType.IProperty)
            .Select(p => p.InternalName));

        foreach (var (internalName, value) in values)
        {
            var propInfo = typeof(PartData).GetProperty(internalName);

            if (propInfo != null && !string.IsNullOrEmpty(value))
            {
                propInfo.SetValue(partData, value);
            }
        }

        var userValues = mgr.GetMappedProperties(PropertyMetadataRegistry.UserDefinedProperties.Select(p => p.InternalName));
        foreach (var userDefinedProperty in PropertyMetadataRegistry.UserDefinedProperties)
        {
            var propertyName = userDefinedProperty.InventorPropertyName ?? "";
            partData.UserDefinedProperties[propertyName] = userValues[userDefinedProperty.InternalName];
        }
    }

    private static void SetExpressionStates(PartData partData, PropertyManager mgr, IReadOnlySet<string> displayedColumns)
    {
        partData.BeginExpressionBatch();
        try
        {
            foreach (var property in PropertyMetadataRegistry.GetEditableProperties().Where(displayedColumns.Contains))
            {
                var isExpression = mgr.IsMappedPropertyExpression(property);
                partData.SetPropertyExpressionState(property, isExpression);
            }
            foreach (var userProperty in PropertyMetadataRegistry.UserDefinedProperties.Where(p => displayedColumns.Contains(p.InternalName)))
            {
                var isExpression = mgr.IsMappedPropertyExpression(userProperty.InternalName);
                var propertyName = userProperty.InventorPropertyName ?? "";
                partData.SetPropertyExpressionState($"UserDefinedProperties[{propertyName}]", isExpression);
            }
        }
        finally
        {
            partData.EndExpressionBatch();
        }
    }

    public void FillPropertyData(ObservableCollection<PartData> partsData, string propertyName)
    {
        if (partsData.Count == 0)
            return;

        // Standard property values are read with the part, an editable column only needs its fx states
        var standardProperty = PropertyMetadataRegistry.Properties.GetValueOrDefault(propertyName);

        if (standardProperty is { IsEditable: false })
        {
            return;
        }

        // For user-defined properties, extract the actual property name from InternalName
        var actualPropertyName = PropertyMetadataRegistry.IsUserDefinedProperty(propertyName)
            ? PropertyMetadataRegistry.GetInventorNameFromUserDefinedInternalName(propertyName)
            : propertyName;

        foreach (var partData in partsData)
        {
            var partDoc = _documentScanner.DocumentCache.GetCachedPartDocument(partData.FullDocumentName) ?? _inventorManager.FindPartDocument(partData.FullDocumentName);
            if (partDoc != null)
            {
                var mgr = new PropertyManager((Document)partDoc);

                if (standardProperty != null)
                {
                    partData.SetPropertyExpressionState(propertyName, mgr.IsMappedPropertyExpression(propertyName));
                    continue;
                }

                var value = mgr.GetMappedProperty(propertyName) ?? "";
                partData.AddUserDefinedProperty(actualPropertyName, value);

                var isExpression = mgr.IsMappedPropertyExpression(propertyName);
                partData.SetPropertyExpressionState($"UserDefinedProperties[{actualPropertyName}]", isExpression);
            }
        }
    }

    public void UpdateQuantitiesWithMultiplier(ObservableCollection<PartData> partsData, int multiplier)
    {
        foreach (var partData in partsData)
        {
            partData.IsOverridden = false;
            partData.SetQuantityInternal(partData.OriginalQuantity * multiplier);
            partData.IsMultiplied = multiplier > 1;
        }
    }

    public DocumentInfo GetDocumentInfo(Document? doc)
    {
        if (doc == null)
        {
            return new DocumentInfo
            {
                DocumentType = "",
                PartNumber = "",
                Description = "",
                ModelState = "",
                IsPrimaryModelState = true
            };
        }

        var mgr = new PropertyManager(doc);
        return new DocumentInfo
        {
            DocumentType = doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? LocalizationManager.Instance.GetString("DocumentType_Assembly") : LocalizationManager.Instance.GetString("DocumentType_Part"),
            PartNumber = mgr.GetMappedProperty("PartNumber"),
            Description = mgr.GetMappedProperty("Description"),
            ModelState = mgr.GetModelState(),
            IsPrimaryModelState = mgr.IsPrimaryModelState()
        };
    }
}

public class DocumentInfo
{
    public string DocumentType { get; set; } = "";
    public string PartNumber { get; set; } = "";
    public string Description { get; set; } = "";
    public string ModelState { get; set; } = "";
    public bool IsPrimaryModelState { get; set; } = true;
}