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

    public PartData GetPartData(ScannedPart part, IReadOnlyDictionary<string, string> rootProperties, int quantity, int itemNumber, bool loadThumbnail = true)
    {
        var partDoc = _documentScanner.DocumentCache.GetCachedPartDocument(part.FullDocumentName) ?? _inventorManager.FindPartDocument(part.FullDocumentName);
        if (partDoc == null) return null!;

        var partData = GetPartData(partDoc, quantity, itemNumber, loadThumbnail);
        partData.BomItem = part.BomItem;
        partData.RootProperties = rootProperties;
        return partData;
    }

    public PartData GetPartData(PartDocument partDoc, int quantity, int itemNumber, bool loadThumbnail = true)
    {
        var partData = new PartData
        {
            Item = itemNumber,
            OriginalQuantity = quantity
        };

        var mgr = new PropertyManager((Document)partDoc);

        ReadAllPropertiesFromPart(partDoc, partData, mgr);

        var userValues = mgr.GetMappedProperties(PropertyMetadataRegistry.UserDefinedProperties.Select(p => p.InternalName));
        foreach (var userDefinedProperty in PropertyMetadataRegistry.UserDefinedProperties)
        {
            var propertyName = userDefinedProperty.InventorPropertyName ?? "";
            partData.UserDefinedProperties[propertyName] = userValues[userDefinedProperty.InternalName];
        }

        if (loadThumbnail)
        {
            partData.Preview = ThumbnailGenerator.GetThumbnail(partData.FullFileName);
        }

        partData.SetQuantityInternal(quantity);

        return partData;
    }

    private static void ReadAllPropertiesFromPart(PartDocument _, PartData partData, PropertyManager mgr)
    {
        partData.FileName = mgr.GetFileName();
        partData.FullFileName = mgr.GetFullFileName();
        partData.FullDocumentName = mgr.GetFullDocumentName();
        partData.ModelState = mgr.GetModelState();
        partData.HasFlatPattern = mgr.HasFlatPattern();
        partData.Thickness = mgr.GetThickness();

        SetExpressionStatesForAllProperties(partData, mgr);

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
    }

    private static void SetExpressionStatesForAllProperties(PartData partData, PropertyManager mgr)
    {
        partData.BeginExpressionBatch();
        try
        {
            foreach (var property in PropertyMetadataRegistry.GetEditableProperties())
            {
                var isExpression = mgr.IsMappedPropertyExpression(property);
                partData.SetPropertyExpressionState(property, isExpression);
            }
            foreach (var userProperty in PropertyMetadataRegistry.UserDefinedProperties)
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

        var isStandardProperty = PropertyMetadataRegistry.Properties.ContainsKey(propertyName);

        if (isStandardProperty)
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