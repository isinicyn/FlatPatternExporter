# PropertySets Object

#### Description

Object that manages the collection of objects and provides the ability to add new property sets to the collection. See the article in the overviews section.

#### Methods

| Name | Description |
| --- | --- |
| [Add](#add) | Adds a new PropertySet. The new set's FMTID can be optionally provided (as a string). |
| [FlushToFile](#flushtofile) | Flush all of the Properties in each of the PropertySets onto the file. The 'dirty' flags are reset. Please note that this method is available in Apprentice only. |
| [PropertySetExists](#propertysetexists) | Function that returns a Boolean to indicate whether a PropertySet with the specified name exists in the PropertySets collection. |
| [RefreshFromFile](#refreshfromfile) | Refresh all of the Properties in each of the PropertySets from the File. The 'dirty' flags are reset and any edits are lost. Please note that this method is available in Apprentice only. |

#### Properties

| Name | Description |
| --- | --- |
| [Count](#count) | Property that returns the number of items in this collection. |
| [Dirty](#dirty) | Property that returns a Boolean flag that indicates whether any of the Properties have been edited, deleted or created. |
| [Item](#item) | Gets the set in this collection in a sequences fashion; by index, or by its name -- Display or Internal. |
| [Parent](#parent) | Property that returns the parent object from whom this object can logically be reached. |
| [Type](#type) | Returns an ObjectTypeEnum indicating this object's type. |

#### Accessed From

- [ApprenticeServerDocument.FilePropertySets](apprenticeserverdocument.md#filepropertysets)
- [ApprenticeServerDocument.PropertySets](apprenticeserverdocument.md#propertysets)
- [ApprenticeServerDrawingDocument.FilePropertySets](apprenticeserverdrawingdocument.md#filepropertysets)
- [ApprenticeServerDrawingDocument.PropertySets](apprenticeserverdrawingdocument.md#propertysets)
- [AssemblyDocument.FilePropertySets](assemblydocument.md#filepropertysets)
- [AssemblyDocument.PropertySets](assemblydocument.md#propertysets)
- [BOMRow.OccurrencePropertySets](bomrow.md#occurrencepropertysets)
- [ComponentOccurrence.OccurrencePropertySets](componentoccurrence.md#occurrencepropertysets)
- [ComponentOccurrenceProxy.OccurrencePropertySets](componentoccurrenceproxy.md#occurrencepropertysets)
- [Document.FilePropertySets](document.md#filepropertysets)
- [Document.PropertySets](document.md#propertysets)
- [DrawingDocument.FilePropertySets](drawingdocument.md#filepropertysets)
- [DrawingDocument.PropertySets](drawingdocument.md#propertysets)
- [PartDocument.FilePropertySets](partdocument.md#filepropertysets)
- [PartDocument.PropertySets](partdocument.md#propertysets)
- [PresentationDocument.FilePropertySets](presentationdocument.md#filepropertysets)
- [PresentationDocument.PropertySets](presentationdocument.md#propertysets)
- [PropertySet.Parent](propertyset.md#parent)
- [VirtualComponentDefinition.PropertySets](virtualcomponentdefinition.md#propertysets)

#### Samples

| Name | Description |
| --- | --- |
| [Using the BOM APIs](../API_Samples/bom_sample.md) | This sample demonstrates the Bill of Materials API functionality in assemblies. |
| [Update iProperty values using Apprentice](../API_Samples/ipropertyapprentice_sample.md) | Updates some iProperty values using Apprentice.  The document specified in the code for the Open method must exist. |
| [Create custom iProperties](../API_Samples/ipropertycreatecustom_sample.md) | Creates custom iProperties of various types. A document must be open when this sample is run. |
| [Get value of iProperty](../API_Samples/ipropertygetvalue_sample.md) | Demonstrates getting the values of the "Part Number" iProperty.  Any property can be retrieved by accesing the correct property set and property.A document must be open when this sample is run. |

#### Version

Introduced in version 4

## Methods

<a name="add"></a>
### Add Method

**Parent Object:** [PropertySets](#propertysets)

#### Description

Adds a new PropertySet. The new set's FMTID can be optionally provided (as a string).

#### Syntax

PropertySets.Add( Name As String, [InternalName] As Variant ) As [PropertySet](propertyset.md)

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| Name | String | Name of the PropertySet. If the name begins with an underscore the property set is hidden and can only be retrieved by asking for it by name. |
| InternalName | Variant | Input Variant that specifies the internal name of the PropertySet to be added. This is an optional argument whose default value is null. |

#### Version

Introduced in version 4


<a name="flushtofile"></a>
### FlushToFile Method

**Parent Object:** [PropertySets](#propertysets)

#### Description

Flush all of the Properties in each of the PropertySets onto the file. The 'dirty' flags are reset. Please note that this method is available in Apprentice only.

#### Syntax

PropertySets.FlushToFile()

#### Version

Introduced in version 4


<a name="propertysetexists"></a>
### PropertySetExists Method

**Parent Object:** [PropertySets](#propertysets)

#### Description

Function that returns a Boolean to indicate whether a PropertySet with the specified name exists in the PropertySets collection.

#### Syntax

PropertySets.PropertySetExists( PropertySetName As String, [PropertySet] As Variant ) As Boolean

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| PropertySetName | String | Input string value that indicating the property set name. This can be the display name or internal name. |
| PropertySet | Variant | Optional output the property set if found with the specified name, otherwise this will return Nothing. This is an optional argument whose default value is null. |

#### Version

Introduced in version 2013


<a name="refreshfromfile"></a>
### RefreshFromFile Method

**Parent Object:** [PropertySets](#propertysets)

#### Description

Refresh all of the Properties in each of the PropertySets from the File. The 'dirty' flags are reset and any edits are lost. Please note that this method is available in Apprentice only.

#### Syntax

PropertySets.RefreshFromFile()

#### Version

Introduced in version 4


---

## Properties

<a name="count"></a>
### Count Property

**Parent Object:** [PropertySets](#propertysets)

#### Description

Property that returns the number of items in this collection.

#### Syntax

PropertySets.Count() As Long

#### Property Value

This is a read only property whose value is a Long.

#### Version

Introduced in version 4


<a name="dirty"></a>
### Dirty Property

**Parent Object:** [PropertySets](#propertysets)

#### Description

Property that returns a Boolean flag that indicates whether any of the Properties have been edited, deleted or created.

#### Syntax

PropertySets.Dirty() As Boolean

#### Property Value

This is a read only property whose value is a Boolean.

#### Version

Introduced in version 4


<a name="item"></a>
### Item Property

**Parent Object:** [PropertySets](#propertysets)

#### Description

Gets the set in this collection in a sequences fashion; by index, or by its name -- Display or Internal.

#### Syntax

PropertySets.Item( Index As Variant ) As [PropertySet](propertyset.md)

#### Property Value

This is a read only property whose value is a [PropertySet](propertyset.md).

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| Index | Variant | Input Variant value that specifies the index of the PropertySet to return. |

#### Version

Introduced in version 4


<a name="parent"></a>
### Parent Property

**Parent Object:** [PropertySets](#propertysets)

#### Description

Property that returns the parent object from whom this object can logically be reached.

#### Syntax

PropertySets.Parent() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 4


<a name="type"></a>
### Type Property

**Parent Object:** [PropertySets](#propertysets)

#### Description

Returns an ObjectTypeEnum indicating this object's type.

#### Syntax

PropertySets.Type() As [ObjectTypeEnum](../API_Enums/objecttypeenum.md)

#### Property Value

This is a read only property whose value is an [ObjectTypeEnum](../API_Enums/objecttypeenum.md).

#### Version

Introduced in version 4


---

