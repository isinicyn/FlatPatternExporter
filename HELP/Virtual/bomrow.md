# BOMRow Object

#### Description

The BOMRow object represents an item in the BOM based on the parent BOMView.

#### Methods

| Name | Description |
| --- | --- |
| [GetReferenceKey](#getreferencekey) | Method that generates and returns the reference key for this entity. |

#### Properties

| Name | Description |
| --- | --- |
| [Application](#application) | Returns the top-level parent application object.  When used the context of Inventor, an Application object is returned.  When used in the context of Apprentice, an ApprenticeServer object is returned. |
| [AttributeSets](#attributesets) | Property that returns the AttributeSets collection object associated with this object. |
| [BOMStructure](#bomstructure) | Gets and sets how the BOM item represented by this row is used/viewed relating to design. |
| [ChildRows](#childrows) | Property that gets the BOMRowsEnumerator object containing the locally-stored rows under this BOMRow. This property will return Nothing for BOMRows in a parts-only view and if there are no sub-rows for this BOMRow. |
| [ComponentDefinitions](#componentdefinitions) | Property that returns the ComponentDefinitions associated with this row in the BOM. These could be part, sheet metal, assembly, weldment or a virtual component definitions. This enumerator will return just one component definition unless this row is a merged one, in which case all associated component definitions are returned. The first component definition in the enumerator is always the primary component definition. |
| [ComponentOccurrences](#componentoccurrences) | Gets the ComponentOccurrences associated with this row in the BOM. |
| [ItemNumber](#itemnumber) | Gets and sets the item number of the BOM item. |
| [ItemNumberLocked](#itemnumberlocked) | Gets and sets whether the item identifier can be edited. |
| [ItemQuantity](#itemquantity) | Property that gets the number of instances not marked as phantom or reference represented by this BOM row. |
| [Merged](#merged) | Property that returns whether this row is a result of a merging multiple rows. If true, the ComponentDefinitions property returns all the associated component definitions. This property will return False for all rows in the data BOMView. |
| [OccurrencePropertySets](#occurrencepropertysets) | Read-only property that returns a PropertySets object associated with this BOMRow. This only applies to non merged component rows with Instance Properties. |
| [Parent](#parent) | Property that returns the parent BOMView or the BOMRow object. |
| [Promoted](#promoted) | Property that returns whether this row was promoted (for instance, as a result of the parent subassembly being marked phantom). This property will return False for all rows in the data BOMView. |
| [ReferencedFileDescriptor](#referencedfiledescriptor) | Gets the FileDescriptor for the component referenced by this row. This only applies to non merged component rows and non local components and immediately referenced components. Therefore this would only be useful in the data view. |
| [RolledUp](#rolledup) | Indicates whether this row is a result of rolling up multiple promoted rows of the same ComponentDefinition. |
| [TotalQuantity](#totalquantity) | Gets and sets the total quantity of the BOM item. Overrides cannot be set for parts only views. |
| [TotalQuantityOverridden](#totalquantityoverridden) | Gets and sets whether the TotalQuantity property has been overridden. This property can only be set to False, in which case the override on the value will be removed. |
| [Type](#type) | Returns an ObjectTypeEnum indicating this object's type. |

#### Accessed From

- [BOMRowsEnumerator.Item](bomrowsenumerator.md#item)
- [DrawingBOMRow.BOMRow](drawingbomrow.md#bomrow)

#### Samples

| Name | Description |
| --- | --- |
| [Using the BOM APIs](../API_Samples/bom_sample.md) | This sample demonstrates the Bill of Materials API functionality in assemblies. |

#### Version

Introduced in version 10

## Methods

<a name="getreferencekey"></a>
### GetReferenceKey Method

**Parent Object:** [BOMRow](#bomrow)

#### Description

Method that generates and returns the reference key for this entity.

#### Syntax

BOMRow.GetReferenceKey( ReferenceKey() As Byte, [KeyContext] As Long )

#### Remarks

The reference key is an array of bytes that can be used as a persistent reference for this entity. To obtain the entity at a later time using the reference key you use the BindKeyToObject method of the object. The ReferenceKeyManager object is obtained using the ReferenceKeyManager property of the Document object.

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| ReferenceKey | Byte | Input/output array of Bytes that contains the reference key. |
| KeyContext | Long | Input Long that specifies the key context. The key context must be supplied when working with any B-Rep entities (and SurfaceBody, FaceShell, Face, Edge, EdgeUse and Vertex objects). A key context is created using the CreateKeyContext method of the ReferenceKeyManager object. For all other object types, the key context argument is not used and is ignored if provided. This is an optional argument whose default value is 0. |

#### Version

Introduced in version 2009


---

## Properties

<a name="application"></a>
### Application Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Returns the top-level parent application object. When used the context of Inventor, an Application object is returned. When used in the context of Apprentice, an ApprenticeServer object is returned.

#### Syntax

BOMRow.Application() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 10


<a name="attributesets"></a>
### AttributeSets Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that returns the AttributeSets collection object associated with this object.

#### Syntax

BOMRow.AttributeSets() As [AttributeSets](attributesets.md)

#### Property Value

This is a read only property whose value is an [AttributeSets](attributesets.md).

#### Version

Introduced in version 2009


<a name="bomstructure"></a>
### BOMStructure Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets and sets how the BOM item represented by this row is used/viewed relating to design.

#### Syntax

BOMRow.BOMStructure() As [BOMStructureEnum](../API_Enums/bomstructureenum.md)

#### Property Value

This is a read/write property whose value is a [BOMStructureEnum](../API_Enums/bomstructureenum.md).

#### Version

Introduced in version 10


<a name="childrows"></a>
### ChildRows Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that gets the BOMRowsEnumerator object containing the locally-stored rows under this BOMRow. This property will return Nothing for BOMRows in a parts-only view and if there are no sub-rows for this BOMRow.

#### Syntax

BOMRow.ChildRows() As [BOMRowsEnumerator](bomrowsenumerator.md)

#### Property Value

This is a read only property whose value is a [BOMRowsEnumerator](bomrowsenumerator.md).

#### Version

Introduced in version 11


<a name="componentdefinitions"></a>
### ComponentDefinitions Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that returns the ComponentDefinitions associated with this row in the BOM. These could be part, sheet metal, assembly, weldment or a virtual component definitions. This enumerator will return just one component definition unless this row is a merged one, in which case all associated component definitions are returned. The first component definition in the enumerator is always the primary component definition.

#### Syntax

BOMRow.ComponentDefinitions() As [ComponentDefinitionsEnumerator](componentdefinitionsenumerator.md)

#### Property Value

This is a read only property whose value is a [ComponentDefinitionsEnumerator](componentdefinitionsenumerator.md).

#### Version

Introduced in version 10


<a name="componentoccurrences"></a>
### ComponentOccurrences Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets the ComponentOccurrences associated with this row in the BOM.

#### Syntax

BOMRow.ComponentOccurrences() As [ComponentOccurrencesEnumerator](componentoccurrencesenumerator.md)

#### Property Value

This is a read only property whose value is a [ComponentOccurrencesEnumerator](componentoccurrencesenumerator.md).

#### Version

Introduced in version 2022


<a name="itemnumber"></a>
### ItemNumber Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets and sets the item number of the BOM item.

#### Syntax

BOMRow.ItemNumber() As String

#### Property Value

This is a read/write property whose value is a String.

#### Version

Introduced in version 10


<a name="itemnumberlocked"></a>
### ItemNumberLocked Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets and sets whether the item identifier can be edited.

#### Syntax

BOMRow.ItemNumberLocked() As Boolean

#### Property Value

This is a read/write property whose value is a Boolean.

#### Version

Introduced in version 11


<a name="itemquantity"></a>
### ItemQuantity Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that gets the number of instances not marked as phantom or reference represented by this BOM row.

#### Syntax

BOMRow.ItemQuantity() As Long

#### Property Value

This is a read only property whose value is a Long.

#### Version

Introduced in version 10


<a name="merged"></a>
### Merged Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that returns whether this row is a result of a merging multiple rows. If true, the ComponentDefinitions property returns all the associated component definitions. This property will return False for all rows in the data BOMView.

#### Syntax

BOMRow.Merged() As Boolean

#### Property Value

This is a read only property whose value is a Boolean.

#### Version

Introduced in version 10


<a name="occurrencepropertysets"></a>
### OccurrencePropertySets Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Read-only property that returns a PropertySets object associated with this BOMRow. This only applies to non merged component rows with Instance Properties.

#### Syntax

BOMRow.OccurrencePropertySets() As [PropertySets](propertysets.md)

#### Property Value

This is a read only property whose value is a [PropertySets](propertysets.md).

#### Version

Introduced in version 2022


<a name="parent"></a>
### Parent Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that returns the parent BOMView or the BOMRow object.

#### Syntax

BOMRow.Parent() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 10


<a name="promoted"></a>
### Promoted Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Property that returns whether this row was promoted (for instance, as a result of the parent subassembly being marked phantom). This property will return False for all rows in the data BOMView.

#### Syntax

BOMRow.Promoted() As Boolean

#### Property Value

This is a read only property whose value is a Boolean.

#### Version

Introduced in version 10


<a name="referencedfiledescriptor"></a>
### ReferencedFileDescriptor Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets the FileDescriptor for the component referenced by this row. This only applies to non merged component rows and non local components and immediately referenced components. Therefore this would only be useful in the data view.

#### Syntax

BOMRow.ReferencedFileDescriptor() As [FileDescriptor](filedescriptor.md)

#### Property Value

This is a read only property whose value is a [FileDescriptor](filedescriptor.md).

#### Version

Introduced in version 2008


<a name="rolledup"></a>
### RolledUp Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Indicates whether this row is a result of rolling up multiple promoted rows of the same ComponentDefinition.

#### Syntax

BOMRow.RolledUp() As Boolean

#### Property Value

This is a read only property whose value is a Boolean.

#### Version

Introduced in version 10


<a name="totalquantity"></a>
### TotalQuantity Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets and sets the total quantity of the BOM item. Overrides cannot be set for parts only views.

#### Syntax

BOMRow.TotalQuantity() As String

#### Property Value

This is a read/write property whose value is a String.

#### Version

Introduced in version 10


<a name="totalquantityoverridden"></a>
### TotalQuantityOverridden Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Gets and sets whether the TotalQuantity property has been overridden. This property can only be set to False, in which case the override on the value will be removed.

#### Syntax

BOMRow.TotalQuantityOverridden() As Boolean

#### Property Value

This is a read/write property whose value is a Boolean.

#### Version

Introduced in version 10


<a name="type"></a>
### Type Property

**Parent Object:** [BOMRow](#bomrow)

#### Description

Returns an ObjectTypeEnum indicating this object's type.

#### Syntax

BOMRow.Type() As [ObjectTypeEnum](../API_Enums/objecttypeenum.md)

#### Property Value

This is a read only property whose value is an [ObjectTypeEnum](../API_Enums/objecttypeenum.md).

#### Version

Introduced in version 10


---

