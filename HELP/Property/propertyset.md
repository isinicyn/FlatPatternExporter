# PropertySet Object

#### Description

Object that represents a PropertySet. This is a collection of related Properties. See the article in the overviews section.

#### Methods

| Name | Description |
| --- | --- |
| [Add](#add) | Adds a new Property to this PropertySet. |
| [Delete](#delete) | Method that deletes this PropertySet. |
| [GetPropertyInfo](#getpropertyinfo) | Method that returns property info in the PropertySet. |
| [SetPropertyValues](#setpropertyvalues) | Method that batch sets property values in the PropertySet. If a specified name is not existent in the property set a new property with the specified name will be created. |

#### Properties

| Name | Description |
| --- | --- |
| [Count](#count) | Property that returns the number of items in this collection. |
| [Dirty](#dirty) | Property that returns a Boolean flag that indicates whether any of the Properties have been edited, deleted or created. |
| [DisplayName](#displayname) | Gets/Sets the human-readable name associated with this Property Set. |
| [InternalName](#internalname) | Gets the unambiguous, internal name (FMTID) associated with this PropertySet. |
| [Item](#item) | Gets the Property given either its name or its sequential index. |
| [ItemByPropId](#itembypropid) | Gets the Property in this set by its PropId. |
| [Name](#name) | Gets the name of this PropertySet. |
| [Parent](#parent) | Property that returns the parent object from whom this object can logically be reached. |
| [Type](#type) | Returns an ObjectTypeEnum indicating this object's type. |

#### Accessed From

- [Property.Parent](property.md#parent)
- [PropertySets.Add](propertysets.md#add)
- [PropertySets.Item](propertysets.md#item)

#### Samples

| Name | Description |
| --- | --- |
| [Using the BOM APIs](../API_Samples/bom_sample.md) | This sample demonstrates the Bill of Materials API functionality in assemblies. |
| [Update iProperty values using Apprentice](../API_Samples/ipropertyapprentice_sample.md) | Updates some iProperty values using Apprentice.  The document specified in the code for the Open method must exist. |
| [Create custom iProperties](../API_Samples/ipropertycreatecustom_sample.md) | Creates custom iProperties of various types. A document must be open when this sample is run. |
| [Create or update custom iProperty](../API_Samples/ipropertycreateupdatecustom_sample.md) | This example creates a custom iProperty if it doesn't exist and updates the value if it does already exist.  A part document must be open before runnin the sample. |
| [Get value of iProperty](../API_Samples/ipropertygetvalue_sample.md) | Demonstrates getting the values of the "Part Number" iProperty.  Any property can be retrieved by accesing the correct property set and property.A document must be open when this sample is run. |

#### Version

Introduced in version 4

## Methods

<a name="add"></a>
### Add Method

**Parent Object:** [PropertySet](#propertyset)

#### Description

Adds a new Property to this PropertySet.

#### Syntax

PropertySet.Add( PropValue As Variant, [Name] As Variant, [PropId] As Variant ) As [Property](property.md)

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| PropValue | Variant | Input Variant that specifies the value of the Property to add to the set. |
| Name | Variant | Input Variant that specifies the name of the Property. When add a property in a custom property set but not the built-in "User Defined Properties" set(whose internal name is {D5CDD505-2E9C-101B-9397-08002B2CF9AE}), if this name is prefixed with an "_" character, then this property is created as a hidden property and can only be accessed if indexed by its name or propID. PropertySet.Count will not account for such hidden properties. This is an optional argument whose default value is null. |
| PropId | Variant | Input Variant that specifies the PropertyID of the Property to add to the set. Valid propids are 2 through 254 and 256 through 0x80000000. Other values are reserved. This is an optional argument whose default value is null. |

#### Version

Introduced in version 4


<a name="delete"></a>
### Delete Method

**Parent Object:** [PropertySet](#propertyset)

#### Description

Method that deletes this PropertySet.

#### Syntax

PropertySet.Delete()

#### Version

Introduced in version 4


<a name="getpropertyinfo"></a>
### GetPropertyInfo Method

**Parent Object:** [PropertySet](#propertyset)

#### Description

Method that returns property info in the PropertySet.

#### Syntax

PropertySet.GetPropertyInfo( Ids() As Long, Names() As String, Values() As Variant )

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| Ids | Long |  |
| Names | String |  |
| Values | Variant |  |

#### Version

Introduced in version 2018


<a name="setpropertyvalues"></a>
### SetPropertyValues Method

**Parent Object:** [PropertySet](#propertyset)

#### Description

Method that batch sets property values in the PropertySet. If a specified name is not existent in the property set a new property with the specified name will be created.

#### Syntax

PropertySet.SetPropertyValues( PropertyNames() As String, PropertyValues() As Variant )

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| PropertyNames | String | Input String array that indicates the property names. |
| PropertyValues | Variant | Input Variant array that indicates the property values. The sequence of the values are the same as the corresponding PropNames. |

#### Version

Introduced in version 2023


---

## Properties

<a name="count"></a>
### Count Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Property that returns the number of items in this collection.

#### Syntax

PropertySet.Count() As Long

#### Property Value

This is a read only property whose value is a Long.

#### Version

Introduced in version 4


<a name="dirty"></a>
### Dirty Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Property that returns a Boolean flag that indicates whether any of the Properties have been edited, deleted or created.

#### Syntax

PropertySet.Dirty() As Boolean

#### Property Value

This is a read only property whose value is a Boolean.

#### Version

Introduced in version 4


<a name="displayname"></a>
### DisplayName Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Gets/Sets the human-readable name associated with this Property Set.

#### Syntax

PropertySet.DisplayName() As String

#### Property Value

This is a read/write property whose value is a String.

#### Version

Introduced in version 4


<a name="internalname"></a>
### InternalName Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Gets the unambiguous, internal name (FMTID) associated with this PropertySet.

#### Syntax

PropertySet.InternalName() As String

#### Property Value

This is a read only property whose value is a String.

#### Version

Introduced in version 4


<a name="item"></a>
### Item Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Gets the Property given either its name or its sequential index.

#### Syntax

PropertySet.Item( Index As Variant ) As [Property](property.md)

#### Property Value

This is a read only property whose value is a [Property](property.md).

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| Index | Variant | Input Variant value that specifies the index of the Property to return. |

#### Version

Introduced in version 4


<a name="itembypropid"></a>
### ItemByPropId Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Gets the Property in this set by its PropId.

#### Syntax

PropertySet.ItemByPropId( PropId As Long ) As [Property](property.md)

#### Property Value

This is a read only property whose value is a [Property](property.md).

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| PropId | Long | Input Long that specifies the PropertyID of the Property to get from the set. Valid PropIds are 2 through 254 and 256 through 0x80000000. Other values are reserved. |

#### Version

Introduced in version 4


<a name="name"></a>
### Name Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Gets the name of this PropertySet.

#### Syntax

PropertySet.Name() As String

#### Property Value

This is a read only property whose value is a String.

#### Version

Introduced in version 8


<a name="parent"></a>
### Parent Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Property that returns the parent object from whom this object can logically be reached.

#### Syntax

PropertySet.Parent() As [PropertySets](propertysets.md)

#### Property Value

This is a read only property whose value is a [PropertySets](propertysets.md).

#### Version

Introduced in version 4


<a name="type"></a>
### Type Property

**Parent Object:** [PropertySet](#propertyset)

#### Description

Returns an ObjectTypeEnum indicating this object's type.

#### Syntax

PropertySet.Type() As [ObjectTypeEnum](../API_Enums/objecttypeenum.md)

#### Property Value

This is a read only property whose value is an [ObjectTypeEnum](../API_Enums/objecttypeenum.md).

#### Version

Introduced in version 4


---

