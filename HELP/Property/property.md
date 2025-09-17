# Property Object

#### Description

Object that represents a Property. See the article in the overviews section.

#### Methods

| Name | Description |
| --- | --- |
| [Delete](#delete) | Method that deletes this Property from its PropertySet. |

#### Properties

| Name | Description |
| --- | --- |
| [Dirty](#dirty) | Property that returns a Boolean flag that indicates whether this property has been edited or created. |
| [DisplayName](#displayname) | Gets/Sets the human-readable name associated with this Property. |
| [Expression](#expression) | Gets/Sets expression that defines the value of this property. |
| [Name](#name) | Gets the human-readable name of this Property, if any. |
| [Parent](#parent) | Property that returns the parent object from whom this object can logically be reached. |
| [PropId](#propid) | Gets the identifier (PROPID) of this Property. |
| [Type](#type) | Returns an ObjectTypeEnum indicating this object's type. |
| [Value](#value) | Gets/Sets the value of this Property. |

#### Accessed From

- [PropertySet.Add](propertyset.md#add)
- [PropertySet.Item](propertyset.md#item)
- [PropertySet.ItemByPropId](propertyset.md#itembypropid)

#### Samples

| Name | Description |
| --- | --- |
| [Create or update custom iProperty](../API_Samples/ipropertycreateupdatecustom_sample.md) | This example creates a custom iProperty if it doesn't exist and updates the value if it does already exist.  A part document must be open before runnin the sample. |
| [Get value of iProperty](../API_Samples/ipropertygetvalue_sample.md) | Demonstrates getting the values of the "Part Number" iProperty.  Any property can be retrieved by accesing the correct property set and property.A document must be open when this sample is run. |

#### Version

Introduced in version 4

## Methods

<a name="delete"></a>
### Delete Method

**Parent Object:** [Property](#property)

#### Description

Method that deletes this Property from its PropertySet.

#### Syntax

Property.Delete()

#### Version

Introduced in version 4


---

## Properties

<a name="dirty"></a>
### Dirty Property

**Parent Object:** [Property](#property)

#### Description

Property that returns a Boolean flag that indicates whether this property has been edited or created.

#### Syntax

Property.Dirty() As Boolean

#### Property Value

This is a read only property whose value is a Boolean.

#### Version

Introduced in version 4


<a name="displayname"></a>
### DisplayName Property

**Parent Object:** [Property](#property)

#### Description

Gets/Sets the human-readable name associated with this Property.

#### Syntax

Property.DisplayName() As String

#### Property Value

This is a read/write property whose value is a String.

#### Version

Introduced in version 6


<a name="expression"></a>
### Expression Property

**Parent Object:** [Property](#property)

#### Description

Gets/Sets expression that defines the value of this property.

#### Syntax

Property.Expression() As String

#### Property Value

This is a read/write property whose value is a String.

#### Version

Introduced in version 2008


<a name="name"></a>
### Name Property

**Parent Object:** [Property](#property)

#### Description

Gets the human-readable name of this Property, if any.

#### Syntax

Property.Name() As String

#### Property Value

This is a read only property whose value is a String.

#### Version

Introduced in version 4


<a name="parent"></a>
### Parent Property

**Parent Object:** [Property](#property)

#### Description

Property that returns the parent object from whom this object can logically be reached.

#### Syntax

Property.Parent() As [PropertySet](propertyset.md)

#### Property Value

This is a read only property whose value is a [PropertySet](propertyset.md).

#### Version

Introduced in version 4


<a name="propid"></a>
### PropId Property

**Parent Object:** [Property](#property)

#### Description

Gets the identifier (PROPID) of this Property.

#### Syntax

Property.PropId() As Long

#### Property Value

This is a read only property whose value is a Long.

#### Version

Introduced in version 4


<a name="type"></a>
### Type Property

**Parent Object:** [Property](#property)

#### Description

Returns an ObjectTypeEnum indicating this object's type.

#### Syntax

Property.Type() As [ObjectTypeEnum](../API_Enums/objecttypeenum.md)

#### Property Value

This is a read only property whose value is an [ObjectTypeEnum](../API_Enums/objecttypeenum.md).

#### Version

Introduced in version 4


<a name="value"></a>
### Value Property

**Parent Object:** [Property](#property)

#### Description

Gets/Sets the value of this Property.

#### Syntax

Property.Value() As Variant

#### Property Value

This is a read/write property whose value is a Variant.

#### Version

Introduced in version 4


---

