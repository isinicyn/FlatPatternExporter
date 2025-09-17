# VirtualComponentDefinition Object

**Derived from:** [ComponentDefinition](componentdefinition.md)

#### Description

This object derives from the ComponentDefinition object. It represents a ComponentDefinition that exists solely for the BOM.

#### Methods

| Name | Description |
| --- | --- |
| [FindUsingPoint](#findusingpoint) | Method that finds all the entities of the specified type at the specified location. |
| [FindUsingVector](#findusingvector) | Method that finds all the entities of the specified type along the specified vector using either a cylinder or cone that to define the tolerance within the defined vector. |
| [GetUnusedGeometries](#getunusedgeometries) | Method that gets the unused sketches and work features. |
| [PurgeUnusedGeometries](#purgeunusedgeometries) | Method that purges unused sketches and work features. |
| [RepositionObject](#repositionobject) | Method that repositions the specifies object(s) to the new position within the collection of the object in the document. |

#### Properties

| Name | Description |
| --- | --- |
| [ActiveMaterial](#activematerial) | Gets and sets the material for the VirtualComponentDefinition. |
| [Application](#application) | Returns the top-level parent application object.  When used the context of Inventor, an Application object is returned.  When used in the context of Apprentice, an ApprenticeServer object is returned. |
| [AttributeSets](#attributesets) | Property that returns the AttributeSets collection object associated with this object. |
| [BOMQuantity](#bomquantity) | Property that returns the BOMQuantity object. |
| [BOMStructure](#bomstructure) | Gets and sets how the component is used/viewed in a BOM. |
| [ClientGraphicsCollection](#clientgraphicscollection) | Property that returns the ClientGraphicsCollection object. |
| [DataIO](#dataio) | Gets the object that directly deals with I/O to and from a storage-medium, including Streams(IStream). |
| [DisplayName](#displayname) | Gets and sets the name of the virtual component. |
| [Document](#document) | Property that returns the containing Document object. |
| [ModelGeometryVersion](#modelgeometryversion) | Property that returns a string that can be used to determine if the document has been modified. This version string is changed every time the assembly is modified. By saving a previous version string, you can compare the current version string to see if the assembly has been modified. |
| [Occurrences](#occurrences) | Property that returns the collection object. |
| [OrientedMinimumRangeBox](#orientedminimumrangebox) | Read-only property that returns the oriented minimum range box for this object. |
| [PreciseRangeBox](#preciserangebox) | Gets a bounding box that tightly encloses all the solid and surface bodies under the ComponentDefinition. |
| [PropertySets](#propertysets) | Property that gets the PropertySets object associated with the virtual component. |
| [RangeBox](#rangebox) | Property that returns a Box object which contains the opposing points of a rectangular box that is guaranteed to enclose this object. |
| [SurfaceBodies](#surfacebodies) | Property that returns all of theresultSurfaceBody objects contained within this ComponentDefinition. |
| [Type](#type) | Returns an ObjectTypeEnum indicating this object's type. |

#### Samples

| Name | Description |
| --- | --- |
| [Using the BOM APIs](../API_Samples/bom_sample.md) | This sample demonstrates the Bill of Materials API functionality in assemblies. |

#### Version

Introduced in version 10

## Methods

<a name="findusingpoint"></a>
### FindUsingPoint Method

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Method that finds all the entities of the specified type at the specified location.

#### Syntax

VirtualComponentDefinition.FindUsingPoint( Point As [Point](point.md), ObjectTypes() As [SelectionFilterEnum](../API_Enums/selectionfilterenum.md), [ProximityTolerance] As Variant, [VisibleObjectsOnly] As Boolean ) As [ObjectsEnumerator](objectsenumerator.md)

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| Point | [Point](point.md) | Input Point object that specifies the model space point at which to find the entities. |
| ObjectTypes | [SelectionFilterEnum](../API_Enums/selectionfilterenum.md) | Input array of SelelctionFilterEnum values that indicate the type of objects to find. The values are the enum values from the SelectionFilterEnum and can be combined to specify multiple object types. |
| ProximityTolerance | Variant | Input Double that specifies the tolerance value for the search. This value defines the radius of a sphere at the input point. All objects that intersect this sphere will be returned. If not specified, the default internal tolerance is used. This is an optional argument whose default value is null. |
| VisibleObjectsOnly | Boolean | Optional input Boolean that indicates whether or not invisible objects should be included in the search. Defaults to True indicating that invisible objects will be ignored. This is an optional argument whose default value is True. |

#### Version

Introduced in version 2009


<a name="findusingvector"></a>
### FindUsingVector Method

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Method that finds all the entities of the specified type along the specified vector using either a cylinder or cone that to define the tolerance within the defined vector.

#### Syntax

VirtualComponentDefinition.FindUsingVector( OriginPoint As [Point](point.md), Direction As [UnitVector](unitvector.md), ObjectTypes() As [SelectionFilterEnum](../API_Enums/selectionfilterenum.md), [UseCylinder] As Boolean, [ProximityTolerance] As Variant, [VisibleObjectsOnly] As Boolean, [LocationPoints] As Variant ) As [ObjectsEnumerator](objectsenumerator.md)

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| OriginPoint | [Point](point.md) | Input Point that defines the model space location to position the vector. |
| Direction | [UnitVector](unitvector.md) | Input UnitVector that defines direction to find all entities that the vector penetrates. The vector is treated as infinite in both directions from the origin point so all entities that lie in the path of the vector on either side of the origin point will be returned. |
| ObjectTypes | [SelectionFilterEnum](../API_Enums/selectionfilterenum.md) | Input array of SelelctionFilterEnum values that indicate the type of objects to find. The values are the enum values from the SelectionFilterEnum and can be combined to specify multiple object types. |
| UseCylinder | Boolean | Input argument that specifies if the vector defines a cylinder or cone when checking to see if an entity is hit by the ray. This is an optional argument whose default value is True. |
| ProximityTolerance | Variant | Optional input Double that specifies the tolerance value for the search. This value defines the radius of a cylinder if UseCylinder is True or the angle of the cone if False. If not specified, the default internal tolerance is used. This is an optional argument whose default value is null. |
| VisibleObjectsOnly | Boolean | Optional input Boolean that indicates whether or not invisible objects should be included in the search. Defaults to True indicating that invisible objects will be ignored. This is an optional argument whose default value is True. |
| LocationPoints | Variant | Optional output that returns a set of Point objects. One Point object is returned for each entity in the FoundEntities list. The point defines the coordinate of the intersection between the vector and the corresponding entity. This is an optional argument whose default value is null. |

#### Version

Introduced in version 2011


<a name="getunusedgeometries"></a>
### GetUnusedGeometries Method

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Method that gets the unused sketches and work features.

#### Syntax

VirtualComponentDefinition.GetUnusedGeometries( UnusedGeometries As [ObjectCollection](objectcollection.md) )

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| UnusedGeometries | [ObjectCollection](objectcollection.md) | Output ObjectCollection object the includes the unused sketches and work features. |

#### Version

Introduced in version 2024


<a name="purgeunusedgeometries"></a>
### PurgeUnusedGeometries Method

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Method that purges unused sketches and work features.

#### Syntax

VirtualComponentDefinition.PurgeUnusedGeometries( [UnusedGeometries] As Variant )

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| UnusedGeometries | Variant | Optional input ObjectCollection that including the sketches and work features to purge. The unused sketches and work features can be retrieved using GetUnusedGeometries method. If this is not provided then all unused sketches and work features will be purged. This is an optional argument whose default value is null. |

#### Version

Introduced in version 2024


<a name="repositionobject"></a>
### RepositionObject Method

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Method that repositions the specifies object(s) to the new position within the collection of the object in the document.

#### Syntax

VirtualComponentDefinition.RepositionObject( TargetObject As Object, Before As Boolean, StartObject As Object, [EndObject] As Variant )

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| TargetObject | Object | Input the Object that specifies the target object to move other objects next to. Valid object includes: PartFeature, ComponentOccurrence, Sketch, Sketch3D, WorkFeature. |
| Before | Boolean | Input Boolean that indicates whether to position other object(s) before or after the target object.  A value of True indicates that the object(s) will be positioned before the target object. |
| StartObject | Object | Input Object that specifies the object to be repositioned.  Valid object includes: PartFeature, ComponentOccurrence, Sketch, Sketch3D, WorkFeature. |
| EndObject | Variant | Optional input Object that specifies the object to be repositioned.  If specified, all the objects from the StartObject to the EndObject, both inclusive, will be repositioned to the specified position in the document.  If not specified, only the StartObject will be repositioned. Valid object includes: PartFeature, ComponentOccurrence, Sketch, Sketch3D, WorkFeature. This is an optional argument whose default value is null. |

#### Version

Introduced in version 2025


---

## Properties

<a name="activematerial"></a>
### ActiveMaterial Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Gets and sets the material for the VirtualComponentDefinition.

#### Syntax

VirtualComponentDefinition.ActiveMaterial() As [Asset](asset.md)

#### Property Value

This is a read/write property whose value is an [Asset](asset.md).

#### Version

Introduced in version 2024


<a name="application"></a>
### Application Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Returns the top-level parent application object. When used the context of Inventor, an Application object is returned. When used in the context of Apprentice, an ApprenticeServer object is returned.

#### Syntax

VirtualComponentDefinition.Application() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 2010


<a name="attributesets"></a>
### AttributeSets Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns the AttributeSets collection object associated with this object.

#### Syntax

VirtualComponentDefinition.AttributeSets() As [AttributeSets](attributesets.md)

#### Property Value

This is a read only property whose value is an [AttributeSets](attributesets.md).

#### Version

Introduced in version 10


<a name="bomquantity"></a>
### BOMQuantity Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns the BOMQuantity object.

#### Syntax

VirtualComponentDefinition.BOMQuantity() As [BOMQuantity](bomquantity.md)

#### Property Value

This is a read only property whose value is a [BOMQuantity](bomquantity.md).

#### Version

Introduced in version 10


<a name="bomstructure"></a>
### BOMStructure Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Gets and sets how the component is used/viewed in a BOM.

#### Syntax

VirtualComponentDefinition.BOMStructure() As [BOMStructureEnum](../API_Enums/bomstructureenum.md)

#### Property Value

This is a read/write property whose value is a [BOMStructureEnum](../API_Enums/bomstructureenum.md).

#### Version

Introduced in version 10


<a name="clientgraphicscollection"></a>
### ClientGraphicsCollection Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns the ClientGraphicsCollection object.

#### Syntax

VirtualComponentDefinition.ClientGraphicsCollection() As [ClientGraphicsCollection](clientgraphicscollection.md)

#### Property Value

This is a read only property whose value is a [ClientGraphicsCollection](clientgraphicscollection.md).

#### Version

Introduced in version 10


<a name="dataio"></a>
### DataIO Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Gets the object that directly deals with I/O to and from a storage-medium, including Streams(IStream).

#### Syntax

VirtualComponentDefinition.DataIO() As [DataIO](dataio.md)

#### Property Value

This is a read only property whose value is a [DataIO](dataio.md).

#### Version

Introduced in version 10


<a name="displayname"></a>
### DisplayName Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Gets and sets the name of the virtual component.

#### Syntax

VirtualComponentDefinition.DisplayName() As String

#### Property Value

This is a read/write property whose value is a String.

#### Version

Introduced in version 10


<a name="document"></a>
### Document Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns the containing Document object.

#### Syntax

VirtualComponentDefinition.Document() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 10


<a name="modelgeometryversion"></a>
### ModelGeometryVersion Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns a string that can be used to determine if the document has been modified. This version string is changed every time the assembly is modified. By saving a previous version string, you can compare the current version string to see if the assembly has been modified.

#### Syntax

VirtualComponentDefinition.ModelGeometryVersion() As String

#### Property Value

This is a read only property whose value is a String.

#### Version

Introduced in version 10


<a name="occurrences"></a>
### Occurrences Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns the collection object.

#### Syntax

VirtualComponentDefinition.Occurrences() As [ComponentOccurrences](componentoccurrences.md)

#### Property Value

This is a read only property whose value is a [ComponentOccurrences](componentoccurrences.md).

#### Version

Introduced in version 10


<a name="orientedminimumrangebox"></a>
### OrientedMinimumRangeBox Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Read-only property that returns the oriented minimum range box for this object.

#### Syntax

VirtualComponentDefinition.OrientedMinimumRangeBox() As [OrientedBox](orientedbox.md)

#### Property Value

This is a read only property whose value is an [OrientedBox](orientedbox.md).

#### Version

Introduced in version 2024


<a name="preciserangebox"></a>
### PreciseRangeBox Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Gets a bounding box that tightly encloses all the solid and surface bodies under the ComponentDefinition.

#### Syntax

VirtualComponentDefinition.PreciseRangeBox() As [Box](box.md)

#### Property Value

This is a read only property whose value is a [Box](box.md).

#### Version

Introduced in version 2023


<a name="propertysets"></a>
### PropertySets Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that gets the PropertySets object associated with the virtual component.

#### Syntax

VirtualComponentDefinition.PropertySets() As [PropertySets](propertysets.md)

#### Property Value

This is a read only property whose value is a [PropertySets](propertysets.md).

#### Version

Introduced in version 10


<a name="rangebox"></a>
### RangeBox Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns a Box object which contains the opposing points of a rectangular box that is guaranteed to enclose this object.

#### Syntax

VirtualComponentDefinition.RangeBox() As [Box](box.md)

#### Property Value

This is a read only property whose value is a [Box](box.md).

#### Version

Introduced in version 11


<a name="surfacebodies"></a>
### SurfaceBodies Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Property that returns all of the result SurfaceBody objects contained within this ComponentDefinition.

#### Syntax

VirtualComponentDefinition.SurfaceBodies() As [SurfaceBodies](surfacebodies.md)

#### Property Value

This is a read only property whose value is a [SurfaceBodies](surfacebodies.md).

#### Version

Introduced in version 10


<a name="type"></a>
### Type Property

**Parent Object:** [VirtualComponentDefinition](#virtualcomponentdefinition)

#### Description

Returns an ObjectTypeEnum indicating this object's type.

#### Syntax

VirtualComponentDefinition.Type() As [ObjectTypeEnum](../API_Enums/objecttypeenum.md)

#### Property Value

This is a read only property whose value is an [ObjectTypeEnum](../API_Enums/objecttypeenum.md).

#### Version

Introduced in version 10


---

