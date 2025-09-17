# DesignProjectManager Object

#### Description

The DesignProjectManager object provides access to project files related functionality in Inventor.

#### Methods

| Name | Description |
| --- | --- |
| [AddOptionsButton](#addoptionsbutton) | Method that adds an options button to the Projects dialog. The returned button object can be used to receive an OnClick event fired when the user clicks the option button. |
| [IsFileInActiveProject](#isfileinactiveproject) | Method that returns whether the given file is located within the active project using the resolution rules of the project, and additionally returns the path type (library, workspace, workgroup) and its name. |
| [ResolveFile](#resolvefile) | Method that runs the file resolver from the source path and attempts to find the destination file name, in the active project. The full file name of the resolved file is returned. A null string is returned if no file was resolved to. |

#### Properties

| Name | Description |
| --- | --- |
| [ActiveDesignProject](#activedesignproject) | Property that returns the currently active design project. Use DesignProject.Activate method to activate a project. |
| [Application](#application) | Returns the top-level parent application object.  When used the context of Inventor, an Application object is returned.  When used in the context of Apprentice, an ApprenticeServer object is returned. |
| [DesignProjects](#designprojects) | Property that returns the DesignProjects collection object containing all the projects. |
| [Parent](#parent) | Property that returns the parent Application or ApprenticeServerComponent object. |
| [Type](#type) | Returns an ObjectTypeEnum indicating this object's type. |

#### Accessed From

- [Application.DesignProjectManager](application.md#designprojectmanager)
- [ApprenticeServer.DesignProjectManager](apprenticeserver.md#designprojectmanager)
- [ApprenticeServerComponent.DesignProjectManager](apprenticeservercomponent.md#designprojectmanager)
- [DesignProject.Parent](designproject.md#parent)
- [InventorServer.DesignProjectManager](InventorServer_DesignProjectManager.htm)
- [InventorServerObject.DesignProjectManager](InventorServerObject_DesignProjectManager.htm)

#### Samples

| Name | Description |
| --- | --- |
| [Set active project](../API_Samples/projectactivate_sample.md) | The following sample demonstrates the activation of an Inventor project. |
| [Create project](../API_Samples/projectcreation_sample.md) | The following sample demonstrates the creation of an Inventor project. |
| [Query and create library paths](../API_Samples/projectlibrarypaths_sample.md) | The following sample demonstrates querying existing library paths associated with a project and adding a new library path. |

#### Version

Introduced in version 2011

## Methods

<a name="addoptionsbutton"></a>
### AddOptionsButton Method

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Method that adds an options button to the Projects dialog. The returned button object can be used to receive an OnClick event fired when the user clicks the option button.

#### Syntax

DesignProjectManager.AddOptionsButton( ClientId As String, DisplayName As String, ToolTipText As String, [StandardIcon] As Variant ) As [ProjectOptionsButton](projectoptionsbutton.md)

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| ClientId | String | Input string that uniquely identifies the client. |
| DisplayName | String | Input string that specifies the dispaly name of the control. |
| ToolTipText | String | Input string that specifies the tooltip text. |
| StandardIcon | Variant | Optional input Picture (IPictureDisp) object that specifies the icon to use for the control. This is an optional argument whose default value is null. |

#### Version

Introduced in version 2012


<a name="isfileinactiveproject"></a>
### IsFileInActiveProject Method

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Method that returns whether the given file is located within the active project using the resolution rules of the project, and additionally returns the path type (library, workspace, workgroup) and its name.

#### Syntax

DesignProjectManager.IsFileInActiveProject( FileName As String, ProjectPathType As [LocationTypeEnum](../API_Enums/locationtypeenum.md), ProjectPathName As String ) As Boolean

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| FileName | String | Input String that specifies the name of a file. This can either be a full file name (recommended), a relative file name or a file name with no path. |
| ProjectPathType | [LocationTypeEnum](../API_Enums/locationtypeenum.md) | Output LocationTypeEnum that returns where the input file was found. Possible returns values are: kLibraryLocation, kWorkspaceLocation, kWorkgroupLocation, kUnknownLocation (if none of the above or if the file is not found within the project). |
| ProjectPathName | String | Output String that returns the name of the library, workspace or workgroup that the file was found in. Returns a null string if none of the above. |

#### Version

Introduced in version 2011


<a name="resolvefile"></a>
### ResolveFile Method

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Method that runs the file resolver from the source path and attempts to find the destination file name, in the active project. The full file name of the resolved file is returned. A null string is returned if no file was resolved to.

#### Syntax

DesignProjectManager.ResolveFile( SourcePath As String, DestinationFileName As String, [Options] As Variant ) As String

#### Parameters

| Name | Type | Description |
| --- | --- | --- |
| SourcePath | String | Input String that specifies the source path to start the file resolution from. |
| DestinationFileName | String | Input String that specifies the destination file name to resolve to. This can either be a relative file name or a full file name. |
| Options | Variant | Optional input Variant reserved for future use. Currently ignored. This is an optional argument whose default value is null. |

#### Version

Introduced in version 2011


---

## Properties

<a name="activedesignproject"></a>
### ActiveDesignProject Property

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Property that returns the currently active design project. Use DesignProject.Activate method to activate a project.

#### Syntax

DesignProjectManager.ActiveDesignProject() As [DesignProject](designproject.md)

#### Property Value

This is a read only property whose value is a [DesignProject](designproject.md).

#### Version

Introduced in version 2011


<a name="application"></a>
### Application Property

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Returns the top-level parent application object. When used the context of Inventor, an Application object is returned. When used in the context of Apprentice, an ApprenticeServer object is returned.

#### Syntax

DesignProjectManager.Application() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 2011


<a name="designprojects"></a>
### DesignProjects Property

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Property that returns the DesignProjects collection object containing all the projects.

#### Syntax

DesignProjectManager.DesignProjects() As [DesignProjects](designprojects.md)

#### Property Value

This is a read only property whose value is a [DesignProjects](designprojects.md).

#### Version

Introduced in version 2011


<a name="parent"></a>
### Parent Property

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Property that returns the parent Application or ApprenticeServerComponent object.

#### Syntax

DesignProjectManager.Parent() As Object

#### Property Value

This is a read only property whose value is an Object.

#### Version

Introduced in version 2011


<a name="type"></a>
### Type Property

**Parent Object:** [DesignProjectManager](#designprojectmanager)

#### Description

Returns an ObjectTypeEnum indicating this object's type.

#### Syntax

DesignProjectManager.Type() As [ObjectTypeEnum](../API_Enums/objecttypeenum.md)

#### Property Value

This is a read only property whose value is an [ObjectTypeEnum](../API_Enums/objecttypeenum.md).

#### Version

Introduced in version 2011


---

