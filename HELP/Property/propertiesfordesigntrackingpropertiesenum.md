# PropertiesForDesignTrackingPropertiesEnum Enumerator

#### Description

Design Tracking Properties, InternalName: {32853F0F-3444-11d1-9E93-0060B03C1CA6}, Name: Design Tracking Properties - This Enum contains values that correspond with the ID's of the Property objects within the Design Tracking Properties PropertySet object.

#### Enumeration Values

| Name | Value | Description |
| --- | --- | --- |
| kAppearanceDesignTrackingProperties | 72 | Appearance : String (VT_BSTR),Read Only, No UI, Need not be present, Appearance. |
| kAuthorityDesignTrackingProperties | 43 | Authority: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), If non-NULL, this signifies the authorization of this model. |
| kCatalogWebLinkDesignTrackingProperties | 23 | Catalog Web Link: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), If this is a Part from a Catalog, the Web address of the Catalog provider. |
| kCategoriesDesignTrackingProperties | 56 | Categories: String (VT_BSTR), Read Only, No UI, XML string providing the complete categorization of this model, along with the mappings of the Category Parameters to parameters within this Document. |
| kCheckedByDesignTrackingProperties | 10 | Checked By: String (VT_BSTR), Editable, UI on Status Page, Always present (default=empty string), Name of the person checking the Part. |
| kCostCenterDesignTrackingProperties | 9 | Cost Center: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), Cost Center responsible for the Project. |
| kCostDesignTrackingProperties | 36 | Cost: Currency (VT_CY), Editable, UI on Project Page, Always present (default=0), Cost of the Part. |
| kCreationDateDesignTrackingProperties | 4 | Creation Time: Date (VT_FILETIME), Read Only, UI on Project Page, Always present (default=data creation time), Time the data in this file was first created. Will remain the same after Save As or Copy operations. |
| kDateCheckedDesignTrackingProperties | 11 | Date Checked: Date (VT_FILETIME), Editable, UI on Status Page, Always present (default=0), Time on which the check was performed. |
| kDateEngrApprovedDesignTrackingProperties | 13 | Engr Date Approved: Date (VT_FILETIME), Editable, UI on Status Page, Always present (default=0), Time on which the approval of the Part's engineering was performed. |
| kDateMfgApprovedDesignTrackingProperties | 35 | Mfg Date Approved: Date (VT_FILETIME), Editable, UI on Project Page, Always present (default=0), Time on which the approval of the Part's manufacturing was performed. |
| kDensityDesignTrackingProperties | 61 | Density: Double (VT_R4), Read Only, No UI, Always present (default=-1), Density. |
| kDescriptionDesignTrackingProperties | 29 | Description: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), A user-defined description of the contents of this file. |
| kDesignerDesignTrackingProperties | 41 | Designer: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), If non-NULL, this is the name of the designer of this model. |
| kDesignStatusDesignTrackingProperties | 40 | Design Status: Long (VT_I4), Editable, UI on Status Page, Always present (default=0), 0=Not Set, 1=Work In Progress, 2=Pending, 3=Released. |
| kDocSubTypeDesignTrackingProperties | 31 | Document SubType: GUID (VT_CLSID), Read Only, No UI, Always present (default=Document's CLSID), Finer classification of this Document's type. Document's type can be retrived by calling ::ReadClassStg on the root-storage. |
| kDocSubTypeNameDesignTrackingProperties | 32 | Document SubType Name: String (VT_BSTR), Read Only, No UI, Always present (default=Document Type name), Display name of the finer classification of this Document's type. |
| kDrawingDeferUpdateDesignTrackingProperties | 51 | Defer Updates: Boolean (VT_BOOL), Editable, No UI, Always present?? (default=FALSE), Whether to suspend drawing updates that reflect changes in its referenced files. |
| kEngineerDesignTrackingProperties | 42 | Engineer: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), If non-NULL, this is the name of the engineer responsible for this model. |
| kEngrApprovedByDesignTrackingProperties | 12 | Engr Approved By: String (VT_BSTR), Editable, UI on Status Page, Always present (default=empty string), Name of the person approving the Part's engineering aspect. |
| kExternalPropRevIdDesignTrackingProperties | 46 | External Property Revision Id: GUID (VT_CLSID), Editable, No UI, Always present (default=initial GUID), GUID representing revision of external property sets (created by Third Parties). |
| kFlatPatternDeferUpdateDesignTrackingProperties | 73 | Flat Pattern Defer Update: Boolean (VT_BOOL), Editable, No UI, Always present?? (default=FALSE), Whether to suspend flat pattern updates. |
| kFlatPatternExtentsAreaDesignTrackingProperties | 65 | Flat Pattern Extents Area: Double (VT_R4), Read Only, No UI, Need not be present, Flat Pattern Extents Area. |
| kFlatPatternExtentsLengthDesignTrackingProperties | 64 | Flat Pattern Extents Length: Double (VT_R4), Read Only, No UI, Need not be present, Flat Pattern Extents Length. |
| kFlatPatternExtentsWidthDesignTrackingProperties | 63 | Flat Pattern Extents Width: Double (VT_R4), Read Only, No UI, Need not be present, Flat Pattern Extents Width. |
| kLanguageDesignTrackingProperties | 50 | Language: String (VT_BSTR), Read Only, No UI, Always present (default=empty string), If non-NULL, this is the language (specified in ISO 639-2 alpha-3) being used for displaying the user-friendly names of various pieces of data. |
| kLastUpdatedWithDesignTrackingProperties | 67 | Last Updated With: string (VT_BSTR),Read Only, UI on Detail Page, Always present (default=empty string), Last updated with which Inventor version. |
| kManufacturerDesignTrackingProperties | 48 | Manufacturer: String (VT_BSTR), Editable, No UI, Always present (default=Standard), This is the name of the manufacturer or supplier of this component. |
| kMassDesignTrackingProperties | 58 | Mass: Double (VT_R4), Read Only, No UI, Always present (default=-1), Mass. |
| kMaterialDesignTrackingProperties | 20 | Material: String (VT_BSTR), Read Only, UI in Material command, Always present (default=empty string), Name of the material used to build the Part. |
| kMaterialIdentifierDesignTrackingProperties | 71 | Material identifier : String (VT_BSTR),Read Only, No UI, Need not be present, Material Identifier. |
| kMfgApprovedByDesignTrackingProperties | 34 | Mfg Approved By: String (VT_BSTR), Editable, UI on Project Page, Need not be present, Name of the person approving the Part's manufacturing aspect. |
| kParameterizedTemplateDesignTrackingProperties | 44 | Parameterized Template: Bool (VT_BOOL), Read Only, No UI, Always present (default=FALSE), Whether this model is a parameterized template to produce other models. |
| kPartIconDesignTrackingProperties | 28 | Part Icon: Clipdata (VT_CLIPDATA), Read Only, No UI, Need not be present, Special icon to be used wherever appropriate (eg: Windows Explorer). |
| kPartNumberDesignTrackingProperties | 5 | Part Number: String (VT_BSTR), Editable, UI on Project Page, Always present (default=Filename), Tracking number assigned to the Part. |
| kPartPropRevIdDesignTrackingProperties | 21 | Part Property Revision Id: GUID (VT_CLSID), Editable, No UI, Always present (default=initial GUID), GUID representing revision of Part properties (SummaryInformation, DocumentSummaryInformation, UserDefinedProperties and DesignTrackingProperties). |
| kProjectDesignTrackingProperties | 7 | Project: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), Project for which this Part was generated. |
| kProxyRefreshDateDesignTrackingProperties | 33 | Proxy Refresh Date: Date (VT_FILETIME), Read Only, No UI, Always present (default=0), If this is Proxy file for a Part, the date/time when it was last refreshed with respect to the Part. |
| kRevitCategoryDesignTrackingProperties | 74 | Revit Category: Long (VT_I4), Editable, UI on Project Page, Always present (default=0), 0=. |
| kSheetMetalRuleDesignTrackingProperties | 66 | Sheet Metal Rule : String (VT_BSTR),Read Only, No UI, Always present (default=empty string), Name of the Sheet Metal Rule used for Sheet Metal. |
| kStandardDesignTrackingProperties | 37 | Standard: String (VT_BSTR), Editable, No UI, Always present (default=empty string), If non-NULL, this is the specific standard that governs this component family. |
| kStandardRevisionDesignTrackingProperties | 47 | Standard Revision: String (VT_BSTR), Editable, No UI, Always present (default=empty string), Applicable when Standard is non-NULL. Indicates the particular version of the international standard being applied to the component. |
| kStandardsOrganizationDesignTrackingProperties | 49 | Standards Organization: String (VT_BSTR), Editable, No UI, Always present (default=Custom), This is the name of the organization that developed the specific standard to which this family might belong. |
| kStockNumberDesignTrackingProperties | 55 | Stock Number: String (VT_BSTR), Editable, UI on Project Page, Always present (default=Filename), Stock number assigned to the Part. |
| kSurfaceAreaDesignTrackingProperties | 59 | SurfaceArea: Double (VT_R4), Read Only, No UI, Always present (default=-1), SurfaceArea. |
| kTemplateRowDesignTrackingProperties | 45 | Template Row: GUID (VT_CLSID), Read Only, No UI, Need not be present, The row of parameters in the template from which this model was derived. |
| kUserStatusDesignTrackingProperties | 17 | User Status: String (VT_BSTR), Editable, UI on Status Page, Always present (default=empty string), Status of the Part. |
| kValidMassPropsDesignTrackingProperties | 62 | Valid MassProps: Long (VT_I4), Read Only, No UI, Always present (default=0), BitFlags indicating which MassProperties are Valid/Invalid. |
| kVendorDesignTrackingProperties | 30 | Vendor: String (VT_BSTR), Editable, UI on Project Page, Always present (default=empty string), Name of the vendor from which this Part is purchased. |
| kVolumeDesignTrackingProperties | 60 | Volume: Double (VT_R4), Read Only, No UI, Always present (default=-1), Volume. |
| kWeldMaterialDesignTrackingProperties | 57 | Weld Material: String (VT_BSTR), Read Only, No UI, Always present (default=empty string), Name of the material used for Weld Beads. |

#### Version

Introduced in version 4

