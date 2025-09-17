# PropertiesForSummaryInformationEnum Enumerator

#### Description

Summary Information, InternalName: {F29F85E0-4FF9-1068-AB91-08002B27B3D9}, Name: Inventor Summary Information - This Enum contains values that correspond with the ID's of the Property objects within the Summary Information PropertySet object.

#### Enumeration Values

| Name | Value | Description |
| --- | --- | --- |
| kAuthorSummaryInformation | 4 | Author: String (VT_STR), Editable, UI on Summary Page, Always present (default=current login name), Creator of this file. |
| kCommentsSummaryInformation | 6 | Comments: String (VT_STR), Editable, UI on Summary Page, Always present (default=empty string), Additional information. |
| kCreationTimeSummaryInformation | 12 | Creation Time: Date (VT_FILETIME), Read Only, No UI, Always present (default=file creation time), Time this file came into existence. *** DEFUNCT. Use Creation Date in Design Tracking Property Set instead. |
| kKeywordsSummaryInformation | 5 | Keywords: String (VT_STR), Editable, UI on Summary Page, Always present (default=empty string), Distinctive words that best describe this file. |
| kLastSavedBySummaryInformation | 8 | Last Saved By: String (VT_STR), Editable, UI on Summary Page, Always present (default=empty string), Last saved by author. |
| kRevisionSummaryInformation | 9 | Revision: String (VT_STR), Editable, UI on Project Page, Always present (default=empty string), Number denoting the current version of the document. |
| kSubjectSummaryInformation | 3 | Subject: String (VT_STR), Editable, UI on Summary Page, Always present (default=empty string), Topic to which this file pertains. |
| kThumbnailSummaryInformation | 17 | Thumbnail: Bytes (VT_CLIPDATA), Editable, Read Only UI on File Dialogs, Always present (default=empty bitmap), A bitmap picture of the document. |
| kTitleSummaryInformation | 2 | Title: String (VT_STR), Editable, UI on Summary Page, Always present (default=empty string), Heading for this file. |

#### Version

Introduced in version 4

