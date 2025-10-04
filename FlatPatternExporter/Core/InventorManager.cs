using System.Runtime.InteropServices;
using System.Windows;
using DefineEdge;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using FlatPatternExporter.UI.Windows;
using Inventor;

namespace FlatPatternExporter.Core;

public class InventorManager
{
    private Inventor.Application? _thisApplication;
    private string _projectName = string.Empty;
    private string _projectWorkspacePath = string.Empty;

    public Inventor.Application? Application => _thisApplication;
    public string ProjectName => _projectName;
    public string ProjectWorkspacePath => _projectWorkspacePath;

    public bool EnsureInventorConnection()
    {
        try
        {
            _thisApplication = (Inventor.Application)MarshalCore.GetActiveObject("Inventor.Application");
            if (_thisApplication != null)
            {
                InitializeProjectData();
                return true;
            }
        }
        catch (COMException)
        {
            CustomMessageBox.Show(
                LocalizationManager.Instance.GetString("Error_InventorConnection"), LocalizationManager.Instance.GetString("MessageBox_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
            _thisApplication = null;
        }
        catch (Exception ex)
        {
            CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_InventorConnection", ex.Message), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK,
                MessageBoxImage.Error);
            _thisApplication = null;
        }

        return false;
    }

    public void InitializeInventor()
    {
        EnsureInventorConnection();
    }

    private void InitializeProjectData()
    {
        if (_thisApplication != null)
        {
            try
            {
                SetProjectFolderInfo();
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_ProjectDataInit", ex.Message), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public void SetProjectFolderInfo()
    {
        if (_thisApplication == null) return;

        try
        {
            var activeProject = _thisApplication.DesignProjectManager.ActiveDesignProject;
            _projectName = activeProject.Name;
            _projectWorkspacePath = activeProject.WorkspacePath;
        }
        catch (Exception ex)
        {
            CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_ProjectInfoGet", ex.Message), LocalizationManager.Instance.GetString("MessageBox_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public DocumentValidationResult ValidateActiveDocument()
    {
        if (!EnsureInventorConnection())
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = LocalizationManager.Instance.GetString("Error_InventorConnectionFailed")
            };
        }

        var doc = _thisApplication?.ActiveDocument;
        if (doc == null)
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = LocalizationManager.Instance.GetString("Error_NoActiveDocument")
            };
        }

        var docType = doc.DocumentType switch
        {
            DocumentTypeEnum.kAssemblyDocumentObject => DocumentType.Assembly,
            DocumentTypeEnum.kPartDocumentObject => DocumentType.Part,
            _ => DocumentType.Invalid
        };

        if (docType == DocumentType.Invalid)
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = LocalizationManager.Instance.GetString("Error_NoDocumentOpen")
            };
        }

        return new DocumentValidationResult
        {
            Document = doc,
            DocType = docType,
            IsValid = true,
            DocumentTypeName = docType == DocumentType.Assembly ? "Assembly" : "Part"
        };
    }

    public void SetInventorUserInterfaceState(bool disableInteraction)
    {
        if (_thisApplication?.UserInterfaceManager != null)
        {
            _thisApplication.UserInterfaceManager.UserInteractionDisabled = disableInteraction;
        }
    }

    public void OpenInventorDocument(string filePath, string? modelState = null)
    {
        if (!System.IO.File.Exists(filePath))
        {
            CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_FileNotFound", filePath), LocalizationManager.Instance.GetString("MessageBox_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            if (!string.IsNullOrEmpty(modelState))
            {
                var pathWithModelState = $"{filePath}<{modelState}>";
                _thisApplication?.Documents?.Open(pathWithModelState);
            }
            else
            {
                _thisApplication?.Documents?.Open(filePath);
            }
        }
        catch (Exception ex)
        {
            CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_FileOpen", filePath, ex.Message), LocalizationManager.Instance.GetString("MessageBox_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public PartDocument? OpenPartDocument(string partNumber)
    {
        var docs = _thisApplication?.Documents;
        if (docs == null) return null;

        foreach (Document doc in docs)
            if (doc is PartDocument pd)
            {
                var mgr = new PropertyManager((Document)pd);
                if (mgr.GetMappedProperty("PartNumber") == partNumber)
                    return pd;
            }

        CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_DocumentNotFound", partNumber), LocalizationManager.Instance.GetString("MessageBox_Error"),
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null;
    }

    public string? GetPartDocumentFullPath(string partNumber)
    {
        var docs = _thisApplication?.Documents;
        if (docs == null) return null;

        foreach (Document doc in docs)
            if (doc is PartDocument pd)
            {
                var mgr = new PropertyManager((Document)pd);
                if (mgr.GetMappedProperty("PartNumber") == partNumber)
                    return pd.FullFileName;
            }

        CustomMessageBox.Show(LocalizationManager.Instance.GetString("Error_DocumentNotFound", partNumber), LocalizationManager.Instance.GetString("MessageBox_Error"),
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null;
    }

    public bool IsLibraryComponent(string fullFileName)
    {
        try
        {
            if (_thisApplication?.DesignProjectManager == null)
                return false;

            _thisApplication.DesignProjectManager.IsFileInActiveProject(
                fullFileName,
                out var projectPathType,
                out _);

            return projectPathType == LocationTypeEnum.kLibraryLocation;
        }
        catch
        {
            return false;
        }
    }
}