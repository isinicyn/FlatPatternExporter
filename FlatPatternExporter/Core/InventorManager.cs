using System.Runtime.InteropServices;
using System.Windows;
using DefineEdge;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using Inventor;
using MessageBox = System.Windows.MessageBox;

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
            MessageBox.Show(
                "Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            _thisApplication = null;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Произошла ошибка при подключении к Inventor: " + ex.Message, "Ошибка", MessageBoxButton.OK,
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
                MessageBox.Show($"Ошибка при инициализации данных проекта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
            MessageBox.Show($"Не удалось получить информацию о проекте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public DocumentValidationResult ValidateActiveDocument()
    {
        if (!EnsureInventorConnection())
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = "Не удалось подключиться к Inventor"
            };
        }

        var doc = _thisApplication?.ActiveDocument;
        if (doc == null)
        {
            return new DocumentValidationResult
            {
                IsValid = false,
                ErrorMessage = "Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова."
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
                ErrorMessage = "Откройте сборку или деталь для работы с приложением."
            };
        }

        return new DocumentValidationResult
        {
            Document = doc,
            DocType = docType,
            IsValid = true,
            DocumentTypeName = docType == DocumentType.Assembly ? "Сборка" : "Деталь"
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
            MessageBox.Show($"Файл по пути {filePath} не найден.", "Ошибка",
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
            MessageBox.Show($"Ошибка при открытии файла по пути {filePath}: {ex.Message}", "Ошибка",
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

        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
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

        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
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