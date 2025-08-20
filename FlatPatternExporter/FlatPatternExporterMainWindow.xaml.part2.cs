using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using DxfGenerator;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Inventor;
using Binding = System.Windows.Data.Binding;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Style = System.Windows.Style;

namespace FlatPatternExporter;

public partial class FlatPatternExporterMainWindow : Window
{
    /// <summary>
    /// Общий метод для чтения всех свойств из документа с использованием PropertyManager
    /// </summary>
    private void ReadAllPropertiesFromPart(PartDocument partDoc, PartData partData, PropertyManager mgr)
    {
        // КАТЕГОРИЯ 2: Свойства документа (не iProperty)
        partData.FileName = mgr.GetFileName();
        partData.FullFileName = mgr.GetFullFileName();
        partData.ModelState = mgr.GetModelState();
        partData.Thickness = mgr.GetThickness();
        partData.HasFlatPattern = mgr.HasFlatPattern();

        // Всегда обязательные свойства (независимо от наличия колонок)
        partData.PartNumber = mgr.GetMappedProperty("PartNumber");
        partData.PartNumberIsExpression = mgr.IsMappedPropertyExpression("PartNumber");
        partData.Description = mgr.GetMappedProperty("Description");
        partData.Material = mgr.GetMappedProperty("Material");

        // КАТЕГОРИЯ 4: Расширенные iProperty (загружаем все)
        partData.Author = mgr.GetMappedProperty("Author");
        partData.Revision = mgr.GetMappedProperty("Revision");
        partData.Title = mgr.GetMappedProperty("Title");
        partData.Subject = mgr.GetMappedProperty("Subject");
        partData.Keywords = mgr.GetMappedProperty("Keywords");
        partData.Comments = mgr.GetMappedProperty("Comments");
        partData.Category = mgr.GetMappedProperty("Category");
        partData.Manager = mgr.GetMappedProperty("Manager");
        partData.Company = mgr.GetMappedProperty("Company");
        partData.Project = mgr.GetMappedProperty("Project");
        partData.StockNumber = mgr.GetMappedProperty("StockNumber");
        partData.CreationTime = mgr.GetMappedProperty("CreationTime");
        partData.CostCenter = mgr.GetMappedProperty("CostCenter");
        partData.CheckedBy = mgr.GetMappedProperty("CheckedBy");
        partData.EngApprovedBy = mgr.GetMappedProperty("EngApprovedBy");
        partData.UserStatus = mgr.GetMappedProperty("UserStatus");
        partData.CatalogWebLink = mgr.GetMappedProperty("CatalogWebLink");
        partData.Vendor = mgr.GetMappedProperty("Vendor");
        partData.MfgApprovedBy = mgr.GetMappedProperty("MfgApprovedBy");
        partData.DesignStatus = mgr.GetMappedProperty("DesignStatus");
        partData.Designer = mgr.GetMappedProperty("Designer");
        partData.Engineer = mgr.GetMappedProperty("Engineer");
        partData.Authority = mgr.GetMappedProperty("Authority");
        partData.Mass = mgr.GetMappedProperty("Mass");
        partData.SurfaceArea = mgr.GetMappedProperty("SurfaceArea");
        partData.Volume = mgr.GetMappedProperty("Volume");
        partData.SheetMetalRule = mgr.GetMappedProperty("SheetMetalRule");
        partData.FlatPatternWidth = mgr.GetMappedProperty("FlatPatternWidth");
        partData.FlatPatternLength = mgr.GetMappedProperty("FlatPatternLength");
        partData.FlatPatternArea = mgr.GetMappedProperty("FlatPatternArea");
        partData.Appearance = mgr.GetMappedProperty("Appearance");
    }
    

    // Перегрузка для сборок - открывает документ по partNumber
    private async Task<PartData> GetPartDataAsync(string partNumber, int quantity, int itemNumber)
    {
        var partDoc = OpenPartDocument(partNumber);
        if (partDoc == null) return null!;
        
        return await GetPartDataAsync(partDoc, quantity, itemNumber);
    }

    // Основной метод - работает с уже открытым документом  
    private async Task<PartData> GetPartDataAsync(PartDocument partDoc, int quantity, int itemNumber)
    {

        // Создаем новый объект PartData
        var partData = new PartData
        {
            // КАТЕГОРИЯ 1: Системные свойства приложения
            Item = itemNumber,
            
            // КАТЕГОРИЯ 6: Свойства количества и состояния
            OriginalQuantity = quantity,
            Quantity = quantity
        };

        // Создаем единый экземпляр PropertyManager для всех операций
        var mgr = new PropertyManager((Document)partDoc);
        
        // КАТЕГОРИЯ 2-4: Чтение всех свойств документа и iProperty
        ReadAllPropertiesFromPart(partDoc, partData, mgr);

        // КАТЕГОРИЯ 5: Получаем значения для пользовательских iProperty
        foreach (var customProperty in _customPropertiesList)
        {
            partData.CustomProperties[customProperty] = mgr.GetMappedProperty(customProperty);
        }

        // КАТЕГОРИЯ 2: Дополнительные свойства документа (изображения)
        partData.Preview = await GetThumbnailAsync(partDoc);        

        return partData;
    }

    private BitmapImage GenerateDxfThumbnail(string dxfDirectory, string partNumber)
    {
        var searchPattern = partNumber + "*.dxf"; // Шаблон поиска
        var dxfFiles = Directory.GetFiles(dxfDirectory, searchPattern);

        if (dxfFiles.Length == 0) return null!;

        try
        {
            var dxfFilePath = dxfFiles[0]; // Берем первый найденный файл, соответствующий шаблону
            var generator = new DxfThumbnailGenerator();
            var bitmap = generator.GenerateThumbnail(dxfFilePath);

            BitmapImage? bitmapImage = null;

            // Инициализация изображения должна выполняться в UI потоке
            Dispatcher.Invoke(() =>
            {
                using (var memoryStream = new MemoryStream())
                {
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Position = 0;

                    bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memoryStream;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                }
            });

            return bitmapImage!;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show($"Ошибка при генерации миниатюры DXF: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null!;
        }
    }

    private string GetModelStateName(Document doc)
    {
        try
        {
            return doc.ModelStateName ?? "Ошибка получения имени";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при получении Model State: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
            return "Ошибка получения имени";
        }
    }

    private async Task<BitmapImage> GetThumbnailAsync(PartDocument document)
    {
        try
        {
            BitmapImage? bitmap = null;
            await Dispatcher.InvokeAsync(() =>
            {
                var apprentice = new ApprenticeServerComponent();
                var apprenticeDoc = apprentice.Open(document.FullDocumentName);

                var thumbnail = apprenticeDoc.Thumbnail;
                var img = IPictureDispConverter.PictureDispToImage(thumbnail);

                if (img != null)
                    using (var memoryStream = new MemoryStream())
                    {
                        img.Save(memoryStream, ImageFormat.Png);
                        memoryStream.Position = 0;

                        bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = memoryStream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                    }
            });

            return bitmap!;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show("Ошибка при получении миниатюры: " + ex.Message, "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null!;
        }
    }

    private static string GetFullFileName(ComponentOccurrence occ)
    {
        string fullFileName = "";
        if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = occ.Definition.Document as PartDocument;
            if (partDoc != null)
                fullFileName = partDoc.FullFileName;
        }
        else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = occ.Definition.Document as AssemblyDocument;
            if (asmDoc != null)
                fullFileName = asmDoc.FullFileName;
        }
        return fullFileName;
    }

    private void ProcessComponentOccurrences(ComponentOccurrences occurrences, Dictionary<string, int> sheetMetalParts, IProgress<ScanProgress>? scanProgress = null, CancellationToken cancellationToken = default)
    {
        // Предварительно считаем только компоненты, которые пройдут фильтрацию
        var filteredOccurrences = new List<ComponentOccurrence>();
        foreach (ComponentOccurrence occ in occurrences)
        {
            if (occ.Suppressed) continue;
            
            // Проверяем, является ли компонент виртуальным
            if (occ.Definition is VirtualComponentDefinition) continue;
            
            try
            {
                var fullFileName = GetFullFileName(occ);
                if (!ShouldExcludeComponent(occ.BOMStructure, fullFileName))
                    filteredOccurrences.Add(occ);
            }
            catch
            {
                // Игнорируем компоненты с ошибками при фильтрации
            }
        }
        
        var totalOccurrences = filteredOccurrences.Count;
        var processedOccurrences = 0;

        foreach (ComponentOccurrence occ in filteredOccurrences)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var partDoc = occ.Definition.Document as PartDocument;
                    if (partDoc != null && partDoc.SubType == PropertyManager.SheetMetalSubType)
                    {
                        var mgr = new PropertyManager((Document)partDoc);
                        var partNumber = mgr.GetMappedProperty("PartNumber");
                        if (!string.IsNullOrEmpty(partNumber))
                        {
                            // Добавляем в трекер конфликтов
                            var modelState = mgr.GetModelState();
                            AddPartToConflictTracker(partNumber, partDoc.FullFileName, modelState);
                            
                            if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                                sheetMetalParts[partNumber]++;
                            else
                                sheetMetalParts.Add(partNumber, 1);
                        }
                    }
                }
                else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    ProcessComponentOccurrences((ComponentOccurrences)occ.SubOccurrences, sheetMetalParts, cancellationToken: cancellationToken); // Рекурсивный вызов без прогресса
                }
                
                // Обновляем прогресс
                processedOccurrences++;
                scanProgress?.Report(new ScanProgress
                {
                    ProcessedItems = processedOccurrences,
                    TotalItems = totalOccurrences,
                    CurrentOperation = "Сканирование компонентов",
                    CurrentItem = $"Компонент {processedOccurrences} из {totalOccurrences}"
                });
            }
            catch (Exception ex)
            {
                // Логирование ошибки
                Debug.WriteLine($"Ошибка при обработке компонента: {ex.Message}");
                
                // Обновляем прогресс даже при ошибке
                processedOccurrences++;
                scanProgress?.Report(new ScanProgress
                {
                    ProcessedItems = processedOccurrences,
                    TotalItems = totalOccurrences,
                    CurrentOperation = "Сканирование компонентов",
                    CurrentItem = $"Компонент {processedOccurrences} из {totalOccurrences}"
                });
            }
        }
    }
    private void ProcessBOM(BOM bom, Dictionary<string, int> sheetMetalParts, IProgress<ScanProgress>? scanProgress = null, CancellationToken cancellationToken = default)
    {
        try
        {
            // Получаем все компоненты из BOM рекурсивно в один плоский список с количествами
            var allBomRows = GetAllBOMRowsRecursively(bom);
            
            Debug.WriteLine($"[ProcessBOM] Всего найдено {allBomRows.Count} строк BOM во всей структуре");
            
            var processedRows = 0;
            var totalRows = allBomRows.Count;

            foreach (var (row, parentQuantity) in allBomRows)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    ProcessBOMRowSimple(row, sheetMetalParts, parentQuantity);
                    
                    // Обновляем прогресс
                    processedRows++;
                    scanProgress?.Report(new ScanProgress
                    {
                        ProcessedItems = processedRows,
                        TotalItems = totalRows,
                        CurrentOperation = "Сканирование спецификации",
                        CurrentItem = $"Строка {processedRows} из {totalRows}"
                    });
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
                {
                    _hasMissingReferences = true;
                    Debug.WriteLine($"Обнаружен компонент с потерянной ссылкой: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка при обработке строки BOM: {ex.Message}");
                    _hasMissingReferences = true;
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Общая ошибка при обработке BOM: {ex.Message}");
            _hasMissingReferences = true;
        }
    }

    private List<(BOMRow Row, int ParentQuantity)> GetAllBOMRowsRecursively(BOM bom, int parentQuantity = 1)
    {
        var allRows = new List<(BOMRow, int)>();
        
        try
        {
            // Проверяем настройку скрытия подавленных компонентов
            var hideSuppressed = bom.HideSuppressedComponentsInBOM;
            
            BOMView? bomView = null;
            foreach (BOMView view in bom.BOMViews)
            {
                if (view.ViewType == BOMViewTypeEnum.kModelDataBOMViewType)
                {
                    bomView = view;
                    break;
                }
            }

            if (bomView == null) return allRows;

            try
            {
                var bomRows = bomView.BOMRows.Cast<BOMRow>().ToArray();
                
                foreach (BOMRow row in bomRows)
                {
                    // Фильтруем подавленные компоненты
                    if (!hideSuppressed && row.ItemQuantity <= 0)
                        continue;

                    // Применяем фильтрацию BOM структуры перед добавлением в список
                    var componentDefinition = row.ComponentDefinitions[1];
                    
                    // Проверяем, является ли компонент виртуальным
                    if (componentDefinition is VirtualComponentDefinition)
                    {
                        // Виртуальные компоненты пропускаем
                        continue;
                    }
                    
                    if (componentDefinition?.Document is Document document && 
                        ShouldExcludeComponent(row.BOMStructure, document.FullFileName))
                        continue;
                        
                    // Добавляем строку с учетом количества родительского компонента
                    allRows.Add((row, parentQuantity));
                    
                    // Рекурсивно получаем строки из подсборок
                    try
                    {
                        if (componentDefinition?.Document is AssemblyDocument asmDoc)
                        {
                            // Передаем общее количество: parentQuantity * количество данной подсборки
                            var totalQuantity = parentQuantity * row.ItemQuantity;
                            var subRows = GetAllBOMRowsRecursively(asmDoc.ComponentDefinition.BOM, totalQuantity);
                            allRows.AddRange(subRows);
                        }
                    }
                    catch (Exception)
                    {
                        // Игнорируем ошибки при получении подсборок
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
            {
                _hasMissingReferences = true;
                Debug.WriteLine("Ошибка при доступе к BOMRows. Возможно, в сборке есть потерянные ссылки.");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка при рекурсивном получении BOM: {ex.Message}");
        }
        
        return allRows;
    }

    private void ProcessBOMRowSimple(BOMRow row, Dictionary<string, int> sheetMetalParts, int parentQuantity = 1)
    {
        try
        {
            var componentDefinition = row.ComponentDefinitions[1];
            if (componentDefinition == null) return;

            var document = componentDefinition.Document as Document;
            if (document == null) return;

            // Обрабатываем только детали из листового металла (не подсборки)
            if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                var partDoc = document as PartDocument;
                if (partDoc != null && partDoc.SubType == PropertyManager.SheetMetalSubType)
                {
                    var mgr = new PropertyManager((Document)partDoc);
                    var partNumber = mgr.GetMappedProperty("PartNumber");
                    if (!string.IsNullOrEmpty(partNumber))
                    {
                        // Добавляем в трекер конфликтов
                        var modelState = mgr.GetModelState();
                        AddPartToConflictTracker(partNumber, partDoc.FullFileName, modelState);
                        
                        // Учитываем общее количество: количество детали * количество родительских сборок
                        var totalQuantity = row.ItemQuantity * parentQuantity;
                        
                        if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                            sheetMetalParts[partNumber] += totalQuantity;
                        else
                            sheetMetalParts.Add(partNumber, totalQuantity);
                    }
                }
            }
            // Подсборки больше не обрабатываем - они уже развернуты в GetAllBOMRowsRecursively
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка при простой обработке строки BOM: {ex.Message}");
        }
    }

    private void InitializeLibraryPaths()
    {
        if (_thisApplication == null) return;

        try
        {
            var project = _thisApplication.DesignProjectManager.ActiveDesignProject;
            foreach (ProjectPath projectPath in project.LibraryPaths)
            {
                _libraryPaths.Add(projectPath.Path);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при инициализации путей библиотек: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    private bool IsLibraryComponent(string fullFileName)
    {
        return _libraryPaths.Any(path => fullFileName.StartsWith(path, StringComparison.OrdinalIgnoreCase));
    }
    private bool ShouldExcludeComponent(BOMStructureEnum bomStructure, string fullFileName)
    {
        if (ExcludeReferenceParts && bomStructure == BOMStructureEnum.kReferenceBOMStructure)
            return true;

        if (ExcludePurchasedParts && bomStructure == BOMStructureEnum.kPurchasedBOMStructure)
            return true;

        if (ExcludePhantomParts && bomStructure == BOMStructureEnum.kPhantomBOMStructure)
            return true;

        if (!IncludeLibraryComponents && !string.IsNullOrEmpty(fullFileName) && IsLibraryComponent(fullFileName))
            return true;

        return false;
    }
    private void SelectFixedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            FixedFolderPath = dialog.SelectedPath;
        }
    }

/// <summary>
/// Централизованная подготовка контекста экспорта
/// </summary>
private async Task<ExportContext> PrepareExportContextAsync(Document document, bool requireScan = true, bool showProgress = false)
{
    var context = new ExportContext();
    
    try
    {
        // Проверяем документ на валидность
        if (requireScan && document != _lastScannedDocument)
        {
            context.IsValid = false;
            context.ErrorMessage = "Активный документ изменился. Пожалуйста, повторите сканирование перед экспортом.";
            return context;
        }

        // Подготавливаем папку экспорта
        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch))
        {
            context.IsValid = false;
            context.ErrorMessage = "Ошибка при подготовке параметров экспорта";
            return context;
        }

        context.TargetDirectory = targetDir;
        context.Multiplier = multiplier;

        // Сканируем документ для получения данных
        var sheetMetalParts = new Dictionary<string, int>();

        if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = (AssemblyDocument)document;
            _partNumberTracker.Clear();
            
            if (SelectedProcessingMethod == ProcessingMethod.Traverse)
                await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts, showProgress ? null : null));
            else if (SelectedProcessingMethod == ProcessingMethod.BOM)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, showProgress ? null : null));

            FilterConflictingParts(sheetMetalParts);
        }
        else if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)document;
            if (partDoc.SubType == PropertyManager.SheetMetalSubType)
            {
                var mgr = new PropertyManager((Document)partDoc);
                var partNumber = mgr.GetMappedProperty("PartNumber");
                sheetMetalParts.Add(partNumber, 1);
            }
        }

        context.SheetMetalParts = sheetMetalParts;
        context.IsValid = true;
    }
    catch (Exception ex)
    {
        context.IsValid = false;
        context.ErrorMessage = $"Ошибка при подготовке экспорта: {ex.Message}";
    }

    return context;
}

private bool PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch)
{
    stopwatch = new Stopwatch();
    targetDir = "";
    multiplier = 1; // Присваиваем значение по умолчанию

    // Обрабатываем выбор папки экспорта через enum
    switch (SelectedExportFolder)
    {
        case ExportFolderType.ChooseFolder:
            // Открываем диалог выбора папки
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                targetDir = dialog.SelectedPath;
            else
                return false;
            break;
            
        case ExportFolderType.ComponentFolder:
            // Папка компонента
            targetDir = Path.GetDirectoryName(_thisApplication?.ActiveDocument?.FullFileName) ?? "";
            break;
            
        case ExportFolderType.FixedFolder:
            if (string.IsNullOrEmpty(_fixedFolderPath))
            {
                MessageBox.Show("Выберите фиксированную папку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            targetDir = _fixedFolderPath;
            break;
            
        case ExportFolderType.ProjectFolder:
            try
            {
                targetDir = _thisApplication?.DesignProjectManager?.ActiveDesignProject?.WorkspacePath ?? "";
                if (string.IsNullOrEmpty(targetDir))
                {
                    MessageBox.Show("Не удалось получить путь проекта.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении папки проекта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            break;
            
        case ExportFolderType.PartFolder:
            // Обрабатывается отдельно в ExportDXF методе
            break;
    }

    if (EnableSubfolder && !string.IsNullOrEmpty(SubfolderNameTextBox.Text))
    {
        targetDir = Path.Combine(targetDir!, SubfolderNameTextBox.Text);
        if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
    }

    if (!int.TryParse(MultiplierTextBox.Text, out multiplier))
    {
        MessageBox.Show("Введите допустимое целое число для множителя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return false;
    }


    SetUIState(UIState.Exporting);

    stopwatch = Stopwatch.StartNew();

    return true;
}

    private void FinalizeExport(bool isCancelled, Stopwatch stopwatch, int processedCount, int skippedCount)
    {
        stopwatch.Stop();
        var elapsedTime = GetElapsedTime(stopwatch.Elapsed);

        _isExporting = false;
        
        var result = new OperationResult
        {
            ProcessedCount = processedCount,
            SkippedCount = skippedCount,
            ElapsedTime = stopwatch.Elapsed,
            WasCancelled = isCancelled
        };

        SetUIStateAfterOperation(result, OperationType.Export);
        ShowOperationResult(result, OperationType.Export);
    }

    private List<PartData> CreatePartDataListFromParts(Dictionary<string, int> parts, int multiplier)
    {
        return parts.Select(part => new PartData
        {
            PartNumber = part.Key,
            OriginalQuantity = part.Value,
            Quantity = part.Value * multiplier
        }).ToList();
    }

    private async Task ExportWithoutScan()
    {
        // Очистка таблицы перед началом скрытого экспорта
        ClearList_Click(this, null!);

        // Валидация документа
        var validation = ValidateActiveDocument();
        if (!validation.IsValid)
        {
            MessageBox.Show(validation.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Подготовка контекста экспорта (без требования предварительного сканирования и без отображения прогресса)
        var context = await PrepareExportContextAsync(validation.Document!, requireScan: false, showProgress: false);
        if (!context.IsValid)
        {
            MessageBox.Show(context.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Настройка UI для быстрого экспорта
        SetUIState(UIState.Exporting);
        _isExporting = true;
        _operationCts = new CancellationTokenSource();

        var stopwatch = Stopwatch.StartNew();

        // Создаем временный список PartData для экспорта (без UI обновлений)
        var tempPartsDataList = CreatePartDataListFromParts(context.SheetMetalParts, context.Multiplier);

        // Выполнение экспорта через централизованную обработку ошибок
        context.GenerateThumbnails = false;
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = 0;
            await Task.Run(() => ExportDXF(tempPartsDataList, context.TargetDirectory, context.Multiplier, 
                ref processedCount, ref skippedCount, context.GenerateThumbnails, _operationCts.Token), _operationCts.Token);
            
            return new OperationResult
            {
                ProcessedCount = processedCount,
                SkippedCount = skippedCount,
                ElapsedTime = stopwatch.Elapsed,
                WasCancelled = _operationCts.Token.IsCancellationRequested
            };
        }, "быстрого экспорта");

        // Завершение операции
        _isExporting = false;

        // Показ результата быстрого экспорта до установки состояния UI
        ShowOperationResult(result, OperationType.Export, isQuickMode: true);
        
        // Используем единообразный подход для завершения операции
        SetUIStateAfterOperation(result, OperationType.Export);
    }


    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        // Переносим фокус на кнопку сканирования, чтобы убрать выделение с кнопки экспорта
        ScanButton.Focus();
        
        // Если экспорт уже идет, выполняем прерывание
        if (_isExporting)
        {
            _operationCts?.Cancel();
            SetUIState(new UIState { ExportEnabled = false, ExportButtonText = "Прерывание..." });
            return;
        }

        // Режим быстрого экспорта с Ctrl
        if (_isCtrlPressed)
        {
            await ExportWithoutScan();
            return;
        }

        // Валидация документа
        var validation = ValidateActiveDocument();
        if (!validation.IsValid)
        {
            MessageBox.Show(validation.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Подготовка контекста экспорта
        var context = await PrepareExportContextAsync(validation.Document!, requireScan: true, showProgress: true);
        if (!context.IsValid)
        {
            MessageBox.Show(context.ErrorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            if (context.ErrorMessage.Contains("сканирование"))
                MultiplierTextBox.Text = "1";
            return;
        }

        // Настройка UI для экспорта
        SetUIState(UIState.Exporting);
        _isExporting = true;
        _operationCts = new CancellationTokenSource();

        var stopwatch = Stopwatch.StartNew();

        // Выполнение экспорта через централизованную обработку ошибок
        var partsDataList = _partsData.Where(p => context.SheetMetalParts.ContainsKey(p.PartNumber)).ToList();
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = 0;
            await Task.Run(() => ExportDXF(partsDataList, context.TargetDirectory, context.Multiplier, 
                ref processedCount, ref skippedCount, context.GenerateThumbnails, _operationCts.Token), _operationCts.Token);
            
            return new OperationResult
            {
                ProcessedCount = processedCount,
                SkippedCount = skippedCount,
                ElapsedTime = stopwatch.Elapsed,
                WasCancelled = _operationCts.Token.IsCancellationRequested
            };
        }, "экспорта");

        // Завершение операции
        _isExporting = false;

        SetUIStateAfterOperation(result, OperationType.Export);
        
        // Показ результата
        ShowOperationResult(result, OperationType.Export);
    }

    private string GetElapsedTime(TimeSpan timeSpan)
    {
        if (timeSpan.TotalMinutes >= 1)
            return $"{(int)timeSpan.TotalMinutes} мин. {timeSpan.Seconds}.{timeSpan.Milliseconds:D3} сек.";
        return $"{timeSpan.Seconds}.{timeSpan.Milliseconds:D3} сек.";
    }

    private async void ExportSelectedDXF_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        
        var itemsWithoutFlatPattern = selectedItems.Where(p => !p.HasFlatPattern).ToList();
        if (itemsWithoutFlatPattern.Count == selectedItems.Count)
        {
            MessageBox.Show("Выбранные файлы не содержат разверток.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (itemsWithoutFlatPattern.Count > 0)
        {
            var result =
                MessageBox.Show("Некоторые выбранные файлы не содержат разверток и будут пропущены. Продолжить?",
                    "Информация", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.No) return;

            selectedItems = selectedItems.Except(itemsWithoutFlatPattern).ToList();
        }

        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch)) return;


        var processedCount = 0;
        var skippedCount = itemsWithoutFlatPattern.Count;

        _operationCts = new CancellationTokenSource();
        try
        {
            await Task.Run(() =>
            {
                ExportDXF(selectedItems, targetDir, multiplier, ref processedCount, ref skippedCount,
                    true, _operationCts.Token); // true - с генерацией миниатюр
            }, _operationCts.Token);
        }
        catch (OperationCanceledException)
        {
            // Операция была отменена - это нормальное поведение
        }


        FinalizeExport(_operationCts?.Token.IsCancellationRequested ?? false, stopwatch, processedCount, skippedCount);
    }

    private void OverrideQuantity_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        
        // Передаем текущее количество первого выбранного элемента для отображения пользователю
        var currentQuantity = selectedItems.FirstOrDefault()?.Quantity;
        var dialog = new OverrideQuantityDialog(currentQuantity);
        dialog.Owner = this;
        dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

        if (dialog.ShowDialog() == true && dialog.NewQuantity.HasValue)
        {
            foreach (var item in selectedItems)
            {
                item.Quantity = dialog.NewQuantity.Value;
                item.IsOverridden = true;
            }
        }
    }

    private void ExportDXF(IEnumerable<PartData> partsDataList, string targetDir, int multiplier,
        ref int processedCount, ref int skippedCount, bool generateThumbnails, CancellationToken cancellationToken = default)
    {
        var organizeByMaterial = false;
        var organizeByThickness = false;

        // Используем свойства вместо прямого обращения к элементам управления
        organizeByMaterial = OrganizeByMaterial;
        organizeByThickness = OrganizeByThickness;

        var totalParts = partsDataList.Count();

        Dispatcher.Invoke(() =>
        {
            UpdateExportProgress(0);
        });

        var localProcessedCount = processedCount;
        var localSkippedCount = skippedCount;

        foreach (var partData in partsDataList)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var partNumber = partData.PartNumber;
            var qty = partData.IsOverridden ? partData.Quantity : partData.OriginalQuantity * multiplier;

            // Обработка для PartFolder через свойство
            if (SelectedExportFolder == ExportFolderType.PartFolder)
            {
                var partPath = GetPartDocumentFullPath(partNumber);
                if (!string.IsNullOrEmpty(partPath))
                {
                    targetDir = Path.GetDirectoryName(partPath) ?? "";
                }
                else
                {
                    targetDir = "";
                }
            }

            PartDocument? partDoc = null;
            try
            {
                partDoc = OpenPartDocument(partNumber);
                if (partDoc == null) throw new Exception("Файл детали не найден или не может быть открыт");

                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var material = partData.Material;
                var thickness = partData.Thickness;

                var materialDir = organizeByMaterial ? Path.Combine(targetDir, material) : targetDir;
                if (!Directory.Exists(materialDir)) Directory.CreateDirectory(materialDir);

                var thicknessDir = organizeByThickness
                    ? Path.Combine(materialDir, thickness.ToString("F1"))
                    : materialDir;
                if (!Directory.Exists(thicknessDir)) Directory.CreateDirectory(thicknessDir);

                // Генерируем имя файла на основе конструктора
                string fileName;
                if (EnableFileNameConstructor && !string.IsNullOrEmpty(_tokenService.FileNameTemplate))
                {
                    fileName = _tokenService.ResolveTemplate(_tokenService.FileNameTemplate, partData);
                }
                else
                {
                    fileName = partNumber;
                }

                var filePath = Path.Combine(thicknessDir, fileName + ".dxf");

                if (!IsValidPath(filePath)) continue;

                var exportSuccess = false;

                if (smCompDef.HasFlatPattern)
                {
                    var flatPattern = smCompDef.FlatPattern;
                    var oDataIO = flatPattern.DataIO;

                    try
                    {
                        string? options = null;
                        Dispatcher.Invoke(() => PrepareExportOptions(out options));

                        // Интеграция настроек слоев
                        var layerOptionsBuilder = new StringBuilder();
                        var invisibleLayersBuilder = new StringBuilder();

                        foreach (var layer in LayerSettings) // Используем настройки слоев
                        {
                            if (layer.CanBeHidden && !layer.IsChecked)
                            {
                                invisibleLayersBuilder.Append($"{layer.LayerName};");
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(layer.CustomName))
                                layerOptionsBuilder.Append($"&{layer.DisplayName}={layer.CustomName}");

                            if (layer.SelectedLineType != "Default")
                                layerOptionsBuilder.Append(
                                    $"&{layer.DisplayName}LineType={LayerSettingsHelper.GetLineTypeValue(layer.SelectedLineType)}");

                            if (layer.SelectedColor != "White")
                                layerOptionsBuilder.Append(
                                    $"&{layer.DisplayName}Color={LayerSettingsHelper.GetColorValue(layer.SelectedColor)}");
                        }

                        if (invisibleLayersBuilder.Length > 0)
                        {
                            invisibleLayersBuilder.Length -= 1; // Убираем последний символ ";"
                            layerOptionsBuilder.Append($"&InvisibleLayers={invisibleLayersBuilder}");
                        }

                        // Объединяем общие параметры и параметры слоев
                        options += layerOptionsBuilder.ToString();

                        var dxfOptions = $"FLAT PATTERN DXF?{options}";

                        // Debugging output
                        Debug.WriteLine($"DXF Options: {dxfOptions}");

                        // Создаем файл, если он не существует
                        if (!File.Exists(filePath)) File.Create(filePath).Close();

                        // Проверка на занятость файла
                        if (IsFileLocked(filePath))
                        {
                            var result = MessageBox.Show(
                                $"Файл {filePath} занят другим приложением. Прервать операцию?",
                                "Предупреждение",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning,
                                MessageBoxResult.No // Устанавливаем "No" как значение по умолчанию
                            );

                            if (result == MessageBoxResult.Yes)
                            {
                                throw new OperationCanceledException(); // Прерываем операцию
                            }

                            while (IsFileLocked(filePath))
                            {
                                var waitResult = MessageBox.Show(
                                    "Ожидание разблокировки файла. Возможно файл открыт в программе просмотра. Закройте файл и нажмите OK.",
                                    "Информация",
                                    MessageBoxButton.OKCancel,
                                    MessageBoxImage.Information,
                                    MessageBoxResult.Cancel // Устанавливаем "Cancel" как значение по умолчанию
                                );

                                if (waitResult == MessageBoxResult.Cancel)
                                {
                                    throw new OperationCanceledException(); // Прерываем операцию
                                }

                                Thread.Sleep(1000); // Ожидание 1 секунды перед повторной проверкой
                            }

                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        oDataIO.WriteDataToFile(dxfOptions, filePath);
                        exportSuccess = true;

                        // Оптимизация DXF если включена настройка и версия не R12
                        if (OptimizeDxf && exportSuccess)
                        {
                            var selectedVersion = AcadVersions[SelectedAcadVersionIndex].Value;
                            // Не оптимизируем файлы R12, так как netDxf не поддерживает эту версию
                            if (selectedVersion != "R12")
                            {
                                DxfOptimizer.OptimizeDxfFile(filePath, selectedVersion);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Обработка ошибок экспорта DXF
                        Debug.WriteLine($"Ошибка экспорта DXF: {ex.Message}");
                    }
                }

                BitmapImage? dxfPreview = null;
                if (generateThumbnails)
                {
                    var selectedVersion = "";
                    Dispatcher.Invoke(() => 
                    {
                        selectedVersion = AcadVersions[SelectedAcadVersionIndex].Value;
                    });
                    
                    // Не создаваем миниатюры для R12, так как netDxf не поддерживает эту версию
                    if (selectedVersion != "R12")
                    {
                        dxfPreview = GenerateDxfThumbnail(thicknessDir, partNumber);
                    }
                }

                Dispatcher.Invoke(() =>
                {
                    partData.ProcessingStatus = exportSuccess ? ProcessingStatus.Success : ProcessingStatus.Error;
                    partData.DxfPreview = dxfPreview!;
                });

                if (exportSuccess)
                    localProcessedCount++;
                else
                    localSkippedCount++;

                Dispatcher.Invoke(() => 
                { 
                    UpdateExportProgress(totalParts > 0 ? (double)localProcessedCount / totalParts * 100 : 0);
                });
            }
            catch (Exception ex)
            {
                localSkippedCount++;
                Debug.WriteLine($"Ошибка обработки детали: {ex.Message}");
            }
        }

        processedCount = localProcessedCount;
        skippedCount = localSkippedCount;

        Dispatcher.Invoke(() =>
        {
            UpdateExportProgress(100);
        });
    }

    private bool IsFileLocked(string filePath)
    {
        try
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                stream.Close();
            }
        }
        catch (IOException)
        {
            return true;
        }

        return false;
    }

    private PartDocument? OpenPartDocument(string partNumber)
    {
        var docs = _thisApplication?.Documents;
        if (docs == null) return null;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd)
            {
                var mgr = new PropertyManager((Document)pd);
                if (mgr.GetMappedProperty("PartNumber") == partNumber)
                    return pd; // Возвращаем найденный открытый документ
            }

        // Если документ не найден
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    private bool IsValidPath(string path)
    {
        try
        {
            new FileInfo(path);
            return true;
        }
        catch
        {
            return false;
        }
    }
    private void MultiplierTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
        {
            UpdateQuantitiesWithMultiplier(multiplier);

            // Проверка на null перед изменением состояния кнопки
            if (ClearMultiplierButton != null)
                ClearMultiplierButton.IsEnabled = multiplier > 1;
        }
        else
        {
            // Если введенное значение некорректное, сбрасываем текст на "1" и выключаем кнопку сброса
            MultiplierTextBox.Text = "1";
            UpdateQuantitiesWithMultiplier(1);

            if (ClearMultiplierButton != null) ClearMultiplierButton.IsEnabled = false;
        }
    }

    private void ClearMultiplierButton_Click(object sender, RoutedEventArgs e)
    {
        // Сбрасываем множитель на 1
        MultiplierTextBox.Text = "1";
        UpdateQuantitiesWithMultiplier(1);

        // Выключаем кнопку сброса
        ClearMultiplierButton.IsEnabled = false;
    }

    private void UpdateDocumentInfo(string documentType, Document? doc)
    {
        if (string.IsNullOrEmpty(documentType) || doc == null)
        {
            var noDocState = UIState.Initial;
            noDocState.ProgressText = "Информация о документе не доступна";
            SetUIState(noDocState);
            DocumentTypeLabel.Text = "";
            PartNumberLabel.Text = "";
            DescriptionLabel.Text = "";
            ModelStateLabel.Text = "";
            IsPrimaryModelState = true; // Сбрасываем состояние модели по умолчанию
            return;
        }

        var mgr = new PropertyManager(doc);
        var partNumber = mgr.GetMappedProperty("PartNumber");
        var description = mgr.GetMappedProperty("Description");
        var modelStateInfo = GetModelStateName(doc);
        var isPrimaryModelState = mgr.IsPrimaryModelState();

        // Заполняем отдельные поля в блоке информации о документе
        DocumentTypeLabel.Text = documentType;
        PartNumberLabel.Text = partNumber;
        DescriptionLabel.Text = description;
        ModelStateLabel.Text = modelStateInfo;
        
        // Устанавливаем свойство для триггера стиля
        IsPrimaryModelState = isPrimaryModelState;

    }

    private void ClearList_Click(object sender, RoutedEventArgs e)
    {
        // Очищаем данные в таблице или любом другом источнике данных
        _partsData.Clear();
        SetUIState(UIState.CreateClearedState());

        // Обнуляем информацию о документе
        UpdateDocumentInfo("", null);

        // Отключаем кнопку "Конфликты" и очищаем список конфликтов
        ConflictFilesButton.IsEnabled = false;
        _conflictFileDetails.Clear(); // Очищаем список с конфликтами (или используем нужную коллекцию)
        _partNumberTracker.Clear(); // Очищаем трекер конфликтов

    }

    private void RemoveSelectedRows_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        
        var result = MessageBox.Show($"Вы действительно хотите удалить {selectedItems.Count} строк(и)?", 
            "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);
        
        if (result == MessageBoxResult.Yes)
        {
            foreach (var item in selectedItems)
            {
                _partsData.Remove(item);
            }
            
            if (_partsData.Count == 0)
            {
                ClearList_Click(sender, e);
            }
            else
            {
                var deletionState = UIState.CreateAfterOperationState(_partsData.Count > 0, false, $"Удалено {selectedItems.Count} строк(и)");
                SetUIState(deletionState);
            }
        }
    }

    private void partsDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_partsData.Count == 0) 
        {
            e.Handled = true; // Предотвращаем открытие контекстного меню, если таблица пуста
            return;
        }
        
        // Управляем доступностью пунктов меню в зависимости от выделения
        var hasSelection = PartsDataGrid.SelectedItems.Count > 0;
        var hasSingleSelection = PartsDataGrid.SelectedItems.Count == 1;
        
        ExportSelectedDXFMenuItem.IsEnabled = hasSelection;
        OverrideQuantityMenuItem.IsEnabled = hasSelection;
        OpenFileLocationMenuItem.IsEnabled = hasSelection;
        OpenSelectedModelsMenuItem.IsEnabled = hasSelection;
        EditIPropertyMenuItem.IsEnabled = hasSingleSelection;
        RemoveSelectedRowsMenuItem.IsEnabled = hasSelection;
    }

    private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();

        foreach (var item in selectedItems)
        {
            var partNumber = item.PartNumber;
            var fullPath = GetPartDocumentFullPath(partNumber);

            // Проверка на null перед использованием fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Проверка существования файла
            if (File.Exists(fullPath))
                try
                {
                    var argument = "/select, \"" + fullPath + "\"";
                    Process.Start("explorer.exe", argument);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии проводника для файла по пути {fullPath}: {ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            else
                MessageBox.Show($"Файл по пути {fullPath}, связанный с номером детали {partNumber}, не найден.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenSelectedModels_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        
        foreach (var item in selectedItems)
        {
            var partNumber = item.PartNumber;
            var fullPath = GetPartDocumentFullPath(partNumber);
            var targetModelState = item.ModelState;

            // Проверка на null перед использованием fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Открываем файл с указанием состояния модели
            OpenInventorDocument(fullPath, targetModelState);
        }
    }

    private string? GetPartDocumentFullPath(string partNumber)
    {
        var docs = _thisApplication?.Documents;
        if (docs == null) return null;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd)
            {
                var mgr = new PropertyManager((Document)pd);
                if (mgr.GetMappedProperty("PartNumber") == partNumber)
                    return pd.FullFileName; // Возвращаем полный путь найденного документа
            }

        // Если документ не найден среди открытых
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    /// <summary>
    /// Централизованный метод для открытия файла в Inventor с указанием состояния модели
    /// </summary>
    /// <param name="filePath">Полный путь к файлу</param>
    /// <param name="modelState">Имя состояния модели (может быть null или пустая строка)</param>
    public void OpenInventorDocument(string filePath, string? modelState = null)
    {
        if (!File.Exists(filePath))
        {
            MessageBox.Show($"Файл по пути {filePath} не найден.", "Ошибка", 
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            if (!string.IsNullOrEmpty(modelState))
            {
                // Используем синтаксис с угловыми скобками для указания состояния модели
                var pathWithModelState = $"{filePath}<{modelState}>";
                _thisApplication?.Documents?.Open(pathWithModelState);
            }
            else
            {
                // Открываем обычным способом, если состояние модели не указано
                _thisApplication?.Documents?.Open(filePath);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при открытии файла по пути {filePath}: {ex.Message}", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void partsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Активация пункта "Редактировать iProperty" только при выборе одной строки
        EditIPropertyMenuItem.IsEnabled = PartsDataGrid.SelectedItems.Count == 1;
        
        // Обновляем предпросмотр с выбранной строкой
        var selectedPart = PartsDataGrid.SelectedItem as PartData;
        _tokenService.UpdatePreviewWithSelectedData(selectedPart);
    }

    public async Task FillPropertyDataAsync(string propertyName)
    {
        // Если нет данных для заполнения, выходим
        if (_partsData.Count == 0)
            return;

        // Проверяем, является ли это стандартным свойством (уже загружено в ReadAllPropertiesFromPart)
        var propInfo = typeof(PartData).GetProperty(propertyName);
        var isStandardProperty = propInfo != null && propInfo.PropertyType == typeof(string);

        // Если это стандартное свойство, то оно уже загружено - просто уведомляем UI
        if (isStandardProperty)
        {
            foreach (var partData in _partsData)
            {
                partData.OnPropertyChanged(propertyName);
                await Task.Delay(10);
            }
            return;
        }

        // Для кастомных свойств загружаем данные из файлов
        foreach (var partData in _partsData)
        {
            var partDoc = OpenPartDocument(partData.PartNumber);
            if (partDoc != null)
            {
                var mgr = new PropertyManager((Document)partDoc);
                var value = mgr.GetMappedProperty(propertyName) ?? "";
                partData.AddCustomProperty(propertyName, value);
            }

            await Task.Delay(10);
        }
    }
    private void RemoveCustomIPropertyColumn(string propertyName)
    {
        // Удаление столбца из DataGrid
        var columnToRemove = PartsDataGrid.Columns.FirstOrDefault(c => c.Header as string == propertyName);
        if (columnToRemove != null) PartsDataGrid.Columns.Remove(columnToRemove);

        // Удаление всех данных, связанных с этим Custom IProperty
        foreach (var partData in _partsData)
            if (partData.CustomProperties.ContainsKey(propertyName))
                partData.RemoveCustomProperty(propertyName);

        // Обновляем список доступных свойств
        _customPropertiesList.Remove(propertyName);
    }

    /// <summary>
    /// Централизованный метод для создания текстовых колонок DataGrid
    /// </summary>
    private DataGridTextColumn CreateTextColumn(string header, string bindingPath)
    {
        return new DataGridTextColumn
        {
            Header = header,
            Binding = new Binding(bindingPath),
            ElementStyle = PartsDataGrid.FindResource("CenteredCellStyle") as Style
        };
    }

    public void AddCustomIPropertyColumn(string propertyName)
    {
        // Проверяем, существует ли уже колонка с таким именем или заголовком
        if (PartsDataGrid.Columns.Any(c => c.Header as string == propertyName))
        {
            MessageBox.Show($"Столбец с именем '{propertyName}' уже существует.", "Предупреждение", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        // Создаем текстовую колонку через централизованный метод
        var column = CreateTextColumn(propertyName, $"CustomProperties[{propertyName}]");

        PartsDataGrid.Columns.Add(column);

        // Дозаполняем данные для новой колонки (асинхронно)
        _ = FillPropertyDataAsync(propertyName);
    }

    private void EditIProperty_Click(object sender, RoutedEventArgs e)
    {
        var selectedItem = PartsDataGrid.SelectedItem as PartData;
        if (selectedItem == null) return;
        
        // Открываем документ детали
        var partDoc = OpenPartDocument(selectedItem.PartNumber);
        if (partDoc == null)
        {
            MessageBox.Show("Не удалось открыть документ детали для редактирования.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        var mgr = new PropertyManager((Document)partDoc);
        
        // Получаем выражения или значения для передачи в диалог
        var partNumberForEdit = selectedItem.PartNumberIsExpression 
            ? mgr.GetMappedPropertyExpression("PartNumber") 
            : selectedItem.PartNumber;
            
        var descriptionIsExpression = mgr.IsMappedPropertyExpression("Description");
        var descriptionForEdit = descriptionIsExpression 
            ? mgr.GetMappedPropertyExpression("Description") 
            : selectedItem.Description;
        
        var editDialog = new EditIPropertyDialog(partNumberForEdit, descriptionForEdit);
        if (editDialog.ShowDialog() == true)
        {
            // Обновляем iProperty в файле детали
            // Если изначально было выражение и новое значение начинается с "=", устанавливаем выражение
            if (selectedItem.PartNumberIsExpression && editDialog.PartNumber.StartsWith("="))
            {
                mgr.SetMappedPropertyExpression("PartNumber", editDialog.PartNumber);
            }
            else
            {
                mgr.SetMappedProperty("PartNumber", editDialog.PartNumber);
            }
            
            if (descriptionIsExpression && editDialog.Description.StartsWith("="))
            {
                mgr.SetMappedPropertyExpression("Description", editDialog.Description);
            }
            else
            {
                mgr.SetMappedProperty("Description", editDialog.Description);
            }

            // Сохраняем изменения
            partDoc.Save2();

            // Обновляем свойства детали в таблице после сохранения
            // Получаем актуальные значения (не выражения) для отображения
            selectedItem.PartNumber = mgr.GetMappedProperty("PartNumber");
            selectedItem.PartNumberIsExpression = mgr.IsMappedPropertyExpression("PartNumber");
            selectedItem.Description = mgr.GetMappedProperty("Description");
        }
    }


    private void AboutButton_Click(object sender, RoutedEventArgs e)
    {
        var aboutWindow = new AboutWindow(this); // Передаем текущий экземпляр MainWindow
        aboutWindow.ShowDialog(); // Отображаем окно как диалоговое
    }


    private void UpdateNoColumnsOverlayVisibility()
    {
        if (NoColumnsOverlay != null)
        {
            NoColumnsOverlay.Visibility = PartsDataGrid.Columns.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void PartsDataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Delete && PartsDataGrid.SelectedItems.Count > 0)
        {
            RemoveSelectedRows_Click(this, new RoutedEventArgs());
            e.Handled = true;
        }
    }
}