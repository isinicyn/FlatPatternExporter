using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using DxfRenderer;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using FlatPatternExporter.Converters;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Models;
using FlatPatternExporter.Services;
using FlatPatternExporter.Utilities;
using Inventor;
using Svg.Skia;
using Binding = System.Windows.Data.Binding;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Style = System.Windows.Style;

namespace FlatPatternExporter.UI.Windows;

public partial class FlatPatternExporterMainWindow : Window
{
    /// <summary>
    /// Конвертирует SVG строку в BitmapImage
    /// </summary>
    private static BitmapImage? ConvertSvgToBitmapImage(string svgContent)
    {
        if (string.IsNullOrEmpty(svgContent))
            return null;

        try
        {
            var svg = new SKSvg();
            var svgDocument = svg.FromSvg(svgContent);
            
            if (svgDocument == null)
                return null;

            var bounds = svgDocument.CullRect;
            
            if (bounds.Width <= 0 || bounds.Height <= 0)
                return null;

            using var surface = SkiaSharp.SKSurface.Create(new SkiaSharp.SKImageInfo((int)bounds.Width, (int)bounds.Height));
            if (surface?.Canvas == null)
                return null;

            surface.Canvas.Clear(SkiaSharp.SKColors.White);
            surface.Canvas.DrawPicture(svgDocument);
            surface.Canvas.Flush();

            using var image = surface.Snapshot();
            using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            
            var bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = new MemoryStream(data.ToArray());
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            bitmapImage.Freeze();

            return bitmapImage;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// Общий метод для чтения всех свойств из документа с использованием PropertyManager
    /// </summary>
    private static void ReadAllPropertiesFromPart(PartDocument _, PartData partData, Services.PropertyManager mgr)
    {
        partData.FileName = mgr.GetFileName();
        partData.FullFileName = mgr.GetFullFileName();
        partData.ModelState = mgr.GetModelState();
        partData.HasFlatPattern = mgr.HasFlatPattern();
        partData.Thickness = mgr.GetThickness();

        // Устанавливаем состояния выражений для всех редактируемых свойств
        SetExpressionStatesForAllProperties(partData, mgr);

        // Автоматически заполняем все свойства типа IProperty из реестра
        foreach (var property in PropertyMetadataRegistry.Properties.Values
                     .Where(p => p.Type == PropertyMetadataRegistry.PropertyType.IProperty))
        {
            var value = mgr.GetMappedProperty(property.InternalName);
            var propInfo = typeof(PartData).GetProperty(property.InternalName);
            
            if (propInfo != null && !string.IsNullOrEmpty(value))
            {
                propInfo.SetValue(partData, value);
            }
        }
    }

    // Перегрузка для сборок - открывает документ по partNumber
    private async Task<PartData> GetPartDataAsync(string partNumber, int quantity, int itemNumber, bool loadThumbnail = true)
    {
        var partDoc = GetCachedPartDocument(partNumber) ?? OpenPartDocument(partNumber);
        if (partDoc == null) return null!;
        
        return await GetPartDataAsync(partDoc, quantity, itemNumber, loadThumbnail);
    }

    // Основной метод - работает с уже открытым документом  
    private async Task<PartData> GetPartDataAsync(PartDocument partDoc, int quantity, int itemNumber, bool loadThumbnail = true)
    {

        // Создаем новый объект PartData
        var partData = new PartData
        {
            Item = itemNumber,
            
            OriginalQuantity = quantity
        };

        // Создаем единый экземпляр PropertyManager для всех операций
        var mgr = new Services.PropertyManager((Document)partDoc);
        
        ReadAllPropertiesFromPart(partDoc, partData, mgr);

        foreach (var userDefinedProperty in PropertyMetadataRegistry.UserDefinedProperties)
        {
            var value = mgr.GetMappedProperty(userDefinedProperty.InternalName);
            // Используем ColumnHeader как ключ для совместимости с биндингами DataGrid
            partData.UserDefinedProperties[userDefinedProperty.ColumnHeader] = value;
        }

        if (loadThumbnail)
        {
            partData.Preview = await GetThumbnailAsync(partDoc);
        }

        partData.SetQuantityInternal(quantity);
        
        return partData;
    }

    /// <summary>
    /// Устанавливает состояния выражений для всех редактируемых свойств
    /// </summary>
    private static void SetExpressionStatesForAllProperties(PartData partData, Services.PropertyManager mgr)
    {
        partData.BeginExpressionBatch();
        try
        {
            foreach (var property in PropertyMetadataRegistry.GetEditableProperties())
            {
                var isExpression = mgr.IsMappedPropertyExpression(property);
                partData.SetPropertyExpressionState(property, isExpression);
            }
            foreach (var userProperty in PropertyMetadataRegistry.UserDefinedProperties)
            {
                var isExpression = mgr.IsMappedPropertyExpression(userProperty.InternalName);
                partData.SetPropertyExpressionState($"UserDefinedProperties[{userProperty.ColumnHeader}]", isExpression);
            }
        }
        finally
        {
            partData.EndExpressionBatch();
        }
    }


    private BitmapImage? GenerateDxfThumbnails(string dxfDirectory, string partNumber)
    {
        var searchPattern = partNumber + "*.dxf"; // Шаблон поиска
        var dxfFiles = Directory.GetFiles(dxfDirectory, searchPattern);

        if (dxfFiles.Length == 0) return null;

        try
        {
            var dxfFilePath = dxfFiles[0]; // Берем первый найденный файл, соответствующий шаблону
            var generator = new DxfThumbnailGenerator();
            var svg = generator.GenerateSvg(dxfFilePath);

            BitmapImage? bitmapImage = null;

            // Конвертируем SVG в BitmapImage
            Dispatcher.Invoke(() =>
            {
                bitmapImage = ConvertSvgToBitmapImage(svg);
            });

            return bitmapImage;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show($"Ошибка при генерации миниатюр DXF: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
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
            if (occ.Definition.Document is PartDocument partDoc)
                fullFileName = partDoc.FullFileName;
        }
        else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            if (occ.Definition.Document is AssemblyDocument asmDoc)
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
                    if (occ.Definition.Document is PartDocument partDoc)
                    {
                        var mgr = new Services.PropertyManager((Document)partDoc);
                        var partNumber = mgr.GetMappedProperty("PartNumber");
                        if (!string.IsNullOrEmpty(partNumber))
                        {
                            // Добавляем документ в кеш
                            AddDocumentToCache(partDoc, partNumber);
                            
                            if (partDoc.SubType == Services.PropertyManager.SheetMetalSubType)
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

            if (componentDefinition.Document is not Document document) return;

            // Обрабатываем только детали из листового металла (не подсборки)
            if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                if (document is PartDocument partDoc)
                {
                    var mgr = new Services.PropertyManager((Document)partDoc);
                    var partNumber = mgr.GetMappedProperty("PartNumber");
                    if (!string.IsNullOrEmpty(partNumber))
                    {
                        // Добавляем документ в кеш
                        AddDocumentToCache(partDoc, partNumber);
                        
                        if (partDoc.SubType == Services.PropertyManager.SheetMetalSubType)
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
            }
            // Подсборки больше не обрабатываем - они уже развернуты в GetAllBOMRowsRecursively
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Ошибка при простой обработке строки BOM: {ex.Message}");
        }
    }

    private bool IsLibraryComponent(string fullFileName)
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
        // Очищаем конфликты только если сменился документ (чтобы не показывать неактуальные конфликты)
        if (requireScan && document != _lastScannedDocument)
        {
            ClearConflictData();
        }
        
        // Проверяем документ на валидность
        if (requireScan && document != _lastScannedDocument)
        {
            context.IsValid = false;
            context.ErrorMessage = "Активный документ изменился. Пожалуйста, повторите сканирование перед экспортом.";
            return context;
        }

        // Подготавливаем папку экспорта
        if (!PrepareForExport(out var targetDir, out var multiplier))
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
            
            if (SelectedProcessingMethod == ProcessingMethod.Traverse)
                await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts, showProgress ? null : null));
            else if (SelectedProcessingMethod == ProcessingMethod.BOM)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts, showProgress ? null : null));

            FilterConflictingParts(sheetMetalParts);
        }
        else if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)document;
            
            // Проверяем, не является ли деталь библиотечным компонентом
            if (!IncludeLibraryComponents && IsLibraryComponent(partDoc.FullFileName))
            {
                // Библиотечный компонент исключается
            }
            else
            {
                var mgr = new Services.PropertyManager((Document)partDoc);
                var partNumber = mgr.GetMappedProperty("PartNumber");
                if (!string.IsNullOrEmpty(partNumber))
                {
                    // Добавляем документ в кеш
                    AddDocumentToCache(partDoc, partNumber);
                    
                    if (partDoc.SubType == Services.PropertyManager.SheetMetalSubType)
                    {
                        sheetMetalParts.Add(partNumber, 1);
                    }
                }
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

    private bool PrepareForExport(out string targetDir, out int multiplier)
    {
        targetDir = "";
        multiplier = 1;

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

        return true;
    }

    private async Task ExportWithoutScan()
    {
        // Очистка таблицы перед началом скрытого экспорта
        ClearList_Click(this, null!);

        // Валидация документа
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Настройка UI для подготовки быстрого экспорта
        InitializeOperation(UIState.PreparingQuickExport, ref _isExporting);

        // Подготовка контекста экспорта (без требования предварительного сканирования и без отображения прогресса)
        var context = await PrepareExportContextOrShowError(validation.Document!, requireScan: false, showProgress: false);
        if (context == null) 
        {
            // Восстанавливаем состояние при ошибке
            SetUIState(UIState.CreateAfterOperationState(false, false, "Ошибка подготовки экспорта"));
            _isExporting = false;
            return;
        }

        // Создаем временный список PartData для экспорта с полными данными (это тоже длительная операция)
        var tempPartsDataList = new List<PartData>();
        var itemCounter = 1;
        var totalParts = context.SheetMetalParts.Count;
        
        foreach (var part in context.SheetMetalParts)
        {
            // Обновляем прогресс каждые 5 деталей или на важных моментах
            if (itemCounter % 5 == 1 || itemCounter == totalParts)
            {
                var progressState = UIState.PreparingQuickExport;
                progressState.ProgressText = $"Подготовка данных деталей ({itemCounter}/{totalParts})...";
                SetUIState(progressState);
                await Task.Delay(1); // Минимальная задержка для обновления UI
            }
            
            var partData = await GetPartDataAsync(part.Key, part.Value * context.Multiplier, itemCounter++, loadThumbnail: false);
            if (partData != null)
            {
                // Не вызываем SetQuantityInternal повторно - количество уже установлено правильно в GetPartDataAsync
                partData.IsMultiplied = context.Multiplier > 1;
                tempPartsDataList.Add(partData);
            }
        }

        // Переключаемся в состояние экспорта когда РЕАЛЬНО начинается экспорт файлов
        SetUIState(UIState.Exporting);
        
        var stopwatch = Stopwatch.StartNew();

        // Выполнение экспорта через централизованную обработку ошибок
        context.GenerateThumbnails = false;
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = 0;
            await Task.Run(() => ExportDXF(tempPartsDataList, context.TargetDirectory, context.Multiplier, 
                ref processedCount, ref skippedCount, context.GenerateThumbnails, _operationCts!.Token), _operationCts!.Token);
            
            return CreateExportOperationResult(processedCount, skippedCount, stopwatch.Elapsed);
        }, "быстрого экспорта");

        // Завершение операции
        CompleteOperation(result, OperationType.Export, ref _isExporting, isQuickMode: true);
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        // Если экспорт уже идет, выполняем прерывание
        if (HandleExportCancellation()) return;

        // Выполняем обычный экспорт
        await PerformNormalExportAsync();
    }

    private static string GetElapsedTime(TimeSpan timeSpan)
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
            var dialogResult =
                MessageBox.Show("Некоторые выбранные файлы не содержат разверток и будут пропущены. Продолжить?",
                    "Информация", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (dialogResult == MessageBoxResult.No) return;

            selectedItems = [.. selectedItems.Except(itemsWithoutFlatPattern)];
        }

        // Валидация документа
        var validation = ValidateDocumentOrShowError();
        if (validation == null) return;

        // Подготовка контекста экспорта
        var context = await PrepareExportContextOrShowError(validation.Document!, requireScan: true, showProgress: false);
        if (context == null) return;

        // Настройка UI для экспорта
        InitializeOperation(UIState.Exporting, ref _isExporting);
        var stopwatch = Stopwatch.StartNew();
        
        // Выполнение экспорта через централизованную обработку ошибок
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = itemsWithoutFlatPattern.Count;
            await Task.Run(() => ExportDXF(selectedItems, context.TargetDirectory, context.Multiplier, 
                ref processedCount, ref skippedCount, context.GenerateThumbnails, _operationCts!.Token), _operationCts!.Token);
            
            return CreateExportOperationResult(processedCount, skippedCount, stopwatch.Elapsed);
        }, "экспорта выделенных деталей");

        // Завершение операции
        CompleteOperation(result, OperationType.Export, ref _isExporting);
    }

    private void ExportDXF(IEnumerable<PartData> partsDataList, string targetDir, int multiplier,
        ref int processedCount, ref int skippedCount, bool generateThumbnails, CancellationToken cancellationToken = default)
    {

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
                partDoc = GetCachedPartDocument(partNumber) ?? OpenPartDocument(partNumber);
                if (partDoc == null) throw new Exception("Файл детали не найден или не может быть открыт");

                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var material = partData.Material;
                var thickness = partData.Thickness;

                var materialDir = OrganizeByMaterial ? Path.Combine(targetDir, material) : targetDir;
                if (!Directory.Exists(materialDir)) Directory.CreateDirectory(materialDir);

                var thicknessDir = OrganizeByThickness
                    ? Path.Combine(materialDir, thickness)
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
                            AcadVersionType selectedEnumVersion = AcadVersionType.V2000;
                            Dispatcher.Invoke(() => { selectedEnumVersion = SelectedAcadVersion; });
                            if (AcadVersionMapping.SupportsOptimization(selectedEnumVersion))
                            {
                                DxfOptimizer.OptimizeDxfFile(filePath, selectedEnumVersion);
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
                    AcadVersionType selectedEnumVersion = AcadVersionType.V2000;
                    Dispatcher.Invoke(() => { selectedEnumVersion = SelectedAcadVersion; });
                    if (AcadVersionMapping.SupportsOptimization(selectedEnumVersion))
                    {
                        dxfPreview = GenerateDxfThumbnails(thicknessDir, partNumber);
                    }
                }

                Dispatcher.Invoke(() =>
                {
                    partData.ProcessingStatus = exportSuccess ? ProcessingStatus.Success : ProcessingStatus.Error;
                    partData.DxfPreview = dxfPreview;
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

    private static bool IsFileLocked(string filePath)
    {
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            stream.Close();
        }
        catch (IOException)
        {
            return true;
        }

        return false;
    }

    private PartDocument? OpenPartDocument(string partNumber)
    {
        // Сначала проверяем кеш
        var cachedDoc = GetCachedPartDocument(partNumber);
        if (cachedDoc != null) return cachedDoc;

        var docs = _thisApplication?.Documents;
        if (docs == null) return null;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd)
            {
                var mgr = new Services.PropertyManager((Document)pd);
                if (mgr.GetMappedProperty("PartNumber") == partNumber)
                    return pd; // Возвращаем найденный открытый документ
            }

        // Если документ не найден
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    private static bool IsValidPath(string path)
    {
        try
        {
            _ = new FileInfo(path);
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

        var mgr = new Services.PropertyManager(doc);
        var partNumber = mgr.GetMappedProperty("PartNumber");
        var description = mgr.GetMappedProperty("Description");
        var modelStateInfo = mgr.GetModelState();
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
        
        // Обновляем TokenService для отображения placeholder'ов
        _tokenService.UpdatePartsData(_partsData);

        // Отключаем кнопку "Конфликты" и очищаем список конфликтов
        ClearConflictData();
        
        // Очищаем кеш документов
        ClearDocumentCache();

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
                // Обновляем TokenService после удаления строк
                _tokenService.UpdatePartsData(_partsData);
                
                var deletionState = UIState.CreateAfterOperationState(_partsData.Count > 0, false, $"Удалено {selectedItems.Count} строк(и)");
                SetUIState(deletionState);
            }
        }
    }

    private void PartsDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_partsData.Count == 0)
        {
            e.Handled = true; // Предотвращаем открытие контекстного меню, если таблица пуста
            return;
        }
        
        // Управляем доступностью пунктов меню в зависимости от выделения
        var hasSelection = PartsDataGrid.SelectedItems.Count > 0;
        var hasSingleSelection = PartsDataGrid.SelectedItems.Count == 1;
        var hasOverriddenSelection = hasSelection && PartsDataGrid.SelectedItems.Cast<PartData>().Any(item => item.IsOverridden);
        
        ExportSelectedDXFMenuItem.IsEnabled = hasSelection;
        OpenFileLocationMenuItem.IsEnabled = hasSelection;
        OpenSelectedModelsMenuItem.IsEnabled = hasSelection;
        ResetQuantityMenuItem.IsEnabled = hasOverriddenSelection;
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
        // Сначала проверяем кеш
        if (_partNumberToFullFileName.TryGetValue(partNumber, out var cachedPath))
            return cachedPath;

        var docs = _thisApplication?.Documents;
        if (docs == null) return null;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd)
            {
                var mgr = new Services.PropertyManager((Document)pd);
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

    private void PartsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Обновляем предпросмотр с выбранной строкой
        var selectedPart = PartsDataGrid.SelectedItem as PartData;
        _tokenService.UpdatePreviewWithSelectedData(selectedPart);
    }

    public void FillPropertyData(string propertyName)
    {
        // Если нет данных для заполнения, выходим
        if (_partsData.Count == 0)
            return;

        // Проверяем, является ли это стандартным свойством через реестр
        var isStandardProperty = PropertyMetadataRegistry.Properties.ContainsKey(propertyName);

        // Если это стандартное свойство, то оно уже загружено - данные готовы
        if (isStandardProperty)
        {
            return;
        }

        // Для User Defined Properties загружаем данные из файлов
        var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.InternalName == propertyName);
        var columnHeader = userProperty?.ColumnHeader ?? propertyName;

        foreach (var partData in _partsData)
        {
            var partDoc = GetCachedPartDocument(partData.PartNumber) ?? OpenPartDocument(partData.PartNumber);
            if (partDoc != null)
            {
                var mgr = new Services.PropertyManager((Document)partDoc);
                var value = mgr.GetMappedProperty(propertyName) ?? "";
                partData.AddUserDefinedProperty(columnHeader, value);
                
                // Устанавливаем состояние выражения для User Defined свойства
                var isExpression = mgr.IsMappedPropertyExpression(propertyName);
                partData.SetPropertyExpressionState($"UserDefinedProperties[{columnHeader}]", isExpression);
            }
        }
    }

    private void RemoveUserDefinedIPropertyColumn(string columnHeaderName)
    {
        // Удаление всех данных, связанных с этим User Defined Property
        // Находим InventorPropertyName для этого columnHeader
        var userProp = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.ColumnHeader == columnHeaderName);
        if (userProp != null)
        {
            foreach (var partData in _partsData)
                if (partData.UserDefinedProperties.ContainsKey(userProp.InventorPropertyName!))
                    partData.RemoveUserDefinedProperty(userProp.InventorPropertyName!);
        }

        // Ищем пользовательское свойство по ColumnHeader для удаления из реестра
        var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.ColumnHeader == columnHeaderName);
        if (userProperty != null)
        {
            PropertyMetadataRegistry.RemoveUserDefinedProperty(userProperty.InternalName);
        }
    }
    
    private void RemoveUserDefinedIPropertyData(string columnHeaderName)
    {
        // Удаление только данных UDP из PartData (без удаления из реестра)
        // Находим InventorPropertyName для этого columnHeader
        var userProp = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.ColumnHeader == columnHeaderName);
        if (userProp != null)
        {
            foreach (var partData in _partsData)
                if (partData.UserDefinedProperties.ContainsKey(userProp.InventorPropertyName!))
                    partData.RemoveUserDefinedProperty(userProp.InventorPropertyName!);
        }
    }
    
    /// <summary>
    /// Централизованный метод для удаления колонки из DataGrid
    /// </summary>
    public void RemoveDataGridColumn(string columnHeaderName, bool removeDataOnly = false)
    {
        // Удаляем колонку из DataGrid
        var columnToRemove = PartsDataGrid.Columns.FirstOrDefault(c => c.Header.ToString() == columnHeaderName);
        if (columnToRemove != null)
        {
            PartsDataGrid.Columns.Remove(columnToRemove);
        }

        // Проверяем, является ли это UDP свойством
        var internalName = PropertyMetadataRegistry.GetInternalNameByColumnHeader(columnHeaderName);
        if (!string.IsNullOrEmpty(internalName) && PropertyMetadataRegistry.IsUserDefinedProperty(internalName))
        {
            if (removeDataOnly)
            {
                // Только удаляем данные UDP (для Adorner)
                RemoveUserDefinedIPropertyData(columnHeaderName);
            }
            else
            {
                // Полное удаление UDP (для SelectIPropertyWindow)
                RemoveUserDefinedIPropertyColumn(columnHeaderName);
            }
        }

        // Обновляем состояние свойств в окне выбора
        var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>()
            .FirstOrDefault();
        selectIPropertyWindow?.UpdatePropertyStates();
    }

    /// <summary>
    /// Централизованный метод для создания текстовых колонок DataGrid
    /// </summary>
    private DataGridTextColumn CreateTextColumn(string header, string bindingPath, bool isSortable = true)
    {
        var binding = new Binding(bindingPath);
        
        return new DataGridTextColumn
        {
            Header = header,
            Binding = binding,
            SortMemberPath = isSortable ? bindingPath : null,
            ElementStyle = PartsDataGrid.FindResource("CenteredCellStyle") as Style,
            IsReadOnly = true
        };
    }

    private Style CreateCellTagStyle(string propertyPath)
    {
        var style = PartsDataGrid.FindResource("DataGridCellStyle") is Style baseStyle
            ? new Style(typeof(DataGridCell), baseStyle)
            : new Style(typeof(DataGridCell));
        style.Setters.Add(new Setter(FrameworkElement.TagProperty, propertyPath));
        return style;
    }

    public void AddUserDefinedIPropertyColumn(string propertyName)
    {
        // Ищем пользовательское свойство в реестре по оригинальному имени
        var userProperty = PropertyMetadataRegistry.UserDefinedProperties.FirstOrDefault(p => p.InventorPropertyName == propertyName);
        var columnHeader = userProperty?.ColumnHeader ?? $"(Пользов.) {propertyName}";

        // Проверяем, существует ли уже колонка с таким заголовком
        if (PartsDataGrid.Columns.Any(c => c.Header as string == columnHeader))
        {
            MessageBox.Show($"Столбец с именем '{columnHeader}' уже существует.", "Предупреждение", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        // Инициализируем пустые значения для всех существующих строк, чтобы избежать KeyNotFoundException
        foreach (var partData in _partsData)
        {
            partData.AddUserDefinedProperty(columnHeader, "");
        }

        // Создаем колонку с универсальным шаблоном и передаем путь через Tag
        var template = FindResource("EditableWithFxTemplate") as DataTemplate ?? throw new ResourceReferenceKeyNotFoundException("EditableWithFxTemplate не найден", "EditableWithFxTemplate");
        var column = new DataGridTemplateColumn
        {
            Header = columnHeader,
            CellTemplate = template,
            SortMemberPath = $"UserDefinedProperties[{columnHeader}]",
            IsReadOnly = false,
            ClipboardContentBinding = new Binding($"UserDefinedProperties[{columnHeader}]"),
            CellStyle = CreateCellTagStyle($"UserDefinedProperties[{columnHeader}]")
        };

        PartsDataGrid.Columns.Add(column);

        // Дозаполняем данные для новой колонки
        FillPropertyData(userProperty?.InternalName ?? $"UDP_{propertyName}");
    }

    private void AboutButton_Click(object sender, RoutedEventArgs e)
    {
        var aboutWindow = new AboutWindow();
        aboutWindow.ShowDialog();
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

    private void ResetQuantity_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0) return;

        var overriddenItems = selectedItems.Where(item => item.IsOverridden).ToList();
        if (overriddenItems.Count == 0)
        {
            MessageBox.Show("В выбранных элементах нет переопределенных количеств.", "Информация", 
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var message = overriddenItems.Count == 1 
            ? $"Сбросить количество для \"{overriddenItems[0].PartNumber}\" до исходного значения {overriddenItems[0].OriginalQuantity}?"
            : $"Сбросить количество для {overriddenItems.Count} выбранных элементов до исходных значений?";

        var result = MessageBox.Show(message, "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        foreach (var item in overriddenItems)
        {
            item.Quantity = item.OriginalQuantity;
        }
    }
}
