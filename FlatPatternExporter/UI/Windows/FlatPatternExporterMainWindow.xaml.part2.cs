using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using DxfRenderer;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using FlatPatternExporter.Converters;
using FlatPatternExporter.Enums;
using FlatPatternExporter.Services;
using Inventor;
using Svg.Skia;
using Binding = System.Windows.Data.Binding;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
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
        var partDoc = _scanService.DocumentCache.GetCachedPartDocument(partNumber) ?? _inventorService.OpenPartDocument(partNumber);
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


    private void SelectFixedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            FixedFolderPath = dialog.SelectedPath;
        }
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
        var exportOptions = CreateExportOptions();
        exportOptions.ShowFileLockedDialogs = true;
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = 0;
            var exportProgress = new Progress<double>(p => { });
            await Task.Run(() => _exportService.ExportDXF(tempPartsDataList, context.TargetDirectory, context.Multiplier,
                exportOptions, ref processedCount, ref skippedCount, context.GenerateThumbnails, exportProgress, _operationCts!.Token), _operationCts!.Token);
            
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
        var exportOptions = CreateExportOptions();
        var result = await ExecuteWithErrorHandlingAsync(async () =>
        {
            var processedCount = 0;
            var skippedCount = itemsWithoutFlatPattern.Count;
            var exportProgress = new Progress<double>(UpdateExportProgress);
            await Task.Run(() => _exportService.ExportDXF(selectedItems, context.TargetDirectory, context.Multiplier,
                exportOptions, ref processedCount, ref skippedCount, context.GenerateThumbnails, exportProgress, _operationCts!.Token), _operationCts!.Token);
            
            return CreateExportOperationResult(processedCount, skippedCount, stopwatch.Elapsed);
        }, "экспорта выделенных деталей");

        // Завершение операции
        CompleteOperation(result, OperationType.Export, ref _isExporting);
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
        _scanService.ConflictAnalyzer.Clear();
        ConflictFilesButton.IsEnabled = false;

        // Очищаем кеш документов
        _scanService.DocumentCache.ClearCache();

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
            var fullPath = _scanService.DocumentCache.GetCachedPartPath(partNumber) ?? _inventorService.GetPartDocumentFullPath(partNumber);

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
            var fullPath = _scanService.DocumentCache.GetCachedPartPath(partNumber) ?? _inventorService.GetPartDocumentFullPath(partNumber);
            var targetModelState = item.ModelState;

            // Проверка на null перед использованием fullPath
            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                continue;
            }

            // Открываем файл с указанием состояния модели
            _inventorService.OpenInventorDocument(fullPath, targetModelState);
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
            var partDoc = _scanService.DocumentCache.GetCachedPartDocument(partData.PartNumber) ?? _inventorService.OpenPartDocument(partData.PartNumber);
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
