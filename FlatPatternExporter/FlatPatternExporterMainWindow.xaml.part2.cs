using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Inventor;
using Binding = System.Windows.Data.Binding;
using Brushes = System.Windows.Media.Brushes;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Style = System.Windows.Style;

namespace FlatPatternExporter;

public partial class FlatPatternExporterMainWindow : Window
{
    private void ReadIPropertyValuesFromPart(PartDocument partDoc, PartData partData)
    {
        var propertySets = partDoc.PropertySets;
        // Считывание имени файла и полного пути
        partData.FullFileName = partDoc.FullFileName; // Полное имя файла с путем
        partData.FileName = Path.GetFileNameWithoutExtension(partDoc.FullFileName); // Имя файла без расширения

        // Набор свойств: Summary Information
        if (IsColumnPresent("Автор")) partData.Author = GetProperty(propertySets["Summary Information"], "Author");
        if (IsColumnPresent("Ревизия")) partData.Revision = GetProperty(propertySets["Summary Information"], "Revision Number");
        if (IsColumnPresent("Название")) partData.Title = GetProperty(propertySets["Summary Information"], "Title");
        if (IsColumnPresent("Тема")) partData.Subject = GetProperty(propertySets["Summary Information"], "Subject");
        if (IsColumnPresent("Ключевые слова")) partData.Keywords = GetProperty(propertySets["Summary Information"], "Keywords");
        if (IsColumnPresent("Примечание")) partData.Comments = GetProperty(propertySets["Summary Information"], "Comments");

        // Набор свойств: Document Summary Information
        if (IsColumnPresent("Категория")) partData.Category = GetProperty(propertySets["Document Summary Information"], "Category");
        if (IsColumnPresent("Менеджер")) partData.Manager = GetProperty(propertySets["Document Summary Information"], "Manager");
        if (IsColumnPresent("Компания")) partData.Company = GetProperty(propertySets["Document Summary Information"], "Company");

        // Набор свойств: Design Tracking Properties
        if (IsColumnPresent("Проект")) partData.Project = GetProperty(propertySets["Design Tracking Properties"], "Project");
        if (IsColumnPresent("Инвентарный номер")) partData.StockNumber = GetProperty(propertySets["Design Tracking Properties"], "Stock Number");
        if (IsColumnPresent("Время создания")) partData.CreationTime = GetProperty(propertySets["Design Tracking Properties"], "Creation Time");
        if (IsColumnPresent("Сметчик")) partData.CostCenter = GetProperty(propertySets["Design Tracking Properties"], "Cost Center");
        if (IsColumnPresent("Проверил")) partData.CheckedBy = GetProperty(propertySets["Design Tracking Properties"], "Checked By");
        if (IsColumnPresent("Нормоконтроль")) partData.EngApprovedBy = GetProperty(propertySets["Design Tracking Properties"], "Engr Approved By");
        if (IsColumnPresent("Статус")) partData.UserStatus = GetProperty(propertySets["Design Tracking Properties"], "User Status");
        if (IsColumnPresent("Веб-ссылка")) partData.CatalogWebLink = GetProperty(propertySets["Design Tracking Properties"], "Catalog Web Link");
        if (IsColumnPresent("Поставщик")) partData.Vendor = GetProperty(propertySets["Design Tracking Properties"], "Vendor");
        if (IsColumnPresent("Утвердил")) partData.MfgApprovedBy = GetProperty(propertySets["Design Tracking Properties"], "Mfg Approved By");
        if (IsColumnPresent("Статус разработки")) partData.DesignStatus = GetProperty(propertySets["Design Tracking Properties"], "Design Status");
        if (IsColumnPresent("Проектировщик")) partData.Designer = GetProperty(propertySets["Design Tracking Properties"], "Designer");
        if (IsColumnPresent("Инженер")) partData.Engineer = GetProperty(propertySets["Design Tracking Properties"], "Engineer");
        if (IsColumnPresent("Нач. отдела")) partData.Authority = GetProperty(propertySets["Design Tracking Properties"], "Authority");
        if (IsColumnPresent("Масса")) partData.Mass = GetProperty(propertySets["Design Tracking Properties"], "Mass");
        if (IsColumnPresent("Площадь поверхности")) partData.SurfaceArea = GetProperty(propertySets["Design Tracking Properties"], "SurfaceArea");
        if (IsColumnPresent("Объем")) partData.Volume = GetProperty(propertySets["Design Tracking Properties"], "Volume");
        if (IsColumnPresent("Правило ЛМ")) partData.SheetMetalRule = GetProperty(propertySets["Design Tracking Properties"], "Sheet Metal Rule");
        if (IsColumnPresent("Ширина развертки")) partData.FlatPatternWidth = GetProperty(propertySets["Design Tracking Properties"], "Flat Pattern Width");
        if (IsColumnPresent("Длинна развертки")) partData.FlatPatternLength = GetProperty(propertySets["Design Tracking Properties"], "Flat Pattern Length");
        if (IsColumnPresent("Площадь развертки")) partData.FlatPatternArea = GetProperty(propertySets["Design Tracking Properties"], "Flat Pattern Area");
        if (IsColumnPresent("Отделка")) partData.Appearance = GetProperty(propertySets["Design Tracking Properties"], "Appearance");
    }

    private async Task<PartData> GetPartDataAsync(string partNumber, int quantity, BOM bom, int itemNumber,
        PartDocument? partDoc = null)
    {
        // Открываем документ, если он не передан
        if (partDoc == null)
        {
            partDoc = OpenPartDocument(partNumber);
            if (partDoc == null) return null;
        }

        // Создаем новый объект PartData и заполняем его основными свойствами
        var partData = new PartData
        {
            PartNumber = await GetPropertyExpressionOrValueAsync(partDoc, "Part Number"),
            Description = await GetPropertyExpressionOrValueAsync(partDoc, "Description"),
            Material = GetMaterialForPart(partDoc),
            Thickness = GetThicknessForPart(partDoc).ToString("F1") + " мм", // Получаем толщину
            ModelState = partDoc.ModelStateName,
            OriginalQuantity = quantity,
            Quantity = quantity,
            QuantityColor = Brushes.Black,
            Item = itemNumber
        };

        // Получаем значения для пользовательских свойств
        foreach (var customProperty in _customPropertiesList)
        {
            // Асинхронно получаем значения пользовательских свойств
            partData.CustomProperties[customProperty] =
                await GetPropertyExpressionOrValueAsync(partDoc, customProperty);
        }

        // Асинхронно получаем превью-изображение (если доступно)
        partData.Preview = await GetThumbnailAsync(partDoc);

        // Проверяем наличие развертки для листовых деталей и задаем цвет
        var smCompDef = partDoc.ComponentDefinition as SheetMetalComponentDefinition;
        var hasFlatPattern = smCompDef != null && smCompDef.HasFlatPattern;
        partData.FlatPatternColor = hasFlatPattern ? Brushes.Green : Brushes.Red;

        // Чтение значений iProperty для Автор, Ревизия, Проект, Инвентарный номер
        ReadIPropertyValuesFromPart(partDoc, partData);

        return partData;
    }

    private BitmapImage GenerateDxfThumbnail(string dxfDirectory, string partNumber)
    {
        var searchPattern = partNumber + "*.dxf"; // Шаблон поиска
        var dxfFiles = Directory.GetFiles(dxfDirectory, searchPattern);

        if (dxfFiles.Length == 0) return null;

        try
        {
            var dxfFilePath = dxfFiles[0]; // Берем первый найденный файл, соответствующий шаблону
            var generator = new DxfThumbnailGenerator.DxfThumbnailGenerator();
            var bitmap = generator.GenerateThumbnail(dxfFilePath);

            BitmapImage bitmapImage = null;

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

            return bitmapImage;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show($"Ошибка при генерации миниатюры DXF: {ex.Message}", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
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
            BitmapImage bitmap = null;
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

            return bitmap;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show("Ошибка при получении миниатюры: " + ex.Message, "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
            return null;
        }
    }

    private void ProcessComponentOccurrences(ComponentOccurrences occurrences, Dictionary<string, int> sheetMetalParts)
    {
        foreach (ComponentOccurrence occ in occurrences)
        {
            if (_isCancelled) break;

            if (occ.Suppressed) continue;

            try
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

                if (ShouldExcludeComponent(occ.BOMStructure, fullFileName)) continue;

                if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var partDoc = occ.Definition.Document as PartDocument;
                    if (partDoc != null && partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                    {
                        var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                        if (!string.IsNullOrEmpty(partNumber))
                        {
                            if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                                sheetMetalParts[partNumber]++;
                            else
                                sheetMetalParts.Add(partNumber, 1);
                        }
                    }
                }
                else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    ProcessComponentOccurrences((ComponentOccurrences)occ.SubOccurrences, sheetMetalParts);
                }
            }
            catch (Exception ex)
            {
                // Логирование ошибки
                Debug.WriteLine($"Ошибка при обработке компонента: {ex.Message}");
            }
        }
    }
    private void ProcessBOM(BOM bom, Dictionary<string, int> sheetMetalParts)
    {
        try
        {
            BOMView? bomView = null;
            foreach (BOMView view in bom.BOMViews)
            {
                bomView = view;
                break;
            }

            if (bomView == null)
            {
                MessageBox.Show("Не найдено ни одного представления спецификации.", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            BOMRow[] bomRows;
            try
            {
                bomRows = bomView.BOMRows.Cast<BOMRow>().ToArray();
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
            {
                _hasMissingReferences = true;
                Debug.WriteLine("Ошибка при доступе к BOMRows. Возможно, в сборке есть потерянные ссылки.");
                return;
            }

            foreach (BOMRow row in bomRows)
            {
                if (_isCancelled) break;

                try
                {
                    ProcessBOMRow(row, sheetMetalParts);
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

    private void ProcessBOMRow(BOMRow row, Dictionary<string, int> sheetMetalParts)
    {
        var componentDefinition = row.ComponentDefinitions[1];
        if (componentDefinition == null) return;

        var document = componentDefinition.Document as Document;
        if (document == null) return;

        if (ShouldExcludeComponent(row.BOMStructure, document.FullFileName)) return;

        if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = document as PartDocument;
            if (partDoc != null && partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                if (!string.IsNullOrEmpty(partNumber))
                {
                    if (sheetMetalParts.TryGetValue(partNumber, out var quantity))
                        sheetMetalParts[partNumber] += row.ItemQuantity;
                    else
                        sheetMetalParts.Add(partNumber, row.ItemQuantity);
                }
            }
        }
        else if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = document as AssemblyDocument;
            if (asmDoc != null) ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts);
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
    private void UpdateIncludeLibraryComponentsState(object sender, RoutedEventArgs e)
    {
        _includeLibraryComponents = IncludeLibraryComponentsCheckBox.IsChecked ?? false;
    }
    private bool IsLibraryComponent(string fullFileName)
    {
        return _libraryPaths.Any(path => fullFileName.StartsWith(path, StringComparison.OrdinalIgnoreCase));
    }
    private bool ShouldExcludeComponent(BOMStructureEnum bomStructure, string fullFileName)
    {
        if (_excludeReferenceParts && bomStructure == BOMStructureEnum.kReferenceBOMStructure)
            return true;

        if (_excludePurchasedParts && bomStructure == BOMStructureEnum.kPurchasedBOMStructure)
            return true;

        if (!_includeLibraryComponents && !string.IsNullOrEmpty(fullFileName) && IsLibraryComponent(fullFileName))
            return true;

        return false;
    }
    private void UpdateExcludeReferencePartsState(object sender, RoutedEventArgs e)
    {
        _excludeReferenceParts = ExcludeReferencePartsCheckBox.IsChecked ?? true;
    }

    private void UpdateExcludePurchasedPartsState(object sender, RoutedEventArgs e)
    {
        _excludePurchasedParts = ExcludePurchasedPartsCheckBox.IsChecked ?? true;
    }
    private void SelectFixedFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            _fixedFolderPath = dialog.SelectedPath;

            if (_fixedFolderPath.Length > 40)
                // Гарантируем, что показываем именно последние 40 символов
                FixedFolderPathTextBlock.Text = "..." + _fixedFolderPath.Substring(_fixedFolderPath.Length - 40);
            else
                FixedFolderPathTextBlock.Text = _fixedFolderPath;
        }
    }

private bool PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch)
{
    stopwatch = new Stopwatch();
    targetDir = string.Empty;
    multiplier = 1; // Присваиваем значение по умолчанию

    // Обрабатываем выбор радио-кнопки
    if (ChooseFolderRadioButton.IsChecked == true)
    {
        // Открываем диалог выбора папки
        var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            targetDir = dialog.SelectedPath;
        else
            return false;
    }
    else if (ComponentFolderRadioButton.IsChecked == true)
    {
        // Папка компонента
        targetDir = Path.GetDirectoryName(_thisApplication.ActiveDocument.FullFileName);
    }
    else if (FixedFolderRadioButton.IsChecked == true)
    {
        if (string.IsNullOrEmpty(_fixedFolderPath))
        {
            MessageBox.Show("Выберите фиксированную папку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }

        targetDir = _fixedFolderPath;
    }
    else if (ProjectFolderRadioButton.IsChecked == true) // Новый код для выбора папки проекта
    {
        try
        {
            targetDir = _thisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath;
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
    }

    if (EnableSubfolderCheckBox.IsChecked == true && !string.IsNullOrEmpty(SubfolderNameTextBox.Text))
    {
        targetDir = Path.Combine(targetDir, SubfolderNameTextBox.Text);
        if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
    }

    if (IncludeQuantityInFileNameCheckBox.IsChecked == true && !int.TryParse(MultiplierTextBox.Text, out multiplier))
    {
        MessageBox.Show("Введите допустимое целое число для множителя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return false;
    }

    CancelButton.IsEnabled = true;
    _isCancelled = false;

    ProgressBar.IsIndeterminate = false;
    ProgressBar.Value = 0;
    ProgressLabel.Text = "Статус: Экспорт данных...";

    stopwatch = Stopwatch.StartNew();

    return true;
}

    private void FinalizeExport(bool isCancelled, Stopwatch stopwatch, int processedCount, int skippedCount)
    {
        stopwatch.Stop();
        var elapsedTime = GetElapsedTime(stopwatch.Elapsed);

        CancelButton.IsEnabled = false;
        ScanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта
        ClearButton.IsEnabled = _partsData.Count > 0; // Активируем кнопку "Очистить" после завершения экспорта

        if (isCancelled)
        {
            ProgressLabel.Text = $"Статус: Прервано ({elapsedTime})";
            MessageBox.Show("Процесс экспорта был прерван.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        else
        {
            ProgressLabel.Text = $"Статус: Завершено ({elapsedTime})";
            MessageBox.Show(
                $"Экспорт DXF завершен.\nВсего файлов обработано: {processedCount + skippedCount}\nПропущено (без разверток): {skippedCount}\nВсего экспортировано: {processedCount}\nВремя выполнения: {elapsedTime}",
                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private async Task ExportWithoutScan()
    {
        // Очистка таблицы перед началом скрытого экспорта
        ClearList_Click(this, null);

        if (_thisApplication == null)
        {
            MessageBox.Show("Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        Document doc = _thisApplication.ActiveDocument;
        if (doc == null)
        {
            MessageBox.Show("Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch)) return;

        ClearButton.IsEnabled = false;
        ScanButton.IsEnabled = false;

        var processedCount = 0;
        var skippedCount = 0;

        // Скрытое сканирование
        var sheetMetalParts = new Dictionary<string, int>();

        if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = (AssemblyDocument)doc;
            if (TraverseRadioButton.IsChecked == true)
                await Task.Run(() =>
                    ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
            else if (BomRadioButton.IsChecked == true)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));

            // Создаем временный список PartData для экспорта
            var tempPartsDataList = new List<PartData>();
            foreach (var part in sheetMetalParts)
                tempPartsDataList.Add(new PartData
                {
                    PartNumber = part.Key,
                    OriginalQuantity = part.Value,
                    Quantity = part.Value * multiplier
                });

            await Task.Run(() =>
                ExportDXF(tempPartsDataList, targetDir, multiplier, ref processedCount, ref skippedCount,
                    false)); // false - без генерации миниатюр
        }
        else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)doc;
            if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                var tempPartData = new List<PartData>
                {
                    new()
                    {
                        PartNumber = partNumber,
                        OriginalQuantity = 1,
                        Quantity = multiplier
                    }
                };
                await Task.Run(() =>
                    ExportDXF(tempPartData, targetDir, multiplier, ref processedCount, ref skippedCount,
                        false)); // false - без генерации миниатюр
            }
            else
            {
                MessageBox.Show("Активный документ не является листовым металлом.", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        ClearButton.IsEnabled = true;
        ScanButton.IsEnabled = true;

        // Здесь мы НЕ обновляем fileCountLabelBottom
        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);

        // Деактивируем кнопку "Экспорт" после завершения скрытого экспорта
        ExportButton.IsEnabled = false;
    }


    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {

        if (_isCtrlPressed)
        {
            await ExportWithoutScan();
            return;
        }

        if (_thisApplication == null)
        {
            MessageBox.Show("Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        SetInventorUserInterfaceState(true);

        Document? doc = _thisApplication.ActiveDocument;
        if (doc == null)
        {
            MessageBox.Show("Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова.", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (doc != _lastScannedDocument)
        {
            MessageBox.Show("Активный документ изменился. Пожалуйста, повторите сканирование перед экспортом.",
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            MultiplierTextBox.Text = "1";
            return;
        }

        if (!PrepareForExport(out var targetDir, out var multiplier, out var stopwatch))
        {
            SetInventorUserInterfaceState(false);
            return;
        }

        ClearButton.IsEnabled = false;
        ScanButton.IsEnabled = false;

        var processedCount = 0;
        var skippedCount = 0;

        if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
        {
            var asmDoc = (AssemblyDocument)doc;
            var sheetMetalParts = new Dictionary<string, int>();
            if (TraverseRadioButton.IsChecked == true)
                await Task.Run(() =>
                    ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
            else if (BomRadioButton.IsChecked == true)
                await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));

            await ExportDXFFiles(sheetMetalParts, targetDir, multiplier, stopwatch);
        }
        else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
        {
            var partDoc = (PartDocument)doc;
            if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                await ExportSinglePartDXF(partDoc, targetDir, multiplier, stopwatch);
            else
                MessageBox.Show("Активный документ не является листовым металлом.", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
        }

        ClearButton.IsEnabled = true;
        ScanButton.IsEnabled = true;

        SetInventorUserInterfaceState(false);

        ExportButton.IsEnabled = true;
    }

    private string GetElapsedTime(TimeSpan timeSpan)
    {
        if (timeSpan.TotalMinutes >= 1)
            return $"{(int)timeSpan.TotalMinutes} мин. {timeSpan.Seconds} сек.";
        return $"{timeSpan.Seconds} сек.";
    }

    private async void ExportSelectedDXF_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для экспорта.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var itemsWithoutFlatPattern = selectedItems.Where(p => p.FlatPatternColor == Brushes.Red).ToList();
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

        ClearButton.IsEnabled = false; // Деактивируем кнопку "Очистить" в процессе экспорта
        ScanButton.IsEnabled = false; // Деактивируем кнопку "Сканировать" в процессе экспорта

        var processedCount = 0;
        var skippedCount = itemsWithoutFlatPattern.Count;

        await Task.Run(() =>
        {
            ExportDXF(selectedItems, targetDir, multiplier, ref processedCount, ref skippedCount,
                true); // true - с генерацией миниатюр
        });

        UpdateFileCountLabel(selectedItems.Count);

        ClearButton.IsEnabled = true; // Активируем кнопку "Очистить" после завершения экспорта
        ScanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта
        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);
    }

    private void OverrideQuantity_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для переопределения количества.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var dialog = new OverrideQuantityDialog();
        dialog.Owner = this;
        dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

        if (dialog.ShowDialog() == true && dialog.NewQuantity.HasValue)
        {
            foreach (var item in selectedItems)
            {
                item.Quantity = dialog.NewQuantity.Value;
                item.IsOverridden = true;
                item.QuantityColor = Brushes.Red; // Переопределенное количество окрашивается в красный цвет
            }

            PartsDataGrid.Items.Refresh();
        }
    }

    private void ExportDXF(IEnumerable<PartData> partsDataList, string targetDir, int multiplier,
        ref int processedCount, ref int skippedCount, bool generateThumbnails)
    {
        var organizeByMaterial = false;
        var organizeByThickness = false;
        var includeQuantityInFileName = false;

        // Выполняем доступ к элементам управления через Dispatcher
        Dispatcher.Invoke(() =>
        {
            organizeByMaterial = OrganizeByMaterialCheckBox.IsChecked == true;
            organizeByThickness = OrganizeByThicknessCheckBox.IsChecked == true;
            includeQuantityInFileName = IncludeQuantityInFileNameCheckBox.IsChecked == true;
        });

        var totalParts = partsDataList.Count();

        Dispatcher.Invoke(() =>
        {
            ProgressBar.Maximum = totalParts;
            ProgressBar.Value = 0;
        });

        var localProcessedCount = processedCount;
        var localSkippedCount = skippedCount;

        foreach (var partData in partsDataList)
        {
            if (_isCancelled) break;

            var partNumber = partData.PartNumber;
            var qty = partData.IsOverridden ? partData.Quantity : partData.OriginalQuantity * multiplier;

            // Доступ к RadioButton через Dispatcher для потока безопасности
            Dispatcher.Invoke(() =>
            {
                if (PartFolderRadioButton.IsChecked == true)
                {
                    targetDir = GetPartDocumentFullPath(partNumber);
                    targetDir = Path.GetDirectoryName(targetDir);
                }
            });

            PartDocument partDoc = null;
            var isPartDocOpenedHere = false;
            try
            {
                partDoc = OpenPartDocument(partNumber);
                if (partDoc == null) throw new Exception("Файл детали не найден или не может быть открыт");

                var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                var material = GetMaterialForPart(partDoc);
                var thickness = GetThicknessForPart(partDoc);

                var materialDir = organizeByMaterial ? Path.Combine(targetDir, material) : targetDir;
                if (!Directory.Exists(materialDir)) Directory.CreateDirectory(materialDir);

                var thicknessDir = organizeByThickness
                    ? Path.Combine(materialDir, thickness.ToString("F1"))
                    : materialDir;
                if (!Directory.Exists(thicknessDir)) Directory.CreateDirectory(thicknessDir);

                var filePath = Path.Combine(thicknessDir,
                    partNumber + (includeQuantityInFileName ? " - " + qty : "") + ".dxf");

                if (!IsValidPath(filePath)) continue;

                var exportSuccess = false;

                if (smCompDef.HasFlatPattern)
                {
                    var flatPattern = smCompDef.FlatPattern;
                    var oDataIO = flatPattern.DataIO;

                    try
                    {
                        string options = null;
                        Dispatcher.Invoke(() => PrepareExportOptions(out options));

                        // Интеграция настроек слоев
                        var layerOptionsBuilder = new StringBuilder();
                        var invisibleLayersBuilder = new StringBuilder();

                        foreach (var layer in LayerSettings) // Используем настройки слоев
                        {
                            if (layer.HasVisibilityOption && !layer.IsVisible)
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
                                _isCancelled = true; // Прерываем операцию
                                break;
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
                                    _isCancelled = true; // Прерываем операцию
                                    break;
                                }

                                Thread.Sleep(1000); // Ожидание 1 секунды перед повторной проверкой
                            }

                            if (_isCancelled) break;
                        }

                        oDataIO.WriteDataToFile(dxfOptions, filePath);
                        exportSuccess = true;
                    }
                    catch (Exception ex)
                    {
                        // Обработка ошибок экспорта DXF
                        Debug.WriteLine($"Ошибка экспорта DXF: {ex.Message}");
                    }
                }

                BitmapImage dxfPreview = null;
                if (generateThumbnails) dxfPreview = GenerateDxfThumbnail(thicknessDir, partNumber);

                Dispatcher.Invoke(() =>
                {
                    partData.ProcessingColor = exportSuccess ? Brushes.Green : Brushes.Red;
                    partData.DxfPreview = dxfPreview;
                    PartsDataGrid.Items.Refresh();
                });

                if (exportSuccess)
                    localProcessedCount++;
                else
                    localSkippedCount++;

                Dispatcher.Invoke(() => { ProgressBar.Value = Math.Min(localProcessedCount, ProgressBar.Maximum); });
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
            ProgressBar.Value = ProgressBar.Maximum; // Установка значения 100% по завершению
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

    private async Task ExportDXFFiles(Dictionary<string, int> sheetMetalParts, string targetDir, int multiplier,
        Stopwatch stopwatch)
    {
        var processedCount = 0;
        var skippedCount = 0;

        var partsDataList = _partsData.Where(p => sheetMetalParts.ContainsKey(p.PartNumber)).ToList();
        await Task.Run(
            () => ExportDXF(partsDataList, targetDir, multiplier, ref processedCount, ref skippedCount,
                true)); // true - с генерацией миниатюр

        UpdateFileCountLabel(partsDataList.Count);
        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);
    }

    private async Task ExportSinglePartDXF(PartDocument partDoc, string targetDir, int multiplier, Stopwatch stopwatch)
    {
        var processedCount = 0;
        var skippedCount = 0;

        var partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
        var partData = _partsData.FirstOrDefault(p => p.PartNumber == partNumber);

        if (partData != null)
        {
            var partsDataList = new List<PartData> { partData };
            await Task.Run(() =>
                ExportDXF(partsDataList, targetDir, multiplier, ref processedCount, ref skippedCount,
                    true)); // true - с генерацией миниатюр
            UpdateFileCountLabel(1);
        }

        FinalizeExport(_isCancelled, stopwatch, processedCount, skippedCount);
    }

    private PartDocument OpenPartDocument(string partNumber)
    {
        var docs = _thisApplication.Documents;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd &&
                GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                return pd; // Возвращаем найденный открытый документ

        // Если документ не найден
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    private string GetMaterialForPart(PartDocument partDoc)
    {
        try
        {
            return partDoc.ComponentDefinition.Material.Name;
        }
        catch (Exception)
        {
            // Обработка ошибок получения материала
        }

        return "Ошибка получения имени";
    }

    private double GetThicknessForPart(PartDocument partDoc)
    {
        try
        {
            var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
            var thicknessParam = smCompDef.Thickness;
            return
                Math.Round((double)thicknessParam.Value * 10,
                    1); // Извлекаем значение толщины и переводим в мм, округляем до одной десятой
        }
        catch (Exception)
        {
            // Обработка ошибок получения толщины
        }

        return 0;
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

    private void IncludeQuantityInFileNameCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
    {
        MultiplierTextBox.IsEnabled = IncludeQuantityInFileNameCheckBox.IsChecked == true;
        if (!MultiplierTextBox.IsEnabled)
        {
            MultiplierTextBox.Text = "1";
            UpdateQuantitiesWithMultiplier(1);
        }
    }

    private void MultiplierTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (int.TryParse(MultiplierTextBox.Text, out var multiplier) && multiplier > 0)
        {
            UpdateQuantitiesWithMultiplier(multiplier);

            // Проверка на null перед изменением видимости кнопки
            if (ClearMultiplierButton != null)
                ClearMultiplierButton.Visibility = multiplier > 1 ? Visibility.Visible : Visibility.Collapsed;
        }
        else
        {
            // Если введенное значение некорректное, сбрасываем текст на "1" и скрываем кнопку сброса
            MultiplierTextBox.Text = "1";
            UpdateQuantitiesWithMultiplier(1);

            if (ClearMultiplierButton != null) ClearMultiplierButton.Visibility = Visibility.Collapsed;
        }
    }

    private void ClearMultiplierButton_Click(object sender, RoutedEventArgs e)
    {
        // Сбрасываем множитель на 1
        MultiplierTextBox.Text = "1";
        UpdateQuantitiesWithMultiplier(1);

        // Скрываем кнопку сброса
        ClearMultiplierButton.Visibility = Visibility.Collapsed;
    }

    private void UpdateFileCountLabel(int count)
    {
        FileCountLabelBottom.Text = $"Найдено листовых деталей: {count}";
    }


    private void ResetProgressBar()
    {
        ProgressBar.IsIndeterminate = false;
        ProgressBar.Minimum = 0;
        ProgressBar.Maximum = 100;
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Статус: ";
        UpdateFileCountLabel(0); // Использование метода для сброса счетчика
        BottomPanel.Opacity = 0.75;
        DocumentInfoLabel.Text = string.Empty;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        _isCancelled = true;
        CancelButton.IsEnabled = false;

        if (!_isScanning)
        {
            // If not scanning, reset the UI
            ResetProgressBar();
            BottomPanel.Opacity = 0.75;
            DocumentInfoLabel.Text = "";
            ExportButton.IsEnabled = false;
            ClearButton.IsEnabled = _partsData.Count > 0; // Обновляем состояние кнопки "Очистить"
            _lastScannedDocument = null;
            UpdateFileCountLabel(0);
        }
    }

    private void UpdateDocumentInfo(string documentType, string partNumber, string description, Document doc)
    {
        if (string.IsNullOrEmpty(documentType) && string.IsNullOrEmpty(partNumber) && string.IsNullOrEmpty(description))
        {
            DocumentInfoLabel.Text = string.Empty;
            ModelStateInfoRunBottom.Text = string.Empty; // Обновленное имя элемента
            BottomPanel.Opacity = 0.75;
            UpdateFileCountLabel(0); // Использование метода для сброса счетчика
            return;
        }

        var modelStateInfo = GetModelStateName(doc);

        DocumentInfoLabel.Text = $"Тип документа: {documentType}\nОбозначение: {partNumber}\nОписание: {description}";

        if (modelStateInfo == "[Primary]" || modelStateInfo == "[Основной]")
        {
            ModelStateInfoRunBottom.Text = "[Основной]";
            ModelStateInfoRunBottom.Foreground =
                new SolidColorBrush(Colors.Black); // Цвет текста - черный (по умолчанию)
        }
        else
        {
            ModelStateInfoRunBottom.Text = modelStateInfo;
            ModelStateInfoRunBottom.Foreground = new SolidColorBrush(Colors.Red); // Цвет текста - красный
        }

        UpdateFileCountLabel(_partsData.Count); // Использование метода для обновления счетчика
        BottomPanel.Opacity = 1.0;
    }

    private void ClearList_Click(object sender, RoutedEventArgs e)
    {
        // Очищаем данные в таблице или любом другом источнике данных
        _partsData.Clear();
        ProgressBar.Value = 0;
        ProgressLabel.Text = "Статус: ";
        DocumentInfoLabel.Text = string.Empty;
        ModelStateInfoRunBottom.Text = string.Empty;
        BottomPanel.Opacity = 0.5;
        ExportButton.IsEnabled = false;
        ClearButton.IsEnabled = false; // Делаем кнопку "Очистить" неактивной после очистки

        // Обнуляем информацию о документе
        UpdateDocumentInfo(string.Empty, string.Empty, string.Empty, null);
        UpdateFileCountLabel(0); // Сброс счетчика файлов

        // Отключаем кнопку "Анализ обозначений" и очищаем список конфликтов
        ConflictFilesButton.IsEnabled = false;
        _conflictFileDetails.Clear(); // Очищаем список с конфликтами (или используем нужную коллекцию)

    }

    private void partsDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_partsData.Count == 0) e.Handled = true; // Предотвращаем открытие контекстного меню, если таблица пуста
    }

    private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = PartsDataGrid.SelectedItems.Cast<PartData>().ToList();
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для открытия расположения файла.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

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
        if (selectedItems.Count == 0)
        {
            MessageBox.Show("Выберите строки для открытия моделей.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

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
                    // Открытие документа в Inventor
                    _thisApplication.Documents.Open(fullPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии файла по пути {fullPath}: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            else
                MessageBox.Show($"Файл по пути {fullPath}, связанный с номером детали {partNumber}, не найден.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string GetPartDocumentFullPath(string partNumber)
    {
        var docs = _thisApplication.Documents;

        // Попытка найти открытый документ
        foreach (Document doc in docs)
            if (doc is PartDocument pd &&
                GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                return pd.FullFileName; // Возвращаем полный путь найденного документа

        // Если документ не найден среди открытых
        MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка",
            MessageBoxButton.OK, MessageBoxImage.Error);
        return null; // Возвращаем null, если документ не найден
    }

    private void partsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Активация пункта "Редактировать iProperty" только при выборе одной строки
        EditIPropertyMenuItem.IsEnabled = PartsDataGrid.SelectedItems.Count == 1;
    }

    public async Task FillCustomPropertyAsync(string customPropertyName)
    {
        foreach (var partData in _partsData)
        {
            var value = await Task.Run(() => GetCustomIPropertyValue(partData.PartNumber, customPropertyName));

            partData.AddCustomProperty(customPropertyName, value);

            PartsDataGrid.Items.Refresh();

            await Task.Delay(50);
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
        _customPropertiesList.Remove(propertyName); // Не забудьте удалить свойство из списка доступных свойств
        PartsDataGrid.Items.Refresh();
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

        // Создаем новую колонку
        var column = new DataGridTextColumn
        {
            Header = propertyName,
            Binding = new Binding($"CustomProperties[{propertyName}]")
            {
                Mode = BindingMode.OneWay // Устанавливаем режим привязки
            },
            Width = DataGridLength.Auto
        };

        // Применяем стиль CenteredCellStyle
        column.ElementStyle = PartsDataGrid.FindResource("CenteredCellStyle") as Style;

        PartsDataGrid.Columns.Add(column);
    }

    private void EditIProperty_Click(object sender, RoutedEventArgs e)
    {
        var selectedItem = PartsDataGrid.SelectedItem as PartData;
        if (selectedItem == null || PartsDataGrid.SelectedItems.Count != 1)
        {
            MessageBox.Show("Выберите одну строку для редактирования iProperty.", "Информация", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var editDialog = new EditIPropertyDialog(selectedItem.PartNumber, selectedItem.Description);
        if (editDialog.ShowDialog() == true)
        {
            // Открываем документ детали
            var partDoc = OpenPartDocument(selectedItem.PartNumber);
            if (partDoc != null)
            {
                // Обновляем iProperty в файле детали
                SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number", editDialog.PartNumber);
                SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description", editDialog.Description);

                // Сохраняем изменения и закрываем документ
                partDoc.Save2();

                // Обновляем свойства детали в таблице
                selectedItem.PartNumber = editDialog.PartNumber;
                selectedItem.Description = editDialog.Description;

                // Обновляем данные в таблице
                PartsDataGrid.Items.Refresh();
            }
            else
            {
                MessageBox.Show("Не удалось открыть документ детали для редактирования.", "Ошибка", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    private void SetProperty(PropertySet propertySet, string propertyName, string value)
    {
        try
        {
            var property = propertySet[propertyName];
            property.Value = value;
        }
        catch (Exception)
        {
            MessageBox.Show($"Не удалось обновить свойство {propertyName}.", "Ошибка", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private string GetCustomIPropertyValue(string partNumber, string propertyName)
    {
        try
        {
            var partDoc = OpenPartDocument(partNumber);
            if (partDoc != null)
            {
                var propertySet = partDoc.PropertySets["Inventor User Defined Properties"];

                if (propertySet[propertyName] is Property property)
                {
                    string value = property.Value?.ToString() ?? string.Empty;
                    return value;
                }
            }
        }
        catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
        {
            // Ловим ошибку и возвращаем пустую строку
            return string.Empty;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при получении свойства {propertyName}: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        return string.Empty;
    }
    private void AboutButton_Click(object sender, RoutedEventArgs e)
    {
        var aboutWindow = new AboutWindow(this); // Передаем текущий экземпляр MainWindow
        aboutWindow.ShowDialog(); // Отображаем окно как диалоговое
    }
}