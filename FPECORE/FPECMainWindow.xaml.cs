using DefineEdge;
using Inventor;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using LayerSettingsApp;

namespace FPECORE
{
    public partial class MainWindow : Window
    {
        private Inventor.Application ThisApplication;
        private bool isCancelled = false;
        private bool isScanning = false;
        private Document? lastScannedDocument = null;
        private ObservableCollection<PartData> partsData = new ObservableCollection<PartData>();
        private int itemCounter = 1; // Инициализация счетчика пунктов

        public ObservableCollection<LayerSetting> LayerSettings { get; set; }
        public ObservableCollection<string> AvailableColors { get; set; }
        public ObservableCollection<string> LineTypes { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            InitializeInventor();
            partsDataGrid.ItemsSource = partsData;
            multiplierTextBox.IsEnabled = includeQuantityInFileNameCheckBox.IsChecked == true;
            UpdateFileCountLabel(0);
            clearButton.IsEnabled = false;

            // Создаем экземпляр MainWindow из LayerSettingsApp
            var layerSettingsWindow = new LayerSettingsApp.MainWindow();

            // Используем его данные для настройки LayerSettings
            LayerSettings = layerSettingsWindow.LayerSettings;
            AvailableColors = layerSettingsWindow.AvailableColors;
            LineTypes = layerSettingsWindow.LineTypes;

            // Устанавливаем DataContext для текущего окна, объединяя данные из LayerSettingsWindow и других источников
            DataContext = this;

            // Добавляем обработчики для отслеживания нажатия и отпускания клавиши Ctrl
            this.KeyDown += new System.Windows.Input.KeyEventHandler(MainWindow_KeyDown);
            this.KeyUp += new System.Windows.Input.KeyEventHandler(MainWindow_KeyUp);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Закрытие всех дочерних окон
            foreach (Window window in System.Windows.Application.Current.Windows)
            {
                if (window != this)
                {
                    window.Close();
                }
            }

            // Если были созданы дополнительные потоки, убедитесь, что они завершены
        }
        
        private void InitializeInventor()
        {
            try
            {
                ThisApplication = (Inventor.Application)MarshalCore.GetActiveObject("Inventor.Application") ?? throw new InvalidOperationException("Не удалось получить активный объект Inventor.");
            }
            catch (COMException)
            {
                System.Windows.MessageBox.Show("Не удалось подключиться к запущенному экземпляру Inventor. Убедитесь, что Inventor запущен.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Произошла ошибка при подключении к Inventor: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
            }
        }

        private bool isCtrlPressed = false;

        private void MainWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.LeftCtrl || e.Key == System.Windows.Input.Key.RightCtrl)
            {
                isCtrlPressed = true;
                exportButton.IsEnabled = true;
            }
        }

        private void MainWindow_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.LeftCtrl || e.Key == System.Windows.Input.Key.RightCtrl)
            {
                isCtrlPressed = false;
                exportButton.IsEnabled = partsData.Count > 0; // Восстанавливаем исходное состояние кнопки
            }
        }

        private void EnableSplineReplacementCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool isChecked = enableSplineReplacementCheckBox.IsChecked == true;
            splineReplacementComboBox.IsEnabled = isChecked;
            splineToleranceTextBox.IsEnabled = isChecked;

            if (isChecked && splineReplacementComboBox.SelectedIndex == -1)
            {
                splineReplacementComboBox.SelectedIndex = 0; // По умолчанию выбираем "Линии"
            }
        }
        private void PrepareExportOptions(out string options)
        {
            var sb = new StringBuilder();

            sb.Append($"AcadVersion={((ComboBoxItem)acadVersionComboBox.SelectedItem).Content}");

            if (enableSplineReplacementCheckBox.IsChecked == true)
            {
                string splineTolerance = splineToleranceTextBox.Text;

                // Получаем текущий разделитель дробной части из системных настроек
                char decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];

                // Заменяем разделитель на актуальный
                splineTolerance = splineTolerance.Replace('.', decimalSeparator).Replace(',', decimalSeparator);

                if (splineReplacementComboBox.SelectedIndex == 0) // Линии
                {
                    sb.Append($"&SimplifySplines=True&SplineTolerance={splineTolerance}");
                }
                else if (splineReplacementComboBox.SelectedIndex == 1) // Дуги
                {
                    sb.Append($"&SimplifySplines=True&SimplifyAsTangentArcs=True&SplineTolerance={splineTolerance}");
                }
            }
            else
            {
                sb.Append("&SimplifySplines=False");
            }

            if (mergeProfilesIntoPolylineCheckBox.IsChecked == true)
            {
                sb.Append("&MergeProfilesIntoPolyline=True");
            }

            if (rebaseGeometryCheckBox.IsChecked == true)
            {
                sb.Append("&RebaseGeometry=True");
            }

            if (trimCenterlinesCheckBox.IsChecked == true) // Проверяем новое поле TrimCenterlinesAtContour
            {
                sb.Append("&TrimCenterlinesAtContour=True");
            }

            options = sb.ToString();
        }
        private void MultiplierTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static bool IsTextAllowed(string text)
        {
            Regex regex = new Regex("^[1-9][0-9]*$"); // Разрешены только положительные числа, исключая ноль
            return regex.IsMatch(text);
        }

        private void UpdateQuantitiesWithMultiplier(int multiplier)
        {
            foreach (var partData in partsData)
            {
                partData.IsOverridden = false; // Сброс переопределённых значений
                partData.Quantity = partData.OriginalQuantity * multiplier;
                partData.QuantityColor = multiplier > 1 ? System.Windows.Media.Brushes.Blue : System.Windows.Media.Brushes.Black;
            }
            partsDataGrid.Items.Refresh();
        }

        private async void ScanButton_Click(object sender, RoutedEventArgs e)
        {
            if (ThisApplication == null)
            {
                MessageBoxHelper.ShowInventorNotRunningError();
                return;
            }

            modelStateInfoRunBottom.Text = string.Empty;

            Document? doc = ThisApplication.ActiveDocument as Document;
            if (doc == null)
            {
                MessageBoxHelper.ShowNoDocumentOpenWarning();
                return;
            }

            if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject && doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
            {
                System.Windows.MessageBox.Show("Откройте сборку или деталь для сканирования.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ResetProgressBar();
            progressBar.IsIndeterminate = true;
            progressLabel.Text = "Статус: Сбор данных...";
            scanButton.IsEnabled = false;
            cancelButton.IsEnabled = true;
            exportButton.IsEnabled = false;
            clearButton.IsEnabled = false; // Кнопка "Очистить" неактивна в процессе сканирования
            isScanning = true;
            isCancelled = false;

            var stopwatch = Stopwatch.StartNew();

            await Task.Delay(100);

            int partCount = 0;
            string documentType = doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject ? "Сборка" : "Деталь";
            string partNumber = string.Empty;
            string description = string.Empty;
            string modelStateInfo = string.Empty;

            partsData.Clear();
            itemCounter = 1; // Сброс счетчика

            var progress = new Progress<PartData>(partData =>
            {
                partData.Item = itemCounter; // Присвоение значения Item как int
                partsData.Add(partData);
                partsDataGrid.Items.Refresh();
                itemCounter++; // Увеличение счетчика после добавления новой строки
            });

            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = (AssemblyDocument)doc;
                var sheetMetalParts = new Dictionary<string, int>();
                if (traverseRadioButton.IsChecked == true)
                {
                    await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
                }
                else if (bomRadioButton.IsChecked == true)
                {
                    await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));
                }
                partCount = sheetMetalParts.Count;
                partNumber = GetProperty(asmDoc.PropertySets["Design Tracking Properties"], "Part Number");
                description = GetProperty(asmDoc.PropertySets["Design Tracking Properties"], "Description");
                modelStateInfo = asmDoc.ComponentDefinition.BOM.BOMViews[1].ModelStateMemberName;

                int itemCounter = 1;
                await Task.Run(async () =>
                {
                    foreach (var part in sheetMetalParts)
                    {
                        if (isCancelled) break;

                        var partData = await GetPartDataAsync(part.Key, part.Value, asmDoc.ComponentDefinition.BOM, itemCounter++);
                        if (partData != null)
                        {
                            ((IProgress<PartData>)progress).Report(partData);
                            await Task.Delay(10); // Увеличиваем время задержки для более плавной анимации
                        }
                    }
                });
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument partDoc = (PartDocument)doc;
                if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    partCount = 1;
                    partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                    description = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description");

                    var partData = await GetPartDataAsync(partNumber, 1, null, 1, partDoc);
                    if (partData != null)
                    {
                        ((IProgress<PartData>)progress).Report(partData);
                        await Task.Delay(10); // Увеличиваем время задержки для более плавной анимации
                    }
                }
            }

            stopwatch.Stop();
            string elapsedTime = GetElapsedTime(stopwatch.Elapsed);

            // Применение множителя после сканирования
            if (int.TryParse(multiplierTextBox.Text, out int multiplier) && multiplier > 0)
            {
                UpdateQuantitiesWithMultiplier(multiplier);
            }

            isScanning = false;
            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;
            progressLabel.Text = isCancelled ? $"Статус: Прервано ({elapsedTime})" : $"Статус: Готово ({elapsedTime})";
            bottomPanel.Opacity = 1.0;
            UpdateDocumentInfo(documentType, partNumber, description, doc);
            UpdateFileCountLabel(isCancelled ? 0 : partCount); // Обновленное использование метода для обновления счетчика

            scanButton.IsEnabled = true;
            cancelButton.IsEnabled = false;
            exportButton.IsEnabled = partCount > 0 && !isCancelled;
            clearButton.IsEnabled = partsData.Count > 0; // Активируем кнопку "Очистить" при наличии данных
            lastScannedDocument = isCancelled ? null : doc;

            if (isCancelled)
            {
                System.Windows.MessageBox.Show("Процесс сканирования был прерван.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private string GetProperty(PropertySet propertySet, string propertyName)
        {
            try
            {
                Property property = propertySet[propertyName];
                return property.Value.ToString();
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private async Task<PartData> GetPartDataAsync(string partNumber, int quantity, BOM bom, int itemNumber, PartDocument? partDoc = null)
        {
            if (partDoc == null)
            {
                partDoc = OpenPartDocument(partNumber);
                if (partDoc == null) return null;
            }

            string modelState = partDoc.ModelStateName;  // Получаем состояние модели
            var previewImage = await GetThumbnailAsync(partDoc);

            var smCompDef = partDoc.ComponentDefinition as SheetMetalComponentDefinition;
            bool hasFlatPattern = smCompDef != null && smCompDef.HasFlatPattern;

            var partData = new PartData
            {
                Item = itemNumber,
                PartNumber = partNumber,
                ModelState = modelState,
                Description = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description"),
                Material = GetMaterialForPart(partDoc),
                Thickness = GetThicknessForPart(partDoc).ToString("F1") + " мм",
                OriginalQuantity = quantity, // Устанавливаем исходное количество
                Quantity = quantity,
                QuantityColor = System.Windows.Media.Brushes.Black, // Устанавливаем исходный цвет
                Preview = previewImage,
                FlatPatternColor = hasFlatPattern ? System.Windows.Media.Brushes.Green : System.Windows.Media.Brushes.Red
            };

            return partData;
        }

        private BitmapImage GenerateDxfThumbnail(string dxfDirectory, string partNumber)
        {
            string searchPattern = partNumber + "*.dxf"; // Шаблон поиска
            string[] dxfFiles = Directory.GetFiles(dxfDirectory, searchPattern);

            if (dxfFiles.Length == 0)
            {
                return null;
            }

            try
            {
                string dxfFilePath = dxfFiles[0]; // Берем первый найденный файл, соответствующий шаблону
                var generator = new DxfThumbnailGenerator.DxfThumbnailGenerator();
                var bitmap = generator.GenerateThumbnail(dxfFilePath);

                BitmapImage bitmapImage = null;

                // Инициализация изображения должна выполняться в UI потоке
                Dispatcher.Invoke(() =>
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
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
                    System.Windows.MessageBox.Show($"Ошибка при генерации миниатюры DXF: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                });
                return null;
            }
        }

        private string GetModelStateName(PartDocument partDoc)
        {
            try
            {
                var modelStates = partDoc.ComponentDefinition.ModelStates;
                if (modelStates.Count > 0)
                {
                    string activeStateName = modelStates.ActiveModelState.Name;
                    return activeStateName;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при получении Model State: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return "Ошибка получения имени";
        }

        private string GetModelStateName(AssemblyDocument asmDoc)
        {
            try
            {
                var modelStates = asmDoc.ComponentDefinition.ModelStates;
                if (modelStates.Count > 0)
                {
                    string activeStateName = modelStates.ActiveModelState.Name;
                    return activeStateName;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при получении Model State: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return "Ошибка получения имени";
        }

        private async Task<BitmapImage> GetThumbnailAsync(PartDocument document)
        {
            try
            {
                BitmapImage bitmap = null;
                await Dispatcher.InvokeAsync(() =>
                {
                    var apprentice = new Inventor.ApprenticeServerComponent();
                    var apprenticeDoc = apprentice.Open(document.FullDocumentName);

                    var thumbnail = apprenticeDoc.Thumbnail;
                    var img = IPictureDispConverter.PictureDispToImage(thumbnail);

                    if (img != null)
                    {
                        using (var memoryStream = new System.IO.MemoryStream())
                        {
                            img.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                            memoryStream.Position = 0;

                            bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.StreamSource = memoryStream;
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                        }
                    }
                });

                return bitmap;
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    System.Windows.MessageBox.Show("Ошибка при получении миниатюры: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                });
                return null;
            }
        }

        private void ProcessComponentOccurrences(Inventor.ComponentOccurrences occurrences, Dictionary<string, int> sheetMetalParts)
        {
            foreach (Inventor.ComponentOccurrence occ in occurrences)
            {
                if (isCancelled) break;

                if (occ.Suppressed) continue;

                try
                {
                    if (occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                    {
                        PartDocument? partDoc = occ.Definition.Document as PartDocument;
                        if (partDoc != null && partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" && occ.BOMStructure != BOMStructureEnum.kReferenceBOMStructure)
                        {
                            string partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                            if (!string.IsNullOrEmpty(partNumber))
                            {
                                if (sheetMetalParts.TryGetValue(partNumber, out int quantity))
                                {
                                    sheetMetalParts[partNumber]++;
                                }
                                else
                                {
                                    sheetMetalParts.Add(partNumber, 1);
                                }
                            }
                        }
                    }
                    else if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        ProcessComponentOccurrences((Inventor.ComponentOccurrences)occ.SubOccurrences, sheetMetalParts);
                    }
                }
                catch (COMException ex)
                {
                    // Обработка ошибок COM
                }
                catch (Exception ex)
                {
                    // Обработка общих ошибок
                }
            }
        }

        private void ProcessBOM(BOM bom, Dictionary<string, int> sheetMetalParts)
        {
            BOMView? bomView = null;
            foreach (BOMView view in bom.BOMViews)
            {
                bomView = view;
                break;
            }

            if (bomView == null)
            {
                System.Windows.MessageBox.Show("Не найдено ни одного представления спецификации.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            foreach (BOMRow row in bomView.BOMRows)
            {
                if (isCancelled) break;

                try
                {
                    var componentDefinition = row.ComponentDefinitions[1];
                    if (componentDefinition == null)
                    {
                        continue;
                    }

                    Document? document = componentDefinition.Document as Document;
                    if (document == null)
                    {
                        continue;
                    }

                    if (row.BOMStructure == BOMStructureEnum.kReferenceBOMStructure || row.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure)
                    {
                        continue;
                    }

                    if (document.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                    {
                        PartDocument? partDoc = document as PartDocument;
                        if (partDoc != null && partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" && row.BOMStructure != BOMStructureEnum.kReferenceBOMStructure)
                        {
                            string partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                            if (!string.IsNullOrEmpty(partNumber))
                            {
                                if (sheetMetalParts.TryGetValue(partNumber, out int quantity))
                                {
                                    sheetMetalParts[partNumber] += (int)row.ItemQuantity;
                                }
                                else
                                {
                                    sheetMetalParts.Add(partNumber, (int)row.ItemQuantity);
                                }
                            }
                        }
                    }
                    else if (document.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        AssemblyDocument asmDoc = document as AssemblyDocument;
                        if (asmDoc != null)
                        {
                            ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts);
                        }
                    }
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80070057))
                {
                    // Обработка ошибок COM
                }
                catch (Exception ex)
                {
                    // Обработка общих ошибок
                }
            }
        }

        private bool PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch)
        {
            targetDir = @"C:\DXF";
            stopwatch = new Stopwatch();  // Инициализируем Stopwatch перед возможным возвратом

            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            multiplier = 1;
            if (includeQuantityInFileNameCheckBox.IsChecked == true && !int.TryParse(multiplierTextBox.Text, out multiplier))
            {
                System.Windows.MessageBox.Show("Введите допустимое целое число для множителя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            cancelButton.IsEnabled = true;
            isCancelled = false;

            progressBar.IsIndeterminate = false;
            progressBar.Value = 0;
            progressLabel.Text = "Статус: Экспорт данных...";

            stopwatch = Stopwatch.StartNew();

            return true;
        }

        private void FinalizeExport(bool isCancelled, Stopwatch stopwatch, int processedCount, int skippedCount)
        {
            stopwatch.Stop();
            string elapsedTime = GetElapsedTime(stopwatch.Elapsed);

            cancelButton.IsEnabled = false;
            scanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта
            clearButton.IsEnabled = partsData.Count > 0; // Активируем кнопку "Очистить" после завершения экспорта

            if (isCancelled)
            {
                progressLabel.Text = $"Статус: Прервано ({elapsedTime})";
                System.Windows.MessageBox.Show("Процесс экспорта был прерван.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                progressLabel.Text = $"Статус: Завершено ({elapsedTime})";
                MessageBoxHelper.ShowExportCompletedInfo(processedCount, skippedCount, elapsedTime);
            }
        }

        private async Task ExportWithoutScan()
        {
            // Очистка таблицы перед началом скрытого экспорта
            ClearList_Click(this, null);

            if (ThisApplication == null)
            {
                MessageBoxHelper.ShowInventorNotRunningError();
                return;
            }

            Document doc = ThisApplication.ActiveDocument as Document;
            if (doc == null)
            {
                MessageBoxHelper.ShowNoDocumentOpenWarning();
                return;
            }

            if (!PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch))
            {
                return;
            }

            clearButton.IsEnabled = false;
            scanButton.IsEnabled = false;

            int processedCount = 0;
            int skippedCount = 0;

            // Скрытое сканирование
            var sheetMetalParts = new Dictionary<string, int>();

            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = (AssemblyDocument)doc;
                if (traverseRadioButton.IsChecked == true)
                {
                    await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
                }
                else if (bomRadioButton.IsChecked == true)
                {
                    await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));
                }

                // Создаем временный список PartData для экспорта
                var tempPartsDataList = new List<PartData>();
                foreach (var part in sheetMetalParts)
                {
                    tempPartsDataList.Add(new PartData
                    {
                        PartNumber = part.Key,
                        OriginalQuantity = part.Value,
                        Quantity = part.Value * multiplier
                    });
                }

                await Task.Run(() => ExportDXF(tempPartsDataList, targetDir, multiplier, ref processedCount, ref skippedCount, false)); // false - без генерации миниатюр
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument partDoc = (PartDocument)doc;
                if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    string partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
                    var tempPartData = new List<PartData>
            {
                new PartData
                {
                    PartNumber = partNumber,
                    OriginalQuantity = 1,
                    Quantity = multiplier
                }
            };
                    await Task.Run(() => ExportDXF(tempPartData, targetDir, multiplier, ref processedCount, ref skippedCount, false)); // false - без генерации миниатюр
                }
                else
                {
                    MessageBoxHelper.ShowNotSheetMetalWarning();
                }
            }

            clearButton.IsEnabled = true;
            scanButton.IsEnabled = true;

            // Здесь мы НЕ обновляем fileCountLabelBottom
            FinalizeExport(isCancelled, stopwatch, processedCount, skippedCount);

            // Деактивируем кнопку "Экспорт" после завершения скрытого экспорта
            exportButton.IsEnabled = false;
        }
        private async void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (isCtrlPressed)
            {
                await ExportWithoutScan();
                return;
            }
            if (ThisApplication == null)
            {
                MessageBoxHelper.ShowInventorNotRunningError();
                return;
            }

            Document? doc = ThisApplication.ActiveDocument as Document;
            if (doc == null)
            {
                MessageBoxHelper.ShowNoDocumentOpenWarning();
                return;
            }

            if (doc != lastScannedDocument)
            {
                System.Windows.MessageBox.Show("Активный документ изменился. Пожалуйста, повторите сканирование перед экспортом.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                multiplierTextBox.Text = "1";
                return;
            }

            if (!PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch))
            {
                return;
            }

            clearButton.IsEnabled = false; // Деактивируем кнопку "Очистить" в процессе экспорта
            scanButton.IsEnabled = false; // Деактивируем кнопку "Сканировать" в процессе экспорта

            int processedCount = 0;
            int skippedCount = 0;

            if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = (AssemblyDocument)doc;
                var sheetMetalParts = new Dictionary<string, int>();
                if (traverseRadioButton.IsChecked == true)
                {
                    await Task.Run(() => ProcessComponentOccurrences(asmDoc.ComponentDefinition.Occurrences, sheetMetalParts));
                }
                else if (bomRadioButton.IsChecked == true)
                {
                    await Task.Run(() => ProcessBOM(asmDoc.ComponentDefinition.BOM, sheetMetalParts));
                }

                await ExportDXFFiles(sheetMetalParts, targetDir, multiplier, stopwatch);
            }
            else if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument partDoc = (PartDocument)doc;
                if (partDoc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                {
                    await ExportSinglePartDXF(partDoc, targetDir, multiplier, stopwatch);
                }
                else
                {
                    MessageBoxHelper.ShowNotSheetMetalWarning();
                }
            }

            clearButton.IsEnabled = true; // Активируем кнопку "Очистить" после завершения экспорта
            scanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта

            // Кнопка "Экспорт" остается активной после обычного экспорта
            exportButton.IsEnabled = true;
        }
        private string GetElapsedTime(TimeSpan timeSpan)
        {
            if (timeSpan.TotalMinutes >= 1)
            {
                return $"{(int)timeSpan.TotalMinutes} мин. {timeSpan.Seconds} сек.";
            }
            else
            {
                return $"{timeSpan.Seconds} сек.";
            }
        }
        private async void ExportSelectedDXF_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = partsDataGrid.SelectedItems.Cast<PartData>().ToList();
            if (selectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show("Выберите строки для экспорта.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var itemsWithoutFlatPattern = selectedItems.Where(p => p.FlatPatternColor == System.Windows.Media.Brushes.Red).ToList();
            if (itemsWithoutFlatPattern.Count == selectedItems.Count)
            {
                System.Windows.MessageBox.Show("Выбранные файлы не содержат разверток.", "Информация", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (itemsWithoutFlatPattern.Count > 0)
            {
                var result = System.Windows.MessageBox.Show("Некоторые выбранные файлы не содержат разверток и будут пропущены. Продолжить?", "Информация", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                {
                    return;
                }

                selectedItems = selectedItems.Except(itemsWithoutFlatPattern).ToList();
            }

            if (!PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch))
            {
                return;
            }

            clearButton.IsEnabled = false; // Деактивируем кнопку "Очистить" в процессе экспорта
            scanButton.IsEnabled = false; // Деактивируем кнопку "Сканировать" в процессе экспорта

            int processedCount = 0;
            int skippedCount = itemsWithoutFlatPattern.Count;

            await Task.Run(() =>
            {
                ExportDXF(selectedItems, targetDir, multiplier, ref processedCount, ref skippedCount, true); // true - с генерацией миниатюр
            });

            UpdateFileCountLabel(selectedItems.Count);

            clearButton.IsEnabled = true; // Активируем кнопку "Очистить" после завершения экспорта
            scanButton.IsEnabled = true; // Активируем кнопку "Сканировать" после завершения экспорта
            FinalizeExport(isCancelled, stopwatch, processedCount, skippedCount);
        }
        private void OverrideQuantity_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = partsDataGrid.SelectedItems.Cast<PartData>().ToList();
            if (selectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show("Выберите строки для переопределения количества.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
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
                    item.QuantityColor = System.Windows.Media.Brushes.Red; // Переопределенное количество окрашивается в красный цвет
                }
                partsDataGrid.Items.Refresh();
            }
        }

        private void ExportDXF(IEnumerable<PartData> partsDataList, string targetDir, int multiplier, ref int processedCount, ref int skippedCount, bool generateThumbnails)
        {
            bool organizeByMaterial = false;
            bool organizeByThickness = false;
            bool includeQuantityInFileName = false;

            Dispatcher.Invoke(() =>
            {
                organizeByMaterial = organizeByMaterialCheckBox.IsChecked == true;
                organizeByThickness = organizeByThicknessCheckBox.IsChecked == true;
                includeQuantityInFileName = includeQuantityInFileNameCheckBox.IsChecked == true;
            });

            int totalParts = partsDataList.Count();

            Dispatcher.Invoke(() =>
            {
                progressBar.Maximum = totalParts;
                progressBar.Value = 0;
            });

            int localProcessedCount = processedCount;
            int localSkippedCount = skippedCount;

            foreach (var partData in partsDataList)
            {
                if (isCancelled) break;

                string partNumber = partData.PartNumber;
                int qty = partData.IsOverridden ? partData.Quantity : partData.OriginalQuantity * multiplier;

                PartDocument partDoc = null;
                bool isPartDocOpenedHere = false;
                try
                {
                    partDoc = OpenPartDocument(partNumber);
                    if (partDoc == null)
                    {
                        throw new Exception("Файл детали не найден или не может быть открыт");
                    }

                    var smCompDef = (SheetMetalComponentDefinition)partDoc.ComponentDefinition;
                    string material = GetMaterialForPart(partDoc);
                    double thickness = GetThicknessForPart(partDoc);

                    string materialDir = organizeByMaterial ? System.IO.Path.Combine(targetDir, material) : targetDir;
                    if (!Directory.Exists(materialDir))
                    {
                        Directory.CreateDirectory(materialDir);
                    }

                    string thicknessDir = organizeByThickness ? System.IO.Path.Combine(materialDir, thickness.ToString("F1")) : materialDir;
                    if (!Directory.Exists(thicknessDir))
                    {
                        Directory.CreateDirectory(thicknessDir);
                    }

                    string filePath = System.IO.Path.Combine(thicknessDir, partNumber + (includeQuantityInFileName ? " - " + qty.ToString() : "") + ".dxf");

                    if (!IsValidPath(filePath))
                    {
                        continue;
                    }

                    bool exportSuccess = false;

                    if (smCompDef.HasFlatPattern)
                    {
                        var flatPattern = smCompDef.FlatPattern;
                        var oDataIO = flatPattern.DataIO;

                        try
                        {
                            string options = null;
                            Dispatcher.Invoke(() => PrepareExportOptions(out options));

                            // Интеграция настроек слоев
                            StringBuilder layerOptionsBuilder = new StringBuilder();
                            StringBuilder invisibleLayersBuilder = new StringBuilder();

                            foreach (var layer in this.LayerSettings) // Используем настройки слоев
                            {
                                if (layer.HasVisibilityOption && !layer.IsVisible)
                                {
                                    invisibleLayersBuilder.Append($"{layer.LayerName};");
                                    continue;
                                }

                                if (!string.IsNullOrWhiteSpace(layer.CustomName))
                                {
                                    layerOptionsBuilder.Append($"&{layer.DisplayName}={layer.CustomName}");
                                }

                                if (layer.SelectedLineType != "Default")
                                {
                                    layerOptionsBuilder.Append($"&{layer.DisplayName}LineType={LayerSettingsApp.MainWindow.GetLineTypeValue(layer.SelectedLineType)}");
                                }

                                if (layer.SelectedColor != "White")
                                {
                                    layerOptionsBuilder.Append($"&{layer.DisplayName}Color={LayerSettingsApp.MainWindow.GetColorValue(layer.SelectedColor)}");
                                }
                            }

                            if (invisibleLayersBuilder.Length > 0)
                            {
                                invisibleLayersBuilder.Length -= 1; // Убираем последний символ ";"
                                layerOptionsBuilder.Append($"&InvisibleLayers={invisibleLayersBuilder}");
                            }

                            // Объединяем общие параметры и параметры слоев
                            options += layerOptionsBuilder.ToString();

                            string dxfOptions = $"FLAT PATTERN DXF?{options}";

                            // Debugging output
                            System.Diagnostics.Debug.WriteLine($"DXF Options: {dxfOptions}");

                            // Создаем файл, если он не существует
                            if (!System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Create(filePath).Close();
                            }

                            // Проверка на занятость файла
                            if (IsFileLocked(filePath))
                            {
                                var result = System.Windows.MessageBox.Show(
                                    $"Файл {filePath} занят другим приложением. Прервать операцию?",
                                    "Предупреждение",
                                    MessageBoxButton.YesNo,
                                    MessageBoxImage.Warning,
                                    MessageBoxResult.No // Устанавливаем "No" как значение по умолчанию
                                );

                                if (result == MessageBoxResult.Yes)
                                {
                                    isCancelled = true; // Прерываем операцию
                                    break;
                                }
                                else
                                {
                                    while (IsFileLocked(filePath))
                                    {
                                        var waitResult = System.Windows.MessageBox.Show(
                                            "Ожидание разблокировки файла. Возможно файл открыт в программе просмотра. Закройте файл и нажмите OK.",
                                            "Информация",
                                            MessageBoxButton.OKCancel,
                                            MessageBoxImage.Information,
                                            MessageBoxResult.Cancel // Устанавливаем "Cancel" как значение по умолчанию
                                        );

                                        if (waitResult == MessageBoxResult.Cancel)
                                        {
                                            isCancelled = true; // Прерываем операцию
                                            break;
                                        }

                                        Thread.Sleep(1000); // Ожидание 1 секунды перед повторной проверкой
                                    }

                                    if (isCancelled)
                                    {
                                        break;
                                    }
                                }
                            }

                            oDataIO.WriteDataToFile(dxfOptions, filePath);
                            exportSuccess = true;
                        }
                        catch (Exception ex)
                        {
                            // Обработка ошибок экспорта DXF
                            System.Diagnostics.Debug.WriteLine($"Ошибка экспорта DXF: {ex.Message}");
                        }
                    }

                    BitmapImage dxfPreview = null;
                    if (generateThumbnails)
                    {
                        dxfPreview = GenerateDxfThumbnail(thicknessDir, partNumber);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        partData.ProcessingColor = exportSuccess ? System.Windows.Media.Brushes.Green : System.Windows.Media.Brushes.Red;
                        partData.DxfPreview = dxfPreview;
                        partsDataGrid.Items.Refresh();
                    });

                    if (exportSuccess)
                    {
                        localProcessedCount++;
                    }
                    else
                    {
                        localSkippedCount++;
                    }

                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Value = Math.Min(localProcessedCount, progressBar.Maximum);
                    });
                }
                catch (Exception ex)
                {
                    localSkippedCount++;
                    System.Diagnostics.Debug.WriteLine($"Ошибка обработки детали: {ex.Message}");
                }
                finally
                {
                    // partDoc?.Close(false); // Убираем закрытие документа
                }
            }

            processedCount = localProcessedCount;
            skippedCount = localSkippedCount;

            Dispatcher.Invoke(() =>
            {
                progressBar.Value = progressBar.Maximum; // Установка значения 100% по завершению
            });
        }

        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
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

        private async Task ExportDXFFiles(Dictionary<string, int> sheetMetalParts, string targetDir, int multiplier, Stopwatch stopwatch)
        {
            int processedCount = 0;
            int skippedCount = 0;

            var partsDataList = partsData.Where(p => sheetMetalParts.ContainsKey(p.PartNumber)).ToList();
            await Task.Run(() => ExportDXF(partsDataList, targetDir, multiplier, ref processedCount, ref skippedCount, true)); // true - с генерацией миниатюр

            UpdateFileCountLabel(partsDataList.Count);
            FinalizeExport(isCancelled, stopwatch, processedCount, skippedCount);
        }

        private async Task ExportSinglePartDXF(PartDocument partDoc, string targetDir, int multiplier, Stopwatch stopwatch)
        {
            int processedCount = 0;
            int skippedCount = 0;

            string partNumber = GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number");
            var partData = partsData.FirstOrDefault(p => p.PartNumber == partNumber);

            if (partData != null)
            {
                var partsDataList = new List<PartData> { partData };
                await Task.Run(() => ExportDXF(partsDataList, targetDir, multiplier, ref processedCount, ref skippedCount, true)); // true - с генерацией миниатюр
                UpdateFileCountLabel(1);
            }

            FinalizeExport(isCancelled, stopwatch, processedCount, skippedCount);
        }

        private PartDocument? OpenPartDocument(string partNumber)
        {
            var docs = ThisApplication.Documents;
            var projectManager = ThisApplication.DesignProjectManager;
            var activeProject = projectManager.ActiveDesignProject;
            PartDocument? partDoc = null;

            foreach (Document doc in docs)
            {
                if (doc is PartDocument pd && GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                {
                    return pd;
                }
            }

            string fullFileName = System.IO.Path.Combine(activeProject.WorkspacePath, partNumber + ".ipt");
            if (System.IO.File.Exists(fullFileName))
            {
                try
                {
                    partDoc = (PartDocument)docs.Open(fullFileName);
                    // Устанавливаем флаг, что документ был открыт здесь
                }
                catch (Exception)
                {
                    // Обработка ошибок открытия документа детали
                }
            }

            return partDoc;
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
                Parameter thicknessParam = smCompDef.Thickness;
                return Math.Round((double)thicknessParam.Value * 10, 1); // Извлекаем значение толщины и переводим в мм, округляем до одной десятой
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
                new System.IO.FileInfo(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void IncludeQuantityInFileNameCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            multiplierTextBox.IsEnabled = includeQuantityInFileNameCheckBox.IsChecked == true;
            if (!multiplierTextBox.IsEnabled)
            {
                multiplierTextBox.Text = "1";
                UpdateQuantitiesWithMultiplier(1);
            }
        }

        private void MultiplierTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (int.TryParse(multiplierTextBox.Text, out int multiplier) && multiplier > 0)
            {
                UpdateQuantitiesWithMultiplier(multiplier);
            }
        }

        private void UpdateFileCountLabel(int count)
        {
            fileCountLabelBottom.Text = $"Найдено листовых деталей: {count}";
        }


        private void ResetProgressBar()
        {
            progressBar.IsIndeterminate = false;
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            progressLabel.Text = "Статус: ";
            UpdateFileCountLabel(0); // Использование метода для сброса счетчика
            bottomPanel.Opacity = 0.75;
            documentInfoLabel.Text = string.Empty;
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            isCancelled = true;
            cancelButton.IsEnabled = false;

            if (!isScanning)
            {
                // If not scanning, reset the UI
                ResetProgressBar();
                bottomPanel.Opacity = 0.75;
                documentInfoLabel.Text = "";
                exportButton.IsEnabled = false;
                clearButton.IsEnabled = partsData.Count > 0; // Обновляем состояние кнопки "Очистить"
                lastScannedDocument = null;
                UpdateFileCountLabel(0);
            }
        }
        private void UpdateDocumentInfo(string documentType, string partNumber, string description, Document doc)
        {
            if (string.IsNullOrEmpty(documentType) && string.IsNullOrEmpty(partNumber) && string.IsNullOrEmpty(description))
            {
                documentInfoLabel.Text = string.Empty;
                modelStateInfoRunBottom.Text = string.Empty; // Обновленное имя элемента
                bottomPanel.Opacity = 0.75;
                UpdateFileCountLabel(0); // Использование метода для сброса счетчика
                return;
            }

            string modelStateInfo = (doc is PartDocument partDoc) ? GetModelStateName(partDoc) :
                                    (doc is AssemblyDocument asmDoc) ? GetModelStateName(asmDoc) :
                                    "Ошибка получения имени";

            documentInfoLabel.Text = $"Тип документа: {documentType}\nОбозначение: {partNumber}\nОписание: {description}";

            if (modelStateInfo == "[Primary]" || modelStateInfo == "[Основной]")
            {
                modelStateInfoRunBottom.Text = "[Основной]";
                modelStateInfoRunBottom.Foreground = new SolidColorBrush(Colors.Black); // Цвет текста - черный (по умолчанию)
            }
            else
            {
                modelStateInfoRunBottom.Text = modelStateInfo;
                modelStateInfoRunBottom.Foreground = new SolidColorBrush(Colors.Red); // Цвет текста - красный
            }

            UpdateFileCountLabel(partsData.Count); // Использование метода для обновления счетчика
            bottomPanel.Opacity = 1.0;
        }
        private void ClearList_Click(object sender, RoutedEventArgs e)
        {
            partsData.Clear();
            progressBar.Value = 0;
            progressLabel.Text = "Статус: ";
            documentInfoLabel.Text = string.Empty;
            modelStateInfoRunBottom.Text = string.Empty;
            bottomPanel.Opacity = 0.5;
            exportButton.IsEnabled = false;
            clearButton.IsEnabled = false; // Делаем кнопку "Очистить" неактивной после очистки

            // Обнуляем информацию о документе
            UpdateDocumentInfo(string.Empty, string.Empty, string.Empty, null);
            UpdateFileCountLabel(0); // Сброс счетчика файлов
        }
        private void partsDataGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (partsData.Count == 0)
            {
                e.Handled = true; // Предотвращаем открытие контекстного меню, если таблица пуста
            }
        }

        private void OpenFileLocation_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = partsDataGrid.SelectedItems.Cast<PartData>().ToList();
            if (selectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show("Выберите строки для открытия расположения файла.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            foreach (var item in selectedItems)
            {
                string partNumber = item.PartNumber;
                string fullPath = GetPartDocumentFullPath(partNumber);
                if (System.IO.File.Exists(fullPath))
                {
                    string argument = "/select, \"" + fullPath + "\"";
                    Process.Start("explorer.exe", argument);
                }
                else
                {
                    System.Windows.MessageBox.Show($"Файл {partNumber} не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OpenSelectedModels_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = partsDataGrid.SelectedItems.Cast<PartData>().ToList();
            if (selectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show("Выберите строки для открытия моделей.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            foreach (var item in selectedItems)
            {
                string partNumber = item.PartNumber;
                string fullPath = GetPartDocumentFullPath(partNumber);
                if (System.IO.File.Exists(fullPath))
                {
                    try
                    {
                        ThisApplication.Documents.Open(fullPath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"Ошибка при открытии файла {partNumber}: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show($"Файл {partNumber} не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string GetPartDocumentFullPath(string partNumber)
        {
            var docs = ThisApplication.Documents;
            foreach (Document doc in docs)
            {
                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    var partDoc = doc as PartDocument;
                    if (partDoc != null && GetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                    {
                        return partDoc.FullFileName;
                    }
                }
            }

            // Если документ не найден в открытых, ищем в проекте
            var projectManager = ThisApplication.DesignProjectManager;
            var activeProject = projectManager.ActiveDesignProject;
            return System.IO.Path.Combine(activeProject.WorkspacePath, partNumber + ".ipt");
        }

        private void partsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }

    public class PartData : INotifyPropertyChanged
    {
        private int quantity;
        private System.Windows.Media.Brush quantityColor;
        private int item; // Новое поле для пункта

        public int Item // Измените тип на int
        {
            get => item;
            set
            {
                item = value;
                OnPropertyChanged();
            }
        }
        public string PartNumber { get; set; }
        public string ModelState { get; set; }
        public string Description { get; set; }
        public string Material { get; set; }
        public string Thickness { get; set; }
        public int OriginalQuantity { get; set; }
        public int Quantity
        {
            get => quantity;
            set
            {
                quantity = value;
                OnPropertyChanged();
            }
        }
        public BitmapImage Preview { get; set; }
        public System.Windows.Media.Brush FlatPatternColor { get; set; }
        public System.Windows.Media.Brush ProcessingColor { get; set; } = System.Windows.Media.Brushes.Transparent;
        public System.Windows.Media.Brush QuantityColor
        {
            get => quantityColor;
            set
            {
                quantityColor = value;
                OnPropertyChanged();
            }
        }
        public bool IsOverridden { get; set; } = false;
        public BitmapImage DxfPreview { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
    public static class MessageBoxHelper
    {
        public static void ShowInventorNotRunningError()
        {
            System.Windows.MessageBox.Show("Inventor не запущен. Пожалуйста, запустите Inventor и попробуйте снова.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static void ShowNoDocumentOpenWarning()
        {
            System.Windows.MessageBox.Show("Нет открытого документа. Пожалуйста, откройте сборку или деталь и попробуйте снова.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        public static void ShowScanningCancelledInfo()
        {
            System.Windows.MessageBox.Show("Процесс сканирования был прерван.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static void ShowExportCancelledInfo()
        {
            System.Windows.MessageBox.Show("Процесс экспорта был прерван.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static void ShowNoFlatPatternWarning()
        {
            System.Windows.MessageBox.Show("Выбранные файлы не содержат разверток.", "Информация", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        public static void ShowPartNotFoundError(string partNumber)
        {
            System.Windows.MessageBox.Show($"Файл {partNumber} не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static void ShowExportCompletedInfo(int processedCount, int skippedCount, string elapsedTime)
        {
            System.Windows.MessageBox.Show($"Экспорт DXF завершен.\nВсего файлов обработано: {processedCount + skippedCount}\nПропущено (без разверток): {skippedCount}\nВсего экспортировано: {processedCount}\nВремя выполнения: {elapsedTime}", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static void ShowErrorMessage(string message)
        {
            System.Windows.MessageBox.Show(message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static void ShowNotSheetMetalWarning()
        {
            System.Windows.MessageBox.Show("Активный документ не является листовым металлом.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

}
