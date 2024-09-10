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
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using System.Reflection;
using System.Windows.Media.Media3D;

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
        private AdornerLayer _adornerLayer;
        private HeaderAdorner _headerAdorner;
        private DataGridColumn _reorderingColumn;
        private bool _isColumnDraggedOutside = false;
        private CollectionViewSource partsDataView;
        private DispatcherTimer searchDelayTimer;
        private const string PlaceholderText = "Поиск...";
        private string actualSearchText = string.Empty; // Поле для хранения фактического текста поиска
        private string fixedFolderPath = string.Empty;
        private bool isEditing = false;
        private List<PartData> originalPartsData = new List<PartData>(); // Для хранения исходного состояния

        public ObservableCollection<LayerSetting> LayerSettings { get; set; }
        public ObservableCollection<string> AvailableColors { get; set; }
        public ObservableCollection<string> LineTypes { get; set; }
        public ObservableCollection<PresetIProperty> PresetIProperties { get; set; }

        private List<string> customPropertiesList = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            InitializeInventor();
            partsDataGrid.ItemsSource = partsData;
            multiplierTextBox.IsEnabled = includeQuantityInFileNameCheckBox.IsChecked == true;
            UpdateFileCountLabel(0);
            clearButton.IsEnabled = false;
            partsDataGrid.PreviewMouseDown += PartsDataGrid_PreviewMouseDown;
            partsDataGrid.PreviewMouseMove += PartsDataGrid_PreviewMouseMove;
            partsDataGrid.PreviewMouseLeftButtonUp += PartsDataGrid_PreviewMouseLeftButtonUp;
            partsDataGrid.ColumnReordering += PartsDataGrid_ColumnReordering;
            partsDataGrid.ColumnReordered += PartsDataGrid_ColumnReordered;
            chooseFolderRadioButton.IsChecked = true;

            partsDataGrid.DragOver += PartsDataGrid_DragOver;
            partsDataGrid.DragLeave += PartsDataGrid_DragLeave;

            // Инициализация CollectionViewSource для фильтрации
            partsDataView = new CollectionViewSource { Source = partsData };
            partsDataView.Filter += PartsData_Filter;

            // Устанавливаем ItemsSource для DataGrid
            partsDataGrid.ItemsSource = partsDataView.View;

            // Подписываемся на событие изменения коллекции
            partsData.CollectionChanged += PartsData_CollectionChanged;

            UpdateEditButtonState(); // Обновляем состояние кнопки после загрузки данных

            // Настраиваем таймер для задержки поиска
            searchDelayTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1) // 1 секунда задержки
            };
            searchDelayTimer.Tick += SearchDelayTimer_Tick;

            // Установка начального текста для searchTextBox
            searchTextBox.Text = PlaceholderText;
            searchTextBox.Foreground = System.Windows.Media.Brushes.Gray;

            // Создаем экземпляр MainWindow из LayerSettingsApp
            var layerSettingsWindow = new LayerSettingsApp.MainWindow();

            // Используем его данные для настройки LayerSettings
            LayerSettings = layerSettingsWindow.LayerSettings;
            AvailableColors = layerSettingsWindow.AvailableColors;
            LineTypes = layerSettingsWindow.LineTypes;


            // Инициализация предустановленных колонок
            PresetIProperties = new ObservableCollection<PresetIProperty>
            {
                new PresetIProperty { InternalName = "ID", DisplayName = "#Нумерация", InventorPropertyName = "Item"},
                new PresetIProperty { InternalName = "Обозначение", DisplayName = "Обозначение", InventorPropertyName = "PartNumber"},
                new PresetIProperty { InternalName = "Наименование", DisplayName = "Наименование", InventorPropertyName = "Description"},
                new PresetIProperty { InternalName = "Состояние модели", DisplayName = "Состояние модели", InventorPropertyName = "ModelState"},
                new PresetIProperty { InternalName = "Материал", DisplayName = "Материал", InventorPropertyName = "Material"},
                new PresetIProperty { InternalName = "Толщина", DisplayName = "Толщина", InventorPropertyName = "Thickness"},
                new PresetIProperty { InternalName = "Количество", DisplayName = "Количество", InventorPropertyName = "Quantity"},
                new PresetIProperty { InternalName = "Изображение детали", DisplayName = "Изображение детали", InventorPropertyName = "Preview"},
                new PresetIProperty { InternalName = "Изображение развертки", DisplayName = "Изображение развертки", InventorPropertyName = "DxfPreview"},

                // Добавление новых предустановленных свойств
                new PresetIProperty { InternalName = "Автор", DisplayName = "Автор", InventorPropertyName = "Author"}, //, ShouldBeAddedOnInit = true or false
                new PresetIProperty { InternalName = "Ревизия", DisplayName = "Ревизия", InventorPropertyName = "Revision Number" },
                new PresetIProperty { InternalName = "Проект", DisplayName = "Проект", InventorPropertyName = "Project"},
                new PresetIProperty { InternalName = "Инвентарный номер", DisplayName = "Инвентарный номер", InventorPropertyName = "Stock Number"}
            };

            // Устанавливаем DataContext для текущего окна, объединяя данные из LayerSettingsWindow и других источников
            DataContext = this;

            // Инициализация DataGrid с предустановленными колонками
            InitializePresetColumns();

            // Добавляем обработчики для отслеживания нажатия и отпускания клавиши Ctrl
            this.KeyDown += new System.Windows.Input.KeyEventHandler(MainWindow_KeyDown);
            this.KeyUp += new System.Windows.Input.KeyEventHandler(MainWindow_KeyUp);

            VersionTextBlock.Text = "Версия программы: " + GetVersion();
        }
        public string GetVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();

            // Попытка получить информацию о версии из различных атрибутов
            var informationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
            var assemblyVersion = assembly.GetName().Version?.ToString();

            // Если версия содержит commit hash, обрезаем его до 7 символов
            if (informationalVersion != null && informationalVersion.Contains("+"))
            {
                // Разбиваем строку на части и обрезаем commit hash до 7 символов
                var parts = informationalVersion.Split('+');
                var shortCommitHash = parts[1].Length > 7 ? parts[1].Substring(0, 7) : parts[1];
                informationalVersion = $"{parts[0]}+{shortCommitHash}";
            }

            // Возвращаем обрезанную версию или другие версии
            return informationalVersion ?? fileVersion ?? assemblyVersion ?? "Версия неизвестна";
        }
        public bool IsColumnPresent(string columnName)
        {
            bool result = false;

            // Используем Dispatcher для доступа к UI-элементам
            Dispatcher.Invoke(() =>
            {
                result = partsDataGrid.Columns.Any(c => c.Header.ToString() == columnName);
            });

            return result;
        }

        private void InitializePresetColumns()
        {
            if (PresetIProperties == null)
            {
                throw new InvalidOperationException("PresetIProperties is not initialized.");
            }

            foreach (var property in PresetIProperties)
            {
                // Если колонка уже есть в таблице, ничего не делаем
                if (IsColumnPresent(property.InternalName))
                {
                    continue;
                }

                // Если колонка не добавлена в XAML, пропускаем её добавление
                if (property.ShouldBeAddedOnInit)
                {
                    AddIPropertyColumn(property);
                }
            }
        }
        private async Task<string> GetPropertyExpressionOrValueAsync(PartDocument partDoc, string propertyName, bool getExpression = false)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var propertySets = partDoc.PropertySets;
                    foreach (PropertySet propSet in propertySets)
                    {
                        foreach (Inventor.Property prop in propSet)
                        {
                            if (prop.Name == propertyName)
                            {
                                if (getExpression)
                                {
                                    // Попытка получить выражение
                                    try
                                    {
                                        string expression = prop.Expression;
                                        if (!string.IsNullOrEmpty(expression))
                                        {
                                            return expression; // Возвращаем выражение
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        // Если нет выражения, вернем значение
                                        return prop.Value?.ToString() ?? string.Empty;
                                    }
                                }
                                return prop.Value?.ToString() ?? string.Empty;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка при получении выражения или значения для {propertyName}: {ex.Message}");
                }

                return string.Empty;
            });
        }

        //private void UpdateCustomProperties(PartDocument partDoc, PartData partData)
        //{
        //    foreach (var customPropertyName in partData.CustomProperties.Keys.ToList())
        //    {
        //        var expressionOrValue = GetPropertyExpressionOrValue(partDoc, customPropertyName);
        //        partData.CustomProperties[customPropertyName] = expressionOrValue;
        //    }
        //}

        private async void EditButton_Click(object sender, RoutedEventArgs e)
        {
            if (isEditing)
            {
                // Отключаем режим редактирования
                isEditing = false;
                partsDataGrid.IsReadOnly = true;
                editButton.Content = "Редактировать";
                saveButton.IsEnabled = false;

                // Восстанавливаем исходное состояние данных
                for (int i = 0; i < partsData.Count; i++)
                {
                    var currentItem = partsData[i];
                    var originalItem = originalPartsData[i];

                    currentItem.PartNumber = originalItem.PartNumber; // Возвращаем значение
                    currentItem.Description = originalItem.Description; // Возвращаем значение

                    // Восстанавливаем кастомные свойства
                    foreach (var key in currentItem.CustomProperties.Keys.ToList())
                    {
                        currentItem.CustomProperties[key] = originalItem.CustomProperties.ContainsKey(key)
                            ? originalItem.CustomProperties[key]
                            : currentItem.CustomProperties[key];
                    }
                }

                partsDataGrid.Items.Refresh(); // Обновляем таблицу
            }
            else
            {
                // Включаем режим редактирования
                isEditing = true;
                partsDataGrid.IsReadOnly = false;
                editButton.Content = "Отмена";
                saveButton.IsEnabled = true;

                // Сохраняем исходное состояние данных
                originalPartsData = partsData.Select(item => new PartData
                {
                    PartNumber = item.PartNumber,
                    PartNumberExpression = item.PartNumberExpression,
                    Description = item.Description,
                    DescriptionExpression = item.DescriptionExpression,
                    CustomProperties = new Dictionary<string, string>(item.CustomProperties),
                    CustomPropertyExpressions = new Dictionary<string, string>(item.CustomPropertyExpressions)
                }).ToList();

                // Проверяем, если чекбокс активен, то отображаем выражения
                if (expressionsCheckBox.IsChecked == true)
                {
                    foreach (var item in partsData)
                    {
                        item.PartNumber = item.PartNumberExpression; // Показываем выражение
                        item.Description = item.DescriptionExpression; // Показываем выражение

                        foreach (var customProperty in item.CustomProperties.Keys.ToList())
                        {
                            item.CustomProperties[customProperty] = item.CustomPropertyExpressions[customProperty]; // Показываем выражение
                        }
                    }
                }

                partsDataGrid.Items.Refresh();
            }


            UpdateEditButtonState();
        }
        private void ExpressionsCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (isEditing) // Применяем изменения только в режиме редактирования
            {
                if (expressionsCheckBox.IsChecked == true)
                {
                    // Показываем выражения
                    foreach (var item in partsData)
                    {
                        item.PartNumber = item.PartNumberExpression; // Показываем выражение
                        item.Description = item.DescriptionExpression; // Показываем выражение

                        foreach (var customProperty in item.CustomProperties.Keys.ToList())
                        {
                            item.CustomProperties[customProperty] = item.CustomPropertyExpressions[customProperty]; // Показываем выражение
                        }
                    }
                }
                else
                {
                    // Показываем обычные значения
                    foreach (var item in partsData)
                    {
                        item.PartNumber = item.PartNumber; // Показываем значение
                        item.Description = item.Description; // Показываем значение

                        foreach (var customProperty in item.CustomProperties.Keys.ToList())
                        {
                            item.CustomProperties[customProperty] = item.CustomProperties[customProperty]; // Показываем значение
                        }
                    }
                }

                partsDataGrid.Items.Refresh(); // Обновляем таблицу после изменения
            }
        }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            isEditing = false;
            partsDataGrid.IsReadOnly = true;
            editButton.Content = "Редактировать";
            saveButton.IsEnabled = false;

            foreach (var item in partsData)
            {
                if (item.IsModified)  // Сохраняем только изменённые элементы
                {
                    await SaveIPropertyChangesAsync(item);
                }

                // После успешного сохранения сбрасываем флаг IsModified и делаем строки только для чтения
                item.IsModified = false;
                item.IsReadOnly = true;
            }

            // Очищаем оригинальные данные, так как изменения сохранены
            originalPartsData.Clear();
            partsDataGrid.Items.Refresh();

            // Восстанавливаем состояние кнопки
            UpdateEditButtonState();
        }
        private async Task SaveIPropertyChangesAsync(PartData item)
        {
            try
            {
                var partDoc = OpenPartDocument(item.OriginalPartNumber);
                if (partDoc != null)
                {
                    SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Part Number", item.PartNumber);
                    SetProperty(partDoc.PropertySets["Design Tracking Properties"], "Description", item.Description);

                    await Task.Run(() =>
                    {
                        partDoc.Save2(true);
                        partDoc.Close();
                    });

                    // После успешного сохранения сбрасываем флаг
                    item.IsModified = false;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PartsDataGrid_DragOver(object sender, System.Windows.DragEventArgs e)
        {
            // Уведомляем SelectIPropertyWindow, что курсор находится над DataGrid
            var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
            if (selectIPropertyWindow != null)
            {
                selectIPropertyWindow.partsDataGrid_DragOver(sender, e);
            }
        }

        private void PartsDataGrid_DragLeave(object sender, System.Windows.DragEventArgs e)
        {
            // Уведомляем SelectIPropertyWindow, что курсор ушел из DataGrid
            var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
            if (selectIPropertyWindow != null)
            {
                selectIPropertyWindow.partsDataGrid_DragLeave(sender, e);
            }
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
        private void PartsData_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            UpdateEditButtonState(); // Обновляем состояние кнопки при изменении данных
        }

        private void UpdateEditButtonState()
        {
            // Если данные в таблице присутствуют, и не активирован режим редактирования
            if (partsData != null && partsData.Count > 0)
            {
                editButton.IsEnabled = true;  // Кнопка "Редактировать/Отмена" всегда активна, если есть данные
                saveButton.IsEnabled = isEditing; // Кнопка "Сохранить" активна только в режиме редактирования
            }
            else
            {
                editButton.IsEnabled = false; // Деактивируем кнопку, если данных нет
                saveButton.IsEnabled = false; // Деактивируем кнопку сохранения, если данных нет
            }
        }

        private void AddPresetIPropertyButton_Click(object sender, RoutedEventArgs e)
        {
            var selectIPropertyWindow = new SelectIPropertyWindow(PresetIProperties, this); // Передаем `this` как ссылку на MainWindow
            selectIPropertyWindow.ShowDialog();

            // После закрытия окна добавляем выбранные свойства
            foreach (var property in selectIPropertyWindow.SelectedProperties)
            {
                AddIPropertyColumn(property);
            }
        }

        private void MainWindow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Получаем исходный элемент
            var clickedElement = e.OriginalSource as DependencyObject;

            // Поднимаемся по дереву визуальных элементов
            while (clickedElement != null && !(clickedElement is Visual || clickedElement is Visual3D))
            {
                // Проверяем, что это не элемент типа Run
                if (clickedElement is Run)
                {
                    // Просто выходим, если наткнулись на Run
                    return;
                }
                clickedElement = VisualTreeHelper.GetParent(clickedElement);
            }
            // Получаем источник события
            var source = e.OriginalSource as DependencyObject;

            // Ищем DataGrid по иерархии визуальных элементов
            while (source != null && !(source is DataGrid))
            {
                source = VisualTreeHelper.GetParent(source);
            }

            // Если DataGrid не найден (клик был вне DataGrid), сбрасываем выделение
            if (source == null)
            {
                partsDataGrid.UnselectAll();
            }
        }

        private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
        {
            // Очищаем текстовое поле и восстанавливаем PlaceholderText
            searchTextBox.Text = string.Empty;
            searchTextBox.Text = PlaceholderText;
            searchTextBox.Foreground = System.Windows.Media.Brushes.Gray;
            clearSearchButton.Visibility = Visibility.Collapsed;
            searchTextBox.CaretIndex = 0; // Перемещаем курсор в начало
        }
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (searchTextBox.Text == PlaceholderText)
            {
                actualSearchText = string.Empty;
                clearSearchButton.Visibility = Visibility.Collapsed;
                return;
            }

            actualSearchText = searchTextBox.Text.Trim().ToLower();

            // Показываем или скрываем кнопку очистки в зависимости от наличия текста
            clearSearchButton.Visibility = string.IsNullOrEmpty(actualSearchText) ? Visibility.Collapsed : Visibility.Visible;

            // Перезапуск таймера при изменении текста
            searchDelayTimer.Stop();
            searchDelayTimer.Start();
        }
        private void SearchTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (searchTextBox.Text == PlaceholderText)
            {
                searchTextBox.Text = string.Empty;
                searchTextBox.Foreground = System.Windows.Media.Brushes.Black;
            }
        }
        private void SearchTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(searchTextBox.Text))
            {
                searchTextBox.Text = PlaceholderText;
                searchTextBox.Foreground = System.Windows.Media.Brushes.Gray;
                actualSearchText = string.Empty; // Очищаем фактический текст поиска
                clearSearchButton.Visibility = Visibility.Collapsed; // Скрываем кнопку очистки
            }
        }
        private void SearchDelayTimer_Tick(object sender, EventArgs e)
        {
            searchDelayTimer.Stop();
            partsDataView.View.Refresh(); // Обновление фильтрации
        }

        private void PartsData_Filter(object sender, FilterEventArgs e)
        {
            if (e.Item is PartData partData)
            {
                if (string.IsNullOrEmpty(actualSearchText))
                {
                    e.Accepted = true;
                    return;
                }

                bool matches = partData.PartNumber?.ToLower().Contains(actualSearchText) == true ||
                               partData.Description?.ToLower().Contains(actualSearchText) == true ||
                               partData.ModelState?.ToLower().Contains(actualSearchText) == true ||
                               partData.Material?.ToLower().Contains(actualSearchText) == true ||
                               partData.Thickness?.ToLower().Contains(actualSearchText) == true ||
                               partData.Quantity.ToString().Contains(actualSearchText) == true;

                // Проверка на пользовательские свойства
                if (!matches)
                {
                    foreach (var property in partData.CustomProperties)
                    {
                        if (property.Value.ToLower().Contains(actualSearchText))
                        {
                            matches = true;
                            break;
                        }
                    }
                }

                e.Accepted = matches;
            }
            else
            {
                e.Accepted = false;
            }
        }
        private void partsDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            // Устанавливаем чередование цветов
            if (e.Row.GetIndex() % 2 == 0)
            {
                e.Row.Background = new SolidColorBrush(Colors.White);
            }
            else
            {
                e.Row.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 245, 245)); // #F5F5F5
            }
        }
        private void partsDataGrid_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(PresetIProperty)))
            {
                var droppedData = e.Data.GetData(typeof(PresetIProperty)) as PresetIProperty;
                if (droppedData != null && !IsColumnPresent(droppedData.InternalName))
                {
                    // Определяем позицию мыши
                    System.Windows.Point position = e.GetPosition(partsDataGrid);

                    // Найти колонку, перед которой должна быть вставлена новая колонка
                    int insertIndex = -1;
                    double totalWidth = 0;

                    for (int i = 0; i < partsDataGrid.Columns.Count; i++)
                    {
                        totalWidth += partsDataGrid.Columns[i].ActualWidth;
                        if (position.X < totalWidth)
                        {
                            insertIndex = i;
                            break;
                        }
                    }

                    // Добавляем колонку на нужную позицию
                    AddIPropertyColumn(droppedData, insertIndex);
                }
            }
        }
        public void AddIPropertyColumn(PresetIProperty iProperty, int? insertIndex = null)
        {
            // Проверяем, существует ли уже колонка с таким заголовком
            if (IsColumnPresent(iProperty.InternalName))
                return;

            DataGridColumn column;

            // Проверка, если это колонка с изображением детали или развертки
            if (iProperty.InternalName == "Изображение детали" || iProperty.InternalName == "Изображение развертки")
            {
                column = new DataGridTemplateColumn
                {
                    Header = iProperty.InternalName,
                    CellTemplate = new DataTemplate
                    {
                        VisualTree = new FrameworkElementFactory(typeof(System.Windows.Controls.Image))
                    }
                };
                // Устанавливаем привязку данных для колонки изображения
                (column as DataGridTemplateColumn).CellTemplate.VisualTree.SetBinding(System.Windows.Controls.Image.SourceProperty, new System.Windows.Data.Binding(iProperty.InventorPropertyName));
            }
            else
            {
                // Создаем текстовую колонку
                var textColumn = new DataGridTextColumn
                {
                    Header = iProperty.InternalName,
                    Binding = new System.Windows.Data.Binding(iProperty.InventorPropertyName)
                };

                // Применяем специальный стиль для колонки "Количество"
                if (iProperty.InternalName == "Количество")
                {
                    textColumn.ElementStyle = partsDataGrid.FindResource("QuantityCellStyle") as System.Windows.Style;
                }
                else
                {
                    // Применяем стиль CenteredCellStyle для всех остальных текстовых колонок
                    textColumn.ElementStyle = partsDataGrid.FindResource("CenteredCellStyle") as System.Windows.Style;
                }

                column = textColumn;
            }

            // Если указан индекс вставки, вставляем колонку в нужное место
            if (insertIndex.HasValue && insertIndex.Value >= 0 && insertIndex.Value < partsDataGrid.Columns.Count)
            {
                partsDataGrid.Columns.Insert(insertIndex.Value, column);
            }
            else
            {
                // В противном случае добавляем колонку в конец
                partsDataGrid.Columns.Add(column);
            }

            // Обновляем список доступных свойств
            var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
            selectIPropertyWindow?.UpdateAvailableProperties();
        }
        private void AddPartData(PartData partData)
        {
            partsData.Add(partData);
            partsDataView.View.Refresh(); // Обновляем фильтрацию и отображение

            // Обновляем состояние кнопки после добавления данных
            UpdateEditButtonState();
        }
        private void PartsDataGrid_ColumnReordering(object sender, DataGridColumnReorderingEventArgs e)
        {
            _reorderingColumn = e.Column;
            _isColumnDraggedOutside = false;

            // Создаем фантомный заголовок с фиксированным размером
            var header = new TextBlock
            {
                Text = _reorderingColumn.Header.ToString(),
                Background = System.Windows.Media.Brushes.LightGray,
                FontSize = 12, // Установите желаемый размер шрифта
                Padding = new Thickness(5),
                Opacity = 0.7, // Полупрозрачность
                TextAlignment = TextAlignment.Center,
                Width = _reorderingColumn.ActualWidth,
                Height = 30,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                IsHitTestVisible = false
            };

            // Создаем и добавляем Adorner
            _adornerLayer = AdornerLayer.GetAdornerLayer(partsDataGrid);
            _headerAdorner = new HeaderAdorner(partsDataGrid, header);
            _adornerLayer.Add(_headerAdorner);

            // Изначально размещаем Adorner там, где находится заголовок
            UpdateAdornerPosition(e);
        }
        private void PartsDataGrid_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                // Определяем, кликнул ли пользователь на заголовке колонки
                var header = e.OriginalSource as FrameworkElement;
                if (header?.Parent is DataGridColumnHeader columnHeader)
                {
                    // Начинаем перетаскивание
                    StartColumnReordering(columnHeader.Column);
                }
            }
        }
        private void StartColumnReordering(DataGridColumn column)
        {
            _reorderingColumn = column;
            _isColumnDraggedOutside = false;

            // Создаем фантомный заголовок с фиксированным размером
            var header = new TextBlock
            {
                Text = _reorderingColumn.Header.ToString(),
                Background = System.Windows.Media.Brushes.LightGray,
                FontSize = 12, // Установите желаемый размер шрифта
                Padding = new Thickness(5),
                Opacity = 0.7, // Полупрозрачность
                TextAlignment = TextAlignment.Center,
                Width = _reorderingColumn.ActualWidth,
                Height = 30,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                IsHitTestVisible = false
            };

            // Создаем и добавляем Adorner
            _adornerLayer = AdornerLayer.GetAdornerLayer(partsDataGrid);
            _headerAdorner = new HeaderAdorner(partsDataGrid, header);
            _adornerLayer.Add(_headerAdorner);

            // Изначально размещаем Adorner там, где находится заголовок
            var position = Mouse.GetPosition(partsDataGrid);
            _headerAdorner.RenderTransform = new TranslateTransform(
                position.X - _headerAdorner.DesiredSize.Width / 2,
                position.Y - _headerAdorner.DesiredSize.Height / 2
            );
        }
        private void PartsDataGrid_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_headerAdorner == null || _reorderingColumn == null)
                return;

            // Обновляем позицию Adorner
            UpdateAdornerPosition(null);

            // Получаем текущую позицию мыши относительно DataGrid
            var position = e.GetPosition(partsDataGrid);

            // Проверяем, если мышь вышла за границы DataGrid
            if (position.X > partsDataGrid.ActualWidth || position.Y < 0 || position.Y > partsDataGrid.ActualHeight)
            {
                _isColumnDraggedOutside = true;
                SetAdornerColor(System.Windows.Media.Brushes.Red); // Меняем цвет на красный и добавляем надпись "УДАЛИТЬ"
            }
            else
            {
                _isColumnDraggedOutside = false;
                SetAdornerColor(System.Windows.Media.Brushes.LightGray); // Возвращаем исходный цвет и удаляем надпись "УДАЛИТЬ"
            }
        }
        private void SetAdornerColor(System.Windows.Media.Brush color)
        {
            if (_headerAdorner != null && _headerAdorner.Child is TextBlock textBlock)
            {
                textBlock.Background = color;

                if (color == System.Windows.Media.Brushes.Red)
                {
                    textBlock.Text = $"❎ {_reorderingColumn.Header}";
                    textBlock.Foreground = System.Windows.Media.Brushes.White; // Белый текст на красном фоне
                    textBlock.FontWeight = FontWeights.Bold; // Жирный шрифт для улучшенной видимости
                    textBlock.Padding = new Thickness(5); // Дополнительное внутреннее пространство
                    textBlock.TextAlignment = TextAlignment.Center; // Центрирование текста
                    textBlock.VerticalAlignment = VerticalAlignment.Center;
                    textBlock.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
                }
                else
                {
                    textBlock.Text = _reorderingColumn.Header.ToString();
                    textBlock.Foreground = System.Windows.Media.Brushes.Black; // Черный текст на сером фоне
                    textBlock.FontWeight = FontWeights.Normal;
                    textBlock.Padding = new Thickness(5);
                }

                // Убираем фиксированные размеры и устанавливаем авторазмеры
                textBlock.Width = Double.NaN;
                textBlock.Height = Double.NaN;
                textBlock.Measure(new System.Windows.Size(Double.PositiveInfinity, Double.PositiveInfinity));
                textBlock.Arrange(new Rect(textBlock.DesiredSize));
            }
        }
        private void PartsDataGrid_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_adornerLayer != null && _headerAdorner != null)
            {
                _adornerLayer.Remove(_headerAdorner);
                _headerAdorner = null;
            }

            if (_reorderingColumn != null && _isColumnDraggedOutside)
            {
                string columnName = _reorderingColumn.Header.ToString();

                // Убираем колонку из DataGrid
                var columnToRemove = partsDataGrid.Columns.FirstOrDefault(c => c.Header.ToString() == columnName);
                if (columnToRemove != null)
                {
                    partsDataGrid.Columns.Remove(columnToRemove);

                    // Удаляем кастомное свойство из списка и данных деталей
                    RemoveCustomIPropertyColumn(columnName);

                    // Обновляем список доступных свойств, если это кастомное свойство
                    var selectIPropertyWindow = System.Windows.Application.Current.Windows.OfType<SelectIPropertyWindow>().FirstOrDefault();
                    selectIPropertyWindow?.UpdateAvailableProperties();
                }
            }

            _reorderingColumn = null;
            _isColumnDraggedOutside = false;
        }
        private void PartsDataGrid_ColumnReordered(object sender, DataGridColumnEventArgs e)
        {
            _reorderingColumn = null;
            _isColumnDraggedOutside = false;

            // Убираем Adorner
            if (_adornerLayer != null && _headerAdorner != null)
            {
                _adornerLayer.Remove(_headerAdorner);
                _headerAdorner = null;
            }
        }
        private void UpdateAdornerPosition(DataGridColumnReorderingEventArgs e)
        {
            if (_headerAdorner == null)
                return;

            // Получаем текущую позицию мыши относительно родителя DataGrid
            var position = Mouse.GetPosition(_adornerLayer);

            // Применяем трансформацию для перемещения Adorner
            _headerAdorner.RenderTransform = new TranslateTransform(
                position.X - (_headerAdorner.RenderSize.Width / 2),
                position.Y - (_headerAdorner.RenderSize.Height / 1.8)
            );
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

        private double GetColumnStartX(DataGrid dataGrid, int columnIndex)
        {
            double startX = 0;
            for (int i = 0; i < columnIndex; i++)
            {
                startX += dataGrid.Columns[i].ActualWidth;
            }
            return startX;
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
            // Проверяем текущий текст вместе с новым вводом
            string newText = (sender as System.Windows.Controls.TextBox).Text.Insert((sender as System.Windows.Controls.TextBox).CaretIndex, e.Text);

            // Проверяем, является ли текст положительным числом больше нуля
            e.Handled = !IsTextAllowed(newText);
        }

        private static bool IsTextAllowed(string text)
        {
            // Регулярное выражение для проверки положительных чисел
            Regex regex = new Regex("^[1-9][0-9]*$"); // Разрешены только положительные числа больше нуля
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

            // Проверка на наличие активного документа
            Document? doc = ThisApplication.ActiveDocument as Document;
            if (doc == null)
            {
                MessageBoxHelper.ShowNoDocumentOpenWarning();
                return;
            }

            bool isBackgroundMode = backgroundModeCheckBox.IsChecked == true;
            if (isBackgroundMode)
            {
                ThisApplication.UserInterfaceManager.UserInteractionDisabled = true;
            }

            modelStateInfoRunBottom.Text = string.Empty;

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
            clearButton.IsEnabled = false;
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
            itemCounter = 1;

            var progress = new Progress<PartData>(partData =>
            {
                partData.Item = itemCounter;
                partsData.Add(partData);
                partsDataGrid.Items.Refresh();
                itemCounter++;
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
                            await Task.Delay(10);
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
                        await Task.Delay(10);
                    }
                }
            }

            stopwatch.Stop();
            string elapsedTime = GetElapsedTime(stopwatch.Elapsed);

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
            UpdateFileCountLabel(isCancelled ? 0 : partCount);

            scanButton.IsEnabled = true;
            cancelButton.IsEnabled = false;
            exportButton.IsEnabled = partCount > 0 && !isCancelled;
            clearButton.IsEnabled = partsData.Count > 0;
            lastScannedDocument = isCancelled ? null : doc;

            if (isBackgroundMode)
            {
                ThisApplication.UserInterfaceManager.UserInteractionDisabled = false;
            }

            if (isCancelled)
            {
                System.Windows.MessageBox.Show("Процесс сканирования был прерван.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private void UpdateSubfolderOptions()
        {
            // Опция "В папку с деталью" - чекбокс и текстовое поле должны быть отключены
            if (partFolderRadioButton.IsChecked == true)
            {
                enableSubfolderCheckBox.IsEnabled = false;
                enableSubfolderCheckBox.IsChecked = false;
                subfolderNameTextBox.IsEnabled = false;
            }
            else
            {
                // Для всех остальных опций чекбокс активен
                enableSubfolderCheckBox.IsEnabled = true;
                subfolderNameTextBox.IsEnabled = enableSubfolderCheckBox.IsChecked == true;
            }
        }
        private void SubfolderNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string forbiddenChars = "\\/:*?\"<>|";
            var textBox = sender as System.Windows.Controls.TextBox;
            int caretIndex = textBox.CaretIndex;

            foreach (var ch in forbiddenChars)
            {
                if (textBox.Text.Contains(ch))
                {
                    textBox.Text = textBox.Text.Replace(ch.ToString(), string.Empty);
                    caretIndex--; // Уменьшаем позицию курсора на 1
                }
            }

            // Возвращаем курсор на правильное место после удаления символов
            textBox.CaretIndex = Math.Max(caretIndex, 0);
        }

        private void PartFolderRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdateSubfolderOptions();
        }

        private void ComponentFolderRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdateSubfolderOptions();
        }

        private void ChooseFolderRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdateSubfolderOptions();
        }

        private void FixedFolderRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdateSubfolderOptions();
        }

        private void EnableSubfolderCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            subfolderNameTextBox.IsEnabled = true;
        }

        private void EnableSubfolderCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            subfolderNameTextBox.IsEnabled = false;
            subfolderNameTextBox.Text = string.Empty;
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
        private void ReadIPropertyValuesFromPart(PartDocument partDoc, PartData partData)
        {
            var propertySets = partDoc.PropertySets;

            // Проверка и чтение свойства "Author"
            if (IsColumnPresent("Автор"))
            {
                partData.Author = GetProperty(propertySets["Summary Information"], "Author");
            }

            // Проверка и чтение свойства "Revision"
            if (IsColumnPresent("Ревизия"))
            {
                partData.Revision = GetProperty(propertySets["Design Tracking Properties"], "Revision Number");
            }

            // Проверка и чтение свойства "Project"
            if (IsColumnPresent("Проект"))
            {
                partData.Project = GetProperty(propertySets["Summary Information"], "Project");
            }

            // Проверка и чтение свойства "Stock number"
            if (IsColumnPresent("Инвентарный номер"))
            {
                partData.StockNumber = GetProperty(propertySets["Design Tracking Properties"], "Stock Number");
            }
        }

        private async Task<PartData> GetPartDataAsync(string partNumber, int quantity, BOM bom, int itemNumber, PartDocument? partDoc = null)
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
                PartNumberExpression = await GetPropertyExpressionOrValueAsync(partDoc, "Part Number", true),
                Description = await GetPropertyExpressionOrValueAsync(partDoc, "Description"),
                DescriptionExpression = await GetPropertyExpressionOrValueAsync(partDoc, "Description", true),
                Material = await GetPropertyExpressionOrValueAsync(partDoc, "Material"),
                Thickness = await GetPropertyExpressionOrValueAsync(partDoc, "Thickness"),
                ModelState = partDoc.ModelStateName,
                OriginalQuantity = quantity,
                Quantity = quantity,
                Item = itemNumber
            };

            // Получаем выражения и значения для пользовательских свойств
            foreach (var customProperty in customPropertiesList)
            {
                // Асинхронно получаем значения пользовательских свойств
                partData.CustomProperties[customProperty] = await GetPropertyExpressionOrValueAsync(partDoc, customProperty);
                partData.CustomPropertyExpressions[customProperty] = await GetPropertyExpressionOrValueAsync(partDoc, customProperty, true);
            }

            // Асинхронно получаем превью-изображение (если доступно)
            partData.Preview = await GetThumbnailAsync(partDoc);

            // Проверяем наличие развертки для листовых деталей и задаем цвет
            var smCompDef = partDoc.ComponentDefinition as SheetMetalComponentDefinition;
            bool hasFlatPattern = smCompDef != null && smCompDef.HasFlatPattern;
            partData.FlatPatternColor = hasFlatPattern ? System.Windows.Media.Brushes.Green : System.Windows.Media.Brushes.Red;

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
        private void SelectFixedFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fixedFolderPath = dialog.SelectedPath;

                if (fixedFolderPath.Length > 40)
                {
                    // Гарантируем, что показываем именно последние 40 символов
                    fixedFolderPathTextBlock.Text = "..." + fixedFolderPath.Substring(fixedFolderPath.Length - 40);
                }
                else
                {
                    fixedFolderPathTextBlock.Text = fixedFolderPath;
                }
            }
        }
        private bool PrepareForExport(out string targetDir, out int multiplier, out Stopwatch stopwatch)
        {
            stopwatch = new Stopwatch();
            targetDir = string.Empty;
            multiplier = 1; // Присваиваем значение по умолчанию

            if (chooseFolderRadioButton.IsChecked == true)
            {
                // Открываем диалог выбора папки
                var dialog = new System.Windows.Forms.FolderBrowserDialog();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    targetDir = dialog.SelectedPath;
                }
                else
                {
                    return false;
                }
            }
            else if (componentFolderRadioButton.IsChecked == true)
            {
                // Папка компонента
                targetDir = System.IO.Path.GetDirectoryName(ThisApplication.ActiveDocument.FullFileName);
            }
            else if (fixedFolderRadioButton.IsChecked == true)
            {
                if (string.IsNullOrEmpty(fixedFolderPath))
                {
                    System.Windows.MessageBox.Show("Выберите фиксированную папку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                targetDir = fixedFolderPath;
            }

            if (enableSubfolderCheckBox.IsChecked == true && !string.IsNullOrEmpty(subfolderNameTextBox.Text))
            {
                targetDir = System.IO.Path.Combine(targetDir, subfolderNameTextBox.Text);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
            }

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

            bool isBackgroundMode = backgroundModeCheckBox.IsChecked == true;
            if (isBackgroundMode)
            {
                ThisApplication.UserInterfaceManager.UserInteractionDisabled = true;
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
                if (isBackgroundMode)
                {
                    ThisApplication.UserInterfaceManager.UserInteractionDisabled = false;
                }
                return;
            }

            clearButton.IsEnabled = false;
            scanButton.IsEnabled = false;

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

            clearButton.IsEnabled = true;
            scanButton.IsEnabled = true;

            if (isBackgroundMode)
            {
                ThisApplication.UserInterfaceManager.UserInteractionDisabled = false;
            }

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

            // Выполняем доступ к элементам управления через Dispatcher
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

                // Доступ к RadioButton через Dispatcher для потока безопасности
                Dispatcher.Invoke(() =>
                {
                    if (partFolderRadioButton.IsChecked == true)
                    {
                        targetDir = GetPartDocumentFullPath(partNumber);
                        targetDir = System.IO.Path.GetDirectoryName(targetDir);
                    }
                });

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

        private PartDocument OpenPartDocument(string partNumber)
        {
            var docs = ThisApplication.Documents;

            // Попытка найти открытый документ
            foreach (Document doc in docs)
            {
                if (doc is PartDocument pd && GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                {
                    return pd; // Возвращаем найденный открытый документ
                }
            }

            // Если документ не найден
            System.Windows.MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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

                // Проверка на null перед изменением видимости кнопки
                if (clearMultiplierButton != null)
                {
                    clearMultiplierButton.Visibility = multiplier > 1 ? Visibility.Visible : Visibility.Collapsed;
                }
            }
            else
            {
                // Если введенное значение некорректное, сбрасываем текст на "1" и скрываем кнопку сброса
                multiplierTextBox.Text = "1";
                UpdateQuantitiesWithMultiplier(1);

                if (clearMultiplierButton != null)
                {
                    clearMultiplierButton.Visibility = Visibility.Collapsed;
                }
            }
        }
        private void ClearMultiplierButton_Click(object sender, RoutedEventArgs e)
        {
            // Сбрасываем множитель на 1
            multiplierTextBox.Text = "1";
            UpdateQuantitiesWithMultiplier(1);

            // Скрываем кнопку сброса
            clearMultiplierButton.Visibility = Visibility.Collapsed;
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

            // Обновляем состояние кнопки после очистки данных
            UpdateEditButtonState();
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

                // Проверка на null перед использованием fullPath
                if (string.IsNullOrEmpty(fullPath))
                {
                    System.Windows.MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    continue;
                }

                // Проверка существования файла
                if (System.IO.File.Exists(fullPath))
                {
                    try
                    {
                        string argument = "/select, \"" + fullPath + "\"";
                        Process.Start("explorer.exe", argument);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"Ошибка при открытии проводника для файла по пути {fullPath}: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show($"Файл по пути {fullPath}, связанный с номером детали {partNumber}, не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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

                // Проверка на null перед использованием fullPath
                if (string.IsNullOrEmpty(fullPath))
                {
                    System.Windows.MessageBox.Show($"Файл, связанный с номером детали {partNumber}, не найден среди открытых.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    continue;
                }

                // Проверка существования файла
                if (System.IO.File.Exists(fullPath))
                {
                    try
                    {
                        // Открытие документа в Inventor
                        ThisApplication.Documents.Open(fullPath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"Ошибка при открытии файла по пути {fullPath}: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show($"Файл по пути {fullPath}, связанный с номером детали {partNumber}, не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string GetPartDocumentFullPath(string partNumber)
        {
            var docs = ThisApplication.Documents;

            // Попытка найти открытый документ
            foreach (Document doc in docs)
            {
                if (doc is PartDocument pd && GetProperty(pd.PropertySets["Design Tracking Properties"], "Part Number") == partNumber)
                {
                    return pd.FullFileName; // Возвращаем полный путь найденного документа
                }
            }

            // Если документ не найден среди открытых
            System.Windows.MessageBox.Show($"Документ с номером детали {partNumber} не найден среди открытых.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return null; // Возвращаем null, если документ не найден
        }

        private void partsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Активация пункта "Редактировать iProperty" только при выборе одной строки
            editIPropertyMenuItem.IsEnabled = partsDataGrid.SelectedItems.Count == 1;
        }
        private async void AddCustomIPropertyButton_Click(object sender, RoutedEventArgs e)
        {
            var inputDialog = new CustomIPropertyDialog();
            if (inputDialog.ShowDialog() == true)
            {
                string customPropertyName = inputDialog.CustomPropertyName;

                // Проверка на существование в списке пользовательских свойств
                if (customPropertiesList.Contains(customPropertyName))
                {
                    System.Windows.MessageBox.Show($"Свойство '{customPropertyName}' уже добавлено.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Проверка на существование колонки
                if (partsDataGrid.Columns.Any(c => (c.Header as string) == customPropertyName))
                {
                    System.Windows.MessageBox.Show($"Столбец с именем '{customPropertyName}' уже существует.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                customPropertiesList.Add(customPropertyName);

                // Добавляем новую колонку в DataGrid
                AddCustomIPropertyColumn(customPropertyName);

                // Если включен фоновый режим, отключаем взаимодействие с UI
                if (backgroundModeCheckBox.IsChecked == true)
                {
                    ThisApplication.UserInterfaceManager.UserInteractionDisabled = true;
                }

                try
                {
                    // Заполняем значения Custom iProperty для уже существующих деталей асинхронно
                    await FillCustomPropertyAsync(customPropertyName);
                }
                finally
                {
                    // Восстанавливаем взаимодействие с UI после завершения асинхронной операции
                    if (backgroundModeCheckBox.IsChecked == true)
                    {
                        ThisApplication.UserInterfaceManager.UserInteractionDisabled = false;
                    }
                }
            }
        }
        private async Task FillCustomPropertyAsync(string customPropertyName)
        {
            foreach (var partData in partsData)
            {
                // Получаем значение custom iProperty асинхронно
                string propertyValue = await Task.Run(() => GetCustomIPropertyValue(partData.PartNumber, customPropertyName));

                // Добавляем значение в коллекцию свойств для конкретного элемента
                partData.AddCustomProperty(customPropertyName, propertyValue);

                // Обновляем UI
                partsDataGrid.Items.Refresh();

                // Проверяем, включен ли фоновый режим
                if (backgroundModeCheckBox.IsChecked == true)
                {
                    // Дополнительная задержка для плавности и снижения нагрузки на UI
                    await Task.Delay(50);
                }
            }
        }
        private void RemoveCustomIPropertyColumn(string propertyName)
        {
            // Удаление столбца из DataGrid
            var columnToRemove = partsDataGrid.Columns.FirstOrDefault(c => (c.Header as string) == propertyName);
            if (columnToRemove != null)
            {
                partsDataGrid.Columns.Remove(columnToRemove);
            }

            // Удаление всех данных, связанных с этим Custom IProperty
            foreach (var partData in partsData)
            {
                if (partData.CustomProperties.ContainsKey(propertyName))
                {
                    partData.RemoveCustomProperty(propertyName);
                }
            }

            // Обновляем список доступных свойств
            customPropertiesList.Remove(propertyName); // Не забудьте удалить свойство из списка доступных свойств
            partsDataGrid.Items.Refresh();
        }

        private void AddCustomIPropertyColumn(string propertyName)
        {
            // Проверяем, существует ли уже колонка с таким именем или заголовком
            if (partsDataGrid.Columns.Any(c => (c.Header as string) == propertyName))
            {
                System.Windows.MessageBox.Show($"Столбец с именем '{propertyName}' уже существует.", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Создаем новую колонку
            var column = new DataGridTextColumn
            {
                Header = propertyName,
                Binding = new System.Windows.Data.Binding($"CustomProperties[{propertyName}]")
                {
                    Mode = BindingMode.OneWay // Устанавливаем режим привязки
                },
                Width = DataGridLength.Auto
            };

            // Применяем стиль CenteredCellStyle
            column.ElementStyle = partsDataGrid.FindResource("CenteredCellStyle") as System.Windows.Style;

            partsDataGrid.Columns.Add(column);
        }
        private void EditIProperty_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = partsDataGrid.SelectedItem as PartData;
            if (selectedItem == null || partsDataGrid.SelectedItems.Count != 1)
            {
                System.Windows.MessageBox.Show("Выберите одну строку для редактирования iProperty.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
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
                    partDoc.Save();
                    // partDoc.Close(); // Можно закрыть документ, если он не нужен больше открыт

                    // Обновляем свойства детали в таблице
                    selectedItem.PartNumber = editDialog.PartNumber;
                    selectedItem.Description = editDialog.Description;

                    // Обновляем данные в таблице
                    partsDataGrid.Items.Refresh();
                }
                else
                {
                    System.Windows.MessageBox.Show("Не удалось открыть документ детали для редактирования.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void SetProperty(PropertySet propertySet, string propertyName, string value)
        {
            try
            {
                Property property = propertySet[propertyName];
                property.Value = value;
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show($"Не удалось обновить свойство {propertyName}.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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

                    // Проверяем, существует ли свойство в данном PropertySet
                    if (propertySet[propertyName] is Property property)
                    {
                        return property.Value?.ToString() ?? string.Empty;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.ErrorCode == unchecked((int)0x80004005))
            {
                // Ловим ошибку и возвращаем пустую строку
                return string.Empty;
            }
            catch (Exception ex)
            {
                MessageBoxHelper.ShowErrorMessage($"Ошибка при получении свойства {propertyName}: {ex.Message}");
            }

            return string.Empty;
        }

    }

    public class PartData : INotifyPropertyChanged
    {
        private int quantity;
        private System.Windows.Media.Brush quantityColor;
        private int item;
        private bool isReadOnly = true;
        private string partNumber;
        private string description;
        private string material;
        private string thickness;
        private string originalPartNumber; // Хранит оригинальный PartNumber
        private bool isModified; // Флаг для отслеживания изменений

        // Выражения (если они есть)
        public string PartNumberExpression { get; set; }
        public string DescriptionExpression { get; set; }

        // Пользовательские iProperty
        private Dictionary<string, string> customProperties = new Dictionary<string, string>();
        public Dictionary<string, string> CustomProperties
        {
            get => customProperties;
            set
            {
                customProperties = value;
                OnPropertyChanged();
            }
        }

        // Выражения для пользовательских iProperty
        public Dictionary<string, string> CustomPropertyExpressions { get; set; } = new Dictionary<string, string>();

        // Порядковый номер элемента
        public int Item
        {
            get => item;
            set
            {
                item = value;
                OnPropertyChanged();
            }
        }

        // Основные свойства детали
        public string PartNumber
        {
            get => partNumber;
            set
            {
                partNumber = value;
                IsModified = true;  // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }

        public string OriginalPartNumber
        {
            get => originalPartNumber;
            set
            {
                originalPartNumber = value;
                OnPropertyChanged();
            }
        }

        public string ModelState { get; set; }

        public string Description
        {
            get => description;
            set
            {
                description = value;
                IsModified = true;  // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }

        public string Material
        {
            get => material;
            set
            {
                material = value;
                IsModified = true;  // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }

        public string Thickness
        {
            get => thickness;
            set
            {
                thickness = value;
                IsModified = true;  // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }

        public int OriginalQuantity { get; set; }

        public int Quantity
        {
            get => quantity;
            set
            {
                quantity = value;
                IsModified = true;  // Устанавливаем флаг изменений при изменении
                OnPropertyChanged();
            }
        }

        // Свойства для отображения изображений и цветов
        public BitmapImage Preview { get; set; }
        public BitmapImage DxfPreview { get; set; }
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

        // Новые свойства для хранения значений iProperty
        public string Author { get; set; }
        public string Revision { get; set; }
        public string Project { get; set; }
        public string StockNumber { get; set; }

        // Свойство для контроля редактируемости
        public bool IsReadOnly
        {
            get => isReadOnly;
            set
            {
                isReadOnly = value;
                OnPropertyChanged();
            }
        }

        // Флаг для отслеживания изменений
        public bool IsModified
        {
            get => isModified;
            set
            {
                isModified = value;
                OnPropertyChanged();
            }
        }

        // Флаг переопределения
        public bool IsOverridden { get; set; } = false;

        // Методы для работы с пользовательскими свойствами
        public void AddCustomProperty(string propertyName, string propertyValue)
        {
            if (customProperties.ContainsKey(propertyName))
            {
                customProperties[propertyName] = propertyValue;
            }
            else
            {
                customProperties.Add(propertyName, propertyValue);
            }
            IsModified = true;  // Устанавливаем флаг изменений
            OnPropertyChanged(nameof(CustomProperties));
        }

        public void RemoveCustomProperty(string propertyName)
        {
            if (customProperties.ContainsKey(propertyName))
            {
                customProperties.Remove(propertyName);
                IsModified = true;  // Устанавливаем флаг изменений
                OnPropertyChanged(nameof(CustomProperties));
            }
        }

        // Реализация интерфейса INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
    public class HeaderAdorner : Adorner
    {
        private readonly VisualCollection _visuals;
        private readonly UIElement _child;

        public HeaderAdorner(UIElement adornedElement, UIElement header)
            : base(adornedElement)
        {
            _visuals = new VisualCollection(this);
            _child = header;
            _visuals.Add(_child);

            IsHitTestVisible = false;
            Opacity = 0.8;  // Полупрозрачность
        }

        public UIElement Child => _child;

        protected override System.Windows.Size ArrangeOverride(System.Windows.Size finalSize)
        {
            _child.Arrange(new Rect(finalSize));
            return finalSize;
        }

        protected override Visual GetVisualChild(int index)
        {
            return _visuals[index];
        }

        protected override int VisualChildrenCount => _visuals.Count;
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

        public static void ShowPartNotFoundError(string partNumber, string filePath = null)
        {
            string message;

            if (string.IsNullOrEmpty(filePath))
            {
                message = $"Файл, связанный с номером детали {partNumber}, не найден среди открытых документов.";
            }
            else
            {
                message = $"Файл по пути {filePath}, связанный с номером детали {partNumber}, не найден.";
            }

            System.Windows.MessageBox.Show(message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
    public class PresetIProperty
    {
        public string InternalName { get; set; }  // Внутреннее имя колонки (например, "#")
        public string DisplayName { get; set; }   // Псевдоним для отображения в списке выбора (например, "#Нумерация")
        public string InventorPropertyName { get; set; }  // Соответствующее имя свойства iProperty в Inventor
        public bool ShouldBeAddedOnInit { get; set; } = false;  // Новый флаг для определения необходимости добавления при старте
    }
}
