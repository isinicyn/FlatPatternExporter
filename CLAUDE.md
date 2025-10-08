# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Важно:** Указывай только основные сведения о проекте без описаний того, что добавлено, удалено или улучшено.

## Стиль кодирования

При внесении изменений в проект придерживайтесь современного C# 12 стиля (.NET 8.0):

### Основные принципы
- Используйте последние возможности языка для улучшения читаемости и производительности
- Предпочитайте краткость без потери ясности
- Следуйте существующим паттернам в кодовой базе

## Структура репозитория

```
FlatPatternExporter/
├── CLAUDE.md                           # Инструкции для Claude
├── LICENSE.txt                         # Лицензия
├── README.md                           # Документация проекта
├── FlatPatternExporter.sln             # Файл решения Visual Studio
├── HELP/                               # Временные файлы справки
├── FlatPatternExporter.Updater/        # Проект апдейтера (WPF)
│   ├── FlatPatternExporter.Updater.csproj
│   ├── App.xaml(.cs)                   # Главное приложение апдейтера
│   ├── MainWindow.xaml(.cs)            # Окно процесса обновления
│   └── FPExport.ico                    # Иконка приложения
└── FlatPatternExporter/                # Основной проект
    ├── FlatPatternExporter.csproj      # Файл проекта
    ├── App.xaml(.cs)                   # Главное приложение
    ├── AssemblyInfo.cs                 # Информация о сборке
    ├── FPExport.ico                    # Иконка приложения
    │
    ├── Core/                           # Бизнес-логика и сервисы ядра
    │   ├── InventorManager.cs          # Работа с Inventor API
    │   ├── DocumentScanner.cs          # Сканирование документов и сборок
    │   ├── DocumentCache.cs            # Управление кешем документов
    │   ├── ConflictAnalyzer.cs         # Анализ конфликтов обозначений
    │   ├── DxfExporter.cs              # Экспорт разверток в DXF
    │   ├── ThumbnailGenerator.cs       # Генерация миниатюр
    │   ├── PartDataReader.cs           # Работа со свойствами деталей
    │   ├── PropertyManager.cs          # Менеджер свойств документов Inventor
    │   └── ConflictDataProcessor.cs    # Обработка данных конфликтов
    │
    ├── Libraries/                      # Внешние независимые библиотеки
    │   ├── DxfRenderer.cs              # Библиотека рендеринга DXF файлов
    │   ├── MarshalCore.cs              # Отдельная библиотека COM interop
    │   └── PopupNotificationService.cs    # Библиотека всплывающих уведомлений
    │
    ├── Enums/                          # Перечисления
    │   └── CommonEnums.cs              # Базовые enum и маппинги
    │
    ├── Models/                         # Модели данных
    │   ├── LayerSettingsClasses.cs     # Модели настроек слоев с INotifyPropertyChanged
    │   ├── AcadVersionItem.cs          # Модель для версий AutoCAD
    │   ├── PartData.cs                 # Основная модель данных детали
    │   ├── PresetIProperty.cs          # Модель предустановленных свойств
    │   ├── ScanProgress.cs             # Модель прогресса сканирования
    │   ├── PartConflictInfo.cs         # Информация о конфликтах
    │   ├── OperationResult.cs          # Результат операции
    │   ├── DocumentValidationResult.cs # Результат валидации документа
    │   ├── ReleaseInfo.cs              # Информация о релизе GitHub
    │   └── UpdateCheckResult.cs        # Результат проверки обновлений
    │
    ├── Services/                       # Вспомогательные сервисы
    │   ├── PropertyMetadataRegistry.cs # Центральный реестр метаданных
    │   ├── TokenService.cs             # Обработка токенов имен файлов
    │   ├── SettingsManager.cs          # Персистентность настроек
    │   ├── TemplatePresetManager.cs    # Управление пресетами шаблонов
    │   ├── VersionInfoService.cs       # Получение версии приложения и коммитов
    │   ├── PropertyListManager.cs      # Управление списками свойств с фильтрацией
    │   ├── LocalizationManager.cs      # Управление локализацией приложения
    │   ├── ExcelExportService.cs       # Экспорт данных в Excel/CSV
    │   ├── GitHubUpdateService.cs      # Работа с GitHub Releases API
    │   └── UpdateManager.cs            # Управление процессом обновления
    │
    ├── Utilities/                      # Утилиты
    │   ├── DxfOptimizer.cs             # Оптимизация DXF файлов
    │   └── VersionComparer.cs          # Сравнение версий приложения
    │
    ├── UI/                             # Пользовательский интерфейс
    │   ├── Windows/                    # Окна приложения
    │   │   ├── FlatPatternExporterMainWindow.xaml(.cs)
    │   │   ├── AboutWindow.xaml(.cs)
    │   │   ├── ConflictDetailsWindow.xaml(.cs)
    │   │   ├── SelectIPropertyWindow.xaml(.cs)
    │   │   ├── CustomMessageBox.xaml(.cs)
    │   │   └── UpdateWindow.xaml(.cs)
    │   │
    │   ├── Controls/                   # Пользовательские элементы
    │   │   ├── CustomTitleBar.xaml(.cs)
    │   │   ├── LayerSettingControl.xaml(.cs)
    │   │   └── HeaderAdorner.cs        # UI компонент для drag&drop колонок
    │   │
    │   └── Models/                     # Модели состояния UI
    │       └── UIState.cs              # Управление состоянием UI
    │
    ├── Converters/                     # WPF конвертеры
    │   ├── ColorToBrushConverter.cs            # Конвертация названия цвета в SolidColorBrush
    │   ├── LineTypeToGeometryConverter.cs      # Визуализация типов линий в LayerSettings
    │   ├── EnumToBooleanConverter.cs           # Универсальная привязка enum к RadioButton
    │   ├── EnumToDisplayNameConverter.cs       # Конвертер enum значений в локализованные имена
    │   ├── DynamicPropertyValueConverter.cs    # Извлечение значений свойств по имени пути
    │   ├── PropertyExpressionByNameConverter.cs # Определение видимости fx индикатора
    │   ├── IPictureDispConverter.cs            # Преобразование IPictureDisp в System.Drawing.Image
    │   ├── LocalizationConverter.cs            # Конвертер для локализации по значению
    │   └── LocalizationKeyConverter.cs         # Конвертер для локализации по ключу в параметре
    │
    ├── Styles/                         # XAML стили
    │   ├── ColorResources.xaml
    │   ├── DarkTheme.xaml
    │   ├── IconResources.xaml
    │   ├── DesignTokens.xaml
    │   ├── BaseControlStyles.xaml
    │   ├── ButtonStyles.xaml
    │   ├── DataGridStyles.xaml
    │   ├── SpecializedControlStyles.xaml
    │   ├── DocumentInfoStyles.xaml
    │   ├── ThemeControlStyles.xaml
    │   ├── ContextMenuStyles.xaml
    │   └── ScrollBarStyles.xaml
    │
    ├── Extensions/                     # Расширения разметки XAML
    │   └── LocalizeExtension.cs        # Расширение разметки для локализации в XAML
    │
    ├── Resources/                      # Ресурсы локализации
    │   ├── Strings.resx                # Строковые ресурсы (английский, по умолчанию)
    │   └── Strings.ru.resx             # Строковые ресурсы (русский)
    │
    └── Properties/                     # Свойства проекта
        ├── launchSettings.json
        └── PublishProfiles/
```

Это C# WPF приложение, которое является дополнением для Autodesk Inventor:

- **FlatPatternExporter** - Приложение для экспорта разверток деталей из листового металла с интегрированным функционалом управления слоями

Приложение построено с использованием:
- .NET 8.0 Windows Desktop
- WPF (Windows Presentation Foundation) 
- Поддержка Windows Forms включена
- Интеграция с Autodesk Inventor Interop API

## Команды сборки

### Сборка проекта
```bash
# Собрать решение
dotnet build FlatPatternExporter.sln --configuration Release

# Собрать конкретную конфигурацию для x64
dotnet build FlatPatternExporter.sln --configuration Release --arch x64
```

### Запуск приложения
```bash
# Запустить в режиме отладки
dotnet run --project FlatPatternExporter\FlatPatternExporter.csproj
```

## Архитектура

### Зависимости проекта
- Autodesk Inventor Interop: `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll`
- netDxf.netstandard (v3.0.1) для работы с DXF файлами
- Svg.Skia (v3.0.6) для конвертации SVG в PNG/BitmapImage
- ClosedXML (v0.105.0) для экспорта данных в Excel
- stdole (v17.14.40260) для COM interop

### Ключевые компоненты

**Core/ - Бизнес-логика ядра (namespace: FlatPatternExporter.Core):**
- `InventorManager` - работа с Inventor API (подключение, валидация документов, управление проектами)
- `DocumentScanner` - сканирование структуры документов, обход сборок, обработка BOM
- `DocumentCache` - управление кешем документов для оптимизации производительности
- `ConflictAnalyzer` - анализ конфликтов обозначений деталей
- `DxfExporter` - экспорт разверток в DXF с настройками слоев и оптимизацией
- `ThumbnailGenerator` - генерация миниатюр из DXF и документов Inventor
- `PartDataReader` - работа со свойствами деталей, чтение iProperties, заполнение данных
- `PropertyManager` - централизованное управление доступом к свойствам документов Inventor
- `ConflictDataProcessor` - обработка и трансформация данных конфликтов
- Классы: `ScanResult`, `ScanOptions`, `ExportContext`, `ExportOptions`, `DocumentInfo`

**Libraries/ - Внешние независимые библиотеки:**
- `DxfRenderer` - библиотека рендеринга DXF файлов (namespace: DxfRenderer)
- `MarshalCore` - COM interop библиотека (namespace: DefineEdge)
- `PopupNotificationService` - библиотека всплывающих уведомлений (namespace: WpfToolkit)

**Enums/ - Перечисления (namespace: FlatPatternExporter.Enums):**
- `CommonEnums` - базовые перечисления (ExportFolderType, ProcessingMethod, ProcessingStatus, AcadVersionType, DocumentType и др.)

**Models/ - Модели данных (namespace: FlatPatternExporter.Models):**
- `LayerSettingsClasses` - модели настроек слоев (LayerSetting, LayerDefaults, валидаторы)
- `AcadVersionItem` - модель для версий AutoCAD в ComboBox
- `PartData` - основная модель данных детали с INotifyPropertyChanged
- `PresetIProperty` - модель предустановленных свойств с INotifyPropertyChanged
- `ScanProgress` - модель прогресса сканирования
- `PartConflictInfo` - информация о конфликтах обозначений
- `OperationResult` - результат операции сканирования/экспорта
- `DocumentValidationResult` - результат валидации документа Inventor

**Services/ - Вспомогательные сервисы (namespace: FlatPatternExporter.Services):**
- `PropertyMetadataRegistry` - централизованный реестр метаданных свойств
- `TokenService` - специализированная обработка токенов имен файлов
- `SettingsManager` - персистентность настроек приложения
- `TemplatePresetManager` - управление пресетами шаблонов
- `VersionInfoService` - получение версии приложения и информации о коммитах
- `PropertyListManager` - управление списками свойств с фильтрацией и состоянием
- `LocalizationManager` - управление локализацией приложения с поддержкой русского и английского языков
- `ExcelExportService` - экспорт данных DataGrid в Excel (XLSX) и CSV форматы с поддержкой изображений

**Utilities/ - Утилиты (namespace: FlatPatternExporter.Utilities):**
- `DxfOptimizer` - оптимизация DXF файлов для различных версий AutoCAD

**UI/ - Интерфейс пользователя:**
- **Windows/ (namespace: FlatPatternExporter.UI.Windows)**: основные окна приложения
  - `FlatPatternExporterMainWindow` - главное окно приложения (координация UI и бизнес-логики)
  - `AboutWindow` - независимое окно информации о программе
  - `ConflictDetailsWindow` - окно конфликтов с инжекцией делегата открытия документов
  - `SelectIPropertyWindow` - окно выбора свойств с полной инжекцией зависимостей
  - `CustomMessageBox` - кастомное окно сообщений с поддержкой локализации и тем оформления
- **Controls/ (namespace: FlatPatternExporter.UI.Controls)**: пользовательские элементы управления
  - `CustomTitleBar` - переиспользуемый компонент кастомного заголовка окна с поддержкой WindowChrome
  - `LayerSettingControl` - контрол настройки слоя DXF
  - `HeaderAdorner` - UI компонент для визуализации drag&drop колонок DataGrid
- **Models/ (namespace: FlatPatternExporter.UI.Models)**: модели состояния UI
  - `UIState` - управление состоянием интерфейса (кнопки, прогресс, текст)

**Стили и ресурсы**:
- `ColorResources.xaml` - Цветовые ресурсы светлой темы
- `DarkTheme.xaml` - Цветовые ресурсы тёмной темы
- `IconResources.xaml` - SVG иконки и геометрия
- `DesignTokens.xaml` - Токены дизайна (размеры, отступы, Thickness)
- `BaseControlStyles.xaml` - Базовые стили для TextBox, ComboBox, CheckBox, RadioButton, Label
- `ButtonStyles.xaml` - Стили кнопок с иконками
- `DataGridStyles.xaml` - Стили для DataGrid элементов
- `SpecializedControlStyles.xaml` - Специализированные стили (ListBox, TreeView, Token, Separator)
- `DocumentInfoStyles.xaml` - Стили для заголовков и информации о документе
- `ThemeControlStyles.xaml` - Тема-зависимые стили для системных контролов
- `ContextMenuStyles.xaml` - Стили для контекстного меню и пунктов меню
- `ScrollBarStyles.xaml` - Стили для ScrollBar и Slider

**Converters/ - WPF конвертеры (namespace: FlatPatternExporter.Converters):**
- `ColorToBrushConverter` - конвертация названия цвета в SolidColorBrush
- `LineTypeToGeometryConverter` - визуализация типов линий в LayerSettings
- `EnumToBooleanConverter` - универсальная привязка enum к RadioButton
- `EnumToDisplayNameConverter` - конвертер enum значений в локализованные имена
- `DynamicPropertyValueConverter` - извлечение значений свойств по имени пути
- `PropertyExpressionByNameConverter` - определение видимости fx индикатора
- `IPictureDispConverter` - преобразование IPictureDisp в System.Drawing.Image
- `LocalizationConverter` - конвертация строковых ключей в локализованные значения
- `LocalizationKeyConverter` - конвертация ключей из параметра конвертера в локализованные значения

### Структура пространств имен
Проект следует конвенции .NET по соответствию пространств имен структуре папок:

**Основные пространства имен:**
- `FlatPatternExporter` - корневое пространство имен (App.xaml.cs)
- `FlatPatternExporter.Core` - бизнес-логика ядра в папке Core/
- `FlatPatternExporter.Enums` - перечисления в папке Enums/
- `FlatPatternExporter.Models` - модели данных в папке Models/
- `FlatPatternExporter.Services` - сервисы в папке Services/
- `FlatPatternExporter.Utilities` - утилиты в папке Utilities/
- `FlatPatternExporter.UI.Windows` - все окна в папке UI/Windows/
- `FlatPatternExporter.UI.Controls` - пользовательские элементы в UI/Controls/
- `FlatPatternExporter.UI.Models` - модели состояния UI в папке UI/Models/
- `FlatPatternExporter.Converters` - WPF конвертеры в Converters/
- `FlatPatternExporter.Extensions` - расширения разметки XAML в Extensions/
- `DxfRenderer` - независимая библиотека рендеринга DXF (Libraries/DxfRenderer.cs)
- `DefineEdge` - независимая библиотека COM interop (Libraries/MarshalCore.cs)
- `WpfToolkit` - независимая библиотека WPF компонентов (Libraries/PopupNotificationService.cs)

**Зависимости между пространствами имен:**
- `UI.Windows` → `Core`, `Enums`, `Models`, `Services`, `Utilities`, `UI.Controls`, `UI.Models`, `WpfToolkit`
- `UI.Controls` → `Models`
- `UI.Models` → отсутствуют внешние зависимости
- `Core` → `Enums`, `Models`, `Services`
- `Services` → `Enums`, `Models`
- `Utilities` → `Enums`
- `Converters` → `Models`
- `Models` → `Enums`
- `Libraries` → `Services` для внешних типов
- `WpfToolkit` → отсутствуют внешние зависимости

### Управление версиями
Проект использует версионирование на основе Git в процессе сборки:
- Формат версии: `1.2.0.{НомерРевизии}`
- Хеш Git коммита включен в InformationalVersion
- Автоматическое увеличение версии на основе количества Git коммитов

### Целевая платформа
- Среда выполнения: `win-x64`
- Фреймворк: `net8.0-windows10.0.26100.0`
- Поддержка Windows Forms и WPF включена
- Nullable reference types включены
- ImplicitUsings включены
- Тип отладочной информации: full для Debug, none для Release

### Особенности функционала экспорта разверток
Приложение специализируется на экспорте разверток деталей из листового металла:
- Сканирует сборки и детали Inventor для поиска компонентов из листового металла
- Собирает данные о деталях (свойства, количество, материалы, толщины)
- Экспортирует развертки в DXF формат с настраиваемыми параметрами слоев
- Поддерживает валидацию имен слоев с фильтрацией недопустимых символов
- Генерирует настраиваемые строки экспорта для DXF
- Анализирует конфликты обозначений деталей
- Поддерживает пакетный экспорт с миниатюрами разверток

### Система кеширования документов
Для оптимизации производительности при работе с множеством открытых документов в Inventor реализована система кеширования:

**Принцип работы:**
- Кеш заполняется автоматически во время обхода структуры сборки при сканировании
- Содержит ссылки на PartDocument объекты, индексированные по номеру детали (PartNumber)
- Дополнительно кеширует полные пути к файлам для быстрого доступа

**Сервис DocumentCache (namespace: FlatPatternExporter.Core):**
- Инкапсулирует всю логику кеширования
- `AddDocumentToCache()` - добавление документа в кеш
- `GetCachedPartDocument()` - получение документа из кеша
- `GetCachedPartPath()` - получение пути к файлу детали
- `ClearCache()` - очистка кеша

**Точки заполнения кеша в DocumentScanner:**
- `ProcessComponentOccurrences` - при методе обработки "Перебор"
- `ProcessBOMRowSimple` - при методе обработки "Спецификация"
- `ScanDocumentAsync` - для одиночных деталей

**Точки использования кеша в UI:**
- `PrepareExportContextAsync` - при быстром экспорте без предварительного сканирования

**Преимущества:**
- Изолирует обработку от других открытых сборок в Inventor
- Исключает множественный поиск по коллекции Documents
- Автоматически обновляется при каждом сканировании
- Кеширует только документы из активной сборки/детали

### Система токенов для имен файлов
Конструктор имени файла с интерактивным интерфейсом:
- Токенизированный ввод с drag&drop интерфейсом
- Real-time предпросмотр результата на основе реальных данных
- Валидация шаблона с визуальной и текстовой индикацией
- Централизованная обработка через TokenService с кэшированием
- Автоматическая санитизация недопустимых символов имени файла
- Система пресетов для сохранения популярных конфигураций шаблонов

### Система управления пресетами шаблонов
Отдельный компонент для управления пользовательскими пресетами шаблонов имен файлов:

**Архитектура системы:**
- `TemplatePresetManager` - инкапсулирует всю логику управления пресетами
- `TemplatePreset` - модель данных для отдельного пресета (имя + шаблон)
- Интеграция с `SettingsManager` для сохранения/загрузки в JSON

**Функциональные возможности:**
- Создание/удаление пресетов с валидацией имен и подтверждением
- Автоматическое предотвращение дублирования имен пресетов
- Сохранение выбранного пресета по индексу (более надежно чем по имени)
- Восстановление выбранного пресета при запуске приложения

**Интеграция с UI:**
- Привязка данных через `PresetManager.TemplatePresets` к ComboBox в XAML
- Делегирование всех операций из MainWindow к менеджеру
- Упрощение кода главного окна за счет инкапсуляции логики

### Централизованная система метаданных свойств
Проект использует единую систему управления метаданными через `PropertyMetadataRegistry`:

**Архитектура системы:**
- Централизованный реестр всех поддерживаемых свойств документов Inventor
- Типизированные определения свойств с полными метаданными
- Поддержка трех типов свойств: IProperty, Document, System
- Автоматическое маппирование внутренних имен на свойства Inventor

**Компоненты системы:**
- `PropertyDefinition` - класс метаданных для одного свойства
- `PropertyType` enum - типы свойств (IProperty, Document, System) 
- `Properties` Dictionary - основной реестр всех свойств
- Метаданные включают: DisplayName, ColumnHeader, Category, PropertySetName, InventorPropertyName

**Функциональные возможности:**
- Автоматическое округление числовых значений с настраиваемой точностью
- Маппинг значений для преобразования системных кодов в читаемый текст
- Поддержка редактируемых и только для чтения свойств
- Настройка поведения колонок (сортировка, поиск, шаблоны отображения)
- Группировка свойств по категориям

**Интеграция:**
- `PropertyManager` использует реестр для доступа к свойствам документов
- `SelectIPropertyWindow` получает список доступных свойств из реестра
- Автоматическое создание предустановленных свойств iProperty с корректными метаданными

### Управление пользовательскими свойствами (Custom iProperties)
- Добавление пользовательских свойств через текстовое поле в окне выбора свойств
- Автоматическое заполнение данных из "Inventor User Defined Properties"
- Поддержка фонового режима обработки

### Система настроек пользователя
- Автоматическое сохранение настроек в JSON файл при закрытии приложения
- Восстановление пользовательских колонок и параметров при запуске
- Расположение файла настроек: `%APPDATA%\FlatPatternExporter\settings.json`
- Сохранение всех свойств iProperty без зависимости от видимых колонок UI

### Архитектура UI - Система привязки данных
Проект использует декларативную архитектуру WPF/MVVM:

**Enum-based система для RadioButton групп:**
- `ExportFolderType` - управление выбором папки экспорта (5 опций)
- `ProcessingMethod` - выбор метода обработки (Перебор/Спецификация)

**Универсальные конвертеры:**
- `EnumToBooleanConverter` - универсальный конвертер для привязки enum к RadioButton
- `DynamicPropertyValueConverter` - извлечение значений свойств по имени пути
- `PropertyExpressionByNameConverter` - определение видимости fx индикатора по состоянию выражения
- Поддерживает любые enum через рефлексию и `Enum.TryParse`

**Система fx индикатора для редактируемых свойств:**
- Универсальный шаблон `EditableWithFxTemplate` для отображения текста с fx индикатором
- Система версионирования состояний выражений через `ExpressionStateVersion` 
- Пакетная обработка изменений состояний выражений через `BeginExpressionBatch`/`EndExpressionBatch`
- Передача пути свойства через `Tag` ячейки DataGrid для универсального связывания

**Система свойств с INotifyPropertyChanged:**
- CheckBox используют двустороннюю привязку данных (`TwoWay` binding)
- Публичные свойства: `ExcludeReferenceParts`, `OrganizeByMaterial`, `EnableSplineReplacement` и др.
- Вычисляемые свойства: `IsSubfolderCheckBoxEnabled` для зависимых состояний

**Декларативные XAML привязки:**
- `ElementName` binding для связи между элементами
- Отсутствие обработчиков событий `Checked`/`Unchecked`
- Отсутствие прямых обращений к UI элементам из code-behind

### Система локализации
Поддержка мультиязычного интерфейса через централизованную систему локализации:

**Архитектура локализации:**
- `LocalizationManager` - синглтон для управления переключением языков
- Поддержка английского (по умолчанию) и русского языков
- Динамическое переключение языка интерфейса в runtime
- Ресурсные файлы: `Strings.resx` (английский), `Strings.ru.resx` (русский)

**Компоненты системы:**
- `LocalizeExtension` - расширение разметки для удобного использования в XAML
- `LocalizationConverter` - конвертер для привязки по значению ключа
- `LocalizationKeyConverter` - конвертер для привязки по параметру
- Интеграция с `INotifyPropertyChanged` для автоматического обновления UI

**Использование в коде:**
- C#: `LocalizationManager.Instance.GetString("ключ")`
- XAML: `{ext:Localize ключ}` или через конвертеры в привязках

### Система кастомных заголовков окон
Все окна приложения используют единообразный кастомный заголовок через переиспользуемый компонент:

**Архитектура:**
- `CustomTitleBar` (UserControl) - автономный компонент заголовка с полной инкапсуляцией логики
- Централизованная конфигурация `WindowChrome` через именованный ресурс в App.xaml
- Интеграция с родительским окном через `Window.GetWindow(this)`

**Функциональные возможности:**
- Гибкая настройка через DependencyProperty (`Title`, `Icon`, `ShowMinimizeButton`, `ShowMaximizeButton`)
- Автоматическая обработка drag'n'drop для перемещения окна
- Двойной клик по заголовку для maximize/restore
- Автоматическое обновление иконки кнопки maximize/restore при изменении состояния окна
- Поддержка плавных системных анимаций Windows (DWM композиция)
- Aero Snap функциональность через WindowChrome

**Интеграция в окна:**
```xaml
<Window shell:WindowChrome.WindowChrome="{StaticResource CustomWindowChrome}" ...>
    <Border>
        <Grid>
            <CustomTitleBar Title="Заголовок" ShowMinimizeButton="False"/>
            <!-- Содержимое окна -->
        </Grid>
    </Border>
</Window>
```

**Преимущества:**
- Полная видимость в design-time режиме Visual Studio
- Hardware rendering без использования `AllowsTransparency`
- Централизованное управление настройками WindowChrome
- Переиспользуемость компонента во всех окнах приложения

### Система кастомных диалоговых окон
Приложение использует кастомное окно сообщений вместо стандартного MessageBox:

**Компонент:**
- `CustomMessageBox` - модальное диалоговое окно с поддержкой локализации и тем оформления

**Функциональные возможности:**
- Поддержка всех типов кнопок: OK, OK/Cancel, Yes/No, Yes/No/Cancel
- Поддержка типов иконок: Information, Warning, Error, Question
- Автоматическая локализация кнопок и заголовков
- Интеграция с системой тем (светлая/темная)
- Использование CustomTitleBar для единообразия интерфейса

**Использование:**
```csharp
// Простое информационное сообщение
CustomMessageBox.Show("Сообщение");

// С заголовком
CustomMessageBox.Show("Сообщение", "Заголовок");

// С кнопками и иконкой
var result = CustomMessageBox.Show(
    "Вы уверены?",
    "Подтверждение",
    MessageBoxButton.YesNo,
    MessageBoxImage.Question
);

// С владельцем окна
CustomMessageBox.Show(
    this,
    "Ошибка операции",
    "Ошибка",
    MessageBoxButton.OK,
    MessageBoxImage.Error
);
```

**Поддерживаемые иконки:**
- `InfoIconGeometry` - информационное сообщение (синий)
- `WarningIconGeometry` - предупреждение (оранжевый)
- `ErrorIconGeometry` - ошибка (красный)
- `QuestionIconGeometry` - вопрос (синий)

### Система автоматических обновлений
Приложение включает полностью автономную систему обновлений, интегрированную с GitHub Releases:

**Архитектура:**
- `GitHubUpdateService` - работа с GitHub Releases API для проверки и загрузки обновлений
- `UpdateManager` - координация процесса обновления (проверка, загрузка, запуск апдейтера)
- `VersionComparer` - сравнение версий в формате "2.1.0.X"
- `UpdateWindow` - UI окно прогресса обновления с ProgressBar и статусами (в основном приложении)
- `FlatPatternExporter.Updater` - отдельное WPF приложение для замены файлов с визуальным интерфейсом

**Функциональные возможности:**
- Проверка обновлений через кнопку "Проверить обновления" в AboutWindow
- Автоматическое сравнение текущей и последней версий
- Загрузка обновления с отображением прогресса
- Резервное копирование текущих файлов перед обновлением
- Автоматический перезапуск приложения после обновления
- Обработка ошибок с возможностью отката

**Процесс обновления:**
1. Пользователь нажимает "Проверить обновления" в окне "О программе"
2. Проверка наличия новых релизов через GitHub API
3. При наличии обновления открывается UpdateWindow с информацией о релизе
4. Пользователь подтверждает установку
5. Скачивание .exe файла с отображением прогресса
6. Запуск Updater с передачей PID процесса и путей к файлам
7. Основное приложение завершается
8. Updater ожидает завершения процесса, создает backup, заменяет файлы
9. Автоматический перезапуск обновленного приложения

**Модели данных:**
- `ReleaseInfo` - информация о релизе (версия, описание, assets, дата публикации)
- `UpdateCheckResult` - результат проверки (доступность обновления, версии, ошибки)

## Заметки по разработке

- Приложение требует установки Autodesk Inventor для функциональности COM interop
- Использует встроенную отладочную информацию для развертывания
- Включает пользовательские иконки и изображения ресурсов
- Поддерживает работу с различными форматами файлов через Inventor API
- Архитектура UI соответствует принципам WPF Data Binding
- Для новых RadioButton групп используется enum + `EnumToBooleanConverter`
- **НЕ добавлять комментарии об изменениях в код** - использовать только чистый код без комментариев о внесенных правках
- **НЕ детализировать содержимое папки HELP** - это временные справочные файлы, упоминать только общее назначение