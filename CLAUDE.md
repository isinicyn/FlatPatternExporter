# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Важно:** Указывай только основные сведения о проекте без описаний того, что добавлено, удалено или улучшено.

## Структура репозитория

```
FlatPatternExporter/
├── CLAUDE.md                           # Инструкции для Claude
├── LICENSE.txt                         # Лицензия
├── README.md                           # Документация проекта
├── FlatPatternExporter.sln             # Файл решения Visual Studio
├── HELP/                               # Временные файлы справки
└── FlatPatternExporter/                # Основной проект
    ├── FlatPatternExporter.csproj      # Файл проекта
    ├── App.xaml(.cs)                   # Главное приложение
    ├── AssemblyInfo.cs                 # Информация о сборке
    ├── FPExport.ico                    # Иконка приложения
    │
    ├── FlatPatternExporterMainWindow.xaml(.cs,.part2.cs)  # Главное окно
    ├── AboutWindow.xaml(.cs)           # Окно "О программе"
    ├── ConflictDetailsWindow.xaml(.cs) # Окно деталей конфликтов
    ├── SelectIPropertyWindow.xaml(.cs) # Окно выбора свойств
    │
    ├── LayerSettingControl.xaml(.cs)   # Элемент управления слоями
    │
    ├── Styles/                         # XAML стили
    │   ├── ColorResources.xaml
    │   ├── IconResources.xaml
    │   ├── ButtonStyles.xaml
    │   ├── DataGridStyles.xaml
    │   └── GeneralStyles.xaml
    │
    ├── Converters/                     # WPF конвертеры
    │   ├── ColorToBrushConverter.cs            # Конвертация названия цвета в SolidColorBrush
    │   ├── LineTypeToGeometryConverter.cs      # Визуализация типов линий в LayerSettings
    │   ├── EnumToBooleanConverter.cs           # Универсальная привязка enum к RadioButton
    │   ├── DynamicPropertyValueConverter.cs    # Извлечение значений свойств по имени пути
    │   ├── PropertyExpressionByNameConverter.cs # Определение видимости fx индикатора
    │   ├── ObjectToBooleanConverter.cs         # Проверка объекта на null для IsEnabled
    │   └── IPictureDispConverter.cs            # Преобразование IPictureDisp в System.Drawing.Image
    │
    ├── MarshalCore.cs                  # COM interop
    ├── DxfGenerator.cs                 # Генератор DXF
    ├── PropertyManager.cs              # Управление свойствами
    ├── PropertyMetadataRegistry.cs     # Централизованная система метаданных свойств
    ├── SettingsManager.cs              # Настройки пользователя
    ├── LayerSettingsClasses.cs         # Классы настроек слоев
    ├── TokenService.cs                 # Сервис работы с токенами
    ├── TemplatePresetManager.cs        # Управление пресетами имен файлов
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
- stdole (v17.14.40260) для COM interop

### Ключевые компоненты

**Основные окна**:
- `FlatPatternExporterMainWindow` - Главное окно приложения для экспорта разверток со встроенным функционалом управления слоями
- `ConflictDetailsWindow` - Окно деталей конфликтов
- `AboutWindow` - Окно "О программе"

**Диалоги управления свойствами**:
- `SelectIPropertyWindow` - Окно выбора свойств с интегрированным функционалом добавления пользовательских свойств

**Компоненты LayerSettings** (интегрированы):
- `LayerSettingControl` - Пользовательский элемент управления для конфигурации слоев
- `LayerSettingsClasses` - Классы моделей данных, конвертеры и вспомогательные методы
- `LayerSettingsHelper` - Статические методы для работы с настройками слоев

**Стили и ресурсы**:
- `ColorResources.xaml` - Централизованные цветовые ресурсы приложения
- `IconResources.xaml` - Централизованные SVG иконки и геометрия
- `ButtonStyles.xaml` - Специализированные стили кнопок с иконками
- `DataGridStyles.xaml` - Стили для DataGrid элементов
- `GeneralStyles.xaml` - Общие стили приложения

**Утилиты**:
- `MarshalCore` - Основная функциональность COM interop  
- `DxfGenerator` - Генератор и обработка DXF файлов
- `IPictureDispConverter` - Конвертер изображений
- `SettingsManager` - Система сохранения и загрузки настроек пользователя в JSON
- `PropertyManager` - Управление пользовательскими свойствами iProperty
- `PropertyMetadataRegistry` - Централизованная система метаданных свойств документов Inventor
- `TextWithFxIndicator` - Компонент с FX индикатором для текстовых полей
- `TokenService` - Централизованная обработка токенов для генерации имен файлов
- `TemplatePresetManager` - Управление пресетами шаблонов имен файлов с сохранением по индексу

### Управление версиями
Проект использует версионирование на основе Git в процессе сборки:
- Формат версии: `1.2.0.{НомерРевизии}`
- Хеш Git коммита включен в InformationalVersion
- Автоматическое увеличение версии на основе количества Git коммитов

### Целевая платформа
- Среда выполнения: `win-x64`
- Фреймворк: `net8.0-windows`
- Тип отладочной информации: embedded для конфигураций Debug и Release

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

**Компоненты системы:**
- `_documentCache` - основной кеш PartDocument по партнамберу
- `_partNumberToFullFileName` - кеш путей к файлам
- `AddDocumentToCache()` - добавление документа в кеш во время обхода
- `GetCachedPartDocument()` - получение документа из кеша
- `ClearDocumentCache()` - очистка кеша при повторном сканировании

**Точки заполнения кеша:**
- `ProcessComponentOccurrences` - при методе обработки "Перебор"
- `ProcessBOMRowSimple` - при методе обработки "Спецификация" 
- `ScanDocumentAsync` - для одиночных деталей
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

## Заметки по разработке

- Приложение требует установки Autodesk Inventor для функциональности COM interop
- Использует встроенную отладочную информацию для развертывания
- Включает пользовательские иконки и изображения ресурсов
- Поддерживает работу с различными форматами файлов через Inventor API
- Архитектура UI соответствует принципам WPF Data Binding
- Для новых RadioButton групп используется enum + `EnumToBooleanConverter`
- **НЕ добавлять комментарии об изменениях в код** - использовать только чистый код без комментариев о внесенных правках
- **НЕ детализировать содержимое папки HELP** - это временные справочные файлы, упоминать только общее назначение

## Стиль кодирования

При внесении изменений в проект придерживайтесь современного C# 12 стиля (.NET 8.0):

### Основные принципы
- Используйте последние возможности языка для улучшения читаемости и производительности
- Предпочитайте краткость без потери ясности
- Следуйте существующим паттернам в кодовой базе

### Инициализация и объявления

**Коллекции:**
```csharp
// Collection expressions (C# 12)
List<string> items = ["item1", "item2", "item3"];
int[] numbers = [1, 2, 3, 4, 5];

// Target-typed new
Dictionary<string, string> colorMap = new()
{
    ["Red"] = "#FF0000",
    ["Blue"] = "#0000FF"
};

// Spread operator (C# 12)
int[] combined = [..array1, ..array2, 100];
```

**Primary constructors (C# 12):**
```csharp
public class LayerSetting(string name, string value)
{
    public string Name { get; set; } = name;
    public string Value { get; set; } = value;
}
```

### Строки и форматирование

```csharp
// Используйте "" вместо string.Empty
public string Name { get; set; } = "";

// String interpolation с форматированием
decimal price = 123.45m;
string formatted = $"Price: {price:C2}";

// Raw string literals для многострочного текста
string query = """
    SELECT * FROM Users
    WHERE Status = 'Active'
    """;
```

### Pattern matching и выражения

```csharp
// Switch expressions
string GetStatusText(Status status) => status switch
{
    Status.Active => "Active",
    Status.Pending => "Pending",
    Status.Deleted => "Deleted",
    _ => "Unknown"
};

// Pattern matching с guards
var result = obj switch
{
    string { Length: > 10 } s => $"Long string: {s}",
    string s => $"Short string: {s}",
    int n when n > 0 => $"Positive number: {n}",
    _ => "Other"
};

// List patterns (C# 11+)
var description = numbers switch
{
    [] => "Empty",
    [var single] => $"One element: {single}",
    [var first, var second] => $"Two elements: {first}, {second}",
    [var first, .., var last] => $"Multiple elements from {first} to {last}"
};
```

### Nullable reference types

```csharp
// Включено по умолчанию в .NET 8.0
public string? OptionalProperty { get; set; }
public string RequiredProperty { get; set; } = "";

// Null-coalescing и null-conditional
string result = value ?? "";
int? length = text?.Length;
string name = user?.Name ?? "Unknown";
```

### Асинхронное программирование

```csharp
// Async/await с cancellation tokens
public async Task<string> LoadDataAsync(CancellationToken ct = default)
{
    await Task.Delay(100, ct);
    return "Data";
}

// Async enumerable
public async IAsyncEnumerable<int> GenerateNumbersAsync()
{
    for (int i = 0; i < 10; i++)
    {
        await Task.Delay(100);
        yield return i;
    }
}
```

### Records и immutability

```csharp
// Record для DTO и immutable данных
public record LayerConfig(string Name, string Value, bool IsEnabled);

// Record с init-only свойствами
public record ExportSettings
{
    public required string OutputPath { get; init; }
    public bool IncludeMetadata { get; init; } = true;
}
```

### File-scoped namespaces и using

```csharp
namespace FlatPatternExporter;

using System.Collections.ObjectModel;

// Весь код файла находится в namespace FlatPatternExporter
public class ExampleClass
{
    // ...
}
```

### LINQ и функциональный стиль

```csharp
// Предпочитайте LINQ методы для работы с коллекциями
var filtered = items
    .Where(x => x.IsActive)
    .OrderBy(x => x.Name)
    .Select(x => new { x.Id, x.Name })
    .ToList();

// Используйте Any(), All(), FirstOrDefault() вместо циклов
bool hasActive = items.Any(x => x.IsActive);
var first = items.FirstOrDefault(x => x.Id == targetId);
```