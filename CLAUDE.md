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
    ├── TextWithFxIndicator.xaml(.cs)   # Компонент с FX индикатором
    │
    ├── Styles/                         # XAML стили
    │   ├── ColorResources.xaml
    │   ├── IconResources.xaml
    │   ├── DataGridStyles.xaml
    │   └── GeneralStyles.xaml
    │
    ├── MarshalCore.cs                  # COM interop
    ├── DxfGenerator.cs                 # Генератор DXF
    ├── IPictureDispConverter.cs        # Конвертер изображений
    ├── PropertyManager.cs              # Управление свойствами
    ├── SettingsManager.cs              # Настройки пользователя
    ├── LayerSettingsClasses.cs         # Классы настроек слоев
    ├── TokenService.cs                 # Сервис работы с токенами
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
- `DataGridStyles.xaml` - Стили для DataGrid элементов
- `GeneralStyles.xaml` - Общие стили приложения

**Утилиты**:
- `MarshalCore` - Основная функциональность COM interop  
- `DxfGenerator` - Генератор и обработка DXF файлов
- `IPictureDispConverter` - Конвертер изображений
- `SettingsManager` - Система сохранения и загрузки настроек пользователя в XML
- `PropertyManager` - Управление пользовательскими свойствами iProperty
- `TextWithFxIndicator` - Компонент с FX индикатором для текстовых полей
- `TokenService` - Централизованная обработка токенов для генерации имен файлов

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

### Система токенов для имен файлов
Конструктор имени файла с интерактивным интерфейсом:
- Токенизированный ввод с drag&drop интерфейсом
- Real-time предпросмотр результата на основе реальных данных
- Валидация шаблона с визуальной и текстовой индикацией
- Централизованная обработка через TokenService с кэшированием
- Автоматическая санитизация недопустимых символов имени файла

### Управление пользовательскими свойствами (Custom iProperties)
- Добавление пользовательских свойств через текстовое поле в окне выбора свойств
- Автоматическое заполнение данных из "Inventor User Defined Properties"
- Поддержка фонового режима обработки

### Система настроек пользователя
- Автоматическое сохранение настроек в XML файл при закрытии приложения
- Восстановление пользовательских колонок и параметров при запуске
- Расположение файла настроек: `%APPDATA%\FlatPatternExporter\settings.xml`
- Сохранение всех свойств iProperty без зависимости от видимых колонок UI

### Архитектура UI - Система привязки данных
Проект использует декларативную архитектуру WPF/MVVM:

**Enum-based система для RadioButton групп:**
- `ExportFolderType` - управление выбором папки экспорта (5 опций)
- `ProcessingMethod` - выбор метода обработки (Перебор/Спецификация)

**Универсальные конвертеры:**
- `EnumToBooleanConverter` - универсальный конвертер для привязки enum к RadioButton
- Поддерживает любые enum через рефлексию и `Enum.TryParse`

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