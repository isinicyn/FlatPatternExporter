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