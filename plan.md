# План рефакторинга проекта FlatPatternExporter

## 1. Анализ текущего состояния

После анализа файлов 'FlatPatternExporterMainWindow.xaml.cs' и 'FlatPatternExporterMainWindow.xaml.part2.cs' выявлены следующие блоки бизнес-логики, смешанные с UI-кодом:

### Основные функциональные блоки:
1. **Работа с Inventor API** - подключение, валидация документов, получение свойств
2. **Сканирование документов** - обход структуры сборок, обработка BOM, анализ листового металла
3. **Управление кешем документов** - кеширование PartDocument для оптимизации
4. **Обработка конфликтов** - анализ дублирующихся обозначений деталей
5. **Экспорт DXF** - генерация файлов, оптимизация, работа со слоями
6. **Генерация миниатюр** - создание превью из DXF и документов
7. **Управление свойствами деталей** - чтение iProperties, обработка выражений
8. **Управление состоянием UI** - централизованное управление состоянием интерфейса

## 2. Новая структура проекта

```
FlatPatternExporter/
├── Core/                               # Бизнес-логика
│   ├── InventorService.cs            # Работа с Inventor API
│   ├── ScanService.cs                 # Сканирование документов и сборок
│   ├── DocumentCacheService.cs       # Управление кешем документов
│   ├── ConflictAnalyzer.cs          # Анализ конфликтов обозначений
│   ├── ExportService.cs              # Экспорт в DXF
│   ├── ThumbnailService.cs           # Генерация миниатюр
│   └── PartDataService.cs            # Работа со свойствами деталей
│
└── UI/
    └── Windows/
        ├── FlatPatternExporterMainWindow.xaml
        └── FlatPatternExporterMainWindow.xaml.cs  # Только UI логика

```

## 3. Детальное описание новых компонентов

### Core/InventorService.cs
**Ответственность**: Подключение к Inventor, валидация документов, управление проектами
**Методы из исходных файлов**:
- `EnsureInventorConnection()`
- `InitializeInventor()`
- `InitializeProjectData()`
- `ValidateActiveDocument()`
- `SetInventorUserInterfaceState()`
- `OpenInventorDocument()`
- `OpenPartDocument()`
- `GetPartDocumentFullPath()`
- `SetProjectFolderInfo()`
- `IsLibraryComponent()`

### Core/ScanService.cs
**Ответственность**: Сканирование структуры документов, обход сборок
**Методы из исходных файлов**:
- `ScanDocumentAsync()`
- `ProcessComponentOccurrences()`
- `ProcessBOM()`
- `GetAllBOMRowsRecursively()`
- `ProcessBOMRowSimple()`
- `ProcessSinglePart()`
- `ShouldExcludeComponent()`
- `GetFullFileName()`
- `FilterConflictingParts()`

### Core/DocumentCacheService.cs
**Ответственность**: Кеширование открытых документов для оптимизации
**Методы и поля из исходных файлов**:
- `_documentCache`
- `_partNumberToFullFileName`
- `AddDocumentToCache()`
- `GetCachedPartDocument()`
- `ClearDocumentCache()`

### Core/ConflictAnalyzer.cs
**Ответственность**: Анализ конфликтов обозначений деталей
**Методы и поля из исходных файлов**:
- `_partNumberTracker`
- `_conflictingParts`
- `_conflictFileDetails`
- `AddPartToConflictTracker()`
- `AnalyzePartNumberConflictsAsync()`
- `ClearConflictData()`
- Класс `PartConflictInfo`

### Core/ExportService.cs
**Ответственность**: Экспорт разверток в DXF
**Методы из исходных файлов**:
- `ExportDXF()`
- `PrepareExportOptions()`
- `PrepareForExport()`
- `PrepareExportContextAsync()`
- `ExportWithoutScan()`
- `IsFileLocked()`
- `IsValidPath()`
- Класс `ExportContext`

### Core/ThumbnailService.cs
**Ответственность**: Генерация миниатюр из DXF и документов
**Методы из исходных файлов**:
- `GenerateDxfThumbnails()`
- `GetThumbnailAsync()`
- `ConvertSvgToBitmapImage()`

### Core/PartDataService.cs
**Ответственность**: Работа со свойствами деталей, чтение iProperties
**Методы из исходных файлов**:
- `GetPartDataAsync()` (обе перегрузки)
- `ReadAllPropertiesFromPart()`
- `SetExpressionStatesForAllProperties()`
- `FillPropertyData()`
- `UpdateQuantitiesWithMultiplier()`
- `UpdateDocumentInfo()`

## 4. План пошагового рефакторинга

### Этап 1: Создание InventorService
1. Создать файл `Core/InventorService.cs`
2. Перенести методы работы с Inventor API
3. Удалить перенесённые методы из исходных файлов
4. Обновить зависимости в MainWindow
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 2: Создание ScanService
1. Создать файл `Core/ScanService.cs`
2. Перенести методы сканирования документов
3. Удалить перенесённые методы из исходных файлов
4. Обновить зависимости
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 3: Создание DocumentCacheService
1. Создать файл `Core/DocumentCacheService.cs`
2. Перенести логику кеширования
3. Удалить из исходных файлов
4. Обновить зависимости
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 4: Создание ConflictAnalyzer
1. Создать файл `Core/ConflictAnalyzer.cs`
2. Перенести логику анализа конфликтов
3. Удалить из исходных файлов
4. Обновить зависимости
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 5: Создание ExportService
1. Создать файл `Core/ExportService.cs`
2. Перенести методы экспорта
3. Удалить из исходных файлов
4. Обновить зависимости
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 6: Создание ThumbnailService
1. Создать файл `Core/ThumbnailService.cs`
2. Перенести методы генерации миниатюр
3. Удалить из исходных файлов
4. Обновить зависимости
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 7: Создание PartDataService
1. Создать файл `Core/PartDataService.cs`
2. Перенести методы работы со свойствами
3. Удалить из исходных файлов
4. Обновить зависимости
5. **Фиксация**: Сборка проекта, проверка компиляции

### Этап 8: Финальная очистка
1. Объединить содержимое `.part2.cs` с основным файлом `.cs` (оставив только UI-логику)
2. Удалить файл `.part2.cs`
3. **Фиксация**: Финальная сборка и проверка

## 5. Важные замечания

- Классы `UIState`, `OperationResult`, `DocumentValidationResult`, `ScanProgress` должны остаться в MainWindow или быть перемещены в отдельную папку Models
- Класс `PartData` должен остаться в Models
- Обработчики событий UI остаются в MainWindow
- Сервисы будут инжектироваться в MainWindow через конструктор или как поля класса
- Все сервисы должны использовать существующий namespace `FlatPatternExporter.Core`

## 6. Статус выполнения

- [x] Этап 1: InventorService - **Завершён успешно**
  - Создан файл Core/InventorService.cs
  - Перенесены методы работы с Inventor API
  - Удалены методы из исходных файлов
  - Обновлены зависимости в MainWindow
  - Проект успешно скомпилирован
- [x] Этап 2: ScanService - **Завершён успешно**
  - Создан файл Core/ScanService.cs с методами сканирования документов
  - Создан файл Core/DocumentCacheService.cs для управления кешем
  - Создан файл Core/ConflictAnalyzer.cs для анализа конфликтов
  - Обновлены все зависимости в MainWindow и part2.cs
  - Проект успешно скомпилирован
- [x] Этап 3: DocumentCacheService - **Завершён успешно**
  - Создан файл Core/DocumentCacheService.cs
  - Перенесена логика кеширования документов
  - Удалены дублирующиеся методы из исходных файлов
  - Обновлены все ссылки на использование сервиса
  - Проект успешно скомпилирован
- [x] Этап 4: ConflictAnalyzer - **Завершён успешно**
  - Создан файл Core/ConflictAnalyzer.cs
  - Перенесена логика анализа конфликтов
  - Удалены дублирующиеся методы и неиспользуемые свойства
  - Обновлены все ссылки на использование сервиса
  - Проект успешно скомпилирован
- [x] Этап 5: ExportService - **Завершён успешно**
  - Создан файл Core/ExportService.cs
  - Перенесены методы ExportDXF, PrepareExportContextAsync, PrepareForExport
  - Перенесены вспомогательные методы PrepareExportOptions, IsFileLocked, IsValidPath
  - Перенесены классы ExportContext и ExportOptions
  - Обновлены зависимости в MainWindow
  - Удалены перенесенные методы из исходных файлов
  - Проект успешно скомпилирован
- [x] Этап 6: ThumbnailService - **Завершён успешно**
  - Создан файл Core/ThumbnailService.cs
  - Перенесены методы генерации миниатюр из MainWindow
  - Обновлены все зависимости
  - Проект успешно скомпилирован
- [x] Этап 7: PartDataService - **Завершён успешно**
  - Создан файл Core/PartDataService.cs
  - Перенесены методы GetPartDataAsync, ReadAllPropertiesFromPart, SetExpressionStatesForAllProperties
  - Перенесены методы FillPropertyData, UpdateQuantitiesWithMultiplier, UpdateDocumentInfo
  - Добавлен новый класс DocumentInfo для передачи информации о документе
  - Обновлены все зависимости в MainWindow и part2.cs
  - Удалены перенесенные методы из исходных файлов
  - Проект успешно скомпилирован
- [x] Этап 8: Финальная очистка - **Завершён успешно**
  - Все методы из part2.cs перенесены в основной файл MainWindow.xaml.cs
  - Добавлены необходимые using директивы (System.IO, System.Windows.Forms)
  - Исправлен конфликт имен File между Inventor.File и System.IO.File
  - Файл FlatPatternExporterMainWindow.xaml.part2.cs удален
  - Проект успешно скомпилирован без ошибок