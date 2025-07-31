# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Важно:** Указывай только основные сведения о проекте без описаний того, что добавлено, удалено или улучшено.

## Структура репозитория

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
- `EditIPropertyDialog` - Диалог редактирования свойств
- `SelectIPropertyWindow` - Окно выбора свойств с интегрированным функционалом добавления пользовательских свойств
- `OverrideQuantityDialog` - Диалог переопределения количества

**Компоненты LayerSettings** (интегрированы):
- `LayerSettingControl` - Пользовательский элемент управления для конфигурации слоев
- `LayerSettingsClasses` - Классы моделей данных, конвертеры и вспомогательные методы
- `LayerSettingsHelper` - Статические методы для работы с настройками слоев

**Утилиты**:
- `MarshalCore` - Основная функциональность COM interop  
- `DXFThumbnailCreator` - Генерация миниатюр DXF файлов
- `IPictureDispConverter` - Конвертер изображений

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

### Управление пользовательскими свойствами (Custom iProperties)
- Добавление пользовательских свойств через текстовое поле в окне выбора свойств
- Автоматическое заполнение данных из "Inventor User Defined Properties"
- Поддержка фонового режима обработки

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