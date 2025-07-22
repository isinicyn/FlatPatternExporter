# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Структура репозитория

Это C# WPF приложение, которое является дополнением для Autodesk Inventor:

- **FPECORE** - Приложение для экспорта свойств файлов с интегрированным функционалом управления слоями

Приложение построено с использованием:
- .NET 8.0 Windows Desktop
- WPF (Windows Presentation Foundation) 
- Поддержка Windows Forms включена
- Интеграция с Autodesk Inventor Interop API

## Команды сборки

### Сборка проекта
```bash
# Собрать решение
dotnet build FPECORE.sln --configuration Release

# Собрать конкретную конфигурацию для x64
dotnet build FPECORE.sln --configuration Release --arch x64
```

### Запуск приложения
```bash
# Запустить в режиме отладки
dotnet run --project FPECORE\FPECore.csproj
```

## Архитектура

### Зависимости проекта
- Autodesk Inventor Interop: `C:\Program Files\Autodesk\Inventor 2026\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll`
- netDxf.netstandard (v3.0.1) для работы с DXF файлами
- stdole (v17.14.40260) для COM interop

### Ключевые компоненты

**Основные окна**:
- `FPECMainWindow` - Главное окно приложения со встроенным функционалом LayerSettings
- `ConflictDetailsWindow` - Окно деталей конфликтов
- `AboutWindow` - Окно "О программе"

**Диалоги управления свойствами**:
- `CustomIPropertyDialog` - Диалог пользовательских свойств
- `EditIPropertyDialog` - Диалог редактирования свойств
- `SelectIPropertyWindow` - Окно выбора свойств
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

### Особенности интеграции LayerSettings
Функционал управления слоями полностью интегрирован в основное приложение:
- UI компоненты встроены в главное окно FPECORE
- Использует единую архитектуру namespace `FPECORE`
- Поддерживает валидацию имен слоев с фильтрацией недопустимых символов
- Генерирует строки экспорта для DXF с настраиваемыми параметрами слоев

## Заметки по разработке

- Приложение требует установки Autodesk Inventor для функциональности COM interop
- Использует встроенную отладочную информацию для развертывания
- Включает пользовательские иконки и изображения ресурсов
- Поддерживает работу с различными форматами файлов через Inventor API