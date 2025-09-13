# Предлагаемая структура проекта FlatPatternExporter

На основе анализа кодовой базы и зависимостей между компонентами.

## 🗂️ **Новая файловая структура:**

```
FlatPatternExporter/
├── App.xaml(.cs)                       # Точка входа (остается в корне)
├── AssemblyInfo.cs                     # Сборка (остается в корне)
├── FPExport.ico                        # Ресурсы (остается в корне)
│
├── Libraries/                          # 📚 Внешние библиотеки для переноса
│   ├── DxfGenerator.cs                 # Будущая отдельная библиотека DXF
│   └── MarshalCore.cs                  # Отдельная библиотека COM interop
│
├── Core/                               # 🎯 Ядро системы и бизнес-логика
│   ├── PropertyMetadataRegistry.cs     # Центральный реестр метаданных (7 зависимостей)
│   ├── PropertyManager.cs              # Менеджер свойств Inventor API
│   ├── TokenService.cs                 # Обработка токенов имен файлов
│   ├── SettingsManager.cs              # Персистентность настроек
│   ├── TemplatePresetManager.cs        # Управление пресетами шаблонов
│   ├── CommonEnums.cs                  # Базовые enum и маппинги
│   └── LayerSettingsClasses.cs         # Модели настроек слоев + хелперы
│
├── UI/                                 # 🖼️ Пользовательский интерфейс
│   ├── Windows/                        # Окна приложения
│   │   ├── FlatPatternExporterMainWindow.xaml(.cs,.part2.cs)
│   │   ├── AboutWindow.xaml(.cs)
│   │   ├── ConflictDetailsWindow.xaml(.cs)
│   │   └── SelectIPropertyWindow.xaml(.cs)
│   │
│   └── Controls/                       # Пользовательские элементы
│       ├── LayerSettingControl.xaml(.cs)
│       └── TemplatePresetManagerControl.xaml(.cs)
│
├── Converters/                         # 🔄 WPF конвертеры (остается как есть)
│   ├── ColorToBrushConverter.cs
│   ├── DynamicPropertyValueConverter.cs
│   ├── EnumToBooleanConverter.cs
│   ├── IPictureDispConverter.cs
│   ├── LineTypeToGeometryConverter.cs
│   └── PropertyExpressionByNameConverter.cs
│
├── Styles/                             # 🎨 XAML стили (остается как есть)
│   ├── ColorResources.xaml
│   ├── IconResources.xaml
│   ├── ButtonStyles.xaml
│   ├── DataGridStyles.xaml
│   └── GeneralStyles.xaml
│
└── Properties/                         # Свойства проекта
    ├── launchSettings.json
    └── PublishProfiles/
```

## 🎯 **Принципы организации:**

### **Libraries/** - Подготовка к выделению
- **DxfGenerator.cs** - готовится к выделению как отдельная библиотека (namespace: `DxfGenerator`)
- **MarshalCore.cs** - готовится к выделению как COM interop библиотека (namespace: `DefineEdge`)
- Оба файла будут легко перенести в отдельные проекты

### **Core/** - Центральные компоненты и бизнес-логика
- **PropertyMetadataRegistry.cs** - самый используемый компонент (7 зависимостей)
- **PropertyManager.cs** - работа с Inventor API, использует реестр
- **TokenService.cs** - специализированная обработка токенов (7 зависимостей)
- **SettingsManager.cs** - персистентность настроек
- **TemplatePresetManager.cs** - управление пресетами шаблонов
- **CommonEnums.cs** - базовые перечисления и маппинги
- **LayerSettingsClasses.cs** - комплексное решение (модели + хелперы + валидаторы)

### **UI/** - Интерфейс пользователя
- **Windows/** - основные окна приложения
- **Controls/** - пользовательские элементы управления
- Четкое разделение окон и элементов управления

## ✅ **Преимущества структуры:**

1. **Соответствует реальным зависимостям** - центральные компоненты в Core/
2. **Подготовка к рефакторингу** - Libraries/ готов к выделению
3. **Minimal disruption** - не ломает текущие зависимости
4. **Упрощенная навигация** - логическая группировка файлов
5. **Совместимость с namespace** - основные файлы остаются в `FlatPatternExporter`
6. **Масштабируемость** - легко добавлять новые компоненты

## 🚀 **Рекомендации по внедрению:**

### **Этап 1** - Подготовка:
1. Создать папки новой структуры

### **Этап 2** - Поэтапное перемещение:
1. **Libraries/** - подготовить к выделению MarshalCore и DxfGenerator
2. **Core/** - перенести центральные компоненты
3. **UI/** - интерфейсные компоненты

### **Этап 3** - Валидация:
1. Компиляция после каждого этапа
2. Обновление using директив
3. Проверка работоспособности

Эта структура основана на **реальном анализе кода и зависимостей** между компонентами проекта.