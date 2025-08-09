# Excel Automation MVP

Автоматизация работы с большими Excel-файлами под Windows через запуск C# консольной утилиты  
(интерфейс предназначен для интеграции с n8n или другим оркестратором через запуск командной строки).

## Назначение

Утилита выполняет типовые тяжёлые операции с Excel-файлами, используя Microsoft Excel Interop:
- Поиск и обновление ссылок/значений ячеек
- Поиск ошибок в столбцах
- Копирование столбцов между файлами
- Форсированный пересчёт формул

## Сценарий запуска

Утилита вызывается через командную строку с указанием параметров действия и файла параметров. Результат работы сохраняется в отдельный файл и/или выводится в stdout.  
Рекомендуемый способ использования — с ноды "Execute Command" в n8n.

### Пример вызова
```bat
ExcelJobRunner.exe action=updateLinks params=C:\tmp\params.json result=C:\tmp\result.json
```

## Поддерживаемые действия (`action`)
1. `updateLinks` — обновление ссылок/значений в ячейках
2. `findErrors` — поиск ошибок в заданных столбцах листа
3. `copyColumns` — копирование столбцов (или ячеек) между файлами
4. `recalculate` — принудительный пересчёт всех формул файла

## Формат входных параметров (`params.json`)

### 1. `updateLinks`
```json
{
  "inputFile": "C:\\files\\example1.xlsx",
  "cells": [
    {"address": "Sheet1!B2", "newValue": "C:\\other\\report.xlsx"},
    {"address": "Sheet1!B3", "newValue": "C:\\other\\plan.xlsx"}
  ],
  "outputFile": "C:\\files\\example1.updated.xlsx"
}
```

### 2. `findErrors`
```json
{
  "inputFile": "C:\\files\\bigfile.xlsb",
  "errors": [ "#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!" ],
  "checks": [
    { "sheet": "Данные", "column": "D" },
    { "sheet": "Данные", "column": "E" }
  ],
  "resultFile": "C:\\files\\errors.json"
}
```

### 3. `copyColumns`
```json
{
  "sourceFile": "C:\\files\\donor.xlsx",
  "targetFile": "C:\\files\\acceptor.xlsx",
  "mappings": [
    {
      "source": { "sheet": "Sheet1", "column": "E", "startRow": 2 },
      "target": { "sheet": "Sheet1", "column": "K", "mode": "append" }
    }
  ]
}
```

### 4. `recalculate`
```json
{
  "inputFile": "C:\\files\\calculations.xlsm",
  "outputFile": "C:\\files\\calculations_recalculated.xlsm"
}
```

## Формат результата (`result`)

- По умолчанию утилита завершает работу с кодом 0 (успех), либо ненулевым кодом (ошибка).
- В файле результата (resultFile/resultPath) можно вернуть:
  - Статус (`OK`, `Fail`)
  - Список изменённых ячеек, ошибок, путь до результата/лога, текст ошибки при сбое.

#### Пример файла результата
```json
{
  "status": "OK",
  "details": { /* спец. полe для каждого типа операции */ }
}
```

## Требования

- Windows с установленным Microsoft Excel
- .NET 6.0+ (для self-contained можно без .NET Framework, просто exe)
- Права на работу с файловой системой по указанным путям

## Интеграция с n8n

1. Используйте встроенный node "Execute Command":
   - Запускайте exe-файл с нужными параметрами.
   - Генерируйте файл params.json на предыдущих шагах workflow.
2. Обрабатывайте результат — заберите файлы/JSON или логи на следующих шагах n8n.

## Важно

- Утилита умеет работать с файлами `.xlsx`, `.xlsb`, `.xlsm`, не удаляет макросы.
- Во время обработки Excel часто запускается в фоне — убедитесь, что нет блокировок по процессам Excel.
- Для масштабной работы желательно запускать задачи по одной — Excel Interop не поддерживает устойчиво параллельные операции!

## Разработка и тесты

- Все тесты пишутся в файлах с именованием `.spec.`
- unit-тесты — для логики забора и подготовки параметров
- интеграционные тесты с моковыми или реальными файлами — в каталоге `/test`
- README.md поддерживается в актуальном состоянии

---
