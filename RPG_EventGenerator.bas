Option Explicit

' ========================================
' RPG Event Generator - Генератор событий для настольной RPG
' Автор: [Yuryi] + Claude Sonnet 4
' Дата создания: 15.07.2025
' Дата последнего обновления: 24.08.2025
' Интеграция: Уникальные события + Универсальные внутренние броски
' Требует модуль: ValidationLogger.bas
' ========================================

' ---------- Типы данных ----------
Private Type PathInfo
    Name As String
    ColumnIndex As Long
End Type

Private Type ClassConfig
    className As String        ' Имя класса (Маг, Ведьмак)
    DataSheetName As String    ' Имя листа с данными
    EventSuffix As String      ' Суффикс для таблиц событий (_маг, _ведьмак)
    PrefixOpasnost As String   ' Префикс для диапазона опасностей
    PrefixPorog As String      ' Префикс для диапазона порогов
    PrefixPath As String       ' Префикс для диапазона путей
    Paths() As PathInfo        ' Массив путей для класса
    IsFullyImplemented As Boolean  ' Флаг готовности класса
End Type

Private Type GeneratorState
    Config As ClassConfig
    IsInitialized As Boolean
    IsValidated As Boolean     ' Флаг успешной валидации
End Type

Private Type EventResult
    BaseEvent As String          ' Базовое событие
    DetailedDescription As String ' Полное описание с результатами бросков
    SubRolls As String           ' Информация о дополнительных бросках
    IsUnique As Boolean          ' Флаг уникальности события
    RequiresReroll As Boolean    ' Требует переброса
End Type

Private Type DiceCommand
    CommandType As String        ' Тип команды: ROLL, IF, SELECT
    DiceExpression As String     ' Выражение броска: 1d10, 2d6+3
    Condition As String          ' Условие: "EVEN", "ODD", ">=5"
    Options As Collection        ' Варианты для SELECT (ключ-значение)
    ResultTemplate As String     ' Шаблон результата с плейсхолдерами
End Type

Private Type RollResult
    RollValue As Long           ' Результат броска
    ProcessedText As String     ' Обработанный текст
    Success As Boolean          ' Успешность обработки
End Type

' ---------- Глобальные переменные (исправлено для совместимости с VBA) ----------
Private State As GeneratorState

' История персонажа (используем отдельные переменные вместо Private Type)
Private UsedUniqueEvents As Collection
Private TotalGenerations As Long

' ---------- Константы настроек ----------
Private Const UI_SHEET As String = "Генератор"
Private Const HISTORY_SHEET As String = "История"

Private Const UI_CLASS_CELL As String = "B2"
Private Const UI_PATH_CELL As String = "B23"
Private Const UI_A25 As String = "A25"
Private Const UI_A26 As String = "A26"
Private Const UI_B24 As String = "B24"
Private Const UI_B25 As String = "B25"
Private Const UI_B26 As String = "B26"

' ========================================
' ПУБЛИЧНЫЕ ФУНКЦИИ
' ========================================

' Главная процедура инициализации и валидации
Public Sub Main()
    On Error GoTo ErrHandler
    
    ' Инициализация систем
    Randomize
    ValidationLogger.InitializeValidation
    InitializeCharacterHistory
    
    ' Выполняем валидацию и показываем результаты
    If Not ValidateAndShowResults() Then Exit Sub
    
    ' Генерируем первое событие с расширенной логикой
    GenerateEnhancedEvent_SaveHistory
    
    Exit Sub

ErrHandler:
    ValidationLogger.AddValidationError "Main", "КРИТИЧНО", _
                                       "Ошибка " & Err.Number & ": " & Err.Description, _
                                       "Строка: " & Erl
    ValidationLogger.ShowValidationResults
    ValidationLogger.ClearValidation
End Sub

' Процедура генерации события с поддержкой уникальных событий и внутренних бросков
Public Sub GenerateEnhancedEvent_SaveHistory()
    On Error GoTo ErrHandler
    
    ' Объявление переменных
    Dim wsUI As Worksheet, wsHist As Worksheet
    Dim colIndex As Long, path As String, className As String
    Dim danger As Boolean, variantDanger As Variant
    Dim roll100 As Long, Porog As Variant
    
    ' Проверка UI листа
    Set wsUI = ThisWorkbook.Worksheets(UI_SHEET)
    
    ' Чтение входных данных (без полной валидации)
    If Not ReadInputData(wsUI, className, path, colIndex, False) Then
        MsgBox "Ошибка чтения входных данных. Проверьте выбор класса и пути.", vbExclamation
        Exit Sub
    End If
    
    ' Загрузка истории персонажа для отслеживания уникальных событий
    LoadCharacterHistoryFromSheet
    
    ' Генерация опасности
    GenerateDanger wsUI, path, colIndex, danger, variantDanger, roll100, Porog
    
    ' Генерация события с расширенной логикой
    Dim enhancedEvent As EventResult
    enhancedEvent = GenerateUniversalEnhancedEvent(colIndex)
    
    ' Записываем результаты в UI
    WriteEnhancedResultsToUI wsUI, enhancedEvent
    
    ' Сохранение в историю
    Set wsHist = GetOrCreateHistorySheet()
    SaveEnhancedToHistory wsHist, path, roll100, Porog, danger, variantDanger, enhancedEvent
    
    ' Сигнализация об успешной генерации
    Dim message As String
    message = "Генерация завершена. Результат записан в '" & HISTORY_SHEET & "'."
    If enhancedEvent.SubRolls <> "" Then
        message = message & vbCrLf & "Дополнительные броски: " & enhancedEvent.SubRolls
    End If
    If enhancedEvent.IsUnique Then
        message = message & vbCrLf & "Событие отмечено как уникальное."
    End If
    
    MsgBox message, vbInformation
    
    Exit Sub

ErrHandler:
    MsgBox "Ошибка генерации: " & Err.Description, vbCritical
End Sub

' ========================================
' ИНИЦИАЛИЗАЦИЯ ИСТОРИИ ПЕРСОНАЖА (ИСПРАВЛЕНО)
' ========================================

' Инициализация истории персонажа
 Private Sub InitializeCharacterHistory()
    Set UsedUniqueEvents = New Collection
    TotalGenerations = 0
End Sub

' Загрузка истории из листа "История"
Private Sub LoadCharacterHistoryFromSheet()
    InitializeCharacterHistory
    
    Dim wsHist As Worksheet
    Set wsHist = GetHistorySheet()
    If wsHist Is Nothing Then Exit Sub
    
    Dim lastRow As Long, i As Long
    lastRow = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row
    
    ' Читаем все события персонажа
    For i = 2 To lastRow ' Пропускаем заголовок
        Dim eventText As String
        eventText = Trim(CStr(wsHist.Cells(i, "H").Value)) ' Колонка "Тип события"
        
        ' Добавляем уникальные события в коллекцию
        If IsUniqueEvent(eventText) Then
            On Error Resume Next
            UsedUniqueEvents.Add eventText, eventText
            On Error GoTo 0
        End If
        
        TotalGenerations = TotalGenerations + 1
    Next i
End Sub

Public Function GetClassConfig(ByVal className As String) As ClassConfig
    Select Case UCase(Trim(className))
        Case "МАГ"
            With GetClassConfig
                .className = "Маг"
                .DataSheetName = "Источник_данных_Маг"
                .EventSuffix = "_маг"
                .PrefixOpasnost = "Опасность_Маг"
                .PrefixPorog = "Порог_Опасности_Маг"
                .PrefixPath = "Путь_Маг"
                .IsFullyImplemented = True
                
                ReDim .Paths(1 To 4)
                .Paths(1).Name = "Осторожность": .Paths(1).ColumnIndex = 2
                .Paths(2).Name = "Политиканство": .Paths(2).ColumnIndex = 3
                .Paths(3).Name = "Эксперименты": .Paths(3).ColumnIndex = 4
                .Paths(4).Name = "Магические исследования": .Paths(4).ColumnIndex = 5
            End With
            
        Case "ВЕДЬМАК"
            With GetClassConfig
                .className = "Ведьмак"
                .DataSheetName = "Источник_данных_Ведьмак"
                .EventSuffix = "_ведьмак"
                .PrefixOpasnost = "Опасность_Ведьмак"
                .PrefixPorog = "Порог_Опасности_Ведьмак"
                .PrefixPath = "Путь_Ведьмак"
                .IsFullyImplemented = False ' TODO: Завершить реализацию класса
                
                ReDim .Paths(1 To 4)
                ' TODO: Добавить пути для Ведьмака
            End With
            
        Case Else
            ValidationLogger.AddValidationError "GetClassConfig", "КРИТИЧНО", _
                                               "Неизвестный класс: " & className
    End Select
End Function

' Получение листа истории (без создания)
Private Function GetHistorySheet() As Worksheet
    On Error Resume Next
    Set GetHistorySheet = ThisWorkbook.Worksheets(HISTORY_SHEET)
    On Error GoTo 0
End Function

' ========================================
' ФУНКЦИИ КОНФИГУРАЦИИ
' ========================================

' Инициализация состояния генератора
Private Sub InitializeState(ByVal className As String)
    If Not State.IsInitialized Or State.Config.className <> className Then
        State.Config = GetClassConfig(className)
        State.IsInitialized = True
        State.IsValidated = False  ' Требуется новая валидация
    End If
End Sub

' ========================================
' ФУНКЦИИ ВАЛИДАЦИИ И ИНИЦИАЛИЗАЦИИ
' ========================================

' Валидация с показом результатов
Private Function ValidateAndShowResults() As Boolean
    ValidateAndShowResults = PerformInitialValidation()
    
    ' Если есть критические ошибки
    If Not ValidateAndShowResults Then
        ValidationLogger.ShowValidationResults
        ValidationLogger.ClearValidation
        Exit Function
    End If
    
    ' Если есть предупреждения (но нет критических ошибок), показываем их
    If ValidationLogger.HasErrors() Then
        ValidationLogger.ShowValidationResults
    End If
End Function

' Начальная валидация системы (вызывается один раз)
Private Function PerformInitialValidation() As Boolean
    PerformInitialValidation = False
    
    ' Объявление переменных
    Dim wsUI As Worksheet
    Dim className As String, path As String, colIndex As Long
    
    ' Проверка UI листа
    If Not ValidationLogger.ValidateWorksheet(UI_SHEET, "PerformInitialValidation") Then
        Exit Function
    End If
    
    Set wsUI = ThisWorkbook.Worksheets(UI_SHEET)
    
    ' Чтение и валидация входных данных
    If Not ReadInputData(wsUI, className, path, colIndex, True) Then
        Exit Function
    End If
    
    ' Проверяем готовность класса
    If Not State.Config.IsFullyImplemented Then
        If ValidationLogger.HasCriticalErrors() Then
            MsgBox "Класс '" & className & "' не полностью реализован. Генерация невозможна.", vbCritical
            Exit Function
        End If
    End If

        ' Проверяем критические ошибки после валидации
    If ValidationLogger.HasCriticalErrors() Then
        Exit Function
    End If

    PerformInitialValidation = True
End Function

' Чтение и валидация входных данных
Private Function ReadInputData(ByRef wsUI As Worksheet, ByRef className As String, _
                              ByRef path As String, ByRef colIndex As Long, _
                              Optional ByVal performValidation As Boolean = False) As Boolean
    ' Читаем класс
    className = Trim(CStr(wsUI.Range(UI_CLASS_CELL).Value))
    If className = "" Then
        If performValidation Then
            ValidationLogger.AddValidationError "ReadInputData", "КРИТИЧНО", _
                                               "Не выбран класс", "Выберите класс в ячейке " & UI_CLASS_CELL
        End If
        ReadInputData = False
        Exit Function
    End If

    ' Инициализируем состояние для класса (если необходимо)
    InitializeState className
    
    ' Валидируем конфигурацию класса (если требуется)
    If performValidation And Not State.IsValidated Then
        ValidateClassConfiguration
        State.IsValidated = True
    End If
    
    ' Читаем путь
    path = Trim(CStr(wsUI.Range(UI_PATH_CELL).Value))
    If path = "" Then
        If performValidation Then
            ValidationLogger.AddValidationError "ReadInputData", "КРИТИЧНО", _
                                               "Не выбран путь", "Выберите путь в ячейке " & UI_PATH_CELL
        End If
        ReadInputData = False
        Exit Function
    End If
    
    ' Получаем индекс колонки
    colIndex = GetColumnIndexForClassPath(path)
    ReadInputData = (colIndex > 0)
End Function

' Валидация конфигурации текущего класса
Private Sub ValidateClassConfiguration() As Boolean
    If State.Config.IsFullyImplemented = False Then
        ValidationLogger.AddValidationError "ValidateClassConfiguration", "КРИТИЧНО", _
                                           "Класс '" & State.Config.className & "' не полностью реализован", _
                                           "Завершите реализацию класса в ClassConfiguration.bas"
        Exit Sub
    End If
    With State.Config
        ' Создаем массив путей для передачи в ValidationLogger
        Dim pathsArray As Variant
        ReDim pathsArray(LBound(.Paths) To UBound(.Paths))
        Dim i As Long
        For i = LBound(.Paths) To UBound(.Paths)
            pathsArray(i) = .Paths(i).Name
        Next i
        
        ' Вызываем валидацию из ValidationLogger
        ValidationLogger.ValidateClassConfiguration .className, .DataSheetName, _
                                                   .PrefixOpasnost, .PrefixPorog, .PrefixPath, pathsArray
    End With
End Sub

' Получение индекса колонки для пути
Private Function GetColumnIndexForClassPath(ByVal path As String) As Long
    Dim i As Long
    
    For i = LBound(State.Config.Paths) To UBound(State.Config.Paths)
        If State.Config.Paths(i).Name = path Then
            GetColumnIndexForClassPath = State.Config.Paths(i).ColumnIndex
            Exit Function
        End If
    Next i
    
    ValidationLogger.AddValidationError "GetColumnIndexForClassPath", "КРИТИЧНО", _
                                       "Недопустимый путь для класса", _
                                       "Путь '" & path & "' не найден для класса '" & State.Config.className & "'"
    GetColumnIndexForClassPath = 0
End Function

' ========================================
' УНИВЕРСАЛЬНАЯ СИСТЕМА ГЕНЕРАЦИИ СОБЫТИЙ
' ========================================

' Обновленная функция генерации событий с интеграцией обеих систем
Private Function GenerateUniversalEnhancedEvent(ByVal colIndex As Long) As EventResult
    Dim maxAttempts As Long: maxAttempts = 50
    Dim attempts As Long: attempts = 0
    
NextAttempt:  ' Добавили метку для перехода
    attempts = attempts + 1
    
    If attempts >= maxAttempts Then GoTo FallbackEvent  ' Проверка на превышение попыток
    
    ' Генерируем базовое событие
    Dim rngPath As Range
    Set rngPath = ValidationLogger.ValidateAndCacheRange(State.Config.PrefixPath, _
                                                       "GenerateUniversalEnhancedEvent", _
                                                       State.Config.DataSheetName)
    If rngPath Is Nothing Then GoTo FallbackEvent
    
    Dim eventType As Variant
    eventType = SafeVLookup(rngPath, Roll(10), colIndex, "GenerateUniversalEnhancedEvent", "тип события")
    If IsError(eventType) Then GoTo FallbackEvent
    
    Dim baseEvent As String
    baseEvent = Trim(CStr(eventType))
    
    ' Проверяем уникальность (система из кода А)
    If IsUniqueEvent(baseEvent) Then
        If IsEventAlreadyUsed(baseEvent) Then
            GoTo NextAttempt  ' ИСПРАВЛЕНО: заменили "Continue Do" на "GoTo NextAttempt"
        End If
    End If
    
    ' Определяем имя таблицы событий для поиска метаданных
    Dim tableName As String
    tableName = Replace(baseEvent & State.Config.EventSuffix, " ", "_")
    
    ' Обрабатываем универсальными методами (система из кода Б)
    GenerateUniversalEnhancedEvent = ProcessUniversalInternalRolls(baseEvent, tableName)
    
    ' Сохраняем уникальное событие
    If GenerateUniversalEnhancedEvent.IsUnique Then
        On Error Resume Next
        UsedUniqueEvents.Add baseEvent, baseEvent
        On Error GoTo 0
    End If
    
    Exit Function

FallbackEvent:
    ' Fallback при превышении попыток или ошибках
    With GenerateUniversalEnhancedEvent
        .BaseEvent = "Спокойное десятилетие"
        .DetailedDescription = "Это десятилетие прошло без особых событий."
        .SubRolls = ""
        .IsUnique = False
        .RequiresReroll = False
    End With
End Function

' ========================================
' УНИВЕРСАЛЬНАЯ СИСТЕМА ВНУТРЕННИХ БРОСКОВ
' ========================================

' Главная функция обработки внутренних бросков
Private Function ProcessUniversalInternalRolls(ByVal baseEvent As String, _
                                             ByVal eventTableName As String) As EventResult
    
    ProcessUniversalInternalRolls.BaseEvent = baseEvent
    ProcessUniversalInternalRolls.DetailedDescription = baseEvent
    ProcessUniversalInternalRolls.SubRolls = ""
    ProcessUniversalInternalRolls.IsUnique = IsUniqueEvent(baseEvent)
    ProcessUniversalInternalRolls.RequiresReroll = False
    
    ' Получаем команды бросков из таблицы
    Dim commands As String
    commands = GetRollCommandsFromTable(eventTableName, baseEvent)
    
    If Trim(commands) = "" Then Exit Function
    
    ' Обрабатываем команды
    Dim result As RollResult
    result = ExecuteRollCommands(commands, baseEvent)
    
    If result.Success Then
        ProcessUniversalInternalRolls.DetailedDescription = result.ProcessedText
        ProcessUniversalInternalRolls.SubRolls = "Дополнительные броски выполнены"
    End If
End Function

' Получение команд бросков из таблицы событий
Private Function GetRollCommandsFromTable(ByVal tableName As String, ByVal eventText As String) As String
    GetRollCommandsFromTable = ""
    
    ' Получаем диапазон таблицы событий
    Dim rng As Range
    Set rng = ValidationLogger.ValidateAndCacheRange(tableName, _
                                                   "GetRollCommandsFromTable", _
                                                   State.Config.DataSheetName)
    If rng Is Nothing Then Exit Function
    
    ' Ищем строку с нашим событием и получаем команды из колонки D
    Dim i As Long
    For i = 1 To rng.Rows.Count
        If Trim(CStr(rng.Cells(i, 2).Value)) = Trim(eventText) Then ' Колонка B содержит название события
            If rng.Columns.Count >= 4 Then ' Проверяем наличие колонки D
                GetRollCommandsFromTable = Trim(CStr(rng.Cells(i, 4).Value)) ' Колонка D - команды
            End If
            Exit Function
        End If
    Next i
End Function

' Выполнение команд бросков
Private Function ExecuteRollCommands(ByVal commands As String, ByVal baseText As String) As RollResult
    ExecuteRollCommands.Success = False
    ExecuteRollCommands.ProcessedText = baseText
    ExecuteRollCommands.RollValue = 0
    
    On Error GoTo ErrHandler
    
    ' Разбиваем команды по разделителю "|"
    Dim commandParts As Variant
    commandParts = Split(commands, "|")
    
    Dim rollValue As Long
    Dim i As Long, processed As String
    processed = baseText
    
    For i = 0 To UBound(commandParts)
        Dim currentCommand As String
        currentCommand = Trim(CStr(commandParts(i)))
        
        If Left(currentCommand, 5) = "ROLL:" Then
            ' Выполняем бросок: ROLL:1d10
            rollValue = ExecuteSingleRoll(Mid(currentCommand, 6))
            ExecuteRollCommands.RollValue = rollValue
            
        ElseIf Left(currentCommand, 7) = "SELECT:" Then
            ' Обрабатываем выборку: SELECT:1d10|1-3:вариант1|4-6:вариант2
            processed = ProcessSelectCommand(currentCommand, commandParts, i, rollValue)
            Exit For ' SELECT завершает обработку
            
        ElseIf Left(currentCommand, 3) = "IF:" Then
            ' Обрабатываем условие: IF:EVEN:активен или IF:1-6:короткий
            Dim conditionResult As String
            conditionResult = ProcessIfCommand(currentCommand, rollValue)
            If conditionResult <> "" Then
                processed = processed & " " & conditionResult
            End If
        End If
    Next i
    
    ExecuteRollCommands.ProcessedText = processed
    ExecuteRollCommands.Success = True
    Exit Function
    
ErrHandler:
    ExecuteRollCommands.ProcessedText = baseText & " (Ошибка обработки команд)"
    ExecuteRollCommands.Success = False
End Function

' Выполнение одного броска кубика
Private Function ExecuteSingleRoll(ByVal diceExpression As String) As Long
    ExecuteSingleRoll = 0
    
    ' Парсим выражение типа "1d10", "2d6+3", "1d20-2"
    Dim expression As String
    expression = UCase(Trim(diceExpression))
    
    ' Простейший парсер для стандартных выражений
    If InStr(expression, "D") > 0 Then
        Dim parts As Variant
        parts = Split(expression, "D")
        
        If UBound(parts) >= 1 Then
            Dim diceCount As Long, diceType As Long, modifier As Long
            diceCount = CLng(parts(0))
            
            ' Обрабатываем модификаторы
            Dim rightPart As String
            rightPart = CStr(parts(1))
            
            If InStr(rightPart, "+") > 0 Then
                Dim plusParts As Variant
                plusParts = Split(rightPart, "+")
                diceType = CLng(plusParts(0))
                modifier = CLng(plusParts(1))
            ElseIf InStr(rightPart, "-") > 0 Then
                Dim minusParts As Variant
                minusParts = Split(rightPart, "-")
                diceType = CLng(minusParts(0))
                modifier = -CLng(minusParts(1))
            Else
                diceType = CLng(rightPart)
                modifier = 0
            End If
            
            ' Выполняем броски
            Dim total As Long, j As Long
            For j = 1 To diceCount
                total = total + Roll(diceType)
            Next j
            
            ExecuteSingleRoll = total + modifier
        End If
    Else
        ' Простое число
        ExecuteSingleRoll = CLng(expression)
    End If
End Function

' Обработка команды SELECT
Private Function ProcessSelectCommand(ByVal selectCmd As String, ByVal allCommands As Variant, _
                                    ByVal startIndex As Long, ByVal rollValue As Long) As String
    ProcessSelectCommand = ""
    
    Dim actualRoll As Long
    If rollValue = 0 Then
        ' Если не было предыдущего броска, выполняем его
        actualRoll = ExecuteSingleRoll(Mid(selectCmd, 8)) ' Убираем "SELECT:"
    Else
        actualRoll = rollValue
    End If
    
    ' Проходим по оставшимся командам как по вариантам
    Dim i As Long
    For i = startIndex + 1 To UBound(allCommands)
        Dim variant As String
        variant = Trim(CStr(allCommands(i)))
        
        If InStr(variant, ":") > 0 Then
            Dim variantParts As Variant
            variantParts = Split(variant, ":", 2)
            
            Dim range As String, result As String
            range = Trim(CStr(variantParts(0)))
            result = Trim(CStr(variantParts(1)))
            
            If IsRollInRange(actualRoll, range) Then
                ProcessSelectCommand = result
                Exit Function
            End If
        End If
    Next i
End Function

' Обработка команды IF
Private Function ProcessIfCommand(ByVal ifCmd As String, ByVal rollValue As Long) As String
    ProcessIfCommand = ""
    
    ' IF:EVEN:активен или IF:1-6:короткий
    Dim parts As Variant
    parts = Split(Mid(ifCmd, 4), ":", 2) ' Убираем "IF:" и делим на условие:результат
    
    If UBound(parts) >= 1 Then
        Dim condition As String, result As String
        condition = UCase(Trim(CStr(parts(0))))
        result = Trim(CStr(parts(1)))
        
        Dim conditionMet As Boolean
        conditionMet = False
        
        Select Case condition
            Case "EVEN"
                conditionMet = (rollValue Mod 2 = 0)
            Case "ODD"
                conditionMet = (rollValue Mod 2 = 1)
            Case Else
                ' Проверяем диапазон чисел
                conditionMet = IsRollInRange(rollValue, condition)
        End Select
        
        If conditionMet Then
            ProcessIfCommand = result
        End If
    End If
End Function

' Проверка попадания броска в диапазон
Private Function IsRollInRange(ByVal rollValue As Long, ByVal rangeStr As String) As Boolean
    IsRollInRange = False
    
    If InStr(rangeStr, "-") > 0 Then
        ' Диапазон: "1-6", "7-8"
        Dim rangeParts As Variant
        rangeParts = Split(rangeStr, "-")
        If UBound(rangeParts) >= 1 Then
            Dim minVal As Long, maxVal As Long
            minVal = CLng(rangeParts(0))
            maxVal = CLng(rangeParts(1))
            IsRollInRange = (rollValue >= minVal And rollValue <= maxVal)
        End If
    ElseIf InStr(rangeStr, ">=") > 0 Then
        ' Больше или равно: ">=5"
        Dim threshold As Long
        threshold = CLng(Mid(rangeStr, 3))
        IsRollInRange = (rollValue >= threshold)
    ElseIf InStr(rangeStr, "<=") > 0 Then
        ' Меньше или равно: "<=3"
        Dim threshold As Long
        threshold = CLng(Mid(rangeStr, 3))
        IsRollInRange = (rollValue <= threshold)
    ElseIf InStr(rangeStr, ">") > 0 Then
        ' Больше: ">5"
        Dim threshold As Long
        threshold = CLng(Mid(rangeStr, 2))
        IsRollInRange = (rollValue > threshold)
    ElseIf InStr(rangeStr, "<") > 0 Then
        ' Меньше: "<3"
        Dim threshold As Long
        threshold = CLng(Mid(rangeStr, 2))
        IsRollInRange = (rollValue < threshold)
    Else
        ' Точное значение: "5"
        IsRollInRange = (rollValue = CLng(rangeStr))
    End If
End Function

' ========================================
' ФУНКЦИИ ПРОВЕРКИ УНИКАЛЬНОСТИ (ИСПРАВЛЕНО)
' ========================================

' Проверка является ли событие уникальным
Private Function IsUniqueEvent(ByVal eventText As String) As Boolean
    IsUniqueEvent = False
    
    ' Список фраз, указывающих на уникальность события
    Dim uniqueIndicators As Variant
    uniqueIndicators = Array( _
        "Если у вас уже есть это знание, вы перебрасываете", _
        "Если у вас уже есть это преимущество, вы перебрасываете", _
        "перебрасываете этот результат" _
    )
    
    Dim i As Long
    For i = 0 To UBound(uniqueIndicators)
        If InStr(eventText, CStr(uniqueIndicators(i))) > 0 Then
            IsUniqueEvent = True
            Exit Function
        End If
    Next i
End Function

' Проверка использовалось ли событие ранее
Private Function IsEventAlreadyUsed(ByVal eventText As String) As Boolean
    On Error GoTo NotFound
    
    ' Пытаемся найти событие в коллекции
    UsedUniqueEvents.Item(eventText)  ' ИСПРАВЛЕНО: убрали "CharHistory."
    IsEventAlreadyUsed = True
    Exit Function
    
NotFound:
    IsEventAlreadyUsed = False
End Function

' ========================================
' ФУНКЦИИ ГЕНЕРАЦИИ (ОСНОВНЫЕ)
' ========================================

' Генерация опасности (чистая логика)
Private Function GenerateDangerData(ByVal path As String, ByVal colIndex As Long, _
                                   ByRef danger As Boolean, ByRef variantDanger As Variant, _
                                   ByRef roll100 As Long, ByRef Porog As Variant) As Boolean
    GenerateDangerData = False
    
    ' Получаем порог опасности с использованием кэша
    Dim rngPorog As Range
    Set rngPorog = ValidationLogger.ValidateAndCacheRange(State.Config.PrefixPorog, _
                                                         "GenerateDangerData", State.Config.DataSheetName)
    If rngPorog Is Nothing Then Exit Function
    
    Porog = SafeVLookup(rngPorog, path, 2, "GenerateDangerData", "порог опасности")
    If IsError(Porog) Then Exit Function
    
    ' Делаем бросок и проверяем опасность
    roll100 = Roll(100)
    danger = (roll100 < CLng(Porog))
    
    ' Получаем тип опасности
    variantDanger = ""
    If danger Then
        Dim rngOpasnost As Range
        Set rngOpasnost = ValidationLogger.ValidateAndCacheRange(State.Config.PrefixOpasnost, _
                                                               "GenerateDangerData", State.Config.DataSheetName)
        If Not rngOpasnost Is Nothing Then
            variantDanger = SafeVLookup(rngOpasnost, Roll(10), colIndex, "GenerateDangerData", "тип опасности")
            If IsError(variantDanger) Then
                variantDanger = "Ошибка. Проверьте диапазон '" & State.Config.PrefixOpasnost & "'"
            End If
        End If
    End If
    
    GenerateDangerData = True
End Function

' Запись расширенных результатов в UI
Private Sub WriteEnhancedResultsToUI(ByRef wsUI As Worksheet, ByRef eventResult As EventResult)
    With wsUI
        .Range(UI_B24).Value = eventResult.BaseEvent
        .Range(UI_B25).Value = eventResult.DetailedDescription
        .Range(UI_B26).Value = eventResult.SubRolls
    End With
End Sub

' Обновленная процедура генерации с использованием новых функций
Private Sub GenerateDanger(ByRef wsUI As Worksheet, ByVal path As String, ByVal colIndex As Long, _
                         ByRef danger As Boolean, ByRef variantDanger As Variant, _
                         ByRef roll100 As Long, ByRef Porog As Variant)
    
    If Not GenerateDangerData(path, colIndex, danger, variantDanger, roll100, Porog) Then
        Exit Sub
    End If
    
    ' Записываем результаты опасности в интерфейс
    With wsUI
        .Range(UI_A25).Value = danger
        .Range(UI_A26).Value = variantDanger
    End With
End Sub

' ========================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ========================================

' Безопасный VLookup с логированием ошибок
Private Function SafeVLookup(ByVal rng As Range, ByVal lookupValue As Variant, _
                           ByVal colIndex As Long, ByVal source As String, _
                           ByVal dataType As String) As Variant
    On Error GoTo ErrHandler
    
    If rng.Columns.Count < colIndex Then
        ValidationLogger.AddValidationError source, "КОНФИГУРАЦИЯ", _
                                           "Недостаточно колонок для " & dataType, _
                                           "Требуется " & colIndex & " колонок, найдено " & rng.Columns.Count
        SafeVLookup = CVErr(xlErrRef)
        Exit Function
    End If
    
    SafeVLookup = Application.WorksheetFunction.VLookup(lookupValue, rng, colIndex, False)
    Exit Function

ErrHandler:
    ValidationLogger.AddValidationError source, "ВНИМАНИЕ", _
                                       "Ошибка поиска " & dataType, _
                                       "Значение '" & lookupValue & "' не найдено"
    SafeVLookup = CVErr(xlErrNA)
End Function

' Универсальная функция броска кости
Private Function Roll(maxValue As Long) As Long
    Roll = Int(Rnd * maxValue) + 1
End Function

' Получение или создание листа истории
Private Function GetOrCreateHistorySheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateHistorySheet = ThisWorkbook.Worksheets(HISTORY_SHEET)
    
    If GetOrCreateHistorySheet Is Nothing Then
        Set GetOrCreateHistorySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateHistorySheet.Name = HISTORY_SHEET
        
        With GetOrCreateHistorySheet
            .Cells(1, "A").Value = "Дата"
            .Cells(1, "B").Value = "Класс"
            .Cells(1, "C").Value = "Путь"
            .Cells(1, "D").Value = "Бросок"
            .Cells(1, "E").Value = "Порог"
            .Cells(1, "F").Value = "Опасность"
            .Cells(1, "G").Value = "Тип опасности"
            .Cells(1, "H").Value = "Тип события"
            .Cells(1, "I").Value = "Детали события"
            .Cells(1, "J").Value = "Дополнительные броски"
            .Cells(1, "K").Value = "Уникальное"
            
            .Range("A1:K1").Font.Bold = True
            .Columns("A:K").AutoFit
            .Columns("I").ColumnWidth = 50
            .Columns("J").ColumnWidth = 30
        End With
    End If
    
    On Error GoTo 0
End Function

' Сохранение расширенных результатов в историю
Private Sub SaveEnhancedToHistory(ByRef wsHist As Worksheet, ByVal path As String, _
                                 ByVal roll100 As Long, ByVal Porog As Variant, _
                                 ByVal danger As Boolean, ByVal variantDanger As Variant, _
                                 ByRef eventResult As EventResult)
    With wsHist
        Dim nextRow As Long
        nextRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
        
        .Cells(nextRow, "A").Value = Now
        .Cells(nextRow, "B").Value = State.Config.className
        .Cells(nextRow, "C").Value = path
        .Cells(nextRow, "D").Value = roll100
        .Cells(nextRow, "E").Value = Porog
        .Cells(nextRow, "F").Value = IIf(danger, "Да", "Нет")
        .Cells(nextRow, "G").Value = variantDanger
        .Cells(nextRow, "H").Value = eventResult.BaseEvent
        .Cells(nextRow, "I").Value = eventResult.DetailedDescription
        .Cells(nextRow, "J").Value = eventResult.SubRolls
        .Cells(nextRow, "K").Value = IIf(eventResult.IsUnique, "Да", "")
    End With
End Sub

' ========================================
' СОВМЕСТИМОСТЬ И ОБРАТНАЯ СОВМЕСТИМОСТЬ
' ========================================

' Устаревшая процедура для обратной совместимости
Public Sub GenerateEvent_SaveHistory()
    ' Перенаправляем на новую систему
    GenerateEnhancedEvent_SaveHistory
End Sub