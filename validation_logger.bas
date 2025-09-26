Option Explicit

' ========================================
' ValidationLogger - Модуль валидации и логирования ошибок
' Автор: [Yuriy] + Claude Sonnet 4
' Назначение: Централизованная система сбора и отображения ошибок валидации
' ========================================

' ---------- Глобальные переменные ----------
Private m_ValidationErrors As Collection    ' Коллекция ошибок валидации (массивы Variant)
Private m_IsValidationEnabled As Boolean    ' Флаг включения валидации
Private m_CacheNames As Collection          ' Кэш имен диапазонов
Private m_CacheRanges As Collection         ' Кэш объектов Range
Private m_CacheTimestamps As Collection     ' Кэш временных меток
Private m_CacheTimeout As Long             ' Таймаут кэша в минутах

' ---------- Константы ----------
Private Const ERROR_TYPE_CRITICAL As String = "КРИТИЧНО"
Private Const ERROR_TYPE_WARNING As String = "ВНИМАНИЕ" 
Private Const ERROR_TYPE_CONFIG As String = "КОНФИГУРАЦИЯ"
Private Const CACHE_TIMEOUT_MINUTES As Long = 30
Private Const ERROR_SHEET_NAME As String = "Ошибки_Валидации"

' ========================================
' ПУБЛИЧНЫЕ ФУНКЦИИ - Интерфейс модуля
' ========================================

' Инициализация системы валидации
Public Sub InitializeValidation()
    Set m_ValidationErrors = New Collection
    Set m_CacheNames = New Collection
    Set m_CacheRanges = New Collection
    Set m_CacheTimestamps = New Collection
    m_IsValidationEnabled = True
    m_CacheTimeout = CACHE_TIMEOUT_MINUTES
End Sub

' Добавление ошибки в коллекцию
Public Sub AddValidationError(ByVal source As String, ByVal errorType As String, _
                             ByVal message As String, Optional ByVal details As String = "")
    If Not m_IsValidationEnabled Then Exit Sub
    
    ' Проверяем инициализацию коллекции
    If m_ValidationErrors Is Nothing Then Set m_ValidationErrors = New Collection
    
    ' Создаем массив для хранения данных об ошибке
    Dim errorData(0 To 4) As Variant
    errorData(0) = Now          ' Timestamp
    errorData(1) = source       ' Source
    errorData(2) = errorType    ' ErrorType
    errorData(3) = message      ' Message
    errorData(4) = details      ' Details
    
    m_ValidationErrors.Add errorData
End Sub

' Проверка наличия ошибок
Public Function HasErrors() As Boolean
    If m_ValidationErrors Is Nothing Then
        HasErrors = False
    Else
        HasErrors = (m_ValidationErrors.Count > 0)
    End If
End Function

' Проверка наличия критических ошибок
Public Function HasCriticalErrors() As Boolean
    HasCriticalErrors = False
    
    ' Проверяем инициализацию
    If m_ValidationErrors Is Nothing Then Exit Function
    If m_ValidationErrors.Count = 0 Then Exit Function
    
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim currentError As Variant ' Массив с данными об ошибке
    
    For i = 1 To m_ValidationErrors.Count
        currentError = m_ValidationErrors(i)
        
        ' Проверяем, что currentError является массивом и имеет нужный размер
        If IsArray(currentError) Then
            If UBound(currentError) >= 2 Then
                If currentError(2) = ERROR_TYPE_CRITICAL Then ' ErrorType
                    HasCriticalErrors = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    ' Если возникла ошибка, считаем что критических ошибок нет
    HasCriticalErrors = False
End Function

' Отображение результатов валидации
Public Sub ShowValidationResults()
    If Not HasErrors() Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = GetOrCreateErrorSheet()
    
    ' Записываем ошибки в лист
    WriteErrorsToSheet ws
    
    ' Показываем уведомление
    If HasCriticalErrors() Then
        MsgBox "Обнаружены критические ошибки конфигурации!" & vbCrLf & _
               "Подробности на листе '" & ERROR_SHEET_NAME & "'", vbCritical
    Else
        MsgBox "Обнаружены предупреждения валидации." & vbCrLf & _
               "Подробности на листе '" & ERROR_SHEET_NAME & "'", vbExclamation
    End If
End Sub

' Очистка системы валидации
Public Sub ClearValidation()
    Set m_ValidationErrors = Nothing
    m_IsValidationEnabled = False
    ClearRangeCache
End Sub

' ========================================
' ФУНКЦИИ ВАЛИДАЦИИ КОНФИГУРАЦИИ
' ========================================

' Валидация существования листа
Public Function ValidateWorksheet(ByVal sheetName As String, ByVal source As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If ws Is Nothing Then
        AddValidationError source, ERROR_TYPE_CRITICAL, _
                          "Отсутствует лист: " & sheetName, _
                          "Создайте лист '" & sheetName & "' или проверьте название"
        ValidateWorksheet = False
    Else
        ValidateWorksheet = True
    End If
    
    On Error GoTo 0
End Function

' Валидация именованного диапазона с кэшированием
Public Function ValidateAndCacheRange(ByVal rangeName As String, ByVal source As String, _
                                     ByVal dataSheetName As String) As Range
    Set ValidateAndCacheRange = Nothing
    
    Debug.Print "Ищем диапазон: " & rangeName & " на листе: " & dataSheetName
    
    ' Проверяем кэш
    Set ValidateAndCacheRange = GetFromCache(rangeName)
    If Not ValidateAndCacheRange Is Nothing Then 
        Debug.Print "Найден в кэше: " & rangeName
        Exit Function
    End If
    
    ' Ищем диапазон
    On Error GoTo ErrHandler
    
    ' Сначала ищем в именованных диапазонах
    Dim testRange As Range
    Set testRange = Nothing
    
    ' Пробуем найти именованный диапазон
    Set testRange = ThisWorkbook.Names(rangeName).RefersToRange
    
    ' Если найден именованный диапазон
    If Not testRange Is Nothing Then
        Debug.Print "Найден именованный диапазон: " & rangeName
        Set ValidateAndCacheRange = testRange
        
        ' БЕЗОПАСНО добавляем в кэш
        AddToCache rangeName, ValidateAndCacheRange
        On Error GoTo 0
        Exit Function
    End If
    
ErrHandler:
    ' Если ошибка при поиске именованного диапазона, пробуем на листе
    Debug.Print "Именованный диапазон не найден, ищем на листе: " & dataSheetName
    
    On Error GoTo ErrHandler2
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(dataSheetName)
    
    If Not ws Is Nothing Then
        Set testRange = Nothing
        Set testRange = ws.Range(rangeName)
        
        If Not testRange Is Nothing Then
            Debug.Print "Найден диапазон на листе: " & rangeName
            Set ValidateAndCacheRange = testRange
            
            ' БЕЗОПАСНО добавляем в кэш
            AddToCache rangeName, ValidateAndCacheRange
            On Error GoTo 0
            Exit Function
        End If
    End If
    
ErrHandler2:
    Debug.Print "ДИАПАЗОН НЕ НАЙДЕН: " & rangeName
    AddValidationError source, ERROR_TYPE_CRITICAL, _
                      "Отсутствует диапазон: " & rangeName, _
                      "Создайте именованный диапазон или проверьте название на листе " & dataSheetName
    Set ValidateAndCacheRange = Nothing
    On Error GoTo 0
End Function

' Валидация структуры диапазона (количество колонок)
Public Function ValidateRangeStructure(ByVal rng As Range, ByVal expectedColumns As Long, _
                                      ByVal rangeName As String, ByVal source As String) As Boolean
    If rng Is Nothing Then
        ValidateRangeStructure = False
        Exit Function
    End If
    
    If rng.Columns.Count < expectedColumns Then
        AddValidationError source, ERROR_TYPE_CONFIG, _
                          "Недостаточно колонок в диапазоне: " & rangeName, _
                          "Ожидается " & expectedColumns & " колонок, найдено " & rng.Columns.Count
        ValidateRangeStructure = False
    Else
        ValidateRangeStructure = True
    End If
End Function

' Валидация конфигурации класса
Public Sub ValidateClassConfiguration(ByVal className As String, ByVal dataSheetName As String, _
                                     ByVal prefixOpasnost As String, ByVal prefixPorog As String, _
                                     ByVal prefixPath As String, ByVal paths As Variant)
    Dim source As String
    source = "ValidateClassConfiguration[" & className & "]"
    
    ' Проверяем лист данных
    ValidateWorksheet dataSheetName, source
    
    ' Проверяем основные диапазоны
    Dim rngOpasnost As Range, rngPorog As Range, rngPath As Range
    
    Set rngOpasnost = ValidateAndCacheRange(prefixOpasnost, source, dataSheetName)
    Set rngPorog = ValidateAndCacheRange(prefixPorog, source, dataSheetName)
    Set rngPath = ValidateAndCacheRange(prefixPath, source, dataSheetName)
    
    ' Проверяем структуру диапазонов
    If Not rngOpasnost Is Nothing Then
        ValidateRangeStructure rngOpasnost, 5, prefixOpasnost, source ' 4 пути + 1 ключ
    End If
    
    If Not rngPorog Is Nothing Then
        ValidateRangeStructure rngPorog, 2, prefixPorog, source ' путь + порог
    End If
    
    If Not rngPath Is Nothing Then
        ValidateRangeStructure rngPath, 5, prefixPath, source ' 4 пути + 1 ключ
    End If
End Sub

' ========================================
' ФУНКЦИИ КЭШИРОВАНИЯ
' ========================================

' Получение диапазона из кэша
Private Function GetFromCache(ByVal rangeName As String) As Range
    Set GetFromCache = Nothing
    
    ' Проверяем инициализацию кэша
    If m_CacheNames Is Nothing Then Exit Function
    If m_CacheNames.Count = 0 Then Exit Function
    
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim foundIndex As Long
    foundIndex = 0
    
    ' Ищем индекс по имени
    For i = 1 To m_CacheNames.Count
        If CStr(m_CacheNames(i)) = rangeName Then
            foundIndex = i
            Exit For
        End If
    Next i
    
    ' Если не найден
    If foundIndex = 0 Then Exit Function
    
    ' Проверяем корректность индексов перед обращением
    If foundIndex > m_CacheTimestamps.Count Or foundIndex > m_CacheRanges.Count Then
        Debug.Print "Несоответствие размеров коллекций кэша для: " & rangeName
        Exit Function
    End If
    
    ' Проверяем таймаут
    Dim cacheTime As Date
    cacheTime = CDate(m_CacheTimestamps(foundIndex))
    
    If DateDiff("n", cacheTime, Now) <= m_CacheTimeout Then
        ' Кэш актуален, возвращаем диапазон
        Set GetFromCache = m_CacheRanges(foundIndex)
        ' НЕ ОБНОВЛЯЕМ время доступа, чтобы избежать ошибок с коллекциями
    Else
        ' Кэш устарел, удаляем запись
        RemoveCacheEntry foundIndex
    End If
    
    Exit Function
    
ErrHandler:
    ' В случае ошибки возвращаем Nothing
    Set GetFromCache = Nothing
    Debug.Print "Ошибка GetFromCache: " & Err.Number & " - " & Err.Description & " для: " & rangeName
End Function

' Добавление диапазона в кэш
Private Sub AddToCache(ByVal rangeName As String, ByVal rng As Range)
    ' Проверяем инициализацию кэшей
    If m_CacheNames Is Nothing Then Set m_CacheNames = New Collection
    If m_CacheRanges Is Nothing Then Set m_CacheRanges = New Collection
    If m_CacheTimestamps Is Nothing Then Set m_CacheTimestamps = New Collection
    
    On Error GoTo ErrHandler
    
    ' БЕЗОПАСНАЯ ПРОВЕРКА: ищем существующую запись
    Dim foundIndex As Long
    foundIndex = 0
    
    Dim i As Long
    For i = 1 To m_CacheNames.Count
        If CStr(m_CacheNames(i)) = rangeName Then
            foundIndex = i
            Exit For
        End If
    Next i
    
    ' Если найдена существующая запись - удаляем её полностью
    If foundIndex > 0 Then
        Debug.Print "Обновляем существующую запись кэша: " & rangeName
        
        ' БЕЗОПАСНОЕ УДАЛЕНИЕ: проверяем размеры коллекций
        If foundIndex <= m_CacheNames.Count Then m_CacheNames.Remove foundIndex
        If foundIndex <= m_CacheRanges.Count Then m_CacheRanges.Remove foundIndex
        If foundIndex <= m_CacheTimestamps.Count Then m_CacheTimestamps.Remove foundIndex
    End If
    
    ' Добавляем новую запись во все коллекции
    m_CacheNames.Add rangeName
    m_CacheRanges.Add rng
    m_CacheTimestamps.Add Now
    
    Debug.Print "Успешно добавлен в кэш: " & rangeName & " (всего записей: " & m_CacheNames.Count & ")"
    Exit Sub
    
ErrHandler:
    ' В случае ошибки очищаем весь кэш и добавляем запись заново
    Debug.Print "КРИТИЧЕСКАЯ ОШИБКА AddToCache: " & Err.Number & " - " & Err.Description & " для диапазона: " & rangeName
    Debug.Print "Очищаем весь кэш и добавляем запись заново"
    
    ' Полная очистка кэша
    Set m_CacheNames = New Collection
    Set m_CacheRanges = New Collection
    Set m_CacheTimestamps = New Collection
    
    ' Добавляем текущую запись
    On Error Resume Next
    m_CacheNames.Add rangeName
    m_CacheRanges.Add rng
    m_CacheTimestamps.Add Now
    On Error GoTo 0
End Sub

' Удаление записи из кэша по индексу
Private Sub RemoveCacheEntry(ByVal index As Long)
    On Error Resume Next
    
    ' Проверяем корректность индекса перед удалением
    If index > 0 And index <= m_CacheNames.Count Then
        m_CacheNames.Remove index
    End If
    
    If index > 0 And index <= m_CacheRanges.Count Then
        m_CacheRanges.Remove index
    End If
    
    If index > 0 And index <= m_CacheTimestamps.Count Then
        m_CacheTimestamps.Remove index
    End If
    
    On Error GoTo 0
End Sub

' Очистка кэша
Public Sub ClearRangeCache()
    Set m_CacheNames = Nothing
    Set m_CacheRanges = Nothing
    Set m_CacheTimestamps = Nothing
    Debug.Print "Кэш очищен"
End Sub

' ========================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ========================================

' Создание или получение листа ошибок
Private Function GetOrCreateErrorSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateErrorSheet = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)
    
    If GetOrCreateErrorSheet Is Nothing Then
        Set GetOrCreateErrorSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateErrorSheet.Name = ERROR_SHEET_NAME
        
        ' Инициализация заголовков
        With GetOrCreateErrorSheet
            .Cells(1, "A").Value = "Дата/Время"
            .Cells(1, "B").Value = "Источник"
            .Cells(1, "C").Value = "Тип"
            .Cells(1, "D").Value = "Сообщение"
            .Cells(1, "E").Value = "Детали"
            
            .Range("A1:E1").Font.Bold = True
            .Range("A1:E1").Interior.Color = RGB(200, 200, 200)
            .Columns("A").ColumnWidth = 20
            .Columns("B").ColumnWidth = 25
            .Columns("C").ColumnWidth = 15
            .Columns("D").ColumnWidth = 40
            .Columns("E").ColumnWidth = 50
        End With
    End If
    
    On Error GoTo 0
End Function

' Запись ошибок в лист
Private Sub WriteErrorsToSheet(ByVal ws As Worksheet)
    If m_ValidationErrors Is Nothing Then Exit Sub
    If m_ValidationErrors.Count = 0 Then Exit Sub
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim errorItem As Variant ' Массив: (Timestamp, Source, ErrorType, Message, Details)
    
    For i = 1 To m_ValidationErrors.Count
        errorItem = m_ValidationErrors(i)
        
        ' Проверяем, что errorItem является массивом и имеет нужный размер
        If IsArray(errorItem) And UBound(errorItem) >= 4 Then
            With ws
                .Cells(nextRow, "A").Value = errorItem(0) ' Timestamp
                .Cells(nextRow, "B").Value = errorItem(1) ' Source
                .Cells(nextRow, "C").Value = errorItem(2) ' ErrorType
                .Cells(nextRow, "D").Value = errorItem(3) ' Message
                .Cells(nextRow, "E").Value = errorItem(4) ' Details
                
                ' Цветовое кодирование
                Select Case CStr(errorItem(2)) ' ErrorType
                    Case ERROR_TYPE_CRITICAL
                        .Range(.Cells(nextRow, "A"), .Cells(nextRow, "E")).Interior.Color = RGB(255, 200, 200)
                    Case ERROR_TYPE_CONFIG
                        .Range(.Cells(nextRow, "A"), .Cells(nextRow, "E")).Interior.Color = RGB(255, 255, 200)
                    Case ERROR_TYPE_WARNING
                        .Range(.Cells(nextRow, "A"), .Cells(nextRow, "E")).Interior.Color = RGB(240, 240, 240)
                End Select
            End With
            
            nextRow = nextRow + 1
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    ' В случае ошибки просто выходим
    Debug.Print "Ошибка в WriteErrorsToSheet: " & Err.Number & " - " & Err.Description
End Sub


    
Public Sub DiagnosticsCache()
    Debug.Print "=== Диагностика кэша ==="
    
    If m_CacheNames Is Nothing Then
        Debug.Print "Кэш имен: не инициализирован"
    Else
        Debug.Print "Кэш имен: " & m_CacheNames.Count & " элементов"
        Dim i As Long
        For i = 1 To m_CacheNames.Count
            Debug.Print "  " & i & ": " & CStr(m_CacheNames(i))
        Next i
    End If
    
    If m_CacheRanges Is Nothing Then
        Debug.Print "Кэш диапазонов: не инициализирован"
    Else
        Debug.Print "Кэш диапазонов: " & m_CacheRanges.Count & " элементов"
    End If
    
    If m_CacheTimestamps Is Nothing Then
        Debug.Print "Кэш времени: не инициализирован"
    Else
        Debug.Print "Кэш времени: " & m_CacheTimestamps.Count & " элементов"
    End If
    
    Debug.Print "==========================="
End Sub