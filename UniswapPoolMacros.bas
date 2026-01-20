Attribute VB_Name = "UniswapPoolMacros"
'===============================================================================
' Модуль: UniswapPoolMacros
' Описание: Набор макросов для автоматизации операций с пулами ликвидности Uniswap
' Версия: 1.0
' Дата: 2026-01-20
'===============================================================================

Option Explicit

'===============================================================================
' Константы для имён листов
'===============================================================================
Const SHEET_POOLS As String = "Пулы"
Const SHEET_FEES As String = "Заработанные комиссии"
Const SHEET_HISTORY As String = "История операций"

'===============================================================================
' Константы для столбцов листа "Пулы"
'===============================================================================
Const COL_POOL_ID As Integer = 1        ' A - ID позиции
Const COL_POOL_DATE As Integer = 2      ' B - Дата открытия
Const COL_POOL_PAIR As Integer = 3      ' C - Пара токенов
Const COL_POOL_RANGE As Integer = 4     ' D - Диапазон цен
Const COL_POOL_AMOUNT As Integer = 5    ' E - Сумма позиции
Const COL_POOL_STATUS As Integer = 6    ' F - Статус

'===============================================================================
' Константы для столбцов листа "Заработанные комиссии"
'===============================================================================
Const COL_FEE_ID As Integer = 1         ' A - ID записи
Const COL_FEE_DATE As Integer = 2       ' B - Дата сбора
Const COL_FEE_POOL_ID As Integer = 3    ' C - ID пула
Const COL_FEE_AMOUNT As Integer = 4     ' D - Сумма комиссии
Const COL_FEE_TOKEN As Integer = 5      ' E - Токен

'===============================================================================
' Макрос: AddNewPosition
' Описание: Добавляет новую строку в лист "Пулы" с автозаполнением ID и даты
' Параметры: Нет
' Возврат: Нет
'===============================================================================
Sub AddNewPosition()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newRow As Long
    Dim newID As Long
    Dim currentDate As Date
    Dim answer As VbMsgBoxResult

    ' Проверка существования листа
    If Not WorksheetExists(SHEET_POOLS) Then
        MsgBox "Лист '" & SHEET_POOLS & "' не найден!" & vbCrLf & _
               "Пожалуйста, создайте лист с именем '" & SHEET_POOLS & "' перед использованием макроса.", _
               vbCritical, "Ошибка"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(SHEET_POOLS)

    ' Подтверждение действия
    answer = MsgBox("Добавить новую позицию в пул?", vbQuestion + vbYesNo, "Подтверждение")
    If answer = vbNo Then Exit Sub

    ' Находим последнюю заполненную строку
    lastRow = ws.Cells(ws.Rows.Count, COL_POOL_ID).End(xlUp).Row

    ' Если это первая строка после заголовка
    If lastRow = 1 Or ws.Cells(lastRow, COL_POOL_ID).Value = "ID" Then
        newRow = 2
        newID = 1
    Else
        newRow = lastRow + 1
        ' Генерируем новый ID (на 1 больше последнего)
        newID = CLng(ws.Cells(lastRow, COL_POOL_ID).Value) + 1
    End If

    currentDate = Date

    ' Отключаем обновление экрана для ускорения
    Application.ScreenUpdating = False

    ' Заполняем новую строку
    With ws
        .Cells(newRow, COL_POOL_ID).Value = newID
        .Cells(newRow, COL_POOL_DATE).Value = currentDate
        .Cells(newRow, COL_POOL_DATE).NumberFormat = "DD.MM.YYYY"
        .Cells(newRow, COL_POOL_PAIR).Value = ""
        .Cells(newRow, COL_POOL_RANGE).Value = ""
        .Cells(newRow, COL_POOL_AMOUNT).Value = 0
        .Cells(newRow, COL_POOL_AMOUNT).NumberFormat = "#,##0.00"
        .Cells(newRow, COL_POOL_STATUS).Value = "Активна"

        ' Применяем форматирование
        .Rows(newRow).Font.Name = "Arial"
        .Rows(newRow).Font.Size = 10

        ' Выделяем новую строку
        .Rows(newRow).Select
    End With

    Application.ScreenUpdating = True

    MsgBox "Новая позиция добавлена!" & vbCrLf & _
           "ID: " & newID & vbCrLf & _
           "Дата: " & Format(currentDate, "DD.MM.YYYY") & vbCrLf & _
           "Статус: Активна", _
           vbInformation, "Успех"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Произошла ошибка при добавлении позиции:" & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, _
           vbCritical, "Ошибка"
End Sub

'===============================================================================
' Макрос: CollectFees
' Описание: Добавляет запись о сборе комиссий в лист "Заработанные комиссии"
' Параметры: Нет
' Возврат: Нет
'===============================================================================
Sub CollectFees()
    On Error GoTo ErrorHandler

    Dim wsPool As Worksheet
    Dim wsFees As Worksheet
    Dim lastRow As Long
    Dim newRow As Long
    Dim newID As Long
    Dim currentDate As Date
    Dim poolID As Variant
    Dim feeAmount As Variant
    Dim feeToken As String
    Dim answer As VbMsgBoxResult

    ' Проверка существования листов
    If Not WorksheetExists(SHEET_POOLS) Then
        MsgBox "Лист '" & SHEET_POOLS & "' не найден!", vbCritical, "Ошибка"
        Exit Sub
    End If

    If Not WorksheetExists(SHEET_FEES) Then
        MsgBox "Лист '" & SHEET_FEES & "' не найден!" & vbCrLf & _
               "Пожалуйста, создайте лист с именем '" & SHEET_FEES & "' перед использованием макроса.", _
               vbCritical, "Ошибка"
        Exit Sub
    End If

    Set wsPool = ThisWorkbook.Sheets(SHEET_POOLS)
    Set wsFees = ThisWorkbook.Sheets(SHEET_FEES)

    ' Запрос ID пула
    poolID = Application.InputBox("Введите ID пула для сбора комиссий:", "ID пула", Type:=1)
    If poolID = False Then Exit Sub ' Пользователь отменил

    ' Проверка существования пула
    If Not PoolExists(wsPool, CLng(poolID)) Then
        MsgBox "Пул с ID " & poolID & " не найден!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Проверка статуса пула
    If Not IsPoolActive(wsPool, CLng(poolID)) Then
        answer = MsgBox("Пул с ID " & poolID & " не активен!" & vbCrLf & _
                       "Продолжить?", vbQuestion + vbYesNo, "Предупреждение")
        If answer = vbNo Then Exit Sub
    End If

    ' Запрос суммы комиссии
    feeAmount = Application.InputBox("Введите сумму собранной комиссии:", "Сумма комиссии", Type:=1)
    If feeAmount = False Then Exit Sub ' Пользователь отменил

    If feeAmount <= 0 Then
        MsgBox "Сумма комиссии должна быть больше нуля!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Запрос токена
    feeToken = InputBox("Введите токен комиссии (например, ETH, USDC):", "Токен комиссии", "ETH")
    If feeToken = "" Then Exit Sub ' Пользователь отменил

    ' Подтверждение
    answer = MsgBox("Добавить запись о сборе комиссий?" & vbCrLf & _
                   "ID пула: " & poolID & vbCrLf & _
                   "Сумма: " & feeAmount & " " & feeToken, _
                   vbQuestion + vbYesNo, "Подтверждение")
    If answer = vbNo Then Exit Sub

    ' Находим последнюю заполненную строку
    lastRow = wsFees.Cells(wsFees.Rows.Count, COL_FEE_ID).End(xlUp).Row

    ' Если это первая строка после заголовка
    If lastRow = 1 Or wsFees.Cells(lastRow, COL_FEE_ID).Value = "ID" Then
        newRow = 2
        newID = 1
    Else
        newRow = lastRow + 1
        newID = CLng(wsFees.Cells(lastRow, COL_FEE_ID).Value) + 1
    End If

    currentDate = Date

    Application.ScreenUpdating = False

    ' Заполняем новую строку
    With wsFees
        .Cells(newRow, COL_FEE_ID).Value = newID
        .Cells(newRow, COL_FEE_DATE).Value = currentDate
        .Cells(newRow, COL_FEE_DATE).NumberFormat = "DD.MM.YYYY"
        .Cells(newRow, COL_FEE_POOL_ID).Value = CLng(poolID)
        .Cells(newRow, COL_FEE_AMOUNT).Value = CDbl(feeAmount)
        .Cells(newRow, COL_FEE_AMOUNT).NumberFormat = "#,##0.00000000"
        .Cells(newRow, COL_FEE_TOKEN).Value = UCase(Trim(feeToken))

        ' Применяем форматирование
        .Rows(newRow).Font.Name = "Arial"
        .Rows(newRow).Font.Size = 10
    End With

    Application.ScreenUpdating = True

    MsgBox "Запись о сборе комиссий добавлена!" & vbCrLf & _
           "ID записи: " & newID & vbCrLf & _
           "Дата: " & Format(currentDate, "DD.MM.YYYY") & vbCrLf & _
           "Пул: " & poolID & vbCrLf & _
           "Сумма: " & feeAmount & " " & feeToken, _
           vbInformation, "Успех"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Произошла ошибка при добавлении записи:" & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, _
           vbCritical, "Ошибка"
End Sub

'===============================================================================
' Макрос: ClosePosition
' Описание: Закрывает позицию, обновляя её статус на "Закрыта"
' Параметры: Нет
' Возврат: Нет
'===============================================================================
Sub ClosePosition()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim poolID As Variant
    Dim poolRow As Long
    Dim answer As VbMsgBoxResult
    Dim oldStatus As String

    ' Проверка существования листа
    If Not WorksheetExists(SHEET_POOLS) Then
        MsgBox "Лист '" & SHEET_POOLS & "' не найден!", vbCritical, "Ошибка"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(SHEET_POOLS)

    ' Запрос ID пула
    poolID = Application.InputBox("Введите ID позиции для закрытия:", "ID позиции", Type:=1)
    If poolID = False Then Exit Sub ' Пользователь отменил

    ' Поиск пула
    poolRow = FindPoolRow(ws, CLng(poolID))

    If poolRow = 0 Then
        MsgBox "Позиция с ID " & poolID & " не найдена!", vbExclamation, "Ошибка"
        Exit Sub
    End If

    ' Получаем текущий статус
    oldStatus = ws.Cells(poolRow, COL_POOL_STATUS).Value

    ' Проверка, не закрыта ли уже позиция
    If UCase(Trim(oldStatus)) = "ЗАКРЫТА" Then
        MsgBox "Позиция с ID " & poolID & " уже закрыта!", vbInformation, "Информация"
        Exit Sub
    End If

    ' Подтверждение
    answer = MsgBox("Закрыть позицию ID " & poolID & "?" & vbCrLf & _
                   "Текущий статус: " & oldStatus, _
                   vbQuestion + vbYesNo, "Подтверждение")
    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' Обновляем статус
    ws.Cells(poolRow, COL_POOL_STATUS).Value = "Закрыта"

    ' Применяем форматирование (серый цвет для закрытой позиции)
    ws.Rows(poolRow).Interior.Color = RGB(220, 220, 220)
    ws.Rows(poolRow).Font.Color = RGB(100, 100, 100)

    Application.ScreenUpdating = True

    MsgBox "Позиция закрыта!" & vbCrLf & _
           "ID: " & poolID & vbCrLf & _
           "Предыдущий статус: " & oldStatus & vbCrLf & _
           "Новый статус: Закрыта", _
           vbInformation, "Успех"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Произошла ошибка при закрытии позиции:" & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, _
           vbCritical, "Ошибка"
End Sub

'===============================================================================
' Макрос: UpdateAllData
' Описание: Принудительный пересчёт всех формул и обновление связей
' Параметры: Нет
' Возврат: Нет
'===============================================================================
Sub UpdateAllData()
    On Error GoTo ErrorHandler

    Dim answer As VbMsgBoxResult
    Dim startTime As Double
    Dim endTime As Double

    ' Подтверждение
    answer = MsgBox("Обновить все данные и формулы?" & vbCrLf & _
                   "Это может занять некоторое время для больших таблиц.", _
                   vbQuestion + vbYesNo, "Подтверждение")
    If answer = vbNo Then Exit Sub

    startTime = Timer

    ' Показываем процесс
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Обновление данных..."

    ' Обновляем все связи
    On Error Resume Next
    ThisWorkbook.UpdateLinks
    On Error GoTo ErrorHandler

    ' Пересчитываем все формулы
    Application.CalculateFull

    ' Обновляем сводные таблицы (если есть)
    Dim pt As PivotTable
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    On Error GoTo ErrorHandler

    ' Восстанавливаем настройки
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    endTime = Timer

    Application.StatusBar = False

    MsgBox "Обновление завершено!" & vbCrLf & _
           "Время выполнения: " & Format(endTime - startTime, "0.00") & " сек.", _
           vbInformation, "Успех"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Произошла ошибка при обновлении данных:" & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, _
           vbCritical, "Ошибка"
End Sub

'===============================================================================
' Вспомогательные функции
'===============================================================================

'-------------------------------------------------------------------------------
' Функция: WorksheetExists
' Описание: Проверяет существование листа с указанным именем
' Параметры: sheetName - имя листа
' Возврат: True если лист существует, False в противном случае
'-------------------------------------------------------------------------------
Private Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' Функция: FindPoolRow
' Описание: Находит строку с указанным ID пула
' Параметры: ws - лист для поиска, poolID - ID пула
' Возврат: Номер строки или 0 если не найдено
'-------------------------------------------------------------------------------
Private Function FindPoolRow(ws As Worksheet, poolID As Long) As Long
    Dim lastRow As Long
    Dim i As Long

    FindPoolRow = 0
    lastRow = ws.Cells(ws.Rows.Count, COL_POOL_ID).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, COL_POOL_ID).Value = poolID Then
            FindPoolRow = i
            Exit Function
        End If
    Next i
End Function

'-------------------------------------------------------------------------------
' Функция: PoolExists
' Описание: Проверяет существование пула с указанным ID
' Параметры: ws - лист для поиска, poolID - ID пула
' Возврат: True если пул существует, False в противном случае
'-------------------------------------------------------------------------------
Private Function PoolExists(ws As Worksheet, poolID As Long) As Boolean
    PoolExists = (FindPoolRow(ws, poolID) > 0)
End Function

'-------------------------------------------------------------------------------
' Функция: IsPoolActive
' Описание: Проверяет активность пула
' Параметры: ws - лист для поиска, poolID - ID пула
' Возврат: True если пул активен, False в противном случае
'-------------------------------------------------------------------------------
Private Function IsPoolActive(ws As Worksheet, poolID As Long) As Boolean
    Dim poolRow As Long
    Dim status As String

    IsPoolActive = False
    poolRow = FindPoolRow(ws, poolID)

    If poolRow > 0 Then
        status = UCase(Trim(ws.Cells(poolRow, COL_POOL_STATUS).Value))
        IsPoolActive = (status = "АКТИВНА" Or status = "АКТИВЕН" Or status = "ACTIVE")
    End If
End Function
