VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoadBigCSVFile 
   Caption         =   "Загрузщик CSV файлоф с большим количеством строк v.2"
   ClientHeight    =   9360.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9705.001
   OleObjectBlob   =   "frmLoadBigCSVFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoadBigCSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MINBUFFER As Long = 10000 'Минимальный размер буфера в байтах
Const MAXBUFFER As Long = 2147483647 'Максимальный размер буфера в байтах
Const MINKOLROWS As Long = 2 'Минимальное количество строк
Const MAXKOLROWS As Long = 1048576 'Максимальное количсетво строк (ограничение офиса 2010)
Const MINKOLROWSTEMPARR As Long = 1 'Минимальное количество строк временного массива
Const MAXKOLROWSTEMPARR As Long = 1048576 'Максимальное количсетво строк временного массива (ограничение офиса 2010)

Private ThisBook As Workbook

'Для прогресс бара
Private pgOnePx  As Single

'Иницилизация формы
Private Sub UserForm_Initialize()
    'Заполняем командБокс данными
    CBrazdelitel.AddItem "Точка с запятой"
    CBrazdelitel.AddItem "Табуляция"
    CBrazdelitel.AddItem "Точка"
    CBrazdelitel.AddItem "Запятая"
    CBrazdelitel.ListIndex = 0
          
    Call EnDisElements(False) 'Отключаем элементы
    
    ProgressBar.Width = 0 'Сброс прогресс бара
    
    Set ThisBook = Application.ActiveWorkbook 'Для возможности переключения на эту книгу
End Sub

'Выбор файла и запуска подсчета строк и колонок
Private Sub CBVibFil_Click()
    Dim TempFileName As Variant
    If ProverkaDannih() Then
        TempFileName = Application.GetOpenFilename(, , "Выберите файл", , False)
    
        If TempFileName <> False Then
            LfileName.Caption = TempFileName
            LsaveFileName.Caption = LoadBigCSVFile.GetNameFile(TempFileName, "FoldName") & ".xlsb" 'Задаем имя сохраняемого файла
            Call SetOption 'Устанавливаем настройки
            Call LoadBigCSVFile.PodschetKolRows 'Запускаем подсчет
            Call SetRezult   'Записываем полученные данные
            Call EnDisElements(True) 'Активируем кнопки загрузки данных
        End If
    End If
End Sub

'Изменение имени сохраняемого файла
Private Sub CBSave_Click()
    Dim TempFileName As Variant
    
    TempFileName = LoadBigCSVFile.GetNameFile(LsaveFileName.Caption, "FoldName")
    TempFileName = Application.GetSaveAsFilename(TempFileName, _
        "Двоичный файл Excel (*.xlsb),*.xlsb,Файл Excel (*.xlsx),*.xlsx,Файл Excel с поддержкой макросов (*.xlsm),*.xlsm", _
        0, "Укажите имя и место сохранения файла")
        
    If TempFileName <> False Then
        LsaveFileName.Caption = TempFileName
    End If
End Sub

'Запуск загрузки данных
Private Sub CBLoad_Click()
    If ProverkaDannih("Data") Then
        Call SetOption("Data") 'Устанавливаем настройки
        Call EnDisElements(False) 'Отключаем кнопки
        Call LoadBigCSVFile.LoadDataOnList 'Запускаем загрузку
        MsgBox ("Загрузка данных завершена, файл сохранен.")
        ThisBook.Activate
    End If
End Sub

'Проверка данных
'Парметры:  "KolStr" - данные для подсчета строк и загрузки данных (по умолчанию)
'           "Data" - для загрузки данных
Private Function ProverkaDannih(Optional Etap As String = "KolStr") As Boolean
    Dim tempMaxKolRows As Long
      
    'Проверяем данные буфера
    ProverkaDannih = ProverkaChisla(TBbufferSize, MINBUFFER, MAXBUFFER)
    
    'Если проверка буффера не пройдена
    If Not ProverkaDannih Then
        Exit Function
    End If
            
    'Если проверка велась для подсчета количества строк, то дальнейшие проверки не нужны
    If Etap = "KolStr" Then
        Exit Function
    End If
    
    'Проверяем данные количества строк
    tempMaxKolRows = RaschetMaxKolRows()
    ProverkaDannih = ProverkaChisla(TBkolRowsOnList, MINKOLROWS, tempMaxKolRows)
   
    'Если проверка количества строк не пройдена
    If Not ProverkaDannih Then
        Exit Function
    End If
    
    'Проверяем данные количества строк временного масива
    ProverkaDannih = ProverkaChisla(TBmaxKolRowsInArr, MINKOLROWSTEMPARR, MAXKOLROWSTEMPARR)
        
    'Если проверка количества строк временного масива не пройдена
    If Not ProverkaDannih Then
        Exit Function
    End If
End Function

'Проверка числа
Private Function ProverkaChisla(ByRef TB As Object, ByVal Min As Long, ByVal Max As Long) As Boolean
    ProverkaChisla = True
    If IsNumeric(TB.Value) Then
        If CCur(TB.Value) < Min Or CCur(TB.Value) > Max Then
            ProverkaChisla = False
        End If
    Else
        ProverkaChisla = False
    End If
    
    'Если проверка буффера не пройдена
    If Not ProverkaChisla Then
        MsgBox ("Введите число в диапозоне от " & Min & " до " & Max)
        TB.SetFocus
        Exit Function
    End If
End Function

'Расчет максимального количества строк
Private Function RaschetMaxKolRows() As Long
    If CLng(LkolRows.Caption) < MAXKOLROWS Then
        RaschetMaxKolRows = CLng(LkolRows.Caption)
    Else
        RaschetMaxKolRows = MAXKOLROWS
    End If
End Function

'Убрать фокус с комбоБокса после выбора
Private Sub CBrazdelitel_Change()
    CBVibFil.SetFocus
End Sub

'Выбор типа разделителя
Private Function TypeRazdelitelja() As String
    Select Case CBrazdelitel.ListIndex
        Case 0
            TypeRazdelitelja = ";"
        Case 1
            TypeRazdelitelja = vbTab
        Case 2
            TypeRazdelitelja = "."
        Case 3
            TypeRazdelitelja = ","
    End Select
End Function

'Прервать процесс
Private Sub CBstop_Click()
    End
End Sub

'Расчет количества необходимых листов
Private Sub RaschetListov()
    Dim tempKolRow As Long
    
    If Len(TBkolRowsOnList.Value) > 0 Then
        If IsNumeric(TBkolRowsOnList.Value) Then
            tempKolRow = CLng(TBkolRowsOnList.Value)
            If tempKolRow >= MINKOLROWS And tempKolRow <= RaschetMaxKolRows() Then
                LkolListov.Caption = CStr( _
                    WorksheetFunction.RoundUp((CLng(LkolRows.Caption) - 1) / (tempKolRow - 1), 0))
                Exit Sub
            End If
        End If
    End If
    LkolListov.Caption = "-"
End Sub

Private Sub TBkolRowsOnList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ProverkaVvodaCifr(KeyAscii)
End Sub

Private Sub TBbufferSize_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ProverkaVvodaCifr(KeyAscii)
End Sub

Private Sub TBmaxKolRowsInArr_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ProverkaVvodaCifr(KeyAscii)
End Sub


'Проверка ввода цифр
Private Sub ProverkaVvodaCifr(ByRef KeyAscii As MSForms.ReturnInteger)
    If KeyAscii.Value < Asc(0) Or KeyAscii.Value > Asc(9) Then
        KeyAscii.Value = 0
    End If
End Sub

Private Sub TBkolRowsOnList_Change()
    Call RaschetListov
End Sub

'Отключение и включение элементов
Private Sub EnDisElements(ByVal sostojanie As Boolean)
    CBLoad.Enabled = sostojanie
    TBkolRowsOnList.Enabled = sostojanie
    CBSave.Enabled = sostojanie
End Sub

'Устанавливаем настройки
'Парметры:  "KolStr" - данные для подсчета строк и загрузки данных (по умолчанию)
'           "Data" - для загрузки данных
Private Sub SetOption(Optional Etap As String = "KolStr")
    LoadBigCSVFile.bufferSize = CLng(TBbufferSize.Value)
    LoadBigCSVFile.razdelitel = TypeRazdelitelja()
    LoadBigCSVFile.fileName = LfileName.Caption
    LoadBigCSVFile.ViklObnovlenieEkrana = CBviklObnovlenieEkrana.Value
    Set LoadBigCSVFile.log = Llog
    If Etap = "Data" Then
        LoadBigCSVFile.saveFileName = LsaveFileName.Caption
        LoadBigCSVFile.kolListov = CInt(LkolListov.Caption)
        LoadBigCSVFile.kolRowsOnList = CLng(TBkolRowsOnList.Value)
        LoadBigCSVFile.maxKolRowsInArr = CLng(TBmaxKolRowsInArr.Value)
    End If
End Sub

'Выводим результат расета строк и колонок
Private Sub SetRezult()
    LfileSize.Caption = CStr(fileSize) & " байт"
    LkolRows.Caption = CStr(LoadBigCSVFile.kolRows)
    LkolCols.Caption = CStr(LoadBigCSVFile.kolCols)
    TBsoderjFirstLine.Value = LoadBigCSVFile.soderjFirstLine
        
    If LoadBigCSVFile.kolRows < CLng(TBkolRowsOnList.Value) Then
        TBkolRowsOnList.Value = CStr(LoadBigCSVFile.kolRows) 'Устанавливаем количество строк на лист
    End If
    Call RaschetListov 'Подсчитываем необходимое количество листов
    If Len(LoadBigCSVFile.soderjFirstLine) > 100 Then 'Добавляем полосу прокрутки если более 100 символов с строке
        Frame1.Height = Frame1.Height + 16
        TBsoderjFirstLine.Height = 32
    End If
End Sub

'Иницилизация прогресс бара
Public Sub PgInit(ByVal Min As Long, ByVal Max As Long, ByVal Value As Long)
    pgOnePx = (Max - Min) / (ProgressBarRamka.Width - 1)
    PgUpdate (Value)
End Sub

'Обновить прогресс прогрессбара
Public Sub PgUpdate(ByVal Value As Long)
    If Round(Value / pgOnePx, 0) <> ProgressBar.Width Then
        ProgressBar.Width = Round(Value / pgOnePx, 0)
        DoEvents
    End If
End Sub

