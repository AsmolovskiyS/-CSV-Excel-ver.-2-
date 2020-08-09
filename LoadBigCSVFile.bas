Attribute VB_Name = "LoadBigCSVFile"
Option Explicit

'Входящие параметры
Public bufferSize As Long
Public razdelitel As String
Public fileName As String
Public saveFileName As String
Public kolRowsOnList As Long
Public ViklObnovlenieEkrana As Boolean
Public log As Object 'Вывод информации о текущей процессе
Public maxKolRowsInArr As Long

'Расчитанные данные
Public fileSize As Long
Public kolRows, kolCols As Long
Public kolListov As Integer
Public soderjFirstLine As String

'Данные программы
Private Book As Workbook
Private ArrRange() As Variant
Private rowInArrRange As Long
Private OstRowOnList As Long

'Подсчитываем
Public Sub Start()
    frmLoadBigCSVFile.Show vbModal
End Sub

'Подсчитываем количество строк и столбцов
Public Sub PodschetKolRows()
    Dim LfAnsi As String
    Dim F As Integer
    Dim Buffer() As Byte
    Dim bufPos As Long
    Dim LineCount As Long
    Dim FirstLine As Boolean
    Dim strBuffer  As String
    Dim BytesLeft As Long
    Dim T0 As Single
    
    T0 = Timer() 'начинаем отсчет времени
      
    Call SendLog("Открытие файла.")
    
    LfAnsi = StrConv(vbLf, vbFromUnicode) 'Символ конца строки
    F = FreeFile 'Берем свободный номер фала
    Open fileName For Binary Access Read As #F ' Открываем файл для чтения
    fileSize = LOF(F) 'Размер файла
                
    'Иницилизируем прогресбар
    Call frmLoadBigCSVFile.PgInit(0, fileSize, 0)
           
    Call SendLog("Производится подсчет количества строк и столбцов")
    
    ReDim Buffer(bufferSize - 1) ' Создаем массив
    BytesLeft = fileSize
    bufPos = 0
    FirstLine = False
    Do Until BytesLeft = 0
        If bufPos = 0 Then
            'Копируем файл в буффер по частям
            If BytesLeft < bufferSize Then
                ReDim Buffer(BytesLeft - 1)
            End If
            Get #F, , Buffer ' Загружаем часть файл в массив
            strBuffer = Buffer 'Копируем в строку
            BytesLeft = BytesLeft - LenB(strBuffer)
            bufPos = 1
        End If
        'Считаем строки
        Do Until bufPos = 0
            bufPos = InStrB(bufPos, strBuffer, LfAnsi)
            If bufPos > 0 Then
                'Если это первая строка то сохранить ее для шапки и посчитать количество колонок
                If Not FirstLine Then
                    soderjFirstLine = StrConv(LeftB(strBuffer, bufPos - 1), vbUnicode)
                    kolCols = UBound(Split(soderjFirstLine, razdelitel)) + 1
                    FirstLine = True
                End If
                LineCount = LineCount + 1
                bufPos = bufPos + 1
            End If
        Loop
        'Обновить прогресс бар
        Call frmLoadBigCSVFile.PgUpdate(fileSize - BytesLeft)
    Loop

    kolRows = LineCount + 1 'Плюс последняя строка
    
    Call SendLog("Подсчет строк и колонок завершен за " & Format$(Timer() - T0, "0.0#") & " c.")
    
    Close #F
End Sub

'Загрузка данных
Sub LoadDataOnList()
    Dim List As Integer
    Dim TempRowsList As Long
    Dim RowsOnPx As Long
    Dim Pos, oldPos  As Long
    Dim PropuskFirstLine As Boolean
    Dim strBufferB, strBuffer As String
    Dim BytesLeft As Long
    Dim F As Integer
    Dim Buffer() As Byte
    Dim Row, RowList As Long
    Dim TempString As String
    Dim PosledniyFragment As Boolean
    Dim T0 As Single
       
    T0 = Timer() 'начинаем отсчет времени
    
    ' Отклюение обновления экрана прироста скорости особо не дает
    If ViklObnovlenieEkrana Then
        Call Uskorenie(True)
    End If
        
    'Иницилизируем прогресбар
    Call frmLoadBigCSVFile.PgInit(0, kolRows, 0)
    
    Call SendLog("Открытие файла.")
    
    F = FreeFile 'Берем свободный номер фала
    Open fileName For Binary Access Read As #F ' Открываем файл для чтения
            
    Call SendLog("Создание книги и листов.")
    Call AddNewBook 'Создание книги и листов
       
    Call SendLog("Загрузка данных")
    
    List = 1
    RowList = 1 'Устанавливаем на чальную строку на лсите
    TempRowsList = TekusheeKolRowsOnList(List) 'Узнаем сколько строк должно быть на листе, и переходим на данный лист
    OstRowOnList = TempRowsList - 1
    Call CreateArray ' Создаем временный массив для строк
      
    ReDim Buffer(bufferSize - 1) ' Создаем массив
    BytesLeft = fileSize
    Pos = 0
    PropuskFirstLine = False
    Row = 0
    TempString = ""
    strBuffer = ""
    PosledniyFragment = False
    Do Until BytesLeft = 0
        'копируем фрагмент фалйа и добавляем его к недообработанном укуску стоки
        If Pos = 0 Then
            If BytesLeft < bufferSize Then
                ReDim Buffer(BytesLeft - 1)
                PosledniyFragment = True
            End If
            Get #F, , Buffer ' Загружаем часть файл в массив
            strBufferB = Buffer 'Копируем в строку
            BytesLeft = BytesLeft - LenB(strBufferB)
            strBuffer = strBuffer + StrConv(strBufferB, vbUnicode) 'Конвертируем
            oldPos = 0
            Pos = 1
        End If
        'Обрабатываем скопированную строку (разбиваем по строкам и колонкам и вставляем на лист)
        Do Until Pos = 0
            Pos = InStr(Pos, strBuffer, vbLf)
            If Pos > 0 Then
                If Not PropuskFirstLine Then
                    'Если первоя строка, то пропускаем ее (шапка ужевезде вставлена)
                    Row = Row + 1
                    oldPos = Pos
                    Pos = Pos + 1
                    PropuskFirstLine = True
                Else
                    'Готовим строку и вставляем данные на лист
                    TempString = Mid(strBuffer, oldPos + 1, Pos - oldPos - 1)
                    Row = Row + 1
                    RowList = RowList + 1
                    Call VstavkaRowInList(TempString, RowList, TempRowsList)
                    oldPos = Pos
                    Pos = Pos + 1
                End If
            Else
                'Проверяем это последний фрагмент файла или нет
                If PosledniyFragment Then
                    'Если последний фрагмент файла тогда
                    'обрабатываем оставшийся фрагмент текста как последнюю строку
                    TempString = Right(strBuffer, Len(strBuffer) - oldPos)
                    Row = Row + 1
                    RowList = RowList + 1
                    Call VstavkaRowInList(TempString, RowList, TempRowsList)
                Else
                    'Если не последний фрагмент файла тогда
                    'корпируем остаток строки для ее последующей обработки
                    strBuffer = Right(strBuffer, Len(strBuffer) - oldPos)
                End If
            End If
            
            'Проверяем, необхдимо ли перейти на новый лист
            If RowList = TempRowsList And List < kolListov Then
                List = List + 1
                RowList = 1 'Устанавливаем на чальную строку на лсите
                TempRowsList = TekusheeKolRowsOnList(List) 'Узнаем сколько строк должно быть на листе, и переходим на данный лист
                OstRowOnList = TempRowsList - 1
                Call CreateArray ' Создаем временный массив для строк
            End If
            
        'Обновить прогресс бар
        Call frmLoadBigCSVFile.PgUpdate(Row)
        Loop
    Loop
     
    Call SendLog("Сохранение книги")
    Book.Save
    
    ' Включаем обновления экрана
    If ViklObnovlenieEkrana Then
        Call Uskorenie(False)
    End If
    
    Call SendLog("Загрузка данных завершена за " & Format$(Timer() - T0, "0.0#") & " c. Файл сохранен.")
End Sub

'Создание книги и лситов с шапками
Private Sub AddNewBook()
    Dim TempList As Worksheet
    Dim i As Integer
    Dim TempCols() As String
    Dim Cols() As String
     
    'Создаем и сохраняем книгу
    Set Book = Workbooks.Add
    Select Case GetNameFile(saveFileName, "Extn")
        Case "xlsb"
            LoadBigCSVFile.Book.SaveAs fileName:=saveFileName, FileFormat:=50
        Case "xlsx"
            LoadBigCSVFile.Book.SaveAs fileName:=saveFileName, FileFormat:=51
        Case "xlsm"
            LoadBigCSVFile.Book.SaveAs fileName:=saveFileName, FileFormat:=52
    End Select
    
    'Подготавливаем шапку столбцов
    ReDim Cols(0, kolCols - 1)
    TempCols = Split(soderjFirstLine, razdelitel)
    For i = LBound(TempCols) To UBound(TempCols)
        Cols(0, i) = TempCols(i)
    Next i
    
    'Создаем нужное количество листов и удаляем стандартные
    'Создаем в обратном порядке, что бы номера шли с лева на право
    For i = kolListov To 1 Step -1
        Set TempList = Book.Worksheets.Add
        TempList.Name = CStr(i)
        'Добавляем шапку столбцов
        TempList.Range(Cells(1, 1), Cells(1, kolCols)).Value = Cols
    Next i
    
    'Удаляем листы созданные по умолчанию
    Application.DisplayAlerts = False
    For Each TempList In Book.Worksheets
        If TempList.Name = "Лист1" Or TempList.Name = "Лист2" Or TempList.Name = "Лист3" Then
            TempList.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

'Узнаем сколько строк должно быть на листе, и переходим на данный лист
Private Function TekusheeKolRowsOnList(ByVal List As Integer) As Long
    If List < kolListov Or kolListov = 1 Then
        TekusheeKolRowsOnList = kolRowsOnList
    Else
        TekusheeKolRowsOnList = kolRows - kolRowsOnList - (List - 2) * (kolRowsOnList - 1) + 1 'plus 1, так как первоя строка это шапка
    End If
    Book.Worksheets(CStr(List)).Select 'Делаем данный лист активным
End Function

'Создаем временный массив со строками
Private Sub CreateArray()
    If maxKolRowsInArr > OstRowOnList Then
        ReDim ArrRange(OstRowOnList - 1, kolCols - 1)
    Else
        ReDim ArrRange(maxKolRowsInArr - 1, kolCols - 1)
    End If
    rowInArrRange = -1
End Sub

'Вставка данных на лист
Private Sub VstavkaRowInList(ByVal txt As String, ByVal RowList As Long, ByVal TempRowsList As Long)
    Dim ArrCols() As String
    Dim i As Long

    'Добавляем строку в массив
    rowInArrRange = rowInArrRange + 1
    ArrCols = Split(txt, razdelitel)
    For i = LBound(ArrCols) To UBound(ArrCols)
        If IsNumeric(ArrCols(i)) Then
            ArrRange(rowInArrRange, i) = CDbl(ArrCols(i))
        Else
            ArrRange(rowInArrRange, i) = CStr(ArrCols(i))
        End If
    Next i
    
    'Проверяем достиг масиив максимальног околичества строк
    'Если достиг вставляем данные на лист и очищаем массив
    'Если нет увеличиваем размер массива
    If rowInArrRange + 1 = maxKolRowsInArr Or RowList = TempRowsList Then
        Range(Cells(RowList - rowInArrRange, 1), Cells(RowList, kolCols)).Value = ArrRange
    End If
    
    If rowInArrRange + 1 = maxKolRowsInArr And RowList <> TempRowsList Then
        OstRowOnList = OstRowOnList - maxKolRowsInArr
        Call CreateArray ' Создаем временный массив для строк
    End If
    
End Sub

'Отправка лога
Private Sub SendLog(ByVal Text As String)
    log.Caption = Text
    DoEvents
End Sub

'Функция для получения частей от полного имени файла (FN)
'(Пример для FN = "C:\Users\Asmolovskiy\Docunents\Test.txt")
'Part:  "Fold" - поный путь к папке (Вернет "C:\Users\Asmolovskiy\Docunents\t")
'       "FoldName" - путь и имя файла без расширения (Вернет "C:\Users\Asmolovskiy\Docunents\Test")
'       "Name" - имя файла без расширением (Вернет "Test")
'       "NameExtn" - имя файла с расширением (Вернет "Test.txt")
'       "Extn" - раширение файла (Вернет "txt")
Function GetNameFile(ByVal FN As String, ByVal Part As String) As String
    Select Case Part
        Case "Fold"
            GetNameFile = Left(FN, InStrRev(FN, "\"))
        Case "FoldName"
            GetNameFile = Left(FN, InStrRev(FN, ".") - 1)
        Case "Name"
            GetNameFile = Mid(FN, InStrRev(FN, "\") + 1, Len(FN) - InStrRev(FN, "."))
        Case "NameExtn"
            GetNameFile = Right(FN, Len(FN) - InStrRev(FN, "\"))
        Case "Extn"
            GetNameFile = Right(FN, Len(FN) - InStrRev(FN, "."))
        Case Else
            GetNameFile = ""
    End Select
End Function

'Включить ускорение
Sub Uskorenie(ByVal Value As Boolean)
    Application.ScreenUpdating = Not Value
    Application.EnableEvents = Not Value
    Application.DisplayStatusBar = Not Value
    Application.DisplayAlerts = Not Value
End Sub




