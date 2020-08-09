Attribute VB_Name = "LoadBigCSVFile"
Option Explicit

'�������� ���������
Public bufferSize As Long
Public razdelitel As String
Public fileName As String
Public saveFileName As String
Public kolRowsOnList As Long
Public ViklObnovlenieEkrana As Boolean
Public log As Object '����� ���������� � ������� ��������
Public maxKolRowsInArr As Long

'����������� ������
Public fileSize As Long
Public kolRows, kolCols As Long
Public kolListov As Integer
Public soderjFirstLine As String

'������ ���������
Private Book As Workbook
Private ArrRange() As Variant
Private rowInArrRange As Long
Private OstRowOnList As Long

'������������
Public Sub Start()
    frmLoadBigCSVFile.Show vbModal
End Sub

'������������ ���������� ����� � ��������
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
    
    T0 = Timer() '�������� ������ �������
      
    Call SendLog("�������� �����.")
    
    LfAnsi = StrConv(vbLf, vbFromUnicode) '������ ����� ������
    F = FreeFile '����� ��������� ����� ����
    Open fileName For Binary Access Read As #F ' ��������� ���� ��� ������
    fileSize = LOF(F) '������ �����
                
    '������������� ����������
    Call frmLoadBigCSVFile.PgInit(0, fileSize, 0)
           
    Call SendLog("������������ ������� ���������� ����� � ��������")
    
    ReDim Buffer(bufferSize - 1) ' ������� ������
    BytesLeft = fileSize
    bufPos = 0
    FirstLine = False
    Do Until BytesLeft = 0
        If bufPos = 0 Then
            '�������� ���� � ������ �� ������
            If BytesLeft < bufferSize Then
                ReDim Buffer(BytesLeft - 1)
            End If
            Get #F, , Buffer ' ��������� ����� ���� � ������
            strBuffer = Buffer '�������� � ������
            BytesLeft = BytesLeft - LenB(strBuffer)
            bufPos = 1
        End If
        '������� ������
        Do Until bufPos = 0
            bufPos = InStrB(bufPos, strBuffer, LfAnsi)
            If bufPos > 0 Then
                '���� ��� ������ ������ �� ��������� �� ��� ����� � ��������� ���������� �������
                If Not FirstLine Then
                    soderjFirstLine = StrConv(LeftB(strBuffer, bufPos - 1), vbUnicode)
                    kolCols = UBound(Split(soderjFirstLine, razdelitel)) + 1
                    FirstLine = True
                End If
                LineCount = LineCount + 1
                bufPos = bufPos + 1
            End If
        Loop
        '�������� �������� ���
        Call frmLoadBigCSVFile.PgUpdate(fileSize - BytesLeft)
    Loop

    kolRows = LineCount + 1 '���� ��������� ������
    
    Call SendLog("������� ����� � ������� �������� �� " & Format$(Timer() - T0, "0.0#") & " c.")
    
    Close #F
End Sub

'�������� ������
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
       
    T0 = Timer() '�������� ������ �������
    
    ' ��������� ���������� ������ �������� �������� ����� �� ����
    If ViklObnovlenieEkrana Then
        Call Uskorenie(True)
    End If
        
    '������������� ����������
    Call frmLoadBigCSVFile.PgInit(0, kolRows, 0)
    
    Call SendLog("�������� �����.")
    
    F = FreeFile '����� ��������� ����� ����
    Open fileName For Binary Access Read As #F ' ��������� ���� ��� ������
            
    Call SendLog("�������� ����� � ������.")
    Call AddNewBook '�������� ����� � ������
       
    Call SendLog("�������� ������")
    
    List = 1
    RowList = 1 '������������� �� ������� ������ �� �����
    TempRowsList = TekusheeKolRowsOnList(List) '������ ������� ����� ������ ���� �� �����, � ��������� �� ������ ����
    OstRowOnList = TempRowsList - 1
    Call CreateArray ' ������� ��������� ������ ��� �����
      
    ReDim Buffer(bufferSize - 1) ' ������� ������
    BytesLeft = fileSize
    Pos = 0
    PropuskFirstLine = False
    Row = 0
    TempString = ""
    strBuffer = ""
    PosledniyFragment = False
    Do Until BytesLeft = 0
        '�������� �������� ����� � ��������� ��� � ���������������� ������ �����
        If Pos = 0 Then
            If BytesLeft < bufferSize Then
                ReDim Buffer(BytesLeft - 1)
                PosledniyFragment = True
            End If
            Get #F, , Buffer ' ��������� ����� ���� � ������
            strBufferB = Buffer '�������� � ������
            BytesLeft = BytesLeft - LenB(strBufferB)
            strBuffer = strBuffer + StrConv(strBufferB, vbUnicode) '������������
            oldPos = 0
            Pos = 1
        End If
        '������������ ������������� ������ (��������� �� ������� � �������� � ��������� �� ����)
        Do Until Pos = 0
            Pos = InStr(Pos, strBuffer, vbLf)
            If Pos > 0 Then
                If Not PropuskFirstLine Then
                    '���� ������ ������, �� ���������� �� (����� �������� ���������)
                    Row = Row + 1
                    oldPos = Pos
                    Pos = Pos + 1
                    PropuskFirstLine = True
                Else
                    '������� ������ � ��������� ������ �� ����
                    TempString = Mid(strBuffer, oldPos + 1, Pos - oldPos - 1)
                    Row = Row + 1
                    RowList = RowList + 1
                    Call VstavkaRowInList(TempString, RowList, TempRowsList)
                    oldPos = Pos
                    Pos = Pos + 1
                End If
            Else
                '��������� ��� ��������� �������� ����� ��� ���
                If PosledniyFragment Then
                    '���� ��������� �������� ����� �����
                    '������������ ���������� �������� ������ ��� ��������� ������
                    TempString = Right(strBuffer, Len(strBuffer) - oldPos)
                    Row = Row + 1
                    RowList = RowList + 1
                    Call VstavkaRowInList(TempString, RowList, TempRowsList)
                Else
                    '���� �� ��������� �������� ����� �����
                    '��������� ������� ������ ��� �� ����������� ���������
                    strBuffer = Right(strBuffer, Len(strBuffer) - oldPos)
                End If
            End If
            
            '���������, ��������� �� ������� �� ����� ����
            If RowList = TempRowsList And List < kolListov Then
                List = List + 1
                RowList = 1 '������������� �� ������� ������ �� �����
                TempRowsList = TekusheeKolRowsOnList(List) '������ ������� ����� ������ ���� �� �����, � ��������� �� ������ ����
                OstRowOnList = TempRowsList - 1
                Call CreateArray ' ������� ��������� ������ ��� �����
            End If
            
        '�������� �������� ���
        Call frmLoadBigCSVFile.PgUpdate(Row)
        Loop
    Loop
     
    Call SendLog("���������� �����")
    Book.Save
    
    ' �������� ���������� ������
    If ViklObnovlenieEkrana Then
        Call Uskorenie(False)
    End If
    
    Call SendLog("�������� ������ ��������� �� " & Format$(Timer() - T0, "0.0#") & " c. ���� ��������.")
End Sub

'�������� ����� � ������ � �������
Private Sub AddNewBook()
    Dim TempList As Worksheet
    Dim i As Integer
    Dim TempCols() As String
    Dim Cols() As String
     
    '������� � ��������� �����
    Set Book = Workbooks.Add
    Select Case GetNameFile(saveFileName, "Extn")
        Case "xlsb"
            LoadBigCSVFile.Book.SaveAs fileName:=saveFileName, FileFormat:=50
        Case "xlsx"
            LoadBigCSVFile.Book.SaveAs fileName:=saveFileName, FileFormat:=51
        Case "xlsm"
            LoadBigCSVFile.Book.SaveAs fileName:=saveFileName, FileFormat:=52
    End Select
    
    '�������������� ����� ��������
    ReDim Cols(0, kolCols - 1)
    TempCols = Split(soderjFirstLine, razdelitel)
    For i = LBound(TempCols) To UBound(TempCols)
        Cols(0, i) = TempCols(i)
    Next i
    
    '������� ������ ���������� ������ � ������� �����������
    '������� � �������� �������, ��� �� ������ ��� � ���� �� �����
    For i = kolListov To 1 Step -1
        Set TempList = Book.Worksheets.Add
        TempList.Name = CStr(i)
        '��������� ����� ��������
        TempList.Range(Cells(1, 1), Cells(1, kolCols)).Value = Cols
    Next i
    
    '������� ����� ��������� �� ���������
    Application.DisplayAlerts = False
    For Each TempList In Book.Worksheets
        If TempList.Name = "����1" Or TempList.Name = "����2" Or TempList.Name = "����3" Then
            TempList.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

'������ ������� ����� ������ ���� �� �����, � ��������� �� ������ ����
Private Function TekusheeKolRowsOnList(ByVal List As Integer) As Long
    If List < kolListov Or kolListov = 1 Then
        TekusheeKolRowsOnList = kolRowsOnList
    Else
        TekusheeKolRowsOnList = kolRows - kolRowsOnList - (List - 2) * (kolRowsOnList - 1) + 1 'plus 1, ��� ��� ������ ������ ��� �����
    End If
    Book.Worksheets(CStr(List)).Select '������ ������ ���� ��������
End Function

'������� ��������� ������ �� ��������
Private Sub CreateArray()
    If maxKolRowsInArr > OstRowOnList Then
        ReDim ArrRange(OstRowOnList - 1, kolCols - 1)
    Else
        ReDim ArrRange(maxKolRowsInArr - 1, kolCols - 1)
    End If
    rowInArrRange = -1
End Sub

'������� ������ �� ����
Private Sub VstavkaRowInList(ByVal txt As String, ByVal RowList As Long, ByVal TempRowsList As Long)
    Dim ArrCols() As String
    Dim i As Long

    '��������� ������ � ������
    rowInArrRange = rowInArrRange + 1
    ArrCols = Split(txt, razdelitel)
    For i = LBound(ArrCols) To UBound(ArrCols)
        If IsNumeric(ArrCols(i)) Then
            ArrRange(rowInArrRange, i) = CDbl(ArrCols(i))
        Else
            ArrRange(rowInArrRange, i) = CStr(ArrCols(i))
        End If
    Next i
    
    '��������� ������ ������ ������������ ����������� �����
    '���� ������ ��������� ������ �� ���� � ������� ������
    '���� ��� ����������� ������ �������
    If rowInArrRange + 1 = maxKolRowsInArr Or RowList = TempRowsList Then
        Range(Cells(RowList - rowInArrRange, 1), Cells(RowList, kolCols)).Value = ArrRange
    End If
    
    If rowInArrRange + 1 = maxKolRowsInArr And RowList <> TempRowsList Then
        OstRowOnList = OstRowOnList - maxKolRowsInArr
        Call CreateArray ' ������� ��������� ������ ��� �����
    End If
    
End Sub

'�������� ����
Private Sub SendLog(ByVal Text As String)
    log.Caption = Text
    DoEvents
End Sub

'������� ��� ��������� ������ �� ������� ����� ����� (FN)
'(������ ��� FN = "C:\Users\Asmolovskiy\Docunents\Test.txt")
'Part:  "Fold" - ����� ���� � ����� (������ "C:\Users\Asmolovskiy\Docunents\t")
'       "FoldName" - ���� � ��� ����� ��� ���������� (������ "C:\Users\Asmolovskiy\Docunents\Test")
'       "Name" - ��� ����� ��� ����������� (������ "Test")
'       "NameExtn" - ��� ����� � ����������� (������ "Test.txt")
'       "Extn" - ��������� ����� (������ "txt")
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

'�������� ���������
Sub Uskorenie(ByVal Value As Boolean)
    Application.ScreenUpdating = Not Value
    Application.EnableEvents = Not Value
    Application.DisplayStatusBar = Not Value
    Application.DisplayAlerts = Not Value
End Sub




