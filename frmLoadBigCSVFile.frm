VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoadBigCSVFile 
   Caption         =   "��������� CSV ������ � ������� ����������� ����� v.2"
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
Const MINBUFFER As Long = 10000 '����������� ������ ������ � ������
Const MAXBUFFER As Long = 2147483647 '������������ ������ ������ � ������
Const MINKOLROWS As Long = 2 '����������� ���������� �����
Const MAXKOLROWS As Long = 1048576 '������������ ���������� ����� (����������� ����� 2010)
Const MINKOLROWSTEMPARR As Long = 1 '����������� ���������� ����� ���������� �������
Const MAXKOLROWSTEMPARR As Long = 1048576 '������������ ���������� ����� ���������� ������� (����������� ����� 2010)

Private ThisBook As Workbook

'��� �������� ����
Private pgOnePx  As Single

'������������ �����
Private Sub UserForm_Initialize()
    '��������� ���������� �������
    CBrazdelitel.AddItem "����� � �������"
    CBrazdelitel.AddItem "���������"
    CBrazdelitel.AddItem "�����"
    CBrazdelitel.AddItem "�������"
    CBrazdelitel.ListIndex = 0
          
    Call EnDisElements(False) '��������� ��������
    
    ProgressBar.Width = 0 '����� �������� ����
    
    Set ThisBook = Application.ActiveWorkbook '��� ����������� ������������ �� ��� �����
End Sub

'����� ����� � ������� �������� ����� � �������
Private Sub CBVibFil_Click()
    Dim TempFileName As Variant
    If ProverkaDannih() Then
        TempFileName = Application.GetOpenFilename(, , "�������� ����", , False)
    
        If TempFileName <> False Then
            LfileName.Caption = TempFileName
            LsaveFileName.Caption = LoadBigCSVFile.GetNameFile(TempFileName, "FoldName") & ".xlsb" '������ ��� ������������ �����
            Call SetOption '������������� ���������
            Call LoadBigCSVFile.PodschetKolRows '��������� �������
            Call SetRezult   '���������� ���������� ������
            Call EnDisElements(True) '���������� ������ �������� ������
        End If
    End If
End Sub

'��������� ����� ������������ �����
Private Sub CBSave_Click()
    Dim TempFileName As Variant
    
    TempFileName = LoadBigCSVFile.GetNameFile(LsaveFileName.Caption, "FoldName")
    TempFileName = Application.GetSaveAsFilename(TempFileName, _
        "�������� ���� Excel (*.xlsb),*.xlsb,���� Excel (*.xlsx),*.xlsx,���� Excel � ���������� �������� (*.xlsm),*.xlsm", _
        0, "������� ��� � ����� ���������� �����")
        
    If TempFileName <> False Then
        LsaveFileName.Caption = TempFileName
    End If
End Sub

'������ �������� ������
Private Sub CBLoad_Click()
    If ProverkaDannih("Data") Then
        Call SetOption("Data") '������������� ���������
        Call EnDisElements(False) '��������� ������
        Call LoadBigCSVFile.LoadDataOnList '��������� ��������
        MsgBox ("�������� ������ ���������, ���� ��������.")
        ThisBook.Activate
    End If
End Sub

'�������� ������
'��������:  "KolStr" - ������ ��� �������� ����� � �������� ������ (�� ���������)
'           "Data" - ��� �������� ������
Private Function ProverkaDannih(Optional Etap As String = "KolStr") As Boolean
    Dim tempMaxKolRows As Long
      
    '��������� ������ ������
    ProverkaDannih = ProverkaChisla(TBbufferSize, MINBUFFER, MAXBUFFER)
    
    '���� �������� ������� �� ��������
    If Not ProverkaDannih Then
        Exit Function
    End If
            
    '���� �������� ������ ��� �������� ���������� �����, �� ���������� �������� �� �����
    If Etap = "KolStr" Then
        Exit Function
    End If
    
    '��������� ������ ���������� �����
    tempMaxKolRows = RaschetMaxKolRows()
    ProverkaDannih = ProverkaChisla(TBkolRowsOnList, MINKOLROWS, tempMaxKolRows)
   
    '���� �������� ���������� ����� �� ��������
    If Not ProverkaDannih Then
        Exit Function
    End If
    
    '��������� ������ ���������� ����� ���������� ������
    ProverkaDannih = ProverkaChisla(TBmaxKolRowsInArr, MINKOLROWSTEMPARR, MAXKOLROWSTEMPARR)
        
    '���� �������� ���������� ����� ���������� ������ �� ��������
    If Not ProverkaDannih Then
        Exit Function
    End If
End Function

'�������� �����
Private Function ProverkaChisla(ByRef TB As Object, ByVal Min As Long, ByVal Max As Long) As Boolean
    ProverkaChisla = True
    If IsNumeric(TB.Value) Then
        If CCur(TB.Value) < Min Or CCur(TB.Value) > Max Then
            ProverkaChisla = False
        End If
    Else
        ProverkaChisla = False
    End If
    
    '���� �������� ������� �� ��������
    If Not ProverkaChisla Then
        MsgBox ("������� ����� � ��������� �� " & Min & " �� " & Max)
        TB.SetFocus
        Exit Function
    End If
End Function

'������ ������������� ���������� �����
Private Function RaschetMaxKolRows() As Long
    If CLng(LkolRows.Caption) < MAXKOLROWS Then
        RaschetMaxKolRows = CLng(LkolRows.Caption)
    Else
        RaschetMaxKolRows = MAXKOLROWS
    End If
End Function

'������ ����� � ���������� ����� ������
Private Sub CBrazdelitel_Change()
    CBVibFil.SetFocus
End Sub

'����� ���� �����������
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

'�������� �������
Private Sub CBstop_Click()
    End
End Sub

'������ ���������� ����������� ������
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


'�������� ����� ����
Private Sub ProverkaVvodaCifr(ByRef KeyAscii As MSForms.ReturnInteger)
    If KeyAscii.Value < Asc(0) Or KeyAscii.Value > Asc(9) Then
        KeyAscii.Value = 0
    End If
End Sub

Private Sub TBkolRowsOnList_Change()
    Call RaschetListov
End Sub

'���������� � ��������� ���������
Private Sub EnDisElements(ByVal sostojanie As Boolean)
    CBLoad.Enabled = sostojanie
    TBkolRowsOnList.Enabled = sostojanie
    CBSave.Enabled = sostojanie
End Sub

'������������� ���������
'��������:  "KolStr" - ������ ��� �������� ����� � �������� ������ (�� ���������)
'           "Data" - ��� �������� ������
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

'������� ��������� ������ ����� � �������
Private Sub SetRezult()
    LfileSize.Caption = CStr(fileSize) & " ����"
    LkolRows.Caption = CStr(LoadBigCSVFile.kolRows)
    LkolCols.Caption = CStr(LoadBigCSVFile.kolCols)
    TBsoderjFirstLine.Value = LoadBigCSVFile.soderjFirstLine
        
    If LoadBigCSVFile.kolRows < CLng(TBkolRowsOnList.Value) Then
        TBkolRowsOnList.Value = CStr(LoadBigCSVFile.kolRows) '������������� ���������� ����� �� ����
    End If
    Call RaschetListov '������������ ����������� ���������� ������
    If Len(LoadBigCSVFile.soderjFirstLine) > 100 Then '��������� ������ ��������� ���� ����� 100 �������� � ������
        Frame1.Height = Frame1.Height + 16
        TBsoderjFirstLine.Height = 32
    End If
End Sub

'������������ �������� ����
Public Sub PgInit(ByVal Min As Long, ByVal Max As Long, ByVal Value As Long)
    pgOnePx = (Max - Min) / (ProgressBarRamka.Width - 1)
    PgUpdate (Value)
End Sub

'�������� �������� ������������
Public Sub PgUpdate(ByVal Value As Long)
    If Round(Value / pgOnePx, 0) <> ProgressBar.Width Then
        ProgressBar.Width = Round(Value / pgOnePx, 0)
        DoEvents
    End If
End Sub

