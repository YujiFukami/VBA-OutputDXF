Attribute VB_Name = "ModFile"
Option Explicit
Sub SaveSheetAsBook(TargetSheet As Worksheet, Optional SaveName$, Optional SavePath$, _
                           Optional MessageIruNaraTrue As Boolean = False)
'�w��̃V�[�g��ʃu�b�N�ŕۑ�����
'20210719�쐬
                           
    '���͈����̒���
    If SaveName = "" Then
        SaveName = TargetSheet.Name
    End If
    If SavePath = "" Then
        SavePath = TargetSheet.Parent.Path
    End If
    
    '�ʃu�b�N�ŕۑ�
    TargetSheet.Copy
    ActiveWorkbook.SaveAs SavePath & "\" & SaveName
    ActiveWorkbook.Close
    
    If MessageIruNaraTrue Then
        MsgBox ("�V�[�g���u" & TargetSheet.Name & "�v��" & vbLf & _
               "�u" & SavePath & "�v��" & vbLf & _
               "�t�@�C�����u" & SaveName & ".xlsx�v�ŕۑ����܂����B")
    End If
    
End Sub

Function GetSheetByName(SheetName$) As Worksheet
'�w��̖��O�̃V�[�g�����[�N�V�[�g�I�u�W�F�N�g�Ƃ��Ď擾����
'20210715�쐬

    Dim Output As Worksheet
    On Error Resume Next
    Set Output = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    
    If Output Is Nothing Then
        MsgBox ("�u" & SheetName & "�v�V�[�g������܂���I�I")
        End
    End If
    
    Set GetSheetByName = Output

End Function

Function InputCSV(CSVPath$)
'CSV�t�@�C����ǂݍ���Ŕz��`���ŕԂ�
'20210706�쐬

    '���͒l�m�F
    Dim Dummy
    If Dir(CSVPath, vbDirectory) = "" Then
        Dummy = MsgBox(CSVPath & "�̃t�@�C���͑��݂��܂���", vbOKOnly + vbCritical)
        Exit Function
    End If
    
    Dim intFree As Integer
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpStr$, TmpSplit
    Dim StrList
    Dim Output

    intFree = FreeFile '��ԍ����擾
    Open CSVPath For Input As #intFree 'CSV�t�@�B�����I�[�v��
    
    K = 0
    ReDim StrList(1 To 1)
    Do Until EOF(intFree)
        Line Input #intFree, TmpStr '1�s�ǂݍ���
        K = K + 1
        ReDim Preserve StrList(1 To K)
        StrList(K) = TmpStr
        
        M = WorksheetFunction.Max(UBound(Split(TmpStr, ",")) + 1, M)
    Loop
        
    Close #intFree
    N = K
    ReDim Output(1 To N, 1 To M)
    
    For I = 1 To N
        TmpStr = StrList(I)
        TmpSplit = Split(TmpStr, ",")
        
        For J = 0 To UBound(TmpSplit)
            Output(I, J + 1) = TmpSplit(J)
        Next J
    Next I
        
    InputCSV = Output
    
End Function

Function InputBook(BookFolderPath$, BookName$, SheetName$, StartCellAddress$, Optional EndCellAddress$)
'�u�b�N���J���Ȃ��Ńf�[�^���擾����
'ExecuteExcel4Macro���g�p����̂ŁAExcel�̃o�[�W�����A�b�v�̎��ɒ���
'20210720

'BookFolderPath�E�E�E�w��u�b�N�̃t�H���_�p�X
'BookName�E�E�E�w��u�b�N�̖��O �g���q�܂�
'SheetName�E�E�E�w��u�b�N�̎擾�ΏۂƂȂ�V�[�g�̖��O
'StartCellAddress�E�E�E�擾�͈͂̍ŏ��̃Z���A�h���X(��:"A1")
'EndCellAddress�E�E�E�擾�͈͂̍Ō�̃Z��(��F"B3")�i�ȗ��Ȃ�StartCellAddress�Ɠ����j
    
    Dim Rs&, Re&, Cs&, Ce& '�n�[�s,��ԍ�����яI�[�s,��ԍ�(Long�^)
    Dim strRC$
    With Range(StartCellAddress)
        Rs = .Row
        Cs = .Column
    End With
    
    If EndCellAddress = "" Then
        Re = Rs
        Ce = Cs
    Else
        With Range(EndCellAddress)
            Re = .Row
            Ce = .Column
        End With
    End If
    
    '�n�_�A�I�_�̔��]���Ă���ꍇ�̏���
    Dim Dummy&
    If Re < Rs Then
        Dummy = Rs
        Re = Rs
        Rs = Dummy
    End If
    
    If Ce < Cs Then
        Dummy = Cs
        Ce = Cs
        Cs = Dummy
    End If

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    If Rs = Re And Cs = Ce Then
        '1�̃Z����������擾����ꍇ�͂��̒l��Ԃ�
        strRC = "R" & Rs & "C" & Cs
        Output = ExecuteExcel4Macro("'" & BookFolderPath & "\[" & BookName & "]" & SheetName & "'!" & strRC)
    Else
        '�����Z������擾����ꍇ�͔z��ŕԂ�
        ReDim Output(1 To Re - Rs + 1, 1 To Ce - Cs + 1)
        
        For I = Rs To Re
            For J = Cs To Ce
                strRC = "R" & I & "C" & J
                Output(I, J) = ExecuteExcel4Macro("'" & BookFolderPath & "\[" & BookName & "]" & SheetName & "'!" & strRC)
            Next J
        Next I
    End If
    
    InputBook = Output
    
End Function

Private Sub TestSelectFile()
'SelectFile�̎��s�T���v��
'20210720

    Dim FolderPath$
    Dim strFileName$
    Dim strExtentions$
    FolderPath = "" 'ActiveWorkbook.Path
    strFileName = "" '"Excel�u�b�N"   '����������������������������������������������
    strExtentions = "" '"*.xls; *.xlsx; *.xlsm" '����������������������������������������������
    
    Dim FilePath$
    FilePath = SelectFile(FolderPath, strFileName, strExtentions)
    
End Sub

Function SelectFile(Optional FolderPath$, Optional strFileName$ = "", Optional strExtentions$ = "")
'�t�@�C����I������_�C�A���O��\�����ăt�@�C����I��������
'�I�������t�@�C���̃t���p�X��Ԃ�
'20210720

'FolderPath�E�E�E�ŏ��ɊJ���t�H���_ �w�肵�Ȃ��ꍇ�̓J�����g�t�H���_�p�X
'strFileName�E�E�E�I������t�@�C���̖��O  ��FExcel�u�b�N
'strExtentions�E�E�E�I������t�@�C���̊g���q�@��F"*.xls; *.xlsx; *.xlsm"

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir '�J�����g�t�H���_
    End If
    
    Dim Output$
    
    With FD
        With .Filters
            .Clear
            .Add strFileName, strExtentions, 1
        End With
        .InitialFileName = FolderPath & "\"
        If .Show = True Then
            Output = .SelectedItems(1)
        Else
            MsgBox ("�t�@�C�����I������Ȃ������̂ŏI�����܂�")
            End
        End If
    End With
    
    SelectFile = Output
    
End Function

Private Sub TestSelectFolder()
'SelectFolder�̎��s�T���v��
'20210720

    Dim FolderPath$
    FolderPath = ActiveWorkbook.Path
    
    Dim FilePath$
    FilePath = SelectFolder(FolderPath)
    
End Sub

Function SelectFolder(Optional FolderPath$)
'�t�H���_��I������_�C�A���O��\�����ăt�@�C����I��������
'�I�������t�H���_�̃t���p�X��Ԃ�
'20210720

'FolderPath�E�E�E�ŏ��ɊJ���t�H���_ �w�肵�Ȃ��ꍇ�̓J�����g�t�H���_�p�X

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir '�J�����g�t�H���_
    End If
    
    Dim Output$
    
    With FD
        With .Filters
            .Clear
        End With
        .InitialFileName = FolderPath & "\"
        If .Show = True Then
            Output = .SelectedItems(1)
        Else
            MsgBox ("�t�H���_���I������Ȃ������̂ŏI�����܂�")
            End
        End If
    End With
    
    SelectFolder = Output
    
End Function

Function GetFileDateTime(FilePath$)
'�t�@�C���̃^�C���X�^���v���擾����B
'�֐��v���o���p
'20210720

'FilePath�E�E�E�^�C���X�^���v���擾����t�@�C���̃t���p�X

    GetFileDateTime = FileDateTime(FilePath)
    
End Function

Sub MakeFolder(FolderPath$)
'�t�H���_���쐬����
'20210720

'FilePath�E�E�E�쐬����t�H���_�̃t���p�X

    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
End Sub

Sub TestGetRowCountTextFile()
    
    Dim FilePath$
    FilePath = ActiveWorkbook.Path & "\" & "TestText.txt"
    
    Dim RowCount&
    RowCount = GetRowCountTextFile(FilePath)
    
End Sub

Function GetRowCountTextFile(FilePath$)
'�e�L�X�g�t�@�C���ACSV�t�@�C���̍s�����擾����
'20210720

    '�t�@�C���̑��݊m�F
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("�u" & FilePath & "�v������܂���" & vbLf & _
                "�I�����܂�")
        End
    End If
    
    Dim Output&
    With CreateObject("Scripting.FileSystemObject")
        Output = .OpenTextFile(FilePath, 8).Line
    End With
    
    GetRowCountTextFile = Output
    
End Function

Function GetCurrentFolder()
'�J�����g�t�H���_�̃p�X���擾
'�֐��v���o���p
'20210720

    GetCurrentFolder = CurDir
    
End Function

Sub SetCurrentFolder(FolderPath$)
'�w��t�H���_�p�X���J�����g�t�H���_��ݒ�
'�t�H���_�p�X���l�b�g���[�N�h���C�u��̃t�H���_�������I�ɔ��肵��
'�l�b�g���[�N�h���C�u��̃t�H���_���J�����g�t�H���_�ɐݒ�ł���
'20210720

    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("�u" & FolderPath & "�v������܂���" & vbLf & _
                "�I�����܂�")
        End
    End If
    
    If Mid(FolderPath, 1, 2) = "\\" Then
        '�l�b�g���[�N�h���C�u�̏ꍇ
        Call SetCurrentFolderNetworkDrive(FolderPath)
    Else
        
        '�J�����g�h���C�u���قȂ�ꍇ�͐�ɐݒ肷��K�v������
        If Mid(FolderPath, 1, 1) <> Mid(CurDir, 1, 1) Then
            ChDrive Mid(FolderPath, 1, 1)
        End If
        
        '�J�����g�t�H���_�ݒ�
        ChDir FolderPath
    End If
    
End Sub

Sub SetCurrentFolderNetworkDrive(NetworkFolderPath$)
'�l�b�g���[�N�h���C�u��̃t�H���_�p�X���J�����g�t�H���_�ɐݒ肷��
'20210720

    With CreateObject("WScript.Shell")
        .CurrentDirectory = NetworkFolderPath
    End With
    
End Sub

Private Sub TestGetExtension()
    
    Dim Dummy
    Dummy = GetExtension(ActiveWorkbook.Path & "\" & ActiveWorkbook.Name)
    
End Sub

Function GetExtension(FilePath$)
'�t�@�C���̊g���q���擾����
'20210720

    Dim Output$
    With CreateObject("Scripting.FileSystemObject")
        Output = .GetExtensionName(FilePath)
    End With
    GetExtension = Output
    
End Function

Sub OpenFolder(FolderPath$)
'�w��p�X�̃t�H���_���N������B
'20210721
    
    Shell "C:\Windows\explorer.exe " & FolderPath, vbNormalFocus

End Sub

Sub OpenFile(FilePath$)
'�w��p�X�̃t�@�C�����N������B
'20210726
    
    Dim WSH As Object
    Set WSH = CreateObject("WScript.Shell")
    WSH.Run FilePath
    
End Sub

Sub OpenApplication(ApplicationPath$)
'�w��p�X�̃A�v�����N������
'��)�d��Ȃ�"calc.exe"�Ȃ�
'20210726
    
    Shell ApplicationPath, vbNormalFocus

End Sub

Sub TestOutputCSV()

    Dim FolderPath$, FileName$, OutputHairetu
    FolderPath = ActiveWorkbook.Path
    FileName = "Test"
    OutputHairetu = Range("B3:I1832").Value
    Call OutputCSV(FolderPath, FileName, OutputHairetu)

End Sub

Sub OutputCSV(FolderPath$, FileName$, ByVal OutputHairetu)
'�w��z���CSV�ŏo�͂���
'20210721

'FolderPath�E�E�E�o�͐�̃t�H���_�p�X
'FileName�E�E�E�o�͂���t�@�C�����i�g���q�͕t���Ȃ��j
'OutputHairetu�E�E�E�o�͂���z��

    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    '1�����z���2�����z��ɕϊ�
    OutputHairetu = Lib���͔z��������p�ɕϊ�(OutputHairetu)
    
    N = UBound(OutputHairetu, 1)
    M = UBound(OutputHairetu, 2)
    Dim fp
    
    ' FreeFile�l�̎擾(�ȍ~���̒l�œ��o�͂���)
    fp = FreeFile
    ' �w��t�@�C����OPEN(�o�̓��[�h)
    Open FolderPath & "\" & FileName & ".csv" For Output As #fp
    ' �ŏI�s�܂ŌJ��Ԃ�
    
    For I = 1 To N
        For J = 1 To M - 1
            ' ���R�[�h���o��
            Print #fp, OutputHairetu(I, J) & ",";
        Next J
        Print #fp, OutputHairetu(I, M)
    Next I
    ' �w��t�@�C����CLOSE
    Close fp

End Sub

Sub TestOutputText()

    Dim FolderPath$, FileName$, OutputHairetu
    FolderPath = ActiveWorkbook.Path
    FileName = "Test"
    OutputHairetu = Range("B3:I1832").Value
    Call OutputText(FolderPath, FileName, OutputHairetu, Chr(9))

End Sub

Sub OutputText(FolderPath$, FileName$, ByVal OutputHairetu, Optional KugiriMoji$ = ",")
'�w��z���txt�ŏo�͂���
'20210721
   
'FolderPath�E�E�E�o�͐�̃t�H���_�p�X
'FileName�E�E�E�o�͂���t�@�C�����i�g���q�͂���j
'OutputHairetu�E�E�E�o�͂���z��
'KugiriMoji�E�E�E������Ԃ̋�؂蕶��

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    '1�����z���2�����z��ɕϊ�
    OutputHairetu = Lib���͔z��������p�ɕϊ�(OutputHairetu)
    
    N = UBound(OutputHairetu, 1)
    M = UBound(OutputHairetu, 2)
    Dim fp
    
    ' FreeFile�l�̎擾(�ȍ~���̒l�œ��o�͂���)
    fp = FreeFile
    ' �w��t�@�C����OPEN(�o�̓��[�h)
    Open FolderPath & "\" & FileName For Output As #fp
    ' �ŏI�s�܂ŌJ��Ԃ�
    
    For I = 1 To N
        For J = 1 To M - 1
            ' ���R�[�h���o��
            Print #fp, OutputHairetu(I, J) & KugiriMoji;
        Next J
        
        If I < N Then
            Print #fp, OutputHairetu(I, M)
        Else
            Print #fp, OutputHairetu(I, M);
        End If
    Next I
    ' �w��t�@�C����CLOSE
    Close fp

End Sub

Function InputText(FilePath$, Optional KugiriMoji$ = "")
'�e�L�X�g�t�@�C����ǂݍ���Ŕz��ŕԂ�
'�����R�[�h�͎����I�ɔ��肵�ēǍ��`����ύX����
'20210721

'FilePath�E�E�E�e�L�X�g�t�@�C���̃t���p�X
'KugiriMoji�E�E�E�e�L�X�g�t�@�C����ǂݍ���ŋ�؂蕶���ŋ�؂��Ĕz��ŏo�͂���ꍇ�̋�؂蕶��

    '�e�L�X�g�t�@�C���̑��݊m�F
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("�u" & FilePath & "�v" & vbLf & _
               "�̑��݂��m�F�ł��܂���B" & vbLf & _
               "�������I�����܂��B")
        End
    End If
    
    '�e�L�X�g�t�@�C���̕����R�[�h���擾
    Dim strCode
    strCode = fncGetCharset(FilePath)
    If strCode = "UTF-8 BOM" Or strCode = "UTF-8" Then
        strCode = "UTF-8"
    ElseIf strCode = "UTF-16 LE BOM" Or strCode = "UTF-16 BE BOM" Then
        strCode = "UTF-16LE"
    Else
        strCode = Empty
    End If
    
    '�e�L�X�g�t�@�C���Ǎ�
    Dim Output
    Dim RowCount&
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim FileNo%, Buffer$
    
    If IsEmpty(strCode) = False Then 'UTF8�ł̏ꍇ����������������������������������
   
        Output = InputTextUTF8(FilePath, KugiriMoji)
    
    Else 'Shift-JIS�ł̏ꍇ����������������������������������
        
        Output = InputTextShiftJIS(FilePath, KugiriMoji)
     
    End If

    InputText = Output
    
End Function

Function InputTextShiftJIS(FilePath$, Optional KugiriMoji$ = "")
'�e�L�X�g�t�@�C����ǂݍ��� ShiftJIS�`����p
'20210721

'FilePath�E�E�E�e�L�X�g�t�@�C���̃t���p�X
'KugiriMoji�E�E�E�e�L�X�g�t�@�C����ǂݍ���ŋ�؂蕶���ŋ�؂��Ĕz��ŏo�͂���ꍇ�̋�؂蕶��
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim FileNo%, Buffer$, SplitBuffer
    Dim Output1, Output2
    ' FreeFile�l�̎擾(�ȍ~���̒l�œ��o�͂���)
    FileNo = FreeFile
    
    N = GetRowCountTextFile(FilePath)
    ReDim Output1(1 To N)
    ' �w��t�@�C����OPEN(���̓��[�h)
    Open FilePath For Input As #FileNo
            
    ' �t�@�C����EOF(End of File)�܂ŌJ��Ԃ�
    I = 0
    M = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, Buffer
        I = I + 1
        Output1(I) = Buffer '1���Ǎ���
        
        If KugiriMoji <> "" Then '�����ŋ�؂�ꍇ�͋�؂�����v�Z
            '��؂蕶���ɂ���؂���̍ő�l����ɍX�V���Ă���
            M = WorksheetFunction.Max(M, UBound(Split(Buffer, KugiriMoji)) + 1)
        End If
    Loop
    
    Close #FileNo
    
    '��؂蕶���̏���
    If KugiriMoji = "" Then
        '��؂蕶���Ȃ�
        Output2 = Output1
    Else
        ReDim Output2(1 To N, 1 To M)
        For I = 1 To N
            Buffer = Output1(I)
            SplitBuffer = Split(Buffer, KugiriMoji)
            For J = 0 To UBound(SplitBuffer)
                Output2(I, J + 1) = SplitBuffer(J)
            Next J
        Next I
    End If
    
    InputTextShiftJIS = Output2

End Function

Function InputTextUTF8(FilePath$, Optional KugiriMoji$ = "")
'�e�L�X�g�t�@�C����ǂݍ��� UTF8�`����p
'20210721

'FilePath�E�E�E�e�L�X�g�t�@�C���̃t���p�X
'KugiriMoji�E�E�E�e�L�X�g�t�@�C����ǂݍ���ŋ�؂蕶���ŋ�؂��Ĕz��ŏo�͂���ꍇ�̋�؂蕶��

    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim Buffer$, SplitBuffer
    Dim Output1, Output2
    
    N = GetRowCountTextFile(FilePath)
    ReDim Output1(1 To N)
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2 ' �t�@�C���̃^�C�v(1:�o�C�i�� 2:�e�L�X�g)
        .Open
        .LineSeparator = 10 '���s�R�[�h
        .LoadFromFile (FilePath)
        
        For I = 1 To N
            Buffer = .ReadText(-2)
            Output1(I) = Buffer
            If KugiriMoji <> "" Then '�����ŋ�؂�ꍇ�͋�؂�����v�Z
                '��؂蕶���ɂ���؂���̍ő�l����ɍX�V���Ă���
                M = WorksheetFunction.Max(M, UBound(Split(Buffer, KugiriMoji)) + 1)
            End If
        Next I
        .Close
    End With
    
    '��؂蕶���̏���
    If KugiriMoji = "" Then
        '��؂蕶���Ȃ�
        Output2 = Output1
    Else
        ReDim Output2(1 To N, 1 To M)
        For I = 1 To N
            Buffer = Output1(I)
            SplitBuffer = Split(Buffer, KugiriMoji)
            For J = 0 To UBound(SplitBuffer)
                Output2(I, J + 1) = SplitBuffer(J)
            Next J
        Next I
    End If
    
    InputTextUTF8 = Output2
    
End Function

'fncGetCharset Ver1.4
Function fncGetCharset(FileName As String) As String
'20200909�ǉ�
'�e�L�X�g�t�@�C���̕����R�[�h��Ԃ�
'�Q�lhttps://popozure.info/20190515/14201

    Dim I                   As Long
    
    Dim hdlFile             As Long
    Dim lngFileLen          As Long
    
    Dim bytFile()           As Byte
    Dim b1                  As Byte
    Dim b2                  As Byte
    Dim b3                  As Byte
    Dim b4                  As Byte
    
    Dim lngSJIS             As Long
    Dim lngUTF8             As Long
    Dim lngEUC              As Long
    
    On Error Resume Next
    
    '�t�@�C���ǂݍ���
    lngFileLen = FileLen(FileName)
    ReDim bytFile(lngFileLen)
    If (Err.Number <> 0) Then
        Exit Function
    End If
    
    hdlFile = FreeFile()
    Open FileName For Binary As #hdlFile
    Get #hdlFile, , bytFile
    Close #hdlFile
    If (Err.Number <> 0) Then
        Exit Function
    End If
    
    'BOM�ɂ�锻�f
    If (bytFile(0) = &HEF And bytFile(1) = &HBB And bytFile(2) = &HBF) Then
        fncGetCharset = "UTF-8 BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFF And bytFile(1) = &HFE) Then
        fncGetCharset = "UTF-16 LE BOM"
        Exit Function
    ElseIf (bytFile(0) = &HFE And bytFile(1) = &HFF) Then
        fncGetCharset = "UTF-16 BE BOM"
        Exit Function
    End If
    
    'BINARY
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 >= &H0 And b1 <= &H8) Or (b1 >= &HA And b1 <= &H9) Or (b1 >= &HB And b1 <= &HC) Or (b1 >= &HE And b1 <= &H19) Or (b1 >= &H1C And b1 <= &H1F) Or (b1 = &H7F) Then
            fncGetCharset = "BINARY"
            Exit Function
        End If
    Next I
           
    'SJIS
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Or (b1 >= &HB0 And b1 <= &HDF) Then
            lngSJIS = lngSJIS + 1
        Else
            If (I < lngFileLen - 2) Then
                b2 = bytFile(I + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                   ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   lngSJIS = lngSJIS + 2
                   I = I + 1
                End If
            End If
        End If
    Next I
           
    'UTF-8
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 = &H9) Or (b1 = &HA) Or (b1 = &HD) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngUTF8 = lngUTF8 + 1
        Else
            If (I < lngFileLen - 2) Then
                b2 = bytFile(I + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   lngUTF8 = lngUTF8 + 2
                   I = I + 1
                Else
                    If (I < lngFileLen - 3) Then
                        b3 = bytFile(I + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) Then
                            lngUTF8 = lngUTF8 + 3
                            I = I + 2
                        Else
                            If (I < lngFileLen - 4) Then
                                b4 = bytFile(I + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And (b2 >= &H80 And b2 <= &HBF) And (b3 >= &H80 And b3 <= &HBF) And (b4 >= &H80 And b3 <= &HBF) Then
                                    lngUTF8 = lngUTF8 + 4
                                    I = I + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next I

    'EUC-JP
    For I = 0 To lngFileLen - 1
        b1 = bytFile(I)
        If (b1 = &H7) Or (b1 = 10) Or (b1 = 13) Or (b1 >= &H20 And b1 <= &H7E) Then
            lngEUC = lngEUC + 1
        Else
            If (I < lngFileLen - 2) Then
                b2 = bytFile(I + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And _
                   (b2 >= &HA1 And b2 <= &HFE)) Or _
                   ((b1 = &H8E) And (b2 >= &HA1 And b2 <= &HDF)) Then
                   lngEUC = lngEUC + 2
                   I = I + 1
                End If
            End If
        End If
    Next I
           
    '�����R�[�h�o�����ʂɂ�锻�f
    If (lngSJIS <= lngUTF8) And (lngEUC <= lngUTF8) Then
        fncGetCharset = "UTF-8"
        Exit Function
    End If
    If (lngUTF8 <= lngSJIS) And (lngEUC <= lngSJIS) Then
        fncGetCharset = "Shift_JIS"
        Exit Function
    End If
    If (lngUTF8 <= lngEUC) And (lngSJIS <= lngEUC) Then
        fncGetCharset = "EUC-JP"
        Exit Function
    End If
    fncGetCharset = ""
    
End Function

Function GetFiles(FolderPath$, ParamArray Extensions())
'�t�H���_���̃t�@�C���̃��X�g���擾����
'20210721

'�uMicrosoft Scripting Runtime�v���C�u�������Q�Ƃ��邱��

'FolderPath�E�E�E�����Ώۂ̃t�H���_�p�X
'Extensions�E�E�E�擾�Ώۂ̊g���q�A�ϒ������z��œ��͉\

    '�t�H���_�̊m�F
    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("�u" & FolderPath & "�v" & vbLf & _
               "�̃t�H���_�̑��݂��m�F�ł��܂���B" & vbLf & _
               "�������I�����܂��B")
    End If
    
    '�g���q�̘A�z�z����쐬
    Dim ExtensionDict As Object, TmpExtension
    
    If IsMissing(Extensions) <> True Then
        '�g���q�����͂���Ă���ꍇ
        Set ExtensionDict = CreateObject("Scripting.Dictionary")
        For Each TmpExtension In Extensions
            TmpExtension = StrConv(TmpExtension, vbLowerCase)
            ExtensionDict.Add TmpExtension, ""
        Next
    End If
    
    Dim FSO As New FileSystemObject
    Dim myFolder As Folder
    Dim myFiles As Files
    Dim TmpFile As File, TmpFileName$
    Set myFolder = FSO.GetFolder(FolderPath)
    Set myFiles = myFolder.Files
    
    If myFiles.Count = 0 Then
        '�t�@�C������
        Exit Function
    End If
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    ReDim Output(1 To 1)
    
    If IsMissing(Extensions) = True Then
        N = myFiles.Count
        ReDim Output(1 To N)
    End If
    
    K = 0
    For Each TmpFile In myFiles
        TmpFileName = TmpFile.Name
        
        If IsMissing(Extensions) <> True Then
            TmpExtension = StrConv(FSO.GetExtensionName(FolderPath & "\" & TmpFileName), vbLowerCase)
            If ExtensionDict.Exists(TmpExtension) Then
                K = K + 1
                ReDim Preserve Output(1 To K)
                Output(K) = TmpFileName
            End If
        Else
            K = K + 1
            Output(K) = TmpFileName
        End If
    Next
    
    GetFiles = Output
    
End Function

Function GetSubFolders(FolderPath$)
'�t�H���_���̃T�u�t�H���_�̃��X�g���擾����
'20210721

'FolderPath�E�E�E�����Ώۂ̃t�H���_�p�X

    '�t�H���_�̊m�F
    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("�u" & FolderPath & "�v" & vbLf & _
               "�̃t�H���_�̑��݂��m�F�ł��܂���B" & vbLf & _
               "�������I�����܂��B")
    End If
    
    '�g���q�̘A�z�z����쐬
    Dim ExtensionDict As Object, TmpExtension
    
    Dim FSO As New FileSystemObject
    Dim myFolder As Folder
    Dim mySubFolder As Folders, TmpSubFolder As Folder
    Set myFolder = FSO.GetFolder(FolderPath)
    Set mySubFolder = myFolder.SubFolders
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    N = mySubFolder.Count
    
    If N = 0 Then
        '�T�u�t�H���_����
        Exit Function
    End If
    ReDim Output(1 To N)
    
    K = 0
    For Each TmpSubFolder In mySubFolder
       K = K + 1
       Output(K) = TmpSubFolder.Name
    Next
    
    GetSubFolders = Output
    
End Function

Sub OutputPDF(TargetSheet As Worksheet, Optional FolderPath$, Optional FileName$, _
              Optional MessageIruNaraTrue As Boolean = True)
'�w��V�[�g��PDF������
'20210721

'TargetSheet�E�E�EPDF������Ώۂ̃V�[�g
'FolderPath�E�E�E�o�͐�t�H���_ �w�肵�Ȃ��ꍇ�̓u�b�N�Ɠ����t�H���_
'FileName�E�E�E�o��PDF�̃t�@�C���� �w�肵�Ȃ��ꍇ�̓V�[�g�̖��O

    If FolderPath = "" Then
        FolderPath = TargetSheet.Parent.Path
    End If
    
    If FileName = "" Then
        FileName = TargetSheet.Name
    End If
    
    '�o�͐�t�H���_���Ȃ��ꍇ�͍쐬����B
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    Dim OutputFileName$
    
    OutputFileName = FolderPath & "\" & FileName & ".pdf"

    TargetSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=OutputFileName
    
    If MessageIruNaraTrue Then
        If MsgBox("�u" & FileName & ".pdf" & "�v" & vbLf & "���쐬���܂���" & vbLf & _
            "�o�͐�t�H���_���N�����܂���?", vbYesNo + vbQuestion) = vbYes Then
            Shell "C:\Windows\explorer.exe " & FolderPath, vbNormalFocus
        End If
    End If
    
End Sub

Sub OutputXML(Title$, InputList, TateTableList, YokoTableList, _
              Optional TateTableName$ = "DATA", Optional YokoTableName$ = "ID", _
              Optional FolderPath$, Optional FileName$)
              
'�e�[�u���f�[�^����XML�f�[�^���o�͂���
'�uMicrosoft XML, v6.0�v���C�u�������Q�Ƃ��邱��
'�Q�l�Fhttp://www.openreference.org/articles/view/651
'20210727
'20210824����

'Title          :�^�C�g��
'InputList      :XML�f�[�^���̃f�[�^���X�g�i2�����z��j
'TateTableList  :�c�����̃e�[�u�����̃��X�g�i1�����z��j
'YokoTableList  :�������̃e�[�u�����̃��X�g�i1�����z��j
'TateTableName  :�c�����̃e�[�u�����B�f�t�H���g��"DATA"
'YokoTableName  :�������̃e�[�u�����B�f�t�H���g��"ID"
'FolderPath     :XML�f�[�^�̏o�͂����̃t�H���_�p�X�B�f�t�H���g�͎�Excel�u�b�N�̃p�X
'FileName       :XML�f�[�^���o�͂���t�@�C�����i�g���q"xml"�͊܂܂��j�@�f�t�H���g�̓^�C�g��(Title)�Ɠ���

    '�����̃`�F�b�N
    Call CheckArray1D(TateTableList, "TateTableList")
    Call CheckArray1D(YokoTableList, "YokoTableList")
    Call CheckArray2D(InputList, "InputList")
    Call CheckArray1DStart1(TateTableList, "TateTableList")
    Call CheckArray1DStart1(YokoTableList, "YokoTableList")
    Call CheckArray2DStart1(InputList, "InputList")
    
    If UBound(TateTableList, 1) <> UBound(InputList, 1) Then
        MsgBox ("�uTateTableList�v�̗v�f����" & vbLf & _
                "�uInputList�v�̏c�v�f������v�����Ă�������")
        Stop
        End
    End If
    
    If UBound(YokoTableList, 1) <> UBound(InputList, 2) Then
        MsgBox ("�uYokoTableList�v�̗v�f����" & vbLf & _
                "�uInputList�v�̉��v�f������v�����Ă�������")
        Stop
        End
    End If
    
    '�����̃f�t�H���g�l�ݒ�
    If FolderPath = "" Then
        FolderPath = ThisWorkbook.Path
    End If
    
    If FileName = "" Then
        FileName = Title
    End If
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(InputList, 1)
    M = UBound(InputList, 2)
    
    Dim XMLDoc As New MSXML2.DOMDocument60
    Dim xmlRoot As IXMLDOMNode
    Dim xmlData As IXMLDOMNode
    Dim xmlChildData As IXMLDOMNode
    Dim xmlAttr As MSXML2.IXMLDOMAttribute

    With XMLDoc
        'XML�錾�𐶐�
        Call .appendChild(.createProcessingInstruction("xml", "version=""1.0"" encoding=""Shift_JIS"""))
        
        '�v�f�𐶐�
        Set xmlRoot = .appendChild(.createElement(Title))
        
        For I = 1 To N
            
            '�v�f�𐶐�
            Set xmlData = .createElement(YokoTableName)      '���v�f�𐶐�
            Set xmlAttr = .createAttribute(TateTableName)    '�c�v�f�𐶐�
            xmlAttr.NodeValue = TateTableList(I)             '���v�f�̒l��ݒ�
            Call xmlData.Attributes.setNamedItem(xmlAttr)    '�v�f��id������ݒ�
            
            '�v�f�̎q�v�f�𐶐����ėv�f�ɒǉ�
            For J = 1 To M
                Set xmlChildData = xmlData.appendChild(.createElement(YokoTableList(J)))
                xmlChildData.Text = InputList(I, J)
            Next J
            
            Call xmlRoot.appendChild(xmlData)
        Next I
        
        'XML�h�L�������g���o��
        .Save (FolderPath & "\" & FileName & ".xml")
    End With

End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Function GetFileName$(FilePath$)
'�t�@�C���̃t���p�X����t�@�C�����擾
'�֐��v���o���p
'20210824
    
    Dim Output$
    Dim TmpList
    TmpList = Split(FilePath, "\")
    Output = TmpList(UBound(TmpList))
    GetFileName = Output
    
End Function

Private Function Lib���͔z��������p�ɕϊ�(InputHairetu)
'���͂����z��������p�ɕϊ�����
'1�����z��2�����z��
'���l��������2�����z��(1,1)
'�v�f�̊J�n�ԍ���1�ɂ���
'20210721

    Dim Output
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Base1%, Base2%
    If IsArray(InputHairetu) = False Then
        '�z��łȂ��ꍇ(���l��������)
        ReDim Output(1 To 1, 1 To 1)
        Output(1, 1) = InputHairetu
    Else
        On Error Resume Next
        M = UBound(InputHairetu, 2)
        On Error GoTo 0
        If M = 0 Then
            '1�����z��
            Output = WorksheetFunction.Transpose(InputHairetu)
        Else
            '2�����z��
            Base1 = LBound(InputHairetu, 1)
            Base2 = LBound(InputHairetu, 2)
            
            If Base1 <> 1 Or Base2 <> 1 Then
                N = UBound(InputHairetu, 1)
                If N = Base1 Then
                    '(1,M)�z��
                    ReDim Output(1 To 1, 1 To M - Base2 + 1)
                    For I = 1 To M - Base2 + 1
                        Output(1, I) = InputHairetu(Base1, Base2 + I - 1)
                    Next I
                Else
                    Output = WorksheetFunction.Transpose(InputHairetu)
                    Output = WorksheetFunction.Transpose(Output)
                End If
            Else
                Output = InputHairetu
            End If
        End If
    End If
    
    Lib���͔z��������p�ɕϊ� = Output
    
End Function



