Attribute VB_Name = "ModOutputDXFLine"
Option Explicit

'OutputDXFLine	���ꏊ�FFukamiAddins3.ModDXF
'OutputText	���ꏊ�FFukamiAddins3.ModFile
'Lib���͔z��������p�ɕϊ�	���ꏊ�FFukamiAddins3.ModFile
'GetFileName	���ꏊ�FFukamiAddins3.ModFile
'CheckArray2D	���ꏊ�FFukamiAddins3.ModArray
'CheckArray2DStart1	���ꏊ�FFukamiAddins3.ModArray


'------------------------------


'------------------------------

'------------------------------


'�z��̏����֌W�̃v���V�[�W��

'------------------------------


Sub OutputDXFLine(InputArray2D, FilePath$)
'�񎟌��z�񂩂�DXF�t�@�C�����쐬����

    Call CheckArray2D(InputArray2D)
    Call CheckArray2DStart1(InputArray2D)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(InputArray2D, 1)
    
    Dim RowCount&
    RowCount = (N - 1) * (8 + 4) + 5 + 3
    
    Dim Output
    ReDim Output(1 To RowCount)
    
    '�`���̕�
    Output(1) = "  0"
    Output(2) = "SECTION"
    Output(3) = "  2"
    Output(4) = "ENTITIES"
    Output(5) = "  0"
    
    '�I���̕�
    Output(RowCount - 2) = "ENDSEC"
    Output(RowCount - 1) = "  0"
    Output(RowCount) = "EOF"
    
    Dim StartNum&
    StartNum = 5 '����������������������������������������������
    
    Dim StartX#, StartY#, EndX#, EndY#
    
    K = StartNum
    For I = 1 To N - 1
        If I <> 1 Then
            K = K + 12
        End If
        
        StartX = InputArray2D(I, 1)
        StartY = InputArray2D(I, 2)
        EndX = InputArray2D(I + 1, 1)
        EndY = InputArray2D(I + 1, 2)
        
        Output(K + 1) = "LINE"
        Output(K + 2) = "  8"
        Output(K + 3) = "0" '���C���[��
                
        Output(K + 4) = " 10"
        Output(K + 5) = Format(StartX, "0.000000")
        Output(K + 6) = " 20"
        Output(K + 7) = Format(StartY, "0.000000")
        Output(K + 8) = " 11"
        Output(K + 9) = Format(EndX, "0.000000")
        Output(K + 10) = " 21"
        Output(K + 11) = Format(EndY, "0.000000")
        
        Output(K + 12) = "  0"
    
    Next I
    
    Dim TmpFilePath$, TmpFileName$
    TmpFileName = GetFileName(FilePath)
    TmpFilePath = Replace(FilePath, "\" & TmpFileName, "")
        
    Call OutputText(TmpFilePath, TmpFileName, Output)

End Sub

Private Sub OutputText(FolderPath$, FileName$, ByVal OutputHairetu, Optional KugiriMoji$ = ",")
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

Private Function GetFileName$(FilePath$)
'�t�@�C���̃t���p�X����t�@�C�����擾
'�֐��v���o���p
'20210824
    
    Dim Output$
    Dim TmpList
    TmpList = Split(FilePath, "\")
    Output = TmpList(UBound(TmpList))
    GetFileName = Output
    
End Function

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

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

