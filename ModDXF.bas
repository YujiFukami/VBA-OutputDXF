Attribute VB_Name = "ModDXF"
Option Explicit

Sub DXF�����e�X�g()
    
    Dim XYList
    XYList = Sheet1.Range("C3:D23").Value
    
    Dim FilePath$
    FilePath = ThisWorkbook.Path & "\" & "DXFTestByVBA.dxf"
    
    Call OutputDXFLine(XYList, FilePath)
    
End Sub

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

