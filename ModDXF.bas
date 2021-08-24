Attribute VB_Name = "ModDXF"
Option Explicit

Sub DXF生成テスト()
    
    Dim XYList
    XYList = Sheet1.Range("C3:D23").Value
    
    Dim FilePath$
    FilePath = ThisWorkbook.Path & "\" & "DXFTestByVBA.dxf"
    
    Call OutputDXFLine(XYList, FilePath)
    
End Sub

Sub OutputDXFLine(InputArray2D, FilePath$)
'二次元配列からDXFファイルを作成する

    Call CheckArray2D(InputArray2D)
    Call CheckArray2DStart1(InputArray2D)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(InputArray2D, 1)
    
    Dim RowCount&
    RowCount = (N - 1) * (8 + 4) + 5 + 3
    
    Dim Output
    ReDim Output(1 To RowCount)
    
    '冒頭の文
    Output(1) = "  0"
    Output(2) = "SECTION"
    Output(3) = "  2"
    Output(4) = "ENTITIES"
    Output(5) = "  0"
    
    '終了の文
    Output(RowCount - 2) = "ENDSEC"
    Output(RowCount - 1) = "  0"
    Output(RowCount) = "EOF"
    
    Dim StartNum&
    StartNum = 5 '←←←←←←←←←←←←←←←←←←←←←←←
    
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
        Output(K + 3) = "0" 'レイヤー名
                
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
