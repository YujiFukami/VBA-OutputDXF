Attribute VB_Name = "ModOutputDXFLine"
Option Explicit

'OutputDXFLine            ・・・元場所：FukamiAddins3.ModDXF  
'CheckArray2D             ・・・元場所：FukamiAddins3.ModArray
'CheckArray2DStart1       ・・・元場所：FukamiAddins3.ModArray
'GetFileName              ・・・元場所：FukamiAddins3.ModFile 
'OutputText               ・・・元場所：FukamiAddins3.ModFile 
'Lib入力配列を処理用に変換・・・元場所：FukamiAddins3.ModFile 



Public Sub OutputDXFLine(InputArray2D, FilePath As String)
'二次元配列からDXFファイルを作成する

    Call CheckArray2D(InputArray2D)
    Call CheckArray2DStart1(InputArray2D)
    
    Dim I        As Long
    Dim K        As Long
    Dim N        As Long
    Dim RowCount As Long
    Dim Output
    N = UBound(InputArray2D, 1)
    RowCount = (N - 1) * (8 + 4) + 5 + 3
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
    
    Dim StartNum As Long
    StartNum = 5 '←←←←←←←←←←←←←←←←←←←←←←←
    
    Dim StartX As Double
    Dim StartY As Double
    Dim EndX   As Double
    Dim EndY   As Double
    
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
    
    Dim TmpFilePath As String
    Dim TmpFileName As String
    TmpFileName = GetFileName(FilePath)
    TmpFilePath = Replace(FilePath, "\" & TmpFileName, "")
        
    Call OutputText(TmpFilePath, TmpFileName, Output)

End Sub

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function GetFileName(FilePath As String)
'ファイルのフルパスからファイル名取得
'関数思い出し用
'20210824
    
    Dim Output As String
    Dim TmpList
    TmpList = Split(FilePath, "\")
    Output = TmpList(UBound(TmpList))
    GetFileName = Output
    
End Function

Private Sub OutputText(FolderPath As String, FileName As String, ByVal OutputHairetu, Optional KugiriMoji As String = ",")
'指定配列をtxtで出力する
'20210721
   
'FolderPath   ・・・出力先のフォルダパス
'FileName     ・・・出力するファイル名（拡張子はつける）
'OutputHairetu・・・出力する配列
'KugiriMoji   ・・・文字列間の区切り文字

    Dim I As Integer
    Dim J As Integer
    Dim M As Integer
    Dim N As Integer
    
    '1次元配列は2次元配列に変換
    OutputHairetu = Lib入力配列を処理用に変換(OutputHairetu)
    
    N = UBound(OutputHairetu, 1)
    M = UBound(OutputHairetu, 2)
    Dim fp
    
    ' FreeFile値の取得(以降この値で入出力する)
    fp = FreeFile
    ' 指定ファイルをOPEN(出力モード)
    Open FolderPath & "\" & FileName For Output As #fp
    ' 最終行まで繰り返す
    
    For I = 1 To N
        For J = 1 To M - 1
            ' レコードを出力
            Print #fp, OutputHairetu(I, J) & KugiriMoji;
        Next J
        
        If I < N Then
            Print #fp, OutputHairetu(I, M)
        Else
            Print #fp, OutputHairetu(I, M);
        End If
    Next I
    ' 指定ファイルをCLOSE
    Close fp

End Sub

Private Function Lib入力配列を処理用に変換(InputHairetu)
'入力した配列を処理用に変換する
'1次元配列→2次元配列
'数値か文字列→2次元配列(1,1)
'要素の開始番号を1にする
'20210721

    Dim Output
    Dim I     As Integer
    Dim M     As Integer
    Dim N     As Integer
    Dim Base1 As Integer
    Dim Base2 As Integer
    If IsArray(InputHairetu) = False Then
        '配列でない場合(数値か文字列)
        ReDim Output(1 To 1, 1 To 1)
        Output(1, 1) = InputHairetu
    Else
        On Error Resume Next
        M = UBound(InputHairetu, 2)
        On Error GoTo 0
        If M = 0 Then
            '1次元配列
            Output = WorksheetFunction.Transpose(InputHairetu)
        Else
            '2次元配列
            Base1 = LBound(InputHairetu, 1)
            Base2 = LBound(InputHairetu, 2)
            
            If Base1 <> 1 Or Base2 <> 1 Then
                N = UBound(InputHairetu, 1)
                If N = Base1 Then
                    '(1,M)配列
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
    
    Lib入力配列を処理用に変換 = Output
    
End Function


