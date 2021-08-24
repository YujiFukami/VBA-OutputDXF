Attribute VB_Name = "ModFile"
Option Explicit
Sub SaveSheetAsBook(TargetSheet As Worksheet, Optional SaveName$, Optional SavePath$, _
                           Optional MessageIruNaraTrue As Boolean = False)
'指定のシートを別ブックで保存する
'20210719作成
                           
    '入力引数の調整
    If SaveName = "" Then
        SaveName = TargetSheet.Name
    End If
    If SavePath = "" Then
        SavePath = TargetSheet.Parent.Path
    End If
    
    '別ブックで保存
    TargetSheet.Copy
    ActiveWorkbook.SaveAs SavePath & "\" & SaveName
    ActiveWorkbook.Close
    
    If MessageIruNaraTrue Then
        MsgBox ("シート名「" & TargetSheet.Name & "」を" & vbLf & _
               "「" & SavePath & "」に" & vbLf & _
               "ファイル名「" & SaveName & ".xlsx」で保存しました。")
    End If
    
End Sub

Function GetSheetByName(SheetName$) As Worksheet
'指定の名前のシートをワークシートオブジェクトとして取得する
'20210715作成

    Dim Output As Worksheet
    On Error Resume Next
    Set Output = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    
    If Output Is Nothing Then
        MsgBox ("「" & SheetName & "」シートがありません！！")
        End
    End If
    
    Set GetSheetByName = Output

End Function

Function InputCSV(CSVPath$)
'CSVファイルを読み込んで配列形式で返す
'20210706作成

    '入力値確認
    Dim Dummy
    If Dir(CSVPath, vbDirectory) = "" Then
        Dummy = MsgBox(CSVPath & "のファイルは存在しません", vbOKOnly + vbCritical)
        Exit Function
    End If
    
    Dim intFree As Integer
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpStr$, TmpSplit
    Dim StrList
    Dim Output

    intFree = FreeFile '空番号を取得
    Open CSVPath For Input As #intFree 'CSVファィルをオープン
    
    K = 0
    ReDim StrList(1 To 1)
    Do Until EOF(intFree)
        Line Input #intFree, TmpStr '1行読み込み
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
'ブックを開かないでデータを取得する
'ExecuteExcel4Macroを使用するので、Excelのバージョンアップの時に注意
'20210720

'BookFolderPath・・・指定ブックのフォルダパス
'BookName・・・指定ブックの名前 拡張子含む
'SheetName・・・指定ブックの取得対象となるシートの名前
'StartCellAddress・・・取得範囲の最初のセルアドレス(例:"A1")
'EndCellAddress・・・取得範囲の最後のセル(例："B3")（省略ならStartCellAddressと同じ）
    
    Dim Rs&, Re&, Cs&, Ce& '始端行,列番号および終端行,列番号(Long型)
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
    
    '始点、終点の反転している場合の処理
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

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    
    If Rs = Re And Cs = Ce Then
        '1つのセルだけから取得する場合はその値を返す
        strRC = "R" & Rs & "C" & Cs
        Output = ExecuteExcel4Macro("'" & BookFolderPath & "\[" & BookName & "]" & SheetName & "'!" & strRC)
    Else
        '複数セルから取得する場合は配列で返す
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
'SelectFileの実行サンプル
'20210720

    Dim FolderPath$
    Dim strFileName$
    Dim strExtentions$
    FolderPath = "" 'ActiveWorkbook.Path
    strFileName = "" '"Excelブック"   '←←←←←←←←←←←←←←←←←←←←←←←
    strExtentions = "" '"*.xls; *.xlsx; *.xlsm" '←←←←←←←←←←←←←←←←←←←←←←←
    
    Dim FilePath$
    FilePath = SelectFile(FolderPath, strFileName, strExtentions)
    
End Sub

Function SelectFile(Optional FolderPath$, Optional strFileName$ = "", Optional strExtentions$ = "")
'ファイルを選択するダイアログを表示してファイルを選択させる
'選択したファイルのフルパスを返す
'20210720

'FolderPath・・・最初に開くフォルダ 指定しない場合はカレントフォルダパス
'strFileName・・・選択するファイルの名前  例：Excelブック
'strExtentions・・・選択するファイルの拡張子　例："*.xls; *.xlsx; *.xlsm"

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir 'カレントフォルダ
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
            MsgBox ("ファイルが選択されなかったので終了します")
            End
        End If
    End With
    
    SelectFile = Output
    
End Function

Private Sub TestSelectFolder()
'SelectFolderの実行サンプル
'20210720

    Dim FolderPath$
    FolderPath = ActiveWorkbook.Path
    
    Dim FilePath$
    FilePath = SelectFolder(FolderPath)
    
End Sub

Function SelectFolder(Optional FolderPath$)
'フォルダを選択するダイアログを表示してファイルを選択させる
'選択したフォルダのフルパスを返す
'20210720

'FolderPath・・・最初に開くフォルダ 指定しない場合はカレントフォルダパス

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir 'カレントフォルダ
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
            MsgBox ("フォルダが選択されなかったので終了します")
            End
        End If
    End With
    
    SelectFolder = Output
    
End Function

Function GetFileDateTime(FilePath$)
'ファイルのタイムスタンプを取得する。
'関数思い出し用
'20210720

'FilePath・・・タイムスタンプを取得するファイルのフルパス

    GetFileDateTime = FileDateTime(FilePath)
    
End Function

Sub MakeFolder(FolderPath$)
'フォルダを作成する
'20210720

'FilePath・・・作成するフォルダのフルパス

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
'テキストファイル、CSVファイルの行数を取得する
'20210720

    'ファイルの存在確認
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("「" & FilePath & "」がありません" & vbLf & _
                "終了します")
        End
    End If
    
    Dim Output&
    With CreateObject("Scripting.FileSystemObject")
        Output = .OpenTextFile(FilePath, 8).Line
    End With
    
    GetRowCountTextFile = Output
    
End Function

Function GetCurrentFolder()
'カレントフォルダのパスを取得
'関数思い出し用
'20210720

    GetCurrentFolder = CurDir
    
End Function

Sub SetCurrentFolder(FolderPath$)
'指定フォルダパスをカレントフォルダを設定
'フォルダパスがネットワークドライブ上のフォルダか自動的に判定して
'ネットワークドライブ上のフォルダもカレントフォルダに設定できる
'20210720

    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("「" & FolderPath & "」がありません" & vbLf & _
                "終了します")
        End
    End If
    
    If Mid(FolderPath, 1, 2) = "\\" Then
        'ネットワークドライブの場合
        Call SetCurrentFolderNetworkDrive(FolderPath)
    Else
        
        'カレントドライブが異なる場合は先に設定する必要がある
        If Mid(FolderPath, 1, 1) <> Mid(CurDir, 1, 1) Then
            ChDrive Mid(FolderPath, 1, 1)
        End If
        
        'カレントフォルダ設定
        ChDir FolderPath
    End If
    
End Sub

Sub SetCurrentFolderNetworkDrive(NetworkFolderPath$)
'ネットワークドライブ上のフォルダパスをカレントフォルダに設定する
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
'ファイルの拡張子を取得する
'20210720

    Dim Output$
    With CreateObject("Scripting.FileSystemObject")
        Output = .GetExtensionName(FilePath)
    End With
    GetExtension = Output
    
End Function

Sub OpenFolder(FolderPath$)
'指定パスのフォルダを起動する。
'20210721
    
    Shell "C:\Windows\explorer.exe " & FolderPath, vbNormalFocus

End Sub

Sub OpenFile(FilePath$)
'指定パスのファイルを起動する。
'20210726
    
    Dim WSH As Object
    Set WSH = CreateObject("WScript.Shell")
    WSH.Run FilePath
    
End Sub

Sub OpenApplication(ApplicationPath$)
'指定パスのアプリを起動する
'例)電卓なら"calc.exe"など
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
'指定配列をCSVで出力する
'20210721

'FolderPath・・・出力先のフォルダパス
'FileName・・・出力するファイル名（拡張子は付けない）
'OutputHairetu・・・出力する配列

    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    '1次元配列は2次元配列に変換
    OutputHairetu = Lib入力配列を処理用に変換(OutputHairetu)
    
    N = UBound(OutputHairetu, 1)
    M = UBound(OutputHairetu, 2)
    Dim fp
    
    ' FreeFile値の取得(以降この値で入出力する)
    fp = FreeFile
    ' 指定ファイルをOPEN(出力モード)
    Open FolderPath & "\" & FileName & ".csv" For Output As #fp
    ' 最終行まで繰り返す
    
    For I = 1 To N
        For J = 1 To M - 1
            ' レコードを出力
            Print #fp, OutputHairetu(I, J) & ",";
        Next J
        Print #fp, OutputHairetu(I, M)
    Next I
    ' 指定ファイルをCLOSE
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
'指定配列をtxtで出力する
'20210721
   
'FolderPath・・・出力先のフォルダパス
'FileName・・・出力するファイル名（拡張子はつける）
'OutputHairetu・・・出力する配列
'KugiriMoji・・・文字列間の区切り文字

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
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

Function InputText(FilePath$, Optional KugiriMoji$ = "")
'テキストファイルを読み込んで配列で返す
'文字コードは自動的に判定して読込形式を変更する
'20210721

'FilePath・・・テキストファイルのフルパス
'KugiriMoji・・・テキストファイルを読み込んで区切り文字で区切って配列で出力する場合の区切り文字

    'テキストファイルの存在確認
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("「" & FilePath & "」" & vbLf & _
               "の存在が確認できません。" & vbLf & _
               "処理を終了します。")
        End
    End If
    
    'テキストファイルの文字コードを取得
    Dim strCode
    strCode = fncGetCharset(FilePath)
    If strCode = "UTF-8 BOM" Or strCode = "UTF-8" Then
        strCode = "UTF-8"
    ElseIf strCode = "UTF-16 LE BOM" Or strCode = "UTF-16 BE BOM" Then
        strCode = "UTF-16LE"
    Else
        strCode = Empty
    End If
    
    'テキストファイル読込
    Dim Output
    Dim RowCount&
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim FileNo%, Buffer$
    
    If IsEmpty(strCode) = False Then 'UTF8版の場合※※※※※※※※※※※※※※※※※
   
        Output = InputTextUTF8(FilePath, KugiriMoji)
    
    Else 'Shift-JIS版の場合※※※※※※※※※※※※※※※※※
        
        Output = InputTextShiftJIS(FilePath, KugiriMoji)
     
    End If

    InputText = Output
    
End Function

Function InputTextShiftJIS(FilePath$, Optional KugiriMoji$ = "")
'テキストファイルを読み込む ShiftJIS形式専用
'20210721

'FilePath・・・テキストファイルのフルパス
'KugiriMoji・・・テキストファイルを読み込んで区切り文字で区切って配列で出力する場合の区切り文字
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim FileNo%, Buffer$, SplitBuffer
    Dim Output1, Output2
    ' FreeFile値の取得(以降この値で入出力する)
    FileNo = FreeFile
    
    N = GetRowCountTextFile(FilePath)
    ReDim Output1(1 To N)
    ' 指定ファイルをOPEN(入力モード)
    Open FilePath For Input As #FileNo
            
    ' ファイルのEOF(End of File)まで繰り返す
    I = 0
    M = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, Buffer
        I = I + 1
        Output1(I) = Buffer '1次読込み
        
        If KugiriMoji <> "" Then '文字で区切る場合は区切り個数を計算
            '区切り文字による区切り個数の最大値を常に更新していく
            M = WorksheetFunction.Max(M, UBound(Split(Buffer, KugiriMoji)) + 1)
        End If
    Loop
    
    Close #FileNo
    
    '区切り文字の処理
    If KugiriMoji = "" Then
        '区切り文字なし
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
'テキストファイルを読み込む UTF8形式専用
'20210721

'FilePath・・・テキストファイルのフルパス
'KugiriMoji・・・テキストファイルを読み込んで区切り文字で区切って配列で出力する場合の区切り文字

    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim Buffer$, SplitBuffer
    Dim Output1, Output2
    
    N = GetRowCountTextFile(FilePath)
    ReDim Output1(1 To N)
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2 ' ファイルのタイプ(1:バイナリ 2:テキスト)
        .Open
        .LineSeparator = 10 '改行コード
        .LoadFromFile (FilePath)
        
        For I = 1 To N
            Buffer = .ReadText(-2)
            Output1(I) = Buffer
            If KugiriMoji <> "" Then '文字で区切る場合は区切り個数を計算
                '区切り文字による区切り個数の最大値を常に更新していく
                M = WorksheetFunction.Max(M, UBound(Split(Buffer, KugiriMoji)) + 1)
            End If
        Next I
        .Close
    End With
    
    '区切り文字の処理
    If KugiriMoji = "" Then
        '区切り文字なし
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
'20200909追加
'テキストファイルの文字コードを返す
'参考https://popozure.info/20190515/14201

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
    
    'ファイル読み込み
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
    
    'BOMによる判断
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
           
    '文字コード出現順位による判断
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
'フォルダ内のファイルのリストを取得する
'20210721

'「Microsoft Scripting Runtime」ライブラリを参照すること

'FolderPath・・・検索対象のフォルダパス
'Extensions・・・取得対象の拡張子、可変長引数配列で入力可能

    'フォルダの確認
    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("「" & FolderPath & "」" & vbLf & _
               "のフォルダの存在が確認できません。" & vbLf & _
               "処理を終了します。")
    End If
    
    '拡張子の連想配列を作成
    Dim ExtensionDict As Object, TmpExtension
    
    If IsMissing(Extensions) <> True Then
        '拡張子が入力されている場合
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
        'ファイル無し
        Exit Function
    End If
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
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
'フォルダ内のサブフォルダのリストを取得する
'20210721

'FolderPath・・・検索対象のフォルダパス

    'フォルダの確認
    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("「" & FolderPath & "」" & vbLf & _
               "のフォルダの存在が確認できません。" & vbLf & _
               "処理を終了します。")
    End If
    
    '拡張子の連想配列を作成
    Dim ExtensionDict As Object, TmpExtension
    
    Dim FSO As New FileSystemObject
    Dim myFolder As Folder
    Dim mySubFolder As Folders, TmpSubFolder As Folder
    Set myFolder = FSO.GetFolder(FolderPath)
    Set mySubFolder = myFolder.SubFolders
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    N = mySubFolder.Count
    
    If N = 0 Then
        'サブフォルダ無し
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
'指定シートをPDF化する
'20210721

'TargetSheet・・・PDF化する対象のシート
'FolderPath・・・出力先フォルダ 指定しない場合はブックと同じフォルダ
'FileName・・・出力PDFのファイル名 指定しない場合はシートの名前

    If FolderPath = "" Then
        FolderPath = TargetSheet.Parent.Path
    End If
    
    If FileName = "" Then
        FileName = TargetSheet.Name
    End If
    
    '出力先フォルダがない場合は作成する。
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    Dim OutputFileName$
    
    OutputFileName = FolderPath & "\" & FileName & ".pdf"

    TargetSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=OutputFileName
    
    If MessageIruNaraTrue Then
        If MsgBox("「" & FileName & ".pdf" & "」" & vbLf & "を作成しました" & vbLf & _
            "出力先フォルダを起動しますか?", vbYesNo + vbQuestion) = vbYes Then
            Shell "C:\Windows\explorer.exe " & FolderPath, vbNormalFocus
        End If
    End If
    
End Sub

Sub OutputXML(Title$, InputList, TateTableList, YokoTableList, _
              Optional TateTableName$ = "DATA", Optional YokoTableName$ = "ID", _
              Optional FolderPath$, Optional FileName$)
              
'テーブルデータからXMLデータを出力する
'「Microsoft XML, v6.0」ライブラリを参照すること
'参考：http://www.openreference.org/articles/view/651
'20210727
'20210824改良

'Title          :タイトル
'InputList      :XMLデータ内のデータリスト（2次元配列）
'TateTableList  :縦方向のテーブル名のリスト（1次元配列）
'YokoTableList  :横方向のテーブル名のリスト（1次元配列）
'TateTableName  :縦方向のテーブル名。デフォルトは"DATA"
'YokoTableName  :横方向のテーブル名。デフォルトは"ID"
'FolderPath     :XMLデータの出力する先のフォルダパス。デフォルトは自Excelブックのパス
'FileName       :XMLデータを出力するファイル名（拡張子"xml"は含まず）　デフォルトはタイトル(Title)と同じ

    '引数のチェック
    Call CheckArray1D(TateTableList, "TateTableList")
    Call CheckArray1D(YokoTableList, "YokoTableList")
    Call CheckArray2D(InputList, "InputList")
    Call CheckArray1DStart1(TateTableList, "TateTableList")
    Call CheckArray1DStart1(YokoTableList, "YokoTableList")
    Call CheckArray2DStart1(InputList, "InputList")
    
    If UBound(TateTableList, 1) <> UBound(InputList, 1) Then
        MsgBox ("「TateTableList」の要素数と" & vbLf & _
                "「InputList」の縦要素数を一致させてください")
        Stop
        End
    End If
    
    If UBound(YokoTableList, 1) <> UBound(InputList, 2) Then
        MsgBox ("「YokoTableList」の要素数と" & vbLf & _
                "「InputList」の横要素数を一致させてください")
        Stop
        End
    End If
    
    '引数のデフォルト値設定
    If FolderPath = "" Then
        FolderPath = ThisWorkbook.Path
    End If
    
    If FileName = "" Then
        FileName = Title
    End If
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(InputList, 1)
    M = UBound(InputList, 2)
    
    Dim XMLDoc As New MSXML2.DOMDocument60
    Dim xmlRoot As IXMLDOMNode
    Dim xmlData As IXMLDOMNode
    Dim xmlChildData As IXMLDOMNode
    Dim xmlAttr As MSXML2.IXMLDOMAttribute

    With XMLDoc
        'XML宣言を生成
        Call .appendChild(.createProcessingInstruction("xml", "version=""1.0"" encoding=""Shift_JIS"""))
        
        '要素を生成
        Set xmlRoot = .appendChild(.createElement(Title))
        
        For I = 1 To N
            
            '要素を生成
            Set xmlData = .createElement(YokoTableName)      '横要素を生成
            Set xmlAttr = .createAttribute(TateTableName)    '縦要素を生成
            xmlAttr.NodeValue = TateTableList(I)             '横要素の値を設定
            Call xmlData.Attributes.setNamedItem(xmlAttr)    '要素にid属性を設定
            
            '要素の子要素を生成して要素に追加
            For J = 1 To M
                Set xmlChildData = xmlData.appendChild(.createElement(YokoTableList(J)))
                xmlChildData.Text = InputList(I, J)
            Next J
            
            Call xmlRoot.appendChild(xmlData)
        Next I
        
        'XMLドキュメントを出力
        .Save (FolderPath & "\" & FileName & ".xml")
    End With

End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2%, Dummy3%
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

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Function GetFileName$(FilePath$)
'ファイルのフルパスからファイル名取得
'関数思い出し用
'20210824
    
    Dim Output$
    Dim TmpList
    TmpList = Split(FilePath, "\")
    Output = TmpList(UBound(TmpList))
    GetFileName = Output
    
End Function

Private Function Lib入力配列を処理用に変換(InputHairetu)
'入力した配列を処理用に変換する
'1次元配列→2次元配列
'数値か文字列→2次元配列(1,1)
'要素の開始番号を1にする
'20210721

    Dim Output
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Base1%, Base2%
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



