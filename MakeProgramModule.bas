Attribute VB_Name = "MakeProgramModule"
'
' プログラム作成
'
Sub プログラム作成()
    ' イベント発生を抑制
    Call EventChange(False)

    ' カレントワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook

    ' エントリー一覧シート
    Call SheetActivate(S_ENTRY_SHEET_NAME)
    Dim oEntrySheet As Worksheet
    Set oEntrySheet = ActiveSheet
    
    ' プログラムシートを作成（ヘッダ行まで）
    Call MakeSheet(oWorkBook, S_PROGRAM_SHEE_TNAME)
    Dim oProgramSheet As Worksheet
    Set oProgramSheet = ActiveSheet

    ' エントリー一覧読み込み
    Dim oEntryList As Object
    Call ReadEntrySheet(S_ENTRY_TABLE_NAME, oEntryList)

    ' 予選なし決勝の設定
    If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
        Call CheckFinal(oEntryList)
    End If

    ' プログラム作成
    Call MakeProgram(oProgramSheet, S_ENTRY_TABLE_NAME, oEntryList)

    ' プログラムの名前設定
    Call SetProgramName(oProgramSheet)
    
    ' プログラムの印刷エリア設定
    Call SetPrintArea(oProgramSheet)

    ' イベント発生を再開
    Call EventChange(True)
    
    ' シートを保存
    oWorkBook.Save
End Sub

'
' エントリー一覧読み込み
'
' sTableName    IN      テーブル名
' oEntryList    OUT     エントリー一覧(Dictionary)
' └プロNo
' 　└組
' 　　└レーン = 行番号
'
Public Sub ReadEntrySheet(sTableName As String, oEntryList As Object)

    ' 出力用エントリー一覧
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    Dim oProNo As Object    ' プロNo
    Dim oHeats As Object    ' 組
    
    ' プログラムNo毎に読み込み
    For Each cProNo In Range(sTableName & "[プロNo]")
        If Not oEntryList.Exists(cProNo.Value) Then
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add cProNo.Value, oProNo
        End If
        
        ' 行番号
        nHeat = cProNo.Offset(0, Range(sTableName & "[組]").Column - Range(sTableName & "[プロNo]").Column).Value
        nLane = cProNo.Offset(0, Range(sTableName & "[レーン]").Column - Range(sTableName & "[プロNo]").Column).Value
        
        ' 組毎に読み込み
        If Not oProNo.Exists(nHeat) Then
            Set oHeats = CreateObject("Scripting.Dictionary")
            oProNo.Add nHeat, oHeats
        End If
        
        ' レーン重複チェック
        If oHeats.Exists(nLane) Then
            MsgBox "プロNo：" & Str(cProNo.Value) & vbCrLf & _
                    "組　　：" & Str(nHeat) & vbCrLf & _
                    "レーン：" & Str(nLane) & vbCrLf & _
                    "が重複しています。"
            Range(sTableName).Parent.Activate
            Range(Cells(cProNo.Row, Range(sTableName & "[レースNo]").Column), _
                    Cells(cProNo.Row, Range(sTableName & "[レーン]").Column)).Select
            cProNo.Activate
            End
        End If
        ' レーン登録
        oHeats.Add nLane, cProNo.Row
    Next

End Sub

'
' 予選決勝確認（選手権用）
'
' 予選が１組しかない場合は
'
' oEntryList    OUT     エントリー一覧(Dictionary)
'
Sub CheckFinal(oEntryList As Object)

    Dim oProNo As Object
    Dim nFinalNo As Integer
    Dim oHeats As Object
    
    ' プログラム番号毎
    For Each vProNo In GetAreaKeyData("選手権種目区分")
        ' 申込みのあるプロNo
        If oEntryList.Exists(vProNo.Value) Then
            Set oProNo = oEntryList.Item(vProNo.Value)
            
            ' 決勝番号を取得
            nFinalNo = VLookupArea(vProNo.Value, "選手権種目区分", "決勝番号")
            
            ' 予選の場合
            If vProNo.Value <> nFinalNo Then
            
                ' １組しかない場合
                If oProNo.Count = 1 Then
                    ' 直接決勝にする
                    oEntryList.Add nFinalNo, oProNo
                    ' 予選には予選キーに決勝文字列を記載
                    oEntryList.Remove vProNo.Value
                    Set oProNo = CreateObject("Scripting.Dictionary")
                    oEntryList.Add vProNo.Value, oProNo
                    oProNo.Add "予選", "予選なし-->決勝へ No." & CStr(nFinalNo)
                ' 予選がある場合
                Else
                    ' 決勝キーに大会記録、標準記録を登録
                    Set oProNo = CreateObject("Scripting.Dictionary")
                    oEntryList.Add nFinalNo, oProNo
                    
                    ' 決勝キーに空の組を入れておく
                    Set oHeats = CreateObject("Scripting.Dictionary")
                    oHeats.Add "決勝", vProNo.Value
                    oProNo.Add "決勝", oHeats
                End If
            End If
        End If
    Next vProNo

End Sub

'
' プログラムシートを作成
'
' oWorkBook     IN      ワークシート
' sSheetName    OUT     シート名
'
Sub MakeSheet(oWorkBook As Workbook, sSheetName As String)

    If IsSheetExists(sSheetName) Then
        ' シートが存在する場合は内容をすべて削除
        Sheets(sSheetName).Activate
        Cells.Select
        Selection.Delete Shift:=xlUp
    Else
        ' 存在しない場合は作成する
        oWorkBook.Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = sSheetName
    End If
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ヘッダ行作成
    Call CopyHeaderCell(oWorkSheet, "Header通番")
    Call CopyHeaderCell(oWorkSheet, "HeaderプロNo")
    Call CopyHeaderCell(oWorkSheet, "Header組")
    Call CopyHeaderCell(oWorkSheet, "Headerレーン")
    Call CopyHeaderCell(oWorkSheet, "Header氏名")
    Call CopyHeaderCell(oWorkSheet, "Header種目")
    Call CopyHeaderCell(oWorkSheet, "Header所属前")
    Call CopyHeaderCell(oWorkSheet, "Header所属")
    Call CopyHeaderCell(oWorkSheet, "Header所属後")
    Call CopyHeaderCell(oWorkSheet, "Header区分")
    Call CopyHeaderCell(oWorkSheet, "Header時間")
    Call CopyHeaderCell(oWorkSheet, "Header順位")
    Call CopyHeaderCell(oWorkSheet, "Header備考")
    Call CopyHeaderCell(oWorkSheet, "Header大会記録")
    Call CopyHeaderCell(oWorkSheet, "Header申込み記録")
    Call CopyHeaderCell(oWorkSheet, "HeaderレースNo")
    Call CopyHeaderCell(oWorkSheet, "Headerソート区分")

    If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
        Call CopyHeaderCell(oWorkSheet, "Header標準記録")
    End If

End Sub

' ヘッダーセルをコピー
'
' 値、表示形式、縦幅、横幅、縦位置、横位置を設定
'
' Worksheet     IN      ワークシート
' sCellName     IN      セルの名前
'
Sub CopyHeaderCell(oWorkSheet As Worksheet, sCellName As String)

    Dim oRange As Range
    Set oRange = GetRange(sCellName)
    With oWorkSheet.Cells(1, oRange.Column)
        .NumberFormatLocal = oRange.NumberFormatLocal
        .ColumnWidth = oRange.ColumnWidth
        .RowHeight = oRange.RowHeight
        .HorizontalAlignment = oRange.HorizontalAlignment
        .VerticalAlignment = oRange.VerticalAlignment
        .Value = oRange.Value
    End With
End Sub

'
' プログラム作成
'
' oWorkSheet    IN      プログラムシート
' sTableName    IN      テーブル名
' oEntryList    IN      エントリー一覧
'
Sub MakeProgram(oWorkSheet As Worksheet, sTableName As String, oEntryList As Object)

    oWorkSheet.Activate

    Dim nCurrentRow As Integer
    nCurrentRow = 1

    ' ヘッダ行を作成
    Call SetNo(oWorkSheet, nCurrentRow)

    Dim oProNo As Object
    Dim oHeats As Object
    
    Dim nMaxProNo As Integer
    Dim nMaxHeat As Integer
    Dim nRaceNo As Integer
    nMaxProNo = GetRange(GetMaster(GetRange("大会名").Value)).Columns(1).Rows().Count
    
    Dim sMessage As String
    
    ' プログラム番号毎
    For Each nProNo In GetAreaKeyData(GetMaster(GetRange("大会名").Value))
        If oEntryList.Exists(Int(nProNo)) Then
            ' 申込みのあるプロNo
            Set oProNo = oEntryList.Item(Int(nProNo))
            nMaxHeat = oProNo.Count
        Else
            ' 申込みのないプロNo
            Set oProNo = Nothing
            nMaxHeat = 1
        End If
        
        ' プログラムヘッダ作成
        Call SetNo(oWorkSheet, nCurrentRow)
        Call MakeProgramHeader(oWorkSheet, sTableName, nCurrentRow, Int(nProNo))
        'Call CopyFormat(nCurrentRow - 1, "Prog組ヘッダフォーマット")
        
        ' 組番号毎
        For nHeat = 1 To nMaxHeat
            sMessage = ""
            If oProNo Is Nothing Then
                ' 申込みのないプロNoの場合は空の１組目を出力
                Set oHeats = Nothing
            ElseIf oProNo.Exists(nHeat) Then
                ' 組が存在する場合は組の値を出力
                Set oHeats = oProNo.Item(nHeat)
            ElseIf nHeat = 1 Then
                If oProNo.Exists("予選") Then
                ' 選手権の予選なしの場合は決勝へのメッセージを出力
                    sMessage = oProNo.Item("予選")
                ' 選手権の予選のある決勝の場合は大会記録、レース番号を入れる
                ElseIf oProNo.Exists("決勝") Then
                    Set oHeats = oProNo.Item("決勝")
                    nRaceNo = nRaceNo + 1
                End If
            Else
                ' 組が存在しない場合（異常系）
                Set oHeats = Nothing
            End If

            ' 組ヘッダ作成
            'Call CopyFormat(nCurrentRow, "Prog組フォーマット")
            Call SetNo(oWorkSheet, nCurrentRow)
            Call MakeHeatHeader(oWorkSheet, sTableName, nCurrentRow, Int(nHeat))
            
            ' タイトル修正
            Call SetTitleMenu("プログラム作成中: " & Str(nProNo) & "/" & Str(nMaxProNo))

            If sMessage <> "" Then
                ' 直接決勝へ
                Call SetNo(oWorkSheet, nCurrentRow)
                Call SetNo(oWorkSheet, nCurrentRow)
                Call CopyCell(oWorkSheet, nCurrentRow, "HeaderプロNo", nProNo)
                Cells(nCurrentRow, GetRange("Header氏名").Column).Value = sMessage
            Else
                ' レーン毎
                For nLane = N_MIN_LANE_OF_RACE To N_MAX_LANE_OF_RACE
                    Call SetNo(oWorkSheet, nCurrentRow)
                    
                    If oHeats Is Nothing Then
                        ' 申込みのないProNoの場合はデフォルト表示
                        Call MakeHeatDefault(oWorkSheet, nCurrentRow, Int(nProNo), Int(nHeat), Int(nLane))
                    ElseIf oHeats.Exists("決勝") Then
                        ' 選手権の決勝の場合は大会記録、標準記録、レース番号を追加
                        Call MakeHeatDefault(oWorkSheet, nCurrentRow, Int(nProNo), Int(nHeat), Int(nLane), CStr(nRaceNo))
                    ElseIf oHeats.Exists(nLane) Then
                        ' 申込みのあるProNoでエントリがあるレーンの場合はデータを記述
                        Call MakeHeat(oWorkSheet, sTableName, nCurrentRow, Int(oHeats.Item(nLane)), Int(nProNo), Int(nHeat))
                    Else
                        ' 申込みのあるProNoでエントリがないレーンの場合はデフォルト表示
                        Call MakeHeatDefault(oWorkSheet, nCurrentRow, Int(nProNo), Int(nHeat), Int(nLane))
                    End If
                
                    ' レース番号を記録しておく
                    If Cells(nCurrentRow, GetRange("HeaderレースNo").Column).Value <> "" Then
                        nRaceNo = Cells(nCurrentRow, GetRange("HeaderレースNo").Column).Value
                    End If
                Next
            End If
            ' 空行を２行入れる
            Call SetNo(oWorkSheet, nCurrentRow)
            Call SetNo(oWorkSheet, nCurrentRow)
        Next
    Next
    
    ' タイトル修正
     Call SetTitleMenu("プログラム作完了: " & Str(nMaxProNo) & "/" & Str(nMaxProNo))
End Sub

'
' 通番設定
'
' プログラムのNo行を作成
'
' oWorkSheet    IN      プログラムシート
' nCurrentRow   IN      通番
'
Sub SetNo(oWorkSheet As Worksheet, nCurrentRow As Integer)
    nCurrentRow = nCurrentRow + 1
    With oWorkSheet.Cells(nCurrentRow, GetRange("Header通番").Column)
        .Value = Str(nCurrentRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'
' 書式コピー
'
' nCurrentRow   IN      現在の行数
' sRangeName    IN      範囲の名前
'
Sub CopyFormat(nCurrentRow As Integer, sRangeName As String)

    ' 元をコピー
    GetRange(sRangeName).Copy

    ' 書式をコピー
    Cells(nCurrentRow, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

End Sub


'
' プログラムヘッダ作成
'
' oWorkBook     IN      ワークシート
' sTableName    IN      テーブル名
' nCurrentRow   IN      カレント行数
' nProNo        IN      プログラム番号
'
Sub MakeProgramHeader(oWorkSheet As Worksheet, sTableName As String, nCurrentRow As Integer, nProNo As Integer)

    Dim sMaster As String
    sMaster = GetMaster(GetRange("大会名").Value)

    Call CopyCell(oWorkSheet, nCurrentRow, "ProgプロNo")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog種目区分")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog種目名")

    With Range(sTableName).ListObject
        Cells(nCurrentRow, GetRange("ProgプロNo").Column).Value = nProNo
        Cells(nCurrentRow, GetRange("Prog種目区分").Column).Value = _
            VLookupArea(nProNo, sMaster, "区分") & _
            VLookupArea(nProNo, sMaster, "性別")

        Cells(nCurrentRow, Range("Prog種目名").Column).Value = _
            VLookupArea(nProNo, sMaster, "距離") & _
            VLookupArea(nProNo, sMaster, "種目")
    
        ' 横須賀選手権は標準記録、大会記録を出力
        If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
            
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog決勝")
            Cells(nCurrentRow, Range("Prog決勝").Column).Value = _
                VLookupArea(nProNo, sMaster, "予選／決勝")
            
            
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog記録")
            Dim nFinalNo As Integer
            nFinalNo = VLookupArea(nProNo, "選手権種目区分", "決勝番号")
            Dim nQualify As Long
            nQualify = VLookupArea(nProNo, sMaster, "標準記録")
            Dim sFormat As String
            If nQualify < 10000 Then
                sQualifyFormat = "##"".""#"
            Else
                sQualifyFormat = "0"":""##"".""#"
            End If
            Dim nRecord As Long
            nRecord = VLookupArea(nFinalNo, "選手権大会記録", "記録")
            Dim sRecordFormat As String
            If nRecord < 10000 Then
                sRecordFormat = "##"".""##"
            Else
                sRecordFormat = "0"":""##"".""##"
            End If
            Cells(nCurrentRow, Range("Prog記録").Column).Value = _
                "（標準記録 " & Format(nQualify / 10, sQualifyFormat) & ", " & _
                "大会記録 " & Format(nRecord, sRecordFormat) & "）"
        End If
    
    End With

    ' 下線を引く
    With Range(Cells(nCurrentRow, Range("Header組").Column), Cells(nCurrentRow, Range("Header大会記録").Column)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

End Sub

' セルをコピー
'
' oWorkSheet    IN      ワークシート
' nRow          IN      行数
' sCellName     IN      デフォルトのセル名
' vOverRide     IN      コピーする文字列
'
Sub CopyCell(oWorkSheet As Worksheet, nRow As Integer, sCellName As String, Optional vOverRide As Variant = Empty)

    Dim oRange As Range
    Set oRange = GetRange(sCellName)
    With oWorkSheet.Cells(nRow, oRange.Column)
        .ShrinkToFit = oRange.ShrinkToFit
        .NumberFormatLocal = oRange.NumberFormatLocal
        .Font.Name = oRange.Font.Name
        .Font.Size = oRange.Font.Size
        .Font.Underline = oRange.Font.Underline
        .Font.Bold = oRange.Font.Bold
        .HorizontalAlignment = oRange.HorizontalAlignment
        .VerticalAlignment = oRange.VerticalAlignment
        .IndentLevel = oRange.IndentLevel
        If IsEmpty(vOverRide) Then
            .Value = Range(sCellName).Value
        Else
            .Value = CStr(vOverRide)
        End If
    End With
End Sub

'
' 組ヘッダ作成
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
' nCurrentRow   IN      カレント行番号
' nHeat         IN      組番号
'
Sub MakeHeatHeader(oWorkSheet As Worksheet, sTableName As String, nCurrentRow As Integer, nHeat As Integer)
    
    Call CopyCell(oWorkSheet, nCurrentRow, "Header組")
    Call CopyCell(oWorkSheet, nCurrentRow, "Headerレーン")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header氏名")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header所属前")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header所属")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header所属後")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header区分")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header時間")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header順位")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header備考")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header大会記録")

    'With Range(sTableName).ListObject
        Cells(nCurrentRow, Range("Prog組番").Column).Value = _
            "<" & Trim(Str(nHeat)) & "組>"
    'End With

End Sub

'
' 選手レコード作成
'
' oWorkSheet    IN      ワークシート
' nCurrentRow   IN      カレント行番号
' nProNo        IN      プログラム番号
' nHeat         IN      組番号
' nLane         IN      レーン番号
'
Sub MakeHeatDefault(oWorkSheet As Worksheet, nCurrentRow As Integer, _
nProNo As Integer, nHeat As Integer, nLane As Integer, _
Optional sRaceNo As String = Empty)
    
    Call CopyCell(oWorkSheet, nCurrentRow, "HeaderプロNo", nProNo)
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog組番", Format(nProNo, "0#") & "-" & Format(nHeat, "#"))
    Call CopyCell(oWorkSheet, nCurrentRow, "Progレーン", nLane)
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog氏名")
    Range(Cells(nCurrentRow, GetRange("Prog氏名").Column), _
        Cells(nCurrentRow, GetRange("Prog種目").Column)).Merge
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属前")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属後")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog区分")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog時間")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog順位")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog備考")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog大会記録")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog申込み記録")
    Call CopyCell(oWorkSheet, nCurrentRow, "ProgレースNo", sRaceNo)
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog標準記録")

End Sub

'
' 組情報作成
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
' nCurrentRow   IN      カレント行番号(プログラムシート)
' nRow          IN      カレント行番号(テーブル)
' nProNo        IN      プログラム番号
' nHeat         IN      組番号
'
Sub MakeHeat(oWorkSheet As Worksheet, sTableName As String, nCurrentRow As Integer, _
nRow As Integer, nProNo As Integer, nHeat As Integer)

    oWorkSheet.Activate

    With Range(sTableName).ListObject
        
        Call CopyCell(oWorkSheet, nCurrentRow, "HeaderプロNo", nProNo)
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog組番", _
                             Format(nProNo, "0#") & "-" & CStr(nHeat))
        Call CopyCell(oWorkSheet, nCurrentRow, "Progレーン", _
                            .ListColumns("レーン").Range(nRow).Value)
        
        If .ListColumns("選手名").Range(nRow).Value <> "" Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog氏名", _
                            .ListColumns("選手名").Range(nRow).Value)
        Else
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog氏名")
        End If
        Range(Cells(nCurrentRow, GetRange("Prog氏名").Column), _
            Cells(nCurrentRow, GetRange("Prog種目").Column)).Merge
        
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属前")
        If Trim(.ListColumns("学校名").Range(nRow).Value) <> "" Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属", _
                            .ListColumns("学校名").Range(nRow).Value)
        Else
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属", _
                            .ListColumns("チーム名").Range(nRow).Value)
        End If
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog所属後")
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog区分", _
                            .ListColumns("区分").Range(nRow).Value)
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog時間")
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog順位")
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog備考")

        ' 横須賀選手権水泳大会
        If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
            Dim nFinalNo As Integer
            nFinalNo = VLookupArea(.ListColumns("プロNo").Range(nRow).Value, "選手権種目区分", "決勝番号")
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog大会記録", _
                    VLookupArea(nFinalNo, "選手権大会記録", "記録"))
        ' 横須賀市民体育大会
        ElseIf GetRange("大会名").Value = "横須賀市民体育大会" Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog大会記録", _
                    VLookupArea(.ListColumns("プロNo").Range(nRow).Value & _
                    .ListColumns("区分").Range(nRow).Value, "市民大会記録", "記録"))
        ' 学童マスターズ大会
        Else
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog大会記録", _
                    VLookupArea(.ListColumns("プロNo").Range(nRow).Value & _
                    .ListColumns("ソート区分").Range(nRow).Value, "学マ大会記録", "記録"))
        End If

        
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog申込み記録", _
                            .ListColumns("申込み時間").Range(nRow).Value)
        
        Call CopyCell(oWorkSheet, nCurrentRow, "Progソート区分", _
                            .ListColumns("ソート区分").Range(nRow).Value)
        Call CopyCell(oWorkSheet, nCurrentRow, "ProgレースNo", _
                            .ListColumns("レースNo").Range(nRow).Value)

        ' 横須賀選手権水泳大会は標準記録も記載
        If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog標準記録", _
                    VLookupArea(.ListColumns("プロNo").Range(nRow).Value, "選手権種目区分", "標準記録"))
        End If
    
    End With

End Sub

'
' プログラム名前定義
'
' 「プログラム作成マクロ」からボタンで実行される
'
Sub プログラム名前定義()
    Sheets(S_PROGRAM_SHEE_TNAME).Activate
    Call SetProgramName(ActiveSheet)
End Sub

'
' プログラムシートの名前定義
'
' oWorkBook     IN      ワークシート
'
Sub SetProgramName(oWorkSheet As Worksheet)
    Call DeleteName("プログラム*")
    Call SetNoName(oWorkSheet)
    Call SetProNoName(oWorkSheet)
    Call SetProNoListName(oWorkSheet)
    Call SetHeatName(oWorkSheet)
    Call SetRaceName(oWorkSheet)
    Call SetSameRaceLabel(oWorkSheet)
End Sub

'
' プログラムシートの通番列の名前定義
'
' 名前「プログラム通番」を定義
'
' プログラムシートの２行目(A2)から最下位行までの
'
' oWorkBook     IN      ワークシート
'
Sub SetNoName(oWorkSheet As Worksheet)
    oWorkSheet.Activate
    Cells(2, GetRange("Header通番").Column).Select
    Range(Selection, Selection.End(xlDown)).Select
    Call DefineName("プログラム通番", Selection.Address(ReferenceStyle:=xlA1))
    Range("$A$1").Select
End Sub

'
' プログラム番号一覧の名前定義
'
' 名前「プログラム種目番号」の定義
'
' oWorkBook     IN      ワークシート
'
Sub SetProNoName(oWorkSheet As Worksheet)
    
    ' プロNo
    Dim nProNo As Integer
    nProNo = 1

    ' アドレス文字列格納用
    Dim sAddress As String
    sAddress = ""

    ' セルオブジェクト
    Dim oCell As Range

    ' プログラム通番をシークしながら処理をする
    For Each vNo In GetRange("プログラム通番")
        Set oCell = oWorkSheet.Cells(vNo.Row, Range("Header組").Column)
        ' 組列でプロNoと同じ場合はプログラム番号のセル
        If oCell.Value = nProNo Then
            If sAddress = "" Then
                sAddress = oCell.Address(ReferenceStyle:=xlA1)
            Else
                sAddress = sAddress & "," & oCell.Address(ReferenceStyle:=xlA1)
            End If
            ' プロNoをインクリメント
            nProNo = nProNo + 1
        End If
    Next vNo

    Call DefineName("プログラム種目番号", sAddress)

End Sub

'
' 記録画面検索用の名前定義
'
' 名前「プログラム番号N」の定義
'
' N：プログラム番号
'
' oWorkBook     IN      ワークシート
'
Sub SetProNoListName(oWorkSheet As Worksheet)
    
    ' プロNo
    Dim nProNo As Integer
    nProNo = 1
    
    ' アドレス文字列格納用
    Dim oRange As Range
    Set oRange = Nothing
    
    ' セルオブジェクト
    Dim oCell As Range
    
    ' プログラム通番をシークしながら処理をする
    For Each vNo In Range("プログラム通番")
        Set oCell = oWorkSheet.Cells(vNo.Row, GetRange("HeaderプロNo").Column)
        ' プロNo列でプロNoより大きくなった場合に登録
        If oCell.Value > nProNo Then
            ' アドレスが空でなければ名前を登録する
            If Not (oRange Is Nothing) Then
                Call DefineName("プログラム番号" & Trim(Str(nProNo)), oRange.Address)
                Set oRange = Nothing
                ' プロNoをインクリメント
                nProNo = nProNo + 1
            End If
        End If
        ' プロNo列でプロNoと同じ場合はプログラム番号のセル
        If oCell.Value = nProNo Then
            If oRange Is Nothing Then
                Set oRange = oCell
            Else
                Set oRange = Application.Union(oRange, oCell)
            End If
        End If
    Next vNo

    ' アドレスが空でなければ名前を登録する
    If Not (oRange Is Nothing) Then
        Call DefineName("プログラム番号" & Trim(Str(nProNo)), oRange.Address)
    End If
End Sub

'
' 記録画面検索用の組の名前定義
'
' 名前「プログラム組NN-X」の定義
'
' NN：プログラム番号
'  X：組番
'
' oWorkBook     IN      ワークシート
'
Sub SetHeatName(oWorkSheet As Worksheet)
   
    ' プログラム番号
    Dim nProNo As Integer
    nProNo = 0
    
    ' 次のプログラム番号
    Dim nNextProNo As Integer
    nNextProNo = 1
    
    ' 組番号
    Dim nHeat As Integer
    ' 組名
    Dim sHeatName As String
    
    ' アドレス文字列格納用
    Dim oRange As Range
    Set oRange = Nothing

    ' セルオブジェクト
    Dim oCell As Range

    For Each vNo In Range("プログラム通番")
        Set oCell = oWorkSheet.Cells(vNo.Row, GetRange("Header組").Column)
        ' 次のプログラム番号に変わる場合
        If oCell.Value = nNextProNo Then
            nProNo = nNextProNo         ' プログラム番号をインクリメント
            nNextProNo = nNextProNo + 1 ' 次のプログラム番号をインクリメント
            nHeat = 1                   ' 組番号の初期化
        End If
        ' 組名のフォーマット
        sHeatName = Format(nProNo, "0#") & "-" & Trim(Str(nHeat))
        ' 組と一致する場合は名前の範囲
        If oCell.Value = sHeatName Then
            If oRange Is Nothing Then
                Set oRange = oCell
            Else
                Set oRange = Application.Union(oRange, oCell)
            End If
        End If

        ' 空行で名前範囲がある場合
        If oCell.Value = "" And Not (oRange Is Nothing) Then
            ' 名前を定義する
            Call DefineName("プログラム組" & Replace(sHeatName, "-", "_"), oRange.Address)

            ' 名前範囲と組番号を初期化
            Set oRange = Nothing
            nHeat = nHeat + 1
        End If
    Next vNo
End Sub

'
' 記録画面検索用の名前定義
'
' 名前「プログラムレースNN」の定義
'
' NN：レース番号
'
' oWorkBook     IN      ワークシート
'
Sub SetRaceName(oWorkSheet As Worksheet)
    
    Dim nRaceNo As Integer
    nRaceNo = 0
        
    ' アドレス文字列格納用
    Dim oRange As Range
    Set oRange = Nothing
    
    ' セルオブジェクト
    Dim oCell As Range

    ' プログラム通番をシークしながら処理をする
    For Each vNo In Range("プログラム通番")
        Set oCell = oWorkSheet.Cells(vNo.Row, GetRange("HeaderレースNo").Column)
        ' 空白以外の場合
        If oCell.Value <> "" Then
            If oCell.Value > nRaceNo Then
                ' アドレスが空でなければ名前を登録する
                If Not (oRange Is Nothing) Then
                    Call DefineName("プログラムレース" & Trim(Str(nRaceNo)), oRange.Address)
                    Set oRange = Nothing
                End If
                nRaceNo = oCell.Value
            End If
            ' プロNo列でプロNoと同じ場合はプログラム番号のセル
            If oCell.Value = nRaceNo Then
                If oRange Is Nothing Then
                    Set oRange = oCell
                Else
                    Set oRange = Application.Union(oRange, oCell)
                End If
            End If
        End If
    Next vNo

    ' アドレスが空でなければ名前を登録する
    If Not (oRange Is Nothing) Then
        Call DefineName("プログラムレース" & Trim(Str(nRaceNo)), oRange.Address)
    End If

End Sub

'
' 同一レースラベル作成
'
' 同一レースの場合に「X-X-X 同一レース」という文言を追記する
'
' oWorkBook     IN      ワークシート
'
Sub SetSameRaceLabel(oWorkSheet As Worksheet)
    
    Dim oRaceNo As Object
    Set oRaceNo = CreateObject("Scripting.Dictionary")
    
    ' レースNoに対するプロNoを読込み
    Call ReadSameRace(oWorkSheet, oRaceNo)
    
    ' 同一レースラベルを書込み
    Call WriteSameRaceLabel(oRaceNo)

End Sub

'
' レースNoに対するプロNoを読込み
'
' oWorkBook     IN      ワークシート
' oRaceNo       OUT     レースNo配列
'  └レースNo
'  　└プロNo：1
'
Sub ReadSameRace(oWorkSheet As Worksheet, oRaceNo As Object)
    Dim nRaceNo As Integer
    Dim oProNo As Object
    For Each vNo In GetRange("プログラム通番")
        ' レースNoを取得
        nRaceNo = oWorkSheet.Cells(vNo.Row, GetRange("HeaderレースNo").Column).Value
        If nRaceNo > 0 Then
            If Not oRaceNo.Exists(nRaceNo) Then
                ' レースNoを追加
                Set oProNo = CreateObject("Scripting.Dictionary")
                oRaceNo.Add nRaceNo, oProNo
            End If
            ' プロNoを取得
            nProNo = Cells(vNo.Row, Range("HeaderプロNo").Column).Value
            If Not oProNo.Exists(nProNo) Then
                ' プロNoを追加
                oProNo.Add nProNo, 1
            End If
        
        End If
    Next vNo
End Sub

'
' 同一レースラベル書込み
'
' 記述する場所はProNoの１行前、氏名と同じ列
'
' oRaceNo       IN      レースNo配列
'
Sub WriteSameRaceLabel(oRaceNo As Object)
    Dim cProNo As Range
    For Each vRaceNo In oRaceNo
        Set oProNo = oRaceNo.Item(vRaceNo)
        If oProNo.Count > 1 Then
            aryProNo = oProNo.Keys()
            sLabel = Join(aryProNo, "-") & " 同一レース"
            For Each vProNo In aryProNo
                Set cProNo = GetProNoRow(Int(vProNo))
                cProNo.Offset(-1, GetRange("Prog氏名").Column - GetRange("ProgプロNo").Column).Value = sLabel
            Next vProNo
        End If
    Next vRaceNo
End Sub

'
' プログラム番号の行数を取得
'
' 名前「プログラム種目番号」からプログラムヘッダのProNoセルを取得
'
' oRaceNo       IN      レースNo配列
'
Function GetProNoRow(nProNo As Integer) As Range
    Dim sName As String
    sName = "プログラム種目番号"

    For Each vProNo In GetRange(sName)
        If vProNo.Value = nProNo Then
            Set GetProNoRow = vProNo
            Exit Function
        End If
    Next vProNo
End Function

'
' 印刷範囲設定
'
' oWorkBook     IN      ワークシート
'
Sub SetPrintArea(oWorkSheet As Worksheet)
    oWorkSheet.Activate
    
    ' 印刷エリアのクリア
    ActiveSheet.PageSetup.PrintArea = ""
    ' 改ページのクリア
    ActiveSheet.ResetAllPageBreaks
    
    ' 印刷エリアの設定
    Dim nBottom As Integer
    nBottom = Range("$A$1").End(xlDown).Row
    
    ' 選手権大会の場合は大会記録を印刷しない
    If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
        ActiveSheet.PageSetup.PrintArea = _
            Range(Cells(GetRange("Header組").Row, GetRange("Header組").Column), _
            Cells(nBottom, GetRange("Header備考").Column)).Address
        Cells(1, GetRange("Header氏名").Column).ColumnWidth = 20
        Cells(1, GetRange("Header種目").Column).ColumnWidth = 20
        Cells(1, GetRange("Header備考").Column).ColumnWidth = 20
    Else
        ActiveSheet.PageSetup.PrintArea = _
            Range(Cells(3, GetRange("Header組").Column), Cells(nBottom, GetRange("Header大会記録").Column)).Address
    End If

    ' 印刷エリアの設定（横１ページ）
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .FitToPagesWide = 1
        .CenterFooter = "－&P－"
    End With
    Application.PrintCommunication = True

    
    ' 改ページプレビュ
    ActiveWindow.View = xlPageBreakPreview
    
    ' 改ページ設定
    Dim nNum As Integer
    nNum = 0
    Dim bFlag As Boolean
    bFlag = True
    Dim nProNo As Integer
    For Each vNo In GetRange("プログラム通番")
        nProNo = Cells(vNo.Row, GetRange("HeaderプロNo").Column).Value
        If nProNo > 0 Then
            If bFlag Then
                nNum = nNum + 1
            End If
            bFlag = False
        Else
            If bFlag = False And nNum Mod 5 = 0 Then
                ' 改行ページ
                nRow = vNo.Row + 1
                If nRow < nBottom Then
                    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=Cells(nRow, GetRange("Header組").Column)
                End If
            End If
            bFlag = True
        End If
    Next vNo

    ' 改ページプレビュを戻す
    ActiveWindow.View = xlNormalView
    Range("$A$1").Select
    
    ' １行の高さ
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 17
    Range("$A$1").Select

End Sub


