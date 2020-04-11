Attribute VB_Name = "ReadEntryModule"
'
' エントリーファイル一覧の読み込み
'
' 特定のフォルダを指定して、中にあるエントリーファイルを
' すべて読み込み一覧シートに出力する。
'
Sub エントリー読込み()
    ' イベント発生を抑制
    Call EventChange(False)

    ' エクセルシートを選択
    Call SheetActivate(sEntrySheetName)

    ' 出力用ワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' 出力用ワークシート
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    ' エントリー一覧読込用配列
    Dim oGameList As Object
    Set oGameList = CreateObject("Scripting.Dictionary")

    ' エントリーファイル読込み
    Call ReadEntryFiles(oGameList)

    ' エントリーシートの書き込み
    Call WriteEntrySheet(oWorkSheet, sEntryTableName, oGameList)
    
    ' ProNo、ソート区分、申込み時間でソート
    Call SortByProNo(oWorkSheet, sEntryTableName)

End Sub

'
' エントリーファイル一覧の読み込み
'
' フォルダを指定して、その中に含まれるエントリーシート（*.xlsx）をすべて詠み込む
'
' oGameList     OUT     エントリー一覧
'
Sub ReadEntryFiles(ByRef oGameList As Object)

    ' ファイル一覧を取得
    '
    Dim sPathName As String
    sPathName = SelectDir()
    Dim FileList As Collection
    Set FileList = GetFiles(sPathName, "\*.xlsx")

    Dim nMax As Integer
    nMax = FileList.Count
    Dim nCount As Integer
    nCount = 0

    '
    ' ファイル毎に処理する
    '
    For Each vFile In FileList
        
        ' タイトル修正
        nCount = nCount + 1
        Call SetTitleMenu("プログラム読込中: " & Str(nCount) & "/" & Str(nMax))
        
        '
        ' ファイルを開く（読み取り専用）
        '
        Set SubBook = Workbooks.Open(Filename:=sPathName + "\" + vFile, ReadOnly:=True)
        Worksheets("記入票").Activate

        ' エントリー一覧の読込み
        Call ReadEntryFile(oGameList)
    
        ' 警告なしでファイルを閉じる（保存しない）
        Application.DisplayAlerts = False
        SubBook.Close
        Application.DisplayAlerts = True
    Next
    
    ' タイトル修正
    Call SetTitleMenu("プログラム読込完了: " & Str(nCount) & "/" & Str(nMax))
    
End Sub



'
' エントリーファイルの読込み
'
' oGameList
' │
' └─大会名：oTeamList       ・・・Range("大会名")
' 　　│
' 　　└─チーム名：oEntryList・・・Range("チーム名")
' 　　　　│
' 　　　　└─選手番号：oEntry・・・Range("選手番号")
'
' oEntry
' │
' ├─性別：Range("選手性別")
' │
' ├─選手名：Range("選手名")
' │
' ├─フリガナ：Range("選手フリガナ")
' │
' ├─区分：Range("選手区分")
' │
' └─1or2：oLines
'
' oLines
' │
' ├─種目名：Range("種目名")
' │
' ├─距離：Range("種目距離")
' │
' └─申込み時間：Range("選手分")＋Range("選手秒")＋Range("選手ミリ秒")
'
' oRelayEntry
' │
' └─L1〜L24：oRelayLines
'
' oRelayLines
' │
' ├─種目番号：Range("リレー種目")
' │
' ├─区分：Range("リレー区分")
' │
' └─申込み時間：Range("リレー分")＋Range("リレー秒")＋Range("リレーミリ秒")
'
' oGameList     OUT     エントリー一覧
'
Sub ReadEntryFile(ByRef oGameList As Object)

    ' 大会名
    Dim sGameName As String
    sGameName = GetRange("大会名").Value
    Dim oTeamList As Object
    If oGameList.Exists(sGameName) Then
        Set oTeamList = oGameList.Item(sGameName)
    Else
        Set oTeamList = CreateObject("Scripting.Dictionary")
        oGameList.Add sGameName, oTeamList
    End If
    
    ' チーム名
    Dim sTeamName As String
    sTeamName = RTrim(LTrim(GetRange("チーム名").Value))
    
    ' チーム申込み保存オブジェクト
    Dim nNum As Integer
    Dim oEntryList As Object
    If oTeamList.Exists(sTeamName) Then
        If sTeamName = "個人" Then
            Set oEntryList = oTeamList.Item(sTeamName)
            nNum = oEntryList.Count
        Else
            MsgBox "チーム名が重複しています。" + vbCrLf + sGameName + " : " + sTeamName
            End
        End If
    Else
        Set oEntryList = CreateObject("Scripting.Dictionary")
        oTeamList.Add sTeamName, oEntryList
    End If

    Dim nIdx As Integer

    ' 個人番号範囲をすべて読み込む
    For Each oCell In GetRange("選手番号")
        
        If oCell.MergeCells Then
            If oCell.Address = oCell.MergeArea.Item(1).Address Then
        
                ' 選手毎のエントリーリスト
                Dim oEntry As Object
                Set oEntry = CreateObject("Scripting.Dictionary")
            
                ' 個人No
                If sTeamName = "個人" Then
                    nNum = nNum + 1
                Else
                    nNum = oCell.Value
                End If

                oEntry.Add "性別", Cells(oCell.Row, GetRange("選手性別").Column).Value + "子"
                oEntry.Add "フリガナ", ReplaceName(Cells(oCell.Row, GetRange("選手フリガナ").Column).Value)

                If Range("大会名").Value = "横須賀選手権水泳大会" Then
                    
                    oEntry.Add "選手名", ReplaceName(Cells(oCell.Row + 1, GetRange("選手名").Column).Value)
                    oEntry.Add "区分", Cells(oCell.Row + 1, GetRange("選手区分").Column).Value
                    nIdx = 1
                
                ElseIf Range("大会名").Value = "横須賀市民体育大会" Then
                    
                    oEntry.Add "選手名", ReplaceName(Cells(oCell.Row + 2, GetRange("選手名").Column).Value)
                    oEntry.Add "学校名", Trim(Cells(oCell.Row + 4, GetRange("選手学校名").Column).Value)
                    If Cells(oCell.Row, GetRange("選手区分").Column).Value <> "" Then
                        oEntry.Add "区分", Cells(oCell.Row, GetRange("選手区分").Column).Value
                    Else
                        oEntry.Add "区分", "年齢区分"
                    End If
                    oEntry.Add "年齢", Cells(oCell.Row + 3, GetRange("選手年齢").Column).Value
                    nIdx = 3
                
                ElseIf Range("大会名").Value = "横須賀マスターズ大会" Then
                
                    oEntry.Add "選手名", ReplaceName(Cells(oCell.Row + 1, GetRange("選手名").Column).Value)
                    oEntry.Add "年齢", Cells(oCell.Row, GetRange("選手年齢").Column).Value
                    nIdx = 1
                
                Else
            
                    oEntry.Add "選手名", ReplaceName(Cells(oCell.Row + 1, GetRange("選手名").Column).Value)
                    oEntry.Add "区分", Cells(oCell.Row, GetRange("選手学年").Column).Value
                    nIdx = 1
                
                End If
                
                ' １行目
                Call ReadEntryLine(1, oCell.Row, oEntry)
                Call CheckEntry(oCell.Row, oEntry, 1)
    
                ' ２行目
                Call ReadEntryLine(2, oCell.Row + nIdx, oEntry)
                Call CheckEntry(oCell.Row + nIdx, oEntry, 2)
                If oEntry.Item("選手名") <> "" Then
                    oEntryList.Add nNum, oEntry
                End If
            End If
        End If
    Next oCell

    ' リレー種目番号範囲をすべて読み込む
    Dim nRelayNum As Integer
    nRelayNum = 0
    Dim oRelayEntry As Object
    Set oRelayEntry = Nothing
    For Each oCell In GetRange("リレー種目")
        ' 値が設定されている場合は読み込む
        If oCell.Value <> "" Then
            ' リレーのエントリーリスト
            If oRelayEntry Is Nothing Then
                Set oRelayEntry = CreateObject("Scripting.Dictionary")
                oEntryList.Add nNum, oRelayEntry
            End If

            nRelayNum = nRelayNum + 1
            Call ReadRelayEntryLine(nRelayNum, oCell.Row, oRelayEntry)
        End If
    Next oCell

End Sub

'
' エントリーファイルの個人種目行読み込み
'
' 種目名、距離、申込み時間を取得する
'
' nNum          IN      個人エントリー行(1,2)
' nRow          IN      行数
' oEntry        OUT     種目行
'
Sub ReadEntryLine(nNum As Integer, nRow As Integer, oEntry As Object)
    
    Dim oLines As Object
    Dim sStyle As String
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    ' 番号範囲をすべて読み込む
    For Each oCell In GetRange("種目名")
        If oCell.Value <> "" Then
            sStyle = oCell.Value
        End If
        ' 種目選択が空または全角空白以外の場合は選択されたものとする
        If Trim(Cells(nRow, oCell.Column).Value) <> "" And Cells(nRow, oCell.Column).Value <> "　" Then
            Set oLines = CreateObject("Scripting.Dictionary")
            oEntry.Add nNum, oLines
            
            oLines.Add "種目番号", VLookupArea(Cells(nRow, oCell.Column).Value, "種目番号区分", "種目番号")
            oLines.Add "種目区分", VLookupArea(Cells(nRow, oCell.Column).Value, "種目番号区分", "区分")
            oLines.Add "種目名", ReplaceStyle(sStyle)
            oLines.Add "距離", ReplaceDistance(Cells(GetRange("種目距離").Row, oCell.Column).Value)
            nMin = Cells(nRow, GetRange("選手分").Column).Value
            nSec = Cells(nRow, GetRange("選手秒").Column).Value
            nMil = Cells(nRow, GetRange("選手ミリ秒").Column).Value
            oLines.Add "申込み時間", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
            Exit Sub
        End If
    Next oCell
End Sub

'
' エントリーの種目番号が正しいかを確認
'
' nRow          IN      行番号
' oEntry        IN      種目行
' nNum          IN      個人エントリー行(1,2)
'
Sub CheckEntry(nRow As Integer, oEntry As Object, nNum As Integer)
    
    If IsEmpty(oEntry.Item(nNum)) Then
        Exit Sub
    End If
    
    Dim oLines As Object
    Set oLines = oEntry.Item(nNum)
    
    Dim sGender As String
    Dim sDistance As String
    Dim sStyle As String
    
    sGender = VLookupArea(oLines.Item("種目番号"), "種目番号区分", "性別")
    sDistance = VLookupArea(oLines.Item("種目番号"), "種目番号区分", "距離")
    sStyle = VLookupArea(oLines.Item("種目番号"), "種目番号区分", "種目")
    
    If sGender <> oEntry.Item("性別") Or sDistance <> oLines.Item("距離") Or sStyle <> oLines.Item("種目名") Then
        MsgBox CStr(nRow) & "行目：種目番号が正しくありません。：" & oLines.Item("種目番号")
        End
    End If

End Sub



'
' 種目名称の置換
'
' sStyle        IN      種目
'
Function ReplaceStyle(sStyle)
    Dim sTemp As String
    sTemp = sStyle
    sTemp = Replace(sTemp, "ﾊﾞﾀﾌﾗｲ", "バタフライ")
    sTemp = Replace(sTemp, "個メ", "個人メドレー")
    ReplaceStyle = sTemp
End Function

'
' 距離名称の置換
'
' sDistance     IN      距離
'
Function ReplaceDistance(sDistance)
    Dim sTemp As String
    sTemp = sDistance
    sTemp = Replace(sTemp, "二五", "25M")
    sTemp = Replace(sTemp, "五〇", "50M")
    sTemp = Replace(sTemp, "一〇〇", "100M")
    sTemp = Replace(sTemp, "二〇〇", "200M")
    sTemp = Replace(sTemp, "四〇〇", "400M")
    ReplaceDistance = sTemp
End Function

'
' 選手名の置換
'
' 姓が１文字の場合は性に全角空白を足す
' 性が２文字以内で名が１文字の場合は名に全角空白を足す
'
' sName         IN      選手名
'
Function ReplaceName(sName)
    
    ' 空白の場合は何もしない
    If Trim(sName) = "" Then
        ReplaceName = ""
        Exit Function
    End If
    
    Dim sTemp As String
    sTemp = STrim(sName)
    
    Dim sTemps As Variant
    sTemps = Split(sTemp, " ")
    ' 姓が１文字の場合は性に全角空白を足す
    If Len(sTemps(0)) = 1 Then
        sTemps(0) = sTemps(0) & "　"
    End If
    ' 性が２文字以内で名が１文字の場合は名に全角空白を足す
    If Len(sTemps(1)) = 1 And Len(sTemps(0)) <= 2 Then
        sTemps(1) = "　" & sTemps(1)
    End If
        
    ReplaceName = sTemps(0) & "　" & sTemps(1)
End Function

'
' エントリーファイルのリレー種目行読み込み
'
' nNum          IN    リレー番号
' nRow          IN    カレント行数
' oRelayEntry   I/O   種目行
'
Sub ReadRelayEntryLine(nNum As Integer, nRow As Integer, oRelayEntry As Object)
    
    Dim oLines As Object
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    If Cells(nRow, GetRange("リレー種目").Column).Value <> "" Then
        Set oRelayLines = CreateObject("Scripting.Dictionary")
        oRelayEntry.Add "L" + Str(nNum), oRelayLines

        oRelayLines.Add "種目番号", Cells(nRow, GetRange("リレー種目").Column).Value
        If IsNameExists("リレー区分") Then
            oRelayLines.Add "区分", Cells(nRow, GetRange("リレー区分").Column).Value
        End If
        nMin = Cells(nRow, GetRange("リレー分").Column).Value
        nSec = Cells(nRow, GetRange("リレー秒").Column).Value
        nMil = Cells(nRow, GetRange("リレーミリ秒").Column).Value
        oRelayLines.Add "申込み時間", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
    End If
End Sub

'
' 申込みをシートに出力
'
' oWorkBook     IN     出力するシートのあるエクセルWorkBook
' sTable        IN     テーブル名
' oTeamList     IN     読み込んだチーム申込み一覧
'
Sub WriteEntrySheet(oWorkSheet As Worksheet, sTable As String, oGameList As Object)
    
    ' エントリーテーブルを初期化
    Call DeleteTable(oWorkSheet, sTable)
    
    ' エントリー一覧の出力
    Call WriteTeamEntries(oWorkSheet, sTable, oGameList)

End Sub

'
' 申込みをシートに出力
'
' oWorkBook     IN     出力するシートのあるエクセルWorkBook
' sTable        IN     テーブル名
' oTeamList     IN     読み込んだチーム申込み一覧
'
Sub WriteTeamEntries(oWorkSheet As Worksheet, sTable As String, oGameList As Object)

    oWorkSheet.Activate

    Dim nPersonNo As Integer
    Dim nTeamNo As Integer
    nTeamNo = 1
    
    Dim nRow As Integer
    nRow = 1
    For Each sGame In oGameList
        Dim oTeamList As Object
        Set oTeamList = oGameList.Item(sGame)
        For Each sTeam In oTeamList
            Dim oEntryList As Object
            Set oEntryList = oTeamList.Item(sTeam)
            
            Dim oLine As Object
            For Each sNum In oEntryList
                Dim oEntry As Object
                Set oEntry = oEntryList.Item(sNum)
                nPersonNo = nTeamNo * 100 + Int(sNum)
                
                If oEntry.Exists("選手名") Then
                    ' 個人
                    For i = 1 To 2
                        If Not IsEmpty(oEntry.Item(i)) Then
                            nRow = nRow + 1
                            Set oLine = oEntry.Item(i)
                            Call WriteLine(sTable, nRow, nPersonNo, sGame, sTeam, oEntry, oLine)
                        End If
                    Next i
                Else
                    ' リレー
                    Dim sKey As String
                    For i = 1 To 24
                        sKey = "L" & Str(i)
                        If oEntry.Exists(sKey) Then
                            nRow = nRow + 1
                            Set oLine = oEntry.Item(sKey)
                            Call WriteRelayLine(sTable, nRow, nTeamNo, sGame, sTeam, oEntry, oLine)
                        End If
                    Next i
                End If
            Next
            ' チーム番号をインクリメント
            nTeamNo = nTeamNo + 1
        Next
    Next
End Sub

'
' エントリーテーブルを初期化
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
'
Sub DeleteTable(oWorkSheet As Worksheet, sTableName As String)
    Dim myTable As ListObject
    Set myTable = oWorkSheet.ListObjects(sTableName)
    If Not (myTable.DataBodyRange Is Nothing) Then
        myTable.DataBodyRange.Delete
    End If
End Sub

'
' 申込み行を出力
'
' sTable        IN      テーブル名
' nRow          IN      出力行番号
' nPersonNo     IN      選手番号
' sGame         IN      大会名
' sTeam         IN      チーム名
' oEntry        IN      エントリー
' oLine         IN      種目、申込み時間
'
Sub WriteLine( _
    sTable As String, _
    nRow As Integer, _
    nPersonNo As Integer, _
    sGame As Variant, _
    sTeam As Variant, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[個人No]").Column).Value = nPersonNo
    Cells(nRow, Range(sTable & "[プロNo]").Column).Value = oLine.Item("種目番号")
    Cells(nRow, Range(sTable & "[チーム名]").Column).Value = CStr(sTeam)
    Cells(nRow, Range(sTable & "[選手名]").Column).Value = oEntry.Item("選手名")
    Cells(nRow, Range(sTable & "[フリガナ]").Column).Value = oEntry.Item("フリガナ")
    Cells(nRow, Range(sTable & "[性別]").Column).Value = oEntry.Item("性別")
    Cells(nRow, Range(sTable & "[距離]").Column).Value = oLine.Item("距離")
    Cells(nRow, Range(sTable & "[種目名]").Column).Value = oLine.Item("種目名")
    Cells(nRow, Range(sTable & "[申込み時間]").Column).Value = oLine.Item("申込み時間")
    If oLine.Item("申込み時間") >= 10000 Then
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    Dim nColumn As Integer
    If CStr(sGame) = "横須賀選手権水泳大会" Then
    
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = ""
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = ""
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oEntry.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    
    ElseIf CStr(sGame) = "横須賀市民体育大会" Then
    
        Cells(nRow, Range(sTable & "[学校名]").Column).Value = oEntry.Item("学校名")
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = oEntry.Item("年齢")
        
        ' 個人年齢区分
        If oEntry.Item("区分") = "年齢区分" Then
            nColumn = VLookupArea(oLine.Item("種目番号"), "市民種目区分", "タイプ")
            Cells(nRow, Range(sTable & "[区分]").Column).Value = _
                Application.WorksheetFunction.VLookup(oEntry.Item("年齢"), GetRange("市民年齢区分"), nColumn, False)
            Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = _
                VLookupArea(oEntry.Item("年齢"), "市民年齢区分", "ソート")
        ' 個人中高
        Else
            Cells(nRow, Range(sTable & "[種目区分]").Column).Value = oEntry.Item("区分")
            Cells(nRow, Range(sTable & "[区分]").Column).Value = oEntry.Item("区分")
            Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
        End If
    
    ElseIf CStr(sGame) = "横須賀マスターズ大会" Then
        
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = ""
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = oEntry.Item("年齢")
        Cells(nRow, Range(sTable & "[区分]").Column).Value = _
            VLookupArea(oEntry.Item("年齢"), "学マ年齢区分", "M年齢区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = _
            VLookupArea(oEntry.Item("年齢"), "学マ年齢区分", "M年齢区分")

    ElseIf CStr(sGame) = "横須賀市学童水泳競技大会" Then
        
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = oLine.Item("種目区分")
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = ""
        Cells(nRow, Range(sTable & "[区分]").Column).Value = _
            VLookupArea(oEntry.Item("区分"), "学マ学年表示", "学年表示")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    
    End If
    
End Sub

'
' リレー申込み行を出力
'
' sTable        IN      テーブル名
' nRow          IN      出力行番号
' nTeamNo       IN      チーム番号
' sGame         IN      大会名
' sTeam         IN      チーム名
' oEntry        IN      エントリー
' oLine         IN      種目、申込み時間
'
Sub WriteRelayLine( _
    sTable As String, _
    nRow As Integer, _
    nTeamNo As Integer, _
    sGame As Variant, _
    sTeam As Variant, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[個人No]").Column).Value = nTeamNo
    Cells(nRow, Range(sTable & "[チーム名]").Column).Value = CStr(sTeam)
    
    Cells(nRow, Range(sTable & "[プロNo]").Column).Value = oLine.Item("種目番号")
    
    Dim sMasterName As String
    sMasterName = GetMaster(CStr(sGame))
    
    Cells(nRow, Range(sTable & "[種目区分]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "区分")
    
    Cells(nRow, Range(sTable & "[性別]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "性別")
    
    Cells(nRow, Range(sTable & "[距離]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "距離")
    
    Cells(nRow, Range(sTable & "[種目名]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "種目")

    Cells(nRow, Range(sTable & "[申込み時間]").Column).Value = oLine.Item("申込み時間")
    If oLine.Item("申込み時間") >= 10000 Then
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    If CStr(sGame) = "横須賀選手権水泳大会" Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oLine.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    
    ElseIf CStr(sGame) = "横須賀市民体育大会" Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oLine.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = oLine.Item("区分")
    
    ElseIf CStr(sGame) = "横須賀マスターズ大会" Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oLine.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = oLine.Item("区分")
    
    ElseIf CStr(sGame) = "横須賀市学童水泳競技大会" Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = "小学"
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    End If
    
End Sub

'
' シートのテーブルをソートする
'
' 第１キー  プロNo      昇順
' 第２キー  ソート区分  昇順
' 第３キー  申込み時間  昇順
'
' oWorkSheet    IN      ワークシート
' sTableName    OUT     テーブル名
'
Public Sub SortByProNo(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate

    With ActiveSheet.ListObjects(sTableName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range(sTableName + "[プロNo]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range(sTableName + "[ソート区分]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range(sTableName + "[申込み時間]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

