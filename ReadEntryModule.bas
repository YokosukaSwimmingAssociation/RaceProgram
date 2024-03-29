Attribute VB_Name = "ReadEntryModule"
Option Explicit    ''←変数の宣言を強制する

'
' エントリーファイル一覧の読み込み
'
' 特定のフォルダを指定して、中にあるエントリーファイルを
' すべて読み込み一覧シートに出力する。
'
Public Sub エントリー読込み(Optional sPathName As String = "")
    ' エクセルシートを選択
    Call SheetActivate(エントリーシート)

    ' 出力用ワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' 出力用ワークシート
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    oWorkSheet.Select
    Call SetForcusTop

    ' イベント発生を抑制
    Call EventChange(False)

    ' エントリー一覧読込用配列
    Dim oGameList As Object
    Set oGameList = CreateObject("Scripting.Dictionary")

    ' エントリーテーブルを初期化
    Call DeleteTable(oWorkSheet, エントリーテーブル)

    ' エントリーファイル読込み
    Call ReadEntryFiles(oGameList, sPathName)

    ' エントリーシートの書き込み
    Call WriteEntrySheet(oWorkSheet, エントリーテーブル, oGameList)
    
    ' ProNo、ソート区分、申込み時間でソート
    Call SetTitleMenu("並び替え中")
    Call SortByProNo(oWorkSheet, エントリーテーブル)

    Call SetTitleMenu("")
    
    ' シートを保存
    oWorkBook.Save

End Sub

'
' テーブルを初期化
'
Public Sub エントリー一覧初期化()
    ' エントリーテーブルを初期化
    Call DeleteTable(Sheets(エントリーシート), エントリーテーブル)
End Sub

'
' エントリーファイル一覧の読み込み
'
' フォルダを指定して、その中に含まれるエントリーシート（*.xlsx）をすべて詠み込む
'
' oGameList     OUT     エントリー一覧
' sDirName      OUT     フォルダ名
'
Private Sub ReadEntryFiles(ByRef oGameList As Object, Optional sPathName As String = "")

    ' ファイル一覧を取得
    '
    If sPathName = "" Then
        sPathName = SelectDir()
    End If
    Dim FileList As Collection
    Set FileList = GetFiles(sPathName, "\*.xlsx")

    Dim nMax As Integer
    nMax = FileList.Count
    Dim nCount As Integer
    nCount = 0

    '
    ' ファイル毎に処理する
    '
    Dim SubBook As Workbook
    Dim vFile As Variant
    For Each vFile In FileList
        
        ' タイトル修正
        nCount = nCount + 1
        Call SetTitleMenu("プログラム読込中: " & Str(nCount) & "/" & Str(nMax))
        
        '
        ' ファイルを開く（読み取り専用）
        '
        Set SubBook = Workbooks.Open(Filename:=sPathName + "\" + vFile, ReadOnly:=True)
        SheetActivate ("記入票")

        ' エントリー一覧の読込み
        Call ReadEntryFile(oGameList)
    
        ' 警告なしでファイルを閉じる（保存しない）
        Application.DisplayAlerts = False
        SubBook.Close
        Application.DisplayAlerts = True
    Next vFile
    
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
' ├─種目名：Range("種目")
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
Private Sub ReadEntryFile(ByRef oGameList As Object)

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

    ' 個人用エントリーの読込み
    Call ReadPersonEntry(nNum, sTeamName, oEntryList)

    ' リレー用エントリーの読込み
    ' ※リレー無しの特別対応(7/3戻す)
    Call ReadRelayEntry(nNum, oEntryList)

End Sub

'
' 個人用エントリーの読込み
'
' nNum          IN      個人エントリー行(1,2)
' sTeam         IN      所属
' oEntryList    OUT     種目行
'
Private Sub ReadPersonEntry(nNum As Integer, sTeamName As String, ByRef oEntryList As Object)
    ' 個人番号範囲をすべて読み込む
    Dim oCell As Range
    For Each oCell In GetRange("選手番号")
        
        ' 個人番号は結合されていて
        If oCell.MergeCells Then
            ' 結合の先頭行で処理する
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

                ' 選手情報を登録
                Call ReadEntrySwimmer(nNum, oCell, oEntry)
                
                ' １行目
                Call ReadEntryLine(1, oCell.Row, oEntry)
                Call CheckEntry(1, oCell.Row, oEntry)
    
                ' ２行目
                Call ReadEntryLine(2, oCell.Row, oEntry)
                Call CheckEntry(2, oCell.Row, oEntry)
                If oEntry.Item("選手名") <> "" Then
                    oEntryList.Add nNum, oEntry
                End If
            End If
        End If
    Next oCell
End Sub

'
' エントリーファイルの個人情報の読込み
'
' 性別、選手名、フリガナ、区分等を読み込む
'
' nNum          IN      個人エントリー行(1,2)
' nRow          IN      行数
' oEntry        OUT     種目行
'
Private Sub ReadEntrySwimmer(nNum As Integer, oCell As Range, ByRef oEntry As Object)

    oEntry.Add "性別", GetOffset(oCell, GetRange("選手性別").Column).Value + "子"
    oEntry.Add "フリガナ", ReplaceName(GetOffset(oCell, GetRange("選手フリガナ").Column).Value)
    
    If Range("大会名").Value = 選手権大会 Then
        
        oEntry.Add "選手名", ReplaceName(GetOffset(oCell, GetRange("選手名").Column).Offset(1).Value)
        oEntry.Add "区分", GetOffset(oCell, GetRange("選手区分").Column).Value
    
    ElseIf Range("大会名").Value = 市民大会 Then
        oEntry.Add "選手名", ReplaceName(GetOffset(oCell, GetRange("選手名").Column).Offset(1).Value)
        oEntry.Add "学校名", Trim(GetOffset(oCell, GetRange("選手学校名").Column).Offset(2).Value)
        If GetOffset(oCell, GetRange("選手区分").Column).Value <> "" Then
            oEntry.Add "区分", GetOffset(oCell, GetRange("選手区分").Column).Value
        Else
            oEntry.Add "区分", "年齢区分"
        End If
        oEntry.Add "年齢", GetOffset(oCell, GetRange("選手年齢").Column).Offset(1).Value
    
    ElseIf Range("大会名").Value = マスターズ大会 Then
    
        oEntry.Add "選手名", ReplaceName(GetOffset(oCell, GetRange("選手名").Column).Offset(1).Value)
        oEntry.Add "年齢", GetOffset(oCell, GetRange("選手年齢").Column).Value
    
    ElseIf Range("大会名").Value = 室内記録会 Then
    
        oEntry.Add "選手名", ReplaceName(GetOffset(oCell, GetRange("選手名").Column).Offset(1).Value)
        oEntry.Add "年齢", GetOffset(oCell, GetRange("選手年齢").Column).Value
        oEntry.Add "区分", GetOffset(oCell, GetRange("選手学年").Column).Offset(1).Value
        oEntry.Add "検定", GetOffset(oCell, GetRange("選手検定").Column).Value
    
    Else
    
        oEntry.Add "選手名", ReplaceName(GetOffset(oCell, GetRange("選手名").Column).Offset(1).Value)
        oEntry.Add "区分", GetOffset(oCell, GetRange("選手学年").Column).Value
    
    End If

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
Private Sub ReadEntryLine(nNum As Integer, nRow As Integer, oEntry As Object)
    
    Dim oLines As Object
    Dim sStyle As String
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    Dim oProNo As Range
    
    ' 番号範囲をすべて読み込む
    Dim oCell As Range
    For Each oCell In GetRange("種目一覧")
        If oCell.Value <> "" Then
            sStyle = oCell.Value
        End If
        ' 種目選択が空以外の場合は選択されたものとする
        Set oProNo = GetRowOffset(oCell, nRow).Offset(nNum - 1)
        If Trim(oProNo.Value) <> "" Then
            Set oLines = CreateObject("Scripting.Dictionary")
            oEntry.Add nNum, oLines
            
            oLines.Add "種目番号", VLookupArea(oProNo.Value, "種目番号区分", "種目番号")
            oLines.Add "種目区分", VLookupArea(oProNo.Value, "種目番号区分", "種目区分") ' 今年だけ区分にしている
            oLines.Add "種目", ReplaceStyle(sStyle)
            oLines.Add "距離", ReplaceDistance(GetRowOffset(oCell, GetRange("種目距離").Row).Value)
            nMin = GetOffset(oProNo, GetRange("選手分").Column).Value
            nSec = GetOffset(oProNo, GetRange("選手秒").Column).Value
            nMil = GetOffset(oProNo, GetRange("選手ミリ秒").Column).Value
            oLines.Add "申込み時間", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
            Exit Sub
        End If
    Next oCell
End Sub

'
' エントリーの種目番号が正しいかを確認
'
' nNum          IN      個人エントリー行(1,2)
' nRow          IN      行番号
' oEntry        IN      種目行
'
Private Sub CheckEntry(nNum As Integer, nRow As Integer, oEntry As Object)
    
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
    
    If sGender <> oEntry.Item("性別") Or sDistance <> oLines.Item("距離") Or sStyle <> oLines.Item("種目") Then
        MsgBox CStr(nRow) & "行目：種目番号が正しくありません。：" & oLines.Item("種目番号")
        End
    End If

End Sub

'
' リレー種目の読込み
'
' 種目名、距離、申込み時間を取得する
'
' nNum          IN      エントリー行
' oEntryList    OUT     エントリー一覧
'
Private Sub ReadRelayEntry(nNum As Integer, ByRef oEntryList As Object)

    ' リレー種目番号範囲をすべて読み込む
    Dim nRelayNum As Integer
    nRelayNum = 0
    Dim oRelayEntry As Object
    Set oRelayEntry = Nothing
    Dim oCell As Range
    For Each oCell In GetRange("リレー種目")
        ' 値が設定されている場合は読み込む
        If oCell.Value <> "" Then
            ' リレーのエントリーリスト
            If oRelayEntry Is Nothing Then
                Set oRelayEntry = CreateObject("Scripting.Dictionary")
                oEntryList.Add nNum, oRelayEntry
            End If

            nRelayNum = nRelayNum + 1
            Call ReadRelayEntryLine(nRelayNum, oCell, oRelayEntry)
        End If
    Next oCell

End Sub

'
' エントリーファイルのリレー種目行読み込み
'
' nNum          IN    リレー番号
' oCell         IN    カレントセル
' oRelayEntry   I/O   種目行
'
Private Sub ReadRelayEntryLine(nNum As Integer, oCell As Range, oRelayEntry As Object)
    
    Dim oLines As Object
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    If oCell.Value <> "" Then
        Dim oRelayLines As Object
        Set oRelayLines = CreateObject("Scripting.Dictionary")
        oRelayEntry.Add "L" + Str(nNum), oRelayLines

        oRelayLines.Add "種目番号", oCell.Value
        If IsNameExists("リレー区分") Then
            oRelayLines.Add "区分", GetOffset(oCell, GetRange("リレー区分").Column).Value
        End If
        nMin = GetOffset(oCell, GetRange("リレー分").Column).Value
        nSec = GetOffset(oCell, GetRange("リレー秒").Column).Value
        nMil = GetOffset(oCell, GetRange("リレーミリ秒").Column).Value
        oRelayLines.Add "申込み時間", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
    End If
End Sub

'
' 種目名称の置換
'
' sStyle        IN      種目
'
Private Function ReplaceStyle(sStyle) As String
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
Private Function ReplaceDistance(sDistance) As String
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
Private Function ReplaceName(sName) As String
    
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
' 申込みをシートに出力
'
' oWorkSheet    IN     ワークシート
' sTable        IN     テーブル名
' oTeamList     IN     読み込んだチーム申込み一覧
'
Private Sub WriteEntrySheet(oWorkSheet As Worksheet, sTable As String, oGameList As Object)

    oWorkSheet.Select

    Dim oTeamList As Object
    Dim oEntryList As Object

    Dim nTeamToal As Integer
    nTeamToal = 0
    Dim nTeamNo As Integer
    nTeamNo = 1
    Dim nRow As Integer
    nRow = 1
    
    Dim vGame As Variant
    For Each vGame In oGameList.Keys()
        ' 出力する大会か判定
        If IsSameGame(CStr(vGame)) Then
            Set oTeamList = oGameList.Item(CStr(vGame))
        Else
            ' 出力する大会以外は捨てる
            Set oTeamList = CreateObject("Scripting.Dictionary")
        End If
        
        nTeamToal = nTeamToal + oTeamList.Count
        
        Dim vTeam As Variant
        For Each vTeam In oTeamList.Keys()
            Set oEntryList = oTeamList.Item(vTeam)
            
            ' 進捗表示
            Call SetTitleMenu("プログラム出力: " & CStr(nTeamNo) & "/" & CStr(nTeamToal))
            
            ' チームごとに出力する
            Call WriteTeamEntry(sTable, nRow, CStr(vGame), CStr(vTeam), nTeamNo, oEntryList)
            
            ' チーム番号をインクリメント
            nTeamNo = nTeamNo + 1
        Next
    Next
End Sub

'
' 出力するゲームか判定
'
' sGame         IN      大会名
'
Private Function IsSameGame(sGameName As String) As Boolean
    If GetRange("大会名").Value = sGameName Then
        IsSameGame = True
    ElseIf GetRange("大会名").Value = 学マ大会 Then
        If sGameName = 学童大会 Or sGameName = マスターズ大会 Then
            IsSameGame = True
        Else
            IsSameGame = False
        End If
    Else
        IsSameGame = False
    End If

End Function

'
' チームの出力をする
'
' sTable        IN      テーブル名
' nRow          IN/OUT  出力する行数
' sGame         IN      大会名
' sTeam         IN      チーム名
' nTeamNo       IN      チーム番号
' oEntryList    IN      エントリー一覧
'
Private Sub WriteTeamEntry(sTable As String, ByRef nRow As Integer, _
sGame As String, sTeam As String, _
nTeamNo As Integer, oEntryList As Object)
    Dim oEntry As Object
    Dim oLine As Object
    Dim nPersonNo As Integer
    
    Dim vNum As Variant
    For Each vNum In oEntryList.Keys()
        Set oEntry = oEntryList.Item(vNum)
        nPersonNo = nTeamNo * 100 + CInt(vNum)
        
        If oEntry.Exists("選手名") Then
            ' 個人
            Call WriteTeamPersonLine(sTable, nRow, sGame, sTeam, nPersonNo, oEntry, oLine)
        Else
            ' リレー
            Call WriteTeamRelayLine(sTable, nRow, sGame, sTeam, nTeamNo, oEntry, oLine)
        End If
    Next vNum
End Sub

'
' 個人の出力をする
'
' sTable        IN      テーブル名
' nRow          IN/OUT  出力する行数
' sGame         IN      大会名
' sTeam         IN      チーム名
' nPersonNo     IN      個人番号
' oEntry        IN      エントリー
' oLine         IN      種目、申込み時間
'
Private Sub WriteTeamPersonLine(sTable As String, ByRef nRow As Integer, _
sGame As String, sTeam As String, nPersonNo As Integer, _
oEntry As Object, oLine As Object)
    Dim i As Integer
    For i = 1 To 個人最大行数
        If Not IsEmpty(oEntry.Item(i)) Then
            nRow = nRow + 1
            Set oLine = oEntry.Item(i)
            Call WriteLine(sTable, nRow, sGame, sTeam, nPersonNo, oEntry, oLine)
        End If
    Next i
End Sub

'
' リレーの出力をする
'
' sTable        IN      テーブル名
' nRow          IN/OUT  出力する行数
' sGame         IN      大会名
' sTeam         IN      チーム名
' nTeamNo       IN      チーム番号
' oEntry        IN      エントリー
' oLine         IN      種目、申込み時間
'
Private Sub WriteTeamRelayLine(sTable As String, ByRef nRow As Integer, _
sGame As String, sTeam As String, nTeamNo As Integer, _
oEntry As Object, oLine As Object)
    Dim i As Integer
    Dim sKey As String
    For i = 1 To リレー最大行数
        sKey = "L" & Str(i)
        If oEntry.Exists(sKey) Then
            nRow = nRow + 1
            Set oLine = oEntry.Item(sKey)
            Call WriteRelayLine(sTable, nRow, sGame, sTeam, nTeamNo, oEntry, oLine)
        End If
    Next i
End Sub


'
' 申込み行を出力
'
' sTable        IN      テーブル名
' nRow          IN      出力行番号
' sGame         IN      大会名
' sTeam         IN      チーム名
' nPersonNo     IN      選手番号
' oEntry        IN      エントリー
' oLine         IN      種目、申込み時間
'
Private Sub WriteLine( _
    sTable As String, _
    nRow As Integer, _
    sGame As String, _
    sTeam As String, _
    nPersonNo As Integer, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[個人No]").Column).Value = nPersonNo
    Cells(nRow, Range(sTable & "[プロNo]").Column).Value = oLine.Item("種目番号")
    Cells(nRow, Range(sTable & "[チーム名]").Column).Value = sTeam
    Cells(nRow, Range(sTable & "[選手名]").Column).Value = oEntry.Item("選手名")
    Cells(nRow, Range(sTable & "[フリガナ]").Column).Value = oEntry.Item("フリガナ")
    Cells(nRow, Range(sTable & "[性別]").Column).Value = oEntry.Item("性別")
    Cells(nRow, Range(sTable & "[距離]").Column).Value = oLine.Item("距離")
    Cells(nRow, Range(sTable & "[種目]").Column).Value = oLine.Item("種目")
    Cells(nRow, Range(sTable & "[申込み時間]").Column).Value = oLine.Item("申込み時間")
    If oLine.Item("申込み時間") >= 10000 Then
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    If sGame = 選手権大会 Then
    
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = ""
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = ""
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oEntry.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    
    ElseIf sGame = 市民大会 Then
    
        Cells(nRow, Range(sTable & "[学校名]").Column).Value = oEntry.Item("学校名")
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = oEntry.Item("年齢")
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = oEntry.Item("区分")
        
        ' 個人年齢区分
        If oEntry.Item("区分") = "年齢区分" Then
            Dim nColumn As Integer
            nColumn = VLookupArea(oLine.Item("種目番号"), "市民種目区分", "タイプ")
            Dim sClass As String
            sClass = Application.WorksheetFunction.VLookup(oEntry.Item("年齢"), GetRange("市民年齢区分"), nColumn, False)
            Cells(nRow, Range(sTable & "[区分]").Column).Value = sClass
            If sClass = "一般" Then
                Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = "20"
            Else
                Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = Left(sClass, 2)
            End If
        ' 個人中高
        Else
            Cells(nRow, Range(sTable & "[区分]").Column).Value = oEntry.Item("区分")
            Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
        End If
    
    ElseIf sGame = マスターズ大会 Then
        
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = ""
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = oEntry.Item("年齢")
        Cells(nRow, Range(sTable & "[区分]").Column).Value = _
            VLookupArea(oEntry.Item("年齢"), "学マ年齢区分", "M年齢区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = _
            VLookupArea(oEntry.Item("年齢"), "学マ年齢区分", "M年齢区分")

    ElseIf sGame = 学童大会 Then
        
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = oLine.Item("種目区分")
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = ""
        Cells(nRow, Range(sTable & "[区分]").Column).Value = _
            VLookupArea(oEntry.Item("区分"), "学マ学年表示", "学年表示")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    
    ElseIf sGame = 室内記録会 Then
        
        Cells(nRow, Range(sTable & "[種目区分]").Column).Value = ""
        Cells(nRow, Range(sTable & "[年齢]").Column).Value = oEntry.Item("年齢")
        'Cells(nRow, Range(sTable & "[区分]").Column).Value = _
        '    VLookupArea(oEntry.Item("年齢") & "_" & oEntry.Item("区分"), "記録会年齢区分", "区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = _
            VLookupArea(oEntry.Item("年齢") & "_" & oEntry.Item("区分"), "記録会年齢区分", "ソート")
        Cells(nRow, Range(sTable & "[検定]").Column).Value = oEntry.Item("検定")
    
    End If
    
End Sub

'
' リレー申込み行を出力
'
' sTable        IN      テーブル名
' nRow          IN      出力行番号
' sGame         IN      大会名
' sTeam         IN      チーム名
' nTeamNo       IN      チーム番号
' oEntry        IN      エントリー
' oLine         IN      種目、申込み時間
'
Private Sub WriteRelayLine( _
    sTable As String, _
    nRow As Integer, _
    sGame As String, _
    sTeam As String, _
    nTeamNo As Integer, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[個人No]").Column).Value = nTeamNo
    Cells(nRow, Range(sTable & "[チーム名]").Column).Value = sTeam
    
    Cells(nRow, Range(sTable & "[プロNo]").Column).Value = oLine.Item("種目番号")
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGame)
    
    Cells(nRow, Range(sTable & "[種目区分]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "種目区分")
    
    Cells(nRow, Range(sTable & "[性別]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "性別")
    
    Cells(nRow, Range(sTable & "[距離]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "距離")
    
    Cells(nRow, Range(sTable & "[種目]").Column).Value = _
        VLookupArea(oLine.Item("種目番号"), sMasterName, "種目")

    Cells(nRow, Range(sTable & "[申込み時間]").Column).Value = oLine.Item("申込み時間")
    If oLine.Item("申込み時間") >= 10000 Then
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[申込み時間]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    If sGame = 選手権大会 Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oLine.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    
    ElseIf sGame = 市民大会 Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oLine.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = oLine.Item("区分")
    
    ElseIf sGame = マスターズ大会 Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = oLine.Item("区分")
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = oLine.Item("区分")
    
    ElseIf sGame = 学童大会 Then
        Cells(nRow, Range(sTable & "[区分]").Column).Value = "小学"
        Cells(nRow, Range(sTable & "[ソート区分]").Column).Value = ""
    End If
    
End Sub

'
' エントリーテーブルを初期化
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
'
Public Sub DeleteTable(oWorkSheet As Worksheet, sTableName As String)
    Dim myTable As ListObject
    Set myTable = oWorkSheet.ListObjects(sTableName)
    If Not (myTable.DataBodyRange Is Nothing) Then
        myTable.DataBodyRange.Delete
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
' sTableName    IN      テーブル名
'
Public Sub SortByProNo(oWorkSheet As Worksheet, sTableName As String)

    With oWorkSheet.ListObjects(sTableName).Sort
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

