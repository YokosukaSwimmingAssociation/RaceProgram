Attribute VB_Name = "CommonModule"
Option Explicit    ''←変数の宣言を強制する

' 大会名
Public Const 学童大会 As String = "横須賀市学童水泳競技大会"
Public Const マスターズ大会 As String = "横須賀マスターズ大会"
Public Const 学マ大会 As String = "学童マスターズ大会"
Public Const 市民大会 As String = "横須賀市民体育大会"
Public Const 選手権大会 As String = "横須賀選手権水泳大会"
Public Const 室内記録会 As String = "短水路室内水泳記録会"

Public Const トップページシート As String = "トップページ"
Public Const エントリーシート As String = "エントリー一覧"
Public Const エントリーテーブル As String = "エントリーテーブル"
Public Const プログラムシート As String = "プログラム"
Public Const フォーマットシート As String = "プログラムフォーマット"
Public Const 記録画面シート As String = "記録画面"
Public Const 設定各種シート As String = "設定各種"

Public Const 平均分け組数 As Integer = 3        ' 平均分け方式にする組数
Public Const 個人最大行数 As Integer = 2        ' 個人の申込み行数
Public Const リレー最大行数 As Integer = 24     ' リレーの最大申込み行数
Public Const ページレース数 As Integer = 5      ' １ページのレース数

Public Const 選手名ブランク As String = "　　．　　．　　．"
Public Const タイムブランク As String = "　　：　　．  "
Public Const 順位ブランク As String = "　　"
Private Const ARRAYSIZE = 10000

Public Type RaceNumber
    nProNo As Integer
    nRance As Integer
End Type

'
' イベントの発生、画面描画のOn/Off
'
' bFlag     IN      True：再開／False：抑制
'
Public Sub EventChange(bFlag As Boolean)
    With Application
        If bFlag Then
            .EnableEvents = True                    ' イベントの発生を再開する
            .ScreenUpdating = True                  ' 描画の更新を行う
            .Calculation = xlCalculationAutomatic   ' セル値の自動計算
        Else
            .EnableEvents = False                   ' イベント抑制する
            .ScreenUpdating = False                 ' 描画の更新を抑制する
            .Calculation = xlCalculationManual      ' セル値の手動計算
        End If
    End With
End Sub

'
' シートの保護
'
' シートの保護／解除を行う。
' 設定前のシートの保護状態を返す。
'
' bFlag         IN      True：保護／False：解除
' oWorkSheet    IN      保護するシートオブジェクト
'
Public Function SheetProtect(bFlag As Boolean, Optional oWorkSheet As Worksheet = Nothing) As Boolean

    If oWorkSheet Is Nothing Then
        Set oWorkSheet = ActiveSheet
    End If

    If ActiveSheet.ProtectContents = True Then
        SheetProtect = True
    Else
        SheetProtect = False
    End If

    If bFlag Then
        oWorkSheet.Protect DrawingObjects:=True, Contents:=True, _
            Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Else
        oWorkSheet.Unprotect
    End If

End Function

'
' オートフィルタの設定
'
' sName         IN      範囲の名前
' bFlag         IN      True：フィルタ表示／False：解除
'
Public Sub SetAutoFilter(sName As String, bFlag As Boolean)
    If GetRange(sName).Parent.AutoFilterMode <> bFlag Then
        GetRange(sName).AutoFilter
    End If
End Sub

'
' セルを選択
'
Public Sub SetForcusTop(Optional sRange As String = "$A$1")
    Range(sRange).Select
End Sub

'
' セルを選択
'
Public Sub SetForcus(sAddress As String)
    Range(sAddress).Select
End Sub

'
' タイトルバーの変更
'
' sTitle    IN      タイトルバー文字列
'
Public Sub SetTitleMenu(sTitle As String)
    Dim bFlag As Boolean
    bFlag = Application.EnableEvents
    If bFlag = False Then
        Call EventChange(True)
    End If
    Application.Caption = sTitle
    DoEvents
    If bFlag = False Then
        Call EventChange(False)
    End If
End Sub

'
' 大会名から種目区分を返す
'
Public Function GetMaster(sGameName As String)
    
    GetMaster = VLookupArea(sGameName, "設定各種", "種目区分範囲名")

End Function

'
' 空白置換
'
' 全角空白、連続した空白を半角空白１つに変換
'
' sStr          IN      文字列
'
Public Function STrim(sStr) As String
    Dim sTemp As String
    sTemp = sStr
    sTemp = Replace(sTemp, "　", " ")
    
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")
    
    '正規表現の指定
    With oReg
        .Pattern = "[ ]+"     'パターンを指定
        .IgnoreCase = False     '大文字と小文字を区別するか(False)、しないか(True)
        .Global = True          '文字列全体を検索するか(True)、しないか(False)
    End With
    sTemp = oReg.Replace(sTemp, " ")
    
    STrim = RTrim(LTrim(sTemp))
End Function

'
' 空白置換
'
' 全角空白、半角空白をなくす
'
' sStr          IN      文字列
'
Public Function STrimAll(sStr As String) As String
    STrimAll = Replace(Replace(sStr, "　", ""), " ", "")
End Function


'
' シートの存在チェック付きアクティベート
'
' シートの存在を確認しVisibleがTrueでなければTrueにしてから
' Activateする
'
' sSheetName    IN      シート名
'
Public Function SheetActivate(sSheetName As String) As Variant
    If IsSheetExists(sSheetName) Then
        If Worksheets(sSheetName).Visible <> True Then
            Worksheets(sSheetName).Visible = True
        End If
        Worksheets(sSheetName).Activate
        Set SheetActivate = ActiveSheet
    Else
        MsgBox "「" & sSheetName & "」シートが存在しません。" & vbCrLf & _
                "正しいファイルをお使いください。", vbOKOnly
        End
    End If
End Function


'
' シートのVisibleを返す
'
' xlSheetVisible/True(-1)
' xlSheetHidden/False(0)
' xlSheetVeryHidden(2)
' Empty(3)
'
Public Function GetSheetVisible(sSheetName As String) As Variant
    If IsSheetExists(sSheetName) Then
        GetSheetVisible = Worksheets(sSheetName).Visible
    Else
        GetSheetVisible = 3
    End If
End Function


'
' シートの存在チェック付き非表示
'
'
' sSheetName    IN      シート名
'
Public Sub SheetVisible(sSheetName As String, vVisible As Variant)
    If IsSheetExists(sSheetName) Then
        Worksheets(sSheetName).Visible = vVisible
    End If
End Sub


'
' シートの存在チェック
'
' sSheetName        IN      シート名
'
Public Function IsSheetExists(sSheetName As String) As Boolean
    IsSheetExists = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = sSheetName Then
            IsSheetExists = True
        End If
    Next ws
End Function

'' Dictionaryを参照引数にし、これをソートする破壊的プロシージャ。
'' Variantの二次元配列のまま扱いたかったが、引数にできないようなので配列を使用している
'' KeyもValueもLong型のみを許容
Public Sub DictQuickSort(ByRef dic As Object, Optional sIndex As String = "Key")
  
    Dim nIndex As Integer
    If sIndex = "Key" Then
        nIndex = 0
    Else
        nIndex = 1
    End If
  
    Dim i As Long
    Dim j As Long
    Dim dicSize As Long
    Dim varTmp(ARRAYSIZE, 2) As Long
  
    dicSize = dic.Count
  
    ' Dictionaryが空か、サイズが1以下であればソート不要
    If dic Is Nothing Or dicSize < 2 Then
        Exit Sub
    End If
  
    ' Dictionaryから二元配列に転写
    i = 0
    Dim Key As Variant
    For Each Key In dic
        varTmp(i, 0) = Key
        varTmp(i, 1) = dic(Key)
        i = i + 1
    Next
  
    'クイックソート
    Call QuickSort(varTmp, 0, dicSize - 1, nIndex)
  
    dic.RemoveAll
  
    For i = 0 To dicSize - 1
        dic(varTmp(i, 0)) = varTmp(i, 1)
    Next
End Sub


'' Long型の二次元配列を受け取り、これの２列目（Value）でクイックソートする
Private Sub QuickSort(ByRef targetVar() As Long, ByVal min As Long, ByVal max As Long, nIndex As Integer)
    Dim i, j As Long
    Dim tmp As Long
    Dim pivot As Long
    
    If min < max Then
        i = min
        j = max
        pivot = med3(targetVar(i, nIndex), targetVar(Int(i + j / 2), nIndex), targetVar(j, nIndex))
        Do
            Do While targetVar(i, nIndex) < pivot
                i = i + 1
            Loop
            Do While pivot < targetVar(j, nIndex)
                j = j - 1
            Loop
            If i >= j Then Exit Do
            
            tmp = targetVar(i, 0)
            targetVar(i, 0) = targetVar(j, 0)
            targetVar(j, 0) = tmp
        
            tmp = targetVar(i, 1)
            targetVar(i, 1) = targetVar(j, 1)
            targetVar(j, 1) = tmp
        
            i = i + 1
            j = j - 1
        
        Loop
        Call QuickSort(targetVar, min, i - 1, nIndex)
        Call QuickSort(targetVar, j + 1, max, nIndex)
        
    End If
End Sub

'' Long, y, z を辞書順比較し二番目のものを返す
Private Function med3(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    If x < y Then
        If y < z Then
            med3 = y
        ElseIf z < x Then
            med3 = x
        Else
            med3 = z
        End If
    Else
        If z < y Then
            med3 = y
        ElseIf x < z Then
            med3 = x
        Else
            med3 = z
        End If
    End If
End Function

'
' 配列に値が存在するか確認する
'
' vAry          IN  配列
' vValue        IN  値
'
Public Function isAryExist(vAry As Variant, vValue As Variant) As Boolean

    Dim vTemp As Variant
    For Each vTemp In vAry
        If vTemp = vValue Then
            isAryExist = True
            Exit Function
        End If
    Next
    isAryExist = False

End Function

'
' 名前チェック付きRangeオブジェクト取得
'
' sName             IN      名前
'
Public Function GetRange(sName As String) As Range
    If IsNameExists(sName) Then
        Set GetRange = Range(sName)
    Else
        MsgBox "名前「" & sName & "」が定義されていません。" & vbCrLf & _
                "正しいファイルをお使いください。", vbOKOnly
        End
    End If
End Function

'
' 名前の存在確認
'
' sName             IN      名前
'
Public Function IsNameExists(sName As String) As Boolean
    IsNameExists = False
    Dim vName As Variant
    For Each vName In ActiveWorkbook.Names
        If vName.Name = sName Then
            IsNameExists = True
            Exit For
        End If
    Next
End Function

'
' 名前の定義
'
' sName             IN      名前
' sRange            IN      レンジ範囲(A1形式)
'
Public Sub DefineName(sName As String, sRange As String)
    If IsNameExists(sName) Then
        ActiveWorkbook.Names(sName).Delete
    End If
    ActiveWorkbook.Names.Add Name:=sName, RefersTo:="=" & sRange
    ActiveWorkbook.Names(sName).Comment = ""
End Sub

'
' 名前削除
'
' sRegStr           IN      削除する名前の文字列
'
Public Sub DeleteName(sRegStr As String)
    Dim vName As Variant
    For Each vName In ActiveWorkbook.Names
        If vName.Name Like sRegStr Then
            vName.Delete
        End If
    Next
End Sub

'
' 同一行の範囲取得
'
' 指定したセルから最右の範囲のアドレスを返す
'
' 例： $A$1 -> $A$1:$F$1
'
' sTopAddres IN      先頭のセルアドレス
'
Public Function ColumnRange(sTopAddress As String) As Range

    Set ColumnRange = Range(Range(sTopAddress), _
                    Range(sTopAddress).End(xlToRight))

End Function

'
' 同一行の範囲取得
'
' 指定したセルから最右の範囲のアドレスを返す
'
' 例： $A$1 -> $A$1:$F$1
'
' sTopAddres IN      先頭のセルアドレス
'
Public Function ColumnRangeAddress(sTopAddress As String) As String

    ColumnRangeAddress = ColumnRange(sTopAddress).Address

End Function

'
' 同一列の範囲取得
'
' 指定したセルから最下層の範囲のアドレスを返す
'
' 例： $A$1 -> $A$1:$A$50
'
' sTopAddres IN      先頭のセルアドレス
'
Public Function RowRange(sTopAddress As String) As Range
    Set RowRange = Range(Range(sTopAddress), Range(sTopAddress).End(xlDown))
End Function

'
' 同一列の範囲取得
'
' 指定したセルから最下層の範囲のアドレスを返す
'
' 例： $A$1 -> $A$1:$A$50
'
' sTopAddres IN      先頭のセルアドレス
'
Public Function RowRangeAddress(sTopAddress As String) As String

    RowRangeAddress = RowRange(sTopAddress).Address

End Function

'
' 行列の範囲取得
'
' 指定したセルから最下層、際右端範囲のアドレスを返す
'
' 例： $A$1 -> $A1$1:$F$50
'
' sTopAddres IN      先頭のセルアドレス
'
Public Function TableRange(sTopAddress As String) As Range

    Dim oRng As Range
    Set oRng = Range(Range(sTopAddress), Range(sTopAddress).End(xlDown))
    Set TableRange = Range(oRng, oRng.End(xlToRight))

End Function

'
' 行列の範囲取得
'
' 指定したセルから最下層、際右端範囲のアドレスを返す
'
' 例： $A$1 -> $A1$1:$F$50
'
' sTopAddres IN      先頭のセルアドレス
'
Public Function TableRangeAddress(sTopAddress As String) As String

    TableRangeAddress = TableRange(sTopAddress).Address

End Function

'
' 範囲の最左列番号を返す
'
' sName      IN      範囲名
'
Public Function GetAreaLeftColumn(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaLeftColumn = oRange.Column
End Function

'
' 範囲の最右列番号を返す
'
' sName      IN      範囲名
'
Public Function GetAreaRightColumn(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaRightColumn = oRange.Column + oRange.Columns.Count - 1
End Function

'
' 範囲の最上行番号を返す
'
' sName      IN      範囲名
'
Public Function GetAreaTopRow(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaTopRow = oRange.Row
End Function

'
' 範囲の最下行番号を返す
'
' sName      IN      範囲名
'
Public Function GetAreaBottomRow(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaBottomRow = oRange.Row + oRange.Rows.Count - 1
End Function

'
' 基準セルからオフセット位置の値セルを返す
'
' vCell         IN      基準セル
' nColumn       IN      列番号
'
Public Function GetOffset(vCell As Variant, nColumn As Integer, Optional nRow As Integer = 0) As Range
    Set GetOffset = vCell.Offset(nRow, nColumn - vCell.Column)
End Function

'
' 基準セルから列のオフセット位置の値セルを返す
'
' vCell         IN      基準セル
' nColumn       IN      列番号
'
Public Function GetRowOffset(vCell As Variant, nRow As Integer, Optional nColumn As Integer = 0) As Range
    Set GetRowOffset = vCell.Offset(nRow - vCell.Row, nColumn)
End Function

'
' 範囲から指定した名前の列番号を返す
' 存在しない場合はゼロを返す
'
' sName      IN      範囲名
' sColName   IN      カラム名
'
Public Function GetColIdx(sName As String, sColName As String) As Integer
    Dim nIndex As Integer
    nIndex = 1
    Dim oCell As Range
    For Each oCell In GetRange(sName).Rows(1).Columns
        If STrimAll(oCell.Value) = sColName Then
            GetColIdx = nIndex
            Exit Function
        End If
        nIndex = nIndex + 1
    Next oCell
    GetColIdx = 0
End Function

'
' 範囲のヘッダ行を除くキー列を取得する
'
' 範囲の１列目のヘッダ行を除く列を返す
'
' sName      IN      範囲名
'
Public Function GetAreaKeyData(sName As String) As Range
    Dim oRange As Range
    Set oRange = Range(sName)
    Set GetAreaKeyData = oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, 1).Rows()
End Function

'
' 範囲のキー列名を返す
'
' 範囲の１列目の名前を返す
'
' sName      IN      範囲名
'
Public Function GetAreaKeyName(sName As String) As Variant
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaKeyName = oRange.Cells(1, 1).Value
End Function

'
' 範囲から指定した値に対応するカラム値を返す
'
' vValue     IN      検索値
' sName      IN      範囲名
' sColName   IN      カラム名
' bFlag      IN      VLookuUp関数の検索の型（False:完全一致／True:一番近いデータ）
'            OUT     一致した値
'
Public Function VLookupArea(vValue As Variant, sName As String, sColName As String, Optional bFlag As Boolean = False) As Variant
On Error GoTo ErrorHandler_VLookupArea
    
    VLookupArea = Application.WorksheetFunction.VLookup(vValue, GetRange(sName), GetColIdx(sName, sColName), bFlag)
    Exit Function

ErrorHandler_VLookupArea:
    MsgBox "範囲名：" & sName & vbCrLf & _
            "項目名：" & sColName & vbCrLf & _
            "検索値：" & CStr(vValue) & vbCrLf & _
            "が見つかりません。シートを確認してください。", vbOKOnly
    End
End Function

'
' 結合エリアのどの部分かを返す
'
' 1:先頭
' 2:末尾
' 3:中間
'
' oCell     IN      セル
'
Public Function CheckMergeArea(oCell As Range)
    Dim nIdx As Integer
    nIdx = 1
    Dim nMax As Integer
    nMax = oCell.MergeArea.Rows.Count
    Dim vCell As Range
    For Each vCell In oCell.MergeArea.Rows
        If vCell.Address = oCell.Address Then
            If nIdx = 1 Then
                CheckMergeArea = 1
            ElseIf nIdx = nMax Then
                CheckMergeArea = 2
            Else
                CheckMergeArea = 3
            End If
        End If
        nIdx = nIdx + 1
    Next vCell
End Function

'
' 罫線を引く
'
Public Sub SetBorder(oRange As Range)
    With oRange
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub



'
' フォルダを選択するダイアログを表示
'
' 初期はマクロファイルと同じフォルダを開く
'
' SelectDir     OUT     フォルダ名フルパス
'
Public Function SelectDir()
    ' 新しいファイルを開く
    Dim FileSysObj As Object
    Dim sPathName As String
    Set FileSysObj = CreateObject("Scripting.FileSystemObject")
    sPathName = FileSysObj.GetParentFolderName(ActiveWorkbook.FullName)
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = sPathName + "\"
        .AllowMultiSelect = False
        If .Show = True Then
            SelectDir = .SelectedItems(1)
        Else
            MsgBox "処理を中止します。"
            End
        End If
    End With

End Function

'
' 特定のフォルダのファイルを開く
'
' sPathName     IN      フォルダパス
' sExt          IN      拡張子の指定
' cFileList     OUT     ファイルリスト
'
Public Function GetFiles(sPathName As String, sExt As String) As Collection
    Dim sFile As String
    Dim cFileList As Collection
    Set cFileList = New Collection
    sFile = Dir(sPathName & sExt)
    Do While sFile <> ""
        cFileList.Add Item:=sFile
        sFile = Dir()
    Loop
    Set GetFiles = cFileList
End Function

'
' プログラム補正用：セルの修正
'
' sRange     IN      セルのアドレス
' sValue     IN      変更する値
'
Public Sub ModCell(sRange As String, sValue As String)

    With Range(sRange)
        .Value = sValue
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

End Sub

'
' プログラム補正用：プログラム番号のアドレスを探す
'
' nProNo     IN      プログラム番号
' sName      IN      選手名
' sColName   IN      カラム名（レースNo、組、レーン）
'            OUT     アドレス文字列
'
Public Function SearchCell(nProNo As Integer, sName As String, sColName As String) As String

    Dim vCell As Range
    For Each vCell In GetAreaKeyData(エントリーテーブル & "[プロNo]")
        If GetOffset(vCell, Range(エントリーテーブル & "[プロNo]").Column).Value = nProNo And _
            GetOffset(vCell, Range(エントリーテーブル & "[選手名]").Column).Value = sName Then
            SearchCell = GetOffset(vCell, Range(エントリーテーブル & "[" & sColName & "]").Column).Address
            Exit Function
        End If
    Next vCell

End Function

'
' プログラム補正用：プログラム番号のアドレスを探す
'
' nProNo     IN      プログラム番号
' sTeam      IN      所属名
' sClass     IN      区分
' sColName   IN      カラム名（Prog氏名）
'            OUT     アドレス文字列
'
Public Function SearchRelayCell(nProNo As Integer, _
sTeam As String, sClass As String, sColName As String) As String

    Dim sName As String
    sName = "プログラム番号" & Trim(CStr(nProNo))

    Dim nLane As Integer
    Dim nOrder As Integer
    Dim vCell As Range
    
    If IsNameExists(sName) Then
        Dim vProNo As Range
        For Each vProNo In GetRange(sName)
            If GetOffset(vProNo, Range("Prog所属").Column).Value = sTeam And _
                GetOffset(vProNo, Range("Prog区分").Column).Value = sClass Then
                SearchRelayCell = GetOffset(vProNo, Range(sColName).Column).Address
                Exit Function
            End If
        Next vProNo
    End If

End Function

'
' プログラム入力用：記録画面の種目選択
'
' nProNo     IN      プログラム番号
' nHeat      IN      組
'
Public Sub SetRace(nProNo As Integer, nHeat As Integer)
    GetRange("記録画面種目番号").Value = nProNo
    GetRange("記録画面組").Value = nHeat
End Sub

'
' プログラム入力用：記録画面のタイム入力
'
' nIndex        IN      順位
' nLean         IN      レーン
' sTime         IN      時間
' sAdditional   IN      備考
'
Public Sub SetLean(nIndex As Integer, nLean As Integer, sTime As String, Optional sAdditional As String = "")
    GetRange("記録画面レーン").Rows(nIndex).Value = nLean
    GetRange("記録画面タイム").Rows(nIndex).Value = sTime
    If sAdditional <> "" Then
        GetRange("記録画面備考").Rows(nIndex).Value = sAdditional
    End If
End Sub


