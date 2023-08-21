Attribute VB_Name = "DefineProgramSheetModule"
'
' 名前を定義する
'
Sub ワークブック名前定義()
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    Call EventChange(False)
  
    Call Header名前定義(フォーマットシート)
    Call Prog名前定義(フォーマットシート)
    Call 記録画面名前定義(記録画面シート)
    Call 各種設定名前定義(設定各種シート)
    Call トップページ定義(トップページシート)
    Call シート非表示
    
    Call EventChange(True)
    oWorkSheet.Activate
End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Private Sub Header名前定義(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Range("$A$1").Select

    ' 名前をすべて削除
    Call DeleteName("Header*")

    Dim oCell As Range
    Dim sName As String
    For nColumn = 1 To ActiveCell.SpecialCells(xlCellTypeLastCell).Column
        Set oCell = Cells(1, nColumn)
        sName = STrimAll(oCell.Value)
        If sName <> "" Then
            Call DefineName("Header" & sName, oCell.Address(ReferenceStyle:=xlA1))
            If sName = "所属" Then
                Call DefineName("Header" & sName & "前", oCell.Offset(0, -1).Address(ReferenceStyle:=xlA1))
                Call DefineName("Header" & sName & "後", oCell.Offset(0, 1).Address(ReferenceStyle:=xlA1))
            End If
        End If
    Next nColumn

    ' シートのロック
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Private Sub Prog名前定義(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    Range("$A$1").Select

    ' 名前をすべて削除
    Call DeleteName("Prog*")

    ' プログラムヘッダ
    Call DefineName("ProgプロNo", "$C$3")
    Call DefineName("Prog種目区分", "$D$3")
    Call DefineName("Prog種目名", "$F$3")
    Call DefineName("Prog決勝", "$I$3")
    Call DefineName("Prog記録", "$K$3")

    ' 組ヘッダ
    Call DefineName("Prog組", "$C$4")
   
    ' レーンデータ
    Call DefineName("Prog組番", "$C$5")
    Call DefineName("Progレーン", "$D$5")
    Call DefineName("Prog氏名", "$E$5")
    Call DefineName("Prog種目", "$F$5")
    Call DefineName("Prog所属前", "$G$5")
    Call DefineName("Prog所属", "$H$5")
    Call DefineName("Prog所属後", "$I$5")
    Call DefineName("Prog区分", "$J$5")
    Call DefineName("Prog時間", "$K$5")
    Call DefineName("Prog順位", "$L$5")
    'Call DefineName("Prog検定", "$L$5")
    Call DefineName("Prog備考", "$M$5")
    Call DefineName("Prog大会記録", "$N$5")
    Call DefineName("Prog申込み記録", "$O$5")
    Call DefineName("ProgレースNo", "$P$5")
    Call DefineName("Progソート区分", "$Q$5")
    Call DefineName("Prog標準記録", "$R$5")

    ' 組ヘッダ
    Call DefineName("Prog組ヘッダフォーマット", "A$2:$R$3")
    Call DefineName("Prog組フォーマット", "A$4:$R$13")
     
    ' シートのロック
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' 記録画面の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 記録画面名前定義(sSheetName As String)
    ' アクティブ／解除
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 名前をすべて削除
    Call DeleteName("記録画面*")

    Call DefineName("記録画面種目番号", "$B$1")
    Call DefineName("記録画面種目名", "$C$1")
    Call DefineName("記録画面組", "$B$2")
    Call DefineName("記録画面レースNo", "$B$3")
    Call DefineName("記録画面レーン", "$B$5:$B$11")
    Call DefineName("記録画面タイム", "$C$5:$C$11")
    Call DefineName("記録画面選手名", "$D$5:$D$11")
    Call DefineName("記録画面チーム名", "$E$5:$E$11")
    Call DefineName("記録画面備考", "$F$5:$F$11")
    Call DefineName("記録画面違反", "$G$5:$G$11")

    Call 記録画面違反定義

    ' シートのロック
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' 記録画面の違反
'
' sValue        IN      ダミー
'
Private Sub 記録画面違反定義()
    With GetRange("記録画面違反").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="　,スタート失格,失格,OP"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' 各種シートの設定
'
' sSheetName        IN      シート名
'
Public Sub 各種設定名前定義(sSheetName As String)

    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 名前をすべて削除
    Call DeleteName("設定*")

    ' シートを定義
    Call DefineName("設定各種", TableRangeAddress("$A$1"))

    ' 各種設定を行う
    For Each vGameName In GetAreaKeyData("設定各種")
        If VLookupArea(vGameName, "設定各種", "対象") = 1 Then
        
            'Debug.Print vGameName
        
            ' 名前をすべて削除
            Call DeleteName(VLookupArea(vGameName, "設定各種", "変数名先頭") & "*")
        
            ' 種目区分の設定
            Call DefineTableRange(VLookupArea(vGameName, "設定各種", "種目区分シート名"), _
                                    VLookupArea(vGameName, "設定各種", "種目区分範囲名"))
            ' 種目区分の設定
            If VLookupArea(vGameName, "設定各種", "種目区分関数名") <> "" Then
                Application.Run VLookupArea(vGameName, "設定各種", "種目区分関数名"), _
                                    VLookupArea(vGameName, "設定各種", "種目区分シート名")
            End If
        
            ' 大会記録の設定
            If VLookupArea(vGameName, "設定各種", "大会記録関数名") <> "" Then
                ' 特殊関数を実施
                Application.Run VLookupArea(vGameName, "設定各種", "大会記録関数名"), _
                                    VLookupArea(vGameName, "設定各種", "大会記録シート名"), _
                                    VLookupArea(vGameName, "設定各種", "大会記録範囲名")
            ElseIf VLookupArea(vGameName, "設定各種", "大会記録シート名") <> "" Then
                Call DefineRecordSheet(VLookupArea(vGameName, "設定各種", "大会記録シート名"), _
                                    VLookupArea(vGameName, "設定各種", "大会記録範囲名"))
            End If
        
            ' 優勝者の設定
            If VLookupArea(vGameName, "設定各種", "大会記録関数名") <> "" Then
                ' 特殊関数を実施
                Application.Run VLookupArea(vGameName, "設定各種", "大会記録関数名"), _
                                    VLookupArea(vGameName, "設定各種", "優勝者シート名"), _
                                    VLookupArea(vGameName, "設定各種", "優勝者範囲名")
            ElseIf VLookupArea(vGameName, "設定各種", "優勝者シート名") <> "" Then
                Call DefineWinnerSheet(VLookupArea(vGameName, "設定各種", "優勝者シート名"), _
                                    VLookupArea(vGameName, "設定各種", "優勝者範囲名"))
            End If
        
            ' 賞状の設定
            If VLookupArea(vGameName, "設定各種", "賞状関数名") <> "" Then
                Application.Run VLookupArea(vGameName, "設定各種", "賞状関数名"), _
                                    VLookupArea(vGameName, "設定各種", "賞状シート名")
            End If
        
        End If
    Next vGameName
    
    ' シートのロック
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible

End Sub

'
' シートのテーブル範囲名を定義する
'
' sSheetName    IN      シート名
' sAreaName     IN      範囲名
'
Private Sub DefineTableRange(sSheetName As String, sAreaName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Dim bProtect As Boolean
    bProtect = SheetProtect(False, oWorkSheet)
    
    Call DefineName(sAreaName, TableRangeAddress("$A$1")) ' 種目番号から各要素を引く

    ' シートの表示／保護
    Call SheetProtect(bProtect, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 学童マスターズ大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 学マ大会種目区分設定(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Call DefineName("学マ年齢区分", TableRangeAddress("$H$1"))
    Call DefineName("学マ学童区分", TableRangeAddress("$K$1"))
    Call DefineName("学マ学年表示", TableRangeAddress("$N$1"))

    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 市民大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 市民大会種目区分設定(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Call DefineName("市民選手年齢区分", RowRangeAddress("$H$1"))
    Call DefineName("市民リレー年齢区分", RowRangeAddress("$IJ$1"))
    Call DefineName("市民年齢区分", TableRangeAddress("$K$1"))

    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 室内記録会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 室内記録会種目区分設定(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Call DefineName("記録会年齢区分", TableRangeAddress("$J$1"))
    Call DefineName("検定年齢区分", TableRangeAddress("$O$1"))

    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 優勝者シートの設定
'
' sSheetName    IN      シート名
' sAreaName     IN      範囲名
'
'
Public Sub DefineWinnerSheet(sSheetName As String, sAreaName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' 範囲名の設定
    Call DefineTableRange(sSheetName, sAreaName)

    ' フィルタ設定
    Call SetAutoFilter(sAreaName, True)
    
    ' 印刷範囲の設定
    Sheets(sSheetName).PageSetup.PrintArea = GetRange(sAreaName).Address

    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 大会記録シートの設定
'
' sSheetName    IN      シート名
' sAreaName     IN      範囲名
'
'
Public Sub DefineRecordSheet(sSheetName As String, sAreaName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' 範囲名の設定
    Call DefineTableRange(sSheetName, sAreaName)

    ' フィルタ設定
    Call SetAutoFilter(sAreaName, True)
    
    ' 印刷範囲の設定
    Dim oRange As Range
    Set oRange = GetRange(sAreaName)
    Sheets(sSheetName).PageSetup.PrintArea = oRange.Offset(0, 1).Resize(oRange.Rows.Count, oRange.Columns.Count - 1).Address

    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 大会記録／優勝者シートの設定（選手権用）
'
' sSheetName    IN      シート名
' sAreaName     IN      範囲名
'
'
Public Sub 選手権大会記録設定(sSheetName As String, sAreaName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' 範囲名の設定
    Call DefineName(sAreaName, TableRangeAddress("$A$2"))

    ' フィルタ設定
    Call SetAutoFilter(sAreaName, False)
    
    ' 行の高さの設定
    Call SetWinnerRowHeight(sAreaName, 16)
    
    ' 印刷範囲の設定
    Dim oRange As Range
    Set oRange = GetRange(sAreaName)
    Sheets(sSheetName).PageSetup.PrintArea = oRange.Offset(-1, 2).Resize(oRange.Rows.Count + 1, oRange.Columns.Count - 2).Address

    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 大会記録／優勝者シートの高さ設定（選手権用）
'
' sSheetName    IN      シート名
' nHeight       IN      高さ
'
Private Sub SetWinnerRowHeight(sAreaName As String, nHeight As Integer)
    Dim vKey As Variant
    For Each vKey In GetAreaKeyData(sAreaName)
        ' リレー種目は4倍
        If GetOffset(vKey, GetColIdx(sAreaName, "種目")).MergeArea.Item(1).Value Like "*リレー" Then
            vKey.RowHeight = nHeight * 4
        Else
            vKey.RowHeight = nHeight
        End If
    Next vKey
End Sub

'
' 学マ大会の賞状定義
'
' sSheetName    IN      シート名
'
Sub 学マ賞状名前定義(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 名前をすべて削除
    Call DeleteName("賞状*")

    Call DefineName("賞状種目区分", "$E$9")
    Call DefineName("賞状距離", "$I$9")
    Call DefineName("賞状種目", "$N$9")
    Call DefineName("賞状順位", "$B$13")
    Call DefineName("賞状タイム", "$N$14")
    Call DefineName("賞状大会新", "$U$14")
    Call DefineName("賞状氏名", "$D$20")
    Call DefineName("賞状所属", "$D$24")
 
    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 市民大会の賞状定義
'
' sSheetName    IN      シート名
'
Sub 市民賞状名前定義(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 名前をすべて削除
    Call DeleteName("賞状*")

    Call DefineName("賞状種目区分", "$AC$4")
    Call DefineName("賞状種目距離区分", "$AC$16")
    Call DefineName("賞状順位", "$AA$10")
    Call DefineName("賞状タイム", "$Y$10")
    Call DefineName("賞状大会新", "$Y$27")
    Call DefineName("賞状氏名", "$U$9")
    Call DefineName("賞状所属", "$W$6")
    
    Call DefineName("賞状大会回数１", "$C$7")
    Call DefineName("賞状大会回数２", "$R$15")
    Call DefineName("賞状年", "$F$5")
    Call DefineName("賞状月", "$F$15")
    Call DefineName("賞状日", "$F$20")
 
    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 選手権大会の賞状定義
'
' sSheetName    IN      シート名
'
Sub 選手権賞状名前定義(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 名前をすべて削除
    Call DeleteName("賞状*")

    Call DefineName("賞状性別", "$H$8")
    Call DefineName("賞状距離", "$L$8")
    Call DefineName("賞状種目", "$S$8")
    Call DefineName("賞状順位", "$C$6")
    Call DefineName("賞状タイム", "$H$10")
    Call DefineName("賞状大会新", "$W$10")
    Call DefineName("賞状氏名", "$H$12")
    Call DefineName("賞状所属", "$H$14")
 
    Call DefineName("賞状大会回数", "$C$4")
    Call DefineName("賞状年", "$G$25")
    Call DefineName("賞状月", "$N$25")
    Call DefineName("賞状日", "$R$25")
 
    ' シートの表示／保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' トップページの定義
'
' sSheetName    IN      シート名
'
Private Sub トップページ定義(sSheetName As String)
    ' 表示／アクティブ／解除
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 名前をすべて削除
    Call DeleteName("大会*")
    
    Call 大会名定義
    Call 大会年定義
    Call プリンタ定義
    Call 組合せ方式定義
    Call 組最少人数定義
    Call レース定員定義
    Call 最小レーン番号定義
    Call 賞状定数定義

    ' シートのロック
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' 配列の設定定義
'
' sName         IN      名前
' sAddress      IN      アドレス
' sAry          IN      リスト
'
Private Sub DefineListValidation(sName As String, sAddress As String, sAry() As String)

    Call DefineName(sName, sAddress)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=Join(sAry, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    ' デフォルト値の設定
    If Not isAryExist(sAry, Range(sName).Value) Then
        Range(sName).Value = sAry(0)
    End If
    
End Sub

'
' 配列の設定定義
'
' sName         IN      名前
' sAddress      IN      アドレス
' nMin          IN      最小値
' nMax          IN      最大値
' nDefault      IN      デフォルト値
'
Private Sub DefineBetweenValidation(sName As String, sAddress As String, _
nMin As Integer, nMax As Integer, Optional nDefault As Variant = Empty)

    Call DefineName(sName, sAddress)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=CStr(nMin), Formula2:=CStr(nMax)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = sName & "は数字だけで入力してください。"
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = CStr(nMin) & "～" & CStr(nMax) & "までの数字を入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
    
    ' デフォルト値の設定
    If Range(sName).Value < nMin Or Range(sName).Value > nMax Then
        If IsEmpty(nDefault) Then
            Range(sName).Value = ""
        Else
            Range(sName).Value = nDefault
        End If
    End If
    
End Sub

'
' 大会名を定義
'
Private Sub 大会名定義()
    
    Dim sAry(3) As String
    sAry(0) = 学マ大会
    sAry(1) = 市民大会
    sAry(2) = 選手権大会
    sAry(3) = 室内記録会

    Call DefineListValidation("大会名", "$B$1", sAry)
    
End Sub

'
' 大会年を定義
'
Private Sub 大会年定義()
    
    Call DefineBetweenValidation("大会年", "$E$6", 2000, 2050, Year(Now))

End Sub

'
' 賞状定数を定義
'
Private Sub 賞状定数定義()
    Call 賞状回数定義
    Call 賞状年定義
    Call 賞状月定義
    Call 賞状日定義
End Sub

'
' 賞状大会回数を定義
'
Private Sub 賞状回数定義()
    
    Call DefineBetweenValidation("大会回数", "$E$9", 1, 150)

End Sub

'
' 賞状年を定義
'
Private Sub 賞状年定義()
    
    Call DefineName("大会元号年", "$E$10")

End Sub

'
' 賞状月を定義
'
Private Sub 賞状月定義()
    
    Call DefineBetweenValidation("大会月", "$E$11", 1, 12)

End Sub

'
' 賞状日を定義
'
Private Sub 賞状日定義()
    
    Call DefineBetweenValidation("大会日", "$E$12", 1, 31)

End Sub


'
' プリンタを定義
'
Private Sub プリンタ定義()
    Call プリンタ名定義
    Call 賞状印刷プレビュー定義
End Sub


'
' プリンタを定義
'
Private Sub プリンタ名定義()

    Dim sAry() As String
    sAry = GetPrinters
    Call DefineListValidation("大会プリンタ名", "$E$7", sAry)
    
End Sub

'
' プリンタの一覧を取得する
'
Private Function GetPrinters() As String()
    ' プリンタ一覧の取得
    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    
    ReDim sDeviceAry(oShell.Namespace(4).Items.Count) As String
    Dim i As Integer
    i = 0
    Dim vPrinters As Variant
    For Each vPrinters In oShell.Namespace(4).Items
        sDeviceAry(i) = vPrinters.Name
        i = i + 1
    Next
    GetPrinters = sDeviceAry
End Function

'
'賞状印刷プレビュー定義
'
Private Sub 賞状印刷プレビュー定義()
    
    Dim sAry(2) As String
    sAry(0) = "しない"
    sAry(1) = "する"
    Call DefineListValidation("大会印刷プレビュー", "$E$8", sAry)
    
End Sub


'
'組合せ方式定義
'
' sValue        IN      ダミー
'
Private Sub 組合せ方式定義(Optional sValue As String = "")
    
    Dim sAry(2) As String
    sAry(0) = "単純方式"
    sAry(1) = "混合分け方式"
    Call DefineListValidation("大会組合せ方式", "$E$5", sAry)
    
End Sub

'
' 組最小人数定義
'
' sValue        IN      ダミー
'
Private Sub 組最少人数定義(Optional sValue As String = "")

    Dim sAry(2) As String
    sAry(0) = "3"
    sAry(1) = "4"
    sAry(2) = "2"
    Call DefineListValidation("大会組最少人数", "$E$2", sAry)

End Sub

'
' レース定員定義
'
' sValue        IN      ダミー
'
Private Sub レース定員定義(Optional sValue As String = "")

    Dim sAry(2) As String
    sAry(0) = "7"
    sAry(1) = "6"
    sAry(2) = "5"
    Call DefineListValidation("大会組レース定員", "$E$3", sAry)

End Sub

'
' 最小レーン番号定義
'
' sValue        IN      ダミー
'
Private Sub 最小レーン番号定義(Optional sValue As String = "")

    Dim sAry(2) As String
    sAry(0) = "3"
    sAry(1) = "2"
    sAry(2) = "1"
    Call DefineListValidation("大会組最小レーン番号", "$E$4", sAry)

End Sub




'
' シート非表示の設定
'
' sValue        IN      ダミー
'
Public Sub シート非表示(Optional sValue As String = "")

    For Each vGameName In GetAreaKeyData("設定各種")
        If VLookupArea(vGameName, "設定各種", "対象") = 1 Then
            If GetRange("大会名").Value = CStr(vGameName) Then
                Call SetSheetVisible(CStr(vGameName), True)
            Else
                Call SetSheetVisible(CStr(vGameName), False)
            End If
        End If
    Next vGameName

End Sub

'
' 各種シート表示／非表示
'
' vGameName IN  大会名
' bFlag     IN  True:表示／False:非表示
'
Private Sub SetSheetVisible(vGameName As String, bFlag As Boolean)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    Call SheetVisible(VLookupArea(vGameName, "設定各種", "種目区分シート名"), bFlag)
    Call SheetVisible(VLookupArea(vGameName, "設定各種", "大会記録シート名"), bFlag)
    Call SheetVisible(VLookupArea(vGameName, "設定各種", "優勝者シート名"), bFlag)
    Call SheetVisible(VLookupArea(vGameName, "設定各種", "賞状シート名"), bFlag)
    ' 賞状の設定
    If bFlag And _
        VLookupArea(vGameName, "設定各種", "賞状関数名") <> "" Then
        Application.Run VLookupArea(vGameName, "設定各種", "賞状関数名"), _
                            VLookupArea(vGameName, "設定各種", "賞状シート名")
    End If
    oWorkSheet.Activate
End Sub

'
' 各種設定シートを表示
'
Public Sub 設定各種表示()
    Call SheetVisible(設定各種シート, True)
End Sub


