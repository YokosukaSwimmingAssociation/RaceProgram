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
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
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
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Private Sub Prog名前定義(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
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
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 記録画面の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 記録画面名前定義(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

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
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
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

    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("設定*")

    ' シートを定義
    Call DefineName("設定各種", TableRangeAddress("$A$1"))

    ' 各種設定を行う
    For Each vGameName In GetAreaKeyData("設定各種")
        If VLookupArea(vGameName, "設定各種", "対象") = 1 Then
        
            Debug.Print vGameName
        
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
            Call DefineTableRange(VLookupArea(vGameName, "設定各種", "大会記録シート名"), _
                                    VLookupArea(vGameName, "設定各種", "大会記録範囲名"))
        
            ' 優勝者の設定
            Call DefineColumnRange(VLookupArea(vGameName, "設定各種", "優勝者シート名"), _
                                    VLookupArea(vGameName, "設定各種", "優勝者範囲名"))
        
            ' 賞状の設定
            If VLookupArea(vGameName, "設定各種", "賞状関数名") <> "" Then
                Application.Run VLookupArea(vGameName, "設定各種", "賞状関数名"), _
                                    VLookupArea(vGameName, "設定各種", "賞状シート名")
            End If
        
        End If
    Next vGameName
    
    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible

End Sub

'
' シートのテーブル範囲名を定義する
'
' sSheetName    IN      シート名
' sAreaName     IN      範囲名
'
Private Sub DefineTableRange(sSheetName As String, sAreaName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    Call DefineName(sAreaName, TableRangeAddress("$A$1")) ' 種目番号から各要素を引く
   
    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' シートのヘッダ行の範囲名を定義する
'
' sSheetName    IN      シート名
' sAreaName     IN      範囲名
'
Private Sub DefineColumnRange(sSheetName As String, sAreaName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    Call DefineName(sAreaName, ColumnRangeAddress("$A$1"))
    
    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub


'
' 学童マスターズ大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 学マ大会種目区分設定(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    Call DefineName("学マ年齢区分", TableRangeAddress("$H$1"))
    Call DefineName("学マ学童区分", TableRangeAddress("$K$1"))
    Call DefineName("学マ学年表示", TableRangeAddress("$N$1"))

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 市民大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Private Sub 市民大会種目区分設定(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    Call DefineName("市民選手年齢区分", RowRangeAddress("$H$1"))
    Call DefineName("市民リレー年齢区分", RowRangeAddress("$IJ$1"))
    Call DefineName("市民年齢区分", TableRangeAddress("$K$1"))

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 学マ大会の賞状定義
'
' sSheetName    IN      シート名
'
Sub 学マ賞状名前定義(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("賞状*")

    Call DefineName("賞状種目区分", "$C$9")
    Call DefineName("賞状距離", "$G$9")
    Call DefineName("賞状種目", "$L$9")
    Call DefineName("賞状順位", "$A$13")
    Call DefineName("賞状タイム", "$L$14")
    Call DefineName("賞状大会新", "$S$14")
    Call DefineName("賞状氏名", "$C$20")
    Call DefineName("賞状所属", "$C$24")
 
    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 市民大会の賞状定義
'
' sSheetName    IN      シート名
'
Sub 市民賞状名前定義(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("賞状*")

    Call DefineName("賞状性別", "$AC$4")
    Call DefineName("賞状種目距離区分", "$AC$16")
    Call DefineName("賞状順位", "$AA$7")
    Call DefineName("賞状タイム", "$Y$10")
    Call DefineName("賞状大会新", "$Y$27")
    Call DefineName("賞状氏名", "$U$9")
    Call DefineName("賞状所属", "$W$6")
 
    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 選手権大会の賞状定義
'
' sSheetName    IN      シート名
'
Sub 選手権賞状名前定義(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

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
 
    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' トップページの定義
'
' sSheetName    IN      シート名
'
Private Sub トップページ定義(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    Call 大会名定義
    Call 大会年定義
    Call プリンタ定義
    Call 組合せ方式定義
    Call 組最少人数定義

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = True
End Sub

'
' 大会名を定義
'
' sValue        IN      ダミー
'
Private Sub 大会名定義()
    
    Call DefineName("大会名", "$B$1")
    With Range("大会名").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="学童マスターズ大会,横須賀市民体育大会,横須賀選手権水泳大会"
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
' 大会年を定義
'
' sValue        IN      ダミー
'
Private Sub 大会年定義()
    
    Call DefineName("大会年", "$E$4")
    With Range("大会年").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="2000", Formula2:="2050"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "開催年は数字だけで入力してください。"
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "2000～2050までの数字を入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
    Range("大会年").Value = Year(Now)

End Sub

'
' 大会年を定義
'
' sValue        IN      ダミー
'
Private Sub プリンタ定義()
    Call DefineName("プリンタ名", "$E$5")
End Sub


'
'組合せ方式定義
'
' sValue        IN      ダミー
'
Private Sub 組合せ方式定義(Optional sValue As String = "")
    
    Call DefineName("組合せ方式", "$E$3")
    With Range("組合せ方式").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="単純方式,混合分け方式"
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
    Range("組合せ方式").Value = "単純方式"
    
End Sub

'
' 組最小人数定
'
' sValue        IN      ダミー
'
Private Sub 組最少人数定義(Optional sValue As String = "")

    Call DefineName("組最少人数", "$E$2")
    With Range("組最少人数").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="3,4"
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
    Range("組最少人数").Value = 4

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
Sub 各種設定表示()
    Call SheetVisible(設定各種シート, True)
End Sub


'
' モジュール読込み
'
Public Sub モジュール読込み()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ImportAll(ActiveWorkbook, sPathName)
End Sub

'
' モジュールExport
'
Public Sub モジュール出力()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ExportAll(ActiveWorkbook, sPathName)
End Sub
