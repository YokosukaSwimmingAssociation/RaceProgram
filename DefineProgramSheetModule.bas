Attribute VB_Name = "DefineProgramSheetModule"
'
' 名前を定義する
'
Sub ワークブック名前定義()
    Call EventChange(False)
  
    Call Header名前定義(S_PROGRAM_FORMAT_SHEET_NAME)
    Call Prog名前定義(S_PROGRAM_FORMAT_SHEET_NAME)
    Call 記録画面名前定義("記録画面")
    Call 学童マスターズ大会種目区分名前定義
    Call 学童マスターズ大会記録名前定義
    Call 学童マスターズ大会優勝者名前定義
    Call 市民大会種目区分名前定義
    Call 市民大会記録名前定義
    Call 市民大会優勝者名前定義
    Call 選手権大会種目区分名前定義
    Call 選手権大会記録名前定義
    Call 選手権大会優勝者名前定義
    Call 賞状名前定義
    Call マクロページ定義
    Call シート非表示
    
    Call EventChange(True)
    Sheets("プログラム作成マクロ").Select
    Range("A1").Select
End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Sub Header名前定義(sSheetName As String)
    Sheets(sSheetName).Visible = True
    Sheets(sSheetName).Select
    Call SheetProtect(False)
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
    Next

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    ActiveSheet.Visible = True

End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Sub Prog名前定義(sSheetName As String)
    Sheets(sSheetName).Visible = True
    Sheets(sSheetName).Select
    Call SheetProtect(False)
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
     
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    ActiveSheet.Visible = True
End Sub

'
' 記録画面の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 記録画面名前定義(sSheetName As String)
    Sheets(sSheetName).Visible = True
    Sheets(sSheetName).Select
    Call SheetProtect(False)

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
    Call SheetProtect(True)
    ActiveSheet.Visible = True
End Sub

'
' 記録画面の違反
'
' sValue        IN      ダミー
'
Sub 記録画面違反定義(Optional sValue As String = "")
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
' 学童マスターズ大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 学童マスターズ大会種目区分名前定義(Optional sValue As String = "")

    Sheets("学童マスターズ種目区分").Visible = True
    Sheets("学童マスターズ種目区分").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("学マ*")
    
    Call DefineName("学マ種目区分", TableRangeAddress("$A$1")) ' 種目番号から各要素を引く
    
    Call DefineName("学マ年齢区分", TableRangeAddress("$H$1"))
    Call DefineName("学マ学童区分", TableRangeAddress("$K$1"))
    Call DefineName("学マ学年表示", TableRangeAddress("$N$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 学童マスターズ大会記録の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 学童マスターズ大会記録名前定義(Optional sValue As String = "")
    Sheets("学童マスターズ大会記録").Visible = True
    Sheets("学童マスターズ大会記録").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("学マ大会記録")
    
    Call DefineName("学マ大会記録", TableRangeAddress("$A$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 学童マスターズ優勝者の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 学童マスターズ大会優勝者名前定義(Optional sValue As String = "")
    Sheets("学童マスターズ優勝者").Visible = True
    Sheets("学童マスターズ優勝者").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("学マ大会優勝者")
    
    Call DefineName("学マ大会優勝者", ColumnRangeAddress("$A$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 市民大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 市民大会種目区分名前定義(Optional sValue As String = "")
    Sheets("市民大会種目区分").Visible = True
    Sheets("市民大会種目区分").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("市民*")
    
    Call DefineName("市民種目区分", TableRangeAddress("$A$1")) ' 種目番号から各要素を引く
    
    Call DefineName("市民選手年齢区分", RowRangeAddress("$H$1"))
    Call DefineName("市民リレー年齢区分", RowRangeAddress("$IJ$1"))
    Call DefineName("市民年齢区分", TableRangeAddress("$K$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 市民大会記録の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 市民大会記録名前定義(Optional sValue As String = "")
    Sheets("市民大会記録").Visible = True
    Sheets("市民大会記録").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("市民大会記録")
    
    Call DefineName("市民大会記録", TableRangeAddress("$A$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 市民大会優勝者の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 市民大会優勝者名前定義(Optional sValue As String = "")
    Sheets("市民大会優勝者").Visible = True
    Sheets("市民大会優勝者").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("市民大会優勝者")
    
    Call DefineName("市民大会優勝者", ColumnRangeAddress("$A$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 選手権大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 選手権大会種目区分名前定義(Optional sValue As String = "")
    Sheets("選手権大会種目区分").Visible = True
    Sheets("選手権大会種目区分").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("選手権*")
    
    Call DefineName("選手権種目区分", TableRangeAddress("$A$1")) ' 種目番号から各要素を引く
   
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 選手権大会記録の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 選手権大会記録名前定義(Optional sValue As String = "")
    Sheets("選手権大会記録").Visible = True
    Sheets("選手権大会記録").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("選手権大会記録")
    
    Call DefineName("選手権大会記録", TableRangeAddress("$A$2"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' 選手権大会優勝者の名前を定義する
'
' sValue        IN      ダミー
'
Sub 選手権大会優勝者名前定義(Optional sValue As String = "")
    Sheets("選手権大会優勝者").Visible = True
    Sheets("選手権大会優勝者").Select
    Call SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("選手権大会優勝者")
    
    Call DefineName("選手権大会優勝者", ColumnRangeAddress("$A$1"))
    
    ' シートのロック
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub


'
' マクロページの定義
'
' sValue        IN      ダミー
'
Sub マクロページ定義(Optional sValue As String = "")

    Sheets("プログラム作成マクロ").Select
    Call SheetProtect(False)

    Call 大会名定義
    Call 大会年定義
    Call 組合せ方式定義
    Call 組最少人数定義

    ' シートのロック
    Call SheetProtect(True)
End Sub

'
' 大会名を定義
'
' sValue        IN      ダミー
'
Sub 大会名定義(Optional sValue As String = "")
    
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
Sub 大会年定義(Optional sValue As String = "")
    
    Call DefineName("大会年", "$E$7")
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
'組合せ方式定義
'
' sValue        IN      ダミー
'
Sub 組合せ方式定義(Optional sValue As String = "")
    
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
Sub 組最少人数定義(Optional sValue As String = "")

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

Sub シート非表示(Optional sValue As String = "")

    If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
        Call 学マ大会シート非表示(False)
        Call 市民大会シート非表示(False)
        Call 選手権大会シート非表示(True)
    ElseIf GetRange("大会名").Value = "横須賀市民体育大会" Then
        Call 学マ大会シート非表示(False)
        Call 市民大会シート非表示(True)
        Call 選手権大会シート非表示(False)
    Else
        Call 学マ大会シート非表示(True)
        Call 市民大会シート非表示(False)
        Call 選手権大会シート非表示(False)
    End If

End Sub

'
' 学童マスターズシート非表示
'
' bFlag     IN  True:表示／False:非表示
'
Sub 学マ大会シート非表示(bFlag As Boolean)
    Sheets("学童マスターズ種目区分").Visible = bFlag
    Sheets("学童マスターズ大会記録").Visible = bFlag
    Sheets("学童マスターズ優勝者").Visible = bFlag
    Sheets("学童マスターズ賞状").Visible = bFlag
End Sub

'
' 学童マスターズシート非表示
'
' bFlag     IN  True:表示／False:非表示
'
Sub 市民大会シート非表示(bFlag As Boolean)
    Sheets("市民大会種目区分").Visible = bFlag
    Sheets("市民大会記録").Visible = bFlag
    Sheets("市民大会優勝者").Visible = bFlag
    'Sheets("市民大会賞状").Visible = bFlag
End Sub

'
' 学童マスターズシート非表示
'
' bFlag     IN  True:表示／False:非表示
'
Sub 選手権大会シート非表示(bFlag As Boolean)
    Sheets("選手権大会種目区分").Visible = bFlag
    Sheets("選手権大会記録").Visible = bFlag
    Sheets("選手権大会優勝者").Visible = bFlag
    'Sheets("選手権大会賞状").Visible = bFlag
End Sub

'
' モジュール読込み
'
Sub モジュール読込み()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ImportAll(ActiveWorkbook, sPathName)
End Sub

'
' モジュールExport
'
Sub モジュール出力()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ExportAll(ActiveWorkbook, sPathName)
End Sub
