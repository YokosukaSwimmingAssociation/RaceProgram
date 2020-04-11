Attribute VB_Name = "DefineProgramSheetModule"
'
' 名前を定義する
'
Sub ワークブック名前定義()
    Call EventChange(False)
  
    Call Header名前定義(sProgramFormatSheetName)
    Call Prog名前定義(sProgramFormatSheetName)
    Call 記録画面名前定義("記録画面")
    Call 学童マスターズ大会種目区分名前定義("学童マスターズ種目区分")
    Call 学童マスターズ大会記録名前定義("学童マスターズ大会記録")
    Call 市民大会種目区分名前定義("市民大会種目区分")
    Call 市民大会記録名前定義("市民大会記録")
    Call 選手権大会種目区分名前定義("選手権大会種目区分")
    Call 選手権大会記録名前定義("選手権大会記録")
    Call 大会名定義("プログラム作成マクロ")
    
    Call EventChange(True)
End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Sub Header名前定義(sSheetName As String)
    Sheets(sSheetName).Select
    ActiveSheet.Unprotect
    Range("$A$1").Select

    ' 名前をすべて削除
    Call DeleteName("Header*")

    Dim oCell As Range
    Dim sName As String
    For nColumn = 1 To ActiveCell.SpecialCells(xlCellTypeLastCell).Column
        Set oCell = Cells(1, nColumn)
        sName = Trim(Replace(oCell.Value, "　", ""))
        If sName <> "" Then
            Call SetName("Header" & sName, oCell.Address(ReferenceStyle:=xlA1))
            If sName = "所属" Then
                Call SetName("Header" & sName & "前", oCell.Offset(0, -1).Address(ReferenceStyle:=xlA1))
                Call SetName("Header" & sName & "後", oCell.Offset(0, 1).Address(ReferenceStyle:=xlA1))
            End If
        End If
    Next

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

'
' プログラムフォーマットのヘッダー名前定義
'
' sSheetName    IN      シート名
'
Sub Prog名前定義(sSheetName As String)
    Sheets(sSheetName).Select
    ActiveSheet.Unprotect
    Range("$A$1").Select

    ' 名前をすべて削除
    Call DeleteName("Prog*")

    ' プログラムヘッダ
    Call SetName("ProgプロNo", "$C$3")
    Call SetName("Prog種目区分", "$D$3")
    Call SetName("Prog種目名", "$G$3")

    ' 組ヘッダ
    Call SetName("Prog組", "$C$4")
   
    ' レーンデータ
    Call SetName("Prog組番", "$C$5")
    Call SetName("Progレーン", "$D$5")
    Call SetName("Prog氏名", "$E$5")
    Call SetName("Prog所属前", "$F$5")
    Call SetName("Prog所属", "$G$5")
    Call SetName("Prog所属後", "$H$5")
    Call SetName("Prog区分", "$I$5")
    Call SetName("Prog時間", "$J$5")
    Call SetName("Prog順位", "$K$5")
    Call SetName("Prog備考", "$L$5")
    Call SetName("Prog大会記録", "$M$5")
    Call SetName("Prog申込み記録", "$N$5")
    Call SetName("ProgレースNo", "$O$5")
    Call SetName("Progソート区分", "$P$5")

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

'
' 記録画面の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 記録画面名前定義(sSheetName As String)

    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("記録画面*")

    Call SetName("記録画面種目番号", "$B$1")
    Call SetName("記録画面種目名", "$C$1")
    Call SetName("記録画面組", "$B$2")
    Call SetName("記録画面レースNo", "$B$3")
    Call SetName("記録画面レーン", "$B$5:$B$11")
    Call SetName("記録画面タイム", "$C$5:$C$11")
    Call SetName("記録画面選手名", "$D$5:$D$11")
    Call SetName("記録画面チーム名", "$E$5:$E$11")
    Call SetName("記録画面大会新", "$F$5:$F$11")

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 学童マスターズ大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 学童マスターズ大会種目区分名前定義(sSheetName As String)

    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("学マ*")
    
    Call SetName("学マ種目番号", TableRangeAddress("$A$1")) ' 種目名から種目番号を引く
    Call SetName("学マ種目区分", "$B$2:$G$73") ' 種目番号から各要素を引く
    
    Call SetName("学マ年齢区分", TableRangeAddress("$L$2"))
    Call SetName("学マ学童区分", TableRangeAddress("$O$2"))
    Call SetName("学マ学年表示", TableRangeAddress("$R$2"))

    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 学童マスターズ大会記録の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 学童マスターズ大会記録名前定義(sSheetName As String)
    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("学マ大会記録")
    
    Call SetName("学マ大会記録", TableRangeAddress("$A$1"))
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 学童マスターズ優勝者の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 学童マスターズ大会優勝者名前定義(sSheetName As String)
    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("学マ大会優勝者")
    
    Call SetName("学マ大会優勝者", TableRangeAddress("$B$1"))
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 市民大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 市民大会種目区分名前定義(sSheetName As String)

    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("市民*")
    
    Call SetName("市民種目番号", TableRangeAddress("$A$1")) ' 種目名から種目番号を引く
    Call SetName("市民種目区分", "$B$2:$G$73") ' 種目番号から各要素を引く
    
    Call SetName("市民選手年齢区分", ColumnRangeAddress("$I$2"))
    Call SetName("市民リレー年齢区分", ColumnRangeAddress("$J$2"))
    Call SetName("市民年齢区分", TableRangeAddress("$P$2"))
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 市民大会記録の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 市民大会記録名前定義(sSheetName As String)
    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("市民大会記録")
    
    Call SetName("市民大会記録", TableRangeAddress("$A$1"))
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 選手権大会種目区分の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 選手権大会種目区分名前定義(sSheetName As String)

    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("選手権*")
    
    Call SetName("選手権種目番号", TableRangeAddress("$A$1")) ' 種目名から種目番号を引く
    Call SetName("選手権種目区分", "$B$2:$H$47") ' 種目番号から各要素を引く
    
    Call SetName("選手権年齢区分", ColumnRangeAddress("$J$2"))
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

'
' 選手権大会記録の名前を定義する
'
' sSheetName    IN      シート名
'
Sub 選手権大会記録名前定義(sSheetName As String)
    Sheets(sSheetName).Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("選手権大会記録")
    
    Call SetName("選手権大会記録", TableRangeAddress("$A$2"))
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Sub 大会名定義(sSheetName As String)

    Sheets(sSheetName).Select
    ActiveSheet.Unprotect
    With Range("$B$1").Validation
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
    Call SetName("大会名", "$B$1")
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
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
