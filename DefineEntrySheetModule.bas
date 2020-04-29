Attribute VB_Name = "DefineEntrySheetModule"
'
' エントリーシートに名前を定義する
'
Public Sub エントリーシート定義()
    Call EventChange(False)
    Call 名前定義
    Call 入力制限定義
    Call 条件付き書式定義
    Call 印刷範囲の設定
    Call EventChange(True)
    ActiveWorkbook.Save
End Sub

'
' シートに名前を定義する
'
Private Sub 名前定義()
    Call 記入票名前定義
    Call 種目番号区分名前定義
End Sub

'
' 記入票シートに名前を定義する
'
Private Sub 記入票名前定義()

    Dim vVisible As Variant
    vVisible = SheetActivate("記入票")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 名前をすべて削除
    Call DeleteName("*")

    ' シートの重要項目定義
    Call DefineName("大会名", "$E$4")
    Call DefineName("チーム名", "$E$5")
    If Range("大会名").Value = 選手権大会 Then
        Call DefineName("申込み期間", "$M$7")
    ElseIf Range("大会名").Value = 市民大会 Then
        Call DefineName("申込み期間", "$L$7")
    Else
        Call DefineName("申込み期間", "$K$7")
    End If

    ' 選手番号とリレー範囲の定義
    ' この範囲を以降の定義で利用する
    If Range("大会名").Value = 市民大会 Then
        Call DefineName("選手番号", "$B$12:$B$71,$B$98:$B$175,$B$194:$B$270,$B$290:$B$366,$B$386:$B$462,$B$482:$B$558")
        Call DefineName("リレー範囲", "$B$74:$B$77,$B$178:$B$181,$B$274:$B$277,$B$370:$B$373,$B$466:$B$469,$B$562:$B$565")
    Else
        Call DefineName("選手番号", "$B$12:$B$31,$B$58:$B$83,$B$102:$B$127,$B$146:$B$171,$B$190:$B$215,$B$234:$B$259")
        Call DefineName("リレー範囲", "$B$34:$B$37,$B$86:$B$89,$B$130:$B$133,$B$174:$B$177,$B$218:$B$221,$B$262:$B$265")
    End If

    Call DefineName("選手性別列", "$C$10")
    Call DefineNameByColumns("選手性別列", "選手性別")

    Call DefineName("選手名列", "$D$10")
    Call DefineName("選手区分列", "$F$10")
    
    If Range("大会名").Value = 選手権大会 Then
        Call DefineName("種目一覧", "$G$10:$S$10")
        Call DefineName("種目距離", "$G$11:$S$11")
   
        Call DefineName("自由形50M列", "$G$11")
        Call DefineName("自由形100M列", "$H$11")
        Call DefineName("自由形200M列", "$I$11")
        Call DefineName("平泳ぎ50M列", "$J$11")
        Call DefineName("平泳ぎ100M列", "$K$11")
        Call DefineName("平泳ぎ200M列", "$L$11")
        Call DefineName("バタフライ50M列", "$M$11")
        Call DefineName("バタフライ100M列", "$N$11")
        Call DefineName("バタフライ200M列", "$O$11")
        Call DefineName("背泳ぎ50M列", "$P$11")
        Call DefineName("背泳ぎ100M列", "$Q$11")
        Call DefineName("背泳ぎ200M列", "$R$11")
        Call DefineName("個人メドレー200M列", "$S$11")
        Call DefineName("フリーリレー4×50M列", "$T$11")
        Call DefineName("メドレーリレー4×50M列", "$U$11")
        
        Call DefineName("選手種目列", "$G$11:$S$11")
        Call DefineName("選手リレー種目列", "$T$11:$U$11")
    
        Call DefineName("選手分列", "$V$12")
        Call DefineName("選手秒列", "$X$12")
        Call DefineName("選手ミリ秒列", "$Z$12")
    
        Call DefineName("リレー分列", "$L$34")
        Call DefineName("リレー秒列", "$N$34")
        Call DefineName("リレーミリ秒列", "$P$34")
    
    ElseIf Range("大会名").Value = 市民大会 Then
        Call DefineName("種目一覧", "$G$10:$P$10")
        Call DefineName("種目距離", "$G$11:$P$11")
        
        Call DefineName("自由形50M列", "$G$11")
        Call DefineName("自由形100M列", "$H$11")
        Call DefineName("自由形200M列", "$I$11")
        Call DefineName("平泳ぎ50M列", "$J$11")
        Call DefineName("平泳ぎ100M列", "$K$11")
        Call DefineName("バタフライ50M列", "$L$11")
        Call DefineName("バタフライ100M列", "$M$11")
        Call DefineName("背泳ぎ50M列", "$N$11")
        Call DefineName("背泳ぎ100M列", "$O$11")
        Call DefineName("個人メドレー200M列", "$P$11")
        Call DefineName("フリーリレー4×50M列", "$Q$11")
        Call DefineName("メドレーリレー4×50M列", "$R$11")
        
        Call DefineName("選手種目列", "$G$11:$P$11")
        Call DefineName("選手リレー種目列", "$Q$11:$R$11")
        
        Call DefineName("選手分列", "$T$12")
        Call DefineName("選手秒列", "$V$12")
        Call DefineName("選手ミリ秒列", "$X$12")
        
        Call DefineName("リレー区分列", "$B$33")
        
        Call DefineName("リレー分列", "$L$34")
        Call DefineName("リレー秒列", "$N$34")
        Call DefineName("リレーミリ秒列", "$P$34")
    
    Else
        ' 学童マスターズ大会
        Call DefineName("種目一覧", "$G$10:$O$10")
        Call DefineName("種目距離", "$G$11:$O$11")
    
        Call DefineName("自由形50M列", "$G$11")
        Call DefineName("自由形100M列", "$H$11")
        Call DefineName("平泳ぎ50M列", "$I$11")
        Call DefineName("平泳ぎ100M列", "$J$11")
        Call DefineName("バタフライ50M列", "$K$11")
        Call DefineName("バタフライ100M列", "$L$11")
        Call DefineName("背泳ぎ50M列", "$M$11")
        Call DefineName("背泳ぎ100M列", "$N$11")
        Call DefineName("個人メドレー200M列", "$O$11")
        Call DefineName("フリーリレー4×50M列", "$P$11")
        Call DefineName("メドレーリレー4×50M列", "$Q$11")
        
        If Range("大会名").Value = マスターズ大会 Then
            Call DefineName("混合フリーリレー4×50M列", "$R$11")
            Call DefineName("混合メドレーリレー4×50M列", "$S$11")
            Call DefineName("選手リレー種目列", "$P$11:$S$11")
            
            Call DefineName("リレー区分列", "$B$33")
        Else
            Call DefineName("選手リレー種目列", "$P$11:$Q$11")
        End If
        
        Call DefineName("選手種目列", "$G$11:$O$11")
    
        Call DefineName("選手分列", "$T$12")
        Call DefineName("選手秒列", "$V$12")
        Call DefineName("選手ミリ秒列", "$X$12")
    
        Call DefineName("リレー分列", "$K$34")
        Call DefineName("リレー秒列", "$M$34")
        Call DefineName("リレーミリ秒列", "$O$34")
    
    End If

    Call DefineName("表示種目番号列", "$AB$11")
    Call DefineName("表示種目区分列", "$AC$11")
    Call DefineName("表示種目性別列", "$AD$11")
    Call DefineName("表示種目距離列", "$AE$11")
    Call DefineName("表示種目名列", "$AF$11")
    Call DefineName("表示区分列", "$AJ$11")
    Call DefineName("表示性別列", "$AK$11")
    Call DefineName("表示距離列", "$AL$11")

    Call DefineName("リレー種目列", "$E$33")
    Call DefineName("リレー種目名列", "$F$33")

    If Range("大会名").Value = 選手権大会 Then
        
        Call DefineNameByEvenOddColumns("選手名列", "選手フリガナ", "選手名")
        Call DefineNameByColumns("選手区分列", "選手区分")
        Call DefineNameByEvenOddColumns("選手種目列", "選手種目偶数", "選手種目奇数")
    
    ElseIf Range("大会名").Value = 市民大会 Then
        
        Call DefineNameByTripleColumns("選手名列", "選手フリガナ", "選手名", "選手学校名")
        Call DefineNameByTripleColumns("選手区分列", "選手区分", "選手年齢", "")
        Call DefineNameByTripleColumns("選手種目列", "選手種目偶数", "選手種目奇数", "")
    
    ElseIf Range("大会名").Value = マスターズ大会 Then
    
        Call DefineNameByEvenOddColumns("選手名列", "選手フリガナ", "選手名")
        Call DefineNameByColumns("選手区分列", "選手年齢")
        Call DefineNameByEvenOddColumns("選手種目列", "選手種目偶数", "選手種目奇数")
        
    Else
        ' 学童大会
        Call DefineNameByEvenOddColumns("選手名列", "選手フリガナ", "選手名")
        Call DefineNameByColumns("選手区分列", "選手学年")
        Call DefineNameByEvenOddColumns("選手種目列", "選手種目偶数", "選手種目奇数")
    
    End If

    Call DefineNameByColumns("選手リレー種目列", "選手リレー種目")
                    
    Call DefineNameByColumns("自由形50M列", "自由形50M")
    Call DefineNameByColumns("自由形100M列", "自由形100M")
    Call DefineNameByColumns("自由形200M列", "自由形200M")
    Call DefineNameByColumns("平泳ぎ50M列", "平泳ぎ50M")
    Call DefineNameByColumns("平泳ぎ100M列", "平泳ぎ100M")
    Call DefineNameByColumns("平泳ぎ200M列", "平泳ぎ200M")
    Call DefineNameByColumns("バタフライ50M列", "バタフライ50M")
    Call DefineNameByColumns("バタフライ100M列", "バタフライ100M")
    Call DefineNameByColumns("バタフライ200M列", "バタフライ200M")
    Call DefineNameByColumns("背泳ぎ50M列", "背泳ぎ50M")
    Call DefineNameByColumns("背泳ぎ100M列", "背泳ぎ100M")
    Call DefineNameByColumns("背泳ぎ200M列", "背泳ぎ200M")
    Call DefineNameByColumns("個人メドレー200M列", "個人メドレー200M")
    Call DefineNameByColumns("フリーリレー4×50M列", "フリーリレー4×50M")
    Call DefineNameByColumns("メドレーリレー4×50M列", "メドレーリレー4×50M")
    Call DefineNameByColumns("混合フリーリレー4×50M列", "混合フリーリレー4×50M")
    Call DefineNameByColumns("混合メドレーリレー4×50M列", "混合メドレーリレー4×50M")
    
    Call DefineNameByColumns("選手分列", "選手分")
    Call DefineNameByColumns("選手秒列", "選手秒")
    Call DefineNameByColumns("選手ミリ秒列", "選手ミリ秒")

    Call DefineNameByColumns("表示種目番号列", "表示種目番号")
    Call DefineNameByColumns("表示種目区分列", "表示種目区分")
    Call DefineNameByColumns("表示種目性別列", "表示種目性別")
    Call DefineNameByColumns("表示種目距離列", "表示種目距離")
    Call DefineNameByColumns("表示種目名列", "表示種目名")
    Call DefineNameByColumns("表示区分列", "表示区分")
    Call DefineNameByColumns("表示性別列", "表示性別")
    Call DefineNameByColumns("表示距離列", "表示距離")

    Call DefineNameByRelayColumns("リレー区分列", "リレー区分")
    Call DefineNameByRelayColumns("リレー種目列", "リレー種目")
    Call DefineNameByRelayColumns("リレー種目名列", "リレー種目名")
    Call DefineNameByRelayColumns("リレー分列", "リレー分")
    Call DefineNameByRelayColumns("リレー秒列", "リレー秒")
    Call DefineNameByRelayColumns("リレーミリ秒列", "リレーミリ秒")

    Sheets("記入票").Select
    Call SetForcusTop

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible

End Sub

'
'種目番号区分シートに名前を定義する
'
Private Sub 種目番号区分名前定義()
    Dim vVisible As Variant
    vVisible = SheetActivate("種目番号区分")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    If Range("大会名").Value = 選手権大会 Then
    
        Call DefineName("種目番号区分", TableRangeAddress("$A$1"))
        Call DefineName("選手年齢区分", RowRangeAddress("$H$2"))
        Call DefineName("申込み期間開始", "$J$2")
        Call DefineName("申込み期間終了", "$J$3")
        Call DefineName("リレー種目番号", RowRangeAddress("$L$2"))
    
    ElseIf Range("大会名").Value = 市民大会 Then

        Call DefineName("種目番号区分", TableRangeAddress("$A$1"))
        Call DefineName("選手年齢区分", RowRangeAddress("$G$2"))
        Call DefineName("リレー年齢区分", RowRangeAddress("$H$2"))
        Call DefineName("申込み期間開始", "$J$2")
        Call DefineName("申込み期間終了", "$J$3")
        Call DefineName("リレー種目番号", RowRangeAddress("$L$2"))

    ElseIf Range("大会名").Value = マスターズ大会 Then
    
        Call DefineName("種目番号区分", TableRangeAddress("$A$1"))
        Call DefineName("リレー年齢区分", RowRangeAddress("$H$2"))
        Call DefineName("申込み期間開始", "$J$2")
        Call DefineName("申込み期間終了", "$J$3")
        Call DefineName("リレー種目番号", RowRangeAddress("$L$2"))
    
    Else
        ' 学童大会
        Call DefineName("種目番号区分", TableRangeAddress("$A$1"))
        Call DefineName("選手区分ＴＢ", TableRangeAddress("$G$2"))
        Call DefineName("申込み期間開始", "$J$2")
        Call DefineName("申込み期間終了", "$J$3")
        Call DefineName("リレー種目番号", RowRangeAddress("$L$2"))
    
    End If

    Sheets("種目番号区分").Select
    Call SetForcusTop

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 列毎の名前を付ける
'
' 列の名前で指定された範囲(選手行)に名前を付ける
'
' 使用例
' Call DefineNameByColumns("選手性別列", "選手性別")
'
' sColName          IN      列の名前
' sName             IN      範囲につける名前
'
Private Sub DefineNameByColumns(sColName As String, sName As String)

    Call DefineNameByRangeColumns(sColName, sName, "選手番号")

End Sub

'
' リレー用列毎の名前を付ける
'
' 列の名前で指定された範囲(リレー行)に名前を付ける
'
' 使用例
' Call DefineNameByRelayColumns("リレー区分列", "リレー区分")
'
' sColName          IN      列の名前
' sName             IN      範囲につける名前
'
Private Sub DefineNameByRelayColumns(sColName As String, sName As String)

    Call DefineNameByRangeColumns(sColName, sName, "リレー範囲")

End Sub

'
' 列毎の名前を付ける
'
' 列の名前で指定された範囲(選手行)に名前を付ける
'
' 使用例
' Call DefineNameByColumns("選手性別列", "選手性別")
'
' sColName          IN      列の名前
' sName             IN      範囲につける名前
' sRangeName        IN      取得する範囲の名前
'
Private Sub DefineNameByRangeColumns(sColName As String, sName As String, sRangeName As String)

    ' 名前がない場合はスキップ
    If Not IsNameExists(sColName) Then
        Exit Sub
    End If

    ' 列番号を取得
    Dim nColumn As Integer
    Dim nCount As Integer
    nColumn = GetRange(sColName).Column
    nCount = GetRange(sColName).Columns.Count

    Dim oRange As Range
    Set oRange = Nothing
    For Each vCell In GetRange(sRangeName)
        If oRange Is Nothing Then
            Set oRange = Cells(vCell.Row, nColumn).Resize(1, nCount)
        Else
            Set oRange = Application.Union(oRange, Cells(vCell.Row, nColumn).Resize(1, nCount))
        End If
    Next vCell

    Call DefineName(sName, oRange.Address(ReferenceStyle:=xlA1))

End Sub

'
' 複数列毎の名前を付ける(偶数、奇数)
'
' 列の名前で指定された範囲(選手行)に偶数行、奇数行それぞれに名前を付ける
'
' 使用例
' Call DefineNameByEvenOddColumns("選手名列", "選手フリガナ", "選手名")
'
' sColName          IN      列の名前
' sEvenName         IN      偶数範囲につける名前
' sOddName          IN      奇数範囲につける名前
'
Private Sub DefineNameByEvenOddColumns(sColName As String, sEvenName As String, sOddName As String)

    ' 名前がない場合はスキップ
    If Not IsNameExists(sColName) Then
        Exit Sub
    End If

    ' 列番号を取得
    Dim nColumn As Integer
    Dim nCount As Integer
    nColumn = GetRange(sColName).Column
    nCount = GetRange(sColName).Columns.Count

    ' Range は非連続領域を46までしか設定できないので文字列でAddressを並べる
    Dim sEvenAddress As String
    Dim sOddAddress As String
    sEvenAddress = ""
    sOddAddress = ""
    For Each vCell In GetRange("選手番号")
        If vCell.MergeCells Then
            If vCell.Address = vCell.MergeArea.Item(1).Address Then
                If sEvenAddress = "" Then
                    sEvenAddress = Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                Else
                    sEvenAddress = sEvenAddress & "," & Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                End If
            
            ElseIf vCell.Address = vCell.MergeArea.Item(2).Address Then

                If sOddAddress = "" Then
                    sOddAddress = Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                Else
                    sOddAddress = sOddAddress & "," & Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                End If
            End If
        End If
    Next vCell

    ' 名前を定義
    Call DefineName(sEvenName, sEvenAddress)
    Call DefineName(sOddName, sOddAddress)
End Sub

'
' 複数列毎の名前を付ける(偶数、奇数)
'
' 列の名前で指定された範囲(選手行)に偶数行、奇数行それぞれに名前を付ける
'
' 使用例
' Call DefineNameByTripleColumns("選手名列", "選手フリガナ", "選手名", "学校名")
'
' sColName          IN      列の名前
' sFirstName        IN      １列目範囲につける名前
' sSecondName       IN      ２列目範囲につける名前
' sThirdName        IN      ３列目範囲につける名前
'
Private Sub DefineNameByTripleColumns(sColName As String, sFirstName As String, sSecondName As String, sThirdName As String)

    ' 名前がない場合はスキップ
    If Not IsNameExists(sColName) Then
        Exit Sub
    End If

    ' 列番号を取得
    Dim nColumn As Integer
    Dim nCount As Integer
    nColumn = GetRange(sColName).Column
    nCount = GetRange(sColName).Columns.Count

    ' 市民大会が6行1セットなので2行目の位置を補正する
    Dim nFirstRow As Integer
    Dim nSecondRow As Integer
    Dim nThirdRow As Integer
    If sFirstName <> "" And sSecondName <> "" Then
        If sThirdName <> "" Then
            nFirstRow = 1
            nSecondRow = 3
            nThirdRow = 5
        Else
            nFirstRow = 1
            nSecondRow = 4
            nThirdRow = 0
        End If
    End If

    ' Range は非連続領域を46までしか設定できないので文字列でAddressを並べる
    Dim sFirstAddress As String
    Dim sSecondAddress As String
    Dim sThirdAddress As String
    sFirstAddress = ""
    sSecondAddress = ""
    sThirdAddress = ""
    For Each vCell In GetRange("選手番号")
        If vCell.MergeCells Then
            If vCell.Address = vCell.MergeArea.Item(nFirstRow).Address Then
                If sFirstAddress = "" Then
                    sFirstAddress = Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                Else
                    sFirstAddress = sFirstAddress & "," & Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                End If
            
            ElseIf vCell.Address = vCell.MergeArea.Item(nSecondRow).Address Then

                If sSecondAddress = "" Then
                    sSecondAddress = Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                Else
                    sSecondAddress = sSecondAddress & "," & Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                End If
            
            ElseIf nThirdRow > 0 And vCell.Address = vCell.MergeArea.Item(nThirdRow).Address Then

                If sThirdAddress = "" Then
                    sThirdAddress = Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                Else
                    sThirdAddress = sThirdAddress & "," & Cells(vCell.Row, nColumn).Resize(1, nCount).Address(ReferenceStyle:=xlA1)
                End If
            
            End If
        End If
    Next vCell

    Call DefineName(sFirstName, sFirstAddress)
    Call DefineName(sSecondName, sSecondAddress)
    If sThirdName <> "" Then
        Call DefineName(sThirdName, sThirdAddress)
    End If
End Sub

'
' 入力制限設定
'
Private Sub 入力制限定義()
    Dim vVisible As Variant
    vVisible = SheetActivate("記入票")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    ' 入力制限全解除
    Call ClearValidation("記入票")
    
    Call DefineEntryDateValidation("申込み期間")
    
    Call DefineGenderValidation("選手性別")
    Call DefineNameValidation("選手名")
    Call DefineRubyValidation("選手フリガナ")
  
    If Range("大会名").Value = 選手権大会 Then
        
        Call DefineClassValidation("選手区分")
        Call DefineSenshukenEntryValidations("")
    
    ElseIf Range("大会名").Value = 市民大会 Then
        
        Call DefineSchoolValidation("選手学校名")
        Call DefineClassValidation("選手区分")
        Call DefineAgeValidation("選手年齢", 12)
        Call DefineShiminEntryValidations("")
        Call DefineRelayClassValidation("リレー区分")
    
    ElseIf Range("大会名").Value = マスターズ大会 Then
        
        Call DefineAgeValidation("選手年齢", 18)
        Call DefineMastersEntryValidations("")
        Call DefineRelayClassValidation("リレー区分")
    
    Else
        ' 学童大会
        Call DefineSchoolGradeValidation("選手学年")
        Call DefineGakudoEntryValidations("")
    
    End If
    
    Call DefineMinuteValidation("選手分")
    Call DefineSecondValidation("選手秒")
    Call DefineMiliSecondValidation("選手ミリ秒")
    
    Call DefineRelayStyleValidation("リレー種目")
    Call DefineMinuteValidation("リレー分")
    Call DefineSecondValidation("リレー秒")
    Call DefineMiliSecondValidation("リレーミリ秒")
    
    Sheets("記入票").Select
    Call SetForcusTop

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 入力制限全解除
'
' sSheetName        IN      シート名
'
Private Sub ClearValidation(sSheetName As String)
    Sheets(sSheetName).Select
    
    Cells.Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Call SetForcusTop
End Sub

'
' 申込み日付の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineEntryDateValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=申込み期間開始", Formula2:="=申込み期間終了"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "正しい日付を入力してください。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' 性別の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineGenderValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="男,女,　"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "性別を選択してください。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' 名前の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineNameValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeHiragana    ' ひらがな
        .ShowInput = False
        .ShowError = False
    End With
End Sub

'
' フリガナの入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineRubyValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "フリガナは自動入力されます。"
        .ErrorTitle = ""
        .InputMessage = "正しく自動入力されない場合はフリガナを上書きしてください。"
        .ErrorMessage = ""
        .IMEMode = xlIMEModeKatakana
        .ShowInput = True
        .ShowError = False
    End With
End Sub

'
' 年令の入力制限設定
'
' sName             IN      範囲の名前
' nAge              IN      年齢の低限
'
Private Sub DefineAgeValidation(sName As String, Optional nAge As Integer = 18)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=CStr(nAge), Formula2:="120"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "年令は数字だけで入力してください。"
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = CStr(nAge) & "～120までの数字を入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' 学童の学年の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineSchoolGradeValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="6"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "学年は数字だけで入力してください。"
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "1～6までの数字を入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' 選手区分の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineClassValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=選手年齢区分"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "区分を選択してください。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' 市民大会の学校入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineSchoolValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeHiragana    ' ひらがな
        .ShowInput = False
        .ShowError = False
    End With
End Sub

'
' 学童大会の種目選択の入力制限設定
'
Private Sub DefineGakudoEntryValidations(sName As String)
    Dim sTarget As String
    
    ' 50M自由形(47～52)
    sTarget = GetRange("自由形50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形50M", _
        "=AND(" & sTarget & ">=47," & sTarget & "<=52)", _
        "47：小学1・2年女子50M自由形" & vbCrLf & "48：小学1・2年男子50M自由形" & vbCrLf & _
        "49：小学3・4年女子50M自由形" & vbCrLf & "50：小学3・4年男子50M自由形" & vbCrLf & _
        "51：小学5・6年女子50M自由形" & vbCrLf & "52：小学5・6年男子50M自由形")
    '100M自由形(20～23)
    sTarget = GetRange("自由形100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形100M", _
        "=AND(" & sTarget & ">=20," & sTarget & "<=23)", _
        "20：小学4年以下女子100M自由形" & vbCrLf & "21：小学4年以下男子100M自由形" & vbCrLf & _
        "22：小学5・6年女子100M自由形" & vbCrLf & "23：小学5・6年男子100M自由形")
    ' 50M平泳ぎ(63～68)
    sTarget = GetRange("平泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ50M", _
        "=AND(" & sTarget & ">=63," & sTarget & "<=68)", _
        "63：小学1・2年女子50M平泳ぎ" & vbCrLf & "64：小学1・2年男子50M平泳ぎ" & vbCrLf & _
        "65：小学3・4年女子50M平泳ぎ" & vbCrLf & "66：小学3・4年男子50M平泳ぎ" & vbCrLf & _
        "67：小学5・6年女子50M平泳ぎ" & vbCrLf & "68：小学5・6年男子50M平泳ぎ")
    '100M平泳ぎ(32～35)
    sTarget = GetRange("平泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ100M", _
        "=AND(" & sTarget & ">=32," & sTarget & "<=35)", _
        "32：小学4年以下女子100M平泳ぎ" & vbCrLf & "33：小学4年以下男子100M平泳ぎ" & vbCrLf & _
        "34：小学5・6年女子100M平泳ぎ" & vbCrLf & "35：小学5・6年男子100M平泳ぎ")
    ' 50Mバタフライ(55～60)
    sTarget = GetRange("バタフライ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ50M", _
        "=AND(" & sTarget & ">=55," & sTarget & "<=60)", _
        "55：小学1・2年女子50Mバタフライ" & vbCrLf & "56：小学1・2年男子50Mバタフライ" & vbCrLf & _
        "57：小学3・4年女子50Mバタフライ" & vbCrLf & "58：小学3・4年男子50Mバタフライ" & vbCrLf & _
        "59：小学5・6年女子50Mバタフライ" & vbCrLf & "60：小学5・6年男子50Mバタフライ")
    '100Mバタフライ(26～29)
    sTarget = GetRange("バタフライ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ100M", _
        "=AND(" & sTarget & ">=26," & sTarget & "<=29)", _
        "26：小学4年以下女子100Mバタフライ" & vbCrLf & "27：小学4年以下男子100Mバタフライ" & vbCrLf & _
        "28：小学5・6年女子100Mバタフライ" & vbCrLf & "29：小学5・6年男子100Mバタフライ")
    ' 50M背泳ぎ(39～44)
    sTarget = GetRange("背泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ50M", _
        "=AND(" & sTarget & ">=39," & sTarget & "<=44)", _
        "39：小学1・2年女子50M背泳ぎ" & vbCrLf & "40：小学1・2年男子50M背泳ぎ" & vbCrLf & _
        "41：小学3・4年女子50M背泳ぎ" & vbCrLf & "42：小学3・4年男子50M背泳ぎ" & vbCrLf & _
        "43：小学5・6年女子50M背泳ぎ" & vbCrLf & "44：小学5・6年男子50M背泳ぎ")
    '100M背泳ぎ(14～17)
    sTarget = GetRange("背泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ100M", _
        "=AND(" & sTarget & ">=14," & sTarget & "<=17)", _
        "14：小学4年以下女子100M背泳ぎ" & vbCrLf & "15：小学4年以下男子100M背泳ぎ" & vbCrLf & _
        "16：小学5・6年女子100M背泳ぎ" & vbCrLf & "17：小学5・6年男子100M背泳ぎ")
    '200M個人メドレー(8～11)
    sTarget = GetRange("個人メドレー200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("個人メドレー200M", _
        "=AND(" & sTarget & ">=8," & sTarget & "<=11)", _
        " 8：小学4年以下女子200M個人メドレー" & vbCrLf & " 9：小学4年以下男子200M個人メドレー" & vbCrLf & _
        "10：小学5・6年女子200M個人メドレー" & vbCrLf & "11：小学5・6年男子200M個人メドレー")
    '4×50Mフリーリレー(71,72)
    sTarget = GetRange("フリーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("フリーリレー4×50M", _
        "=AND(" & sTarget & ">=71," & sTarget & "<=72)", _
        "71：小学女子4×50Mフリーリレー" & vbCrLf & "72：小学男子4×50Mフリーリレー")
    '4×50Mメドレーリレー(3,4)
    sTarget = GetRange("メドレーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("メドレーリレー4×50M", _
        "=AND(" & sTarget & ">=3," & sTarget & "<=4)", _
        "3：小学女子4×50Mメドレーリレー" & vbCrLf & "4：小学男子4×50Mメドレーリレー")
End Sub

'
' マスターズ大会の種目選択の入力制限設定
'
Private Sub DefineMastersEntryValidations(sName As String)
    Dim sTarget As String

    ' 50M自由形(45,46)
    sTarget = GetRange("自由形50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形50M", _
        "=AND(" & sTarget & ">=45," & sTarget & "<=46)", _
        "45：女子50M自由形" & vbCrLf & "46：男子50M自由形")
    '100M自由形(18,19)
    sTarget = GetRange("自由形100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形100M", _
        "=AND(" & sTarget & ">=18," & sTarget & "<=19)", _
        "18：女子100M自由形" & vbCrLf & "19：男子100M自由形")
    ' 50M平泳ぎ(61,62)
    sTarget = GetRange("平泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ50M", _
        "=AND(" & sTarget & ">=61," & sTarget & "<=62)", _
        "61：女子50M平泳ぎ" & vbCrLf & "62：男子50M平泳ぎ")
    '100M平泳ぎ(30,31)
    sTarget = GetRange("平泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ100M", _
        "=AND(" & sTarget & ">=30," & sTarget & "<=31)", _
        "30：女子100M平泳ぎ" & vbCrLf & "31：男子100M平泳ぎ")
    ' 50Mバタフライ(53,54)
    sTarget = GetRange("バタフライ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ50M", _
        "=AND(" & sTarget & ">=53," & sTarget & "<=54)", _
        "53：女子50Mバタフライ" & vbCrLf & "54：男子50Mバタフライ")
    '100Mバタフライ(24,25)
    sTarget = GetRange("バタフライ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ100M", _
        "=AND(" & sTarget & ">=24," & sTarget & "<=25)", _
        "24：女子100Mバタフライ" & vbCrLf & "25：男子100Mバタフライ")
    ' 50M背泳ぎ(37,38)
    sTarget = GetRange("背泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ50M", _
        "=AND(" & sTarget & ">=37," & sTarget & "<=38)", _
        "37：女子50M背泳ぎ" & vbCrLf & "38：男子50M背泳ぎ")
    '100M背泳ぎ(12,13)
    sTarget = GetRange("背泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ100M", _
        "=AND(" & sTarget & ">=12," & sTarget & "<=13)", _
        "12：女子100M背泳ぎ" & vbCrLf & "13：男子100M背泳ぎ")
    '200M個人メドレー(6,7)
    sTarget = GetRange("個人メドレー200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("個人メドレー200M", _
        "=AND(" & sTarget & ">=6," & sTarget & "<=7)", _
        "6：女子200M個人メドレー" & vbCrLf & "7：男子200M個人メドレー")
    '4×50Mフリーリレー(69,70)
    sTarget = GetRange("フリーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("フリーリレー4×50M", _
        "=AND(" & sTarget & ">=69," & sTarget & "<=70)", _
        "69：女子4×50Mフリーリレー" & vbCrLf & "70：男子4×50Mフリーリレー")
    '4×50Mメドレーリレー(1,2)
    sTarget = GetRange("メドレーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("メドレーリレー4×50M", _
        "=AND(" & sTarget & ">=1," & sTarget & "<=2)", _
        "1：女子4×50Mメドレーリレー" & vbCrLf & "2：男子4×50Mメドレーリレー")
    '4×50M混合フリーリレー(36)
    sTarget = GetRange("混合フリーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("混合フリーリレー4×50M", _
        "=AND(" & sTarget & ">=36," & sTarget & "<=36)", _
        "36：4×50M混合フリーリレー")
    '4×50M混合メドレーリレー(5)
    sTarget = GetRange("混合メドレーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("混合メドレーリレー4×50M", _
        "=AND(" & sTarget & ">=5," & sTarget & "<=5)", _
        "5：4×50M混合メドレーリレー")

End Sub

'
' 市民大会の種目選択の入力制限設定
'
Private Sub DefineShiminEntryValidations(sName As String)
    Dim sTarget As String
    
    ' 50M自由形(55～60)
    sTarget = GetRange("自由形50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形50M", _
        "=AND(" & sTarget & ">=55," & sTarget & "<=60)", _
        "55：中学女子50M自由形" & vbCrLf & "56：高校女子50M自由形" & vbCrLf & _
        "57：年齢区分女子50M自由形" & vbCrLf & "58：中学男子50M自由形" & vbCrLf & _
        "59：高校男子50M自由形" & vbCrLf & "60：年齢区分男子50M自由形")
    '100M自由形(37～42)
    sTarget = GetRange("自由形100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形100M", _
        "=AND(" & sTarget & ">=37," & sTarget & "<=42)", _
        "37：中学女子100M自由形" & vbCrLf & "38：高校女子100M自由形" & vbCrLf & _
        "39：年齢区分女子100M自由形" & vbCrLf & "40：中学男子100M自由形" & vbCrLf & _
        "41：高校男子100M自由形" & vbCrLf & "42：年齢区分男子100M自由形")
    '200M自由形(13～18)
    sTarget = GetRange("自由形200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形200M", _
        "=AND(" & sTarget & ">=13," & sTarget & "<=18)", _
        "13：中学女子200M自由形" & vbCrLf & "14：高校女子200M自由形" & vbCrLf & _
        "15：年齢区分女子200M自由形" & vbCrLf & "16：中学男子200M自由形" & vbCrLf & _
        "17：高校男子200M自由形" & vbCrLf & "18：年齢区分男子200M自由形")
    ' 50M平泳ぎ(61～66)
    sTarget = GetRange("平泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ50M", _
        "=AND(" & sTarget & ">=61," & sTarget & "<=66)", _
        "61：中学女子50M平泳ぎ" & vbCrLf & "62：高校女子50M平泳ぎ" & vbCrLf & _
        "63：年齢区分女子50M平泳ぎ" & vbCrLf & "64：中学男子50M平泳ぎ" & vbCrLf & _
        "65：高校男子50M平泳ぎ" & vbCrLf & "66：年齢区分男子50M平泳ぎ")
    '100M平泳ぎ(31～36)
    sTarget = GetRange("平泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ100M", _
        "=AND(" & sTarget & ">=31," & sTarget & "<=36)", _
        "31：中学女子100M平泳ぎ" & vbCrLf & "32：高校女子100M平泳ぎ" & vbCrLf & _
        "33：年齢区分女子100M平泳ぎ" & vbCrLf & "34：中学男子100M平泳ぎ" & vbCrLf & _
        "35：高校男子100M平泳ぎ" & vbCrLf & "36：年齢区分男子100M平泳ぎ")
    ' 50Mバタフライ(49～54)
    sTarget = GetRange("バタフライ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ50M", _
        "=AND(" & sTarget & ">=49," & sTarget & "<=54)", _
        "49：中学女子50Mバタフライ" & vbCrLf & "50：高校女子50Mバタフライ" & vbCrLf & _
        "51：年齢区分女子50Mバタフライ" & vbCrLf & "52：中学男子50Mバタフライ" & vbCrLf & _
        "53：高校男子50Mバタフライ" & vbCrLf & "54：年齢区分男子50Mバタフライ")
    '100Mバタフライ(25～30)
    sTarget = GetRange("バタフライ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ100M", _
        "=AND(" & sTarget & ">=25," & sTarget & "<=30)", _
        "25：中学女子100Mバタフライ" & vbCrLf & "26：高校女子100Mバタフライ" & vbCrLf & _
        "27：年齢区分女子100Mバタフライ" & vbCrLf & "28：中学男子100Mバタフライ" & vbCrLf & _
        "29：高校男子100Mバタフライ" & vbCrLf & "30：年齢区分男子100Mバタフライ")
    ' 50M背泳ぎ(43～48)
    sTarget = GetRange("背泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ50M", _
        "=AND(" & sTarget & ">=43," & sTarget & "<=48)", _
        "43：中学女子50M背泳ぎ" & vbCrLf & "44：高校女子50M背泳ぎ" & vbCrLf & _
        "45：年齢区分女子50M背泳ぎ" & vbCrLf & "46：中学男子50M背泳ぎ" & vbCrLf & _
        "47：高校男子50M背泳ぎ" & vbCrLf & "48：年齢区分男子50M背泳ぎ")
    '100M背泳ぎ(19～24)
    sTarget = GetRange("背泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ100M", _
        "=AND(" & sTarget & ">=19," & sTarget & "<=24)", _
        "19：中学女子100M背泳ぎ" & vbCrLf & "20：高校女子100M背泳ぎ" & vbCrLf & _
        "21：年齢区分女子100M背泳ぎ" & vbCrLf & "22：中学男子100M背泳ぎ" & vbCrLf & _
        "23：高校男子100M背泳ぎ" & vbCrLf & "24：年齢区分男子100M背泳ぎ")
    '200M個人メドレー(7～12)
    sTarget = GetRange("個人メドレー200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("個人メドレー200M", _
        "=AND(" & sTarget & ">=7," & sTarget & "<=12)", _
        " 7：中学女子200M個人メドレー" & vbCrLf & " 8：高校女子200M個人メドレー" & vbCrLf & _
        " 9：年齢区分女子200M個人メドレー" & vbCrLf & "10：中学男子200M個人メドレー" & vbCrLf & _
        "11：高校男子200M個人メドレー" & vbCrLf & "12：年齢区分男子200M個人メドレー")
    '4×50Mフリーリレー(67～72)
    sTarget = GetRange("フリーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("フリーリレー4×50M", _
        "=AND(" & sTarget & ">=67," & sTarget & "<=72)", _
        "67：中学女子4×50Mフリーリレー" & vbCrLf & "68：高校女子4×50Mフリーリレー" & vbCrLf & _
        "69：年齢区分女子4×50Mフリーリレー" & vbCrLf & "70：中学男子4×50Mフリーリレー" & vbCrLf & _
        "71：高校男子4×50Mフリーリレー" & vbCrLf & "72：年齢区分男子4×50Mフリーリレー")
    '4×50Mメドレーリレー(1～6)
    sTarget = GetRange("メドレーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("メドレーリレー4×50M", _
        "=AND(" & sTarget & ">=1," & sTarget & "<=6)", _
        "67：中学女子4×50Mメドレーリレー" & vbCrLf & "68：高校女子4×50Mメドレーリレー" & vbCrLf & _
        "69：年齢区分女子4×50Mメドレーリレー" & vbCrLf & "70：中学男子4×50Mメドレーリレー" & vbCrLf & _
        "71：高校男子4×50Mメドレーリレー" & vbCrLf & "72：年齢区分男子4×50Mメドレーリレー")
End Sub

'
' 選手権の種目選択の入力制限設定
'
Private Sub DefineSenshukenEntryValidations(sName As String)
    Dim sTarget As String
    
    ' 50M自由形(7～8)
    sTarget = GetRange("自由形50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形50M", _
        "=AND(" & sTarget & ">=7," & sTarget & "<=8)", _
        " 7：女子50M自由形" & vbCrLf & " 8：男子50M自由形")
    '100M自由形(15～16)
    sTarget = GetRange("自由形100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形100M", _
        "=AND(" & sTarget & ">=15," & sTarget & "<=16)", _
        "15：女子100M自由形" & vbCrLf & "16：男子100M自由形")
    '200M自由形(25～26)
    sTarget = GetRange("自由形200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("自由形200M", _
        "=AND(" & sTarget & ">=25," & sTarget & "<=26)", _
        "25：女子200M自由形" & vbCrLf & "26：男子200M自由形")
    ' 50M平泳ぎ(5～6)
    sTarget = GetRange("平泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ50M", _
        "=AND(" & sTarget & ">=5," & sTarget & "<=6)", _
        " 5：女子50M平泳ぎ" & vbCrLf & " 6：男子50M平泳ぎ")
    '100M平泳ぎ(13～14)
    sTarget = GetRange("平泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ100M", _
        "=AND(" & sTarget & ">=13," & sTarget & "<=14)", _
        "13：女子100M平泳ぎ" & vbCrLf & "14：男子100M平泳ぎ")
    '200M平泳ぎ(23～24)
    sTarget = GetRange("平泳ぎ200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("平泳ぎ200M", _
        "=AND(" & sTarget & ">=23," & sTarget & "<=24)", _
        "23：女子200M平泳ぎ" & vbCrLf & "24：男子200M平泳ぎ")
    ' 50Mバタフライ(3～4)
    sTarget = GetRange("バタフライ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ50M", _
        "=AND(" & sTarget & ">=3," & sTarget & "<=4)", _
        " 3：女子50Mバタフライ" & vbCrLf & " 4：男子50Mバタフライ")
    '100Mバタフライ(11～12)
    sTarget = GetRange("バタフライ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ100M", _
        "=AND(" & sTarget & ">=11," & sTarget & "<=12)", _
        "11：女子100Mバタフライ" & vbCrLf & "12：男子100Mバタフライ")
    '200Mバタフライ(21～22)
    sTarget = GetRange("バタフライ200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("バタフライ200M", _
        "=AND(" & sTarget & ">=21," & sTarget & "<=22)", _
        "21：女子200Mバタフライ" & vbCrLf & "22：男子200Mバタフライ")
    ' 50M背泳ぎ(1～2)
    sTarget = GetRange("背泳ぎ50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ50M", _
        "=AND(" & sTarget & ">=1," & sTarget & "<=2)", _
        " 1：女子50M背泳ぎ" & vbCrLf & " 2：男子50M背泳ぎ")
    '100M背泳ぎ(9～10)
    sTarget = GetRange("背泳ぎ100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ100M", _
        "=AND(" & sTarget & ">=9," & sTarget & "<=10)", _
        " 9：女子100M背泳ぎ" & vbCrLf & "10：男子100M背泳ぎ")
    '200M背泳ぎ(19～20)
    sTarget = GetRange("背泳ぎ200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("背泳ぎ200M", _
        "=AND(" & sTarget & ">=19," & sTarget & "<=20)", _
        "19：女子200M背泳ぎ" & vbCrLf & "20：男子200M背泳ぎ")
    '200M個人メドレー(17～18)
    sTarget = GetRange("個人メドレー200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("個人メドレー200M", _
        "=AND(" & sTarget & ">=17," & sTarget & "<=18)", _
        "17：女子200M個人メドレー" & vbCrLf & "18：男子200M個人メドレー")
    '4×50Mフリーリレー(45～46)
    sTarget = GetRange("フリーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("フリーリレー4×50M", _
        "=AND(" & sTarget & ">=45," & sTarget & "<=46)", _
        "45：女子4×50Mフリーリレー" & vbCrLf & "46：男子4×50Mフリーリレー")
    '4×50Mメドレーリレー(27～28)
    sTarget = GetRange("メドレーリレー4×50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("メドレーリレー4×50M", _
        "=AND(" & sTarget & ">=27," & sTarget & "<=28)", _
        "27：女子4×50Mメドレーリレー" & vbCrLf & "28：男子4×50Mメドレーリレー")
End Sub

'
' 種目選択の入力制限設定
'
' sName             IN      範囲の名前
' sValidationString IN      入力規制条件関数式
' sErrorMessage     IN      エラー時の文字列
'
Private Sub DefineEntryValidation(sName As String, sValidationString As String, sErrorMessage As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sValidationString
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "入力間違い"
        .InputMessage = ""
        .ErrorMessage = "プログラム番号は以下のいずれかを入力してください。" & vbCrLf & sErrorMessage
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' 分の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineMinuteValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="9"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "1～9の半角数字だけ入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' 秒の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineSecondValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="59"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "0～59の半角数字だけ入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' ミリ秒の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineMiliSecondValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="99"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "入力エラー"
        .InputMessage = ""
        .ErrorMessage = "0～99の半角数字だけ入力してください。"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' リレー年齢区分の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineRelayClassValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=リレー年齢区分"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' リレー種目番号の入力制限設定
'
' sName             IN      範囲の名前
'
Private Sub DefineRelayStyleValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=リレー種目番号"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' 条件付き書式設定
'
Private Sub 条件付き書式定義()
    Dim vVisible As Variant
    vVisible = SheetActivate("記入票")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    ' すべての条件付き書式をクリア
    Cells.FormatConditions.Delete

    Dim nIdx As Integer
    nIdx = 2
    If Range("大会名").Value = 選手権大会 Then
        
        Call DefineGenderNotification("選手性別", "選手区分")
        Call DefineNameNotification("選手名", "選手区分")
        Call DefineRubyNotification("選手フリガナ", "選手区分")
        Call DefineClassNotification("選手区分")
    
    ElseIf Range("大会名").Value = 市民大会 Then
        
        nIdx = 4
        Call DefineGenderNotification("選手性別", "選手年齢", nIdx)
        Call DefineNameNotification("選手名", "選手年齢")
        Call DefineRubyNotification("選手フリガナ", "選手年齢")
        Call DefineSchoolNotification("選手学校名")
        Call DefineClassNotification("選手区分", nIdx)
        Call DefineShiminNotification("選手年齢")
    
    ElseIf Range("大会名").Value = マスターズ大会 Then
        
        Call DefineGenderNotification("選手性別", "選手年齢")
        Call DefineNameNotification("選手名", "選手年齢")
        Call DefineRubyNotification("選手フリガナ", "選手年齢")
        Call DefineClassNotification("選手年齢")
    
    Else
        ' 学童大会
        Call DefineGenderNotification("選手性別", "選手学年")
        Call DefineNameNotification("選手名", "選手学年")
        Call DefineRubyNotification("選手フリガナ", "選手学年")
        Call DefineClassNotification("選手学年")
    
    End If
    
    Call DefineEntryNotification("選手種目偶数", 1, (nIdx - 1))
    Call DefineEntryNotification("選手種目奇数", nIdx, -(nIdx - 1))
    
    Call DefineEntryNotificationRelay("選手リレー種目")
    Call DefineSecondNotification("選手秒")
    
    If Range("大会名").Value = "横須賀マスターズ大会" Or _
        Range("大会名").Value = "横須賀市民体育大会" Then
        Call DefineRelayClassNotification("リレー区分")
    End If
    Call DefineRelayStyleNotification("リレー種目")
    Call DefineRelaySecondNotification("リレー秒")
    
    Sheets("記入票").Select
    Call SetForcusTop

    ' シートのロック
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 性別の注意表示定義
'
' sName             IN      範囲の名前
' sClassName        IN      区分範囲の名前
' nIdx              IN      ２列目の行数
'
'  =OR(AND(TRIM(選手性別)="",OR(TRIM(選手名)<>"",TRIM(選手区分)<>"", COUNTA(選手種目)>0)),
'      AND(表示種目性別1<>"",表示性別1<>"",表示種目性別1<>表示性別1),_
'      AND(表示種目性別2<>"",表示性別2<>"",表示種目性別2<>表示性別2))
'
Private Sub DefineGenderNotification(sName As String, sClassName As String, Optional nIdx As Integer = 2)
    
    Dim 選手性別 As String
    選手性別 = GetRange("選手性別").Rows(1).Address(RowAbsolute:=False)
    Dim 選手名 As String
    選手名 = GetRange("選手名").Rows(1).Address(RowAbsolute:=False)
    Dim 選手区分 As String
    選手区分 = GetRange(sClassName).Rows(1).Address(RowAbsolute:=False)
    Dim 選手種目 As String
    選手種目 = Application.Union(GetRange("選手種目偶数").Rows(1), GetRange("選手種目奇数").Rows(1)).Address(RowAbsolute:=False)
    Dim 表示種目性別1 As String
    表示種目性別1 = GetRange("表示種目性別").Rows(1).Address(RowAbsolute:=False)
    Dim 表示種目性別2 As String
    表示種目性別2 = GetRange("表示種目性別").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示性別1 As String
    表示性別1 = GetRange("表示性別").Rows(1).Address(RowAbsolute:=False)
    Dim 表示性別2 As String
    表示性別2 = GetRange("表示性別").Rows(nIdx).Address(RowAbsolute:=False)
  
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(AND(TRIM(" & 選手性別 & ")="""",OR(TRIM(" & 選手名 & ")<>""""," & _
            "TRIM(" & 選手区分 & ")<>"""",COUNTA(" & 選手種目 & ")>0))," & _
            "AND(" & 表示種目性別1 & "<>""""," & 表示性別1 & "<>""""," & 表示種目性別1 & "<>" & 表示性別1 & ")," & _
            "AND(" & 表示種目性別2 & "<>""""," & 表示性別2 & "<>""""," & 表示種目性別2 & "<>" & 表示性別2 & "))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 選手名の注意表示定義
'
' sName             IN      範囲の名前
' sClassName        IN      区分範囲の名前
'
'  =OR(AND(TRIM(選手名)="",OR(TRIM(選手区分)<>"",COUNTA(選手種目)>0)),
'      AND(TRIM(選手名)<>"",COUNTIF(選手名,"*　*")+COUNTIF(選手名,"* *")=0))
'
Private Sub DefineNameNotification(sName As String, sClassName As String)
   
    Dim 選手名 As String
    選手名 = GetRange("選手名").Rows(1).Address(RowAbsolute:=False)
    Dim 選手区分 As String
    選手区分 = GetRange(sClassName).Rows(1).Address(RowAbsolute:=False)
    Dim 選手種目 As String
    選手種目 = Application.Union(GetRange("選手種目偶数").Rows(1), GetRange("選手種目奇数").Rows(1)).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(AND(TRIM(" & 選手名 & ")="""",OR(TRIM(" & 選手区分 & ")<>"""",COUNTA(" & 選手種目 & ")>0))," & _
                "AND(TRIM(" & 選手名 & ")<>"""",COUNTIF(" & 選手名 & ",""*　*"")+COUNTIF(" & 選手名 & ",""* *"")=0))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 選手フリガナの注意表示定義
'
' sName             IN      範囲の名前
' sClassName        IN      区分範囲の名前
'
'  =AND(TRIM(選手フリガナ)="",OR(TRIM(選手名)<>"",TRIM(選手区分)<>"",COUNTA(選手種目)>0))
'
Private Sub DefineRubyNotification(sName As String, sClassName As String)
    
    Dim 選手名 As String
    選手名 = GetRange("選手名").Rows(1).Address(RowAbsolute:=False)
    Dim 選手フリガナ As String
    選手フリガナ = GetRange("選手フリガナ").Rows(1).Address(RowAbsolute:=False)
    Dim 選手区分 As String
    選手区分 = GetRange(sClassName).Rows(1).Address(RowAbsolute:=False)
    Dim 選手種目 As String
    選手種目 = Application.Union(GetRange("選手種目偶数").Rows(1), GetRange("選手種目奇数").Rows(1)).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & 選手フリガナ & ")="""",OR(TRIM(" & 選手名 & ")<>"""",TRIM(" & 選手区分 & ")<>"""",COUNTA(" & 選手種目 & ")>0))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 選手区分の注意表示定義
'
' sName             IN      範囲の名前
'
'  =OR(AND(TRIM(選手区分)="",COUNTA(選手種目)>0),
'      AND(表示種目区分1<>"",表示区分1<>"",表示種目区分1<>表示区分1),
'      AND(表示種目距離1<>"",表示距離1<>"",表示種目距離1<>表示距離1),
'      AND(表示種目区分2<>"",表示区分2<>"",表示種目区分2<>表示区分2),
'      AND(表示種目距離2<>"",表示距離2<>"",表示種目距離2<>表示距離2))
'
Private Sub DefineClassNotification(sName As String, Optional nIdx As Integer = 2)
    
    Dim 選手区分 As String
    選手区分 = GetRange(sName).Rows(1).Address(RowAbsolute:=False)
    Dim 選手種目 As String
    選手種目 = Application.Union(GetRange("選手種目偶数").Rows(1), GetRange("選手種目奇数").Rows(1)).Address(RowAbsolute:=False)
    Dim 表示種目区分1 As String
    表示種目区分1 = GetRange("表示種目区分").Rows(1).Address(RowAbsolute:=False)
    Dim 表示種目区分2 As String
    表示種目区分2 = GetRange("表示種目区分").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示種目距離1 As String
    表示種目距離1 = GetRange("表示種目距離").Rows(1).Address(RowAbsolute:=False)
    Dim 表示種目距離2 As String
    表示種目距離2 = GetRange("表示種目距離").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示区分1 As String
    表示区分1 = GetRange("表示区分").Rows(1).Address(RowAbsolute:=False)
    Dim 表示区分2 As String
    表示区分2 = GetRange("表示区分").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示距離1 As String
    表示距離1 = GetRange("表示距離").Rows(1).Address(RowAbsolute:=False)
    Dim 表示距離2 As String
    表示距離2 = GetRange("表示距離").Rows(nIdx).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(AND(TRIM(" & 選手区分 & ")="""",COUNTA(" & 選手種目 & ")>0)," & _
            "AND(" & 表示種目区分1 & "<>""""," & 表示区分1 & "<>""""," & 表示種目区分1 & "<>" & 表示区分1 & ")," & _
            "AND(" & 表示種目距離1 & "<>""""," & 表示距離1 & "<>""""," & 表示種目距離1 & "<>" & 表示距離1 & ")," & _
            "AND(" & 表示種目区分2 & "<>""""," & 表示区分2 & "<>""""," & 表示種目区分2 & "<>" & 表示区分2 & ")," & _
            "AND(" & 表示種目距離2 & "<>""""," & 表示距離2 & "<>""""," & 表示種目距離2 & "<>" & 表示距離2 & "))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 市民大会の学校名の注意表示定義
'
' sName             IN      範囲の名前
'
'  =AND(COUNTIF(チーム名,"*中学*")+COUNTIF(チーム名,"*高校*")+COUNTIF(チーム名,"*学校")=0,
'       TRIM(選手学校名)="",OR(TRIM(選手区分)="高校",TRIM(選手区分)="中学"))
'
Private Sub DefineSchoolNotification(sName As String)
    
    Dim 選手学校名 As String
    選手学校名 = GetRange("選手学校名").Rows(1).Address(RowAbsolute:=False)
    Dim 選手区分 As String
    選手区分 = GetRange("選手区分").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(COUNTIF(チーム名,""*中学*"")+COUNTIF(チーム名,""*高校*"")+COUNTIF(チーム名,""*学校"")=0," & _
            "     TRIM(" & 選手学校名 & ")="""",OR(TRIM(" & 選手区分 & ")=""高校"",TRIM(" & 選手区分 & ")=""中学""))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 市民大会の年齢の注意表示定義
'
' sName             IN      範囲の名前
'
'  =AND(TRIM(選手年齢)="",TRIM(選手区分)="年齢区分",COUNTA(選手種目)>0)
'
Private Sub DefineShiminNotification(sName As String)
    
    Dim 選手年齢 As String
    選手年齢 = GetRange("選手年齢").Rows(1).Address(RowAbsolute:=False)
    Dim 選手区分 As String
    選手区分 = GetRange("選手区分").Rows(1).Address(RowAbsolute:=False)
    Dim 選手種目 As String
    選手種目 = Application.Union(GetRange("選手種目偶数").Rows(1), GetRange("選手種目奇数").Rows(1)).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & 選手年齢 & ")="""",TRIM(" & 選手区分 & ")=""年齢区分"",COUNTA(" & 選手種目 & ")>0)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 選手種目の注意表示定義
'
' sName             IN      範囲の名前
' nIdx              IN      行の番号
' nOffset           IN      オフセット
'
'  =OR(COUNTA(選手種目)>1,AND(COUNTA(選手種目)=0,TRIM(選手秒)<>""),
'      AND(選手種目開始セル<>"",OFFSET(選手種目開始セル,1,0)<>""),
'      AND(表示種目区分<>"", 表示区分<>"", 表示種目区分<>表示区分),
'      AND(表示種目性別<>"", 表示性別<>"", 表示種目性別<>表示性別),
'      AND(表示種目距離<>"", 表示距離<>"", 表示種目距離<>表示距離))
'
Private Sub DefineEntryNotification(sName As String, nIdx As Integer, nOffset As Integer)
    
    Dim 選手種目 As String
    選手種目 = GetRange(sName).Rows(1).Address(RowAbsolute:=False)
    Dim 選手秒 As String
    選手秒 = GetRange("選手秒").Rows(nIdx).Address(RowAbsolute:=False)
    
    Dim 選手種目開始セル As String
    選手種目開始セル = GetRange(sName).Rows(1).Columns(1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Dim 表示種目区分 As String
    表示種目区分 = GetRange("表示種目区分").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示種目性別 As String
    表示種目性別 = GetRange("表示種目性別").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示種目距離 As String
    表示種目距離 = GetRange("表示種目距離").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示区分 As String
    表示区分 = GetRange("表示区分").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示性別 As String
    表示性別 = GetRange("表示性別").Rows(nIdx).Address(RowAbsolute:=False)
    Dim 表示距離 As String
    表示距離 = GetRange("表示距離").Rows(nIdx).Address(RowAbsolute:=False)

    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(COUNTA(" & 選手種目 & ")>1,AND(COUNTA(" & 選手種目 & ")=0,TRIM(" & 選手秒 & ")<>"""")," & _
            "AND(" & 選手種目開始セル & "<>"""",OFFSET(" & 選手種目開始セル & "," & nOffset & ",0)<>"""")," & _
            "AND(" & 表示種目区分 & "<>""""," & 表示区分 & "<>""""," & 表示種目区分 & "<>" & 表示区分 & ")," & _
            "AND(" & 表示種目性別 & "<>""""," & 表示性別 & "<>""""," & 表示種目性別 & "<>" & 表示性別 & ")," & _
            "AND(" & 表示種目距離 & "<>""""," & 表示距離 & "<>""""," & 表示種目距離 & "<>" & 表示距離 & "))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 選手種目の注意表示定義（リレー）
'
' sName             IN      範囲の名前
'
'   =AND(選手種目開始セル<>"",VLOOKUP(選手種目開始セル,種目番号区分,3,FALSE)<>"男女混合",VLOOKUP(選手種目開始セル,種目番号区分,3,FALSE)<>表示性別)
'
Private Sub DefineEntryNotificationRelay(sName As String)
    
    Dim 選手種目開始セル As String
    選手種目開始セル = GetRange(sName).Rows(1).Columns(1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Dim 表示性別 As String
    表示性別 = GetRange("表示性別").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(" & 選手種目開始セル & "<>"""",VLOOKUP(" & 選手種目開始セル & ",種目番号区分,3,FALSE)<>""男女混合"",VLOOKUP(" & 選手種目開始セル & ",種目番号区分,3,FALSE)<>" & 表示性別 & ")"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 選手秒の注意表示定義
'
' sName             IN      範囲の名前
'
'   =AND(COUNTA(選手種目)=1,TRIM(選手秒)="")
'
Private Sub DefineSecondNotification(sName As String)
    
    Dim 選手種目偶数 As String
    選手種目偶数 = GetRange("選手種目偶数").Rows(1).Address(RowAbsolute:=False)
    Dim 選手秒 As String
    選手秒 = GetRange("選手秒").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(COUNTA(" & 選手種目偶数 & ")=1,TRIM(" & 選手秒 & ")="""")"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' リレー区分の注意表示定義
'
' sName             IN      範囲の名前
'
'   =AND(TRIM(リレー区分)="",OR(TRIM(リレー種目)<>"",TRIM(リレー分)<>"",TRIM(リレー秒)<>"",TRIM(リレーミリ秒)<>""))
'
Private Sub DefineRelayClassNotification(sName As String)
   
    Dim リレー区分 As String
    リレー区分 = GetRange("リレー区分").Rows(1).Address(RowAbsolute:=False)
    Dim リレー種目 As String
    リレー種目 = GetRange("リレー種目").Rows(1).Address(RowAbsolute:=False)
    Dim リレー分 As String
    リレー分 = GetRange("リレー分").Rows(1).Address(RowAbsolute:=False)
    Dim リレー秒 As String
    リレー秒 = GetRange("リレー秒").Rows(1).Address(RowAbsolute:=False)
    Dim リレーミリ秒 As String
    リレーミリ秒 = GetRange("リレーミリ秒").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & リレー区分 & ")="""",OR(TRIM(" & リレー種目 & ")<>"""",TRIM(" & リレー分 & ")<>"""",TRIM(" & リレー秒 & ")<>"""",TRIM(" & リレーミリ秒 & ")<>""""))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' リレー種目の注意表示定義
'
' sName             IN      範囲の名前
'
'   =AND(TRIM(リレー種目)="",OR(TRIM(リレー分)<>"",TRIM(リレー秒)<>"",TRIM(リレーミリ秒)<>""))
'
Private Sub DefineRelayStyleNotification(sName As String)
    
    Dim リレー種目 As String
    リレー種目 = GetRange("リレー種目").Rows(1).Address(RowAbsolute:=False)
    Dim リレー分 As String
    リレー分 = GetRange("リレー分").Rows(1).Address(RowAbsolute:=False)
    Dim リレー秒 As String
    リレー秒 = GetRange("リレー秒").Rows(1).Address(RowAbsolute:=False)
    Dim リレーミリ秒 As String
    リレーミリ秒 = GetRange("リレーミリ秒").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & リレー種目 & ")="""",OR(TRIM(" & リレー分 & ")<>"""",TRIM(" & リレー秒 & ")<>"""",TRIM(" & リレーミリ秒 & ")<>""""))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' リレー秒の注意表示定義
'
' sName             IN      範囲の名前
'
'   =AND(TRIM(リレー種目)="",OR(TRIM(リレー秒)<>""))
'
Private Sub DefineRelaySecondNotification(sName As String)
    
    Dim リレー種目 As String
    リレー種目 = GetRange("リレー種目").Rows(1).Address(RowAbsolute:=False)
    Dim リレー秒 As String
    リレー秒 = GetRange("リレー秒").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & リレー種目 & ")="""",OR(TRIM(" & リレー秒 & ")<>""""))"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'
' 印刷範囲を設定する
'
Private Sub 印刷範囲の設定()
    Seehts("記入票").Select
    
    Application.PrintCommunication = True
    If Range("大会名").Value = 選手権大会 Then
        With ActiveSheet.PageSetup
            .PrintArea = "$A$1:$Z$265"
            .FitToPagesWide = 1
        End With
    Else
        With ActiveSheet.PageSetup
            .PrintArea = "$A$1:$X$265"
            .FitToPagesWide = 1
        End With
    End If
    Application.PrintCommunication = False
End Sub
