Attribute VB_Name = "ImportExportModule"
'
' モジュール読込み
'
Public Sub モジュール読込み()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ImportAll(ActiveWorkbook, sPathName)
End Sub

'
' モジュールの出力
'
Public Sub モジュール出力()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ExportAll(ActiveWorkbook, sPathName)
End Sub


'
' モジュールの読込み
'
' oWorkBook  IN      WorkBook
' sPath      IN      フォルダパス
'
Public Sub ImportAll(oWorkBook As Workbook, sPath As String)
    On Error Resume Next
    
    Dim oFso        As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim sArModule() As String                   '// モジュールファイル配列
    Dim sModule                                 '// モジュールファイル
    Dim sExt        As String                   '// 拡張子
    Dim iMsg                                    '// MsgBox関数戻り値
    
    iMsg = MsgBox("同名のモジュールは上書きします。よろしいですか？", vbOKCancel, "上書き確認")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// 全モジュールのファイルパスを取得
    Dim FileList As Collection
    Set FileList = GetFiles(sPath, "\*")
        
    '// 全モジュールをループ
    For Each sModule In FileList
        '// 拡張子を小文字で取得
        sExt = LCase(oFso.GetExtensionName(sModule))
        
        '// 拡張子がcls、frm、basのいずれかの場合
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// 同名モジュールを削除
            Call oWorkBook.VBProject.VBComponents.Remove(oWorkBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// モジュールを追加
            Call oWorkBook.VBProject.VBComponents.Import(sModule)
            '// Import確認用ログ出力
            Debug.Print sModule
        End If
    Next
End Sub

'
' モジュールの出力
'
' oWorkBook  IN      WorkBook
' sPath      IN      フォルダパス
'
Public Sub ExportAll(oWorkBook As Workbook, sPath As String)
    On Error Resume Next
    
    Dim sFileName As String
    With ActiveWorkbook.VBProject
        Dim i As Integer
        For i = 1 To .VBComponents.Count
            Debug.Print "Type: " & .VBComponents(i).Type
            Debug.Print "Name: " & .VBComponents(i).Name
            If .VBComponents(i).Type = 1 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".bas"
                .VBComponents(i).Export sFileName
            ElseIf .VBComponents(i).Type = 2 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".cls"
                .VBComponents(i).Export sFileName
            End If
        Next i
    End With
End Sub

