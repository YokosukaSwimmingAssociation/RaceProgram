Attribute VB_Name = "FileBatchReplaceMoudle"
'
' エントリーファイル一覧の読み込み
'
' フォルダを指定して、その中に含まれるエントリーシート（*.xlsx）をすべて詠み込む
'
Sub エントリーファイル一括変換()

    ' ファイル一覧を取得
    '
    Dim sPathName As String
    sPathName = SelectDir()
    Dim FileList As Collection
    Set FileList = GetFiles(sPathName)

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
        Call SetTitleMenu("エントリーファイル変換中: " & Str(nCount) & "/" & Str(nMax))
        
        '
        ' ファイルを開く（読み取り専用）
        '
        Set SubBook = Workbooks.Open(Filename:=sPathName + "\" + vFile, ReadOnly:=False)
        Worksheets("記入票").Activate

        ' エントリー一覧の読込み
        Call エントリーシート全部定義
    
        ' 警告なしでファイルを閉じる（保存しない）
        Application.DisplayAlerts = False
        SubBook.Close
        Application.DisplayAlerts = True
    Next
    
    Call SetTitleMenu("")
    
    
End Sub


