Attribute VB_Name = "BatchFileReplaceModule"
'
' �G���g���[�t�@�C���ꗗ�̓ǂݍ���
'
' �t�H���_���w�肵�āA���̒��Ɋ܂܂��G���g���[�V�[�g�i*.xlsx�j�����ׂĉr�ݍ���
'
Sub �G���g���[�t�@�C���ꊇ�ϊ�()

    ' �t�@�C���ꗗ���擾
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
    ' �t�@�C�����ɏ�������
    '
    For Each vFile In FileList
        
        ' �^�C�g���C��
        nCount = nCount + 1
        Call SetTitleMenu("�G���g���[�t�@�C���ϊ���: " & Str(nCount) & "/" & Str(nMax))
        
        '
        ' �t�@�C�����J���i�ǂݎ���p�j
        '
        Set SubBook = Workbooks.Open(Filename:=sPathName + "\" + vFile, ReadOnly:=False)
        Worksheets("�L���[").Activate

        ' �G���g���[�ꗗ�̓Ǎ���
        Call �G���g���[�t�@�C���ϊ�1
        Call �G���g���[�V�[�g��`
    
        ' �x���Ȃ��Ńt�@�C�������i�ۑ����Ȃ��j
        Application.DisplayAlerts = False
        SubBook.Close
        Application.DisplayAlerts = True
    Next
    
    Call SetTitleMenu("")
    
    
End Sub

Private Sub �G���g���[�t�@�C���ϊ�1()
    Sheets("��ڔԍ��敪").Select
    ActiveSheet.Unprotect
    Range("B1").Value = "��ڋ敪"
End Sub


