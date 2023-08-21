Attribute VB_Name = "ImportExportModule"
'
' ���W���[���Ǎ���
'
Public Sub ���W���[���Ǎ���()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ImportAll(ActiveWorkbook, sPathName)
End Sub

'
' ���W���[���̏o��
'
Public Sub ���W���[���o��()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ExportAll(ActiveWorkbook, sPathName)
End Sub


'
' ���W���[���̓Ǎ���
'
' oWorkBook  IN      WorkBook
' sPath      IN      �t�H���_�p�X
'
Public Sub ImportAll(oWorkBook As Workbook, sPath As String)
    On Error Resume Next
    
    Dim oFso        As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim sArModule() As String                   '// ���W���[���t�@�C���z��
    Dim sModule                                 '// ���W���[���t�@�C��
    Dim sExt        As String                   '// �g���q
    Dim iMsg                                    '// MsgBox�֐��߂�l
    
    iMsg = MsgBox("�����̃��W���[���͏㏑�����܂��B��낵���ł����H", vbOKCancel, "�㏑���m�F")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// �S���W���[���̃t�@�C���p�X���擾
    Dim FileList As Collection
    Set FileList = GetFiles(sPath, "\*")
        
    '// �S���W���[�������[�v
    For Each sModule In FileList
        '// �g���q���������Ŏ擾
        sExt = LCase(oFso.GetExtensionName(sModule))
        
        '// �g���q��cls�Afrm�Abas�̂����ꂩ�̏ꍇ
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// �������W���[�����폜
            Call oWorkBook.VBProject.VBComponents.Remove(oWorkBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// ���W���[����ǉ�
            Call oWorkBook.VBProject.VBComponents.Import(sModule)
            '// Import�m�F�p���O�o��
            Debug.Print sModule
        End If
    Next
End Sub

'
' ���W���[���̏o��
'
' oWorkBook  IN      WorkBook
' sPath      IN      �t�H���_�p�X
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

