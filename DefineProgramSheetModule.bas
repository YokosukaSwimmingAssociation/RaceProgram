Attribute VB_Name = "DefineProgramSheetModule"
'
' ���O���`����
'
Sub ���[�N�u�b�N���O��`()
    Call EventChange(False)
  
    Call Header���O��`(S_PROGRAM_FORMAT_SHEET_NAME)
    Call Prog���O��`(S_PROGRAM_FORMAT_SHEET_NAME)
    Call �L�^��ʖ��O��`("�L�^���")
    Call �w���}�X�^�[�Y����ڋ敪���O��`
    Call �w���}�X�^�[�Y���L�^���O��`
    Call �w���}�X�^�[�Y���D���Җ��O��`
    Call �s������ڋ敪���O��`
    Call �s�����L�^���O��`
    Call �s�����D���Җ��O��`
    Call �I�茠����ڋ敪���O��`
    Call �I�茠���L�^���O��`
    Call �I�茠���D���Җ��O��`
    Call �܏󖼑O��`
    Call �}�N���y�[�W��`
    Call �V�[�g��\��
    
    Call EventChange(True)
    Sheets("�v���O�����쐬�}�N��").Select
    Range("A1").Select
End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Sub Header���O��`(sSheetName As String)
    Sheets(sSheetName).Visible = True
    Sheets(sSheetName).Select
    Call SheetProtect(False)
    Range("$A$1").Select

    ' ���O�����ׂč폜
    Call DeleteName("Header*")

    Dim oCell As Range
    Dim sName As String
    For nColumn = 1 To ActiveCell.SpecialCells(xlCellTypeLastCell).Column
        Set oCell = Cells(1, nColumn)
        sName = STrimAll(oCell.Value)
        If sName <> "" Then
            Call DefineName("Header" & sName, oCell.Address(ReferenceStyle:=xlA1))
            If sName = "����" Then
                Call DefineName("Header" & sName & "�O", oCell.Offset(0, -1).Address(ReferenceStyle:=xlA1))
                Call DefineName("Header" & sName & "��", oCell.Offset(0, 1).Address(ReferenceStyle:=xlA1))
            End If
        End If
    Next

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    ActiveSheet.Visible = True

End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Sub Prog���O��`(sSheetName As String)
    Sheets(sSheetName).Visible = True
    Sheets(sSheetName).Select
    Call SheetProtect(False)
    Range("$A$1").Select

    ' ���O�����ׂč폜
    Call DeleteName("Prog*")

    ' �v���O�����w�b�_
    Call DefineName("Prog�v��No", "$C$3")
    Call DefineName("Prog��ڋ敪", "$D$3")
    Call DefineName("Prog��ږ�", "$F$3")
    Call DefineName("Prog����", "$I$3")
    Call DefineName("Prog�L�^", "$K$3")

    ' �g�w�b�_
    Call DefineName("Prog�g", "$C$4")
   
    ' ���[���f�[�^
    Call DefineName("Prog�g��", "$C$5")
    Call DefineName("Prog���[��", "$D$5")
    Call DefineName("Prog����", "$E$5")
    Call DefineName("Prog���", "$F$5")
    Call DefineName("Prog�����O", "$G$5")
    Call DefineName("Prog����", "$H$5")
    Call DefineName("Prog������", "$I$5")
    Call DefineName("Prog�敪", "$J$5")
    Call DefineName("Prog����", "$K$5")
    Call DefineName("Prog����", "$L$5")
    Call DefineName("Prog���l", "$M$5")
    Call DefineName("Prog���L�^", "$N$5")
    Call DefineName("Prog�\���݋L�^", "$O$5")
    Call DefineName("Prog���[�XNo", "$P$5")
    Call DefineName("Prog�\�[�g�敪", "$Q$5")
    Call DefineName("Prog�W���L�^", "$R$5")

    ' �g�w�b�_
    Call DefineName("Prog�g�w�b�_�t�H�[�}�b�g", "A$2:$R$3")
    Call DefineName("Prog�g�t�H�[�}�b�g", "A$4:$R$13")
     
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    ActiveSheet.Visible = True
End Sub

'
' �L�^��ʂ̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �L�^��ʖ��O��`(sSheetName As String)
    Sheets(sSheetName).Visible = True
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�L�^���*")

    Call DefineName("�L�^��ʎ�ڔԍ�", "$B$1")
    Call DefineName("�L�^��ʎ�ږ�", "$C$1")
    Call DefineName("�L�^��ʑg", "$B$2")
    Call DefineName("�L�^��ʃ��[�XNo", "$B$3")
    Call DefineName("�L�^��ʃ��[��", "$B$5:$B$11")
    Call DefineName("�L�^��ʃ^�C��", "$C$5:$C$11")
    Call DefineName("�L�^��ʑI�薼", "$D$5:$D$11")
    Call DefineName("�L�^��ʃ`�[����", "$E$5:$E$11")
    Call DefineName("�L�^��ʔ��l", "$F$5:$F$11")
    Call DefineName("�L�^��ʈᔽ", "$G$5:$G$11")

    Call �L�^��ʈᔽ��`

    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = True
End Sub

'
' �L�^��ʂ̈ᔽ
'
' sValue        IN      �_�~�[
'
Sub �L�^��ʈᔽ��`(Optional sValue As String = "")
    With GetRange("�L�^��ʈᔽ").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="�@,�X�^�[�g���i,���i,OP"
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
' �w���}�X�^�[�Y����ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �w���}�X�^�[�Y����ڋ敪���O��`(Optional sValue As String = "")

    Sheets("�w���}�X�^�[�Y��ڋ敪").Visible = True
    Sheets("�w���}�X�^�[�Y��ڋ敪").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�w�}*")
    
    Call DefineName("�w�}��ڋ敪", TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
    
    Call DefineName("�w�}�N��敪", TableRangeAddress("$H$1"))
    Call DefineName("�w�}�w���敪", TableRangeAddress("$K$1"))
    Call DefineName("�w�}�w�N�\��", TableRangeAddress("$N$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �w���}�X�^�[�Y���L�^�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �w���}�X�^�[�Y���L�^���O��`(Optional sValue As String = "")
    Sheets("�w���}�X�^�[�Y���L�^").Visible = True
    Sheets("�w���}�X�^�[�Y���L�^").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�w�}���L�^")
    
    Call DefineName("�w�}���L�^", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �w���}�X�^�[�Y�D���҂̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �w���}�X�^�[�Y���D���Җ��O��`(Optional sValue As String = "")
    Sheets("�w���}�X�^�[�Y�D����").Visible = True
    Sheets("�w���}�X�^�[�Y�D����").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�w�}���D����")
    
    Call DefineName("�w�}���D����", ColumnRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �s������ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �s������ڋ敪���O��`(Optional sValue As String = "")
    Sheets("�s������ڋ敪").Visible = True
    Sheets("�s������ڋ敪").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�s��*")
    
    Call DefineName("�s����ڋ敪", TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
    
    Call DefineName("�s���I��N��敪", RowRangeAddress("$H$1"))
    Call DefineName("�s�������[�N��敪", RowRangeAddress("$IJ$1"))
    Call DefineName("�s���N��敪", TableRangeAddress("$K$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �s�����L�^�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �s�����L�^���O��`(Optional sValue As String = "")
    Sheets("�s�����L�^").Visible = True
    Sheets("�s�����L�^").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�s�����L�^")
    
    Call DefineName("�s�����L�^", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �s�����D���҂̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �s�����D���Җ��O��`(Optional sValue As String = "")
    Sheets("�s�����D����").Visible = True
    Sheets("�s�����D����").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�s�����D����")
    
    Call DefineName("�s�����D����", ColumnRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �I�茠����ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �I�茠����ڋ敪���O��`(Optional sValue As String = "")
    Sheets("�I�茠����ڋ敪").Visible = True
    Sheets("�I�茠����ڋ敪").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�I�茠*")
    
    Call DefineName("�I�茠��ڋ敪", TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
   
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �I�茠���L�^�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �I�茠���L�^���O��`(Optional sValue As String = "")
    Sheets("�I�茠���L�^").Visible = True
    Sheets("�I�茠���L�^").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�I�茠���L�^")
    
    Call DefineName("�I�茠���L�^", TableRangeAddress("$A$2"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub

'
' �I�茠���D���҂̖��O���`����
'
' sValue        IN      �_�~�[
'
Sub �I�茠���D���Җ��O��`(Optional sValue As String = "")
    Sheets("�I�茠���D����").Visible = True
    Sheets("�I�茠���D����").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�I�茠���D����")
    
    Call DefineName("�I�茠���D����", ColumnRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
    ActiveSheet.Visible = False
End Sub


'
' �}�N���y�[�W�̒�`
'
' sValue        IN      �_�~�[
'
Sub �}�N���y�[�W��`(Optional sValue As String = "")

    Sheets("�v���O�����쐬�}�N��").Select
    Call SheetProtect(False)

    Call ����`
    Call ���N��`
    Call �g����������`
    Call �g�ŏ��l����`

    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �����`
'
' sValue        IN      �_�~�[
'
Sub ����`(Optional sValue As String = "")
    
    Call DefineName("��", "$B$1")
    With Range("��").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="�w���}�X�^�[�Y���,���{��s���̈���,���{��I�茠���j���"
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
' ���N���`
'
' sValue        IN      �_�~�[
'
Sub ���N��`(Optional sValue As String = "")
    
    Call DefineName("���N", "$E$7")
    With Range("���N").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="2000", Formula2:="2050"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "�J�ÔN�͐��������œ��͂��Ă��������B"
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "2000�`2050�܂ł̐�������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
    Range("���N").Value = Year(Now)

End Sub

'
'�g����������`
'
' sValue        IN      �_�~�[
'
Sub �g����������`(Optional sValue As String = "")
    
    Call DefineName("�g��������", "$E$3")
    With Range("�g��������").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="�P������,������������"
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
    Range("�g��������").Value = "�P������"
    
End Sub

'
' �g�ŏ��l����
'
' sValue        IN      �_�~�[
'
Sub �g�ŏ��l����`(Optional sValue As String = "")

    Call DefineName("�g�ŏ��l��", "$E$2")
    With Range("�g�ŏ��l��").Validation
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
    Range("�g�ŏ��l��").Value = 4

End Sub

Sub �V�[�g��\��(Optional sValue As String = "")

    If GetRange("��").Value = "���{��I�茠���j���" Then
        Call �w�}���V�[�g��\��(False)
        Call �s�����V�[�g��\��(False)
        Call �I�茠���V�[�g��\��(True)
    ElseIf GetRange("��").Value = "���{��s���̈���" Then
        Call �w�}���V�[�g��\��(False)
        Call �s�����V�[�g��\��(True)
        Call �I�茠���V�[�g��\��(False)
    Else
        Call �w�}���V�[�g��\��(True)
        Call �s�����V�[�g��\��(False)
        Call �I�茠���V�[�g��\��(False)
    End If

End Sub

'
' �w���}�X�^�[�Y�V�[�g��\��
'
' bFlag     IN  True:�\���^False:��\��
'
Sub �w�}���V�[�g��\��(bFlag As Boolean)
    Sheets("�w���}�X�^�[�Y��ڋ敪").Visible = bFlag
    Sheets("�w���}�X�^�[�Y���L�^").Visible = bFlag
    Sheets("�w���}�X�^�[�Y�D����").Visible = bFlag
    Sheets("�w���}�X�^�[�Y�܏�").Visible = bFlag
End Sub

'
' �w���}�X�^�[�Y�V�[�g��\��
'
' bFlag     IN  True:�\���^False:��\��
'
Sub �s�����V�[�g��\��(bFlag As Boolean)
    Sheets("�s������ڋ敪").Visible = bFlag
    Sheets("�s�����L�^").Visible = bFlag
    Sheets("�s�����D����").Visible = bFlag
    'Sheets("�s�����܏�").Visible = bFlag
End Sub

'
' �w���}�X�^�[�Y�V�[�g��\��
'
' bFlag     IN  True:�\���^False:��\��
'
Sub �I�茠���V�[�g��\��(bFlag As Boolean)
    Sheets("�I�茠����ڋ敪").Visible = bFlag
    Sheets("�I�茠���L�^").Visible = bFlag
    Sheets("�I�茠���D����").Visible = bFlag
    'Sheets("�I�茠���܏�").Visible = bFlag
End Sub

'
' ���W���[���Ǎ���
'
Sub ���W���[���Ǎ���()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ImportAll(ActiveWorkbook, sPathName)
End Sub

'
' ���W���[��Export
'
Sub ���W���[���o��()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ExportAll(ActiveWorkbook, sPathName)
End Sub
