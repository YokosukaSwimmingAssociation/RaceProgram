Attribute VB_Name = "DefineProgramSheetModule"
'
' ���O���`����
'
Sub ���[�N�u�b�N���O��`()
    Call EventChange(False)
  
    Call Header���O��`(sProgramFormatSheetName)
    Call Prog���O��`(sProgramFormatSheetName)
    Call �L�^��ʖ��O��`("�L�^���")
    Call �w���}�X�^�[�Y����ڋ敪���O��`
    Call �w���}�X�^�[�Y���L�^���O��`
    Call �w���}�X�^�[�Y���D���Җ��O��`
    Call �s������ڋ敪���O��`
    Call �s�����L�^���O��`
    Call �I�茠����ڋ敪���O��`
    Call �I�茠���L�^���O��`
    Call �܏󖼑O��`
    Call ����`
    
    Call EventChange(True)
End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Sub Header���O��`(sSheetName As String)
    Sheets(sSheetName).Select
    Call SheetProtect(False)
    Range("$A$1").Select

    ' ���O�����ׂč폜
    Call DeleteName("Header*")

    Dim oCell As Range
    Dim sName As String
    For nColumn = 1 To ActiveCell.SpecialCells(xlCellTypeLastCell).Column
        Set oCell = Cells(1, nColumn)
        sName = Trim(Replace(oCell.Value, "�@", ""))
        If sName <> "" Then
            Call SetName("Header" & sName, oCell.Address(ReferenceStyle:=xlA1))
            If sName = "����" Then
                Call SetName("Header" & sName & "�O", oCell.Offset(0, -1).Address(ReferenceStyle:=xlA1))
                Call SetName("Header" & sName & "��", oCell.Offset(0, 1).Address(ReferenceStyle:=xlA1))
            End If
        End If
    Next

     ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Sub Prog���O��`(sSheetName As String)
    Sheets(sSheetName).Select
    Call SheetProtect(False)
    Range("$A$1").Select

    ' ���O�����ׂč폜
    Call DeleteName("Prog*")

    ' �v���O�����w�b�_
    Call SetName("Prog�v��No", "$C$3")
    Call SetName("Prog��ڋ敪", "$D$3")
    Call SetName("Prog��ږ�", "$G$3")

    ' �g�w�b�_
    Call SetName("Prog�g", "$C$4")
   
    ' ���[���f�[�^
    Call SetName("Prog�g��", "$C$5")
    Call SetName("Prog���[��", "$D$5")
    Call SetName("Prog����", "$E$5")
    Call SetName("Prog�����O", "$F$5")
    Call SetName("Prog����", "$G$5")
    Call SetName("Prog������", "$H$5")
    Call SetName("Prog�敪", "$I$5")
    Call SetName("Prog����", "$J$5")
    Call SetName("Prog����", "$K$5")
    Call SetName("Prog���l", "$L$5")
    Call SetName("Prog���L�^", "$M$5")
    Call SetName("Prog�\���݋L�^", "$N$5")
    Call SetName("Prog���[�XNo", "$O$5")
    Call SetName("Prog�\�[�g�敪", "$P$5")

     ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, UserInterfaceOnly:=True
End Sub

'
' �L�^��ʂ̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �L�^��ʖ��O��`(sSheetName As String)

    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�L�^���*")

    Call SetName("�L�^��ʎ�ڔԍ�", "$B$1")
    Call SetName("�L�^��ʎ�ږ�", "$C$1")
    Call SetName("�L�^��ʑg", "$B$2")
    Call SetName("�L�^��ʃ��[�XNo", "$B$3")
    Call SetName("�L�^��ʃ��[��", "$B$5:$B$11")
    Call SetName("�L�^��ʃ^�C��", "$C$5:$C$11")
    Call SetName("�L�^��ʑI�薼", "$D$5:$D$11")
    Call SetName("�L�^��ʃ`�[����", "$E$5:$E$11")
    Call SetName("�L�^��ʑ��V", "$F$5:$F$11")

    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �w���}�X�^�[�Y����ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �w���}�X�^�[�Y����ڋ敪���O��`(Optional sValue As String = "")

    Sheets("�w���}�X�^�[�Y��ڋ敪").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�w�}*")
    
    Call SetName("�w�}��ڋ敪", TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
    
    Call SetName("�w�}�N��敪", TableRangeAddress("$H$1"))
    Call SetName("�w�}�w���敪", TableRangeAddress("$K$1"))
    Call SetName("�w�}�w�N�\��", TableRangeAddress("$N$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �w���}�X�^�[�Y���L�^�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �w���}�X�^�[�Y���L�^���O��`(Optional sValue As String = "")
    Sheets("�w���}�X�^�[�Y���L�^").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�w�}���L�^")
    
    Call SetName("�w�}���L�^", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �w���}�X�^�[�Y�D���҂̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �w���}�X�^�[�Y���D���Җ��O��`(Optional sValue As String = "")
    Sheets("�w���}�X�^�[�Y�D����").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�w�}���D����")
    
    Call SetName("�w�}���D����", ColumnRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �s������ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �s������ڋ敪���O��`(Optional sValue As String = "")

    Sheets("�s������ڋ敪").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�s��*")
    
    Call SetName("�s����ڋ敪", TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
    
    Call SetName("�s���I��N��敪", RowRangeAddress("$H$1"))
    Call SetName("�s�������[�N��敪", RowRangeAddress("$IJ$1"))
    Call SetName("�s���N��敪", TableRangeAddress("$K$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �s�����L�^�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �s�����L�^���O��`(Optional sValue As String = "")
    Sheets("�s�����L�^").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�s�����L�^")
    
    Call SetName("�s�����L�^", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �s�����D���҂̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �s�����D���Җ��O��`(Optional sValue As String = "")
    Sheets("�s�����D����").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�s�����D����")
    
    Call SetName("�s�����D����", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �I�茠����ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �I�茠����ڋ敪���O��`(Optional sValue As String = "")

    Sheets("�I�茠����ڋ敪").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�I�茠*")
    
    Call SetName("�I�茠��ڋ敪", TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
   
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �I�茠���L�^�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Sub �I�茠���L�^���O��`(Optional sValue As String = "")
    Sheets("�I�茠���L�^").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�I�茠���L�^")
    
    Call SetName("�I�茠���L�^", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' �I�茠���D���҂̖��O���`����
'
' sValue        IN      �_�~�[
'
Sub �I�茠���D���Җ��O��`(Optional sValue As String = "")
    Sheets("�I�茠���D����").Select
    Call SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�I�茠���D����")
    
    Call SetName("�I�茠���D����", TableRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
End Sub

'
' ���A�N���`
'
' sValue        IN      �_�~�[
'
Sub ����`(Optional sValue As String = "")
    '
    Sheets("�v���O�����쐬�}�N��").Select
    Call SheetProtect(False)
    
    With Range("$B$1").Validation
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
    Call SetName("��", "$B$1")
    Call SetName("���N", "$B$2")
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True)
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
