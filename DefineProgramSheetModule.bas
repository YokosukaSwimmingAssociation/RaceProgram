Attribute VB_Name = "DefineProgramSheetModule"
'
' ���O���`����
'
Sub ���[�N�u�b�N���O��`()
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    Call EventChange(False)
  
    Call Header���O��`(�t�H�[�}�b�g�V�[�g)
    Call Prog���O��`(�t�H�[�}�b�g�V�[�g)
    Call �L�^��ʖ��O��`(�L�^��ʃV�[�g)
    Call �e��ݒ薼�O��`(�ݒ�e��V�[�g)
    Call �g�b�v�y�[�W��`(�g�b�v�y�[�W�V�[�g)
    Call �V�[�g��\��
    
    Call EventChange(True)
    oWorkSheet.Activate
End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Private Sub Header���O��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
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
    Next nColumn

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Private Sub Prog���O��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
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
     
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �L�^��ʂ̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �L�^��ʖ��O��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

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
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �L�^��ʂ̈ᔽ
'
' sValue        IN      �_�~�[
'
Private Sub �L�^��ʈᔽ��`()
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
' �e��V�[�g�̐ݒ�
'
' sSheetName        IN      �V�[�g��
'
Public Sub �e��ݒ薼�O��`(sSheetName As String)

    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�ݒ�*")

    ' �V�[�g���`
    Call DefineName("�ݒ�e��", TableRangeAddress("$A$1"))

    ' �e��ݒ���s��
    For Each vGameName In GetAreaKeyData("�ݒ�e��")
        If VLookupArea(vGameName, "�ݒ�e��", "�Ώ�") = 1 Then
        
            Debug.Print vGameName
        
            ' ���O�����ׂč폜
            Call DeleteName(VLookupArea(vGameName, "�ݒ�e��", "�ϐ����擪") & "*")
        
            ' ��ڋ敪�̐ݒ�
            Call DefineTableRange(VLookupArea(vGameName, "�ݒ�e��", "��ڋ敪�V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "��ڋ敪�͈͖�"))
            ' ��ڋ敪�̐ݒ�
            If VLookupArea(vGameName, "�ݒ�e��", "��ڋ敪�֐���") <> "" Then
                Application.Run VLookupArea(vGameName, "�ݒ�e��", "��ڋ敪�֐���"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "��ڋ敪�V�[�g��")
            End If
        
            ' ���L�^�̐ݒ�
            Call DefineTableRange(VLookupArea(vGameName, "�ݒ�e��", "���L�^�V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "���L�^�͈͖�"))
        
            ' �D���҂̐ݒ�
            Call DefineColumnRange(VLookupArea(vGameName, "�ݒ�e��", "�D���҃V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "�D���Ҕ͈͖�"))
        
            ' �܏�̐ݒ�
            If VLookupArea(vGameName, "�ݒ�e��", "�܏�֐���") <> "" Then
                Application.Run VLookupArea(vGameName, "�ݒ�e��", "�܏�֐���"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "�܏�V�[�g��")
            End If
        
        End If
    Next vGameName
    
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible

End Sub

'
' �V�[�g�̃e�[�u���͈͖����`����
'
' sSheetName    IN      �V�[�g��
' sAreaName     IN      �͈͖�
'
Private Sub DefineTableRange(sSheetName As String, sAreaName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    Call DefineName(sAreaName, TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������
   
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �V�[�g�̃w�b�_�s�͈͖̔����`����
'
' sSheetName    IN      �V�[�g��
' sAreaName     IN      �͈͖�
'
Private Sub DefineColumnRange(sSheetName As String, sAreaName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    Call DefineName(sAreaName, ColumnRangeAddress("$A$1"))
    
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub


'
' �w���}�X�^�[�Y����ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �w�}����ڋ敪�ݒ�(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    Call DefineName("�w�}�N��敪", TableRangeAddress("$H$1"))
    Call DefineName("�w�}�w���敪", TableRangeAddress("$K$1"))
    Call DefineName("�w�}�w�N�\��", TableRangeAddress("$N$1"))

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �s������ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �s������ڋ敪�ݒ�(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    Call DefineName("�s���I��N��敪", RowRangeAddress("$H$1"))
    Call DefineName("�s�������[�N��敪", RowRangeAddress("$IJ$1"))
    Call DefineName("�s���N��敪", TableRangeAddress("$K$1"))

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �w�}���̏܏��`
'
' sSheetName    IN      �V�[�g��
'
Sub �w�}�܏󖼑O��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�܏�*")

    Call DefineName("�܏��ڋ敪", "$C$9")
    Call DefineName("�܏󋗗�", "$G$9")
    Call DefineName("�܏���", "$L$9")
    Call DefineName("�܏󏇈�", "$A$13")
    Call DefineName("�܏�^�C��", "$L$14")
    Call DefineName("�܏���V", "$S$14")
    Call DefineName("�܏󎁖�", "$C$20")
    Call DefineName("�܏󏊑�", "$C$24")
 
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �s�����̏܏��`
'
' sSheetName    IN      �V�[�g��
'
Sub �s���܏󖼑O��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�܏�*")

    Call DefineName("�܏󐫕�", "$AC$4")
    Call DefineName("�܏��ڋ����敪", "$AC$16")
    Call DefineName("�܏󏇈�", "$AA$7")
    Call DefineName("�܏�^�C��", "$Y$10")
    Call DefineName("�܏���V", "$Y$27")
    Call DefineName("�܏󎁖�", "$U$9")
    Call DefineName("�܏󏊑�", "$W$6")
 
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �I�茠���̏܏��`
'
' sSheetName    IN      �V�[�g��
'
Sub �I�茠�܏󖼑O��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("�܏�*")

    Call DefineName("�܏󐫕�", "$H$8")
    Call DefineName("�܏󋗗�", "$L$8")
    Call DefineName("�܏���", "$S$8")
    Call DefineName("�܏󏇈�", "$C$6")
    Call DefineName("�܏�^�C��", "$H$10")
    Call DefineName("�܏���V", "$W$10")
    Call DefineName("�܏󎁖�", "$H$12")
    Call DefineName("�܏󏊑�", "$H$14")
 
    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �g�b�v�y�[�W�̒�`
'
' sSheetName    IN      �V�[�g��
'
Private Sub �g�b�v�y�[�W��`(sSheetName As String)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    Call ����`
    Call ���N��`
    Call �v�����^��`
    Call �g����������`
    Call �g�ŏ��l����`

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = True
End Sub

'
' �����`
'
' sValue        IN      �_�~�[
'
Private Sub ����`()
    
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
Private Sub ���N��`()
    
    Call DefineName("���N", "$E$4")
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
' ���N���`
'
' sValue        IN      �_�~�[
'
Private Sub �v�����^��`()
    Call DefineName("�v�����^��", "$E$5")
End Sub


'
'�g����������`
'
' sValue        IN      �_�~�[
'
Private Sub �g����������`(Optional sValue As String = "")
    
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
Private Sub �g�ŏ��l����`(Optional sValue As String = "")

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

'
' �V�[�g��\���̐ݒ�
'
' sValue        IN      �_�~�[
'
Public Sub �V�[�g��\��(Optional sValue As String = "")

    For Each vGameName In GetAreaKeyData("�ݒ�e��")
        If VLookupArea(vGameName, "�ݒ�e��", "�Ώ�") = 1 Then
            If GetRange("��").Value = CStr(vGameName) Then
                Call SetSheetVisible(CStr(vGameName), True)
            Else
                Call SetSheetVisible(CStr(vGameName), False)
            End If
        End If
    Next vGameName

End Sub

'
' �e��V�[�g�\���^��\��
'
' vGameName IN  ��
' bFlag     IN  True:�\���^False:��\��
'
Private Sub SetSheetVisible(vGameName As String, bFlag As Boolean)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    Call SheetVisible(VLookupArea(vGameName, "�ݒ�e��", "��ڋ敪�V�[�g��"), bFlag)
    Call SheetVisible(VLookupArea(vGameName, "�ݒ�e��", "���L�^�V�[�g��"), bFlag)
    Call SheetVisible(VLookupArea(vGameName, "�ݒ�e��", "�D���҃V�[�g��"), bFlag)
    Call SheetVisible(VLookupArea(vGameName, "�ݒ�e��", "�܏�V�[�g��"), bFlag)
    ' �܏�̐ݒ�
    If bFlag And _
        VLookupArea(vGameName, "�ݒ�e��", "�܏�֐���") <> "" Then
        Application.Run VLookupArea(vGameName, "�ݒ�e��", "�܏�֐���"), _
                            VLookupArea(vGameName, "�ݒ�e��", "�܏�V�[�g��")
    End If
    oWorkSheet.Activate
End Sub

'
' �e��ݒ�V�[�g��\��
'
Sub �e��ݒ�\��()
    Call SheetVisible(�ݒ�e��V�[�g, True)
End Sub


'
' ���W���[���Ǎ���
'
Public Sub ���W���[���Ǎ���()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ImportAll(ActiveWorkbook, sPathName)
End Sub

'
' ���W���[��Export
'
Public Sub ���W���[���o��()
    Dim sPathName As String
    sPathName = SelectDir()
    Call ExportAll(ActiveWorkbook, sPathName)
End Sub
