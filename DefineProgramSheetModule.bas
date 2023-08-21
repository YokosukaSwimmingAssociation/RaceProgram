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
    ' �\���^�A�N�e�B�u�^����
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
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
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' �v���O�����t�H�[�}�b�g�̃w�b�_�[���O��`
'
' sSheetName    IN      �V�[�g��
'
Private Sub Prog���O��`(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
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
    'Call DefineName("Prog����", "$L$5")
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
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' �L�^��ʂ̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �L�^��ʖ��O��`(sSheetName As String)
    ' �A�N�e�B�u�^����
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

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
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
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

    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' ���O�����ׂč폜
    Call DeleteName("�ݒ�*")

    ' �V�[�g���`
    Call DefineName("�ݒ�e��", TableRangeAddress("$A$1"))

    ' �e��ݒ���s��
    For Each vGameName In GetAreaKeyData("�ݒ�e��")
        If VLookupArea(vGameName, "�ݒ�e��", "�Ώ�") = 1 Then
        
            'Debug.Print vGameName
        
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
            If VLookupArea(vGameName, "�ݒ�e��", "���L�^�֐���") <> "" Then
                ' ����֐������{
                Application.Run VLookupArea(vGameName, "�ݒ�e��", "���L�^�֐���"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "���L�^�V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "���L�^�͈͖�")
            ElseIf VLookupArea(vGameName, "�ݒ�e��", "���L�^�V�[�g��") <> "" Then
                Call DefineRecordSheet(VLookupArea(vGameName, "�ݒ�e��", "���L�^�V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "���L�^�͈͖�"))
            End If
        
            ' �D���҂̐ݒ�
            If VLookupArea(vGameName, "�ݒ�e��", "���L�^�֐���") <> "" Then
                ' ����֐������{
                Application.Run VLookupArea(vGameName, "�ݒ�e��", "���L�^�֐���"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "�D���҃V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "�D���Ҕ͈͖�")
            ElseIf VLookupArea(vGameName, "�ݒ�e��", "�D���҃V�[�g��") <> "" Then
                Call DefineWinnerSheet(VLookupArea(vGameName, "�ݒ�e��", "�D���҃V�[�g��"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "�D���Ҕ͈͖�"))
            End If
        
            ' �܏�̐ݒ�
            If VLookupArea(vGameName, "�ݒ�e��", "�܏�֐���") <> "" Then
                Application.Run VLookupArea(vGameName, "�ݒ�e��", "�܏�֐���"), _
                                    VLookupArea(vGameName, "�ݒ�e��", "�܏�V�[�g��")
            End If
        
        End If
    Next vGameName
    
    ' �V�[�g�̃��b�N
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible

End Sub

'
' �V�[�g�̃e�[�u���͈͖����`����
'
' sSheetName    IN      �V�[�g��
' sAreaName     IN      �͈͖�
'
Private Sub DefineTableRange(sSheetName As String, sAreaName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Dim bProtect As Boolean
    bProtect = SheetProtect(False, oWorkSheet)
    
    Call DefineName(sAreaName, TableRangeAddress("$A$1")) ' ��ڔԍ�����e�v�f������

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(bProtect, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �w���}�X�^�[�Y����ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �w�}����ڋ敪�ݒ�(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Call DefineName("�w�}�N��敪", TableRangeAddress("$H$1"))
    Call DefineName("�w�}�w���敪", TableRangeAddress("$K$1"))
    Call DefineName("�w�}�w�N�\��", TableRangeAddress("$N$1"))

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �s������ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �s������ڋ敪�ݒ�(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Call DefineName("�s���I��N��敪", RowRangeAddress("$H$1"))
    Call DefineName("�s�������[�N��敪", RowRangeAddress("$IJ$1"))
    Call DefineName("�s���N��敪", TableRangeAddress("$K$1"))

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �����L�^���ڋ敪�̖��O���`����
'
' sSheetName    IN      �V�[�g��
'
Private Sub �����L�^���ڋ敪�ݒ�(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    Call DefineName("�L�^��N��敪", TableRangeAddress("$J$1"))
    Call DefineName("����N��敪", TableRangeAddress("$O$1"))

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �D���҃V�[�g�̐ݒ�
'
' sSheetName    IN      �V�[�g��
' sAreaName     IN      �͈͖�
'
'
Public Sub DefineWinnerSheet(sSheetName As String, sAreaName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' �͈͖��̐ݒ�
    Call DefineTableRange(sSheetName, sAreaName)

    ' �t�B���^�ݒ�
    Call SetAutoFilter(sAreaName, True)
    
    ' ����͈͂̐ݒ�
    Sheets(sSheetName).PageSetup.PrintArea = GetRange(sAreaName).Address

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���L�^�V�[�g�̐ݒ�
'
' sSheetName    IN      �V�[�g��
' sAreaName     IN      �͈͖�
'
'
Public Sub DefineRecordSheet(sSheetName As String, sAreaName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' �͈͖��̐ݒ�
    Call DefineTableRange(sSheetName, sAreaName)

    ' �t�B���^�ݒ�
    Call SetAutoFilter(sAreaName, True)
    
    ' ����͈͂̐ݒ�
    Dim oRange As Range
    Set oRange = GetRange(sAreaName)
    Sheets(sSheetName).PageSetup.PrintArea = oRange.Offset(0, 1).Resize(oRange.Rows.Count, oRange.Columns.Count - 1).Address

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���L�^�^�D���҃V�[�g�̐ݒ�i�I�茠�p�j
'
' sSheetName    IN      �V�[�g��
' sAreaName     IN      �͈͖�
'
'
Public Sub �I�茠���L�^�ݒ�(sSheetName As String, sAreaName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' �͈͖��̐ݒ�
    Call DefineName(sAreaName, TableRangeAddress("$A$2"))

    ' �t�B���^�ݒ�
    Call SetAutoFilter(sAreaName, False)
    
    ' �s�̍����̐ݒ�
    Call SetWinnerRowHeight(sAreaName, 16)
    
    ' ����͈͂̐ݒ�
    Dim oRange As Range
    Set oRange = GetRange(sAreaName)
    Sheets(sSheetName).PageSetup.PrintArea = oRange.Offset(-1, 2).Resize(oRange.Rows.Count + 1, oRange.Columns.Count - 2).Address

    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���L�^�^�D���҃V�[�g�̍����ݒ�i�I�茠�p�j
'
' sSheetName    IN      �V�[�g��
' nHeight       IN      ����
'
Private Sub SetWinnerRowHeight(sAreaName As String, nHeight As Integer)
    Dim vKey As Variant
    For Each vKey In GetAreaKeyData(sAreaName)
        ' �����[��ڂ�4�{
        If GetOffset(vKey, GetColIdx(sAreaName, "���")).MergeArea.Item(1).Value Like "*�����[" Then
            vKey.RowHeight = nHeight * 4
        Else
            vKey.RowHeight = nHeight
        End If
    Next vKey
End Sub

'
' �w�}���̏܏��`
'
' sSheetName    IN      �V�[�g��
'
Sub �w�}�܏󖼑O��`(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' ���O�����ׂč폜
    Call DeleteName("�܏�*")

    Call DefineName("�܏��ڋ敪", "$E$9")
    Call DefineName("�܏󋗗�", "$I$9")
    Call DefineName("�܏���", "$N$9")
    Call DefineName("�܏󏇈�", "$B$13")
    Call DefineName("�܏�^�C��", "$N$14")
    Call DefineName("�܏���V", "$U$14")
    Call DefineName("�܏󎁖�", "$D$20")
    Call DefineName("�܏󏊑�", "$D$24")
 
    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �s�����̏܏��`
'
' sSheetName    IN      �V�[�g��
'
Sub �s���܏󖼑O��`(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' ���O�����ׂč폜
    Call DeleteName("�܏�*")

    Call DefineName("�܏��ڋ敪", "$AC$4")
    Call DefineName("�܏��ڋ����敪", "$AC$16")
    Call DefineName("�܏󏇈�", "$AA$10")
    Call DefineName("�܏�^�C��", "$Y$10")
    Call DefineName("�܏���V", "$Y$27")
    Call DefineName("�܏󎁖�", "$U$9")
    Call DefineName("�܏󏊑�", "$W$6")
    
    Call DefineName("�܏���񐔂P", "$C$7")
    Call DefineName("�܏���񐔂Q", "$R$15")
    Call DefineName("�܏�N", "$F$5")
    Call DefineName("�܏�", "$F$15")
    Call DefineName("�܏��", "$F$20")
 
    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �I�茠���̏܏��`
'
' sSheetName    IN      �V�[�g��
'
Sub �I�茠�܏󖼑O��`(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

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
 
    Call DefineName("�܏����", "$C$4")
    Call DefineName("�܏�N", "$G$25")
    Call DefineName("�܏�", "$N$25")
    Call DefineName("�܏��", "$R$25")
 
    ' �V�[�g�̕\���^�ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �g�b�v�y�[�W�̒�`
'
' sSheetName    IN      �V�[�g��
'
Private Sub �g�b�v�y�[�W��`(sSheetName As String)
    ' �\���^�A�N�e�B�u�^����
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' ���O�����ׂč폜
    Call DeleteName("���*")
    
    Call ����`
    Call ���N��`
    Call �v�����^��`
    Call �g����������`
    Call �g�ŏ��l����`
    Call ���[�X�����`
    Call �ŏ����[���ԍ���`
    Call �܏�萔��`

    ' �V�[�g�̃��b�N
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = xlSheetVisible
End Sub

'
' �z��̐ݒ��`
'
' sName         IN      ���O
' sAddress      IN      �A�h���X
' sAry          IN      ���X�g
'
Private Sub DefineListValidation(sName As String, sAddress As String, sAry() As String)

    Call DefineName(sName, sAddress)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=Join(sAry, ",")
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
    
    ' �f�t�H���g�l�̐ݒ�
    If Not isAryExist(sAry, Range(sName).Value) Then
        Range(sName).Value = sAry(0)
    End If
    
End Sub

'
' �z��̐ݒ��`
'
' sName         IN      ���O
' sAddress      IN      �A�h���X
' nMin          IN      �ŏ��l
' nMax          IN      �ő�l
' nDefault      IN      �f�t�H���g�l
'
Private Sub DefineBetweenValidation(sName As String, sAddress As String, _
nMin As Integer, nMax As Integer, Optional nDefault As Variant = Empty)

    Call DefineName(sName, sAddress)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=CStr(nMin), Formula2:=CStr(nMax)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = sName & "�͐��������œ��͂��Ă��������B"
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = CStr(nMin) & "�`" & CStr(nMax) & "�܂ł̐�������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
    
    ' �f�t�H���g�l�̐ݒ�
    If Range(sName).Value < nMin Or Range(sName).Value > nMax Then
        If IsEmpty(nDefault) Then
            Range(sName).Value = ""
        Else
            Range(sName).Value = nDefault
        End If
    End If
    
End Sub

'
' �����`
'
Private Sub ����`()
    
    Dim sAry(3) As String
    sAry(0) = �w�}���
    sAry(1) = �s�����
    sAry(2) = �I�茠���
    sAry(3) = �����L�^��

    Call DefineListValidation("��", "$B$1", sAry)
    
End Sub

'
' ���N���`
'
Private Sub ���N��`()
    
    Call DefineBetweenValidation("���N", "$E$6", 2000, 2050, Year(Now))

End Sub

'
' �܏�萔���`
'
Private Sub �܏�萔��`()
    Call �܏�񐔒�`
    Call �܏�N��`
    Call �܏󌎒�`
    Call �܏����`
End Sub

'
' �܏���񐔂��`
'
Private Sub �܏�񐔒�`()
    
    Call DefineBetweenValidation("����", "$E$9", 1, 150)

End Sub

'
' �܏�N���`
'
Private Sub �܏�N��`()
    
    Call DefineName("�����N", "$E$10")

End Sub

'
' �܏󌎂��`
'
Private Sub �܏󌎒�`()
    
    Call DefineBetweenValidation("��", "$E$11", 1, 12)

End Sub

'
' �܏�����`
'
Private Sub �܏����`()
    
    Call DefineBetweenValidation("����", "$E$12", 1, 31)

End Sub


'
' �v�����^���`
'
Private Sub �v�����^��`()
    Call �v�����^����`
    Call �܏����v���r���[��`
End Sub


'
' �v�����^���`
'
Private Sub �v�����^����`()

    Dim sAry() As String
    sAry = GetPrinters
    Call DefineListValidation("���v�����^��", "$E$7", sAry)
    
End Sub

'
' �v�����^�̈ꗗ���擾����
'
Private Function GetPrinters() As String()
    ' �v�����^�ꗗ�̎擾
    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    
    ReDim sDeviceAry(oShell.Namespace(4).Items.Count) As String
    Dim i As Integer
    i = 0
    Dim vPrinters As Variant
    For Each vPrinters In oShell.Namespace(4).Items
        sDeviceAry(i) = vPrinters.Name
        i = i + 1
    Next
    GetPrinters = sDeviceAry
End Function

'
'�܏����v���r���[��`
'
Private Sub �܏����v���r���[��`()
    
    Dim sAry(2) As String
    sAry(0) = "���Ȃ�"
    sAry(1) = "����"
    Call DefineListValidation("������v���r���[", "$E$8", sAry)
    
End Sub


'
'�g����������`
'
' sValue        IN      �_�~�[
'
Private Sub �g����������`(Optional sValue As String = "")
    
    Dim sAry(2) As String
    sAry(0) = "�P������"
    sAry(1) = "������������"
    Call DefineListValidation("���g��������", "$E$5", sAry)
    
End Sub

'
' �g�ŏ��l����`
'
' sValue        IN      �_�~�[
'
Private Sub �g�ŏ��l����`(Optional sValue As String = "")

    Dim sAry(2) As String
    sAry(0) = "3"
    sAry(1) = "4"
    sAry(2) = "2"
    Call DefineListValidation("���g�ŏ��l��", "$E$2", sAry)

End Sub

'
' ���[�X�����`
'
' sValue        IN      �_�~�[
'
Private Sub ���[�X�����`(Optional sValue As String = "")

    Dim sAry(2) As String
    sAry(0) = "7"
    sAry(1) = "6"
    sAry(2) = "5"
    Call DefineListValidation("���g���[�X���", "$E$3", sAry)

End Sub

'
' �ŏ����[���ԍ���`
'
' sValue        IN      �_�~�[
'
Private Sub �ŏ����[���ԍ���`(Optional sValue As String = "")

    Dim sAry(2) As String
    sAry(0) = "3"
    sAry(1) = "2"
    sAry(2) = "1"
    Call DefineListValidation("���g�ŏ����[���ԍ�", "$E$4", sAry)

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
Public Sub �ݒ�e��\��()
    Call SheetVisible(�ݒ�e��V�[�g, True)
End Sub


