Attribute VB_Name = "DefineEntrySheetModule"
'
' �G���g���[�V�[�g�ɖ��O���`����
'
Public Sub �G���g���[�V�[�g��`()
    Call EventChange(False)
    Call ���O��`
    Call ���͐�����`
    Call �����t��������`
    Call ����͈͂̐ݒ�
    Call EventChange(True)
    ActiveWorkbook.Save
End Sub

'
' �V�[�g�ɖ��O���`����
'
Private Sub ���O��`()
    Call �L���[���O��`
    Call ��ڔԍ��敪���O��`
End Sub

'
' �L���[�V�[�g�ɖ��O���`����
'
Private Sub �L���[���O��`()

    Dim vVisible As Variant
    vVisible = SheetActivate("�L���[")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���O�����ׂč폜
    Call DeleteName("*")

    ' �V�[�g�̏d�v���ڒ�`
    Call DefineName("��", "$E$4")
    Call DefineName("�`�[����", "$E$5")
    If Range("��").Value = �I�茠��� Then
        Call DefineName("�\���݊���", "$M$7")
    ElseIf Range("��").Value = �s����� Then
        Call DefineName("�\���݊���", "$L$7")
    Else
        Call DefineName("�\���݊���", "$K$7")
    End If

    ' �I��ԍ��ƃ����[�͈͂̒�`
    ' ���͈̔͂��ȍ~�̒�`�ŗ��p����
    If Range("��").Value = �s����� Then
        Call DefineName("�I��ԍ�", "$B$12:$B$71,$B$98:$B$175,$B$194:$B$270,$B$290:$B$366,$B$386:$B$462,$B$482:$B$558")
        Call DefineName("�����[�͈�", "$B$74:$B$77,$B$178:$B$181,$B$274:$B$277,$B$370:$B$373,$B$466:$B$469,$B$562:$B$565")
    Else
        Call DefineName("�I��ԍ�", "$B$12:$B$31,$B$58:$B$83,$B$102:$B$127,$B$146:$B$171,$B$190:$B$215,$B$234:$B$259")
        Call DefineName("�����[�͈�", "$B$34:$B$37,$B$86:$B$89,$B$130:$B$133,$B$174:$B$177,$B$218:$B$221,$B$262:$B$265")
    End If

    Call DefineName("�I�萫�ʗ�", "$C$10")
    Call DefineNameByColumns("�I�萫�ʗ�", "�I�萫��")

    Call DefineName("�I�薼��", "$D$10")
    Call DefineName("�I��敪��", "$F$10")
    
    If Range("��").Value = �I�茠��� Then
        Call DefineName("��ڈꗗ", "$G$10:$S$10")
        Call DefineName("��ڋ���", "$G$11:$S$11")
   
        Call DefineName("���R�`50M��", "$G$11")
        Call DefineName("���R�`100M��", "$H$11")
        Call DefineName("���R�`200M��", "$I$11")
        Call DefineName("���j��50M��", "$J$11")
        Call DefineName("���j��100M��", "$K$11")
        Call DefineName("���j��200M��", "$L$11")
        Call DefineName("�o�^�t���C50M��", "$M$11")
        Call DefineName("�o�^�t���C100M��", "$N$11")
        Call DefineName("�o�^�t���C200M��", "$O$11")
        Call DefineName("�w�j��50M��", "$P$11")
        Call DefineName("�w�j��100M��", "$Q$11")
        Call DefineName("�w�j��200M��", "$R$11")
        Call DefineName("�l���h���[200M��", "$S$11")
        Call DefineName("�t���[�����[4�~50M��", "$T$11")
        Call DefineName("���h���[�����[4�~50M��", "$U$11")
        
        Call DefineName("�I���ڗ�", "$G$11:$S$11")
        Call DefineName("�I�胊���[��ڗ�", "$T$11:$U$11")
    
        Call DefineName("�I�蕪��", "$V$12")
        Call DefineName("�I��b��", "$X$12")
        Call DefineName("�I��~���b��", "$Z$12")
    
        Call DefineName("�����[����", "$L$34")
        Call DefineName("�����[�b��", "$N$34")
        Call DefineName("�����[�~���b��", "$P$34")
    
    ElseIf Range("��").Value = �s����� Then
        Call DefineName("��ڈꗗ", "$G$10:$P$10")
        Call DefineName("��ڋ���", "$G$11:$P$11")
        
        Call DefineName("���R�`50M��", "$G$11")
        Call DefineName("���R�`100M��", "$H$11")
        Call DefineName("���R�`200M��", "$I$11")
        Call DefineName("���j��50M��", "$J$11")
        Call DefineName("���j��100M��", "$K$11")
        Call DefineName("�o�^�t���C50M��", "$L$11")
        Call DefineName("�o�^�t���C100M��", "$M$11")
        Call DefineName("�w�j��50M��", "$N$11")
        Call DefineName("�w�j��100M��", "$O$11")
        Call DefineName("�l���h���[200M��", "$P$11")
        Call DefineName("�t���[�����[4�~50M��", "$Q$11")
        Call DefineName("���h���[�����[4�~50M��", "$R$11")
        
        Call DefineName("�I���ڗ�", "$G$11:$P$11")
        Call DefineName("�I�胊���[��ڗ�", "$Q$11:$R$11")
        
        Call DefineName("�I�蕪��", "$T$12")
        Call DefineName("�I��b��", "$V$12")
        Call DefineName("�I��~���b��", "$X$12")
        
        Call DefineName("�����[�敪��", "$B$33")
        
        Call DefineName("�����[����", "$L$34")
        Call DefineName("�����[�b��", "$N$34")
        Call DefineName("�����[�~���b��", "$P$34")
    
    Else
        ' �w���}�X�^�[�Y���
        Call DefineName("��ڈꗗ", "$G$10:$O$10")
        Call DefineName("��ڋ���", "$G$11:$O$11")
    
        Call DefineName("���R�`50M��", "$G$11")
        Call DefineName("���R�`100M��", "$H$11")
        Call DefineName("���j��50M��", "$I$11")
        Call DefineName("���j��100M��", "$J$11")
        Call DefineName("�o�^�t���C50M��", "$K$11")
        Call DefineName("�o�^�t���C100M��", "$L$11")
        Call DefineName("�w�j��50M��", "$M$11")
        Call DefineName("�w�j��100M��", "$N$11")
        Call DefineName("�l���h���[200M��", "$O$11")
        Call DefineName("�t���[�����[4�~50M��", "$P$11")
        Call DefineName("���h���[�����[4�~50M��", "$Q$11")
        
        If Range("��").Value = �}�X�^�[�Y��� Then
            Call DefineName("�����t���[�����[4�~50M��", "$R$11")
            Call DefineName("�������h���[�����[4�~50M��", "$S$11")
            Call DefineName("�I�胊���[��ڗ�", "$P$11:$S$11")
            
            Call DefineName("�����[�敪��", "$B$33")
        Else
            Call DefineName("�I�胊���[��ڗ�", "$P$11:$Q$11")
        End If
        
        Call DefineName("�I���ڗ�", "$G$11:$O$11")
    
        Call DefineName("�I�蕪��", "$T$12")
        Call DefineName("�I��b��", "$V$12")
        Call DefineName("�I��~���b��", "$X$12")
    
        Call DefineName("�����[����", "$K$34")
        Call DefineName("�����[�b��", "$M$34")
        Call DefineName("�����[�~���b��", "$O$34")
    
    End If

    Call DefineName("�\����ڔԍ���", "$AB$11")
    Call DefineName("�\����ڋ敪��", "$AC$11")
    Call DefineName("�\����ڐ��ʗ�", "$AD$11")
    Call DefineName("�\����ڋ�����", "$AE$11")
    Call DefineName("�\����ږ���", "$AF$11")
    Call DefineName("�\���敪��", "$AJ$11")
    Call DefineName("�\�����ʗ�", "$AK$11")
    Call DefineName("�\��������", "$AL$11")

    Call DefineName("�����[��ڗ�", "$E$33")
    Call DefineName("�����[��ږ���", "$F$33")

    If Range("��").Value = �I�茠��� Then
        
        Call DefineNameByEvenOddColumns("�I�薼��", "�I��t���K�i", "�I�薼")
        Call DefineNameByColumns("�I��敪��", "�I��敪")
        Call DefineNameByEvenOddColumns("�I���ڗ�", "�I���ڋ���", "�I���ڊ")
    
    ElseIf Range("��").Value = �s����� Then
        
        Call DefineNameByTripleColumns("�I�薼��", "�I��t���K�i", "�I�薼", "�I��w�Z��")
        Call DefineNameByTripleColumns("�I��敪��", "�I��敪", "�I��N��", "")
        Call DefineNameByTripleColumns("�I���ڗ�", "�I���ڋ���", "�I���ڊ", "")
    
    ElseIf Range("��").Value = �}�X�^�[�Y��� Then
    
        Call DefineNameByEvenOddColumns("�I�薼��", "�I��t���K�i", "�I�薼")
        Call DefineNameByColumns("�I��敪��", "�I��N��")
        Call DefineNameByEvenOddColumns("�I���ڗ�", "�I���ڋ���", "�I���ڊ")
        
    Else
        ' �w�����
        Call DefineNameByEvenOddColumns("�I�薼��", "�I��t���K�i", "�I�薼")
        Call DefineNameByColumns("�I��敪��", "�I��w�N")
        Call DefineNameByEvenOddColumns("�I���ڗ�", "�I���ڋ���", "�I���ڊ")
    
    End If

    Call DefineNameByColumns("�I�胊���[��ڗ�", "�I�胊���[���")
                    
    Call DefineNameByColumns("���R�`50M��", "���R�`50M")
    Call DefineNameByColumns("���R�`100M��", "���R�`100M")
    Call DefineNameByColumns("���R�`200M��", "���R�`200M")
    Call DefineNameByColumns("���j��50M��", "���j��50M")
    Call DefineNameByColumns("���j��100M��", "���j��100M")
    Call DefineNameByColumns("���j��200M��", "���j��200M")
    Call DefineNameByColumns("�o�^�t���C50M��", "�o�^�t���C50M")
    Call DefineNameByColumns("�o�^�t���C100M��", "�o�^�t���C100M")
    Call DefineNameByColumns("�o�^�t���C200M��", "�o�^�t���C200M")
    Call DefineNameByColumns("�w�j��50M��", "�w�j��50M")
    Call DefineNameByColumns("�w�j��100M��", "�w�j��100M")
    Call DefineNameByColumns("�w�j��200M��", "�w�j��200M")
    Call DefineNameByColumns("�l���h���[200M��", "�l���h���[200M")
    Call DefineNameByColumns("�t���[�����[4�~50M��", "�t���[�����[4�~50M")
    Call DefineNameByColumns("���h���[�����[4�~50M��", "���h���[�����[4�~50M")
    Call DefineNameByColumns("�����t���[�����[4�~50M��", "�����t���[�����[4�~50M")
    Call DefineNameByColumns("�������h���[�����[4�~50M��", "�������h���[�����[4�~50M")
    
    Call DefineNameByColumns("�I�蕪��", "�I�蕪")
    Call DefineNameByColumns("�I��b��", "�I��b")
    Call DefineNameByColumns("�I��~���b��", "�I��~���b")

    Call DefineNameByColumns("�\����ڔԍ���", "�\����ڔԍ�")
    Call DefineNameByColumns("�\����ڋ敪��", "�\����ڋ敪")
    Call DefineNameByColumns("�\����ڐ��ʗ�", "�\����ڐ���")
    Call DefineNameByColumns("�\����ڋ�����", "�\����ڋ���")
    Call DefineNameByColumns("�\����ږ���", "�\����ږ�")
    Call DefineNameByColumns("�\���敪��", "�\���敪")
    Call DefineNameByColumns("�\�����ʗ�", "�\������")
    Call DefineNameByColumns("�\��������", "�\������")

    Call DefineNameByRelayColumns("�����[�敪��", "�����[�敪")
    Call DefineNameByRelayColumns("�����[��ڗ�", "�����[���")
    Call DefineNameByRelayColumns("�����[��ږ���", "�����[��ږ�")
    Call DefineNameByRelayColumns("�����[����", "�����[��")
    Call DefineNameByRelayColumns("�����[�b��", "�����[�b")
    Call DefineNameByRelayColumns("�����[�~���b��", "�����[�~���b")

    Sheets("�L���[").Select
    Call SetForcusTop

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible

End Sub

'
'��ڔԍ��敪�V�[�g�ɖ��O���`����
'
Private Sub ��ڔԍ��敪���O��`()
    Dim vVisible As Variant
    vVisible = SheetActivate("��ڔԍ��敪")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    If Range("��").Value = �I�茠��� Then
    
        Call DefineName("��ڔԍ��敪", TableRangeAddress("$A$1"))
        Call DefineName("�I��N��敪", RowRangeAddress("$H$2"))
        Call DefineName("�\���݊��ԊJ�n", "$J$2")
        Call DefineName("�\���݊��ԏI��", "$J$3")
        Call DefineName("�����[��ڔԍ�", RowRangeAddress("$L$2"))
    
    ElseIf Range("��").Value = �s����� Then

        Call DefineName("��ڔԍ��敪", TableRangeAddress("$A$1"))
        Call DefineName("�I��N��敪", RowRangeAddress("$G$2"))
        Call DefineName("�����[�N��敪", RowRangeAddress("$H$2"))
        Call DefineName("�\���݊��ԊJ�n", "$J$2")
        Call DefineName("�\���݊��ԏI��", "$J$3")
        Call DefineName("�����[��ڔԍ�", RowRangeAddress("$L$2"))

    ElseIf Range("��").Value = �}�X�^�[�Y��� Then
    
        Call DefineName("��ڔԍ��敪", TableRangeAddress("$A$1"))
        Call DefineName("�����[�N��敪", RowRangeAddress("$H$2"))
        Call DefineName("�\���݊��ԊJ�n", "$J$2")
        Call DefineName("�\���݊��ԏI��", "$J$3")
        Call DefineName("�����[��ڔԍ�", RowRangeAddress("$L$2"))
    
    Else
        ' �w�����
        Call DefineName("��ڔԍ��敪", TableRangeAddress("$A$1"))
        Call DefineName("�I��敪�s�a", TableRangeAddress("$G$2"))
        Call DefineName("�\���݊��ԊJ�n", "$J$2")
        Call DefineName("�\���݊��ԏI��", "$J$3")
        Call DefineName("�����[��ڔԍ�", RowRangeAddress("$L$2"))
    
    End If

    Sheets("��ڔԍ��敪").Select
    Call SetForcusTop

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �񖈂̖��O��t����
'
' ��̖��O�Ŏw�肳�ꂽ�͈�(�I��s)�ɖ��O��t����
'
' �g�p��
' Call DefineNameByColumns("�I�萫�ʗ�", "�I�萫��")
'
' sColName          IN      ��̖��O
' sName             IN      �͈͂ɂ��閼�O
'
Private Sub DefineNameByColumns(sColName As String, sName As String)

    Call DefineNameByRangeColumns(sColName, sName, "�I��ԍ�")

End Sub

'
' �����[�p�񖈂̖��O��t����
'
' ��̖��O�Ŏw�肳�ꂽ�͈�(�����[�s)�ɖ��O��t����
'
' �g�p��
' Call DefineNameByRelayColumns("�����[�敪��", "�����[�敪")
'
' sColName          IN      ��̖��O
' sName             IN      �͈͂ɂ��閼�O
'
Private Sub DefineNameByRelayColumns(sColName As String, sName As String)

    Call DefineNameByRangeColumns(sColName, sName, "�����[�͈�")

End Sub

'
' �񖈂̖��O��t����
'
' ��̖��O�Ŏw�肳�ꂽ�͈�(�I��s)�ɖ��O��t����
'
' �g�p��
' Call DefineNameByColumns("�I�萫�ʗ�", "�I�萫��")
'
' sColName          IN      ��̖��O
' sName             IN      �͈͂ɂ��閼�O
' sRangeName        IN      �擾����͈̖͂��O
'
Private Sub DefineNameByRangeColumns(sColName As String, sName As String, sRangeName As String)

    ' ���O���Ȃ��ꍇ�̓X�L�b�v
    If Not IsNameExists(sColName) Then
        Exit Sub
    End If

    ' ��ԍ����擾
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
' �����񖈂̖��O��t����(�����A�)
'
' ��̖��O�Ŏw�肳�ꂽ�͈�(�I��s)�ɋ����s�A��s���ꂼ��ɖ��O��t����
'
' �g�p��
' Call DefineNameByEvenOddColumns("�I�薼��", "�I��t���K�i", "�I�薼")
'
' sColName          IN      ��̖��O
' sEvenName         IN      �����͈͂ɂ��閼�O
' sOddName          IN      ��͈͂ɂ��閼�O
'
Private Sub DefineNameByEvenOddColumns(sColName As String, sEvenName As String, sOddName As String)

    ' ���O���Ȃ��ꍇ�̓X�L�b�v
    If Not IsNameExists(sColName) Then
        Exit Sub
    End If

    ' ��ԍ����擾
    Dim nColumn As Integer
    Dim nCount As Integer
    nColumn = GetRange(sColName).Column
    nCount = GetRange(sColName).Columns.Count

    ' Range �͔�A���̈��46�܂ł����ݒ�ł��Ȃ��̂ŕ������Address����ׂ�
    Dim sEvenAddress As String
    Dim sOddAddress As String
    sEvenAddress = ""
    sOddAddress = ""
    For Each vCell In GetRange("�I��ԍ�")
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

    ' ���O���`
    Call DefineName(sEvenName, sEvenAddress)
    Call DefineName(sOddName, sOddAddress)
End Sub

'
' �����񖈂̖��O��t����(�����A�)
'
' ��̖��O�Ŏw�肳�ꂽ�͈�(�I��s)�ɋ����s�A��s���ꂼ��ɖ��O��t����
'
' �g�p��
' Call DefineNameByTripleColumns("�I�薼��", "�I��t���K�i", "�I�薼", "�w�Z��")
'
' sColName          IN      ��̖��O
' sFirstName        IN      �P��ڔ͈͂ɂ��閼�O
' sSecondName       IN      �Q��ڔ͈͂ɂ��閼�O
' sThirdName        IN      �R��ڔ͈͂ɂ��閼�O
'
Private Sub DefineNameByTripleColumns(sColName As String, sFirstName As String, sSecondName As String, sThirdName As String)

    ' ���O���Ȃ��ꍇ�̓X�L�b�v
    If Not IsNameExists(sColName) Then
        Exit Sub
    End If

    ' ��ԍ����擾
    Dim nColumn As Integer
    Dim nCount As Integer
    nColumn = GetRange(sColName).Column
    nCount = GetRange(sColName).Columns.Count

    ' �s����6�s1�Z�b�g�Ȃ̂�2�s�ڂ̈ʒu��␳����
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

    ' Range �͔�A���̈��46�܂ł����ݒ�ł��Ȃ��̂ŕ������Address����ׂ�
    Dim sFirstAddress As String
    Dim sSecondAddress As String
    Dim sThirdAddress As String
    sFirstAddress = ""
    sSecondAddress = ""
    sThirdAddress = ""
    For Each vCell In GetRange("�I��ԍ�")
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
' ���͐����ݒ�
'
Private Sub ���͐�����`()
    Dim vVisible As Variant
    vVisible = SheetActivate("�L���[")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    ' ���͐����S����
    Call ClearValidation("�L���[")
    
    Call DefineEntryDateValidation("�\���݊���")
    
    Call DefineGenderValidation("�I�萫��")
    Call DefineNameValidation("�I�薼")
    Call DefineRubyValidation("�I��t���K�i")
  
    If Range("��").Value = �I�茠��� Then
        
        Call DefineClassValidation("�I��敪")
        Call DefineSenshukenEntryValidations("")
    
    ElseIf Range("��").Value = �s����� Then
        
        Call DefineSchoolValidation("�I��w�Z��")
        Call DefineClassValidation("�I��敪")
        Call DefineAgeValidation("�I��N��", 12)
        Call DefineShiminEntryValidations("")
        Call DefineRelayClassValidation("�����[�敪")
    
    ElseIf Range("��").Value = �}�X�^�[�Y��� Then
        
        Call DefineAgeValidation("�I��N��", 18)
        Call DefineMastersEntryValidations("")
        Call DefineRelayClassValidation("�����[�敪")
    
    Else
        ' �w�����
        Call DefineSchoolGradeValidation("�I��w�N")
        Call DefineGakudoEntryValidations("")
    
    End If
    
    Call DefineMinuteValidation("�I�蕪")
    Call DefineSecondValidation("�I��b")
    Call DefineMiliSecondValidation("�I��~���b")
    
    Call DefineRelayStyleValidation("�����[���")
    Call DefineMinuteValidation("�����[��")
    Call DefineSecondValidation("�����[�b")
    Call DefineMiliSecondValidation("�����[�~���b")
    
    Sheets("�L���[").Select
    Call SetForcusTop

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���͐����S����
'
' sSheetName        IN      �V�[�g��
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
' �\���ݓ��t�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineEntryDateValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=�\���݊��ԊJ�n", Formula2:="=�\���݊��ԏI��"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "���������t����͂��Ă��������B"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' ���ʂ̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineGenderValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="�j,��,�@"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "���ʂ�I�����Ă��������B"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' ���O�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
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
        .IMEMode = xlIMEModeHiragana    ' �Ђ炪��
        .ShowInput = False
        .ShowError = False
    End With
End Sub

'
' �t���K�i�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineRubyValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "�t���K�i�͎������͂���܂��B"
        .ErrorTitle = ""
        .InputMessage = "�������������͂���Ȃ��ꍇ�̓t���K�i���㏑�����Ă��������B"
        .ErrorMessage = ""
        .IMEMode = xlIMEModeKatakana
        .ShowInput = True
        .ShowError = False
    End With
End Sub

'
' �N�߂̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
' nAge              IN      �N��̒��
'
Private Sub DefineAgeValidation(sName As String, Optional nAge As Integer = 18)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=CStr(nAge), Formula2:="120"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "�N�߂͐��������œ��͂��Ă��������B"
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = CStr(nAge) & "�`120�܂ł̐�������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' �w���̊w�N�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineSchoolGradeValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="6"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "�w�N�͐��������œ��͂��Ă��������B"
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "1�`6�܂ł̐�������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'
' �I��敪�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineClassValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=�I��N��敪"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "�敪��I�����Ă��������B"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' �s�����̊w�Z���͐����ݒ�
'
' sName             IN      �͈̖͂��O
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
        .IMEMode = xlIMEModeHiragana    ' �Ђ炪��
        .ShowInput = False
        .ShowError = False
    End With
End Sub

'
' �w�����̎�ڑI���̓��͐����ݒ�
'
Private Sub DefineGakudoEntryValidations(sName As String)
    Dim sTarget As String
    
    ' 50M���R�`(47�`52)
    sTarget = GetRange("���R�`50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`50M", _
        "=AND(" & sTarget & ">=47," & sTarget & "<=52)", _
        "47�F���w1�E2�N���q50M���R�`" & vbCrLf & "48�F���w1�E2�N�j�q50M���R�`" & vbCrLf & _
        "49�F���w3�E4�N���q50M���R�`" & vbCrLf & "50�F���w3�E4�N�j�q50M���R�`" & vbCrLf & _
        "51�F���w5�E6�N���q50M���R�`" & vbCrLf & "52�F���w5�E6�N�j�q50M���R�`")
    '100M���R�`(20�`23)
    sTarget = GetRange("���R�`100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`100M", _
        "=AND(" & sTarget & ">=20," & sTarget & "<=23)", _
        "20�F���w4�N�ȉ����q100M���R�`" & vbCrLf & "21�F���w4�N�ȉ��j�q100M���R�`" & vbCrLf & _
        "22�F���w5�E6�N���q100M���R�`" & vbCrLf & "23�F���w5�E6�N�j�q100M���R�`")
    ' 50M���j��(63�`68)
    sTarget = GetRange("���j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��50M", _
        "=AND(" & sTarget & ">=63," & sTarget & "<=68)", _
        "63�F���w1�E2�N���q50M���j��" & vbCrLf & "64�F���w1�E2�N�j�q50M���j��" & vbCrLf & _
        "65�F���w3�E4�N���q50M���j��" & vbCrLf & "66�F���w3�E4�N�j�q50M���j��" & vbCrLf & _
        "67�F���w5�E6�N���q50M���j��" & vbCrLf & "68�F���w5�E6�N�j�q50M���j��")
    '100M���j��(32�`35)
    sTarget = GetRange("���j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��100M", _
        "=AND(" & sTarget & ">=32," & sTarget & "<=35)", _
        "32�F���w4�N�ȉ����q100M���j��" & vbCrLf & "33�F���w4�N�ȉ��j�q100M���j��" & vbCrLf & _
        "34�F���w5�E6�N���q100M���j��" & vbCrLf & "35�F���w5�E6�N�j�q100M���j��")
    ' 50M�o�^�t���C(55�`60)
    sTarget = GetRange("�o�^�t���C50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C50M", _
        "=AND(" & sTarget & ">=55," & sTarget & "<=60)", _
        "55�F���w1�E2�N���q50M�o�^�t���C" & vbCrLf & "56�F���w1�E2�N�j�q50M�o�^�t���C" & vbCrLf & _
        "57�F���w3�E4�N���q50M�o�^�t���C" & vbCrLf & "58�F���w3�E4�N�j�q50M�o�^�t���C" & vbCrLf & _
        "59�F���w5�E6�N���q50M�o�^�t���C" & vbCrLf & "60�F���w5�E6�N�j�q50M�o�^�t���C")
    '100M�o�^�t���C(26�`29)
    sTarget = GetRange("�o�^�t���C100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C100M", _
        "=AND(" & sTarget & ">=26," & sTarget & "<=29)", _
        "26�F���w4�N�ȉ����q100M�o�^�t���C" & vbCrLf & "27�F���w4�N�ȉ��j�q100M�o�^�t���C" & vbCrLf & _
        "28�F���w5�E6�N���q100M�o�^�t���C" & vbCrLf & "29�F���w5�E6�N�j�q100M�o�^�t���C")
    ' 50M�w�j��(39�`44)
    sTarget = GetRange("�w�j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��50M", _
        "=AND(" & sTarget & ">=39," & sTarget & "<=44)", _
        "39�F���w1�E2�N���q50M�w�j��" & vbCrLf & "40�F���w1�E2�N�j�q50M�w�j��" & vbCrLf & _
        "41�F���w3�E4�N���q50M�w�j��" & vbCrLf & "42�F���w3�E4�N�j�q50M�w�j��" & vbCrLf & _
        "43�F���w5�E6�N���q50M�w�j��" & vbCrLf & "44�F���w5�E6�N�j�q50M�w�j��")
    '100M�w�j��(14�`17)
    sTarget = GetRange("�w�j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��100M", _
        "=AND(" & sTarget & ">=14," & sTarget & "<=17)", _
        "14�F���w4�N�ȉ����q100M�w�j��" & vbCrLf & "15�F���w4�N�ȉ��j�q100M�w�j��" & vbCrLf & _
        "16�F���w5�E6�N���q100M�w�j��" & vbCrLf & "17�F���w5�E6�N�j�q100M�w�j��")
    '200M�l���h���[(8�`11)
    sTarget = GetRange("�l���h���[200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�l���h���[200M", _
        "=AND(" & sTarget & ">=8," & sTarget & "<=11)", _
        " 8�F���w4�N�ȉ����q200M�l���h���[" & vbCrLf & " 9�F���w4�N�ȉ��j�q200M�l���h���[" & vbCrLf & _
        "10�F���w5�E6�N���q200M�l���h���[" & vbCrLf & "11�F���w5�E6�N�j�q200M�l���h���[")
    '4�~50M�t���[�����[(71,72)
    sTarget = GetRange("�t���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�t���[�����[4�~50M", _
        "=AND(" & sTarget & ">=71," & sTarget & "<=72)", _
        "71�F���w���q4�~50M�t���[�����[" & vbCrLf & "72�F���w�j�q4�~50M�t���[�����[")
    '4�~50M���h���[�����[(3,4)
    sTarget = GetRange("���h���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���h���[�����[4�~50M", _
        "=AND(" & sTarget & ">=3," & sTarget & "<=4)", _
        "3�F���w���q4�~50M���h���[�����[" & vbCrLf & "4�F���w�j�q4�~50M���h���[�����[")
End Sub

'
' �}�X�^�[�Y���̎�ڑI���̓��͐����ݒ�
'
Private Sub DefineMastersEntryValidations(sName As String)
    Dim sTarget As String

    ' 50M���R�`(45,46)
    sTarget = GetRange("���R�`50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`50M", _
        "=AND(" & sTarget & ">=45," & sTarget & "<=46)", _
        "45�F���q50M���R�`" & vbCrLf & "46�F�j�q50M���R�`")
    '100M���R�`(18,19)
    sTarget = GetRange("���R�`100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`100M", _
        "=AND(" & sTarget & ">=18," & sTarget & "<=19)", _
        "18�F���q100M���R�`" & vbCrLf & "19�F�j�q100M���R�`")
    ' 50M���j��(61,62)
    sTarget = GetRange("���j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��50M", _
        "=AND(" & sTarget & ">=61," & sTarget & "<=62)", _
        "61�F���q50M���j��" & vbCrLf & "62�F�j�q50M���j��")
    '100M���j��(30,31)
    sTarget = GetRange("���j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��100M", _
        "=AND(" & sTarget & ">=30," & sTarget & "<=31)", _
        "30�F���q100M���j��" & vbCrLf & "31�F�j�q100M���j��")
    ' 50M�o�^�t���C(53,54)
    sTarget = GetRange("�o�^�t���C50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C50M", _
        "=AND(" & sTarget & ">=53," & sTarget & "<=54)", _
        "53�F���q50M�o�^�t���C" & vbCrLf & "54�F�j�q50M�o�^�t���C")
    '100M�o�^�t���C(24,25)
    sTarget = GetRange("�o�^�t���C100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C100M", _
        "=AND(" & sTarget & ">=24," & sTarget & "<=25)", _
        "24�F���q100M�o�^�t���C" & vbCrLf & "25�F�j�q100M�o�^�t���C")
    ' 50M�w�j��(37,38)
    sTarget = GetRange("�w�j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��50M", _
        "=AND(" & sTarget & ">=37," & sTarget & "<=38)", _
        "37�F���q50M�w�j��" & vbCrLf & "38�F�j�q50M�w�j��")
    '100M�w�j��(12,13)
    sTarget = GetRange("�w�j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��100M", _
        "=AND(" & sTarget & ">=12," & sTarget & "<=13)", _
        "12�F���q100M�w�j��" & vbCrLf & "13�F�j�q100M�w�j��")
    '200M�l���h���[(6,7)
    sTarget = GetRange("�l���h���[200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�l���h���[200M", _
        "=AND(" & sTarget & ">=6," & sTarget & "<=7)", _
        "6�F���q200M�l���h���[" & vbCrLf & "7�F�j�q200M�l���h���[")
    '4�~50M�t���[�����[(69,70)
    sTarget = GetRange("�t���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�t���[�����[4�~50M", _
        "=AND(" & sTarget & ">=69," & sTarget & "<=70)", _
        "69�F���q4�~50M�t���[�����[" & vbCrLf & "70�F�j�q4�~50M�t���[�����[")
    '4�~50M���h���[�����[(1,2)
    sTarget = GetRange("���h���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���h���[�����[4�~50M", _
        "=AND(" & sTarget & ">=1," & sTarget & "<=2)", _
        "1�F���q4�~50M���h���[�����[" & vbCrLf & "2�F�j�q4�~50M���h���[�����[")
    '4�~50M�����t���[�����[(36)
    sTarget = GetRange("�����t���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�����t���[�����[4�~50M", _
        "=AND(" & sTarget & ">=36," & sTarget & "<=36)", _
        "36�F4�~50M�����t���[�����[")
    '4�~50M�������h���[�����[(5)
    sTarget = GetRange("�������h���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�������h���[�����[4�~50M", _
        "=AND(" & sTarget & ">=5," & sTarget & "<=5)", _
        "5�F4�~50M�������h���[�����[")

End Sub

'
' �s�����̎�ڑI���̓��͐����ݒ�
'
Private Sub DefineShiminEntryValidations(sName As String)
    Dim sTarget As String
    
    ' 50M���R�`(55�`60)
    sTarget = GetRange("���R�`50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`50M", _
        "=AND(" & sTarget & ">=55," & sTarget & "<=60)", _
        "55�F���w���q50M���R�`" & vbCrLf & "56�F���Z���q50M���R�`" & vbCrLf & _
        "57�F�N��敪���q50M���R�`" & vbCrLf & "58�F���w�j�q50M���R�`" & vbCrLf & _
        "59�F���Z�j�q50M���R�`" & vbCrLf & "60�F�N��敪�j�q50M���R�`")
    '100M���R�`(37�`42)
    sTarget = GetRange("���R�`100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`100M", _
        "=AND(" & sTarget & ">=37," & sTarget & "<=42)", _
        "37�F���w���q100M���R�`" & vbCrLf & "38�F���Z���q100M���R�`" & vbCrLf & _
        "39�F�N��敪���q100M���R�`" & vbCrLf & "40�F���w�j�q100M���R�`" & vbCrLf & _
        "41�F���Z�j�q100M���R�`" & vbCrLf & "42�F�N��敪�j�q100M���R�`")
    '200M���R�`(13�`18)
    sTarget = GetRange("���R�`200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`200M", _
        "=AND(" & sTarget & ">=13," & sTarget & "<=18)", _
        "13�F���w���q200M���R�`" & vbCrLf & "14�F���Z���q200M���R�`" & vbCrLf & _
        "15�F�N��敪���q200M���R�`" & vbCrLf & "16�F���w�j�q200M���R�`" & vbCrLf & _
        "17�F���Z�j�q200M���R�`" & vbCrLf & "18�F�N��敪�j�q200M���R�`")
    ' 50M���j��(61�`66)
    sTarget = GetRange("���j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��50M", _
        "=AND(" & sTarget & ">=61," & sTarget & "<=66)", _
        "61�F���w���q50M���j��" & vbCrLf & "62�F���Z���q50M���j��" & vbCrLf & _
        "63�F�N��敪���q50M���j��" & vbCrLf & "64�F���w�j�q50M���j��" & vbCrLf & _
        "65�F���Z�j�q50M���j��" & vbCrLf & "66�F�N��敪�j�q50M���j��")
    '100M���j��(31�`36)
    sTarget = GetRange("���j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��100M", _
        "=AND(" & sTarget & ">=31," & sTarget & "<=36)", _
        "31�F���w���q100M���j��" & vbCrLf & "32�F���Z���q100M���j��" & vbCrLf & _
        "33�F�N��敪���q100M���j��" & vbCrLf & "34�F���w�j�q100M���j��" & vbCrLf & _
        "35�F���Z�j�q100M���j��" & vbCrLf & "36�F�N��敪�j�q100M���j��")
    ' 50M�o�^�t���C(49�`54)
    sTarget = GetRange("�o�^�t���C50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C50M", _
        "=AND(" & sTarget & ">=49," & sTarget & "<=54)", _
        "49�F���w���q50M�o�^�t���C" & vbCrLf & "50�F���Z���q50M�o�^�t���C" & vbCrLf & _
        "51�F�N��敪���q50M�o�^�t���C" & vbCrLf & "52�F���w�j�q50M�o�^�t���C" & vbCrLf & _
        "53�F���Z�j�q50M�o�^�t���C" & vbCrLf & "54�F�N��敪�j�q50M�o�^�t���C")
    '100M�o�^�t���C(25�`30)
    sTarget = GetRange("�o�^�t���C100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C100M", _
        "=AND(" & sTarget & ">=25," & sTarget & "<=30)", _
        "25�F���w���q100M�o�^�t���C" & vbCrLf & "26�F���Z���q100M�o�^�t���C" & vbCrLf & _
        "27�F�N��敪���q100M�o�^�t���C" & vbCrLf & "28�F���w�j�q100M�o�^�t���C" & vbCrLf & _
        "29�F���Z�j�q100M�o�^�t���C" & vbCrLf & "30�F�N��敪�j�q100M�o�^�t���C")
    ' 50M�w�j��(43�`48)
    sTarget = GetRange("�w�j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��50M", _
        "=AND(" & sTarget & ">=43," & sTarget & "<=48)", _
        "43�F���w���q50M�w�j��" & vbCrLf & "44�F���Z���q50M�w�j��" & vbCrLf & _
        "45�F�N��敪���q50M�w�j��" & vbCrLf & "46�F���w�j�q50M�w�j��" & vbCrLf & _
        "47�F���Z�j�q50M�w�j��" & vbCrLf & "48�F�N��敪�j�q50M�w�j��")
    '100M�w�j��(19�`24)
    sTarget = GetRange("�w�j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��100M", _
        "=AND(" & sTarget & ">=19," & sTarget & "<=24)", _
        "19�F���w���q100M�w�j��" & vbCrLf & "20�F���Z���q100M�w�j��" & vbCrLf & _
        "21�F�N��敪���q100M�w�j��" & vbCrLf & "22�F���w�j�q100M�w�j��" & vbCrLf & _
        "23�F���Z�j�q100M�w�j��" & vbCrLf & "24�F�N��敪�j�q100M�w�j��")
    '200M�l���h���[(7�`12)
    sTarget = GetRange("�l���h���[200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�l���h���[200M", _
        "=AND(" & sTarget & ">=7," & sTarget & "<=12)", _
        " 7�F���w���q200M�l���h���[" & vbCrLf & " 8�F���Z���q200M�l���h���[" & vbCrLf & _
        " 9�F�N��敪���q200M�l���h���[" & vbCrLf & "10�F���w�j�q200M�l���h���[" & vbCrLf & _
        "11�F���Z�j�q200M�l���h���[" & vbCrLf & "12�F�N��敪�j�q200M�l���h���[")
    '4�~50M�t���[�����[(67�`72)
    sTarget = GetRange("�t���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�t���[�����[4�~50M", _
        "=AND(" & sTarget & ">=67," & sTarget & "<=72)", _
        "67�F���w���q4�~50M�t���[�����[" & vbCrLf & "68�F���Z���q4�~50M�t���[�����[" & vbCrLf & _
        "69�F�N��敪���q4�~50M�t���[�����[" & vbCrLf & "70�F���w�j�q4�~50M�t���[�����[" & vbCrLf & _
        "71�F���Z�j�q4�~50M�t���[�����[" & vbCrLf & "72�F�N��敪�j�q4�~50M�t���[�����[")
    '4�~50M���h���[�����[(1�`6)
    sTarget = GetRange("���h���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���h���[�����[4�~50M", _
        "=AND(" & sTarget & ">=1," & sTarget & "<=6)", _
        "67�F���w���q4�~50M���h���[�����[" & vbCrLf & "68�F���Z���q4�~50M���h���[�����[" & vbCrLf & _
        "69�F�N��敪���q4�~50M���h���[�����[" & vbCrLf & "70�F���w�j�q4�~50M���h���[�����[" & vbCrLf & _
        "71�F���Z�j�q4�~50M���h���[�����[" & vbCrLf & "72�F�N��敪�j�q4�~50M���h���[�����[")
End Sub

'
' �I�茠�̎�ڑI���̓��͐����ݒ�
'
Private Sub DefineSenshukenEntryValidations(sName As String)
    Dim sTarget As String
    
    ' 50M���R�`(7�`8)
    sTarget = GetRange("���R�`50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`50M", _
        "=AND(" & sTarget & ">=7," & sTarget & "<=8)", _
        " 7�F���q50M���R�`" & vbCrLf & " 8�F�j�q50M���R�`")
    '100M���R�`(15�`16)
    sTarget = GetRange("���R�`100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`100M", _
        "=AND(" & sTarget & ">=15," & sTarget & "<=16)", _
        "15�F���q100M���R�`" & vbCrLf & "16�F�j�q100M���R�`")
    '200M���R�`(25�`26)
    sTarget = GetRange("���R�`200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���R�`200M", _
        "=AND(" & sTarget & ">=25," & sTarget & "<=26)", _
        "25�F���q200M���R�`" & vbCrLf & "26�F�j�q200M���R�`")
    ' 50M���j��(5�`6)
    sTarget = GetRange("���j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��50M", _
        "=AND(" & sTarget & ">=5," & sTarget & "<=6)", _
        " 5�F���q50M���j��" & vbCrLf & " 6�F�j�q50M���j��")
    '100M���j��(13�`14)
    sTarget = GetRange("���j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��100M", _
        "=AND(" & sTarget & ">=13," & sTarget & "<=14)", _
        "13�F���q100M���j��" & vbCrLf & "14�F�j�q100M���j��")
    '200M���j��(23�`24)
    sTarget = GetRange("���j��200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���j��200M", _
        "=AND(" & sTarget & ">=23," & sTarget & "<=24)", _
        "23�F���q200M���j��" & vbCrLf & "24�F�j�q200M���j��")
    ' 50M�o�^�t���C(3�`4)
    sTarget = GetRange("�o�^�t���C50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C50M", _
        "=AND(" & sTarget & ">=3," & sTarget & "<=4)", _
        " 3�F���q50M�o�^�t���C" & vbCrLf & " 4�F�j�q50M�o�^�t���C")
    '100M�o�^�t���C(11�`12)
    sTarget = GetRange("�o�^�t���C100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C100M", _
        "=AND(" & sTarget & ">=11," & sTarget & "<=12)", _
        "11�F���q100M�o�^�t���C" & vbCrLf & "12�F�j�q100M�o�^�t���C")
    '200M�o�^�t���C(21�`22)
    sTarget = GetRange("�o�^�t���C200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�o�^�t���C200M", _
        "=AND(" & sTarget & ">=21," & sTarget & "<=22)", _
        "21�F���q200M�o�^�t���C" & vbCrLf & "22�F�j�q200M�o�^�t���C")
    ' 50M�w�j��(1�`2)
    sTarget = GetRange("�w�j��50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��50M", _
        "=AND(" & sTarget & ">=1," & sTarget & "<=2)", _
        " 1�F���q50M�w�j��" & vbCrLf & " 2�F�j�q50M�w�j��")
    '100M�w�j��(9�`10)
    sTarget = GetRange("�w�j��100M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��100M", _
        "=AND(" & sTarget & ">=9," & sTarget & "<=10)", _
        " 9�F���q100M�w�j��" & vbCrLf & "10�F�j�q100M�w�j��")
    '200M�w�j��(19�`20)
    sTarget = GetRange("�w�j��200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�w�j��200M", _
        "=AND(" & sTarget & ">=19," & sTarget & "<=20)", _
        "19�F���q200M�w�j��" & vbCrLf & "20�F�j�q200M�w�j��")
    '200M�l���h���[(17�`18)
    sTarget = GetRange("�l���h���[200M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�l���h���[200M", _
        "=AND(" & sTarget & ">=17," & sTarget & "<=18)", _
        "17�F���q200M�l���h���[" & vbCrLf & "18�F�j�q200M�l���h���[")
    '4�~50M�t���[�����[(45�`46)
    sTarget = GetRange("�t���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("�t���[�����[4�~50M", _
        "=AND(" & sTarget & ">=45," & sTarget & "<=46)", _
        "45�F���q4�~50M�t���[�����[" & vbCrLf & "46�F�j�q4�~50M�t���[�����[")
    '4�~50M���h���[�����[(27�`28)
    sTarget = GetRange("���h���[�����[4�~50M").Rows(1).Address(RowAbsolute:=False)
    Call DefineEntryValidation("���h���[�����[4�~50M", _
        "=AND(" & sTarget & ">=27," & sTarget & "<=28)", _
        "27�F���q4�~50M���h���[�����[" & vbCrLf & "28�F�j�q4�~50M���h���[�����[")
End Sub

'
' ��ڑI���̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
' sValidationString IN      ���͋K�������֐���
' sErrorMessage     IN      �G���[���̕�����
'
Private Sub DefineEntryValidation(sName As String, sValidationString As String, sErrorMessage As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sValidationString
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���͊ԈႢ"
        .InputMessage = ""
        .ErrorMessage = "�v���O�����ԍ��͈ȉ��̂����ꂩ����͂��Ă��������B" & vbCrLf & sErrorMessage
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' ���̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineMinuteValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="9"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "1�`9�̔��p�����������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' �b�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineSecondValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="59"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "0�`59�̔��p�����������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' �~���b�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineMiliSecondValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="99"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "���̓G���["
        .InputMessage = ""
        .ErrorMessage = "0�`99�̔��p�����������͂��Ă��������B"
        .IMEMode = xlIMEModeAlpha
        .ShowInput = False
        .ShowError = True
    End With
End Sub

'
' �����[�N��敪�̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineRelayClassValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=�����[�N��敪"
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
' �����[��ڔԍ��̓��͐����ݒ�
'
' sName             IN      �͈̖͂��O
'
Private Sub DefineRelayStyleValidation(sName As String)
    With Range(sName).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=�����[��ڔԍ�"
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
' �����t�������ݒ�
'
Private Sub �����t��������`()
    Dim vVisible As Variant
    vVisible = SheetActivate("�L���[")
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    ' ���ׂĂ̏����t���������N���A
    Cells.FormatConditions.Delete

    Dim nIdx As Integer
    nIdx = 2
    If Range("��").Value = �I�茠��� Then
        
        Call DefineGenderNotification("�I�萫��", "�I��敪")
        Call DefineNameNotification("�I�薼", "�I��敪")
        Call DefineRubyNotification("�I��t���K�i", "�I��敪")
        Call DefineClassNotification("�I��敪")
    
    ElseIf Range("��").Value = �s����� Then
        
        nIdx = 4
        Call DefineGenderNotification("�I�萫��", "�I��N��", nIdx)
        Call DefineNameNotification("�I�薼", "�I��N��")
        Call DefineRubyNotification("�I��t���K�i", "�I��N��")
        Call DefineSchoolNotification("�I��w�Z��")
        Call DefineClassNotification("�I��敪", nIdx)
        Call DefineShiminNotification("�I��N��")
    
    ElseIf Range("��").Value = �}�X�^�[�Y��� Then
        
        Call DefineGenderNotification("�I�萫��", "�I��N��")
        Call DefineNameNotification("�I�薼", "�I��N��")
        Call DefineRubyNotification("�I��t���K�i", "�I��N��")
        Call DefineClassNotification("�I��N��")
    
    Else
        ' �w�����
        Call DefineGenderNotification("�I�萫��", "�I��w�N")
        Call DefineNameNotification("�I�薼", "�I��w�N")
        Call DefineRubyNotification("�I��t���K�i", "�I��w�N")
        Call DefineClassNotification("�I��w�N")
    
    End If
    
    Call DefineEntryNotification("�I���ڋ���", 1, (nIdx - 1))
    Call DefineEntryNotification("�I���ڊ", nIdx, -(nIdx - 1))
    
    Call DefineEntryNotificationRelay("�I�胊���[���")
    Call DefineSecondNotification("�I��b")
    
    If Range("��").Value = "���{��}�X�^�[�Y���" Or _
        Range("��").Value = "���{��s���̈���" Then
        Call DefineRelayClassNotification("�����[�敪")
    End If
    Call DefineRelayStyleNotification("�����[���")
    Call DefineRelaySecondNotification("�����[�b")
    
    Sheets("�L���[").Select
    Call SetForcusTop

    ' �V�[�g�̃��b�N
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���ʂ̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
' sClassName        IN      �敪�͈̖͂��O
' nIdx              IN      �Q��ڂ̍s��
'
'  =OR(AND(TRIM(�I�萫��)="",OR(TRIM(�I�薼)<>"",TRIM(�I��敪)<>"", COUNTA(�I����)>0)),
'      AND(�\����ڐ���1<>"",�\������1<>"",�\����ڐ���1<>�\������1),_
'      AND(�\����ڐ���2<>"",�\������2<>"",�\����ڐ���2<>�\������2))
'
Private Sub DefineGenderNotification(sName As String, sClassName As String, Optional nIdx As Integer = 2)
    
    Dim �I�萫�� As String
    �I�萫�� = GetRange("�I�萫��").Rows(1).Address(RowAbsolute:=False)
    Dim �I�薼 As String
    �I�薼 = GetRange("�I�薼").Rows(1).Address(RowAbsolute:=False)
    Dim �I��敪 As String
    �I��敪 = GetRange(sClassName).Rows(1).Address(RowAbsolute:=False)
    Dim �I���� As String
    �I���� = Application.Union(GetRange("�I���ڋ���").Rows(1), GetRange("�I���ڊ").Rows(1)).Address(RowAbsolute:=False)
    Dim �\����ڐ���1 As String
    �\����ڐ���1 = GetRange("�\����ڐ���").Rows(1).Address(RowAbsolute:=False)
    Dim �\����ڐ���2 As String
    �\����ڐ���2 = GetRange("�\����ڐ���").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\������1 As String
    �\������1 = GetRange("�\������").Rows(1).Address(RowAbsolute:=False)
    Dim �\������2 As String
    �\������2 = GetRange("�\������").Rows(nIdx).Address(RowAbsolute:=False)
  
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(AND(TRIM(" & �I�萫�� & ")="""",OR(TRIM(" & �I�薼 & ")<>""""," & _
            "TRIM(" & �I��敪 & ")<>"""",COUNTA(" & �I���� & ")>0))," & _
            "AND(" & �\����ڐ���1 & "<>""""," & �\������1 & "<>""""," & �\����ڐ���1 & "<>" & �\������1 & ")," & _
            "AND(" & �\����ڐ���2 & "<>""""," & �\������2 & "<>""""," & �\����ڐ���2 & "<>" & �\������2 & "))"
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
' �I�薼�̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
' sClassName        IN      �敪�͈̖͂��O
'
'  =OR(AND(TRIM(�I�薼)="",OR(TRIM(�I��敪)<>"",COUNTA(�I����)>0)),
'      AND(TRIM(�I�薼)<>"",COUNTIF(�I�薼,"*�@*")+COUNTIF(�I�薼,"* *")=0))
'
Private Sub DefineNameNotification(sName As String, sClassName As String)
   
    Dim �I�薼 As String
    �I�薼 = GetRange("�I�薼").Rows(1).Address(RowAbsolute:=False)
    Dim �I��敪 As String
    �I��敪 = GetRange(sClassName).Rows(1).Address(RowAbsolute:=False)
    Dim �I���� As String
    �I���� = Application.Union(GetRange("�I���ڋ���").Rows(1), GetRange("�I���ڊ").Rows(1)).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(AND(TRIM(" & �I�薼 & ")="""",OR(TRIM(" & �I��敪 & ")<>"""",COUNTA(" & �I���� & ")>0))," & _
                "AND(TRIM(" & �I�薼 & ")<>"""",COUNTIF(" & �I�薼 & ",""*�@*"")+COUNTIF(" & �I�薼 & ",""* *"")=0))"
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
' �I��t���K�i�̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
' sClassName        IN      �敪�͈̖͂��O
'
'  =AND(TRIM(�I��t���K�i)="",OR(TRIM(�I�薼)<>"",TRIM(�I��敪)<>"",COUNTA(�I����)>0))
'
Private Sub DefineRubyNotification(sName As String, sClassName As String)
    
    Dim �I�薼 As String
    �I�薼 = GetRange("�I�薼").Rows(1).Address(RowAbsolute:=False)
    Dim �I��t���K�i As String
    �I��t���K�i = GetRange("�I��t���K�i").Rows(1).Address(RowAbsolute:=False)
    Dim �I��敪 As String
    �I��敪 = GetRange(sClassName).Rows(1).Address(RowAbsolute:=False)
    Dim �I���� As String
    �I���� = Application.Union(GetRange("�I���ڋ���").Rows(1), GetRange("�I���ڊ").Rows(1)).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & �I��t���K�i & ")="""",OR(TRIM(" & �I�薼 & ")<>"""",TRIM(" & �I��敪 & ")<>"""",COUNTA(" & �I���� & ")>0))"
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
' �I��敪�̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'  =OR(AND(TRIM(�I��敪)="",COUNTA(�I����)>0),
'      AND(�\����ڋ敪1<>"",�\���敪1<>"",�\����ڋ敪1<>�\���敪1),
'      AND(�\����ڋ���1<>"",�\������1<>"",�\����ڋ���1<>�\������1),
'      AND(�\����ڋ敪2<>"",�\���敪2<>"",�\����ڋ敪2<>�\���敪2),
'      AND(�\����ڋ���2<>"",�\������2<>"",�\����ڋ���2<>�\������2))
'
Private Sub DefineClassNotification(sName As String, Optional nIdx As Integer = 2)
    
    Dim �I��敪 As String
    �I��敪 = GetRange(sName).Rows(1).Address(RowAbsolute:=False)
    Dim �I���� As String
    �I���� = Application.Union(GetRange("�I���ڋ���").Rows(1), GetRange("�I���ڊ").Rows(1)).Address(RowAbsolute:=False)
    Dim �\����ڋ敪1 As String
    �\����ڋ敪1 = GetRange("�\����ڋ敪").Rows(1).Address(RowAbsolute:=False)
    Dim �\����ڋ敪2 As String
    �\����ڋ敪2 = GetRange("�\����ڋ敪").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\����ڋ���1 As String
    �\����ڋ���1 = GetRange("�\����ڋ���").Rows(1).Address(RowAbsolute:=False)
    Dim �\����ڋ���2 As String
    �\����ڋ���2 = GetRange("�\����ڋ���").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\���敪1 As String
    �\���敪1 = GetRange("�\���敪").Rows(1).Address(RowAbsolute:=False)
    Dim �\���敪2 As String
    �\���敪2 = GetRange("�\���敪").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\������1 As String
    �\������1 = GetRange("�\������").Rows(1).Address(RowAbsolute:=False)
    Dim �\������2 As String
    �\������2 = GetRange("�\������").Rows(nIdx).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(AND(TRIM(" & �I��敪 & ")="""",COUNTA(" & �I���� & ")>0)," & _
            "AND(" & �\����ڋ敪1 & "<>""""," & �\���敪1 & "<>""""," & �\����ڋ敪1 & "<>" & �\���敪1 & ")," & _
            "AND(" & �\����ڋ���1 & "<>""""," & �\������1 & "<>""""," & �\����ڋ���1 & "<>" & �\������1 & ")," & _
            "AND(" & �\����ڋ敪2 & "<>""""," & �\���敪2 & "<>""""," & �\����ڋ敪2 & "<>" & �\���敪2 & ")," & _
            "AND(" & �\����ڋ���2 & "<>""""," & �\������2 & "<>""""," & �\����ڋ���2 & "<>" & �\������2 & "))"
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
' �s�����̊w�Z���̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'  =AND(COUNTIF(�`�[����,"*���w*")+COUNTIF(�`�[����,"*���Z*")+COUNTIF(�`�[����,"*�w�Z")=0,
'       TRIM(�I��w�Z��)="",OR(TRIM(�I��敪)="���Z",TRIM(�I��敪)="���w"))
'
Private Sub DefineSchoolNotification(sName As String)
    
    Dim �I��w�Z�� As String
    �I��w�Z�� = GetRange("�I��w�Z��").Rows(1).Address(RowAbsolute:=False)
    Dim �I��敪 As String
    �I��敪 = GetRange("�I��敪").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(COUNTIF(�`�[����,""*���w*"")+COUNTIF(�`�[����,""*���Z*"")+COUNTIF(�`�[����,""*�w�Z"")=0," & _
            "     TRIM(" & �I��w�Z�� & ")="""",OR(TRIM(" & �I��敪 & ")=""���Z"",TRIM(" & �I��敪 & ")=""���w""))"
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
' �s�����̔N��̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'  =AND(TRIM(�I��N��)="",TRIM(�I��敪)="�N��敪",COUNTA(�I����)>0)
'
Private Sub DefineShiminNotification(sName As String)
    
    Dim �I��N�� As String
    �I��N�� = GetRange("�I��N��").Rows(1).Address(RowAbsolute:=False)
    Dim �I��敪 As String
    �I��敪 = GetRange("�I��敪").Rows(1).Address(RowAbsolute:=False)
    Dim �I���� As String
    �I���� = Application.Union(GetRange("�I���ڋ���").Rows(1), GetRange("�I���ڊ").Rows(1)).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & �I��N�� & ")="""",TRIM(" & �I��敪 & ")=""�N��敪"",COUNTA(" & �I���� & ")>0)"
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
' �I���ڂ̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
' nIdx              IN      �s�̔ԍ�
' nOffset           IN      �I�t�Z�b�g
'
'  =OR(COUNTA(�I����)>1,AND(COUNTA(�I����)=0,TRIM(�I��b)<>""),
'      AND(�I���ڊJ�n�Z��<>"",OFFSET(�I���ڊJ�n�Z��,1,0)<>""),
'      AND(�\����ڋ敪<>"", �\���敪<>"", �\����ڋ敪<>�\���敪),
'      AND(�\����ڐ���<>"", �\������<>"", �\����ڐ���<>�\������),
'      AND(�\����ڋ���<>"", �\������<>"", �\����ڋ���<>�\������))
'
Private Sub DefineEntryNotification(sName As String, nIdx As Integer, nOffset As Integer)
    
    Dim �I���� As String
    �I���� = GetRange(sName).Rows(1).Address(RowAbsolute:=False)
    Dim �I��b As String
    �I��b = GetRange("�I��b").Rows(nIdx).Address(RowAbsolute:=False)
    
    Dim �I���ڊJ�n�Z�� As String
    �I���ڊJ�n�Z�� = GetRange(sName).Rows(1).Columns(1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Dim �\����ڋ敪 As String
    �\����ڋ敪 = GetRange("�\����ڋ敪").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\����ڐ��� As String
    �\����ڐ��� = GetRange("�\����ڐ���").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\����ڋ��� As String
    �\����ڋ��� = GetRange("�\����ڋ���").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\���敪 As String
    �\���敪 = GetRange("�\���敪").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\������ As String
    �\������ = GetRange("�\������").Rows(nIdx).Address(RowAbsolute:=False)
    Dim �\������ As String
    �\������ = GetRange("�\������").Rows(nIdx).Address(RowAbsolute:=False)

    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=OR(COUNTA(" & �I���� & ")>1,AND(COUNTA(" & �I���� & ")=0,TRIM(" & �I��b & ")<>"""")," & _
            "AND(" & �I���ڊJ�n�Z�� & "<>"""",OFFSET(" & �I���ڊJ�n�Z�� & "," & nOffset & ",0)<>"""")," & _
            "AND(" & �\����ڋ敪 & "<>""""," & �\���敪 & "<>""""," & �\����ڋ敪 & "<>" & �\���敪 & ")," & _
            "AND(" & �\����ڐ��� & "<>""""," & �\������ & "<>""""," & �\����ڐ��� & "<>" & �\������ & ")," & _
            "AND(" & �\����ڋ��� & "<>""""," & �\������ & "<>""""," & �\����ڋ��� & "<>" & �\������ & "))"
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
' �I���ڂ̒��ӕ\����`�i�����[�j
'
' sName             IN      �͈̖͂��O
'
'   =AND(�I���ڊJ�n�Z��<>"",VLOOKUP(�I���ڊJ�n�Z��,��ڔԍ��敪,3,FALSE)<>"�j������",VLOOKUP(�I���ڊJ�n�Z��,��ڔԍ��敪,3,FALSE)<>�\������)
'
Private Sub DefineEntryNotificationRelay(sName As String)
    
    Dim �I���ڊJ�n�Z�� As String
    �I���ڊJ�n�Z�� = GetRange(sName).Rows(1).Columns(1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Dim �\������ As String
    �\������ = GetRange("�\������").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(" & �I���ڊJ�n�Z�� & "<>"""",VLOOKUP(" & �I���ڊJ�n�Z�� & ",��ڔԍ��敪,3,FALSE)<>""�j������"",VLOOKUP(" & �I���ڊJ�n�Z�� & ",��ڔԍ��敪,3,FALSE)<>" & �\������ & ")"
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
' �I��b�̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'   =AND(COUNTA(�I����)=1,TRIM(�I��b)="")
'
Private Sub DefineSecondNotification(sName As String)
    
    Dim �I���ڋ��� As String
    �I���ڋ��� = GetRange("�I���ڋ���").Rows(1).Address(RowAbsolute:=False)
    Dim �I��b As String
    �I��b = GetRange("�I��b").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(COUNTA(" & �I���ڋ��� & ")=1,TRIM(" & �I��b & ")="""")"
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
' �����[�敪�̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'   =AND(TRIM(�����[�敪)="",OR(TRIM(�����[���)<>"",TRIM(�����[��)<>"",TRIM(�����[�b)<>"",TRIM(�����[�~���b)<>""))
'
Private Sub DefineRelayClassNotification(sName As String)
   
    Dim �����[�敪 As String
    �����[�敪 = GetRange("�����[�敪").Rows(1).Address(RowAbsolute:=False)
    Dim �����[��� As String
    �����[��� = GetRange("�����[���").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�� As String
    �����[�� = GetRange("�����[��").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�b As String
    �����[�b = GetRange("�����[�b").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�~���b As String
    �����[�~���b = GetRange("�����[�~���b").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & �����[�敪 & ")="""",OR(TRIM(" & �����[��� & ")<>"""",TRIM(" & �����[�� & ")<>"""",TRIM(" & �����[�b & ")<>"""",TRIM(" & �����[�~���b & ")<>""""))"
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
' �����[��ڂ̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'   =AND(TRIM(�����[���)="",OR(TRIM(�����[��)<>"",TRIM(�����[�b)<>"",TRIM(�����[�~���b)<>""))
'
Private Sub DefineRelayStyleNotification(sName As String)
    
    Dim �����[��� As String
    �����[��� = GetRange("�����[���").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�� As String
    �����[�� = GetRange("�����[��").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�b As String
    �����[�b = GetRange("�����[�b").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�~���b As String
    �����[�~���b = GetRange("�����[�~���b").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & �����[��� & ")="""",OR(TRIM(" & �����[�� & ")<>"""",TRIM(" & �����[�b & ")<>"""",TRIM(" & �����[�~���b & ")<>""""))"
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
' �����[�b�̒��ӕ\����`
'
' sName             IN      �͈̖͂��O
'
'   =AND(TRIM(�����[���)="",OR(TRIM(�����[�b)<>""))
'
Private Sub DefineRelaySecondNotification(sName As String)
    
    Dim �����[��� As String
    �����[��� = GetRange("�����[���").Rows(1).Address(RowAbsolute:=False)
    Dim �����[�b As String
    �����[�b = GetRange("�����[�b").Rows(1).Address(RowAbsolute:=False)
    
    With Range(sName)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=AND(TRIM(" & �����[��� & ")="""",OR(TRIM(" & �����[�b & ")<>""""))"
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
' ����͈͂�ݒ肷��
'
Private Sub ����͈͂̐ݒ�()
    Seehts("�L���[").Select
    
    Application.PrintCommunication = True
    If Range("��").Value = �I�茠��� Then
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
