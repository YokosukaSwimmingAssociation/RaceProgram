Attribute VB_Name = "MakeProgramModule"
Option Explicit    ''���ϐ��̐錾����������

'
' �v���O�����쐬
'
Sub �v���O�����쐬()
    ' �C�x���g������}��
    Call EventChange(False)

    ' �J�����g���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook

    ' �G���g���[�ꗗ�V�[�g
    Dim oEntrySheet As Worksheet
    Set oEntrySheet = SheetActivate(�G���g���[�V�[�g)
    
    ' �v���O�����V�[�g���쐬�i�w�b�_�s�܂Łj
    Call MakeSheet(oWorkBook, �v���O�����V�[�g)
    Dim oProgramSheet As Worksheet
    Set oProgramSheet = ActiveSheet
    
    ' �^�C�g���C��
    Call SetTitleMenu("�G���g���[�ꗗ�Ǎ���")
    
    ' �G���g���[�ꗗ�ǂݍ���
    Dim oEntryList As Object
    Call ReadEntrySheet(�G���g���[�e�[�u��, oEntryList)

    ' �\�I�Ȃ������̐ݒ�
    If GetRange("��").Value = �I�茠��� Then
        Call CheckFinal(oEntryList)
    End If

    ' �v���O�����쐬
    Call MakeProgram(oProgramSheet, �G���g���[�e�[�u��, oEntryList)

    ' �^�C�g���C��
    Call SetTitleMenu("�v���O�������O��`���s��")

    ' �v���O�����̖��O�ݒ�
    Call SetProgramName(oProgramSheet)
    
    ' �v���O�����̈���G���A�ݒ�
    Call SetPrintArea(oProgramSheet)

    ' �C�x���g�������ĊJ
    Call EventChange(True)
    
    ' �V�[�g��ۑ�
    oWorkBook.Save
    
    ' �^�C�g���C��
    Call SetTitleMenu("")

End Sub

'
' �G���g���[�ꗗ�ǂݍ���
'
' sTableName    IN      �e�[�u����
' oEntryList    OUT     �G���g���[�ꗗ(Dictionary)
' ���v��No
' �@���g
' �@�@�����[�� = ProNo��̃Z��
'
Public Sub ReadEntrySheet(sTableName As String, oEntryList As Object)

    ' �o�͗p�G���g���[�ꗗ
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    Dim oProNo As Object    ' �v��No
    Dim oHeats As Object    ' �g
    Dim nHeat  As Integer   ' �g
    Dim nLane  As Integer   ' ���[��
    
    ' �v���O����No���ɓǂݍ���
    Dim vProNo As Variant
    For Each vProNo In Range(sTableName & "[�v��No]")
        If Not IsNumeric(vProNo.Value) Then
            MsgBox CStr(vProNo.Row) & "�s�ڂɕs���Ȓl�����݂��܂��B", vbOKOnly
            End
        ElseIf Not oEntryList.Exists(vProNo.Value) Then
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vProNo.Value, oProNo
        End If
        
        ' �s�ԍ�
        nHeat = GetOffset(vProNo, Range(sTableName & "[�g]").Column).Value
        nLane = GetOffset(vProNo, Range(sTableName & "[���[��]").Column).Value
        
        ' �g���ɓǂݍ���
        If Not oProNo.Exists(nHeat) Then
            Set oHeats = CreateObject("Scripting.Dictionary")
            oProNo.Add nHeat, oHeats
        End If
        
        ' ���[���d���`�F�b�N
        If oHeats.Exists(nLane) Then
            MsgBox "�v��No�F" & CStr(vProNo.Value) & vbCrLf & _
                    "�g�@�@�F" & CStr(nHeat) & vbCrLf & _
                    "���[���F" & CStr(nLane) & vbCrLf & _
                    "���d�����Ă��܂��B"
            Range(sTableName).Parent.Activate
            Range(GetOffset(vProNo, Range(sTableName & "[���[�XNo]").Column), _
                    GetOffset(vProNo, Range(sTableName & "[���[��]").Column)).Select
            vProNo.Activate
            End
        End If
        ' ���[���o�^
        oHeats.Add nLane, vProNo
    Next vProNo

    ' �P�����o�^����ĂȂ���ΏI��
    If oEntryList.Count = 0 Then
        MsgBox "�G���g���[�ꗗ�����݂��܂���B", vbOKOnly
        End
    End If

End Sub

'
' �\�I�����m�F�i�I�茠�p�j
'
' �\�I���P�g�����Ȃ��ꍇ��
'
' oEntryList    OUT     �G���g���[�ꗗ(Dictionary)
'
Private Sub CheckFinal(oEntryList As Object)

    Dim oProNo As Object
    Dim nFinalNo As Integer
    Dim oHeats As Object
    
    ' �v���O�����ԍ���
    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData("�I�茠��ڋ敪")
        ' �\���݂̂���v��No
        If oEntryList.Exists(vProNo.Value) Then
            Set oProNo = oEntryList.Item(vProNo.Value)
            
            ' �����ԍ����擾
            nFinalNo = VLookupArea(vProNo.Value, "�I�茠��ڋ敪", "�����ԍ�")
            
            ' �\�I�̏ꍇ
            If vProNo.Value <> nFinalNo Then
            
                ' �P�g�����Ȃ��ꍇ
                If oProNo.Count = 1 Then
                    ' ���ڌ����ɂ���
                    oEntryList.Add nFinalNo, oProNo
                    ' �\�I�ɂ͗\�I�L�[�Ɍ�����������L��
                    oEntryList.Remove vProNo.Value
                    Set oProNo = CreateObject("Scripting.Dictionary")
                    oEntryList.Add vProNo.Value, oProNo
                    oProNo.Add "�\�I", "�\�I�Ȃ�-->������ No." & CStr(nFinalNo)
                ' �\�I������ꍇ
                Else
                    ' �����L�[�ɑ��L�^�A�W���L�^��o�^
                    Set oProNo = CreateObject("Scripting.Dictionary")
                    oEntryList.Add nFinalNo, oProNo
                    
                    ' �����L�[�ɋ�̑g�����Ă���
                    Set oHeats = CreateObject("Scripting.Dictionary")
                    oHeats.Add "����", vProNo.Value
                    oProNo.Add "����", oHeats
                End If
            End If
        End If
    Next vProNo

End Sub

'
' �v���O�����V�[�g������
'
Public Sub �v���O����������()
    Call DeleteProgramSheet(�v���O�����V�[�g)
    Call �v���O�������O��`
End Sub

'
' �v���O�����V�[�g��������
'
' sSheetName    OUT     �V�[�g��
'
Private Sub DeleteProgramSheet(sSheetName As String)

    If IsSheetExists(sSheetName) Then
        ' �V�[�g�����݂���ꍇ�͓��e�����ׂč폜
        Sheets(sSheetName).Activate
        Cells.Select
        Selection.Delete Shift:=xlUp
    End If

End Sub

'
' �v���O�����V�[�g���쐬
'
' oWorkBook     IN      ���[�N�V�[�g
' sSheetName    OUT     �V�[�g��
'
Private Sub MakeSheet(oWorkBook As Workbook, sSheetName As String)

    If IsSheetExists(sSheetName) Then
        ' �V�[�g�����݂���ꍇ�͓��e�����ׂč폜
        Call DeleteProgramSheet(sSheetName)
    Else
        ' ���݂��Ȃ��ꍇ�͍쐬����
        oWorkBook.Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = sSheetName
    End If
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' �w�b�_�s�쐬
    Call CopyHeaderCell(oWorkSheet, "Header�ʔ�")
    Call CopyHeaderCell(oWorkSheet, "Header�v��No")
    Call CopyHeaderCell(oWorkSheet, "Header�g")
    Call CopyHeaderCell(oWorkSheet, "Header���[��")
    Call CopyHeaderCell(oWorkSheet, "Header����")
    Call CopyHeaderCell(oWorkSheet, "Header���")
    Call CopyHeaderCell(oWorkSheet, "Header�����O")
    Call CopyHeaderCell(oWorkSheet, "Header����")
    Call CopyHeaderCell(oWorkSheet, "Header������")
    Call CopyHeaderCell(oWorkSheet, "Header�敪")
    Call CopyHeaderCell(oWorkSheet, "Header����")
    Call CopyHeaderCell(oWorkSheet, "Header����")
    'Call CopyHeaderCell(oWorkSheet, "Header����")
    Call CopyHeaderCell(oWorkSheet, "Header���l")
    Call CopyHeaderCell(oWorkSheet, "Header���L�^")
    Call CopyHeaderCell(oWorkSheet, "Header�\���݋L�^")
    Call CopyHeaderCell(oWorkSheet, "Header���[�XNo")
    Call CopyHeaderCell(oWorkSheet, "Header�\�[�g�敪")

    If GetRange("��").Value = �I�茠��� Then
        Call CopyHeaderCell(oWorkSheet, "Header�W���L�^")
    End If

End Sub

' �w�b�_�[�Z�����R�s�[
'
' �l�A�\���`���A�c���A�����A�c�ʒu�A���ʒu��ݒ�
'
' Worksheet     IN      ���[�N�V�[�g
' sCellName     IN      �Z���̖��O
'
Private Sub CopyHeaderCell(oWorkSheet As Worksheet, sCellName As String)

    Dim oRange As Range
    Set oRange = GetRange(sCellName)
    With oWorkSheet.Cells(1, oRange.Column)
        .NumberFormatLocal = oRange.NumberFormatLocal
        .ColumnWidth = oRange.ColumnWidth
        .RowHeight = oRange.RowHeight
        .HorizontalAlignment = oRange.HorizontalAlignment
        .VerticalAlignment = oRange.VerticalAlignment
        .Value = oRange.Value
    End With
End Sub

'
' �v���O�����쐬
'
' oWorkSheet    IN      �v���O�����V�[�g
' sTableName    IN      �e�[�u����
' oEntryList    IN      �G���g���[�ꗗ
'
Private Sub MakeProgram(oWorkSheet As Worksheet, sTableName As String, oEntryList As Object)

    oWorkSheet.Activate

    Dim nCurrentRow As Integer
    nCurrentRow = 1

    ' �w�b�_�s���쐬
    Call SetNoCell(oWorkSheet, nCurrentRow)

    Dim oProNo As Object
    Dim oHeats As Object
    
    Dim nMaxProNo As Integer
    Dim nMaxHeat As Integer
    Dim nRaceNo As Integer
    nMaxProNo = GetRange(GetMaster(GetRange("��").Value)).Columns(1).Rows().Count
    
    Dim sMessage As String
    
    ' �v���O�����ԍ���
    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData(GetMaster(GetRange("��").Value))
        If oEntryList.Exists(vProNo.Value) Then
            ' �\���݂̂���v��No
            Set oProNo = oEntryList.Item(vProNo.Value)
            nMaxHeat = oProNo.Count
        Else
            ' �\���݂̂Ȃ��v��No
            Set oProNo = Nothing
            nMaxHeat = 1
        End If
        
        ' �v���O�����w�b�_�쐬
        Call SetNoCell(oWorkSheet, nCurrentRow)
        Call MakeProgramHeader(oWorkSheet, sTableName, nCurrentRow, CInt(vProNo))
        
        ' �g�ԍ���
        Dim nHeat As Integer
        For nHeat = 1 To nMaxHeat
            sMessage = ""
            If oProNo Is Nothing Then
                ' �\���݂̂Ȃ��v��No�̏ꍇ�͋�̂P�g�ڂ��o��
                Set oHeats = Nothing
            ElseIf oProNo.Exists(nHeat) Then
                ' �g�����݂���ꍇ�͑g�̒l���o��
                Set oHeats = oProNo.Item(nHeat)
            ElseIf nHeat = 1 Then
                If oProNo.Exists("�\�I") Then
                ' �I�茠�̗\�I�Ȃ��̏ꍇ�͌����ւ̃��b�Z�[�W���o��
                    sMessage = oProNo.Item("�\�I")
                ' �I�茠�̗\�I�̂��錈���̏ꍇ�͑��L�^�A���[�X�ԍ�������
                ElseIf oProNo.Exists("����") Then
                    Set oHeats = oProNo.Item("����")
                    nRaceNo = nRaceNo + 1
                End If
            Else
                ' �g�����݂��Ȃ��ꍇ�i�ُ�n�j
                Set oHeats = Nothing
            End If

            ' �g�w�b�_�쐬
            'Call CopyFormat(nCurrentRow, "Prog�g�t�H�[�}�b�g")
            Call SetNoCell(oWorkSheet, nCurrentRow)
            Call MakeHeatHeader(oWorkSheet, sTableName, nCurrentRow, CInt(nHeat))
            
            ' �^�C�g���C��
            Call SetTitleMenu("�v���O�����쐬��: " & CStr(vProNo) & "/" & CStr(nMaxProNo))

            If sMessage <> "" Then
                ' ���ڌ�����
                Call SetNoCell(oWorkSheet, nCurrentRow)
                Call SetNoCell(oWorkSheet, nCurrentRow)
                Call CopyCell(oWorkSheet, nCurrentRow, "Header�v��No", vProNo)
                Cells(nCurrentRow, GetRange("Header����").Column).Value = sMessage
            Else
                ' ���[����
                Dim nLane As Integer
                For nLane = �ŏ����[���ԍ� To �ő僌�[���ԍ�
                    Call SetNoCell(oWorkSheet, nCurrentRow)
                    
                    If oHeats Is Nothing Then
                        ' �\���݂̂Ȃ�ProNo�̏ꍇ�̓f�t�H���g�\��
                        Call MakeHeatDefault(oWorkSheet, nCurrentRow, CInt(vProNo), nHeat, nLane)
                    ElseIf oHeats.Exists("����") Then
                        ' �I�茠�̌����̏ꍇ�͑��L�^�A�W���L�^�A���[�X�ԍ���ǉ�
                        Call MakeHeatDefault(oWorkSheet, nCurrentRow, CInt(vProNo), nHeat, nLane, CStr(nRaceNo))
                    ElseIf oHeats.Exists(nLane) Then
                        ' �\���݂̂���ProNo�ŃG���g�������郌�[���̏ꍇ�̓f�[�^���L�q
                        Call MakeHeat(oWorkSheet, sTableName, nCurrentRow, oHeats.Item(nLane).Row, CInt(vProNo), nHeat)
                    Else
                        ' �\���݂̂���ProNo�ŃG���g�����Ȃ����[���̏ꍇ�̓f�t�H���g�\��
                        Call MakeHeatDefault(oWorkSheet, nCurrentRow, CInt(vProNo), nHeat, nLane)
                    End If
                
                    ' ���[�X�ԍ����L�^���Ă���
                    If Cells(nCurrentRow, GetRange("Header���[�XNo").Column).Value <> "" Then
                        nRaceNo = Cells(nCurrentRow, GetRange("Header���[�XNo").Column).Value
                    End If
                Next
            End If
            ' ��s���Q�s�����
            Call SetNoCell(oWorkSheet, nCurrentRow)
            Call SetNoCell(oWorkSheet, nCurrentRow)
        Next nHeat
    Next vProNo
    
    ' �^�C�g���C��
     Call SetTitleMenu("�v���O�����슮��: " & CStr(nMaxProNo) & "/" & CStr(nMaxProNo))
End Sub

'
' �ʔԐݒ�
'
' �v���O������No�s���쐬���J�����g�s���P�C���N�������g����
'
' oWorkSheet    IN      �v���O�����V�[�g
' nCurrentRow   IN      �ʔ�
'
Private Sub SetNoCell(oWorkSheet As Worksheet, nCurrentRow As Integer)
    nCurrentRow = nCurrentRow + 1
    With oWorkSheet.Cells(nCurrentRow, GetRange("Header�ʔ�").Column)
        .Value = CStr(nCurrentRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'
' �����R�s�[
'
' nCurrentRow   IN      ���݂̍s��
' sRangeName    IN      �͈̖͂��O
'
Private Sub CopyFormat(nCurrentRow As Integer, sRangeName As String)

    ' �����R�s�[
    GetRange(sRangeName).Copy

    ' �������R�s�[
    Cells(nCurrentRow, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

End Sub


'
' �v���O�����w�b�_�쐬
'
' oWorkBook     IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
' nCurrentRow   IN      �J�����g�s��
' nProNo        IN      �v���O�����ԍ�
'
Private Sub MakeProgramHeader(oWorkSheet As Worksheet, sTableName As String, nCurrentRow As Integer, nProNo As Integer)

    Dim sMaster As String
    sMaster = GetMaster(GetRange("��").Value)
    
    Dim sQualifyFormat As String

    Call CopyCell(oWorkSheet, nCurrentRow, "Prog�v��No")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog��ڋ敪")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog��ږ�")

    With Range(sTableName).ListObject
        Cells(nCurrentRow, GetRange("Prog�v��No").Column).Value = nProNo
        Cells(nCurrentRow, GetRange("Prog��ڋ敪").Column).Value = _
            VLookupArea(nProNo, sMaster, "��ڋ敪") & _
            VLookupArea(nProNo, sMaster, "����")

        Cells(nCurrentRow, Range("Prog��ږ�").Column).Value = _
            VLookupArea(nProNo, sMaster, "����") & _
            VLookupArea(nProNo, sMaster, "���")
    
        ' ���{��I�茠�͕W���L�^�A���L�^���o��
        If GetRange("��").Value = �I�茠��� Then
            
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
            Cells(nCurrentRow, Range("Prog����").Column).Value = _
                VLookupArea(nProNo, sMaster, "�\�I�^����")
            
            
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog�L�^")
            Dim nFinalNo As Integer
            nFinalNo = VLookupArea(nProNo, "�I�茠��ڋ敪", "�����ԍ�")
            Dim nQualify As Long
            nQualify = VLookupArea(nProNo, sMaster, "�W���L�^")
            Dim sFormat As String
            If nQualify < 10000 Then
                sQualifyFormat = "##"".""#"
            Else
                sQualifyFormat = "0"":""##"".""#"
            End If
            Dim nRecord As Long
            nRecord = VLookupArea(nFinalNo, "�I�茠���L�^", "�L�^")
            Dim sRecordFormat As String
            If nRecord < 10000 Then
                sRecordFormat = "##"".""##"
            Else
                sRecordFormat = "0"":""##"".""##"
            End If
            Cells(nCurrentRow, Range("Prog�L�^").Column).Value = _
                "�i�W���L�^ " & Format(nQualify / 10, sQualifyFormat) & ", " & _
                "���L�^ " & Format(nRecord, sRecordFormat) & "�j"
        End If
    
    End With

    ' ����������
    With Range(Cells(nCurrentRow, Range("Header�g").Column), Cells(nCurrentRow, Range("Header���L�^").Column)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

End Sub

' �Z�����R�s�[
'
' oWorkSheet    IN      ���[�N�V�[�g
' nRow          IN      �s��
' sCellName     IN      �f�t�H���g�̃Z����
' vOverRide     IN      �R�s�[���镶����
'
Private Sub CopyCell(oWorkSheet As Worksheet, nRow As Integer, sCellName As String, Optional vOverRide As Variant = Empty)

    Dim oRange As Range
    Set oRange = GetRange(sCellName)
    With oWorkSheet.Cells(nRow, oRange.Column)
        .ShrinkToFit = oRange.ShrinkToFit
        .NumberFormatLocal = oRange.NumberFormatLocal
        .Font.Name = oRange.Font.Name
        .Font.Size = oRange.Font.Size
        .Font.Underline = oRange.Font.Underline
        .Font.Bold = oRange.Font.Bold
        .HorizontalAlignment = oRange.HorizontalAlignment
        .VerticalAlignment = oRange.VerticalAlignment
        .IndentLevel = oRange.IndentLevel
        If IsEmpty(vOverRide) Then
            .Value = Range(sCellName).Value
        Else
            .Value = CStr(vOverRide)
        End If
    End With
End Sub

'
' �g�w�b�_�쐬
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
' nCurrentRow   IN      �J�����g�s�ԍ�
' nHeat         IN      �g�ԍ�
'
Private Sub MakeHeatHeader(oWorkSheet As Worksheet, sTableName As String, nCurrentRow As Integer, nHeat As Integer)
    
    Call CopyCell(oWorkSheet, nCurrentRow, "Header�g")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header���[��")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header�����O")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header������")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header�敪")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header����")
    'Call CopyCell(oWorkSheet, nCurrentRow, "Header����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header���l")
    Call CopyCell(oWorkSheet, nCurrentRow, "Header���L�^")

    'With Range(sTableName).ListObject
        Cells(nCurrentRow, Range("Prog�g��").Column).Value = _
            "<" & Trim(CStr(nHeat)) & "�g>"
    'End With

End Sub

'
' �I�背�R�[�h�쐬
'
' oWorkSheet    IN      ���[�N�V�[�g
' nCurrentRow   IN      �J�����g�s�ԍ�
' nProNo        IN      �v���O�����ԍ�
' nHeat         IN      �g�ԍ�
' nLane         IN      ���[���ԍ�
'
Private Sub MakeHeatDefault(oWorkSheet As Worksheet, nCurrentRow As Integer, _
nProNo As Integer, nHeat As Integer, nLane As Integer, _
Optional sRaceNo As String = Empty)
    
    Call CopyCell(oWorkSheet, nCurrentRow, "Header�v��No", nProNo)
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog�g��", Format(nProNo, "0#") & "-" & Format(nHeat, "#"))
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog���[��", nLane)
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
    Range(Cells(nCurrentRow, GetRange("Prog����").Column), _
        Cells(nCurrentRow, GetRange("Prog���").Column)).Merge
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog�����O")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog������")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog�敪")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
    'Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog���l")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog���L�^")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog�\���݋L�^")
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog���[�XNo", sRaceNo)
    Call CopyCell(oWorkSheet, nCurrentRow, "Prog�W���L�^")

End Sub

'
' �g���쐬
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
' nCurrentRow   IN      �J�����g�s�ԍ�(�v���O�����V�[�g)
' nRow          IN      �J�����g�s�ԍ�(�e�[�u��)
' nProNo        IN      �v���O�����ԍ�
' nHeat         IN      �g�ԍ�
'
Private Sub MakeHeat(oWorkSheet As Worksheet, sTableName As String, nCurrentRow As Integer, _
nRow As Integer, nProNo As Integer, nHeat As Integer)

    oWorkSheet.Activate

    With Range(sTableName).ListObject
        
        Call CopyCell(oWorkSheet, nCurrentRow, "Header�v��No", nProNo)
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog�g��", _
                             Format(nProNo, "0#") & "-" & CStr(nHeat))
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog���[��", _
                            .ListColumns("���[��").Range(nRow).Value)
        
        If .ListColumns("�I�薼").Range(nRow).Value <> "" Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog����", _
                            .ListColumns("�I�薼").Range(nRow).Value)
        Else
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
        End If
        Range(Cells(nCurrentRow, GetRange("Prog����").Column), _
            Cells(nCurrentRow, GetRange("Prog���").Column)).Merge
        
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog�����O")
        If Trim(.ListColumns("�w�Z��").Range(nRow).Value) <> "" Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog����", _
                            .ListColumns("�w�Z��").Range(nRow).Value)
        Else
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog����", _
                            .ListColumns("�`�[����").Range(nRow).Value)
        End If
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog������")
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog�敪", _
                            .ListColumns("�敪").Range(nRow).Value)
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog����")
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog���l")

        ' ���{��I�茠���j���
        
        If GetRange("��").Value = �I�茠��� Then
            Dim nFinalNo As Integer
            nFinalNo = VLookupArea(.ListColumns("�v��No").Range(nRow).Value, "�I�茠��ڋ敪", "�����ԍ�")
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog���L�^", _
                    VLookupArea(nFinalNo, "�I�茠���L�^", "�L�^"))
        
        ' ���{��s���̈���
        ElseIf GetRange("��").Value = �s����� Then
            ' Call CopyCell(oWorkSheet, nCurrentRow, "Prog���L�^", _
            '        VLookupArea(GetRecordKey(�s�����, _
            '        .ListColumns("�v��No").Range(nRow).Value, _
            '        .ListColumns("�敪").Range(nRow).Value), "�s�����L�^", "�L�^"))
        
        ' �����L�^��
        ElseIf GetRange("��").Value = �����L�^�� Then
            Debug.Print "..."
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog����", _
                            .ListColumns("����").Range(nRow).Value)
            
        ' �w���}�X�^�[�Y���
        Else
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog���L�^", _
                    VLookupArea(GetRecordKey(�w�}���, _
                    .ListColumns("�v��No").Range(nRow).Value, _
                    .ListColumns("�\�[�g�敪").Range(nRow).Value), "�w�}���L�^", "�L�^"))
        End If

        
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog�\���݋L�^", _
                            .ListColumns("�\���ݎ���").Range(nRow).Value)
        
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog�\�[�g�敪", _
                            .ListColumns("�\�[�g�敪").Range(nRow).Value)
        Call CopyCell(oWorkSheet, nCurrentRow, "Prog���[�XNo", _
                            .ListColumns("���[�XNo").Range(nRow).Value)

        ' ���{��I�茠���j���͕W���L�^���L��
        If GetRange("��").Value = �I�茠��� Then
            Call CopyCell(oWorkSheet, nCurrentRow, "Prog�W���L�^", _
                    VLookupArea(.ListColumns("�v��No").Range(nRow).Value, "�I�茠��ڋ敪", "�W���L�^"))
        End If
    
    End With

End Sub

'
' �v���O�������O��`
'
' �u�v���O�����쐬�}�N���v����{�^���Ŏ��s�����
'
Public Sub �v���O�������O��`()
    Sheets(�v���O�����V�[�g).Activate
    Call SetProgramName(ActiveSheet)
End Sub

'
' �v���O�����V�[�g�̖��O��`
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetProgramName(oWorkSheet As Worksheet)
    Call DeleteName("�v���O����*")
    Call SetNoName(oWorkSheet)
    If IsNameExists("�v���O�����ʔ�") Then
        Call SetProNoName(oWorkSheet)
        Call SetProNoListName(oWorkSheet)
        Call SetHeatName(oWorkSheet)
        Call SetRaceName(oWorkSheet)
        Call SetSameRaceLabel(oWorkSheet)
    End If
End Sub

'
' �v���O�����V�[�g�̒ʔԗ�̖��O��`
'
' ���O�u�v���O�����ʔԁv���`
'
' �v���O�����V�[�g�̂Q�s��(A2)����ŉ��ʍs�܂ł�
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetNoName(oWorkSheet As Worksheet)
    oWorkSheet.Activate
    Dim oCell As Range
    Set oCell = Cells(2, GetRange("Header�ʔ�").Column)
    If oCell.Value <> "" Then
        Call DefineName("�v���O�����ʔ�", RowRangeAddress(oCell.Address))
    End If
    Call SetForcusTop
End Sub

'
' �v���O�����ԍ��ꗗ�̖��O��`
'
' ���O�u�v���O������ڔԍ��v�̒�`
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetProNoName(oWorkSheet As Worksheet)
    
    ' �v��No
    Dim nProNo As Integer
    nProNo = 1

    ' �A�h���X������i�[�p
    Dim sAddress As String
    sAddress = ""

    ' �Z���I�u�W�F�N�g
    Dim oCell As Range

    ' �v���O�����ʔԂ��V�[�N���Ȃ��珈��������
    Dim vNo As Variant
    For Each vNo In GetRange("�v���O�����ʔ�")
        Set oCell = oWorkSheet.Cells(vNo.Row, Range("Header�g").Column)
        ' �g��Ńv��No�Ɠ����ꍇ�̓v���O�����ԍ��̃Z��
        If oCell.Value = nProNo Then
            If sAddress = "" Then
                sAddress = oCell.Address(ReferenceStyle:=xlA1)
            Else
                sAddress = sAddress & "," & oCell.Address(ReferenceStyle:=xlA1)
            End If
            ' �v��No���C���N�������g
            nProNo = nProNo + 1
        End If
    Next vNo

    Call DefineName("�v���O������ڔԍ�", sAddress)

End Sub

'
' �L�^��ʌ����p�̖��O��`
'
' ���O�u�v���O�����ԍ�N�v�̒�`
'
' N�F�v���O�����ԍ�
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetProNoListName(oWorkSheet As Worksheet)
    
    ' �v��No
    Dim nProNo As Integer
    nProNo = 1
    
    ' �A�h���X������i�[�p
    Dim oRange As Range
    Set oRange = Nothing
    
    ' �Z���I�u�W�F�N�g
    Dim oCell As Range
    
    ' �v���O�����ʔԂ��V�[�N���Ȃ��珈��������
    Dim vNo As Variant
    For Each vNo In Range("�v���O�����ʔ�")
        Set oCell = oWorkSheet.Cells(vNo.Row, GetRange("Header�v��No").Column)
        ' �v��No��Ńv��No���傫���Ȃ����ꍇ�ɓo�^
        If oCell.Value > nProNo Then
            ' �A�h���X����łȂ���Ζ��O��o�^����
            If Not (oRange Is Nothing) Then
                Call DefineName("�v���O�����ԍ�" & Trim(CStr(nProNo)), oRange.Address)
                Set oRange = Nothing
                ' �v��No���C���N�������g
                nProNo = nProNo + 1
            End If
        End If
        ' �v��No��Ńv��No�Ɠ����ꍇ�̓v���O�����ԍ��̃Z��
        If oCell.Value = nProNo Then
            If oRange Is Nothing Then
                Set oRange = oCell
            Else
                Set oRange = Application.Union(oRange, oCell)
            End If
        End If
    Next vNo

    ' �A�h���X����łȂ���Ζ��O��o�^����
    If Not (oRange Is Nothing) Then
        Call DefineName("�v���O�����ԍ�" & Trim(CStr(nProNo)), oRange.Address)
    End If
End Sub

'
' �L�^��ʌ����p�̑g�̖��O��`
'
' ���O�u�v���O�����gNN-X�v�̒�`
'
' NN�F�v���O�����ԍ�
'  X�F�g��
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetHeatName(oWorkSheet As Worksheet)
   
    ' �v���O�����ԍ�
    Dim nProNo As Integer
    nProNo = 0
    
    ' ���̃v���O�����ԍ�
    Dim nNextProNo As Integer
    nNextProNo = 1
    
    ' �g�ԍ�
    Dim nHeat As Integer
    ' �g��
    Dim sHeatName As String
    
    ' �A�h���X������i�[�p
    Dim oRange As Range
    Set oRange = Nothing

    ' �Z���I�u�W�F�N�g
    Dim oCell As Range

    Dim vNo As Variant
    For Each vNo In Range("�v���O�����ʔ�")
        Set oCell = oWorkSheet.Cells(vNo.Row, GetRange("Header�g").Column)
        ' ���̃v���O�����ԍ��ɕς��ꍇ
        If oCell.Value = nNextProNo Then
            nProNo = nNextProNo         ' �v���O�����ԍ����C���N�������g
            nNextProNo = nNextProNo + 1 ' ���̃v���O�����ԍ����C���N�������g
            nHeat = 1                   ' �g�ԍ��̏�����
        End If
        ' �g���̃t�H�[�}�b�g
        sHeatName = Format(nProNo, "0#") & "-" & Trim(CStr(nHeat))
        ' �g�ƈ�v����ꍇ�͖��O�͈̔�
        If oCell.Value = sHeatName Then
            If oRange Is Nothing Then
                Set oRange = oCell
            Else
                Set oRange = Application.Union(oRange, oCell)
            End If
        End If

        ' ��s�Ŗ��O�͈͂�����ꍇ
        If oCell.Value = "" And Not (oRange Is Nothing) Then
            ' ���O���`����
            Call DefineName("�v���O�����g" & Replace(sHeatName, "-", "_"), oRange.Address)

            ' ���O�͈͂Ƒg�ԍ���������
            Set oRange = Nothing
            nHeat = nHeat + 1
        End If
    Next vNo
End Sub

'
' �L�^��ʌ����p�̖��O��`
'
' ���O�u�v���O�������[�XNN�v�̒�`
'
' NN�F���[�X�ԍ�
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetRaceName(oWorkSheet As Worksheet)
    
    Dim nRaceNo As Integer
    nRaceNo = 0
        
    ' �A�h���X������i�[�p
    Dim oRange As Range
    Set oRange = Nothing
    
    ' �Z���I�u�W�F�N�g
    Dim oCell As Range

    ' �v���O�����ʔԂ��V�[�N���Ȃ��珈��������
    Dim vNo As Variant
    For Each vNo In Range("�v���O�����ʔ�")
        Set oCell = oWorkSheet.Cells(vNo.Row, GetRange("Header���[�XNo").Column)
        ' �󔒈ȊO�̏ꍇ
        If oCell.Value <> "" Then
            If oCell.Value > nRaceNo Then
                ' �A�h���X����łȂ���Ζ��O��o�^����
                If Not (oRange Is Nothing) Then
                    Call DefineName("�v���O�������[�X" & Trim(CStr(nRaceNo)), oRange.Address)
                    Set oRange = Nothing
                End If
                nRaceNo = oCell.Value
            End If
            ' �v��No��Ńv��No�Ɠ����ꍇ�̓v���O�����ԍ��̃Z��
            If oCell.Value = nRaceNo Then
                If oRange Is Nothing Then
                    Set oRange = oCell
                Else
                    Set oRange = Application.Union(oRange, oCell)
                End If
            End If
        End If
    Next vNo

    ' �A�h���X����łȂ���Ζ��O��o�^����
    If Not (oRange Is Nothing) Then
        Call DefineName("�v���O�������[�X" & Trim(CStr(nRaceNo)), oRange.Address)
    End If

End Sub

'
' ���ꃌ�[�X���x���쐬
'
' ���ꃌ�[�X�̏ꍇ�ɁuX-X-X ���ꃌ�[�X�v�Ƃ���������ǋL����
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetSameRaceLabel(oWorkSheet As Worksheet)
    
    Dim oRaceNo As Object
    Set oRaceNo = CreateObject("Scripting.Dictionary")
    
    ' ���[�XNo�ɑ΂���v��No��Ǎ���
    Call ReadSameRace(oWorkSheet, oRaceNo)
    
    ' ���ꃌ�[�X���x����������
    Call WriteSameRaceLabel(oRaceNo)

End Sub

'
' ���[�XNo�ɑ΂���v��No��Ǎ���
'
' oWorkBook     IN      ���[�N�V�[�g
' oRaceNo       OUT     ���[�XNo�z��
'  �����[�XNo
'  �@���v��No�F1
'
Private Sub ReadSameRace(oWorkSheet As Worksheet, oRaceNo As Object)
    
    Dim oProNo As Object
    Dim nProNo As Integer
    Dim nRaceNo As Integer
    Dim vNo As Variant
    For Each vNo In GetRange("�v���O�����ʔ�")
        ' ���[�XNo���擾
        nRaceNo = oWorkSheet.Cells(vNo.Row, GetRange("Header���[�XNo").Column).Value
        If nRaceNo > 0 Then
            If Not oRaceNo.Exists(nRaceNo) Then
                ' ���[�XNo��ǉ�
                Set oProNo = CreateObject("Scripting.Dictionary")
                oRaceNo.Add nRaceNo, oProNo
            End If
            ' �v��No���擾
            nProNo = Cells(vNo.Row, Range("Header�v��No").Column).Value
            If Not oProNo.Exists(nProNo) Then
                ' �v��No��ǉ�
                oProNo.Add nProNo, 1
            End If
        
        End If
    Next vNo
End Sub

'
' ���ꃌ�[�X���x��������
'
' �L�q����ꏊ��ProNo�̂P�s�O�A�����Ɠ�����
'
' oRaceNo       IN      ���[�XNo�z��
'
Private Sub WriteSameRaceLabel(oRaceNo As Object)
    Dim oCell As Range
    Dim oProNo As Object
    
    Dim aryProNo As Variant
    Dim sLabel As String
    
    Dim vRaceNo As Variant
    For Each vRaceNo In oRaceNo
        Set oProNo = oRaceNo.Item(vRaceNo)
        If oProNo.Count > 1 Then
            aryProNo = oProNo.Keys()
            sLabel = Join(aryProNo, "-") & " ���ꃌ�[�X"
            
            Dim vProNo As Variant
            For Each vProNo In aryProNo
                Set oCell = GetProNoCell(CInt(vProNo))
                GetOffset(oCell, GetRange("Prog����").Column).Offset(-1).Value = sLabel
            Next vProNo
        
        End If
    Next vRaceNo
End Sub

'
' �v���O�����ԍ��̍s�����擾
'
' ���O�u�v���O������ڔԍ��v����v���O�����w�b�_��ProNo�Z�����擾
'
' oRaceNo       IN      ���[�XNo�z��
'
Private Function GetProNoCell(nProNo As Integer) As Range
    Dim sName As String
    sName = "�v���O������ڔԍ�"

    Dim vProNo As Variant
    For Each vProNo In GetRange(sName)
        If vProNo.Value = nProNo Then
            Set GetProNoCell = vProNo
            Exit Function
        End If
    Next vProNo
End Function

'
' ����͈͐ݒ�
'
' oWorkBook     IN      ���[�N�V�[�g
'
Private Sub SetPrintArea(oWorkSheet As Worksheet)
    oWorkSheet.Activate
    
    ' ����G���A�̃N���A
    ActiveSheet.PageSetup.PrintArea = ""
    ' ���y�[�W�̃N���A
    ActiveSheet.ResetAllPageBreaks
    
    ' ����G���A�̐ݒ�
    Dim nBottom As Integer
    nBottom = Range("$A$1").End(xlDown).Row
    
    ' �I�茠���̏ꍇ�͑��L�^��������Ȃ�
    If GetRange("��").Value = �I�茠��� Then
        ActiveSheet.PageSetup.PrintArea = _
            Range(Cells(GetRange("Header�g").Row, GetRange("Header�g").Column), _
            Cells(nBottom, GetRange("Header���l").Column)).Address
        Cells(1, GetRange("Header����").Column).ColumnWidth = 20
        Cells(1, GetRange("Header���").Column).ColumnWidth = 20
        Cells(1, GetRange("Header���l").Column).ColumnWidth = 20
    Else
        ActiveSheet.PageSetup.PrintArea = _
            Range(Cells(2, GetRange("Header�g").Column), Cells(nBottom, GetRange("Header���L�^").Column)).Address
    End If

    ' ����G���A�̐ݒ�i���P�y�[�W�j
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .FitToPagesWide = 1
        .CenterFooter = "�|&P�|"
    End With
    Application.PrintCommunication = True

    ' ���y�[�W�ݒ�
    Call SetPageBreaks
    
    ' �P�s�̍���
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 16.7
    Call SetForcusTop

End Sub

'
' ���y�[�W�ݒ�
'
Private Sub SetPageBreaks()
    ' ���y�[�W�v���r��
    ActiveWindow.View = xlPageBreakPreview
    
    ' ���y�[�W�ݒ�
    Dim nNum As Integer     ' ���[�X���J�E���^
    nNum = 0
    Dim nProNo As Integer   ' �v��No�̒l
    Dim bFlag As Boolean    ' �͈̓t���O(True�F�͈͊O�^False�F�͈͓�)
    bFlag = True
    
    Dim nRow As Integer
    Dim nBottom As Integer
    nBottom = GetAreaBottomRow("�v���O�����ʔ�")
    
    Dim vNo As Variant
    For Each vNo In GetRange("�v���O�����ʔ�")
        ' ProNo�̒l���擾
        nProNo = Val(GetOffset(vNo, GetRange("Header�v��No").Column).Value)
        ' �󗓂łȂ��ꍇ
        If nProNo > 0 Then
            ' �A���łȂ��ꍇ�̓J�E���^���グ��
            If bFlag Then
                nNum = nNum + 1
            End If
            bFlag = False
        Else
            ' �󔒂Ŕ͈͊O�ɏo���ꍇ�ibFlag=False)
            If bFlag = False And nNum Mod �y�[�W���[�X�� = 0 Then
                ' ���s�y�[�W
                nRow = vNo.Row + 1
                If nRow < nBottom Then
                    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=Cells(nRow, GetRange("Header�g").Column)
                End If
            End If
            bFlag = True
        End If
    Next vNo
    
    ' ���y�[�W�v���r����߂�
    ActiveWindow.View = xlNormalView
    Call SetForcusTop
End Sub
