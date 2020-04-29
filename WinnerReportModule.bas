Attribute VB_Name = "WinnerReportModule"
Option Explicit    ''���ϐ��̐錾����������

Public Sub �D���҈ꗗ�쐬()

    ' �C�x���g������}��
    Call EventChange(False)

    Dim sGameName As String
    sGameName = GetRange("��").Value

    Dim oWinnerList As Object
    Set oWinnerList = CreateObject("Scripting.Dictionary")

    ' �D���҂̓Ǎ���
    Call ReadWinner(sGameName, oWinnerList)
    
    ' �D���҂̏�����
    Call WriteWinner(sGameName, oWinnerList)

    ' �C�x���g�����𔭐�
    Call EventChange(True)

    ' �ۑ�
    ActiveWorkbook.Save
End Sub

'
' �D���ғǍ���
'
' sGameName     IN  ��
' oWinnerList   OUT �D���҃��X�g
'
' oWinnerList
' ��
' ������ڔԍ�&�敪
' �@�@��
' �@�@�����C���f�b�N�X�F
' �@�@�@�@��
' �@�@�@�@��������
' �@�@�@�@��
' �@�@�@�@��������
' �@�@�@�@��
' �@�@�@�@�����L�^
' �@�@�@�@��
' �@�@�@�@�������V
'
'
Private Sub ReadWinner(sGameName As String, oWinnerList As Object)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' �v���O�����ԍ���
    Dim vProNo As Range
    For Each vProNo In GetAreaKeyData(sMasterName)

        ' �����i�^�C�������j�̏ꍇ
        If Not IsSenshukenQualifyRace(sGameName, vProNo) Then
            ' �v���O�����ԍ�������P�ʂ�T��
            Dim vCell As Range
            For Each vCell In GetRange("�v���O�����ԍ�" & CStr(vProNo))
                ' �P�ʂ̏ꍇ
                If GetOffset(vCell, Range("Header����").Column).Value = 1 Then
                    ' �D���ҏ���o�^
                    Call SetWinnerInfo(sGameName, vProNo, oWinnerList, vCell)
                End If
            Next vCell
        End If
    Next vProNo
End Sub

'
' �I�茠�̗\�I��ڂ�����
'
' True: �\�I�^ False: �����܂��͑I�茠�ȊO
'
' sGameName     IN          ��
' vProNo        IN          ProNo
'
Private Function IsSenshukenQualifyRace(sGameName As String, vProNo As Variant) As Boolean
    ' �I�茠�̗\�I�͔�΂�
    If sGameName = �I�茠��� Then
         If VLookupArea(vProNo, "�I�茠��ڋ敪", "�\�I�^����") = "�\�I" Then
            IsSenshukenQualifyRace = True
         Else
            IsSenshukenQualifyRace = False
         End If
    Else
        IsSenshukenQualifyRace = False
    End If
End Function

'
' �D���҂̑I����o�^
'
' sGameName     IN          ��
' vProNo        IN          ProNo
' oWinnerList   IN/OUT      �D���Ҕz��
' vCell         IN          �D���҂̃Z��
'
Private Sub SetWinnerInfo(sGameName As String, vProNo As Variant, oWinnerList As Object, vCell As Range)
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' �P�ʃ��X�g
    Dim oWinners As Object
    Dim oWinner As Object
    Set oWinner = CreateObject("Scripting.Dictionary")
    
    ' ���L�^���󗓂̏ꍇ����̂�Variant�Ő錾
    Dim nRecord As Variant
    nRecord = GetOffset(vCell, Range("Header���L�^").Column).Value
    
    Dim sKey As String
    sKey = GetWinnerKey(sGameName, sMasterName, CInt(vProNo), _
        GetOffset(vCell, Range("Header�敪").Column).Value)
    
    oWinner.Add "����", GetOffset(vCell, Range("Header����").Column).Value
    oWinner.Add "����", GetOffset(vCell, Range("Header����").Column).Value
    oWinner.Add "�L�^", GetOffset(vCell, Range("Header����").Column).Value
    
    If Not IsNumeric(nRecord) Or oWinner.Item("�L�^") <= nRecord Then
        oWinner.Add "���V", "���V"
    End If
    
    ' �v��No�{�敪�̂P�ʂ����o�^�̏ꍇ
    If Not (oWinnerList.Exists(sKey)) Then
        Set oWinners = CreateObject("Scripting.Dictionary")
        oWinners.Add oWinners.Count + 1, oWinner
        oWinnerList.Add sKey, oWinners
    Else
        Set oWinners = oWinnerList.Item(sKey)
        oWinners.Add oWinners.Count + 1, oWinner
    End If
End Sub

'
' ���̑��L�^�̋敪
'
' sGameName     IN  ��
' sMasterName   IN  �}�X�^�[��
' nProNo        IN  ��ڔԍ�
' sType         IN  �敪
'
Private Function GetWinnerKey(sGameName As String, sMasterName As String, nProNo As Integer, sType As String)

    If sGameName = �I�茠��� Then
        GetWinnerKey = CStr(nProNo)
    ElseIf sGameName = �s����� Then
        GetWinnerKey = CStr(nProNo) & sType
    Else
        ' �敪���擾
        If Trim(VLookupArea(nProNo, sMasterName, "��ڋ敪")) = "" Then
            GetWinnerKey = CStr(nProNo) & sType
        Else
            GetWinnerKey = CStr(nProNo)
        End If
    End If

End Function

'
' �D���ҏ�����
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Private Sub WriteWinner(sGameName As String, oWinnerList As Object)

    ' �D���҃V�[�g��I�����ی������
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' �D���Ҕ͈͖�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)

    ' �폜
    Call DeleteWinnerSheet(sWinnerAreaName)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
        
    ' �o�͂���
    Call WriteWinnerList(sWinnerAreaName, sRecordAreaName, oWinnerList)

    ' �����ݒ�
    Call SetWinnerRecordStyle(sGameName)
    
    ' ����͈͂̐ݒ�
    ActiveSheet.PageSetup.PrintArea = TableRangeAddress(sWinnerAreaName)

End Sub

'
' �D���ҏ�����
'
' sWinnerAreaName   IN  �D���Ҕ͈͖�
' sRecordAreaName   IN  ���L�^�͈͖�
' oWinnerList       IN  �D���҃��X�g
'
Private Sub WriteWinnerList(sWinnerAreaName As String, sRecordAreaName As String, oWinnerList As Object)
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
        
    ' ���L�^��
    Dim vKey As Range
    For Each vKey In GetAreaKeyData(sRecordAreaName)
        If oWinnerList.Exists(CStr(vKey.Value)) Then
            Set oWinners = oWinnerList.Item(CStr(vKey.Value))
            Dim vIdx As Variant
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteWinnerLine(sWinnerAreaName, sRecordAreaName, nRow, vKey, oWinner)
                nRow = nRow + 1
            Next vIdx
        End If
    Next vKey
End Sub


'
' �D���҃V�[�g������
'
Public Sub �D���҃V�[�g������()
    Call DeleteWinnerSheet(GetWinnerSheetName(GetRange("��").Value))
End Sub

'
' �D���҃V�[�g������
'
' sWinnerAreaName   IN  �D���Ҕ͈͖�
' nRow              IN  �D���҃V�[�g�������s
'
Private Sub DeleteWinnerSheet(sWinnerAreaName As String, Optional nRow As Integer = 1)
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    If Cells(oRange.Row + nRow, oRange.Column) <> "" Then
        oRange.Offset(nRow, 0).Resize(oRange.Rows().Count - nRow).EntireRow.Delete
    End If
End Sub

'
' �D���҃V�[�g�L��
'
' sWinnerAreaName   IN  �D���Ҕ͈͖�
' sRecordAreaName   IN  ���L�^�͈͖�
' nRow              IN  �D���҂̍s��
' vKey              IN  ���L�^�̎Q�ƌ��̊�Z��
' oWinner           IN  �D���ҏ��
'
Private Sub WriteWinnerLine(sWinnerAreaName As String, sRecordAreaName As String, nRow As Integer, vKey As Variant, oWinner As Object)
    
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "�v��No.")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "�v��No."))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "����")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "����"))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "����")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "����"))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "���")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "���"))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "�敪")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "�敪"))
    
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "����")) = oWinner.Item("����")
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "����")) = oWinner.Item("����")
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "�L�^")) = oWinner.Item("�L�^")
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "���V")) = oWinner.Item("���V")

End Sub

'
' �D���҃V�[�g�̏�������L�^����R�s�[����
'
' sGameName         IN  ��
'
Private Sub SetWinnerRecordStyle(sGameName As String)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    
     ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    GetRange(sRecordAreaName).Offset(1, 1).Resize(1).Copy
        
    ' �D���҃V�[�g
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    
    ' �D���Ҕ͈͖�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    If Cells(oRange.Row + 1, oRange.Column) <> "" Then
        oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, oRange.Columns.Count).Select
    End If
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Range("A1").Select
End Sub

'
' ���L�^�V�[�g�̏������Q�s�ڂ���R�s�[����
'
' sGameName         IN  ��
'
Private Sub SetRecordWinnerStyle(sGameName As String, sSheetName As String, sAreaName As String)

    Dim nFormatRow As Integer
    nFormatRow = 1
    Dim nStartRow As Integer
    nStartRow = 2

    ' �V�[�g���A�N�e�B�u��
    Sheets(sSheetName).Activate
    
    ' �͈͖���2�s�ځi�f�[�^�P�s�ڂ�I���j
    GetRange(sAreaName).Offset(nFormatRow, 0).Resize(1).Copy
    
    Dim oRange As Range
    Set oRange = TableRange(sAreaName)
    If oRange.Offset(nStartRow).Resize(1, 1).Value <> "" Then
        oRange.Offset(nStartRow, 0).Resize(oRange.Rows.Count - nStartRow).Select
    End If
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Range("A1").Select
End Sub

'
' �D���҃V�[�g��
'
' sGameName     IN      ��
'
Private Function GetWinnerSheetName(sGameName As String) As String
    GetWinnerSheetName = VLookupArea(sGameName, "�ݒ�e��", "�D���҃V�[�g��")
End Function

'
' �D���Ҕ͈͖�
'
' sGameName     IN      ��
'
Private Function GetWinnerAreaName(sGameName As String) As String
    GetWinnerAreaName = VLookupArea(sGameName, "�ݒ�e��", "�D���Ҕ͈͖�")
End Function

'
' ���L�^�V�[�g��
'
' sGameName     IN      ��
'
Private Function GetRecordSheetName(sGameName As String) As String
    GetRecordSheetName = VLookupArea(sGameName, "�ݒ�e��", "���L�^�V�[�g��")
End Function


'
' ���L�^�͈͖�
'
' sGameName     IN      ��
'
Private Function GetRecordAreaName(sGameName As String) As String
    GetRecordAreaName = VLookupArea(sGameName, "�ݒ�e��", "���L�^�͈͖�")
End Function

'
' ���L�^�X�V
'
Public Sub ���L�^�X�V()

    ' �C�x���g������}��
    Call EventChange(False)

    Dim sGameName As String
    sGameName = GetRange("��").Value

    Dim oWinnerList As Object
    Set oWinnerList = CreateObject("Scripting.Dictionary")

    ' ���L�^�̓Ǎ���
    Call ReadRecords(sGameName, oWinnerList)

    ' �V���L�^�̓Ǎ���
    Call ReadNewRecords(sGameName, oWinnerList)
    
    ' ���L�^�̏�����
    Call WriteNewRecords(sGameName, oWinnerList)

    Call �e��ݒ薼�O��`(�ݒ�e��V�[�g)

    ' �C�x���g�����𔭐�
    Call EventChange(True)

    ' �ۑ�
    ActiveWorkbook.Save

    ' ���L�^�V�[�g
    Call SheetActivate(GetRecordSheetName(sGameName))

End Sub

'
' ���L�^�Ǎ���
'
' sGameName     IN  ��
' oRecordList   OUT ���L�^�҃��X�g
'
' oRecordList
' ��
' ������ڔԍ�&�敪
' �@�@��
' �@�@�����C���f�b�N�X�F
' �@�@�@�@��
' �@�@�@�@��������
' �@�@�@�@��
' �@�@�@�@��������
' �@�@�@�@��
' �@�@�@�@�����L�^
' �@�@�@�@��
' �@�@�@�@�������V
'
'
Private Sub ReadRecords(sGameName As String, oRecordList As Object)
    
    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim sKey As String
        
    ' ���L�^��
    Dim vCell As Range
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        sKey = vCell.Value
        
        Set oWinner = CreateObject("Scripting.Dictionary")
        
        ' �J�����̒l�̓o�^
        Dim vKey As Range
        For Each vKey In GetRange(sRecordAreaName).Rows(1).Columns()
            oWinner.Add STrimAll(vKey.Value), GetOffset(vCell, vKey.Column).Value
        Next vKey
        
        ' �v��No�{�敪�̂P�ʂ����o�^�̏ꍇ
        If Not (oRecordList.Exists(sKey)) Then
            Set oWinners = CreateObject("Scripting.Dictionary")
            oWinners.Add oWinners.Count + 1, oWinner
            oRecordList.Add sKey, oWinners
        Else
            Set oWinners = oRecordList.Item(sKey)
            oWinners.Add oWinners.Count + 1, oWinner
        End If
    Next vCell
End Sub

'
' ���L�^�Ǎ���
'
' sGameName     IN  ��
' oRecordList   OUT ���L�^�҃��X�g
'
' oRecordList
' ��
' ������ڔԍ�&�敪
' �@�@��
' �@�@�����C���f�b�N�X�F
' �@�@�@�@��
' �@�@�@�@��������
' �@�@�@�@��
' �@�@�@�@��������
' �@�@�@�@��
' �@�@�@�@�����L�^
' �@�@�@�@��
' �@�@�@�@�������V
'
'
Private Sub ReadNewRecords(sGameName As String, oRecordList As Object)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' �D���҃V�[�g
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select

    ' �D���҂͈̔�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    
    ' ���L�^�͈̔�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    ' ���N
    Dim nYear As Integer
    nYear = GetRange("���N").Value
    
    Dim oWinners As Object
    Dim oWinner As Object
    Dim oWinnerOld As Object

    Dim sKey As String

    ' �D���Җ�
    Dim vCell As Range
    For Each vCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' ���V�Ȃ�i�[����
        If GetOffset(vCell, GetAreaColumnIndex(sWinnerAreaName, "���V")).Value = "���V" Then
                
            ' �L�[���擾
            sKey = GetWinnerKey(sGameName, sMasterName, CInt(vCell.Value), _
                    GetOffset(vCell, GetAreaColumnIndex(sWinnerAreaName, "�敪")))
                
            Set oWinner = CreateObject("Scripting.Dictionary")
            
            ' �J�����̒l�̓o�^
            Dim vKey As Variant
            For Each vKey In GetRange(sWinnerAreaName).Rows(1).Columns()
                oWinner.Add STrimAll(vKey.Value), GetOffset(vCell, vKey.Column).Value
            Next vKey
            oWinner.Add GetAreaKeyName(sRecordAreaName), sKey
            oWinner.Add "�N", nYear
                
            ' �v��No�{�敪�̂P�ʂ����o�^�̏ꍇ
            If Not (oRecordList.Exists(sKey)) Then
                Set oWinners = CreateObject("Scripting.Dictionary")
                oWinners.Add oWinners.Count + 1, oWinner
                oRecordList.Add sKey, oWinners
            Else
                Set oWinners = oRecordList.Item(sKey)
                ' ���ɑ��݂���ꍇ�̓^�C�����r���Â���΍폜
                For Each vKey In oWinners.Keys()
                    Set oWinnerOld = oWinners.Item(vKey)
                    If oWinner.Item("�L�^") < oWinnerOld.Item("�L�^") Then
                        oWinners.Remove vKey
                    End If
                Next vKey
                
                ' �ǉ�����
                oWinners.Add oWinners.Count + 1, oWinner
            End If
        End If
    Next vCell
End Sub

'
' ���L�^������
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Private Sub WriteNewRecords(sGameName As String, oRecordList As Object)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    ' �폜
    Call DeleteWinnerSheet(sRecordAreaName, 2)

    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
    
    ' ���L�^��
    Dim vKey As Variant
    For Each vKey In oRecordList.Keys()
        Set oWinners = oRecordList.Item(vKey)
        Dim vIdx As Variant
        For Each vIdx In oWinners
            Set oWinner = oWinners.Item(vIdx)
            
            ' �J�����̒l�̏�����
            Dim vCell As Range
            For Each vCell In GetRange(sRecordAreaName).Rows(1).Columns()
                Cells(nRow, vCell.Column) = oWinner.Item(STrimAll(vCell.Value))
            Next vCell
            
            nRow = nRow + 1
        Next vIdx
    Next vKey

    ' �����ݒ�
    Call SetRecordWinnerStyle(sGameName, sSheetName, sRecordAreaName)

    ' �V�[�g�̕ی�
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = True
End Sub

'
' ���L�^������
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Private Sub WriteNewRecordsOld(sGameName As String, oRecordList As Object)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    ' ���L�^��
    Dim vCell As Range
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        If oRecordList.Exists(vCell.Value) Then
            Set oWinners = oRecordList.Item(vCell.Value)
            Dim vIdx As Variant
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteRecordLine(sRecordAreaName, vCell, oWinner)
            Next vIdx
        End If
    Next vCell

    ' �V�[�g�̕ی�
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = True
End Sub


'
' ���L�^�V�[�g�L��
'
' sGameName     IN  ��
' vCell         IN  �Q�ƌ��̊�Z��
' oWinner       IN  �D���ҏ��
'
Private Sub WriteRecordLine(sAreaName As String, vCell As Variant, oWinner As Object)

    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "����") - 1).Value = oWinner.Item("����")
    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "����") - 1).Value = oWinner.Item("����")
    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "�L�^") - 1).Value = oWinner.Item("�L�^")
    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "�N") - 1).Value = oWinner.Item("�N")

End Sub
