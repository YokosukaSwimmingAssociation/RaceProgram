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
' �@�@�@�@�����敪
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
    sKey = GetRecordKey(sGameName, CInt(vProNo), _
        GetOffset(vCell, Range("Header�敪").Column).Value)
    
    oWinner.Add "����", GetOffset(vCell, Range("Header����").Column).Value
    oWinner.Add "����", GetOffset(vCell, Range("Header����").Column).Value
    oWinner.Add "�L�^", GetOffset(vCell, Range("Header����").Column).Value
    oWinner.Add "�敪", GetOffset(vCell, Range("Header�敪").Column).Value
    
    If Not IsNumeric(nRecord) Then
        oWinner.Add "���V", "�Q�l�L�^"
    ElseIf oWinner.Item("�L�^") <= nRecord Then
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
' nProNo        IN  ��ڔԍ�
' sClass        IN  �敪
'
Public Function GetRecordKey(sGameName As String, nProNo As Integer, sClass As String) As String

    If sGameName = �I�茠��� Then
        GetRecordKey = CStr(nProNo)
    ElseIf sGameName = �s����� Then
        GetRecordKey = Format(nProNo, "00") & "_" & Replace(STrimAll(sClass), "���", "20��")
    Else
        Dim sMasterName As String
        sMasterName = GetMaster(sGameName)
        ' �敪���擾
        If Trim(VLookupArea(nProNo, sMasterName, "��ڋ敪")) = "" Then
            GetRecordKey = CStr(nProNo) & "_" & sClass
        Else
            GetRecordKey = CStr(nProNo) & "_"
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
    
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' �D���Ҕ͈͖�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
        
    If sGameName = �I�茠��� Then
        ' ������
        Call DeleteWinnerSheetForSenshuken(sWinnerAreaName)
        
        ' �o�͂���
        Call WriteWinnerListForSenshuken(sWinnerAreaName, oWinnerList)
    
        ' �D���҃V�[�g�ݒ�
        Call �I�茠���L�^�ݒ�(sSheetName, sWinnerAreaName)
    Else
        ' ������
        Call DeleteWinnerSheet(sWinnerAreaName)
        
        ' �o�͂���
        Call WriteWinnerList(sWinnerAreaName, sRecordAreaName, oWinnerList)
        
        ' �D���҃V�[�g�ݒ�
        Call DefineWinnerSheet(sSheetName, sWinnerAreaName)
    
        ' �����ݒ�
        Call SetWinnerRecordStyle(sGameName)
    End If

    ' �V�[�g�̕ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' �D���ҏ�����
'
' sWinnerAreaName   IN  �D���Ҕ͈͖�
' sRecordAreaName   IN  ���L�^�͈͖�
' oWinnerList       IN  �D���҃��X�g
'
Private Sub WriteWinnerList(sWinnerAreaName As String, sRecordAreaName As String, oWinnerList As Object)
    
    ' ����̋敪�ɑ΂��ĂP��̂ݎ��{����`�F�b�N�p
    Dim oKeyList As Object
    Set oKeyList = CreateObject("Scripting.Dictionary")
    
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
        
    ' ���L�^��
    Dim vKey As Range
    For Each vKey In GetAreaKeyData(sRecordAreaName)
        If Not oKeyList.Exists(CStr(vKey.Value)) Then
            If oWinnerList.Exists(CStr(vKey.Value)) Then
                Set oWinners = oWinnerList.Item(CStr(vKey.Value))
                Dim vIdx As Variant
                For Each vIdx In oWinners
                    Set oWinner = oWinners.Item(vIdx)
                    Call WriteWinnerLine(sWinnerAreaName, sRecordAreaName, nRow, vKey, oWinner)
                    nRow = nRow + 1
                Next vIdx
            End If
            oKeyList.Add CStr(vKey.Value), 1
        End If
    Next vKey
End Sub

'
' �D���ҏ����݁i�I�茠�p�j
'
' sAreaName         IN  �D���Ҕ͈͖�
' oWinnerList       IN  �D���҃��X�g
'
Private Sub WriteWinnerListForSenshuken(sAreaName As String, oWinnerList As Object)
    
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")
    
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData(sAreaName)
        If vProNo.Value > 0 Then
            ' �D���҂����݂���ꍇ
            If oWinnerList.Exists(CStr(vProNo.Value)) Then
                ' ����ProNo�ɂ͂P�񂾂����{
                If Not oProNo.Exists(CStr(vProNo.Value)) Then
                    Set oWinners = oWinnerList.Item(CStr(vProNo.Value))
                    Dim vIdx As Variant
                    For Each vIdx In oWinners
                        Set oWinner = oWinners.Item(vIdx)
                        If vIdx > 1 Then
                            ' ���^�C���L�^�̏ꍇ�͍s��}��
                            Set vProNo = InsertWinnerRow(vProNo, sAreaName)
                        End If
                        Call WriteWinnerLineForSenshuken(sAreaName, vProNo, oWinner)
                    Next vIdx
                    oProNo.Add CStr(vProNo.Value), 1
                End If
            End If
        End If
    Next vProNo

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
' �D���҃V�[�g�������i�I�茠�p�j
'
' sAreaName         IN  �D���Ҕ͈͖�
'
Private Sub DeleteWinnerSheetForSenshuken(sAreaName As String)

    ' �����s����ꍇ�͍폜�ŏ��Ƃ���
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")
    Dim oDelList As Object
    Set oDelList = CreateObject("Scripting.Dictionary")

    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData(sAreaName)
        If vProNo.Value > 0 Then
            If oProNo.Exists(vProNo.Value) Then
                ' ���[�v���ɍ폜����Ɣ͈͂����������Ȃ�̂�
                ' ��U�폜�Ώۃ��X�g�ɒǉ�����
                oDelList.Add oDelList.Count + 1, vProNo
            Else
                oProNo.Add vProNo.Value, 1
                GetOffset(vProNo, GetColIdx(sAreaName, "����")) = ""
                GetOffset(vProNo, GetColIdx(sAreaName, "����")) = ""
                GetOffset(vProNo, GetColIdx(sAreaName, "�敪")) = ""
                GetOffset(vProNo, GetColIdx(sAreaName, "�L�^")) = ""
                If GetColIdx(sAreaName, "���V") > 0 Then
                    GetOffset(vProNo, GetColIdx(sAreaName, "���V")) = ""
                End If
                If GetColIdx(sAreaName, "�N") > 0 Then
                    GetOffset(vProNo, GetColIdx(sAreaName, "�N")) = ""
                End If
            End If
        End If
    Next vProNo

    ' �폜���X�g���폜����
    Dim oDelCell As Range
    Dim vDelCell As Variant
    For Each vDelCell In oDelList
        Set oDelCell = oDelList.Item(vDelCell)
        oDelCell.EntireRow.Delete
    Next vDelCell

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
Private Sub WriteWinnerLine(sWinnerAreaName As String, sRecordAreaName As String, _
nRow As Integer, vKey As Variant, oWinner As Object)
    
    Cells(nRow, GetColIdx(sWinnerAreaName, "�v��No.")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "�v��No."))
    Cells(nRow, GetColIdx(sWinnerAreaName, "����")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "����"))
    Cells(nRow, GetColIdx(sWinnerAreaName, "����")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "����"))
    Cells(nRow, GetColIdx(sWinnerAreaName, "���")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "���"))
    Cells(nRow, GetColIdx(sWinnerAreaName, "�敪")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "�敪"))
    
    Cells(nRow, GetColIdx(sWinnerAreaName, "����")) = oWinner.Item("����")
    Cells(nRow, GetColIdx(sWinnerAreaName, "����")) = oWinner.Item("����")
    Cells(nRow, GetColIdx(sWinnerAreaName, "�L�^")) = oWinner.Item("�L�^")
    Cells(nRow, GetColIdx(sWinnerAreaName, "���V")) = oWinner.Item("���V")

End Sub

'
' �D���҃V�[�g�L���i�I�茠�p�j
'
' sAreaName         IN  �D���Ҕ͈͖�
' vKey              IN  �D���҂̊�Z��
' oWinner           IN  �D���ҏ��
'
Private Sub WriteWinnerLineForSenshuken(sAreaName As String, vKey As Variant, oWinner As Object)
    
    GetOffset(vKey, GetColIdx(sAreaName, "����")) = Replace(oWinner.Item("����"), "�D", vbCrLf)
    GetOffset(vKey, GetColIdx(sAreaName, "����")) = oWinner.Item("����")
    GetOffset(vKey, GetColIdx(sAreaName, "�敪")) = oWinner.Item("�敪")
    GetOffset(vKey, GetColIdx(sAreaName, "�L�^")) = oWinner.Item("�L�^")
    
    If GetColIdx(sAreaName, "���V") > 0 Then
        GetOffset(vKey, GetColIdx(sAreaName, "���V")) = oWinner.Item("���V")
    End If
    
    If GetColIdx(sAreaName, "�N") > 0 Then
        GetOffset(vKey, GetColIdx(sAreaName, "�N")) = oWinner.Item("�N")
    End If
    

End Sub

'
' �D���҃V�[�g�s�}���i�I�茠�p�j
'
' ���^�C���̏ꍇ�ɍs��}������
'
' oCell         IN      ��̃Z��
' sAreaName         IN  �D���Ҕ͈͖�
'
Private Function InsertWinnerRow(oCell As Variant, sAreaName As String) As Range
    Dim oNewCell As Range
    oCell.Offset(1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Set oNewCell = oCell.Offset(1)
    oNewCell.Value = oCell.Value
    GetOffset(oNewCell, GetColIdx(sAreaName, "����")).Value = GetOffset(oCell, GetColIdx(sAreaName, "����")).Value
    ' ��ڗ�͌���
    Union(GetOffset(oNewCell, GetColIdx(sAreaName, "���")), GetOffset(oCell, GetColIdx(sAreaName, "���")).MergeArea).Merge
    GetOffset(oNewCell, GetColIdx(sAreaName, "����")).Value = GetOffset(oCell, GetColIdx(sAreaName, "����")).Value
    ' �S�̂Ɍr����ݒ�
    Call SetBorder(GetOffset(oNewCell, GetColIdx(sAreaName, "���")) _
        .Resize(1, GetRange(sAreaName).Columns.Count - (GetColIdx(sAreaName, "���") - oCell.Column)))
    ' �V�����s�̊�Z����Ԃ�
    Set InsertWinnerRow = oNewCell
End Function


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
' sSheetName        IN  �V�[�g��
' sAreaName         IN  �͈͖�
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
' ���L�^�V�[�g����ёւ���
'
' sSheetName        IN  �V�[�g��
' sAreaName         IN  �͈͖�
'
Private Sub SortRecordWinner(sSheetName As String, sAreaName As String)

    ' �V�[�g���A�N�e�B�u��
    Sheets(sSheetName).Activate
    Range("A1").Select
    Selection.AutoFilter
   
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:= _
            Range(RowRangeAddress("B2")), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortTextAsNumbers
        .SortFields.Add2 Key:= _
            Range(RowRangeAddress("F2")), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
            :=xlSortNormal

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1").Select
End Sub

'
' ���L�^�V�[�g����ёւ���i�s�����j
'
' sSheetName        IN  �V�[�g��
' sAreaName         IN  �͈͖�
'
Private Sub SortRecordWinnerShimin(sSheetName As String, sAreaName As String)

    ' �V�[�g���A�N�e�B�u��
    Sheets(sSheetName).Activate
    Range("A1").Select
    Selection.AutoFilter
   
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:= _
            Range(RowRangeAddress("A2")), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortNormal

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

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
    
    If sGameName = �I�茠��� Then
        ' ���L�^�̏�����
        Call WriteNewRecordsForSenshuken(sGameName, oWinnerList)
    Else
        ' ���L�^�̏�����
        Call WriteNewRecords(sGameName, oWinnerList)
    End If

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
    
    Call SheetActivate(sSheetName)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    ' ���N
    Dim nYear As Integer
    nYear = GetRange("���N").Value
    
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
        
        ' ���N�ƈقȂ�ꍇ�ǂݍ���
        If oWinner.Item("�N") <> nYear Then
            ' �v��No�{�敪�̂P�ʂ����o�^�̏ꍇ
            If Not (oRecordList.Exists(sKey)) Then
                Set oWinners = CreateObject("Scripting.Dictionary")
                oWinners.Add oWinners.Count + 1, oWinner
                oRecordList.Add sKey, oWinners
            Else
                Set oWinners = oRecordList.Item(sKey)
                oWinners.Add oWinners.Count + 1, oWinner
            End If
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

    Dim sKeyName As String
    Dim sKey As String

    ' �D���Җ�
    Dim vCell As Range
    For Each vCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' ���V�Ȃ�i�[����
        If GetOffset(vCell, GetColIdx(sWinnerAreaName, "���V")).Value = "���V" _
            Or GetOffset(vCell, GetColIdx(sWinnerAreaName, "���V")).Value = "�Q�l�L�^" Then
                
            Set oWinner = CreateObject("Scripting.Dictionary")
            
            ' �L�[���擾
            sKeyName = GetAreaKeyName(sRecordAreaName)
            sKey = GetRecordKey(sGameName, CInt(vCell.Value), _
                    GetOffset(vCell, GetColIdx(sWinnerAreaName, "�敪")))
            
            ' �D���҂̗�l�����ׂēo�^
            Dim vKey As Variant
            For Each vKey In GetRange(sWinnerAreaName).Rows(1).Columns()
                oWinner.Add STrimAll(vKey.Value), GetOffset(vCell, vKey.Column).Value
            Next vKey
            
            ' �I�茠�̏ꍇ��Key���͈͓��ɂ���̂Ń`�F�b�N���Ă���ǉ�
            If Not oWinner.Exists(sKeyName) Then
                oWinner.Add sKeyName, sKey
            End If
            oWinner.Add "�N", nYear
                
            ' �v��No�{�敪�̂P�ʂ����o�^�̏ꍇ
            If Not (oRecordList.Exists(sKey)) Then
                Set oWinners = CreateObject("Scripting.Dictionary")
                oWinners.Add oWinners.Count + 1, oWinner
                oRecordList.Add sKey, oWinners
            Else
                Set oWinners = oRecordList.Item(sKey)
                ' ���ɑ��݂���ꍇ�̓^�C�����r���Â���΍폜
                ' �ߋ��̑��V�����݂��Ȃ��ꍇ���폜����
                For Each vKey In oWinners.Keys()
                    Set oWinnerOld = oWinners.Item(vKey)
                    If oWinner.Item("�L�^") < oWinnerOld.Item("�L�^") Or _
                       oWinnerOld.Item("�L�^") = "" Then
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
    
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    ' �t�B���^����
    Call SetAutoFilter(sRecordAreaName, False)

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

    ' ���ёւ�
    If sGameName = �s����� Then
        Call SortRecordWinnerShimin(sSheetName, sRecordAreaName)
    Else
        Call SortRecordWinner(sSheetName, sRecordAreaName)
    End If

    ' ���L�^�V�[�g�̐ݒ�
    Call DefineRecordSheet(sSheetName, sRecordAreaName)

    ' �V�[�g�̕ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���L�^�����݁i�I�茠�p�j
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Private Sub WriteNewRecordsForSenshuken(sGameName As String, oRecordList As Object)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    
    ' �\���^�A�N�e�B�u�^����
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    ' ������
    Call DeleteWinnerSheetForSenshuken(sRecordAreaName)

    ' ������
    Call WriteWinnerListForSenshuken(sRecordAreaName, oRecordList)

    ' �D���҃V�[�g�ݒ�
    Call �I�茠���L�^�ݒ�(sSheetName, sRecordAreaName)

    ' �V�[�g�̕ی�
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' ���L�^�V�[�g�L��
'
' sGameName     IN  ��
' vCell         IN  �Q�ƌ��̊�Z��
' oWinner       IN  �D���ҏ��
'
Private Sub WriteRecordLine(sAreaName As String, vCell As Variant, oWinner As Object)

    vCell.Offset(0, GetColIdx(sAreaName, "����") - 1).Value = oWinner.Item("����")
    vCell.Offset(0, GetColIdx(sAreaName, "����") - 1).Value = oWinner.Item("����")
    vCell.Offset(0, GetColIdx(sAreaName, "�L�^") - 1).Value = oWinner.Item("�L�^")
    vCell.Offset(0, GetColIdx(sAreaName, "�N") - 1).Value = oWinner.Item("�N")

End Sub
