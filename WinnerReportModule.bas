Attribute VB_Name = "WinnerReportModule"
Sub �D���҈ꗗ�쐬()

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
Sub ReadWinner(sGameName As String, oWinnerList As Object)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' �P�ʃ��X�g
    Dim oWinners As Object
    Set oWinners = Nothing
    ' �P�ʏ��
    Dim oWinner As Object
    Set oWinner = Nothing
    
    Dim sKey As String
    Dim sName As String
    Dim sTeam As String
    Dim nTime As Long
    Dim nRecord As Variant
    Dim bFlag As Boolean
    
    ' �v���O�����ԍ���
    For Each nProNo In GetAreaKeyData(sMasterName)
        ' �I�茠�̗\�I�͔�΂�
        If sGameName = �I�茠��� Then
             If VLookupArea(nProNo, "�I�茠��ڋ敪", "�\�I�^����") = "�\�I" Then
                bFlag = False
             Else
                bFlag = True
             End If
        Else
            bFlag = True
        End If
        
        ' �����i�^�C�������j�̏ꍇ
        If bFlag Then
            ' �v���O�����ԍ�������P�ʂ�T��
            For Each oCell In GetRange("�v���O�����ԍ�" & CStr(nProNo))
                ' �P�ʂ̏ꍇ
                If oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value = 1 Then
                    
                    Set oWinner = CreateObject("Scripting.Dictionary")
                    
                    sName = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                    sTeam = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                    nTime = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                    nRecord = oCell.Offset(0, Range("Header���L�^").Column - Range("Header�v��No").Column).Value
                    sKey = GetWinnerKey(sGameName, sMasterName, CInt(nProNo), _
                        oCell.Offset(0, Range("Header�敪").Column - Range("Header�v��No").Column).Value)
    
                    oWinner.Add "����", sName
                    oWinner.Add "����", sTeam
                    oWinner.Add "�L�^", nTime
                    If Not IsNumeric(nRecord) Or nTime <= nRecord Then
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
                End If
            Next oCell
        End If
    Next nProNo
End Sub

'
' ���̑��L�^�̋敪
'
' sGameName     IN  ��
' sMasterName   IN  �}�X�^�[��
' nProNo        IN  ��ڔԍ�
' sType         IN  �敪
'
Function GetWinnerKey(sGameName As String, sMasterName As String, nProNo As Integer, sType As String)

    If sGameName = �I�茠��� Then
        GetWinnerKey = CStr(nProNo)
    ElseIf sGameName = �s����� Then
        GetWinnerKey = CStr(nProNo) & sType
    Else
        ' �敪���擾
        If Trim(VLookupArea(nProNo, sMasterName, "�敪")) = "" Then
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
Sub WriteWinner(sGameName As String, oWinnerList As Object)

    ' �D���҃V�[�g��I�����ی������
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' �D���Ҕ͈͖�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)

    ' �폜
    Call DeleteWinnerSheet(sWinnerAreaName)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
        
    ' ���L�^��
    For Each vKey In GetAreaKeyData(sRecordAreaName)
        If oWinnerList.Exists(CStr(vKey.Value)) Then
            Set oWinners = oWinnerList.Item(CStr(vKey.Value))
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteWinnerLine(sWinnerAreaName, sRecordAreaName, nRow, vKey, oWinner)
                nRow = nRow + 1
            Next vIdx
        End If
    Next vKey

    ' �����ݒ�
    Call SetWinnerRecordStyle(sGameName)
    
    ' ����͈͂̐ݒ�
    ActiveSheet.PageSetup.PrintArea = TableRangeAddress(sWinnerAreaName)

End Sub

Sub DeleteWinnerSheet(sWinnerAreaName As String)
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    If Cells(oRange.Row + 1, oRange.Column) <> "" Then
        oRange.Offset(1, 0).Resize(oRange.Rows().Count - 1).EntireRow.Delete
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
Sub WriteWinnerLine(sWinnerAreaName As String, sRecordAreaName As String, nRow As Integer, vKey As Variant, oWinner As Object)
    
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
Sub SetWinnerRecordStyle(sGameName As String)

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
' ���L�^�V�[�g�̏�����D���҂���R�s�[����
'
' sGameName         IN  ��
'
Sub SetRecordWinnerStyle(sGameName As String)

    ' �D���҃V�[�g
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    
    ' �D���Ҕ͈͖�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    GetRange(sWinnerAreaName).Offset(1, 0).Resize(1).Copy
    
    ' ���L�^�V�[�g
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    
     ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    Dim oRange As Range
    Set oRange = TableRange(sRecordAreaName)
    If Cells(oRange.Row + 1, oRange.Column) <> "" Then
        oRange.Offset(1, 1).Resize(oRange.Rows.Count - 1, oRange.Columns.Count - 1).Select
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
Function GetWinnerSheetName(sGameName As String)
    If sGameName = "���{��I�茠���j���" Then
        GetWinnerSheetName = "�I�茠���D����"
    ElseIf sGameName = "���{��s���̈���" Then
        GetWinnerSheetName = "�s�����D����"
    Else
        GetWinnerSheetName = "�w���}�X�^�[�Y�D����"
    End If
End Function

'
' �D���Ҕ͈͖�
'
' sGameName     IN      ��
'
Function GetWinnerAreaName(sGameName As String)
    If sGameName = "���{��I�茠���j���" Then
        GetWinnerAreaName = "�I�茠���D����"
    ElseIf sGameName = "���{��s���̈���" Then
        GetWinnerAreaName = "�s�����D����"
    Else
        GetWinnerAreaName = "�w�}���D����"
    End If
End Function

'
' ���L�^�V�[�g��
'
' sGameName     IN      ��
'
Function GetRecordSheetName(sGameName As String)
    If sGameName = "���{��I�茠���j���" Then
        GetRecordSheetName = "�I�茠���L�^"
    ElseIf sGameName = "���{��s���̈���" Then
        GetRecordSheetName = "�s�����L�^"
    Else
        GetRecordSheetName = "�w���}�X�^�[�Y���L�^"
    End If
End Function


'
' ���L�^�͈͖�
'
' sGameName     IN      ��
'
Function GetRecordAreaName(sGameName As String)
    If sGameName = "���{��I�茠���j���" Then
        GetRecordAreaName = "�I�茠���L�^"
    ElseIf sGameName = "���{��s���̈���" Then
        GetRecordAreaName = "�s�����L�^"
    Else
        GetRecordAreaName = "�w�}���L�^"
    End If
End Function

Sub ���L�^�X�V()

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

    Call ���[�N�u�b�N���O��`

    ' �C�x���g�����𔭐�
    Call EventChange(True)

    ' �ۑ�
    ActiveWorkbook.Save

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
Sub ReadRecords(sGameName As String, oRecordList As Object)
    
    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim sKey As String
        
    ' ���L�^��
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        sKey = vCell.Value
        
        Set oWinner = CreateObject("Scripting.Dictionary")
        
        ' �J�����̒l�̓o�^
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
Sub ReadNewRecords(sGameName As String, oRecordList As Object)
    
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

    ' �D���Җ�
    For Each oCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' ���V�Ȃ�i�[����
        If GetOffset(oCell, GetAreaColumnIndex(sWinnerAreaName, "���V")).Value = "���V" Then
                
            ' �L�[���擾
            sKey = GetWinnerKey(sGameName, sMasterName, CInt(oCell.Value), _
                    GetOffset(oCell, GetAreaColumnIndex(sWinnerAreaName, "�敪")))
                
            Set oWinner = CreateObject("Scripting.Dictionary")
            
            ' �J�����̒l�̓o�^
            For Each vKey In GetRange(sWinnerAreaName).Rows(1).Columns()
                oWinner.Add STrimAll(vKey.Value), GetOffset(oCell, vKey.Column).Value
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
    Next oCell
End Sub

'
' ���L�^������
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Sub WriteNewRecords(sGameName As String, oRecordList As Object)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    ' �폜
    Call DeleteWinnerSheet(sRecordAreaName)

    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
    
    ' ���L�^��
    For Each vKey In oRecordList.Keys()
        Set oWinners = oRecordList.Item(vKey)
        For Each vIdx In oWinners
            Set oWinner = oWinners.Item(vIdx)
            
            ' �J�����̒l�̏�����
            For Each vCell In GetRange(sRecordAreaName).Rows(1).Columns()
                Cells(nRow, vCell.Column) = oWinner.Item(STrimAll(vCell.Value))
            Next vCell
            
            nRow = nRow + 1
        Next vIdx
    Next vKey

    ' �����ݒ�
    Call SetRecordWinnerStyle(sGameName)

    ' �V�[�g�̕ی�
    Call SheetProtect(True)
End Sub

'
' ���L�^������
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Sub WriteNewRecordsOld(sGameName As String, oRecordList As Object)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' ���L�^�͈͖�
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    ' ���L�^��
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        If oRecordList.Exists(vCell.Value) Then
            Set oWinners = oRecordList.Item(vCell.Value)
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteRecordLine(sRecordAreaName, vCell, oWinner)
            Next vIdx
        End If
    Next vCell

    ' �V�[�g�̕ی�
    Call SheetProtect(True)
End Sub


'
' ���L�^�V�[�g�L��
'
' sGameName     IN  ��
' vCell         IN  �Q�ƌ��̊�Z��
' oWinner       IN  �D���ҏ��
'
Sub WriteRecordLine(sAreaName As String, vCell As Variant, oWinner As Object)

    nNameCol = GetAreaColumnIndex(sAreaName, "����")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "����")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "�L�^")
    Dim nRecordCol As Integer
    nRecordCol = GetAreaColumnIndex(sAreaName, "�N")
    
    vCell.Offset(0, nNameCol - 1).Value = oWinner.Item("����")
    vCell.Offset(0, nTeamCol - 1).Value = oWinner.Item("����")
    vCell.Offset(0, nTimeCol - 1).Value = oWinner.Item("�L�^")
    vCell.Offset(0, nRecordCol - 1).Value = oWinner.Item("�N")

End Sub
