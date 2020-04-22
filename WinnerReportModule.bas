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
    
    Dim sProType As String
    Dim sKey As String
    Dim sName As String
    Dim sTeam As String
    Dim sType As String
    Dim nTime As Long
    Dim nRecord As Variant
    
    ' �v���O�����ԍ���
    For Each nProNo In GetAreaKeyData(sMasterName)
        sProType = Trim(VLookupArea(nProNo, sMasterName, "�敪"))
        
        ' �v���O�����ԍ�������P�ʂ�T��
        For Each oCell In GetRange("�v���O�����ԍ�" & CStr(nProNo))
            ' �P�ʂ̏ꍇ
            If oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value = 1 Then
                
                Set oWinner = CreateObject("Scripting.Dictionary")
                
                sName = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                sTeam = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                If sProType = "" Then
                    sType = oCell.Offset(0, Range("Header�敪").Column - Range("Header�v��No").Column).Value
                Else
                    sType = ""
                End If
                nTime = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                nRecord = oCell.Offset(0, Range("Header���L�^").Column - Range("Header�v��No").Column).Value
                sKey = CStr(nProNo) & sType
                
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
    Next nProNo
End Sub


'
' �D���ҏ�����
'
' sGameName     IN  ��
' oWinnerList   IN  �D���҃��X�g
'
Sub WriteWinner(sGameName As String, oWinnerList As Object)

    ' �D���҃V�[�g
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
        If oWinnerList.Exists(vKey.Value) Then
            Set oWinners = oWinnerList.Item(vKey.Value)
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteWinnerLine(sWinnerAreaName, nRow, vKey, oWinner)
                nRow = nRow + 1
            Next vIdx
        End If
    Next vKey

    ' �����ݒ�
    Call SetWinnerRecordStyle(sGameName, sWinnerAreaName, sRecordAreaName)
    
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
' sGameName     IN  ��
' nRow          IN  �s��
' vKey          IN  �Q�ƌ��̊�Z��
' oWinner       IN  �D���ҏ��
'
Sub WriteWinnerLine(sAreaName As String, nRow As Integer, vKey As Variant, oWinner As Object)

    Dim nProNoCol As Integer
    nProNoCol = GetAreaColumnIndex(sAreaName, "�v��No.")
    Dim nGenDist As Integer
    nGenDist = GetAreaColumnIndex(sAreaName, "��")
    Dim nStyle As Integer
    nStyle = GetAreaColumnIndex(sAreaName, "��")
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sAreaName, "�敪")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sAreaName, "����")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "����")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "�L�^")
    Dim nRecordCol As Integer
    nRecordCol = GetAreaColumnIndex(sAreaName, "���V")
    
    Cells(nRow, nProNoCol) = vKey.Offset(0, nProNoCol)
    Cells(nRow, nGenDist) = vKey.Offset(0, nGenDist)
    Cells(nRow, nStyle) = vKey.Offset(0, nStyle)
    Cells(nRow, nTypeCol) = vKey.Offset(0, nTypeCol)
    
    Cells(nRow, nNameCol) = oWinner.Item("����")
    Cells(nRow, nTeamCol) = oWinner.Item("����")
    Cells(nRow, nTimeCol) = oWinner.Item("�L�^")
    Cells(nRow, nRecordCol) = oWinner.Item("���V")

End Sub

'
' �D���҃V�[�g�̏�������L�^����R�s�[����
'
' sGameName         IN  ��
' sWinnerAreaName   IN  �D���҂͈͖̔�
' sRecordAreaName   IN  ���L�^�͈͖̔�
'
Sub SetWinnerRecordStyle(sGameName As String, sWinnerAreaName As String, sRecordAreaName As String)

    ' ���L�^�V�[�g
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    GetRange(sRecordAreaName).Offset(1, 1).Resize(1).Copy
    
    ' �D���҃V�[�g
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, oRange.Columns.Count).Select
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
    Call ReadNewRecords(sGameName, oWinnerList)
    
    ' ���L�^�̏�����
    Call WriteNewRecords(sGameName, oWinnerList)

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
Sub ReadNewRecords(sGameName As String, oRecordList As Object)
    
    ' �D���҃V�[�g
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select

    ' �D���҂͈̔�
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    
    ' ���N
    Dim nYear As Integer
    nYear = GetRange("���N").Value
    
    Dim oWinners As Object
    Dim oWinner As Object

    Dim nProNoCol As Integer
    nProNoCol = GetAreaColumnIndex(sWinnerAreaName, "�v��No.")
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sWinnerAreaName, "�敪")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sWinnerAreaName, "����")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sWinnerAreaName, "����")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sWinnerAreaName, "�L�^")
    Dim nRecordCol As Integer
    nRecordCol = GetAreaColumnIndex(sWinnerAreaName, "���V")
    
    ' �D���Җ�
    For Each oCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' ���V�Ȃ�i�[����
        If oCell.Offset(0, nRecordCol - nProNoCol).Value = "���V" Then
                
            Set oWinner = CreateObject("Scripting.Dictionary")
                
            sType = oCell.Offset(0, nTypeCol - nProNoCol).Value
            sName = oCell.Offset(0, nNameCol - nProNoCol).Value
            sTeam = oCell.Offset(0, nTeamCol - nProNoCol).Value
            nTime = oCell.Offset(0, nTimeCol - nProNoCol).Value
            
            ' �w���}�X�^�[�Y�ŏ��w���܂܂��ꍇ
            If sType Like "���w*" Then
                sKey = CStr(oCell.Value)
            Else
                sKey = CStr(oCell.Value) & sType
            End If
                
            oWinner.Add "����", sName
            oWinner.Add "����", sTeam
            oWinner.Add "�L�^", nTime
            oWinner.Add "�N", nYear
                
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
    Next oCell
End Sub


'
' �D���ҏ�����
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
