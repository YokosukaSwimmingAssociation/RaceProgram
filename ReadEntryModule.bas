Attribute VB_Name = "ReadEntryModule"
Option Explicit    ''���ϐ��̐錾����������

'
' �G���g���[�t�@�C���ꗗ�̓ǂݍ���
'
' ����̃t�H���_���w�肵�āA���ɂ���G���g���[�t�@�C����
' ���ׂēǂݍ��݈ꗗ�V�[�g�ɏo�͂���B
'
Public Sub �G���g���[�Ǎ���(Optional sPathName As String = "")
    ' �G�N�Z���V�[�g��I��
    Call SheetActivate(�G���g���[�V�[�g)

    ' �o�͗p���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' �o�͗p���[�N�V�[�g
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    oWorkSheet.Select
    Call SetForcusTop

    ' �C�x���g������}��
    Call EventChange(False)

    ' �G���g���[�ꗗ�Ǎ��p�z��
    Dim oGameList As Object
    Set oGameList = CreateObject("Scripting.Dictionary")

    ' �G���g���[�e�[�u����������
    Call DeleteTable(oWorkSheet, �G���g���[�e�[�u��)

    ' �G���g���[�t�@�C���Ǎ���
    Call ReadEntryFiles(oGameList, sPathName)

    ' �G���g���[�V�[�g�̏�������
    Call WriteEntrySheet(oWorkSheet, �G���g���[�e�[�u��, oGameList)
    
    ' ProNo�A�\�[�g�敪�A�\���ݎ��ԂŃ\�[�g
    Call SetTitleMenu("���ёւ���")
    Call SortByProNo(oWorkSheet, �G���g���[�e�[�u��)

    Call SetTitleMenu("")
    
    ' �V�[�g��ۑ�
    oWorkBook.Save

End Sub

'
' �e�[�u����������
'
Public Sub �G���g���[�ꗗ������()
    ' �G���g���[�e�[�u����������
    Call DeleteTable(Sheets(�G���g���[�V�[�g), �G���g���[�e�[�u��)
End Sub

'
' �G���g���[�t�@�C���ꗗ�̓ǂݍ���
'
' �t�H���_���w�肵�āA���̒��Ɋ܂܂��G���g���[�V�[�g�i*.xlsx�j�����ׂĉr�ݍ���
'
' oGameList     OUT     �G���g���[�ꗗ
' sDirName      OUT     �t�H���_��
'
Private Sub ReadEntryFiles(ByRef oGameList As Object, Optional sPathName As String = "")

    ' �t�@�C���ꗗ���擾
    '
    If sPathName = "" Then
        sPathName = SelectDir()
    End If
    Dim FileList As Collection
    Set FileList = GetFiles(sPathName, "\*.xlsx")

    Dim nMax As Integer
    nMax = FileList.Count
    Dim nCount As Integer
    nCount = 0

    '
    ' �t�@�C�����ɏ�������
    '
    Dim SubBook As Workbook
    Dim vFile As Variant
    For Each vFile In FileList
        
        ' �^�C�g���C��
        nCount = nCount + 1
        Call SetTitleMenu("�v���O�����Ǎ���: " & Str(nCount) & "/" & Str(nMax))
        
        '
        ' �t�@�C�����J���i�ǂݎ���p�j
        '
        Set SubBook = Workbooks.Open(Filename:=sPathName + "\" + vFile, ReadOnly:=True)
        SheetActivate ("�L���[")

        ' �G���g���[�ꗗ�̓Ǎ���
        Call ReadEntryFile(oGameList)
    
        ' �x���Ȃ��Ńt�@�C�������i�ۑ����Ȃ��j
        Application.DisplayAlerts = False
        SubBook.Close
        Application.DisplayAlerts = True
    Next vFile
    
    ' �^�C�g���C��
    Call SetTitleMenu("�v���O�����Ǎ�����: " & Str(nCount) & "/" & Str(nMax))
    
End Sub



'
' �G���g���[�t�@�C���̓Ǎ���
'
' oGameList
' ��
' �������FoTeamList       �E�E�ERange("��")
' �@�@��
' �@�@�����`�[�����FoEntryList�E�E�ERange("�`�[����")
' �@�@�@�@��
' �@�@�@�@�����I��ԍ��FoEntry�E�E�ERange("�I��ԍ�")
'
' oEntry
' ��
' �������ʁFRange("�I�萫��")
' ��
' �����I�薼�FRange("�I�薼")
' ��
' �����t���K�i�FRange("�I��t���K�i")
' ��
' �����敪�FRange("�I��敪")
' ��
' ����1or2�FoLines
'
' oLines
' ��
' ������ږ��FRange("���")
' ��
' ���������FRange("��ڋ���")
' ��
' �����\���ݎ��ԁFRange("�I�蕪")�{Range("�I��b")�{Range("�I��~���b")
'
' oRelayEntry
' ��
' ����L1�`L24�FoRelayLines
'
' oRelayLines
' ��
' ������ڔԍ��FRange("�����[���")
' ��
' �����敪�FRange("�����[�敪")
' ��
' �����\���ݎ��ԁFRange("�����[��")�{Range("�����[�b")�{Range("�����[�~���b")
'
' oGameList     OUT     �G���g���[�ꗗ
'
Private Sub ReadEntryFile(ByRef oGameList As Object)

    ' ��
    Dim sGameName As String
    sGameName = GetRange("��").Value
    Dim oTeamList As Object
    If oGameList.Exists(sGameName) Then
        Set oTeamList = oGameList.Item(sGameName)
    Else
        Set oTeamList = CreateObject("Scripting.Dictionary")
        oGameList.Add sGameName, oTeamList
    End If
    
    ' �`�[����
    Dim sTeamName As String
    sTeamName = RTrim(LTrim(GetRange("�`�[����").Value))
    
    ' �`�[���\���ݕۑ��I�u�W�F�N�g
    Dim nNum As Integer
    Dim oEntryList As Object
    If oTeamList.Exists(sTeamName) Then
        If sTeamName = "�l" Then
            Set oEntryList = oTeamList.Item(sTeamName)
            nNum = oEntryList.Count
        Else
            MsgBox "�`�[�������d�����Ă��܂��B" + vbCrLf + sGameName + " : " + sTeamName
            End
        End If
    Else
        Set oEntryList = CreateObject("Scripting.Dictionary")
        oTeamList.Add sTeamName, oEntryList
    End If

    ' �l�p�G���g���[�̓Ǎ���
    Call ReadPersonEntry(nNum, sTeamName, oEntryList)

    ' �����[�p�G���g���[�̓Ǎ���
    ' �������[�����̓��ʑΉ�(7/3�߂�)
    Call ReadRelayEntry(nNum, oEntryList)

End Sub

'
' �l�p�G���g���[�̓Ǎ���
'
' nNum          IN      �l�G���g���[�s(1,2)
' sTeam         IN      ����
' oEntryList    OUT     ��ڍs
'
Private Sub ReadPersonEntry(nNum As Integer, sTeamName As String, ByRef oEntryList As Object)
    ' �l�ԍ��͈͂����ׂēǂݍ���
    Dim oCell As Range
    For Each oCell In GetRange("�I��ԍ�")
        
        ' �l�ԍ��͌�������Ă���
        If oCell.MergeCells Then
            ' �����̐擪�s�ŏ�������
            If oCell.Address = oCell.MergeArea.Item(1).Address Then
        
                ' �I�薈�̃G���g���[���X�g
                Dim oEntry As Object
                Set oEntry = CreateObject("Scripting.Dictionary")
            
                ' �lNo
                If sTeamName = "�l" Then
                    nNum = nNum + 1
                Else
                    nNum = oCell.Value
                End If

                ' �I�����o�^
                Call ReadEntrySwimmer(nNum, oCell, oEntry)
                
                ' �P�s��
                Call ReadEntryLine(1, oCell.Row, oEntry)
                Call CheckEntry(1, oCell.Row, oEntry)
    
                ' �Q�s��
                Call ReadEntryLine(2, oCell.Row, oEntry)
                Call CheckEntry(2, oCell.Row, oEntry)
                If oEntry.Item("�I�薼") <> "" Then
                    oEntryList.Add nNum, oEntry
                End If
            End If
        End If
    Next oCell
End Sub

'
' �G���g���[�t�@�C���̌l���̓Ǎ���
'
' ���ʁA�I�薼�A�t���K�i�A�敪����ǂݍ���
'
' nNum          IN      �l�G���g���[�s(1,2)
' nRow          IN      �s��
' oEntry        OUT     ��ڍs
'
Private Sub ReadEntrySwimmer(nNum As Integer, oCell As Range, ByRef oEntry As Object)

    oEntry.Add "����", GetOffset(oCell, GetRange("�I�萫��").Column).Value + "�q"
    oEntry.Add "�t���K�i", ReplaceName(GetOffset(oCell, GetRange("�I��t���K�i").Column).Value)
    
    If Range("��").Value = �I�茠��� Then
        
        oEntry.Add "�I�薼", ReplaceName(GetOffset(oCell, GetRange("�I�薼").Column).Offset(1).Value)
        oEntry.Add "�敪", GetOffset(oCell, GetRange("�I��敪").Column).Value
    
    ElseIf Range("��").Value = �s����� Then
        oEntry.Add "�I�薼", ReplaceName(GetOffset(oCell, GetRange("�I�薼").Column).Offset(1).Value)
        oEntry.Add "�w�Z��", Trim(GetOffset(oCell, GetRange("�I��w�Z��").Column).Offset(2).Value)
        If GetOffset(oCell, GetRange("�I��敪").Column).Value <> "" Then
            oEntry.Add "�敪", GetOffset(oCell, GetRange("�I��敪").Column).Value
        Else
            oEntry.Add "�敪", "�N��敪"
        End If
        oEntry.Add "�N��", GetOffset(oCell, GetRange("�I��N��").Column).Offset(1).Value
    
    ElseIf Range("��").Value = �}�X�^�[�Y��� Then
    
        oEntry.Add "�I�薼", ReplaceName(GetOffset(oCell, GetRange("�I�薼").Column).Offset(1).Value)
        oEntry.Add "�N��", GetOffset(oCell, GetRange("�I��N��").Column).Value
    
    ElseIf Range("��").Value = �����L�^�� Then
    
        oEntry.Add "�I�薼", ReplaceName(GetOffset(oCell, GetRange("�I�薼").Column).Offset(1).Value)
        oEntry.Add "�N��", GetOffset(oCell, GetRange("�I��N��").Column).Value
        oEntry.Add "�敪", GetOffset(oCell, GetRange("�I��w�N").Column).Offset(1).Value
        oEntry.Add "����", GetOffset(oCell, GetRange("�I�茟��").Column).Value
    
    Else
    
        oEntry.Add "�I�薼", ReplaceName(GetOffset(oCell, GetRange("�I�薼").Column).Offset(1).Value)
        oEntry.Add "�敪", GetOffset(oCell, GetRange("�I��w�N").Column).Value
    
    End If

End Sub

'
' �G���g���[�t�@�C���̌l��ڍs�ǂݍ���
'
' ��ږ��A�����A�\���ݎ��Ԃ��擾����
'
' nNum          IN      �l�G���g���[�s(1,2)
' nRow          IN      �s��
' oEntry        OUT     ��ڍs
'
Private Sub ReadEntryLine(nNum As Integer, nRow As Integer, oEntry As Object)
    
    Dim oLines As Object
    Dim sStyle As String
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    Dim oProNo As Range
    
    ' �ԍ��͈͂����ׂēǂݍ���
    Dim oCell As Range
    For Each oCell In GetRange("��ڈꗗ")
        If oCell.Value <> "" Then
            sStyle = oCell.Value
        End If
        ' ��ڑI������ȊO�̏ꍇ�͑I�����ꂽ���̂Ƃ���
        Set oProNo = GetRowOffset(oCell, nRow).Offset(nNum - 1)
        If Trim(oProNo.Value) <> "" Then
            Set oLines = CreateObject("Scripting.Dictionary")
            oEntry.Add nNum, oLines
            
            oLines.Add "��ڔԍ�", VLookupArea(oProNo.Value, "��ڔԍ��敪", "��ڔԍ�")
            oLines.Add "��ڋ敪", VLookupArea(oProNo.Value, "��ڔԍ��敪", "��ڋ敪") ' ���N�����敪�ɂ��Ă���
            oLines.Add "���", ReplaceStyle(sStyle)
            oLines.Add "����", ReplaceDistance(GetRowOffset(oCell, GetRange("��ڋ���").Row).Value)
            nMin = GetOffset(oProNo, GetRange("�I�蕪").Column).Value
            nSec = GetOffset(oProNo, GetRange("�I��b").Column).Value
            nMil = GetOffset(oProNo, GetRange("�I��~���b").Column).Value
            oLines.Add "�\���ݎ���", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
            Exit Sub
        End If
    Next oCell
End Sub

'
' �G���g���[�̎�ڔԍ��������������m�F
'
' nNum          IN      �l�G���g���[�s(1,2)
' nRow          IN      �s�ԍ�
' oEntry        IN      ��ڍs
'
Private Sub CheckEntry(nNum As Integer, nRow As Integer, oEntry As Object)
    
    If IsEmpty(oEntry.Item(nNum)) Then
        Exit Sub
    End If
    
    Dim oLines As Object
    Set oLines = oEntry.Item(nNum)
    
    Dim sGender As String
    Dim sDistance As String
    Dim sStyle As String
    
    sGender = VLookupArea(oLines.Item("��ڔԍ�"), "��ڔԍ��敪", "����")
    sDistance = VLookupArea(oLines.Item("��ڔԍ�"), "��ڔԍ��敪", "����")
    sStyle = VLookupArea(oLines.Item("��ڔԍ�"), "��ڔԍ��敪", "���")
    
    If sGender <> oEntry.Item("����") Or sDistance <> oLines.Item("����") Or sStyle <> oLines.Item("���") Then
        MsgBox CStr(nRow) & "�s�ځF��ڔԍ�������������܂���B�F" & oLines.Item("��ڔԍ�")
        End
    End If

End Sub

'
' �����[��ڂ̓Ǎ���
'
' ��ږ��A�����A�\���ݎ��Ԃ��擾����
'
' nNum          IN      �G���g���[�s
' oEntryList    OUT     �G���g���[�ꗗ
'
Private Sub ReadRelayEntry(nNum As Integer, ByRef oEntryList As Object)

    ' �����[��ڔԍ��͈͂����ׂēǂݍ���
    Dim nRelayNum As Integer
    nRelayNum = 0
    Dim oRelayEntry As Object
    Set oRelayEntry = Nothing
    Dim oCell As Range
    For Each oCell In GetRange("�����[���")
        ' �l���ݒ肳��Ă���ꍇ�͓ǂݍ���
        If oCell.Value <> "" Then
            ' �����[�̃G���g���[���X�g
            If oRelayEntry Is Nothing Then
                Set oRelayEntry = CreateObject("Scripting.Dictionary")
                oEntryList.Add nNum, oRelayEntry
            End If

            nRelayNum = nRelayNum + 1
            Call ReadRelayEntryLine(nRelayNum, oCell, oRelayEntry)
        End If
    Next oCell

End Sub

'
' �G���g���[�t�@�C���̃����[��ڍs�ǂݍ���
'
' nNum          IN    �����[�ԍ�
' oCell         IN    �J�����g�Z��
' oRelayEntry   I/O   ��ڍs
'
Private Sub ReadRelayEntryLine(nNum As Integer, oCell As Range, oRelayEntry As Object)
    
    Dim oLines As Object
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    If oCell.Value <> "" Then
        Dim oRelayLines As Object
        Set oRelayLines = CreateObject("Scripting.Dictionary")
        oRelayEntry.Add "L" + Str(nNum), oRelayLines

        oRelayLines.Add "��ڔԍ�", oCell.Value
        If IsNameExists("�����[�敪") Then
            oRelayLines.Add "�敪", GetOffset(oCell, GetRange("�����[�敪").Column).Value
        End If
        nMin = GetOffset(oCell, GetRange("�����[��").Column).Value
        nSec = GetOffset(oCell, GetRange("�����[�b").Column).Value
        nMil = GetOffset(oCell, GetRange("�����[�~���b").Column).Value
        oRelayLines.Add "�\���ݎ���", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
    End If
End Sub

'
' ��ږ��̂̒u��
'
' sStyle        IN      ���
'
Private Function ReplaceStyle(sStyle) As String
    Dim sTemp As String
    sTemp = sStyle
    sTemp = Replace(sTemp, "����ײ", "�o�^�t���C")
    sTemp = Replace(sTemp, "��", "�l���h���[")
    ReplaceStyle = sTemp
End Function

'
' �������̂̒u��
'
' sDistance     IN      ����
'
Private Function ReplaceDistance(sDistance) As String
    Dim sTemp As String
    sTemp = sDistance
    sTemp = Replace(sTemp, "���", "25M")
    sTemp = Replace(sTemp, "�܁Z", "50M")
    sTemp = Replace(sTemp, "��Z�Z", "100M")
    sTemp = Replace(sTemp, "��Z�Z", "200M")
    sTemp = Replace(sTemp, "�l�Z�Z", "400M")
    ReplaceDistance = sTemp
End Function

'
' �I�薼�̒u��
'
' �����P�����̏ꍇ�͐��ɑS�p�󔒂𑫂�
' �����Q�����ȓ��Ŗ����P�����̏ꍇ�͖��ɑS�p�󔒂𑫂�
'
' sName         IN      �I�薼
'
Private Function ReplaceName(sName) As String
    
    ' �󔒂̏ꍇ�͉������Ȃ�
    If Trim(sName) = "" Then
        ReplaceName = ""
        Exit Function
    End If
    
    Dim sTemp As String
    sTemp = STrim(sName)
    
    Dim sTemps As Variant
    sTemps = Split(sTemp, " ")
    ' �����P�����̏ꍇ�͐��ɑS�p�󔒂𑫂�
    If Len(sTemps(0)) = 1 Then
        sTemps(0) = sTemps(0) & "�@"
    End If
    ' �����Q�����ȓ��Ŗ����P�����̏ꍇ�͖��ɑS�p�󔒂𑫂�
    If Len(sTemps(1)) = 1 And Len(sTemps(0)) <= 2 Then
        sTemps(1) = "�@" & sTemps(1)
    End If
        
    ReplaceName = sTemps(0) & "�@" & sTemps(1)
End Function

'
' �\���݂��V�[�g�ɏo��
'
' oWorkSheet    IN     ���[�N�V�[�g
' sTable        IN     �e�[�u����
' oTeamList     IN     �ǂݍ��񂾃`�[���\���݈ꗗ
'
Private Sub WriteEntrySheet(oWorkSheet As Worksheet, sTable As String, oGameList As Object)

    oWorkSheet.Select

    Dim oTeamList As Object
    Dim oEntryList As Object

    Dim nTeamToal As Integer
    nTeamToal = 0
    Dim nTeamNo As Integer
    nTeamNo = 1
    Dim nRow As Integer
    nRow = 1
    
    Dim vGame As Variant
    For Each vGame In oGameList.Keys()
        ' �o�͂��������
        If IsSameGame(CStr(vGame)) Then
            Set oTeamList = oGameList.Item(CStr(vGame))
        Else
            ' �o�͂�����ȊO�͎̂Ă�
            Set oTeamList = CreateObject("Scripting.Dictionary")
        End If
        
        nTeamToal = nTeamToal + oTeamList.Count
        
        Dim vTeam As Variant
        For Each vTeam In oTeamList.Keys()
            Set oEntryList = oTeamList.Item(vTeam)
            
            ' �i���\��
            Call SetTitleMenu("�v���O�����o��: " & CStr(nTeamNo) & "/" & CStr(nTeamToal))
            
            ' �`�[�����Ƃɏo�͂���
            Call WriteTeamEntry(sTable, nRow, CStr(vGame), CStr(vTeam), nTeamNo, oEntryList)
            
            ' �`�[���ԍ����C���N�������g
            nTeamNo = nTeamNo + 1
        Next
    Next
End Sub

'
' �o�͂���Q�[��������
'
' sGame         IN      ��
'
Private Function IsSameGame(sGameName As String) As Boolean
    If GetRange("��").Value = sGameName Then
        IsSameGame = True
    ElseIf GetRange("��").Value = �w�}��� Then
        If sGameName = �w����� Or sGameName = �}�X�^�[�Y��� Then
            IsSameGame = True
        Else
            IsSameGame = False
        End If
    Else
        IsSameGame = False
    End If

End Function

'
' �`�[���̏o�͂�����
'
' sTable        IN      �e�[�u����
' nRow          IN/OUT  �o�͂���s��
' sGame         IN      ��
' sTeam         IN      �`�[����
' nTeamNo       IN      �`�[���ԍ�
' oEntryList    IN      �G���g���[�ꗗ
'
Private Sub WriteTeamEntry(sTable As String, ByRef nRow As Integer, _
sGame As String, sTeam As String, _
nTeamNo As Integer, oEntryList As Object)
    Dim oEntry As Object
    Dim oLine As Object
    Dim nPersonNo As Integer
    
    Dim vNum As Variant
    For Each vNum In oEntryList.Keys()
        Set oEntry = oEntryList.Item(vNum)
        nPersonNo = nTeamNo * 100 + CInt(vNum)
        
        If oEntry.Exists("�I�薼") Then
            ' �l
            Call WriteTeamPersonLine(sTable, nRow, sGame, sTeam, nPersonNo, oEntry, oLine)
        Else
            ' �����[
            Call WriteTeamRelayLine(sTable, nRow, sGame, sTeam, nTeamNo, oEntry, oLine)
        End If
    Next vNum
End Sub

'
' �l�̏o�͂�����
'
' sTable        IN      �e�[�u����
' nRow          IN/OUT  �o�͂���s��
' sGame         IN      ��
' sTeam         IN      �`�[����
' nPersonNo     IN      �l�ԍ�
' oEntry        IN      �G���g���[
' oLine         IN      ��ځA�\���ݎ���
'
Private Sub WriteTeamPersonLine(sTable As String, ByRef nRow As Integer, _
sGame As String, sTeam As String, nPersonNo As Integer, _
oEntry As Object, oLine As Object)
    Dim i As Integer
    For i = 1 To �l�ő�s��
        If Not IsEmpty(oEntry.Item(i)) Then
            nRow = nRow + 1
            Set oLine = oEntry.Item(i)
            Call WriteLine(sTable, nRow, sGame, sTeam, nPersonNo, oEntry, oLine)
        End If
    Next i
End Sub

'
' �����[�̏o�͂�����
'
' sTable        IN      �e�[�u����
' nRow          IN/OUT  �o�͂���s��
' sGame         IN      ��
' sTeam         IN      �`�[����
' nTeamNo       IN      �`�[���ԍ�
' oEntry        IN      �G���g���[
' oLine         IN      ��ځA�\���ݎ���
'
Private Sub WriteTeamRelayLine(sTable As String, ByRef nRow As Integer, _
sGame As String, sTeam As String, nTeamNo As Integer, _
oEntry As Object, oLine As Object)
    Dim i As Integer
    Dim sKey As String
    For i = 1 To �����[�ő�s��
        sKey = "L" & Str(i)
        If oEntry.Exists(sKey) Then
            nRow = nRow + 1
            Set oLine = oEntry.Item(sKey)
            Call WriteRelayLine(sTable, nRow, sGame, sTeam, nTeamNo, oEntry, oLine)
        End If
    Next i
End Sub


'
' �\���ݍs���o��
'
' sTable        IN      �e�[�u����
' nRow          IN      �o�͍s�ԍ�
' sGame         IN      ��
' sTeam         IN      �`�[����
' nPersonNo     IN      �I��ԍ�
' oEntry        IN      �G���g���[
' oLine         IN      ��ځA�\���ݎ���
'
Private Sub WriteLine( _
    sTable As String, _
    nRow As Integer, _
    sGame As String, _
    sTeam As String, _
    nPersonNo As Integer, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[�lNo]").Column).Value = nPersonNo
    Cells(nRow, Range(sTable & "[�v��No]").Column).Value = oLine.Item("��ڔԍ�")
    Cells(nRow, Range(sTable & "[�`�[����]").Column).Value = sTeam
    Cells(nRow, Range(sTable & "[�I�薼]").Column).Value = oEntry.Item("�I�薼")
    Cells(nRow, Range(sTable & "[�t���K�i]").Column).Value = oEntry.Item("�t���K�i")
    Cells(nRow, Range(sTable & "[����]").Column).Value = oEntry.Item("����")
    Cells(nRow, Range(sTable & "[����]").Column).Value = oLine.Item("����")
    Cells(nRow, Range(sTable & "[���]").Column).Value = oLine.Item("���")
    Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).Value = oLine.Item("�\���ݎ���")
    If oLine.Item("�\���ݎ���") >= 10000 Then
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    If sGame = �I�茠��� Then
    
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oEntry.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    
    ElseIf sGame = �s����� Then
    
        Cells(nRow, Range(sTable & "[�w�Z��]").Column).Value = oEntry.Item("�w�Z��")
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = oEntry.Item("�N��")
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = oEntry.Item("�敪")
        
        ' �l�N��敪
        If oEntry.Item("�敪") = "�N��敪" Then
            Dim nColumn As Integer
            nColumn = VLookupArea(oLine.Item("��ڔԍ�"), "�s����ڋ敪", "�^�C�v")
            Dim sClass As String
            sClass = Application.WorksheetFunction.VLookup(oEntry.Item("�N��"), GetRange("�s���N��敪"), nColumn, False)
            Cells(nRow, Range(sTable & "[�敪]").Column).Value = sClass
            If sClass = "���" Then
                Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = "20"
            Else
                Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = Left(sClass, 2)
            End If
        ' �l����
        Else
            Cells(nRow, Range(sTable & "[�敪]").Column).Value = oEntry.Item("�敪")
            Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
        End If
    
    ElseIf sGame = �}�X�^�[�Y��� Then
        
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = oEntry.Item("�N��")
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�N��"), "�w�}�N��敪", "M�N��敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�N��"), "�w�}�N��敪", "M�N��敪")

    ElseIf sGame = �w����� Then
        
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = oLine.Item("��ڋ敪")
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�敪"), "�w�}�w�N�\��", "�w�N�\��")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    
    ElseIf sGame = �����L�^�� Then
        
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = oEntry.Item("�N��")
        'Cells(nRow, Range(sTable & "[�敪]").Column).Value = _
        '    VLookupArea(oEntry.Item("�N��") & "_" & oEntry.Item("�敪"), "�L�^��N��敪", "�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�N��") & "_" & oEntry.Item("�敪"), "�L�^��N��敪", "�\�[�g")
        Cells(nRow, Range(sTable & "[����]").Column).Value = oEntry.Item("����")
    
    End If
    
End Sub

'
' �����[�\���ݍs���o��
'
' sTable        IN      �e�[�u����
' nRow          IN      �o�͍s�ԍ�
' sGame         IN      ��
' sTeam         IN      �`�[����
' nTeamNo       IN      �`�[���ԍ�
' oEntry        IN      �G���g���[
' oLine         IN      ��ځA�\���ݎ���
'
Private Sub WriteRelayLine( _
    sTable As String, _
    nRow As Integer, _
    sGame As String, _
    sTeam As String, _
    nTeamNo As Integer, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[�lNo]").Column).Value = nTeamNo
    Cells(nRow, Range(sTable & "[�`�[����]").Column).Value = sTeam
    
    Cells(nRow, Range(sTable & "[�v��No]").Column).Value = oLine.Item("��ڔԍ�")
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGame)
    
    Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "��ڋ敪")
    
    Cells(nRow, Range(sTable & "[����]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "����")
    
    Cells(nRow, Range(sTable & "[����]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "����")
    
    Cells(nRow, Range(sTable & "[���]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "���")

    Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).Value = oLine.Item("�\���ݎ���")
    If oLine.Item("�\���ݎ���") >= 10000 Then
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    If sGame = �I�茠��� Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oLine.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    
    ElseIf sGame = �s����� Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oLine.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = oLine.Item("�敪")
    
    ElseIf sGame = �}�X�^�[�Y��� Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oLine.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = oLine.Item("�敪")
    
    ElseIf sGame = �w����� Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = "���w"
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    End If
    
End Sub

'
' �G���g���[�e�[�u����������
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
'
Public Sub DeleteTable(oWorkSheet As Worksheet, sTableName As String)
    Dim myTable As ListObject
    Set myTable = oWorkSheet.ListObjects(sTableName)
    If Not (myTable.DataBodyRange Is Nothing) Then
        myTable.DataBodyRange.Delete
    End If
End Sub

'
' �V�[�g�̃e�[�u�����\�[�g����
'
' ��P�L�[  �v��No      ����
' ��Q�L�[  �\�[�g�敪  ����
' ��R�L�[  �\���ݎ���  ����
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
'
Public Sub SortByProNo(oWorkSheet As Worksheet, sTableName As String)

    With oWorkSheet.ListObjects(sTableName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range(sTableName + "[�v��No]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range(sTableName + "[�\�[�g�敪]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range(sTableName + "[�\���ݎ���]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

