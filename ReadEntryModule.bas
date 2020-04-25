Attribute VB_Name = "ReadEntryModule"
'
' �G���g���[�t�@�C���ꗗ�̓ǂݍ���
'
' ����̃t�H���_���w�肵�āA���ɂ���G���g���[�t�@�C����
' ���ׂēǂݍ��݈ꗗ�V�[�g�ɏo�͂���B
'
Sub �G���g���[�Ǎ���()
    ' �C�x���g������}��
    Call EventChange(False)

    ' �G�N�Z���V�[�g��I��
    Call SheetActivate(S_ENTRY_SHEET_NAME)

    ' �o�͗p���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' �o�͗p���[�N�V�[�g
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    ' �G���g���[�ꗗ�Ǎ��p�z��
    Dim oGameList As Object
    Set oGameList = CreateObject("Scripting.Dictionary")

    ' �G���g���[�t�@�C���Ǎ���
    Call ReadEntryFiles(oGameList)

    ' �G���g���[�V�[�g�̏�������
    Call WriteEntrySheet(oWorkSheet, S_ENTRY_TABLE_NAME, oGameList)
    
    ' ProNo�A�\�[�g�敪�A�\���ݎ��ԂŃ\�[�g
    Call SortByProNo(oWorkSheet, S_ENTRY_TABLE_NAME)

    ' �V�[�g��ۑ�
    oWorkBook.Save

End Sub

'
' �G���g���[�t�@�C���ꗗ�̓ǂݍ���
'
' �t�H���_���w�肵�āA���̒��Ɋ܂܂��G���g���[�V�[�g�i*.xlsx�j�����ׂĉr�ݍ���
'
' oGameList     OUT     �G���g���[�ꗗ
'
Sub ReadEntryFiles(ByRef oGameList As Object)

    ' �t�@�C���ꗗ���擾
    '
    Dim sPathName As String
    sPathName = SelectDir()
    Dim FileList As Collection
    Set FileList = GetFiles(sPathName, "\*.xlsx")

    Dim nMax As Integer
    nMax = FileList.Count
    Dim nCount As Integer
    nCount = 0

    '
    ' �t�@�C�����ɏ�������
    '
    For Each vFile In FileList
        
        ' �^�C�g���C��
        nCount = nCount + 1
        Call SetTitleMenu("�v���O�����Ǎ���: " & Str(nCount) & "/" & Str(nMax))
        
        '
        ' �t�@�C�����J���i�ǂݎ���p�j
        '
        Set SubBook = Workbooks.Open(Filename:=sPathName + "\" + vFile, ReadOnly:=True)
        Worksheets("�L���[").Activate

        ' �G���g���[�ꗗ�̓Ǎ���
        Call ReadEntryFile(oGameList)
    
        ' �x���Ȃ��Ńt�@�C�������i�ۑ����Ȃ��j
        Application.DisplayAlerts = False
        SubBook.Close
        Application.DisplayAlerts = True
    Next
    
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
' ������ږ��FRange("��ږ�")
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
Sub ReadEntryFile(ByRef oGameList As Object)

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

    Dim nIdx As Integer

    ' �l�ԍ��͈͂����ׂēǂݍ���
    For Each oCell In GetRange("�I��ԍ�")
        
        If oCell.MergeCells Then
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

                oEntry.Add "����", Cells(oCell.Row, GetRange("�I�萫��").Column).Value + "�q"
                oEntry.Add "�t���K�i", ReplaceName(Cells(oCell.Row, GetRange("�I��t���K�i").Column).Value)

                If Range("��").Value = "���{��I�茠���j���" Then
                    
                    oEntry.Add "�I�薼", ReplaceName(Cells(oCell.Row + 1, GetRange("�I�薼").Column).Value)
                    oEntry.Add "�敪", Cells(oCell.Row + 1, GetRange("�I��敪").Column).Value
                    nIdx = 1
                
                ElseIf Range("��").Value = "���{��s���̈���" Then
                    
                    oEntry.Add "�I�薼", ReplaceName(Cells(oCell.Row + 2, GetRange("�I�薼").Column).Value)
                    oEntry.Add "�w�Z��", Trim(Cells(oCell.Row + 4, GetRange("�I��w�Z��").Column).Value)
                    If Cells(oCell.Row, GetRange("�I��敪").Column).Value <> "" Then
                        oEntry.Add "�敪", Cells(oCell.Row, GetRange("�I��敪").Column).Value
                    Else
                        oEntry.Add "�敪", "�N��敪"
                    End If
                    oEntry.Add "�N��", Cells(oCell.Row + 3, GetRange("�I��N��").Column).Value
                    nIdx = 3
                
                ElseIf Range("��").Value = "���{��}�X�^�[�Y���" Then
                
                    oEntry.Add "�I�薼", ReplaceName(Cells(oCell.Row + 1, GetRange("�I�薼").Column).Value)
                    oEntry.Add "�N��", Cells(oCell.Row, GetRange("�I��N��").Column).Value
                    nIdx = 1
                
                Else
            
                    oEntry.Add "�I�薼", ReplaceName(Cells(oCell.Row + 1, GetRange("�I�薼").Column).Value)
                    oEntry.Add "�敪", Cells(oCell.Row, GetRange("�I��w�N").Column).Value
                    nIdx = 1
                
                End If
                
                ' �P�s��
                Call ReadEntryLine(1, oCell.Row, oEntry)
                Call CheckEntry(oCell.Row, oEntry, 1)
    
                ' �Q�s��
                Call ReadEntryLine(2, oCell.Row + nIdx, oEntry)
                Call CheckEntry(oCell.Row + nIdx, oEntry, 2)
                If oEntry.Item("�I�薼") <> "" Then
                    oEntryList.Add nNum, oEntry
                End If
            End If
        End If
    Next oCell

    ' �����[��ڔԍ��͈͂����ׂēǂݍ���
    Dim nRelayNum As Integer
    nRelayNum = 0
    Dim oRelayEntry As Object
    Set oRelayEntry = Nothing
    For Each oCell In GetRange("�����[���")
        ' �l���ݒ肳��Ă���ꍇ�͓ǂݍ���
        If oCell.Value <> "" Then
            ' �����[�̃G���g���[���X�g
            If oRelayEntry Is Nothing Then
                Set oRelayEntry = CreateObject("Scripting.Dictionary")
                oEntryList.Add nNum, oRelayEntry
            End If

            nRelayNum = nRelayNum + 1
            Call ReadRelayEntryLine(nRelayNum, oCell.Row, oRelayEntry)
        End If
    Next oCell

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
Sub ReadEntryLine(nNum As Integer, nRow As Integer, oEntry As Object)
    
    Dim oLines As Object
    Dim sStyle As String
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    ' �ԍ��͈͂����ׂēǂݍ���
    For Each oCell In GetRange("��ږ�")
        If oCell.Value <> "" Then
            sStyle = oCell.Value
        End If
        ' ��ڑI������܂��͑S�p�󔒈ȊO�̏ꍇ�͑I�����ꂽ���̂Ƃ���
        If Trim(Cells(nRow, oCell.Column).Value) <> "" And Cells(nRow, oCell.Column).Value <> "�@" Then
            Set oLines = CreateObject("Scripting.Dictionary")
            oEntry.Add nNum, oLines
            
            oLines.Add "��ڔԍ�", VLookupArea(Cells(nRow, oCell.Column).Value, "��ڔԍ��敪", "��ڔԍ�")
            oLines.Add "��ڋ敪", VLookupArea(Cells(nRow, oCell.Column).Value, "��ڔԍ��敪", "��ڋ敪")
            oLines.Add "��ږ�", ReplaceStyle(sStyle)
            oLines.Add "����", ReplaceDistance(Cells(GetRange("��ڋ���").Row, oCell.Column).Value)
            nMin = Cells(nRow, GetRange("�I�蕪").Column).Value
            nSec = Cells(nRow, GetRange("�I��b").Column).Value
            nMil = Cells(nRow, GetRange("�I��~���b").Column).Value
            oLines.Add "�\���ݎ���", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
            Exit Sub
        End If
    Next oCell
End Sub

'
' �G���g���[�̎�ڔԍ��������������m�F
'
' nRow          IN      �s�ԍ�
' oEntry        IN      ��ڍs
' nNum          IN      �l�G���g���[�s(1,2)
'
Sub CheckEntry(nRow As Integer, oEntry As Object, nNum As Integer)
    
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
    
    If sGender <> oEntry.Item("����") Or sDistance <> oLines.Item("����") Or sStyle <> oLines.Item("��ږ�") Then
        MsgBox CStr(nRow) & "�s�ځF��ڔԍ�������������܂���B�F" & oLines.Item("��ڔԍ�")
        End
    End If

End Sub



'
' ��ږ��̂̒u��
'
' sStyle        IN      ���
'
Function ReplaceStyle(sStyle)
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
Function ReplaceDistance(sDistance)
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
Function ReplaceName(sName)
    
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
' �G���g���[�t�@�C���̃����[��ڍs�ǂݍ���
'
' nNum          IN    �����[�ԍ�
' nRow          IN    �J�����g�s��
' oRelayEntry   I/O   ��ڍs
'
Sub ReadRelayEntryLine(nNum As Integer, nRow As Integer, oRelayEntry As Object)
    
    Dim oLines As Object
    Dim nMin As Integer
    Dim nSec As Integer
    Dim nMil As Integer
    
    If Cells(nRow, GetRange("�����[���").Column).Value <> "" Then
        Set oRelayLines = CreateObject("Scripting.Dictionary")
        oRelayEntry.Add "L" + Str(nNum), oRelayLines

        oRelayLines.Add "��ڔԍ�", Cells(nRow, GetRange("�����[���").Column).Value
        If IsNameExists("�����[�敪") Then
            oRelayLines.Add "�敪", Cells(nRow, GetRange("�����[�敪").Column).Value
        End If
        nMin = Cells(nRow, GetRange("�����[��").Column).Value
        nSec = Cells(nRow, GetRange("�����[�b").Column).Value
        nMil = Cells(nRow, GetRange("�����[�~���b").Column).Value
        oRelayLines.Add "�\���ݎ���", CLng(nMin * CLng(10000) + nSec * 100 + nMil)
    End If
End Sub

'
' �\���݂��V�[�g�ɏo��
'
' oWorkBook     IN     �o�͂���V�[�g�̂���G�N�Z��WorkBook
' sTable        IN     �e�[�u����
' oTeamList     IN     �ǂݍ��񂾃`�[���\���݈ꗗ
'
Sub WriteEntrySheet(oWorkSheet As Worksheet, sTable As String, oGameList As Object)
    
    ' �G���g���[�e�[�u����������
    Call DeleteTable(oWorkSheet, sTable)
    
    ' �G���g���[�ꗗ�̏o��
    Call WriteTeamEntries(oWorkSheet, sTable, oGameList)

End Sub

'
' �\���݂��V�[�g�ɏo��
'
' oWorkBook     IN     �o�͂���V�[�g�̂���G�N�Z��WorkBook
' sTable        IN     �e�[�u����
' oTeamList     IN     �ǂݍ��񂾃`�[���\���݈ꗗ
'
Sub WriteTeamEntries(oWorkSheet As Worksheet, sTable As String, oGameList As Object)

    oWorkSheet.Activate

    Dim nPersonNo As Integer
    Dim nTeamNo As Integer
    nTeamNo = 1
    
    Dim nRow As Integer
    nRow = 1
    For Each sGame In oGameList
        Dim oTeamList As Object
        Set oTeamList = oGameList.Item(sGame)
        For Each sTeam In oTeamList
            Dim oEntryList As Object
            Set oEntryList = oTeamList.Item(sTeam)
            
            Dim oLine As Object
            For Each sNum In oEntryList
                Dim oEntry As Object
                Set oEntry = oEntryList.Item(sNum)
                nPersonNo = nTeamNo * 100 + Int(sNum)
                
                If oEntry.Exists("�I�薼") Then
                    ' �l
                    For i = 1 To 2
                        If Not IsEmpty(oEntry.Item(i)) Then
                            nRow = nRow + 1
                            Set oLine = oEntry.Item(i)
                            Call WriteLine(sTable, nRow, nPersonNo, sGame, sTeam, oEntry, oLine)
                        End If
                    Next i
                Else
                    ' �����[
                    Dim sKey As String
                    For i = 1 To 24
                        sKey = "L" & Str(i)
                        If oEntry.Exists(sKey) Then
                            nRow = nRow + 1
                            Set oLine = oEntry.Item(sKey)
                            Call WriteRelayLine(sTable, nRow, nTeamNo, sGame, sTeam, oEntry, oLine)
                        End If
                    Next i
                End If
            Next
            ' �`�[���ԍ����C���N�������g
            nTeamNo = nTeamNo + 1
        Next
    Next
End Sub

'
' �G���g���[�e�[�u����������
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
'
Sub DeleteTable(oWorkSheet As Worksheet, sTableName As String)
    Dim myTable As ListObject
    Set myTable = oWorkSheet.ListObjects(sTableName)
    If Not (myTable.DataBodyRange Is Nothing) Then
        myTable.DataBodyRange.Delete
    End If
End Sub

'
' �\���ݍs���o��
'
' sTable        IN      �e�[�u����
' nRow          IN      �o�͍s�ԍ�
' nPersonNo     IN      �I��ԍ�
' sGame         IN      ��
' sTeam         IN      �`�[����
' oEntry        IN      �G���g���[
' oLine         IN      ��ځA�\���ݎ���
'
Sub WriteLine( _
    sTable As String, _
    nRow As Integer, _
    nPersonNo As Integer, _
    sGame As Variant, _
    sTeam As Variant, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[�lNo]").Column).Value = nPersonNo
    Cells(nRow, Range(sTable & "[�v��No]").Column).Value = oLine.Item("��ڔԍ�")
    Cells(nRow, Range(sTable & "[�`�[����]").Column).Value = CStr(sTeam)
    Cells(nRow, Range(sTable & "[�I�薼]").Column).Value = oEntry.Item("�I�薼")
    Cells(nRow, Range(sTable & "[�t���K�i]").Column).Value = oEntry.Item("�t���K�i")
    Cells(nRow, Range(sTable & "[����]").Column).Value = oEntry.Item("����")
    Cells(nRow, Range(sTable & "[����]").Column).Value = oLine.Item("����")
    Cells(nRow, Range(sTable & "[��ږ�]").Column).Value = oLine.Item("��ږ�")
    Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).Value = oLine.Item("�\���ݎ���")
    If oLine.Item("�\���ݎ���") >= 10000 Then
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    Dim nColumn As Integer
    If CStr(sGame) = "���{��I�茠���j���" Then
    
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oEntry.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    
    ElseIf CStr(sGame) = "���{��s���̈���" Then
    
        Cells(nRow, Range(sTable & "[�w�Z��]").Column).Value = oEntry.Item("�w�Z��")
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = oEntry.Item("�N��")
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = oEntry.Item("�敪")
        
        ' �l�N��敪
        If oEntry.Item("�敪") = "�N��敪" Then
            nColumn = VLookupArea(oLine.Item("��ڔԍ�"), "�s����ڋ敪", "�^�C�v")
            sType = Application.WorksheetFunction.VLookup(oEntry.Item("�N��"), GetRange("�s���N��敪"), nColumn, False)
            Cells(nRow, Range(sTable & "[�敪]").Column).Value = sType
            If sType = "���" Then
                Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = "20"
            Else
                Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = Left(sType, 2)
            End If
        ' �l����
        Else
            Cells(nRow, Range(sTable & "[�敪]").Column).Value = oEntry.Item("�敪")
            Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
        End If
    
    ElseIf CStr(sGame) = "���{��}�X�^�[�Y���" Then
        
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = oEntry.Item("�N��")
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�N��"), "�w�}�N��敪", "M�N��敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�N��"), "�w�}�N��敪", "M�N��敪")

    ElseIf CStr(sGame) = "���{��s�w�����j���Z���" Then
        
        Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = oLine.Item("��ڋ敪")
        Cells(nRow, Range(sTable & "[�N��]").Column).Value = ""
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = _
            VLookupArea(oEntry.Item("�敪"), "�w�}�w�N�\��", "�w�N�\��")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    
    End If
    
End Sub

'
' �����[�\���ݍs���o��
'
' sTable        IN      �e�[�u����
' nRow          IN      �o�͍s�ԍ�
' nTeamNo       IN      �`�[���ԍ�
' sGame         IN      ��
' sTeam         IN      �`�[����
' oEntry        IN      �G���g���[
' oLine         IN      ��ځA�\���ݎ���
'
Sub WriteRelayLine( _
    sTable As String, _
    nRow As Integer, _
    nTeamNo As Integer, _
    sGame As Variant, _
    sTeam As Variant, _
    oEntry As Object, _
    oLine As Object _
)

    Cells(nRow, Range(sTable & "[No.]").Column).Value = nRow + 1
    Cells(nRow, Range(sTable & "[�lNo]").Column).Value = nTeamNo
    Cells(nRow, Range(sTable & "[�`�[����]").Column).Value = CStr(sTeam)
    
    Cells(nRow, Range(sTable & "[�v��No]").Column).Value = oLine.Item("��ڔԍ�")
    
    Dim sMasterName As String
    sMasterName = GetMaster(CStr(sGame))
    
    Cells(nRow, Range(sTable & "[��ڋ敪]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "��ڋ敪")
    
    Cells(nRow, Range(sTable & "[����]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "����")
    
    Cells(nRow, Range(sTable & "[����]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "����")
    
    Cells(nRow, Range(sTable & "[��ږ�]").Column).Value = _
        VLookupArea(oLine.Item("��ڔԍ�"), sMasterName, "���")

    Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).Value = oLine.Item("�\���ݎ���")
    If oLine.Item("�\���ݎ���") >= 10000 Then
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = "#"":""##"".""##"
    Else
        Cells(nRow, Range(sTable & "[�\���ݎ���]").Column).NumberFormatLocal = """ :""##"".""##"
    End If
    
    If CStr(sGame) = "���{��I�茠���j���" Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oLine.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
    
    ElseIf CStr(sGame) = "���{��s���̈���" Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oLine.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = oLine.Item("�敪")
    
    ElseIf CStr(sGame) = "���{��}�X�^�[�Y���" Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = oLine.Item("�敪")
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = oLine.Item("�敪")
    
    ElseIf CStr(sGame) = "���{��s�w�����j���Z���" Then
        Cells(nRow, Range(sTable & "[�敪]").Column).Value = "���w"
        Cells(nRow, Range(sTable & "[�\�[�g�敪]").Column).Value = ""
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
' sTableName    OUT     �e�[�u����
'
Public Sub SortByProNo(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate

    With ActiveSheet.ListObjects(sTableName).Sort
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

