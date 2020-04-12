Attribute VB_Name = "DecisionRaceModule"
'
' �g�ݍ��킹����
'
' ���[�XNo�Ƒg�A���[����ݒ肷��
'
Sub �g�ݍ��킹����()
    ' �C�x���g������}��
    Call EventChange(False)

    ' �G�N�Z���V�[�g��I��
    Call SheetActivate(sEntrySheetName)

    ' �o�͗p���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' �o�͗p���[�N�V�[�g
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ProNo�A�\�[�g�敪�A�\���ݎ��ԂŃ\�[�g
    Call SortByProNo(oWorkSheet, sEntryTableName)
    
    ' �g�ݍ��킹�쐬
    Call SetHeatLaneOrder(oWorkSheet, sEntryTableName)
    
    ' ���[�XNo, ���[���Ń\�[�g
    Call SortByRace(oWorkSheet, sEntryTableName)
    
    ' �C�x���g�����𔭐�
    Call EventChange(True)

    ' �V�[�g��ۑ�
    oWorkBook.Save
End Sub

'
' �g�ݍ��킹�쐬
'
' �G���g���[�ꗗ��Ǎ��݁A�g�ݍ��킹���쐬����
' �G���g���[�ꗗ��ProNo�A�\�[�g�敪�A�\���ݎ��ԂŃ\�[�g����Ă���O��
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
'
Sub SetHeatLaneOrder(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate
    
    ' �G���g���[�ꗗ
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^���i�[
    Call ReadProNo(sTableName, oEntryList)

    Dim nRaceNo As Integer      ' ���[�XNo
    Dim nCount As Integer       ' �v��No���̌���
    Dim nRow As Integer         ' �J�����g�̍s�ԍ�
    Dim nHeat As Integer        ' �g�ԍ�
    Dim nPreHeat As Integer     ' ����܂ł̑g�ԍ�
    Dim nNum As Integer
    nTotalHeat = 0
    ' �v���O����No��
    For Each nProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(nProNo)
        nCount = oProNo.Count
        nPreHeat = 0
        
        ' �g��
        For Each nOrder In oProNo.Keys
            ' �J�����g�s�ԍ�
            nRow = oProNo.Item(nOrder)
            ' �g�ԍ�������
            nHeat = GetOrderHeat(nCount, Int(nOrder))
            If nPreHeat <> nHeat Then
                ' �g�ԍ����ς�����ꍇ
                nNum = 1
                nPreHeat = nHeat
                nRaceNo = nRaceNo + 1   ' ���[�XNo���C���N�������g
            Else
                nNum = nNum + 1
            End If
            ' ���[�XNo�A�g�̏�����
            Cells(nRow, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
            Cells(nRow, Range(sTableName & "[�g]").Column).Value = nHeat
                      
            ' ���{��I�茠���j���
            If GetRange("��").Value = "���{��I�茠���j���" Then
                Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
            ' ���{��s���̈���
            ElseIf GetRange("��").Value = "���{��s���̈���" Then
                Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
            Else
                ' ���{��}�X�^�[�Y
                If Cells(nRow, Range(sTableName & "[�\�[�g�敪]").Column).Value <> "" Then
                    Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane2(nHeat, oProNo.Count, nNum, False)
                ' �w��
                Else
                    Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nHeat, oProNo.Count, nNum, False)
                End If
            End If
        Next
    Next

End Sub

'
' �v��No���L�[�Ƀf�[�^�i�[
'
' �v��No���ɒʔԂ�U���Ă��̍s�ԍ����i�[����
'
' sTableName    IN      �e�[�u����
' oEntryList    OUT     �G���g���[�ԍ�
'�@���v��No
'�@�@���ʔԁF�s�ԍ�
'
Sub ReadProNo(sTableName As String, oEntryList As Object)
    ' �f�[�^���i�[
    Dim oProNo As Object
    For Each cProNo In Range(sTableName & "[�v��No]")
        If Not oEntryList.Exists(cProNo.Value) Then
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add cProNo.Value, oProNo
        End If
        oProNo.Add oProNo.Count + 1, cProNo.Row
    Next
End Sub

'
' �g�ԍ�����
'
' �v��No�̑��l���Ə��Ԃ���g�ԍ����Z�o����
'
' �g���A�P�g�ڂ̐l���A�Q�g�ڂ̐l�����Z�o���邱�Ƃ�
' �g�ԍ����Z�o�\
'
' nTotalNum     IN      �v��No�̑��l��
' nOrder        IN      ����
'
Function GetOrderHeat(nTotalNum As Integer, nOrder As Integer)

    ' �g�����Z�o
    Dim nHeats As Integer
    nHeats = GetHeats(nTotalNum)
    
    ' �P�g�ڐl���Z�o
    Dim nFirstHeatNumber As Integer
    nFirstHeatNumber = GetFirstHeatNumber(nTotalNum)
    
    ' �Q�g�ڐl���Z�o
    Dim nSecondHeatNumber As Integer
    nSecondHeatNumber = GetSecondHeatNumber(nTotalNum)
    
    ' �P�g�ڂ̏ꍇ
    If nOrder <= nFirstHeatNumber Then
        GetOrderHeat = 1
    ' �Q�g�ڂ̏ꍇ
    ElseIf nOrder <= nFirstHeatNumber + nSecondHeatNumber Then
        GetOrderHeat = 2
    ' �R�g�ڈȍ~�̏ꍇ
    Else
        GetOrderHeat = GetHeats(nOrder - (nFirstHeatNumber + nSecondHeatNumber)) + 2
    End If

End Function
'
' �g���Z�o
'
' �g���̓��[�X�̑��l�����P���[�X�̐l��
'
' nTotalNum     IN      ���[�X�̑��l��
'
Function GetHeats(nTotalNum As Integer)

    GetHeats = Application.WorksheetFunction.RoundUp(nTotalNum / nNumberOfRace, 0)

End Function

'
' �P�g�ڐl���Z�o
'
' ���l�����R���ȏア��ꍇ�A�Œ�R���͂P�g�ڂɎc��
'
' nTotalNum     IN      ���[�X�̑��l��
'
Function GetFirstHeatNumber(nTotalNum As Integer)

    If nTotalNum <= nNumberOfRace Then
        GetFirstHeatNumber = nTotalNum
    ElseIf nTotalNum Mod nNumberOfRace = 0 Then
        GetFirstHeatNumber = nNumberOfRace
    ElseIf nTotalNum Mod nNumberOfRace <= nMinLaneOfRace Then
        GetFirstHeatNumber = nMinLaneOfRace
    Else
        GetFirstHeatNumber = nTotalNum Mod nNumberOfRace
    End If

End Function

'
' �Q�g�ڐl���Z�o
'
' �P�g�ڂɉ񂷐l���ɂ���ĂQ�g�ڂ��ω�����
'
' nTotalNum     IN      ���[�X�̑��l��
'
Function GetSecondHeatNumber(nTotalNum As Integer)

    If nTotalNum <= nNumberOfRace Then
        GetSecondHeatNumber = 0
    ElseIf nTotalNum Mod nNumberOfRace = 0 Then
        GetSecondHeatNumber = nNumberOfRace
    ElseIf nTotalNum Mod nNumberOfRace <= nMinLaneOfRace Then
        GetSecondHeatNumber = nNumberOfRace + (nTotalNum Mod nNumberOfRace - nMinLaneOfRace)
    Else
        GetSecondHeatNumber = nNumberOfRace
    End If

End Function

'
' ���[������i�w���p�j
'
' ���[���͋��Z�K���̒P�������ŕ��ׂ�
'
' nHeat         IN      �g�ԍ�
' nTotalNum     IN      �v��No�̑��l��
' nOrder        IN      ����
' bFlag         In      True�F�ʏ�^False�F�t��
'
Function GetLane(nHeat As Integer, nTotalNum As Integer, nOrder As Integer, Optional bFlag As Boolean = True)

    Dim nMax As Integer
    Dim nNum As Integer
    
    If nHeat = 1 Then
        nMax = GetFirstHeatNumber(nTotalNum)
    ElseIf nHeat = 2 Then
        nMax = GetSecondHeatNumber(nTotalNum)
    Else
        nMax = nNumberOfRace
    End If
    
    nNum = nMax - nOrder + 1
    
    If bFlag Then
        ' 4->5->3->6->2->7->1
        GetLane = nCenterLane - Application.WorksheetFunction.Power(-1, nNum - 1) _
                    * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        ' 4->3->5->2->6->1->7(�Ȃ����w���͂����炾����)
        GetLane = nCenterLane + Application.WorksheetFunction.Power(-1, nNum - 1) _
                    * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    End If

End Function

'
' ���[������i�}�X�^�[�Y�p�j
'
' ���[���͏����ŕ��ׂ�
'
' nHeat         IN      �g�ԍ�
' nTotalNum     IN      �v��No�̑��l��
' nOrder        IN      ����
' bFlag         In      True�F�ʏ�^False�F�t��
'
Function GetLane2(nHeat As Integer, nTotalNum As Integer, nOrder As Integer, Optional bFlag As Boolean = True)

    Dim nMax As Integer
    
    If nHeat = 1 Then
        nMax = GetFirstHeatNumber(nTotalNum)
    ElseIf nHeat = 2 Then
        nMax = GetSecondHeatNumber(nTotalNum)
    Else
        nMax = nNumberOfRace
    End If
    
    If bFlag Then
        ' 4->5->3->6->2->7->1
        GetLane2 = nCenterLane + nOrder - Application.WorksheetFunction.RoundUp(nMax / 2, 0)
    Else
        ' 4->3->5->2->6->1->7(�Ȃ����w���͂����炾����)
        GetLane2 = nCenterLane + Application.WorksheetFunction.RoundUp(nMax / 2, 0) - (nMax - nOrder) - 1
    End If
End Function


'
' ���[�X�ԍ��C��
'
'
Sub ���[�X�ԍ��C��()
    ' �C�x���g������}��
    Call EventChange(False)

    ' �o�͗p���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook

    ' �o�͗p�V�[�g
    Call SheetActivate(sEntrySheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' �ă\�[�g
    Call SortByRace(oWorkSheet, sEntryTableName)
    
    ' ���[�X�ԍ��C��
    Call SetRaceNo(oWorkSheet, sEntryTableName)

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' ���[�X�ԍ��C��
'
' oWorkSheet    IN  ���[�N�V�[�g
' sTableName    IN  �e�[�u����
'
Sub SetRaceNo(oWorkSheet As Worksheet, sTableName As String)
    
    ' ProNo�A�g�̏d���`�F�b�N
    Dim oEntryList As Object
    Call ReadEntrySheet(sTableName, oEntryList)
    
    ' �G���g���[�ꗗ
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^���i�[�����[�X�ԍ��̃`�F�b�N
    Call ReadRace(sTableName, oEntryList)
    
    Dim nRaceNo As Integer
    nRaceNo = 1
    
    Dim nCurrentRaceNo As Integer
    nCurrentRaceNo = 0

    ' ���[�X�ԍ��C��
    Dim nNewNo As Integer
    nNewNo = 1
    Dim nRow As Integer
    For Each vRaceNo In oEntryList.Keys
        Set oRaceNo = oEntryList.Item(vRaceNo)
        
        ' ���[��
        For Each vLane In oRaceNo.Keys
            nRow = oRaceNo.Item(vLane)
            
            ' ���[�X�ԍ�
            Cells(nRow, Range(sTableName & "[���[�XNo]").Column).Value = nNewNo
        Next
        nNewNo = nNewNo + 1
    Next

End Sub

'
' ���[�X�ԍ��̓Ǎ���
'
' �Ǎ��݂Ȃ��烌�[�X�ԍ����̃��[���d���`�F�b�N���s��
'
' sTableName    IN      �e�[�u����
' oEntryList    I/O     �G���g���[�ꗗ
'
Sub ReadRace(sTableName As String, oEntryList As Object)
    Dim nLane As Integer
    Dim oRaceNo As Object
    For Each cRaceNo In Range(sTableName & "[���[�XNo]")
        ' ���݂��Ȃ����[�XNo�̏ꍇ��
        If Not oEntryList.Exists(cRaceNo.Value) Then
            ' �G���g���[�ꗗ�ɓo�^����
            Set oRaceNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add cRaceNo.Value, oRaceNo
        End If
        ' ���[���ԍ����擾
        nLane = cRaceNo.Offset(0, Range(sTableName & "[���[��]").Column - Range(sTableName & "[���[�XNo]").Column).Value
        ' ���[�X�ԍ��ɑ΂��郌�[���̏d���`�F�b�N
        If oRaceNo.Exists(nLane) Then
            MsgBox "���[�XNo�F" & Str(cRaceNo.Value) & vbCrLf & _
                    "���[���@�F" & Str(nLane) & vbCrLf & _
                    "���d�����Ă��܂��B"
            Range(sTableName).Parent.Activate
            Range(Cells(cRaceNo.Row, Range(sTableName & "[���[�XNo]").Column), _
                    Cells(cRaceNo.Row, Range(sTableName & "[���[��]").Column)).Select
            cRaceNo.Activate
            End
        Else
            oRaceNo.Add nLane, cRaceNo.Row
        End If
    Next cRaceNo
End Sub

'
' ���[�XNo�A�g�Ń\�[�g����
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
'
Sub SortByRace(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate

    With ActiveSheet.ListObjects(sTableName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range(sTableName & "[���[�XNo]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range(sTableName & "[���[��]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

