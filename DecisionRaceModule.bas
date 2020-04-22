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
    Call SheetActivate(S_ENTRY_SHEET_NAME)

    ' �o�͗p���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' �o�͗p���[�N�V�[�g
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ProNo�A�\�[�g�敪�A�\���ݎ��ԂŃ\�[�g
    Call SortByProNo(oWorkSheet, S_ENTRY_TABLE_NAME)
    
    ' �g�ݍ��킹�쐬
    Call SetHeatLaneOrder(oWorkSheet, S_ENTRY_TABLE_NAME)
    
    ' ���[�XNo, ���[���Ń\�[�g
    Call SortByRace(oWorkSheet, S_ENTRY_TABLE_NAME)
    
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
    
    ' �g�ݍ��킹���w�}�����t�ȑΉ��i���ؗp�̎b��j
    Dim bFlag As Boolean
    If GetRange("��").Value = "�w���}�X�^�[�Y���" Then
        bFlag = False
    Else
        bFlag = True
    End If
    
    ' �g�̍ŏ��l��
    Dim nMinNumberOfRace As Integer
    If GetRange("��").Value = "���{��I�茠���j���" Then
        nMinNumberOfRace = N_MIN_NUMBER_OF_RACE
    Else
        nMinNumberOfRace = N_MIN_NUMBER_OF_RACE2
    End If
    
    
    ' �G���g���[�ꗗ
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    Dim oProNo As Object
    
    ' �f�[�^���i�[
    Call ReadProNo(sTableName, oEntryList)

    Dim nRow As Integer             ' �J�����g�̍s�ԍ�
    
    Dim nNumOfProNo As Integer      ' �v���O����No���̐l��
    Dim nRaceNo As Integer          ' ���[�XNo
    Dim nHeat As Integer            ' �g�ԍ�
    Dim nHeats As Integer           ' �g��
    Dim nNumOfHeat As Integer       ' �g�̐l��
    Dim nNumOfHeats() As Integer    ' �g���̐l��
    Dim nNumOfSortType As Integer   ' �\�[�g�敪���̎c��l��
    Dim nRemNumber As Integer       ' �g�̎c��l��
    
    Dim nOrder As Integer           ' ProNo���̏����i�t���j
    Dim nStartLane As Integer       ' �g�̊J�n�ʒu
    Dim nCenterLane As Integer      ' �Z���^�[���[��
    Dim nMax As Integer             ' �g�̒���Lane���肷��l��
    Dim nNum As Integer             ' Lene���肷��l���̒��̏���
    Dim bAverage As Boolean         ' ���ϕ��������̗��p�L��
    Dim nMaxNum As Integer          ' ���ϕ��������̑��l��
    
    nRaceNo = 0
    
    ' �v���O����No��
    For Each nProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(nProNo)
        
        ' �v���O����No���̐l��
        nNumOfProNo = GetNumberOfProNo(oProNo)
        
        ' �g�����Z�o
        nHeats = GetHeats(nNumOfProNo)
        Call GetNumberOfHeat(nNumOfProNo, nHeats, nNumOfHeats, nMinNumberOfRace)
        
        ' ���ϕ��������𗘗p����P�[�X
        bAverage = False
        If GetRange("��").Value = "���{��I�茠���j���" And nHeats >= N_AVERAGE_DEC_RACE Then
            If VLookupArea(nProNo, "�I�茠��ڋ敪", "�\�I�^����") = "�\�I" Then
                bAverage = True
            End If
        End If
        
        ' ProNo���̑I��ʒu
        nOrder = 1
        
        ' �g��
        For nHeat = 1 To nHeats
            ' ���[�XNo���C���N�������g
            nRaceNo = nRaceNo + 10
            
            ' �g�̐l��
            nNumOfHeat = nNumOfHeats(nHeat - 1)
            ' �g�̎c��l��
            nRemNumber = nNumOfHeat
            
            ' ���ϕ�������
            If bAverage And nHeat = (nHeats - N_AVERAGE_DEC_RACE) + 1 Then
                nMaxNum = 0
                For i = nHeat To nHeat + N_AVERAGE_DEC_RACE - 1
                    nMaxNum = nMaxNum + nNumOfHeats(i - 1)
                Next i
                Call AverageMethod(nRaceNo, CInt(nHeat), nMaxNum, nOrder, oProNo, sTableName)
                Exit For
            End If
        
            ' �g�̊J�n�ʒu
            nStartLane = GetStartLane(nNumOfHeat, GetCenterLane(N_NUMBER_OF_RACE, N_MIN_LANE_OF_RACE), bFlag)
            
            ' �g�̐l�����c���Ă����
            While nRemNumber > 0
                ' �\�[�g�敪���̎c��l��
                nNumOfSortType = GetNumberOfSortType(nOrder, oProNo)
                If nNumOfSortType <= nRemNumber Then
                    ' Lane���肷��l��
                    nMax = nNumOfSortType
                Else
                    nMax = nRemNumber
                End If
                
                ' ���[�������肷��
                For nNum = 1 To nMax
                    ' �J�����g�s�ԍ�
                    nRow = GetProNoRow(nOrder, oProNo)
                
                    ' ���[�XNo�A�g�̏�����
                    Cells(nRow, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
                    Cells(nRow, Range(sTableName & "[�g]").Column).Value = nHeat
                
                    ' ���[�XNo�A�g�A���[�����L�q
                    nCenterLane = GetCenterLane(nMax, nStartLane, bFlag)
                    Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nCenterLane, nMax, nNum, bFlag)
                
                    ' ���Ԃ��C���N�������g
                    nOrder = nOrder + 1
                Next
            
                ' �J�n�ʒu��ύX
                nStartLane = nStartLane + nMax
            
                ' �c��l�������Z
                nRemNumber = nRemNumber - nMax
            Wend
        Next nHeat
    
    Next nProNo

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
        Set oRow = CreateObject("Scripting.Dictionary")
        oRow.Add "Row", cProNo.Row
        oRow.Add "SortType", Trim(Cells(cProNo.Row, Range(sTableName & "[�\�[�g�敪]").Column).Value)
        oProNo.Add oProNo.Count + 1, oRow
    Next
End Sub

'
' �v��No�A�\�[�g�敪�̌���
'
' oProNo        IN      �v��No���̃G���g���[�z��
'
Function GetNumberOfProNo(oProNo As Object)
    GetNumberOfProNo = oProNo.Count
End Function

'
' �v��No�̏��Ԃ̍s
'
' nIndex        IN      �����ԍ�
' oProNo        IN      �v��No���̃G���g���[�z��
'
Function GetProNoRow(nIndex As Integer, oProNo As Object)
    Dim oRow As Object
    Set oRow = oProNo.Item(nIndex)
    GetProNoRow = oRow.Item("Row")
End Function

'
' �J�n�ʒu���瓯��\�[�g�敪�̌���
'
' nIndex        IN      �J�n�ʒu
' oProNo        IN      �v��No���̃G���g���[�z��
'
Function GetNumberOfSortType(nIndex As Integer, oProNo As Object)
    
    GetNumberOfSortType = 1
    Dim oRow As Object
    Set oRow = oProNo.Item(nIndex)
    Dim sSortType As String
    sSortType = oRow.Item("SortType")
    
    For i = nIndex + 1 To oProNo.Count
        Set oRow = oProNo.Item(i)
        If oRow.Item("SortType") = sSortType Then
            GetNumberOfSortType = GetNumberOfSortType + 1
        Else
            Exit Function
        End If
    Next
End Function


'
' �g�l���z���ݒ�
'
' nTotalNum     IN      �v��No�̃G���g���[��
' nHeat         IN      �g��
' nNumberOfHeat() OUT   �g���̐l���z��
' nMinNumberOfRace IN   �g�̍ŏ��l��
'
Sub GetNumberOfHeat(nTotalNum As Integer, nHeats As Integer, nNumberOfHeat() As Integer, nMinNumberOfRace As Integer)
    
    ReDim nNumberOfHeat(nHeats - 1) As Integer
    
    ' �P�g�ڐl���Z�o
    nNumberOfHeat(0) = GetFirstHeatNumber(nTotalNum, nMinNumberOfRace)
    
    ' �Q�g�ڐl���Z�o
    If nHeats >= 2 Then
        nNumberOfHeat(1) = GetSecondHeatNumber(nTotalNum, nMinNumberOfRace)
    
        ' �R�g�ڈȍ~
        If nHeats > 2 Then
            For i = 2 To nHeats - 1
                nNumberOfHeat(i) = N_NUMBER_OF_RACE
            Next
        End If
    End If
End Sub

'
' �g���Z�o
'
' �g���̓��[�X�̑��l�����P���[�X�̐l��
'
' nTotalNum     IN      ���[�X�̑��l��
'
Function GetHeats(nTotalNum As Integer)

    GetHeats = Application.WorksheetFunction.RoundUp(nTotalNum / N_NUMBER_OF_RACE, 0)

End Function

'
' �P�g�ڐl���Z�o
'
' ���l�����ŏ��l���ȏア��ꍇ�A�ŏ��l���͂P�g�ڂɎc��
'
' nTotalNum     IN      ���[�X�̑��l��
' nMinNumberOfRace IN   �g�̍ŏ��l��
'
Function GetFirstHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer)

    If nTotalNum <= N_NUMBER_OF_RACE Then
        GetFirstHeatNumber = nTotalNum
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE = 0 Then
        GetFirstHeatNumber = N_NUMBER_OF_RACE
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE <= nMinNumberOfRace Then
        GetFirstHeatNumber = nMinNumberOfRace
    Else
        GetFirstHeatNumber = nTotalNum Mod N_NUMBER_OF_RACE
    End If

End Function

'
' �Q�g�ڐl���Z�o
'
' �P�g�ڂɉ񂷐l���ɂ���ĂQ�g�ڂ��ω�����
'
' nTotalNum     IN      ���[�X�̑��l��
' nMinNumberOfRace IN   �g�̍ŏ��l��
'
Function GetSecondHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer)

    If nTotalNum <= N_NUMBER_OF_RACE Then
        GetSecondHeatNumber = 0
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE = 0 Then
        GetSecondHeatNumber = N_NUMBER_OF_RACE
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE <= nMinNumberOfRace Then
        GetSecondHeatNumber = N_NUMBER_OF_RACE + (nTotalNum Mod N_NUMBER_OF_RACE - nMinNumberOfRace)
    Else
        GetSecondHeatNumber = N_NUMBER_OF_RACE
    End If

End Function

'
' �Z���^�[���[���Z�o
'
' nCount        IN      �l��
' nStart        IN      �J�n�ʒu
'
Function GetCenterLane(nCount As Integer, nStart As Integer, Optional bFlag As Boolean = True)
    If bFlag Then
        GetCenterLane = nStart + Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
    Else
        GetCenterLane = nStart + Application.WorksheetFunction.RoundDown((nCount) / 2, 0)
    End If
End Function

'
' �J�n�ʒu���Z�o
'
' nCount        IN      ���[�X�l��
' nCenterLane   IN      �Z���^�[���[��
'
Function GetStartLane(nCount As Integer, nCenterLane As Integer, Optional bFlag As Boolean = True)
    If bFlag Then
        GetStartLane = nCenterLane - Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
    Else
        GetStartLane = nCenterLane - Application.WorksheetFunction.RoundDown((nCount) / 2, 0)
    End If
End Function

'
' ���[������i�P�������j
'
' ���[���͋��Z�K���̒P�������ŕ��ׂ�
'
' nCenter       IN      �Z���^�[
' nMax          IN      �l��
' nOrder        IN      ����
' bFlag         In      True�F�ʏ�^False�F�t��
'
Function GetLane(nCenter As Integer, nMax As Integer, nOrder As Integer, Optional bFlag As Boolean = True)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    If bFlag Then
        GetLane = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        GetLane = nCenter + Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    End If
End Function

'
' ���[������(���ϕ�������)
'
' ���[���͋��Z�K���̕��ϕ��������ŕ��ׂ�
'
' nCenter       IN      �Z���^�[
' nMax          IN      �l��
' nOrder        IN      ����
'
Function GetLane2(nCenter As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = Application.WorksheetFunction.RoundUp((nMax - nOrder + 1) / N_AVERAGE_DEC_RACE, 0)
    GetLane2 = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
            * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
End Function

'
' �g�ԍ�����(���ϕ�������)
'
' �v��No�̑��l���Ə��Ԃ���g�ԍ����Z�o����
'
' �g�����Z�o����
'
' nHeats        IN      �g��
' nMax          IN      �l��
' nOrder        IN      ����
'
Function GetOrderHeat(nHeats As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    GetOrderHeat = nHeats - (nNum - 1) Mod nHeats
End Function

'
' ���[������(���ϕ�������)
'
' ���[���͋��Z�K���̕��ϕ��������ŕ��ׂ�
'
' nStartRaceNo  IN/OUT  �J�n��RaceNo
' nStartHeat    IN      �J�n�̑g�ԍ�
' nMaxNum       IN      �l��
' nOrder        IN      ����
' oProNo        IN      ����
' sTableName    IN      �e�[�u����
'
Sub AverageMethod(nStartRaceNo As Integer, nStartHeat As Integer, nMaxNum As Integer, _
nOrder As Integer, oProNo As Object, sTableName As String)
    
    Dim nCenterLane As Integer
    Dim nRow As Integer
    Dim nRaceNo As Integer
    Dim nHeat As Integer
    
    Dim nNum As Integer

    ' �g�̃Z���^�[
    nCenterLane = GetCenterLane(N_NUMBER_OF_RACE, N_MIN_LANE_OF_RACE)
    
    ' �g�̐l�����c���Ă����
    For nNum = 1 To nMaxNum
        
        ' �J�����g�s�ԍ�
        nRow = GetProNoRow(nOrder, oProNo)
        
        ' ���[�XNo
        nRaceNo = (GetOrderHeat(N_AVERAGE_DEC_RACE, nMaxNum, nNum) + (nStartRaceNo / 10 - 1)) * 10
        ' �g�ԍ�
        nHeat = GetOrderHeat(N_AVERAGE_DEC_RACE, nMaxNum, nNum) + (nStartHeat - 1)
    
        ' ���[�XNo�A�g�̏�����
        Cells(nRow, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
        Cells(nRow, Range(sTableName & "[�g]").Column).Value = nHeat
    
        ' ���[�XNo�A�g�A���[�����L�q
        Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane2(nCenterLane, nMaxNum, nNum)
    
        ' ���Ԃ��C���N�������g
        nOrder = nOrder + 1
    Next

    nStartRaceNo = (nStartRaceNo / 10 - 1 + N_AVERAGE_DEC_RACE) * 10

End Sub


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
    Call SheetActivate(S_ENTRY_SHEET_NAME)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' �ă\�[�g
    Call SortByRace(oWorkSheet, S_ENTRY_TABLE_NAME)
    
    ' ���[�X�ԍ��C��
    Call SetRaceNo(oWorkSheet, S_ENTRY_TABLE_NAME)

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

