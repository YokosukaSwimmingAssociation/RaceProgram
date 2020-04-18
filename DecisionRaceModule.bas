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
    
    ' �g�ݍ��킹���w�}�����t�ȑΉ��i���ؗp�̎b��j
    Dim bFlag As Boolean
    If GetRange("��").Value = "�w���}�X�^�[�Y���" Then
        bFlag = False
    Else
        bFlag = True
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
    
    Dim nOrder As Integer           ' ProNo���̏����i�t���j
    Dim nStartLane As Integer       ' �g�̊J�n�ʒu
    Dim nCenterLane As Integer      ' �Z���^�[���[��
    Dim nMax As Integer             ' �g�̒���Lane���肷��l��
    Dim nNum As Integer             ' Lene���肷��l���̒��̏���
    
    nRaceNo = 0
    
    ' �v���O����No��
    For Each nProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(nProNo)
        
        ' �v���O����No���̐l��
        nNumOfProNo = GetNumberOfProNo(oProNo)
        
        ' �g�����Z�o
        nHeats = GetHeats(nNumOfProNo)
        Call GetNumberOfHeat(nNumOfProNo, nHeats, nNumOfHeats)
        
        ' ProNo���̑I��ʒu
        nOrder = 1
        
        ' �g��
        For nHeat = 1 To nHeats
            ' ���[�XNo���C���N�������g
            nRaceNo = nRaceNo + 1
            
            ' �g�̐l��
            nNumOfHeat = nNumOfHeats(nHeat - 1)
            ' �g�̎c��l��
            nRemNumber = nNumOfHeat
        
            ' �g�̊J�n�ʒu
            nStartLane = GetStartLane(nNumOfHeat)
            
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
                    nCenterLane = GetCenterLane(nMax, nStartLane)
                    Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane3(nCenterLane, nMax, nNum, bFlag)
                
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
' nRow          IN      �s��
' oProNo        IN      �v��No���̃G���g���[�z��
' nRaceNo       IN      ���[�XNo
' nHeat         IN      �g�ԍ�
' nNum          IN      ����
' nStartLane    IN      �J�n�ʒu
'
Sub SetHeatLaneOrderRow(nRow As Integer, oProNo As Object, _
nRaceNo As Integer, nHeat As Integer, _
nNum As Integer, nStartLane As Integer)
    
    Dim sType As String             ' ��ڋ敪
    
    ' ���[�XNo�A�g�̏�����
    Cells(nRow, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
    Cells(nRow, Range(sTableName & "[�g]").Column).Value = nHeat
    
    ' ���{��I�茠���j���
    If GetRange("��").Value = "���{��I�茠���j���" Then
        Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
    ' ���{��s���̈���
    ElseIf GetRange("��").Value = "���{��s���̈���" Then
        ' ��ڋ敪
        sType = Cells(nRow, Range(sTableName & "[��ڋ敪]").Column).Value
              
        If sType = "�N��敪" Then
            ' �N��敪
            Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane2(nHeat, oProNo.Count, nNum)
        Else
            ' ���w�A���Z
            Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
        End If
    Else
        ' ���{��}�X�^�[�Y
        If Cells(nRow, Range(sTableName & "[�\�[�g�敪]").Column).Value <> "" Then
            Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane2(nHeat, oProNo.Count, nNum, False)
        ' �w��
        Else
            Cells(nRow, Range(sTableName & "[���[��]").Column).Value = GetLane(nHeat, oProNo.Count, nNum, False)
        End If
    End If
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
' nNumberOfHeat() OUT     �g���̐l���z��
'
Sub GetNumberOfHeat(nTotalNum As Integer, nHeats As Integer, nNumberOfHeat() As Integer)
    
    ReDim nNumberOfHeat(nHeats - 1) As Integer
    
    ' �P�g�ڐl���Z�o
    nNumberOfHeat(0) = GetFirstHeatNumber(nTotalNum)
    
    ' �Q�g�ڐl���Z�o
    If nHeats >= 2 Then
        nNumberOfHeat(1) = GetSecondHeatNumber(nTotalNum)
    
        ' �R�g�ڈȍ~
        If nHeats > 2 Then
            For i = 2 To nHeats - 1
                nNumberOfHeat(i) = nNumberOfRace
            Next
        End If
    End If
End Sub

'
' �g�ԍ�����
'
' �v��No�̑��l���Ə��Ԃ���g�ԍ����Z�o����
'
' �g���A�P�g�ڂ̐l���A�Q�g�ڂ̐l�����Z�o���邱�Ƃ�
' �g�ԍ����Z�o�\
'
' nOrder        IN      ����
' nHeats        IN      �g��
' nHeatNumber() IN      �g���̐l���z��
'
Function GetOrderHeat(nOrder As Integer, nHeats As Integer, nHeatNumber() As Integer)

    ' �P�g�ڂ̏ꍇ
    If nOrder <= nHeatNumber(0) Then
        GetOrderHeat = 1
    ' �Q�g�ڂ̏ꍇ
    ElseIf nOrder <= nHeatNumber(0) + nHeatNumber(1) Then
        GetOrderHeat = 2
    ' �R�g�ڈȍ~�̏ꍇ
    Else
        GetOrderHeat = GetHeats(nOrder - (nHeatNumber(0) + nHeatNumber(1))) + 2
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
' �Z���^�[���[���Z�o
'
' nCount        IN      �l��
' nStart        IN      �J�n�ʒu
'
Function GetCenterLane(nCount As Integer, nStart As Integer)
     GetCenterLane = nStart + Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
End Function

'
' �J�n�ʒu���Z�o
'
' nCount        IN      ���[�X�l��
'
Function GetStartLane(nCount As Integer)
     GetStartLane = nCenterLane - Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
End Function

'
' ���[������
'
' ���[���͋��Z�K���̒P�������ŕ��ׂ�
'
' nCenter       IN      �Z���^�[
' nMax          IN      �l��
' nOrder        IN      ����
' bFlag         In      True�F�ʏ�^False�F�t��
'
Function GetLane3(nCenter As Integer, nMax As Integer, nOrder As Integer, Optional bFlag As Boolean = True)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    If bFlag Then
        GetLane3 = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        GetLane3 = nCenter + Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
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

