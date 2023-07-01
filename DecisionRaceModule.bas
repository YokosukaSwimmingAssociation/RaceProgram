Attribute VB_Name = "DecisionRaceModule"
Option Explicit    ''���ϐ��̐錾����������

'
' �g�ݍ��킹����
'
' ���[�XNo�Ƒg�A���[����ݒ肷��
'
Public Sub �g�ݍ��킹����()
    ' �C�x���g������}��
    Call EventChange(False)

    ' �G�N�Z���V�[�g��I��
    Call SheetActivate(�G���g���[�V�[�g)

    ' �o�͗p���[�N�u�b�N
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' �o�͗p���[�N�V�[�g
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ProNo�A�\�[�g�敪�A�\���ݎ��ԂŃ\�[�g
    Call SortByProNo(oWorkSheet, �G���g���[�e�[�u��)
    
    ' �g�ݍ��킹�쐬
    Call SetHeatLaneOrder(oWorkSheet, �G���g���[�e�[�u��)
    
    ' ���[�XNo, ���[���Ń\�[�g
    Call SortByRace(oWorkSheet, �G���g���[�e�[�u��)
    
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
Private Sub SetHeatLaneOrder(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate
    
    ' �G���g���[�ꗗ
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^���i�[
    Call ReadProNo(sTableName, oEntryList)
    
    ' �g�A���[�����o��
    Call WriteHeatLaneOrder(sTableName, oEntryList)

End Sub

'
' �v��No���L�[�Ƀf�[�^�i�[
'
' �v��No���ɒʔԂ�U���Ă��̍s�ԍ����i�[����
'
' sTableName    IN      �e�[�u����
' oEntryList    OUT     �G���g���[�ԍ�
'�@���v��No�F
'�@�@���ʔԁFProNo�s�̃Z���I�u�W�F�N�g
'
Private Sub ReadProNo(sTableName As String, oEntryList As Object)
    ' �f�[�^���i�[
    Dim vProNo As Variant
    Dim oProNo As Object
    For Each vProNo In Range(sTableName & "[�v��No]")
        If Not oEntryList.Exists(vProNo.Value) Then
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vProNo.Value, oProNo
        End If
        oProNo.Add oProNo.Count + 1, vProNo
    Next vProNo
End Sub

'
' �g�A���[�����o�͂���
'
' sTableName    IN      �e�[�u����
' oEntryList    IN      �G���g���[�ԍ�
'
Private Sub WriteHeatLaneOrder(sTableName As String, oEntryList As Object)

    Dim nNumOfProNo As Integer      ' �v���O����No���̐l��
    Dim nHeats As Integer           ' �g��
    Dim nNumOfHeats() As Integer    ' �g���̐l��
    Dim nMinNumberOfRace As Integer ' �g�̍ŏ��l��
    nMinNumberOfRace = GetRange("���g�ŏ��l��").Value
    Dim bAverage As Boolean         ' ���ϕ��������̗��p�L��
    
    Dim nRaceNo As Integer          ' ���[�XNo
    nRaceNo = 0
    
    ' �v���O����No��
    Dim vProNo As Variant
    Dim oProNo As Object            ' �v���O�������̔z��
    For Each vProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(vProNo)
        
        ' �v���O����No���̐l��
        nNumOfProNo = oProNo.Count
        
        ' �g�����Z�o
        nHeats = GetHeats(nNumOfProNo)
        Call GetNumberOfHeat(nNumOfProNo, nHeats, nNumOfHeats, nMinNumberOfRace)
        
        ' ���ϕ��������𗘗p���邩����
        bAverage = IsAverageOrder(vProNo, nHeats)
        
        ' �g���Ɍv�Z
        Call WriteHeatLaneOrderByProNo(nRaceNo, nHeats, nNumOfHeats, oProNo, sTableName, bAverage)
    
    Next vProNo
End Sub

'
' �v���O�����ԍ����ɑg�A���[�����o�͂���
'
' nRaceNo       IN      ���[�XNo
' nHeats        IN      �g��
' nNumOfHeats   IN      �g���̐l��
' oProNo        IN      �v���O�������̔z��
' sTableName    IN      �e�[�u����
' bAverage      IN      ���ϕ��������̗��p�L��
'
Private Sub WriteHeatLaneOrderByProNo(nRaceNo As Integer, nHeats As Integer, nNumOfHeats() As Integer, _
oProNo As Object, sTableName As String, bAverage As Boolean)
    
    Dim nNumOfHeat As Integer       ' �g�̐l��
    Dim nOrder As Integer           ' ProNo���̏����i�t���j
    nOrder = 1
    
    ' �g��
    Dim nHeat As Integer
    For nHeat = 1 To nHeats
        ' ���[�XNo���C���N�������g(�������₷���悤��10�����₷)
        nRaceNo = nRaceNo + 10
        
        ' ���ϕ�������
        If bAverage And nHeat = (nHeats - ���ϕ����g��) + 1 Then
            Call WriteHeatLaneOrderByAverage(nRaceNo, nHeat, nNumOfHeats, nOrder, oProNo, sTableName)
            Exit Sub
        Else
            ' �g�̐l��
            nNumOfHeat = nNumOfHeats(nHeat - 1)
            ' �g���ɏo��
            Call WriteHeatLaneOrderByHeat(nRaceNo, nHeat, nNumOfHeat, nOrder, oProNo, sTableName)
        End If
    
    Next nHeat
End Sub

'
' �g���ɑg�A���[�����o�͂���
'
' nRaceNo       IN      ���[�XNo
' nHeat         IN      �g�ԍ�
' nNumOfHeat    IN      �g�̐l��
' nOrder        IN      ����
' oProNo        IN      ����
' sTableName    IN      �e�[�u����
'
Private Sub WriteHeatLaneOrderByHeat(nRaceNo As Integer, nHeat As Integer, nNumOfHeat As Integer, _
nOrder As Integer, oProNo As Object, sTableName As String)
    
    ' �g�ݍ��킹���w�}�����t�ȑΉ��i���ؗp�̎b��j
    Dim bFlag As Boolean
    If GetRange("��").Value = �w�}��� Then
        bFlag = False
    Else
        bFlag = True
    End If
    
    Dim nStartLane As Integer       ' �g�̊J�n�ʒu
    Dim nTargetNum As Integer             ' �g�̒���Lane���肷��l��
    
    Dim nNumOfSortClass As Integer  ' �\�[�g�敪���̎c��l��
    Dim nRemNumber As Integer       ' �g�̎c��l��
    nRemNumber = nNumOfHeat
    
    ' �g�̊J�n�ʒu
    nStartLane = GetStartLane(nNumOfHeat, GetCenterLane(Range("���g���[�X���").Value, GetRange("���g�ŏ����[���ԍ�").Value), bFlag)
    
    ' �g�̐l�����c���Ă����
    While nRemNumber > 0
        ' �\�[�g�敪���̎c��l��
        nNumOfSortClass = GetNumberOfSortClass(nOrder, oProNo, Range(sTableName & "[�\�[�g�敪]").Column)
        If nNumOfSortClass <= nRemNumber Then
            ' Lane���肷��l��
            nTargetNum = nNumOfSortClass
        Else
            nTargetNum = nRemNumber
        End If
        
        ' �\�[�g�敪���ɑg�A���[�����o��
        Call WriteHeatLaneOrderBySortClass(nRaceNo, nHeat, nTargetNum, nOrder, oProNo, sTableName, nStartLane, bFlag)
    
        ' �J�n�ʒu��ύX
        nStartLane = nStartLane + nTargetNum
    
        ' �c��l�������Z
        nRemNumber = nRemNumber - nTargetNum
    Wend
End Sub

'
' �\�[�g�敪���ɑg�A���[�����o�͂���
'
' nRaceNo       IN      ���[�XNo
' nHeat         IN      �g�ԍ�
' nTargetNum    IN      �Ώۂ̐l��
' nOrder        IN/OUT  ����
' oProNo        IN      ���Ԃ̔z��
' sTableName    IN      �e�[�u����
' nStartLane    IN      �J�n���[��
' bFlag         In      True�F�ʏ�^False�F�t��
'
Private Sub WriteHeatLaneOrderBySortClass(nRaceNo As Integer, nHeat As Integer, _
nTargetNum As Integer, ByRef nOrder As Integer, _
oProNo As Object, sTableName As String, _
nStartLane As Integer, Optional bFlag As Boolean = True)

    Dim nCenterLane As Integer      ' �Z���^�[���[��

    ' ���[�������肷��
    Dim oCell As Range              ' �J�����g�s�̃Z��
    Dim nIndex As Integer           ' Lene���肷��l���̒��̏���
    For nIndex = 1 To nTargetNum
        ' �J�����g�s�ԍ�
        Set oCell = oProNo.Item(nOrder)
    
        ' ���[�XNo�A�g�̏�����
        GetOffset(oCell, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
        GetOffset(oCell, Range(sTableName & "[�g]").Column).Value = nHeat
    
        ' ���[�XNo�A�g�A���[�����L�q
        nCenterLane = GetCenterLane(nTargetNum, nStartLane, bFlag)
        GetOffset(oCell, Range(sTableName & "[���[��]").Column).Value = GetLane(nCenterLane, nTargetNum, nIndex, bFlag)
    
        ' ���Ԃ��C���N�������g
        nOrder = nOrder + 1
    Next nIndex

End Sub

'
' �J�n�ʒu���瓯��\�[�g�敪�̌���
'
' �J�n�ʒu�Ŏw�肳�ꂽ�\�[�g�敪�Ɠ����l�̊Ԃ̓J�E���g����
'
' nIndex            IN      �J�n�ʒu
' oProNo            IN      �v��No���̃G���g���[�z��
' nSortClassColumn  IN      �\�[�g�敪�̃J�����ʒu
'
Private Function GetNumberOfSortClass(nIndex As Integer, oProNo As Object, nSortClassColumn As Integer) As Integer
    
    GetNumberOfSortClass = 1
    
    Dim vProNo As Object
    Set vProNo = oProNo.Item(nIndex)
    Dim sSortClass As String
    sSortClass = GetOffset(vProNo, nSortClassColumn).Value
    
    Dim i As Integer
    For i = nIndex + 1 To oProNo.Count
        Set vProNo = oProNo.Item(i)
        If GetOffset(vProNo, nSortClassColumn).Value = sSortClass Then
            GetNumberOfSortClass = GetNumberOfSortClass + 1
        Else
            Exit Function
        End If
    Next i
End Function

'
' �g�l���z���ݒ�
'
' nTotalNum     IN      �v��No�̃G���g���[��
' nHeat         IN      �g��
' nNumberOfHeat() OUT   �g���̐l���z��
' nMinNumberOfRace IN   �g�̍ŏ��l��
'
Private Sub GetNumberOfHeat(nTotalNum As Integer, nHeats As Integer, nNumberOfHeat() As Integer, nMinNumberOfRace As Integer)
    
    ReDim nNumberOfHeat(nHeats - 1) As Integer
    
    ' �P�g�ڐl���Z�o
    nNumberOfHeat(0) = GetFirstHeatNumber(nTotalNum, nMinNumberOfRace)
    
    ' �Q�g�ڐl���Z�o
    If nHeats >= 2 Then
        nNumberOfHeat(1) = GetSecondHeatNumber(nTotalNum, nMinNumberOfRace)
    
        ' �R�g�ڈȍ~
        If nHeats > 2 Then
            Dim i As Integer
            For i = 2 To nHeats - 1
                nNumberOfHeat(i) = Range("���g���[�X���").Value
            Next i
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
Private Function GetHeats(nTotalNum As Integer) As Integer

    GetHeats = Application.WorksheetFunction.RoundUp(nTotalNum / Range("���g���[�X���").Value, 0)

End Function

'
' �P�g�ڐl���Z�o
'
' ���l�����ŏ��l���ȏア��ꍇ�A�ŏ��l���͂P�g�ڂɎc��
'
' nTotalNum     IN      ���[�X�̑��l��(�����g�̑���)
' nMinNumberOfRace IN   �g�̍ŏ��l��
'
Private Function GetFirstHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer) As Integer

    Dim maxNum
    maxNum = Range("���g���[�X���").Value
    If nTotalNum <= maxNum Then
        GetFirstHeatNumber = nTotalNum
    ElseIf nTotalNum Mod maxNum = 0 Then
        GetFirstHeatNumber = maxNum
    ElseIf nTotalNum Mod maxNum <= nMinNumberOfRace Then
        GetFirstHeatNumber = nMinNumberOfRace
    Else
        GetFirstHeatNumber = nTotalNum Mod maxNum
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
Private Function GetSecondHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer) As Integer

    Dim maxNum
    maxNum = Range("���g���[�X���").Value
    If nTotalNum <= maxNum Then
        GetSecondHeatNumber = 0
    ElseIf nTotalNum Mod maxNum = 0 Then
        GetSecondHeatNumber = maxNum
    ElseIf nTotalNum Mod maxNum <= nMinNumberOfRace Then
        GetSecondHeatNumber = maxNum + (nTotalNum Mod maxNum - nMinNumberOfRace)
    Else
        GetSecondHeatNumber = maxNum
    End If

End Function

'
' �Z���^�[���[���Z�o
'
' nCount        IN      �l��
' nStart        IN      �J�n�ʒu
'
Public Function GetCenterLane(nCount As Integer, nStart As Integer, Optional bFlag As Boolean = True) As Integer
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
Private Function GetStartLane(nCount As Integer, nCenterLane As Integer, Optional bFlag As Boolean = True) As Integer
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
' �t���ɂ��邽�߂ɑ��l�����珇���������Ă���
'
' nCenter       IN      �Z���^�[
' nMaxNum       IN      ���l��
' nOrder        IN      ����
' bFlag         In      True�F�ʏ�^False�F�t��
'
Private Function GetLane(nCenter As Integer, nMaxNum As Integer, nOrder As Integer, Optional bFlag As Boolean = True)
    Dim nNum As Integer
    nNum = nMaxNum - nOrder + 1
    If bFlag Then
        GetLane = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        GetLane = nCenter + Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    End If
End Function

'
'  ���ϕ����������̗p���邩
'
' �I�茠���̗\�I�ŁA���������������w�肳��Ă���ꍇ�̂ݑΏ�
'
' vProNo        IN      �v���O�����ԍ�
' nHeats        IN      �g��
'
Private Function IsAverageOrder(vProNo As Variant, nHeats As Integer) As Boolean
    IsAverageOrder = False
    If GetRange("��").Value = �I�茠��� And nHeats >= ���ϕ����g�� Then
        If GetRange("���g��������").Value = "������������" And _
            VLookupArea(vProNo, "�I�茠��ڋ敪", "�\�I�^����") = "�\�I" Then
            IsAverageOrder = True
        End If
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
Private Function GetLane2(nCenter As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = Application.WorksheetFunction.RoundUp((nMax - nOrder + 1) / ���ϕ����g��, 0)
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
Private Function GetOrderHeat(nHeats As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    GetOrderHeat = nHeats - (nNum - 1) Mod nHeats
End Function


'
' ���[������(���ϕ�������)
'
' nStartRaceNo  IN/OUT  �J�n��RaceNo
' nStartHeat    IN      �J�n�̑g�ԍ�
' nNumOfHeats   IN      �g���̐l��
' nOrder        IN      ����
' oProNo        IN      ����
' sTableName    IN      �e�[�u����
'
Private Sub WriteHeatLaneOrderByAverage(nStartRaceNo As Integer, nStartHeat As Integer, _
nNumOfHeats() As Integer, nOrder As Integer, oProNo As Object, sTableName As String)
    Dim nMaxNum As Integer
    nMaxNum = 0
    Dim i As Integer
    For i = nStartHeat To nStartHeat + ���ϕ����g�� - 1
        nMaxNum = nMaxNum + nNumOfHeats(i - 1)
    Next i
    Call AverageMethod(nStartRaceNo, nStartHeat, nMaxNum, nOrder, oProNo, sTableName)
    nStartRaceNo = (nStartRaceNo / 10 - 1 + ���ϕ����g��) * 10
End Sub


'
' ���[������(���ϕ�������)
'
' ���[���͋��Z�K���̕��ϕ��������ŕ��ׂ�
'
' nStartRaceNo  IN      �J�n��RaceNo
' nStartHeat    IN      �J�n�̑g�ԍ�
' nMaxNum       IN      �l��
' nOrder        IN      ����
' oProNo        IN      ���Ԕz��
' sTableName    IN      �e�[�u����
'
Private Sub AverageMethod(nStartRaceNo As Integer, nStartHeat As Integer, nMaxNum As Integer, _
nOrder As Integer, oProNo As Object, sTableName As String)
    
    Dim nCenterLane As Integer
    Dim nRaceNo As Integer
    Dim nHeat As Integer
    
    ' �g�̃Z���^�[
    nCenterLane = GetCenterLane(Range("���g���[�X���").Value, GetRange("���g�ŏ����[���ԍ�").Value)
    
    ' �g�̐l�����c���Ă����
    Dim oCell As Range              ' �J�����g�s�̃Z��
    Dim nIndex As Integer
    For nIndex = 1 To nMaxNum
        
        ' �J�����g�s�ԍ�
        Set oCell = oProNo.Item(nOrder)
        
        ' ���[�XNo
        nRaceNo = (GetOrderHeat(���ϕ����g��, nMaxNum, nIndex) + (nStartRaceNo / 10 - 1)) * 10
        ' �g�ԍ�
        nHeat = GetOrderHeat(���ϕ����g��, nMaxNum, nIndex) + (nStartHeat - 1)
    
        ' ���[�XNo�A�g�̏�����
        GetOffset(oCell, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
        GetOffset(oCell, Range(sTableName & "[�g]").Column).Value = nHeat
    
        ' ���[�XNo�A�g�A���[�����L�q
        GetOffset(oCell, Range(sTableName & "[���[��]").Column).Value = GetLane2(nCenterLane, nMaxNum, nIndex)
    
        ' ���Ԃ��C���N�������g
        nOrder = nOrder + 1
    Next nIndex
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
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(�G���g���[�V�[�g)
    
    ' �ă\�[�g
    Call SortByRace(oWorkSheet, �G���g���[�e�[�u��)
    
    ' ProNo�A�g�̏d���`�F�b�N
    Dim oEntryList As Object
    Call ReadEntrySheet(�G���g���[�e�[�u��, oEntryList)
    
    ' ���[�X�ԍ��C��
    If GetRange("��").Value = �I�茠��� Then
        Call ModifyRaceNoForSenshuken(�G���g���[�e�[�u��)
    Else
        Call ModifyRaceNo(�G���g���[�e�[�u��)
    End If

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' ���[�X�ԍ��C��
'
' sTableName    IN  �e�[�u����
'
Private Sub ModifyRaceNo(sTableName As String)
    
    ' �G���g���[�ꗗ
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^���i�[�����[�X�ԍ��̃`�F�b�N
    Call ReadEntryByRaceNo(sTableName, oEntryList)
    
    ' �o��
    Call ModifyEntryByRaceNo(sTableName, oEntryList)

End Sub

'
' ���[�X�ԍ��̓Ǎ���
'
' �Ǎ��݂Ȃ��烌�[�X�ԍ����̃��[���d���`�F�b�N���s��
'
' sTableName    IN      �e�[�u����
' oEntryList    IN/OUT  �G���g���[�ꗗ
'   �����[�XNo
'   �@�@�����[���F�v��No��̃Z��
'
Private Sub ReadEntryByRaceNo(sTableName As String, oEntryList As Object)
    Dim nLane As Integer
    Dim oRaceNo As Object
    Dim vRaceNo As Variant
    For Each vRaceNo In Range(sTableName & "[���[�XNo]")
        ' ���݂��Ȃ����[�XNo�̏ꍇ��
        If Not oEntryList.Exists(vRaceNo.Value) Then
            ' �G���g���[�ꗗ�ɓo�^����
            Set oRaceNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vRaceNo.Value, oRaceNo
        End If
        
        ' ���[���ԍ����擾
        nLane = GetOffset(vRaceNo, Range(sTableName & "[���[��]").Column).Value
        
        ' ���[�X�ԍ��ɑ΂��郌�[���̏d���`�F�b�N
        If oRaceNo.Exists(nLane) Then
            MsgBox "���[�XNo�F" & Str(vRaceNo.Value) & vbCrLf & _
                    "���[���@�F" & Str(nLane) & vbCrLf & _
                    "���d�����Ă��܂��B"
            Range(sTableName).Parent.Activate
            Range(vRaceNo, GetOffset(vRaceNo, Range(sTableName & "[���[��]").Column)).Select
            vRaceNo.Activate
            End
        Else
            oRaceNo.Add nLane, vRaceNo
        End If
    Next vRaceNo
End Sub

'
' ���[�X�ԍ����C���o��
'
' sTableName    IN      �e�[�u����
' oEntryList    IN/OUT  �G���g���[�ꗗ
'
Private Sub ModifyEntryByRaceNo(sTableName As String, oEntryList As Object)
    ' ���[�X�ԍ��C��
    Dim nRaceNo As Integer
    nRaceNo = 1
    
    Dim oCell As Range
    Dim vRaceNo As Variant
    Dim oRaceNo As Object
    For Each vRaceNo In oEntryList.Keys
        Set oRaceNo = oEntryList.Item(vRaceNo)
        
        ' ���[��
        Dim vLane As Variant
        For Each vLane In oRaceNo.Keys
            Set oCell = oRaceNo.Item(vLane)
            
            ' ���[�X�ԍ�
            GetOffset(oCell, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
        Next
        nRaceNo = nRaceNo + 1
    Next vRaceNo
End Sub

'
' ���[�X�ԍ��C��(�I�茠�p)
'
' sTableName    IN  �e�[�u����
'
Private Sub ModifyRaceNoForSenshuken(sTableName As String)
    
    ' �G���g���[�ꗗ
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^���i�[�����[�X�ԍ��̃`�F�b�N
    Call ReadRaceForSenshuken(sTableName, oEntryList)
    
    ' �o��
    Call ModifyEntryByProNoForSenshuken(sTableName, oEntryList)

End Sub

'
' ���[�X�ԍ��̓Ǎ���(�I�茠�p)
'
' �Ǎ��݂Ȃ��烌�[�X�ԍ����̃��[���d���`�F�b�N���s��
'
' sTableName    IN      �e�[�u����
' oEntryList    IN/OUT �G���g���[�ꗗ
' ���v��No
'   �����[�XNo
'   �@�@�����[���F�s
'
Private Sub ReadRaceForSenshuken(sTableName As String, oEntryList As Object)
    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim oProNo As Object
    Dim oRaceNo As Object
    
    Dim vProNo As Variant
    For Each vProNo In Range(sTableName & "[�v��No]")
        ' ���݂��Ȃ��v��No�̏ꍇ��
        If Not oEntryList.Exists(vProNo.Value) Then
            ' �G���g���[�ꗗ�ɓo�^����
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vProNo.Value, oProNo
        End If
        
        ' ���݂��Ȃ����[�XNo�̏ꍇ��
        nRaceNo = GetOffset(vProNo, Range(sTableName & "[���[�XNo]").Column).Value
        If Not oProNo.Exists(nRaceNo) Then
            ' �G���g���[�ꗗ�ɓo�^����
            Set oRaceNo = CreateObject("Scripting.Dictionary")
            oProNo.Add nRaceNo, oRaceNo
        End If
        
        ' ���[���ԍ����擾
        nLane = GetOffset(vProNo, Range(sTableName & "[���[��]").Column).Value
        ' ���[�X�ԍ��ɑ΂��郌�[���̏d���`�F�b�N
        If oRaceNo.Exists(nLane) Then
            MsgBox "���[�XNo�F" & Str(nRaceNo) & vbCrLf & _
                    "���[���@�F" & Str(nLane) & vbCrLf & _
                    "���d�����Ă��܂��B"
            Range(sTableName).Parent.Activate
            Range(vProNo, GetOffset(vProNo, Range(sTableName & "[���[��]").Column)).Select
            vProNo.Activate
            End
        Else
            oRaceNo.Add nLane, vProNo
        End If
    Next vProNo
End Sub

'
' ���[�X�ԍ��C��(�I�茠�p)
'
' sTableName    IN  �e�[�u����
' oEntryList    IN  �G���g���[�ꗗ
'
Private Sub ModifyEntryByProNoForSenshuken(sTableName As String, oEntryList As Object)
    
    ' ���[�X�ԍ��C��
    Dim nRaceNo As Integer
    nRaceNo = 1
    
    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData("�I�茠��ڋ敪")
        ' ProNo�̃G���g���[���Ȃ��ꍇ�̓X�L�b�v
        If oEntryList.Exists(vProNo.Value) Then
            
            ' �\�I�����̏C��
            Call ModifyFinalEntry(vProNo, oEntryList, nRaceNo)
            
            ' ���[�XNo�̏C��
            If oEntryList.Exists(vProNo.Value) Then
                Call ModifyEntryByRaceNoForSenshuken(sTableName, vProNo, oEntryList, nRaceNo)
            End If
                 
        End If
    Next vProNo

End Sub

'
' �\�I�����̏C��
'
' vProNo        IN      ProNo��̃Z��
' oEntryList    IN/OUT  �G���g���[�z��
' nRaceNo       IN/OUT  ���[�X�ԍ�
'
Private Sub ModifyFinalEntry(vProNo As Variant, oEntryList As Object, nRaceNo As Integer)
    ' ProNo�̃G���g���[���擾
    Dim oProNo As Object
    Set oProNo = oEntryList.Item(vProNo.Value)
    
    ' �����ԍ����擾
    Dim nFinalNo As Integer
    nFinalNo = VLookupArea(vProNo.Value, "�I�茠��ڋ敪", "�����ԍ�")
    
    ' �\�I�̏ꍇ
    If vProNo.Value <> nFinalNo Then
        ' �g���P�g�̏ꍇ
        If oProNo.Count = 1 Then
            ' ProNo�������ɓ���ւ���
            oEntryList.Add nFinalNo, oProNo
            oEntryList.Remove vProNo.Value
            ' ���[�X�ԍ��͐U��Ȃ�
            Set oProNo = CreateObject("Scripting.Dictionary")
        ElseIf oProNo.Count > 1 Then
            ' �����ɋ�̃G���g���[������Ă���
            Dim oFinalProNo As Object
            Set oFinalProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add nFinalNo, oFinalProNo
        End If
    Else
        ' �����̑g��0�̏ꍇ
        If oProNo.Count = 0 Then
            ' �\�I����������p�^�[���Ȃ̂ŃC���N�������g���Ă���
            nRaceNo = nRaceNo + 1
        End If
    End If
End Sub

'
' ���[�X�ԍ��̏C���o��
'
' sTableName    IN  �e�[�u����
' vProNo        IN      ProNo��̃Z��
' oEntryList    IN/OUT  �G���g���[�z��
' nRaceNo       IN/OUT  ���[�X�ԍ�
'
Private Sub ModifyEntryByRaceNoForSenshuken(sTableName As String, _
vProNo As Variant, oEntryList As Object, nRaceNo As Integer)
    Dim oCell As Range
    Dim oProNo As Object
    Dim oRaceNo As Object
    
    ' ProNo�̃G���g���[���擾
    Set oProNo = oEntryList.Item(vProNo.Value)
    
    Dim vRaceNo As Variant
    For Each vRaceNo In oProNo.Keys()
        Set oRaceNo = oProNo.Item(vRaceNo)
    
        ' ���[��
        Dim vLane As Variant
        For Each vLane In oRaceNo.Keys
            Set oCell = oRaceNo.Item(vLane)
            
            ' ���[�X�ԍ�
            GetOffset(oCell, Range(sTableName & "[���[�XNo]").Column).Value = nRaceNo
        Next
    
        nRaceNo = nRaceNo + 1
    Next vRaceNo
End Sub


'
' ���[�XNo�A�g�Ń\�[�g����
'
' oWorkSheet    IN      ���[�N�V�[�g
' sTableName    IN      �e�[�u����
'
Private Sub SortByRace(oWorkSheet As Worksheet, sTableName As String)

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

