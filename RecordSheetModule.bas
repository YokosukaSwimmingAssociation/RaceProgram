Attribute VB_Name = "RecordSheetModule"
'
' ��ږ��Ǎ���
'
' �L�^��ʂ�ProNo�����͂��ꂽ���ږ���Ǎ��ݕ\������
' ���݂��Ȃ�ProNo�̏ꍇ�͎�ږ��͋󗓂ƂȂ�
'
Sub ��ږ��Ǎ���()
    Sheets("�L�^���").Protect UserInterfaceOnly:=True
    For Each vNo In GetRange("�v���O������ڔԍ�")
        If vNo.Value = GetRange("�L�^��ʎ�ڔԍ�").Value Then
            ' ��ڋ敪�Ǝ�ږ���A�����ĕ\������
            GetRange("�L�^��ʎ�ږ�").Value = vNo.Offset(0, GetRange("Prog��ڋ敪").Column - vNo.Column).Value _
                    & " " & vNo.Offset(0, GetRange("Prog��ږ�").Column - vNo.Column).Value
            Exit Sub
        End If
    Next vNo
    ' �Y�����Ȃ��ꍇ�͏�����
    Range("�L�^��ʎ�ږ�").Value = ""
    Range("�L�^��ʑg").Value = 1
End Sub

'
' ���[�X�ԍ��Ǎ���
'
' �L�^��ʂ�ProNo�܂��͑g�����͂��ꂽ�烌�[�X�ԍ���Ǎ��ݕ\������
' ���݂��Ȃ��g�ݍ��킹�̏ꍇ�̓��[�X�ԍ��͋󗓂ƂȂ�
'
Sub ���[�X�ԍ��Ǎ���()
    Dim nProNo As Integer
    Dim nHeat As Integer

    nProNo = GetRange("�L�^��ʎ�ڔԍ�").Value
    nHeat = GetRange("�L�^��ʑg").Value
    
    Dim sName As String
    sName = "�v���O�����g" & Format(nProNo, "0#") & "_" & Trim(Str(nHeat))

    If IsNameExists(sName) Then
        For Each vLane In Range(sName)
            If vLane.Offset(0, GetRange("Header���[�XNo").Column - vLane.Column).Value <> "" Then
                GetRange("�L�^��ʃ��[�XNo").Value = vLane.Offset(0, GetRange("Header���[�XNo").Column - vLane.Column).Value
                Exit Sub
            End If
        Next vLane
    Else
        ' ���݂��Ȃ��v���O�����ԍ��̏ꍇ
        Range("�L�^��ʃ��[�XNo").Value = ""
    End If
    
End Sub

'
' �I�薼�Ǎ���
'
' �L�^��ʂŃ��[�������͂��ꂽ�ꍇ�ɑI�薼��Ǎ��ݕ\������
' ���[�X�ԍ����󗓂̏ꍇ�͉������Ȃ�
'
Sub �I�薼�Ǎ���()
    Dim nRaceNo As Integer
    nRaceNo = Range("�L�^��ʃ��[�XNo").Value
    If nRaceNo = 0 Then
        Exit Sub
    End If
    
    Dim nLane As Integer
    For Each vLane In Range("�L�^��ʃ��[��")
        nLane = Cells(vLane.Row, Range("�L�^��ʃ��[��").Column).Value
        ' �I�薼
        Cells(vLane.Row, Range("�L�^��ʑI�薼").Column).Value = SearchName(nRaceNo, nLane)
        ' �`�[����
        Cells(vLane.Row, Range("�L�^��ʃ`�[����").Column).Value = SearchTeam(nRaceNo, nLane)
    Next vLane
End Sub

'
' �I�薼����
'
' ���[�X�ԍ��A���[���ԍ�����I�薼����������
'
' ���O�u�v���O�������[�XN�v���烌�[�X�̃Z�����擾���ĒT������
'
' nRaceNo           IN      ���[�X�ԍ�
' nLane             IN      ���[���ԍ�
'
Function SearchName(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(Str(nRaceNo))

    ' ���݂��郌�[�X�ԍ��̏ꍇ
    If IsNameExists(sName) Then
        ' ���[�����ɏ�������
        For Each vLaneNo In Range(sName)
            ' ���[���ԍ����w�肳�ꂽ���[���ԍ��̏ꍇ
            If vLaneNo.Offset(0, Range("Prog���[��").Column - vLaneNo.Column).Value = nLane Then
                SearchName = vLaneNo.Offset(0, Range("Prog����").Column - vLaneNo.Column).Value
                ' ���O���󔒗p������̏ꍇ�͋󔒂ɂ���
                If SearchName = sBlankName Then
                    SearchName = ""
                End If
                Exit Function
            End If
        Next vLaneNo
    End If
    SearchName = ""
End Function

'
' �`�[��������
'
' ���[�X�ԍ��A���[���ԍ�����`�[��������������
'
' ���O�u�v���O�������[�XN�v���烌�[�X�̃Z�����擾���ĒT������
'
' nRaceNo           IN      ���[�X�ԍ�
' nLane             IN      ���[���ԍ�
'
Function SearchTeam(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(Str(nRaceNo))

    ' ���݂��郌�[�X�ԍ��̏ꍇ
    If IsNameExists(sName) Then
        ' ���[�����ɏ�������
        For Each vLaneNo In Range(sName)
            ' ���[���ԍ����w�肳�ꂽ���[���ԍ��̏ꍇ
            If vLaneNo.Offset(0, Range("Prog���[��").Column - vLaneNo.Column).Value = nLane Then
                SearchTeam = vLaneNo.Offset(0, Range("Prog����").Column - vLaneNo.Column).Value
                Exit Function
            End If
        Next vLaneNo
    End If
    SearchTeam = ""
End Function

'
' ���L�^����
'
' �^�C�������͂��ꂽ�ꍇ��
'
' ���O�u�v���O�������[�XN�v���烌�[�X�̃Z�����擾���ĒT������
'
Sub ���L�^����()
    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim nRecordTime As Integer

    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value
   
    For Each vLane In GetRange("�L�^��ʃ��[��")
        nLane = Cells(vLane.Row, GetRange("�L�^��ʃ��[��").Column).Value
        nTime = Cells(vLane.Row, GetRange("�L�^��ʃ^�C��").Column).Value
        
        If nLane > 0 And nTime > 0 Then
            If nTime < SearchRecord(nRaceNo, nLane) Then
                ' ���Ԃ����L�^��菬�����ꍇ�͑��V�i����^�C����NG�j
                Cells(vLane.Row, GetRange("�L�^��ʑ��V").Column).Value = "���V"
            Else
                ' ����ȊO�͋�
                Cells(vLane.Row, GetRange("�L�^��ʑ��V").Column).Value = ""
            End If
        Else
            ' �������͂���Ă��Ȃ����[������
            Cells(vLane.Row, GetRange("�L�^��ʑ��V").Column).Value = ""
        End If
    Next vLane
End Sub

'
' ���L�^�擾
'
' ���[�X�ԍ��A���[���ԍ�������L�^����������
'
' ���O�u�v���O�������[�XN�v���烌�[�X�̃Z�����擾���ĒT������
'
' nRaceNo           IN      ���[�X�ԍ�
' nLane             IN      ���[���ԍ�
'
Function SearchRecord(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(Str(nRaceNo))

    If IsNameExists(sName) Then
        For Each vLaneNo In Range(sName)
            If vLaneNo.Offset(0, GetRange("Prog���[��").Column - vLaneNo.Column).Value = nLane Then
                SearchRecord = vLaneNo.Offset(0, GetRange("Prog���L�^").Column - vLaneNo.Column).Value
                Exit For
            End If
        Next vLaneNo
    End If
End Function

'
' �L�^��ʂ̏��������s��
'
' ���[���̏������͍s������ڔԍ��A�g�̏������͍s��Ȃ�
'
Sub ������()
    Sheets("�L�^���").Protect UserInterfaceOnly:=True
    
    ' �C�x���g������}��
    Call EventChange(False)
    
    For Each vLane In GetRange("�L�^��ʃ��[��")
        vLane.Value = ""
        vLane.Offset(0, GetRange("�L�^��ʃ^�C��").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("�L�^��ʑI�薼").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("�L�^��ʃ`�[����").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("�L�^��ʑ��V").Column - vLane.Column).Value = ""
    Next vLane

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' ���̓f�[�^���v���O�����ɓo�^����
'
' �L�^��ʂœo�^�{�^���������ꂽ�ۂɃv���O�����ɋL������
'
Sub �o�^()
    ' �C�x���g������}��
    Call EventChange(False)

    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim sAdditional As String

    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value
    
    For Each vLane In GetRange("�L�^��ʃ��[��")
        nLane = Cells(vLane.Row, GetRange("�L�^��ʃ��[��").Column).Value
        nTime = Cells(vLane.Row, GetRange("�L�^��ʃ^�C��").Column).Value
        sAdditional = Cells(vLane.Row, GetRange("�L�^��ʑ��V").Column).Value
        
        If nLane <> 0 Then
            Call SetRecord(nRaceNo, nLane, nTime, sAdditional)
        End If
    Next vLane

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' ���̓f�[�^���v���O�����ɓo�^����
'
' nRaceNo           IN      ���[�X�ԍ�
' nLane             IN      ���[���ԍ�
' nTime             IN      �^�C��
' sAdditional       IN      ���V
'
Function SetRecord(nRaceNo As Integer, nLane As Integer, nTime As Long, sAdditional As String)

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(Str(nRaceNo))

    For Each vLaneNo In GetRange(sName)
        If vLaneNo.Offset(0, GetRange("Prog���[��").Column - vLaneNo.Column).Value = nLane Then
            If nTime = 0 Then
                ' �^�C�������͂���Ă��Ȃ��ꍇ�͊���
                vLaneNo.Offset(0, GetRange("Prog���l").Column - vLaneNo.Column).Value = sAdditional
                vLaneNo.Offset(0, GetRange("Prog���l").Column - vLaneNo.Column).Value = "����"
            Else
                ' �^�C�������͂���Ă���ꍇ�͎��ԂƔ��l��ݒ�
                vLaneNo.Offset(0, GetRange("Prog����").Column - vLaneNo.Column).Value = nTime
                vLaneNo.Offset(0, GetRange("Prog���l").Column - vLaneNo.Column).Value = sAdditional
            End If
            Exit Function
        End If
    Next vLaneNo
End Function


'
' ���ʌ���
'
' ���ꃌ�[�X���l�����āA���[�XNo�̒��Ɋ܂܂��v��No�ɑ΂���
' ���ׂď��ʂ�����
'
Sub ���ʌ���()
    ' �C�x���g������}��
    Call EventChange(False)

    Dim nRaceNo As Integer
    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(Str(nRaceNo))

    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    Dim nProNo As Integer
    If IsNameExists(sName) Then
        For Each vRaceNo In GetRange(sName)
            nProNo = vRaceNo.Offset(0, GetRange("Header�v��No").Column - vRaceNo.Column).Value
            ' �ŏ��̂P�񂾂����s
            If Not oProNo.Exists(nProNo) Then
                Call SetOrder(nProNo)
                oProNo.Add nProNo, 1
            End If
        Next vRaceNo
    End If

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' ���Ԃ����߂�
'
' nProNo            IN      ��ڔԍ�
'
Sub SetOrder(nProNo As Integer)

    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(Str(nProNo))
    
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    ' �Ǎ���
    Call ReadOrder(nProNo, sName, oProNo)

    ' �\�[�g
    Call SortDictOrder(nProNo, sName, oProNo)

End Sub

'
' ���Ԃ�ǂݍ���
'
' nProNo            IN      ��ڔԍ�
' sName             IN      ��ڔԍ��̖��O
' oProNo            OUT     �z��
'
Sub ReadOrder(nProNo As Integer, sName As String, oProNo As Object)
    
    ' �Ǎ���
    Dim oSubClass As Object
    If IsNameExists(sName) Then
        ' ���[�XNo���Ɏ��{
        For Each vLane In Range(sName)
            ' ���Ԃ����͂���Ă���ꍇ���Ώ�
            If IsNumeric(vLane.Offset(0, GetRange("Prog����").Column - vLane.Column).Value) Then
                ' �\�[�g�敪�i�N��敪�j���ɏ��ʂ�����
                sSubClass = vLane.Offset(0, GetRange("Header�\�[�g�敪").Column - vLane.Column).Value
                If sSubClass = "" Then
                    ' �\�[�g�敪���Ȃ��ꍇ�͂P�敪�iALL�j�Ƃ��Ă���
                    sSubClass = "ALL"
                End If
                If Not oProNo.Exists(sSubClass) Then
                    Set oSubClass = CreateObject("Scripting.Dictionary")
                    oProNo.Add sSubClass, oSubClass
                End If

                ' Key�i�s�j�FValue�i���ԁj�Ƃ��Ď����^�ɓo�^
                oSubClass.Add vLane.Row, vLane.Offset(0, GetRange("Prog����").Column - vLane.Column).Value
            End If
        Next vLane
    Else
        MsgBox "�v���O�����ԍ����s���ł��B"
        End
    End If
End Sub

'
' ���Ԃ�ǂݍ���
'
' nProNo            IN      ��ڔԍ�
' sName             IN      ��ڔԍ��̖��O
' oProNo            OUT     �z��
'
Sub SortDictOrder(nProNo As Integer, sName As String, oProNo As Object)
    ' ���ёւ�
    Dim oSubClass As Object
    Dim nOrder As Integer
    Dim nCount As Integer
    Dim nTime As Long
    Dim nPreTime As Long
    For Each vProNo In oProNo
        Set oSubClass = oProNo.Item(vProNo)
        ' ���ёւ������{
        Call DictQuickSort(oSubClass, "Value")
        nOrder = 1
        nCount = 1
        nPreTime = 0
        For Each vRow In oSubClass
            nTime = oSubClass.Item(vRow)
            ' ����^�C���łȂ��Ƃ��͏��ʂ��グ��
            If nTime > nPreTime Then
                nOrder = nCount
                nPreTime = nTime
            End If
            ' ���ʂ���������
            Sheets(Range(sName).Parent.Name).Cells(vRow, Range("Prog����").Column).Value = nOrder
            nCount = nCount + 1
        Next vRow
    Next vProNo
End Sub
