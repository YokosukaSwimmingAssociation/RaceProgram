Attribute VB_Name = "RecordSheetModule"
Option Explicit    ''���ϐ��̐錾����������

'
' ��ږ��Ǎ���
'
' �L�^��ʂ�ProNo�����͂��ꂽ���ږ���Ǎ��ݕ\������
' ���݂��Ȃ�ProNo�̏ꍇ�͎�ږ��͋󗓂ƂȂ�
'
Public Sub ��ږ��Ǎ���()
    Sheets("�L�^���").Protect UserInterfaceOnly:=True
    If Trim(GetRange("�L�^��ʎ�ڔԍ�").Value) <> "" Then
        Dim vNo As Range
        For Each vNo In GetRange("�v���O������ڔԍ�")
            If vNo.Value = GetRange("�L�^��ʎ�ڔԍ�").Value Then
                ' ��ڋ敪�Ǝ�ږ���A�����ĕ\������
                GetRange("�L�^��ʎ�ږ�").Value = GetOffset(vNo, GetRange("Prog��ڋ敪").Column).Value _
                        & " " & GetOffset(vNo, GetRange("Prog��ږ�").Column).Value
                Exit Sub
            End If
        Next vNo
    End If
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
Public Sub ���[�X�ԍ��Ǎ���()
    Dim nProNo As Integer
    Dim nHeat As Integer

    nProNo = GetRange("�L�^��ʎ�ڔԍ�").Value
    nHeat = GetRange("�L�^��ʑg").Value
    
    Dim sName As String
    sName = "�v���O�����g" & Format(nProNo, "0#") & "_" & Trim(CStr(nHeat))

    If IsNameExists(sName) Then
        Dim vLane As Range
        For Each vLane In Range(sName)
            If GetOffset(vLane, GetRange("Header���[�XNo").Column).Value <> "" Then
                GetRange("�L�^��ʃ��[�XNo").Value = GetOffset(vLane, GetRange("Header���[�XNo").Column).Value
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
' oLaneCell     IN  �ύX�̂��������[���̃Z��
'
Public Sub �I�薼�Ǎ���(oLaneCell As Range)
    Dim nRaceNo As Integer
    nRaceNo = Range("�L�^��ʃ��[�XNo").Value
    If nRaceNo = 0 Then
        Exit Sub
    End If
    
    Dim nLane As Integer
    nLane = oLaneCell.Value
    ' �I�薼
    GetOffset(oLaneCell, Range("�L�^��ʑI�薼").Column).Value = SearchName(nRaceNo, nLane)
    ' �`�[����
    GetOffset(oLaneCell, Range("�L�^��ʃ`�[����").Column).Value = SearchTeam(nRaceNo, nLane)
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
Private Function SearchName(nRaceNo As Integer, nLane As Integer) As String

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(CStr(nRaceNo))

    ' ���݂��郌�[�X�ԍ��̏ꍇ
    If IsNameExists(sName) Then
        ' ���[�����ɏ�������
        Dim vLaneNo As Range
        For Each vLaneNo In Range(sName)
            ' ���[���ԍ����w�肳�ꂽ���[���ԍ��̏ꍇ
            If GetOffset(vLaneNo, Range("Prog���[��").Column).Value = nLane Then
                SearchName = GetOffset(vLaneNo, Range("Prog����").Column).Value
                ' ���O���󔒗p������̏ꍇ�͋󔒂ɂ���
                If SearchName = �I�薼�u�����N Then
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
Private Function SearchTeam(nRaceNo As Integer, nLane As Integer) As String

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(CStr(nRaceNo))

    ' ���݂��郌�[�X�ԍ��̏ꍇ
    If IsNameExists(sName) Then
        ' ���[�����ɏ�������
        Dim vLaneNo As Range
        For Each vLaneNo In Range(sName)
            ' ���[���ԍ����w�肳�ꂽ���[���ԍ��̏ꍇ
            If GetOffset(vLaneNo, Range("Prog���[��").Column).Value = nLane Then
                SearchTeam = GetOffset(vLaneNo, Range("Prog����").Column).Value
                Exit Function
            End If
        Next vLaneNo
    End If
    SearchTeam = ""
End Function

'
' �ᔽ���f
'
' oDqCell       IN  �ύX�̂������ᔽ�Z��
'
Public Sub �ᔽ���f(oDqCell As Range)
    Application.EnableEvents = False
    
    Dim sDq As String
    sDq = STrimAll(oDqCell.Value)
    
    ' OP��ݒ肳��Ă���ꍇ
    If sDq = "OP" Then
        ' �^�C�����c����OP��ݒ�
        GetOffset(oDqCell, GetRange("�L�^��ʔ��l").Column).Value = sDq
    
    ' ���i��ݒ肳��Ă���ꍇ
    ElseIf sDq <> "" Then
        ' �^�C������ɂ��Ĉᔽ��ݒ�
        GetOffset(oDqCell, GetRange("�L�^��ʃ^�C��").Column).Value = ""
        GetOffset(oDqCell, GetRange("�L�^��ʔ��l").Column).Value = sDq
    
    ' �󔒂ɖ߂����ꍇ
    Else
        Call ���L�^����(GetOffset(oDqCell, GetRange("�L�^��ʃ^�C��").Column))
    End If

    Application.EnableEvents = True
End Sub

'
' ���L�^����
'
' �^�C�������͂��ꂽ�ꍇ��
'
' ���O�u�v���O�������[�XN�v���烌�[�X�̃Z�����擾���ĒT������
'
' oTimeCell       IN  �ύX�̂������^�C���Z��
'
Public Sub ���L�^����(oTimeCell As Range)
    Application.EnableEvents = False
    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim nRecordTime As Long
    Dim nQualifyTime As Long
    
    ' �L�^��p�b��
    GetOffset(oTimeCell, GetRange("�L�^��ʔ��l").Column).Value = ""
    Exit Sub

    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value
   
    nLane = GetOffset(oTimeCell, GetRange("�L�^��ʃ��[��").Column).Value
    nTime = GetOffset(oTimeCell, GetRange("�L�^��ʃ^�C��").Column).Value
    
    ' ���[���A�^�C���ɒl���ݒ肳��Ă���ꍇ
    If nLane > 0 And nTime > 0 Then
        nRecordTime = SearchRecord(nRaceNo, nLane)
        nQualifyTime = SearchQualify(nRaceNo, nLane)
        If nQualifyTime > 0 And nTime > nQualifyTime Then
            ' ���Ԃ��W���L�^���傫���ꍇ�̓^�C�����i
            GetOffset(oTimeCell, GetRange("�L�^��ʔ��l").Column).Value = "�^�C�����i"
        ElseIf nRecordTime = 0 Or nTime < nRecordTime Then
            ' ���Ԃ����L�^��菬�����ꍇ�͑��V�i����^�C����NG�j
            GetOffset(oTimeCell, GetRange("�L�^��ʔ��l").Column).Value = "���V"
        Else
            ' ����ȊO�͋�
            GetOffset(oTimeCell, GetRange("�L�^��ʔ��l").Column).Value = ""
        End If
    Else
        ' �������͂���Ă��Ȃ����[������
        GetOffset(oTimeCell, GetRange("�L�^��ʔ��l").Column).Value = ""
    End If
    GetOffset(oTimeCell, GetRange("�L�^��ʈᔽ").Column).Value = ""

    Application.EnableEvents = True
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
Private Function SearchRecord(nRaceNo As Integer, nLane As Integer) As Long

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(CStr(nRaceNo))

    If IsNameExists(sName) Then
        Dim vLaneNo As Range
        For Each vLaneNo In Range(sName)
            If GetOffset(vLaneNo, GetRange("Prog���[��").Column).Value = nLane Then
                If IsNumeric(GetOffset(vLaneNo, GetRange("Prog���L�^").Column).Value) Then
                    SearchRecord = CLng(GetOffset(vLaneNo, GetRange("Prog���L�^").Column).Value)
                Else
                    SearchRecord = 0
                End If
                Exit For
            End If
        Next vLaneNo
    End If
End Function

'
' �W���L�^�擾
'
' ���[�X�ԍ��A���[���ԍ�����W���L�^����������
'
' ���O�u�v���O�������[�XN�v���烌�[�X�̃Z�����擾���ĒT������
'
' nRaceNo           IN      ���[�X�ԍ�
' nLane             IN      ���[���ԍ�
'
Private Function SearchQualify(nRaceNo As Integer, nLane As Integer) As Long

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(CStr(nRaceNo))

    If IsNameExists(sName) Then
        Dim vLaneNo As Range
        For Each vLaneNo In Range(sName)
            If GetOffset(vLaneNo, GetRange("Prog���[��").Column).Value = nLane Then
                If IsNumeric(GetOffset(vLaneNo, GetRange("Prog�W���L�^").Column).Value) Then
                    SearchQualify = CLng(GetOffset(vLaneNo, GetRange("Prog�W���L�^").Column).Value)
                Else
                    SearchQualify = 0
                End If
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
Public Sub ������()
    Sheets("�L�^���").Protect UserInterfaceOnly:=True
    
    ' �C�x���g������}��
    Call EventChange(False)
    
    Dim vLane As Range
    For Each vLane In GetRange("�L�^��ʃ��[��")
        vLane.Value = ""
        GetOffset(vLane, GetRange("�L�^��ʃ^�C��").Column).Value = ""
        GetOffset(vLane, GetRange("�L�^��ʑI�薼").Column).Value = ""
        GetOffset(vLane, GetRange("�L�^��ʃ`�[����").Column).Value = ""
        GetOffset(vLane, GetRange("�L�^��ʔ��l").Column).Value = ""
        GetOffset(vLane, GetRange("�L�^��ʈᔽ").Column).Value = ""
    Next vLane

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' ���̓f�[�^���v���O�����ɓo�^����
'
' �L�^��ʂœo�^�{�^���������ꂽ�ۂɃv���O�����ɋL������
'
Public Sub �o�^()
    ' �C�x���g������}��
    Call EventChange(False)

    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim sAdditional As String

    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value
    
    Dim vLane As Range
    For Each vLane In GetRange("�L�^��ʃ��[��")
        nLane = GetOffset(vLane, GetRange("�L�^��ʃ��[��").Column).Value
        nTime = GetOffset(vLane, GetRange("�L�^��ʃ^�C��").Column).Value
        sAdditional = GetOffset(vLane, GetRange("�L�^��ʔ��l").Column).Value
        
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
Private Sub SetRecord(nRaceNo As Integer, nLane As Integer, nTime As Long, sAdditional As String)

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(CStr(nRaceNo))

    Dim vLaneNo As Range
    For Each vLaneNo In GetRange(sName)
        If GetOffset(vLaneNo, GetRange("Prog���[��").Column).Value = nLane Then
            If nTime = 0 Then
                ' �^�C�������͂���Ă��Ȃ��ꍇ
                If sAdditional <> "" Then
                    ' ���l�̒l��ݒ�
                    GetOffset(vLaneNo, GetRange("Prog���l").Column).Value = sAdditional
                Else
                    ' ���l���󗓂Ȃ����
                    GetOffset(vLaneNo, GetRange("Prog���l").Column).Value = "����"
                End If
            Else
                ' �^�C�������͂���Ă���ꍇ�͎��ԂƔ��l��ݒ�
                GetOffset(vLaneNo, GetRange("Prog����").Column).Value = nTime
                GetOffset(vLaneNo, GetRange("Prog���l").Column).Value = sAdditional
            End If
            Exit Sub
        End If
    Next vLaneNo
End Sub


'
' ���ʌ���
'
' ���ꃌ�[�X���l�����āA���[�XNo�̒��Ɋ܂܂��v��No�ɑ΂���
' ���ׂď��ʂ�����
'
Public Sub ���ʌ���()
    ' �C�x���g������}��
    Call EventChange(False)

    Dim nRaceNo As Integer
    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value

    Dim sName As String
    sName = "�v���O�������[�X" & Trim(CStr(nRaceNo))

    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    Dim nProNo As Integer
    If IsNameExists(sName) Then
        Dim vRaceNo As Range
        For Each vRaceNo In GetRange(sName)
            nProNo = GetOffset(vRaceNo, GetRange("Header�v��No").Column).Value
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
Private Sub SetOrder(nProNo As Integer)

    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(CStr(nProNo))
    
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    ' �Ǎ���
    Call ReadOrder(nProNo, sName, oProNo)

    ' �\�[�g���ď��Ԃ�ݒ�
    Call SortDictOrder(nProNo, sName, oProNo)

End Sub

'
' ���Ԃ�t���郌�[���ǂݍ���
'
' �L�^���Ȃ��ꍇ�͓ǂݍ��܂Ȃ�
' �^�C�����i�̏ꍇ�͓ǂݍ��܂Ȃ�
' OP�̏ꍇ�͓ǂݍ��܂Ȃ�
'
' nProNo            IN      ��ڔԍ�
' sName             IN      ��ڔԍ��̖��O
' oProNo            OUT     �z��
'
Private Sub ReadOrder(nProNo As Integer, sName As String, oProNo As Object)
    
    ' �Ǎ���
    Dim sSubClass As String
    Dim oSubClass As Object
    If IsNameExists(sName) Then
        ' ���[�XNo���Ɏ��{
        Dim vLane As Range
        For Each vLane In Range(sName)
            ' �L�^������A�^�C�����i�AOP�łȂ������Ώ�
            If IsNumeric(GetOffset(vLane, GetRange("Prog����").Column).Value) And _
                GetOffset(vLane, GetRange("Prog���l").Column).Value <> "�^�C�����i" And _
                GetOffset(vLane, GetRange("Prog���l").Column).Value <> "OP" Then
                ' �\�[�g�敪�i�N��敪�j���ɏ��ʂ�����
                sSubClass = GetOffset(vLane, GetRange("Header�\�[�g�敪").Column).Value
                If sSubClass = "" Then
                    ' �\�[�g�敪���Ȃ��ꍇ�͂P�敪�iALL�j�Ƃ��Ă���
                    sSubClass = "ALL"
                End If
                If Not oProNo.Exists(sSubClass) Then
                    Set oSubClass = CreateObject("Scripting.Dictionary")
                    oProNo.Add sSubClass, oSubClass
                End If

                ' Key�i�s�j�FValue�i���ԁj�Ƃ��Ď����^�ɓo�^
                oSubClass.Add vLane.Row, GetOffset(vLane, GetRange("Prog����").Column).Value
            End If
        Next vLane
    Else
        MsgBox "�v���O�����ԍ����s���ł��B"
        End
    End If
End Sub

'
' �\�[�g���ď��Ԃ�ݒ�
'
' nProNo            IN      ��ڔԍ�
' sName             IN      ��ڔԍ��̖��O
' oProNo            OUT     �z��
'
Private Sub SortDictOrder(nProNo As Integer, sName As String, oProNo As Object)
    ' ���ёւ�
    Dim oSubClass As Object
    Dim nOrder As Integer
    Dim nCount As Integer
    Dim nTime As Long
    Dim nPreTime As Long
    Dim vProNo As Variant
    For Each vProNo In oProNo
        Set oSubClass = oProNo.Item(vProNo)
        ' ���ёւ������{
        Call DictQuickSort(oSubClass, "Value")
        nOrder = 1
        nCount = 1
        nPreTime = 0
        Dim vRow As Variant
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

'
' �\�I�̏ꍇ�Ɍ������쐬
'
Public Sub �����o�^()
    ' �C�x���g������}��
    Call EventChange(False)

    ' �I�茠�ȊO�͖���
    If GetRange("��").Value <> �I�茠��� Then
        Exit Sub
    End If

    Dim nProNo As Integer
    nProNo = GetRange("�L�^��ʎ�ڔԍ�").Value

    Dim nFinalNo As Integer
    nFinalNo = VLookupArea(nProNo, "�I�茠��ڋ敪", "�����ԍ�")

    ' �������Ȃ��ꍇ������
    If nProNo = nFinalNo Then
        Exit Sub
    End If

    ' �����i�o�҂�ǂݍ���
    Dim oFinalist As Object
    Call ReadFinalist(nProNo, oFinalist)

    ' ���L�^�ƕW���L�^
    Dim nRecord As Long
    nRecord = VLookupArea(nFinalNo, "�I�茠���L�^", "�L�^")
    Dim nQualify As Long
    nQualify = VLookupArea(nProNo, "�I�茠��ڋ敪", "�W���L�^")

    ' �����i�o�҂��o��
    Call WriteFinalist(nFinalNo, oFinalist, nRecord, nQualify)

    ' �C�x���g�������ĊJ
    Call EventChange(True)
End Sub

'
' �����i�o�҂�ǂݍ���
'
' nProNo            IN      ��ڔԍ�
' oFinalist         OUT     �����i�o�҂̍s�ԍ��z��
'
Private Sub ReadFinalist(nProNo As Integer, oFinalist As Object)
    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(CStr(nProNo))

    Set oFinalist = CreateObject("Scripting.Dictionary")

    Dim nOrder As Integer

    If IsNameExists(sName) Then
        Dim vProNo As Range
        For Each vProNo In GetRange(sName)
            If STrimAll(GetOffset(vProNo, GetRange("Header����").Column).Value) <> "" Then
                nOrder = STrimAll(GetOffset(vProNo, GetRange("Header����").Column).Value)
                ' �����l���܂ŕۑ�
                If nOrder > 0 And nOrder <= ���[�X��� Then
                    oFinalist.Add nOrder, vProNo
                End If
            End If
        Next vProNo
    End If
End Sub

'
' �����i�o�҂��o�͂���
'
' nProNo            IN      ��ڔԍ�
' oFinalist         IN      �����i�o�҂̍s�ԍ��z��
' nRecord           IN      ���L�^
' nQualify          IN      �W���L�^
'
Private Sub WriteFinalist(nProNo As Integer, oFinalist As Object, nRecord As Long, nQualify As Long)
    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(CStr(nProNo))

    Dim nLane As Integer
    Dim nOrder As Integer
    Dim vCell As Range
    
    If IsNameExists(sName) Then
        Dim vProNo As Range
        For Each vProNo In GetRange(sName)
            ' ���[����
            nLane = GetOffset(vProNo, GetRange("Header���[��").Column).Value
            ' ���[�����珇�ʂ��擾
            nOrder = GetOrderByLane(GetCenterLane(���[�X���, �ŏ����[���ԍ�), nLane)
            ' �\�I�̍s���擾
            Set vCell = oFinalist.Item(nOrder)
            
            GetOffset(vProNo, Range("Prog����").Column).Value = GetOffset(vCell, Range("Prog����").Column).Value
            GetOffset(vProNo, Range("Prog����").Column).Value = GetOffset(vCell, Range("Prog����").Column).Value
            GetOffset(vProNo, Range("Prog�敪").Column).Value = GetOffset(vCell, Range("Prog�敪").Column).Value
            GetOffset(vProNo, Range("Prog�\���݋L�^").Column).Value = GetOffset(vCell, Range("Prog����").Column).Value
            GetOffset(vProNo, Range("Prog���L�^").Column).Value = nRecord
            GetOffset(vProNo, Range("Prog�W���L�^").Column).Value = nQualify
            
        Next vProNo
    End If
End Sub

'
' ���[�����珇�ʂ��Z�o
'
' nCenterLane       IN      �Z���^�[���[��
' nLane             IN      ���[���ԍ�
'
Private Function GetOrderByLane(nCenterLane As Integer, nLane As Integer) As Integer
    Dim nNum As Integer
    nNum = nLane - nCenterLane
    If nNum <= 0 Then
        GetOrderByLane = 2 * (1 - nNum) - 1
    Else
        GetOrderByLane = 2 * nNum
    End If
End Function
