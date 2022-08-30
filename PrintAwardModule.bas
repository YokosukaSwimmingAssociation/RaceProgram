Attribute VB_Name = "PrintAwardModule"
Option Explicit    ''���ϐ��̐錾����������
'
' �܏���
'
' �w�肵�����[�XNo�ɑ��݂���ProNo�̏܏���������
'
Public Sub �܏���()

    ' �C�x���g������}��
    Call EventChange(False)

    ' �L�^��ʃV�[�g��ۑ�
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    ' ���[�X�ԍ�
    Dim nRaceNo As Integer
    nRaceNo = GetRange("�L�^��ʃ��[�XNo").Value

    ' �܏���
    Call PrintAwardByRace(nRaceNo)

    ' �L�^��ʃV�[�g�ɖ߂�
    oWorkSheet.Activate

    ' �C�x���g�������ĊJ
    Call EventChange(True)

End Sub

'
' ���[�X�ԍ��w��܏���
'
' nRaceNo           IN      ���[�XNo
'
Private Function PrintAwardByRace(nRaceNo As Integer)

    ' ���[�X�̖��O�擾
    Dim sName As String
    sName = "�v���O�������[�X" & Trim(Str(nRaceNo))
    If Not IsNameExists(sName) Then
        MsgBox "����Ώۂ��擾�ł��܂���ł����B" & vbCrLf & _
                "���[�XNo���������w�肳��Ă��邩�m�F���Ă��������B" & vbCrLf & _
                "�������w�肳��Ă���ꍇ�́A�v���O�������O��`�����s���Ă݂Ă��������B", vbOKOnly
        End
    End If

    ' �܏���
    Call PrintAwardByName(sName)

End Function

'
' ���O�w��܏���
'
' sName             IN      ���[�X�̖��O��`
'
Private Function PrintAwardByName(sName As String)

    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    Dim nProNo As Integer
    Dim vRaceNo As Variant
    For Each vRaceNo In GetRange(sName)
        nProNo = GetOffset(vRaceNo, GetRange("Header�v��No").Column).Value
        ' �ŏ��̂P�񂾂����s
        If Not oProNo.Exists(nProNo) Then
            ' �܏����Ώۂ��m�F
            If CheckTarget(nProNo) Then
                ' �܏�����
                Call PrintAwardByProNo(nProNo)
            End If
            oProNo.Add nProNo, 1
        End If
    Next vRaceNo

End Function


'
' ����Ώۂ����m�F����
'
' �w�}���͊w���̂ݑΏہi��ڋ敪����Ώۂ��擾�j
' �s�����͂��ׂđΏ�
' �I�茠���͌����̂ݑΏہi��ڋ敪����Ώۂ��擾�j
'
'
' nProNo            IN      �v��No
'
Private Function CheckTarget(nProNo As Integer) As Boolean

    ' ��
    Dim sGameName As String
    sGameName = GetRange("��").Value

    If sGameName = �I�茠��� Then
        ' �\�I�E�����̊m�F
        If VLookupArea(nProNo, "�I�茠��ڋ敪", "�\�I�^����") <> "�\�I" Then
            CheckTarget = True
        Else
            CheckTarget = False
        End If
    ElseIf sGameName = �s����� Then
        CheckTarget = True
    ElseIf sGameName = �w�}��� Then
        ' ��ڂ̑����擾
        If VLookupArea(nProNo, "�w�}��ڋ敪", "���敪") = "�w��" Then
            CheckTarget = True
        Else
            CheckTarget = False
        End If
    Else
        MsgBox "�����������w�肳��Ă��܂���B", vbOKOnly
        End
    End If

End Function

'
' �܏���������
'
' �w�肵���v��No�̒���1�ʁ`3�ʂ̏܏���������
'
' nProNo            IN      �v��No
'
Private Sub PrintAwardByProNo(nProNo As Integer)

    ' ��
    Dim sGameName As String
    sGameName = GetRange("��").Value
    
    ' ��ڋ敪���擾
    Dim sMasterName As String
    sMasterName = GetMaster(GetRange("��").Value)
    
    Dim sRaceClass As String ' �敪
    sRaceClass = VLookupArea(nProNo, sMasterName, "��ڋ敪")
    Dim sGender As String ' ����
    sGender = VLookupArea(nProNo, sMasterName, "����")
    Dim sDistance As String ' ����
    sDistance = Replace(VLookupArea(nProNo, sMasterName, "����"), "M", "")
    Dim sStyle As String ' ���
    sStyle = VLookupArea(nProNo, sMasterName, "���")
    Dim nMaxOrder As Integer ' �o�͂��鏇��
    nMaxOrder = VLookupArea(sGameName, "�ݒ�e��", "�܏󏇈�")
    
    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(CStr(nProNo))
    
    Dim vProNo As Variant
    Dim nOrder As Integer
    For Each vProNo In GetRange(sName)
        nOrder = Val(GetOffset(vProNo, GetRange("Prog����").Column).Value)
        If nOrder >= 1 And nOrder <= nMaxOrder Then
            Call PrintAwardByLine(sGameName, vProNo, sRaceClass, sGender, sDistance, sStyle)
        End If
    Next vProNo
    
End Sub

'
' �s�w��ŏ܏���������
'
' �w�肵���s�̃��R�[�h���������
'
' sGameName         IN      ��
' vProNo            IN      ProNo
' sRaceClass        IN      ��ڋ敪
' sGender           IN      ����
' sDistance         IN      ����
' sStyle            IN      ���
'
Private Sub PrintAwardByLine(sGameName As String, _
vProNo As Variant, _
sRaceClass As String, _
sGender As String, _
sDistance As String, _
sStyle As String)
   
    ' ���ʂ̐ݒ�
    GetRange("�܏󏇈�").Value = GetOffset(vProNo, Range("Prog����").Column).Value
    GetRange("�܏�^�C��").Value = GetOffset(vProNo, Range("Prog����").Column).Value
    If GetOffset(vProNo, Range("Prog���l").Column).Value = "���V" Then
        GetRange("�܏���V").Value = "���V"
    Else
        GetRange("�܏���V").Value = ""
    End If
    GetRange("�܏󎁖�").Value = GetOffset(vProNo, Range("Prog����").Column).Value
    GetRange("�܏󏊑�").Value = GetOffset(vProNo, Range("Prog����").Column).Value
   
    If GetRange("�܏�^�C��").Value >= 10000 Then
        GetRange("�܏�^�C��").NumberFormatLocal = "#""��""##""�b""##"
    Else
        GetRange("�܏�^�C��").NumberFormatLocal = "##""�b""##"
    End If

    ' ���ŗL�̐ݒ�
    If sGameName = �I�茠��� Then
        Call SetAwardValForSenshuken(sGender, sDistance, sStyle)
    ElseIf sGameName = �s����� Then
        Call SetAwardValForShimin(sRaceClass, sGender, sDistance, sStyle, _
                GetOffset(vProNo, Range("Prog�敪").Column).Value)
    Else
        Call SetAwardValForGakudo(sRaceClass, sGender, sDistance, sStyle)
    End If

    ' ���
    Call PrintAward

End Sub

'
' �܏���
'
Private Sub PrintAward()
    ' �v���r���[�L��
    Dim bPreview As Boolean
    If GetRange("������v���r���[").Value = "����" Then
        bPreview = True
    Else
        bPreview = False
    End If
    
    ' �v�����^��
    Dim sPrinterName As String
    sPrinterName = GetRange("�v�����^��").Value
    If sPrinterName = "" Then
        MsgBox "�v�����^�����ݒ肳��Ă��܂���B", vbOKOnly
        End
    End If
    
    ' ���
    GetRange("�܏󎁖�").Parent.PrintOut _
        Copies:=1, Collate:=True, IgnorePrintAreas:=False, Preview:=bPreview, _
        ActivePrinter:=sPrinterName

End Sub


'
' �w�����܏�ϐ��ݒ�
'
Private Sub SetAwardValForGakudo( _
sRaceClass As String, _
sGender As String, _
sDistance As String, _
sStyle As String)
    GetRange("�܏��ڋ敪").Value = sRaceClass & sGender
    GetRange("�܏󋗗�").Value = sDistance
    GetRange("�܏���").Value = sStyle
End Sub


'
' �s�����܏�ϐ��ݒ�
'
Private Sub SetAwardValForShimin( _
sRaceClass As String, _
sGender As String, _
sDistance As String, _
sStyle As String, _
sClass As String)
    If sRaceClass = "�N��敪" Then
        GetRange("�܏��ڋ敪").Value = sGender
        GetRange("�܏��ڋ����敪").Value = sDistance & "�l" & sStyle & "�@" & sClass
    Else
        GetRange("�܏��ڋ敪").Value = sRaceClass & sGender
        GetRange("�܏��ڋ����敪").Value = sDistance & "�l" & sStyle
    End If
    GetRange("�܏���񐔂P").Value = GetRange("����").Value
    GetRange("�܏���񐔂Q").Value = GetRange("����").Value
    GetRange("�܏�N").Value = GetRange("�����N").Value
    GetRange("�܏�").Value = GetRange("��").Value
    GetRange("�܏��").Value = GetRange("����").Value

    ' �J�������̕ύX
    If sStyle Like "*�����[" Then
        GetRange("�܏󎁖�").ColumnWidth = 1.13
        GetRange("�܏󏊑�").ColumnWidth = 2.5
    Else
        GetRange("�܏󎁖�").ColumnWidth = 2.5
        GetRange("�܏󏊑�").ColumnWidth = 1.13
    End If

End Sub

'
' �I�茠�܏�ϐ��ݒ�
'
Private Sub SetAwardValForSenshuken( _
sGender As String, _
sDistance As String, _
sStyle As String)
    GetRange("�܏󐫕�").Value = sGender
    GetRange("�܏󋗗�").Value = sDistance
    GetRange("�܏���").Value = sStyle
    
    GetRange("�܏����").Value = GetRange("����").Value
    GetRange("�܏�N").Value = GetRange("�����N").Value
    GetRange("�܏�").Value = GetRange("��").Value
    GetRange("�܏��").Value = GetRange("����").Value
End Sub

