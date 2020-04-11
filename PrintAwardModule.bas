Attribute VB_Name = "PrintAwardModule"
'
' �܏���
'
' �w�肵�����[�XNo�ɑ��݂���ProNo�̏܏���������
'
Sub �܏���()

    ' �C�x���g������}��
    Call EventChange(False)

    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    ' ���[�X�ԍ�
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
                ' �܏����Ώۂ��m�F
                If CheckTarget(nProNo) Then
                    ' �܏�����
                    Call PrintAward(nProNo)
                End If
                oProNo.Add nProNo, 1
            End If
        Next vRaceNo
    End If

    oWorkSheet.Activate

    ' �C�x���g�������ĊJ
    Call EventChange(True)

End Sub

'
' ����Ώۂ����m�F����
'
' nProNo            IN      �v��No
'
Function CheckTarget(nProNo As Integer) As Boolean

    ' ��ڂ̋敪���擾
    Dim sGameType
    sGameType = Application.WorksheetFunction.VLookup(nProNo, Range("�w�}��ڋ敪"), 6, False)
    
    If sGameType = "�w��" Or sGameType = "�w�������[" Then
        CheckTarget = True
    Else
        CheckTarget = False
    End If

End Function

'
' �܏���������
'
' �w�肵���v��No�̒���1�ʁ`3�ʂ̏܏���������
'
' nProNo            IN      �v��No
'
Sub PrintAward(nProNo As Integer)

    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(Str(nProNo))
    
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = GetRange(sName).Parent
    
    Dim sRaceType As String ' �敪
    sRaceType = Application.WorksheetFunction.VLookup(nProNo, Range("�w�}��ڋ敪"), 2, False)
    Dim sGender As String ' ����
    sGender = Application.WorksheetFunction.VLookup(nProNo, Range("�w�}��ڋ敪"), 3, False)
    Dim sDistance As String ' ����
    sDistance = Application.WorksheetFunction.VLookup(nProNo, Range("�w�}��ڋ敪"), 4, False)
    Dim sStyle As String ' ��ږ�
    sStyle = Application.WorksheetFunction.VLookup(nProNo, Range("�w�}��ڋ敪"), 5, False)
    
    Dim nOrder As Integer
    For Each vProNo In GetRange(sName)
        nOrder = Val(vProNo.Offset(0, GetRange("Prog����").Column - vProNo.Column).Value)
        If nOrder >= 1 And nOrder <= 3 Then
            Call PrintAwardByLine(oWorkSheet, vProNo.Row, sRaceType, sGender, sDistance, sStyle)
        End If
    Next vProNo
    
End Sub

'
' �s�w��ŏ܏���������
'
' �w�肵���s�̃��R�[�h���������
'
' oWorkSheet        IN      ���[�N�V�[�g
' nRow              IN      �s�ԍ�
' sRaceType         IN      ��ڋ敪
' sGender           IN      ����
' sDistance         IN      ����
' sStyle            IN      ��ږ�
'
Sub PrintAwardByLine(oWorkSheet As Worksheet, _
nRow As Integer, _
sRaceType As String, _
sGender As String, _
sDistance As String, _
sStyle As String)
   
    GetRange("�܏��ڋ敪").Value = sRaceType & sGender
    GetRange("�܏󋗗�").Value = sDistance
    GetRange("�܏���").Value = sStyle
    GetRange("�܏󏇈�").Value = oWorkSheet.Cells(nRow, Range("Prog����").Column).Value
    GetRange("�܏�^�C��").Value = oWorkSheet.Cells(nRow, Range("Prog����").Column).Value
    GetRange("�܏���V").Value = oWorkSheet.Cells(nRow, Range("Prog���l").Column).Value
    GetRange("�܏󎁖�").Value = oWorkSheet.Cells(nRow, Range("Prog����").Column).Value
    GetRange("�܏󏊑�").Value = oWorkSheet.Cells(nRow, Range("Prog����").Column).Value

    If GetRange("�܏�^�C��").Value >= 10000 Then
        GetRange("�܏�^�C��").NumberFormatLocal = "#""��""##""�b""##"
    Else
        GetRange("�܏�^�C��").NumberFormatLocal = "##""�b""##"
    End If

    ' ���
    GetRange("�܏��ڋ敪").Parent.Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, Preview:=True

End Sub

Sub �܏󖼑O��`(Optional sValue As String = "")
    Call �w�}�܏󖼑O��`
End Sub

Sub �w�}�܏󖼑O��`(Optional sValue As String = "")
    Sheets("�w�}�܏�").Select
    ActiveSheet.Unprotect

    ' ���O�����ׂč폜
    Call DeleteName("�܏�*")

    Call SetName("�܏��ڋ敪", "$C$9")
    Call SetName("�܏󋗗�", "$G$9")
    Call SetName("�܏���", "$L$9")
    Call SetName("�܏󏇈�", "$A$13")
    Call SetName("�܏�^�C��", "$L$14")
    Call SetName("�܏���V", "$S$14")
    Call SetName("�܏󎁖�", "$C$20")
    Call SetName("�܏󏊑�", "$C$24")
 
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
 
End Sub

