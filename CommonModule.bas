Attribute VB_Name = "CommonModule"
Option Explicit    ''���ϐ��̐錾����������

' ��
Public Const �w����� As String = "���{��s�w�����j���Z���"
Public Const �}�X�^�[�Y��� As String = "���{��}�X�^�[�Y���"
Public Const �w�}��� As String = "�w���}�X�^�[�Y���"
Public Const �s����� As String = "���{��s���̈���"
Public Const �I�茠��� As String = "���{��I�茠���j���"
Public Const �����L�^�� As String = "�Z���H�������j�L�^��"

Public Const �g�b�v�y�[�W�V�[�g As String = "�g�b�v�y�[�W"
Public Const �G���g���[�V�[�g As String = "�G���g���[�ꗗ"
Public Const �G���g���[�e�[�u�� As String = "�G���g���[�e�[�u��"
Public Const �v���O�����V�[�g As String = "�v���O����"
Public Const �t�H�[�}�b�g�V�[�g As String = "�v���O�����t�H�[�}�b�g"
Public Const �L�^��ʃV�[�g As String = "�L�^���"
Public Const �ݒ�e��V�[�g As String = "�ݒ�e��"

Public Const ���ϕ����g�� As Integer = 3        ' ���ϕ��������ɂ���g��
Public Const �l�ő�s�� As Integer = 2        ' �l�̐\���ݍs��
Public Const �����[�ő�s�� As Integer = 24     ' �����[�̍ő�\���ݍs��
Public Const �y�[�W���[�X�� As Integer = 5      ' �P�y�[�W�̃��[�X��

Public Const �I�薼�u�����N As String = "�@�@�D�@�@�D�@�@�D"
Public Const �^�C���u�����N As String = "�@�@�F�@�@�D  "
Public Const ���ʃu�����N As String = "�@�@"
Private Const ARRAYSIZE = 10000

Public Type RaceNumber
    nProNo As Integer
    nRance As Integer
End Type

'
' �C�x���g�̔����A��ʕ`���On/Off
'
' bFlag     IN      True�F�ĊJ�^False�F�}��
'
Public Sub EventChange(bFlag As Boolean)
    With Application
        If bFlag Then
            .EnableEvents = True                    ' �C�x���g�̔������ĊJ����
            .ScreenUpdating = True                  ' �`��̍X�V���s��
            .Calculation = xlCalculationAutomatic   ' �Z���l�̎����v�Z
        Else
            .EnableEvents = False                   ' �C�x���g�}������
            .ScreenUpdating = False                 ' �`��̍X�V��}������
            .Calculation = xlCalculationManual      ' �Z���l�̎蓮�v�Z
        End If
    End With
End Sub

'
' �V�[�g�̕ی�
'
' �V�[�g�̕ی�^�������s���B
' �ݒ�O�̃V�[�g�̕ی��Ԃ�Ԃ��B
'
' bFlag         IN      True�F�ی�^False�F����
' oWorkSheet    IN      �ی삷��V�[�g�I�u�W�F�N�g
'
Public Function SheetProtect(bFlag As Boolean, Optional oWorkSheet As Worksheet = Nothing) As Boolean

    If oWorkSheet Is Nothing Then
        Set oWorkSheet = ActiveSheet
    End If

    If ActiveSheet.ProtectContents = True Then
        SheetProtect = True
    Else
        SheetProtect = False
    End If

    If bFlag Then
        oWorkSheet.Protect DrawingObjects:=True, Contents:=True, _
            Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Else
        oWorkSheet.Unprotect
    End If

End Function

'
' �I�[�g�t�B���^�̐ݒ�
'
' sName         IN      �͈̖͂��O
' bFlag         IN      True�F�t�B���^�\���^False�F����
'
Public Sub SetAutoFilter(sName As String, bFlag As Boolean)
    If GetRange(sName).Parent.AutoFilterMode <> bFlag Then
        GetRange(sName).AutoFilter
    End If
End Sub

'
' �Z����I��
'
Public Sub SetForcusTop(Optional sRange As String = "$A$1")
    Range(sRange).Select
End Sub

'
' �Z����I��
'
Public Sub SetForcus(sAddress As String)
    Range(sAddress).Select
End Sub

'
' �^�C�g���o�[�̕ύX
'
' sTitle    IN      �^�C�g���o�[������
'
Public Sub SetTitleMenu(sTitle As String)
    Dim bFlag As Boolean
    bFlag = Application.EnableEvents
    If bFlag = False Then
        Call EventChange(True)
    End If
    Application.Caption = sTitle
    DoEvents
    If bFlag = False Then
        Call EventChange(False)
    End If
End Sub

'
' �������ڋ敪��Ԃ�
'
Public Function GetMaster(sGameName As String)
    
    GetMaster = VLookupArea(sGameName, "�ݒ�e��", "��ڋ敪�͈͖�")

End Function

'
' �󔒒u��
'
' �S�p�󔒁A�A�������󔒂𔼊p�󔒂P�ɕϊ�
'
' sStr          IN      ������
'
Public Function STrim(sStr) As String
    Dim sTemp As String
    sTemp = sStr
    sTemp = Replace(sTemp, "�@", " ")
    
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")
    
    '���K�\���̎w��
    With oReg
        .Pattern = "[ ]+"     '�p�^�[�����w��
        .IgnoreCase = False     '�啶���Ə���������ʂ��邩(False)�A���Ȃ���(True)
        .Global = True          '������S�̂��������邩(True)�A���Ȃ���(False)
    End With
    sTemp = oReg.Replace(sTemp, " ")
    
    STrim = RTrim(LTrim(sTemp))
End Function

'
' �󔒒u��
'
' �S�p�󔒁A���p�󔒂��Ȃ���
'
' sStr          IN      ������
'
Public Function STrimAll(sStr As String) As String
    STrimAll = Replace(Replace(sStr, "�@", ""), " ", "")
End Function


'
' �V�[�g�̑��݃`�F�b�N�t���A�N�e�B�x�[�g
'
' �V�[�g�̑��݂��m�F��Visible��True�łȂ����True�ɂ��Ă���
' Activate����
'
' sSheetName    IN      �V�[�g��
'
Public Function SheetActivate(sSheetName As String) As Variant
    If IsSheetExists(sSheetName) Then
        If Worksheets(sSheetName).Visible <> True Then
            Worksheets(sSheetName).Visible = True
        End If
        Worksheets(sSheetName).Activate
        Set SheetActivate = ActiveSheet
    Else
        MsgBox "�u" & sSheetName & "�v�V�[�g�����݂��܂���B" & vbCrLf & _
                "�������t�@�C�������g�����������B", vbOKOnly
        End
    End If
End Function


'
' �V�[�g��Visible��Ԃ�
'
' xlSheetVisible/True(-1)
' xlSheetHidden/False(0)
' xlSheetVeryHidden(2)
' Empty(3)
'
Public Function GetSheetVisible(sSheetName As String) As Variant
    If IsSheetExists(sSheetName) Then
        GetSheetVisible = Worksheets(sSheetName).Visible
    Else
        GetSheetVisible = 3
    End If
End Function


'
' �V�[�g�̑��݃`�F�b�N�t����\��
'
'
' sSheetName    IN      �V�[�g��
'
Public Sub SheetVisible(sSheetName As String, vVisible As Variant)
    If IsSheetExists(sSheetName) Then
        Worksheets(sSheetName).Visible = vVisible
    End If
End Sub


'
' �V�[�g�̑��݃`�F�b�N
'
' sSheetName        IN      �V�[�g��
'
Public Function IsSheetExists(sSheetName As String) As Boolean
    IsSheetExists = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = sSheetName Then
            IsSheetExists = True
        End If
    Next ws
End Function

'' Dictionary���Q�ƈ����ɂ��A������\�[�g����j��I�v���V�[�W���B
'' Variant�̓񎟌��z��̂܂܈��������������A�����ɂł��Ȃ��悤�Ȃ̂Ŕz����g�p���Ă���
'' Key��Value��Long�^�݂̂����e
Public Sub DictQuickSort(ByRef dic As Object, Optional sIndex As String = "Key")
  
    Dim nIndex As Integer
    If sIndex = "Key" Then
        nIndex = 0
    Else
        nIndex = 1
    End If
  
    Dim i As Long
    Dim j As Long
    Dim dicSize As Long
    Dim varTmp(ARRAYSIZE, 2) As Long
  
    dicSize = dic.Count
  
    ' Dictionary���󂩁A�T�C�Y��1�ȉ��ł���΃\�[�g�s�v
    If dic Is Nothing Or dicSize < 2 Then
        Exit Sub
    End If
  
    ' Dictionary����񌳔z��ɓ]��
    i = 0
    Dim Key As Variant
    For Each Key In dic
        varTmp(i, 0) = Key
        varTmp(i, 1) = dic(Key)
        i = i + 1
    Next
  
    '�N�C�b�N�\�[�g
    Call QuickSort(varTmp, 0, dicSize - 1, nIndex)
  
    dic.RemoveAll
  
    For i = 0 To dicSize - 1
        dic(varTmp(i, 0)) = varTmp(i, 1)
    Next
End Sub


'' Long�^�̓񎟌��z����󂯎��A����̂Q��ځiValue�j�ŃN�C�b�N�\�[�g����
Private Sub QuickSort(ByRef targetVar() As Long, ByVal min As Long, ByVal max As Long, nIndex As Integer)
    Dim i, j As Long
    Dim tmp As Long
    Dim pivot As Long
    
    If min < max Then
        i = min
        j = max
        pivot = med3(targetVar(i, nIndex), targetVar(Int(i + j / 2), nIndex), targetVar(j, nIndex))
        Do
            Do While targetVar(i, nIndex) < pivot
                i = i + 1
            Loop
            Do While pivot < targetVar(j, nIndex)
                j = j - 1
            Loop
            If i >= j Then Exit Do
            
            tmp = targetVar(i, 0)
            targetVar(i, 0) = targetVar(j, 0)
            targetVar(j, 0) = tmp
        
            tmp = targetVar(i, 1)
            targetVar(i, 1) = targetVar(j, 1)
            targetVar(j, 1) = tmp
        
            i = i + 1
            j = j - 1
        
        Loop
        Call QuickSort(targetVar, min, i - 1, nIndex)
        Call QuickSort(targetVar, j + 1, max, nIndex)
        
    End If
End Sub

'' Long, y, z ����������r����Ԗڂ̂��̂�Ԃ�
Private Function med3(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    If x < y Then
        If y < z Then
            med3 = y
        ElseIf z < x Then
            med3 = x
        Else
            med3 = z
        End If
    Else
        If z < y Then
            med3 = y
        ElseIf x < z Then
            med3 = x
        Else
            med3 = z
        End If
    End If
End Function

'
' �z��ɒl�����݂��邩�m�F����
'
' vAry          IN  �z��
' vValue        IN  �l
'
Public Function isAryExist(vAry As Variant, vValue As Variant) As Boolean

    Dim vTemp As Variant
    For Each vTemp In vAry
        If vTemp = vValue Then
            isAryExist = True
            Exit Function
        End If
    Next
    isAryExist = False

End Function

'
' ���O�`�F�b�N�t��Range�I�u�W�F�N�g�擾
'
' sName             IN      ���O
'
Public Function GetRange(sName As String) As Range
    If IsNameExists(sName) Then
        Set GetRange = Range(sName)
    Else
        MsgBox "���O�u" & sName & "�v����`����Ă��܂���B" & vbCrLf & _
                "�������t�@�C�������g�����������B", vbOKOnly
        End
    End If
End Function

'
' ���O�̑��݊m�F
'
' sName             IN      ���O
'
Public Function IsNameExists(sName As String) As Boolean
    IsNameExists = False
    Dim vName As Variant
    For Each vName In ActiveWorkbook.Names
        If vName.Name = sName Then
            IsNameExists = True
            Exit For
        End If
    Next
End Function

'
' ���O�̒�`
'
' sName             IN      ���O
' sRange            IN      �����W�͈�(A1�`��)
'
Public Sub DefineName(sName As String, sRange As String)
    If IsNameExists(sName) Then
        ActiveWorkbook.Names(sName).Delete
    End If
    ActiveWorkbook.Names.Add Name:=sName, RefersTo:="=" & sRange
    ActiveWorkbook.Names(sName).Comment = ""
End Sub

'
' ���O�폜
'
' sRegStr           IN      �폜���閼�O�̕�����
'
Public Sub DeleteName(sRegStr As String)
    Dim vName As Variant
    For Each vName In ActiveWorkbook.Names
        If vName.Name Like sRegStr Then
            vName.Delete
        End If
    Next
End Sub

'
' ����s�͈͎̔擾
'
' �w�肵���Z������ŉE�͈̔͂̃A�h���X��Ԃ�
'
' ��F $A$1 -> $A$1:$F$1
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Public Function ColumnRange(sTopAddress As String) As Range

    Set ColumnRange = Range(Range(sTopAddress), _
                    Range(sTopAddress).End(xlToRight))

End Function

'
' ����s�͈͎̔擾
'
' �w�肵���Z������ŉE�͈̔͂̃A�h���X��Ԃ�
'
' ��F $A$1 -> $A$1:$F$1
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Public Function ColumnRangeAddress(sTopAddress As String) As String

    ColumnRangeAddress = ColumnRange(sTopAddress).Address

End Function

'
' �����͈͎̔擾
'
' �w�肵���Z������ŉ��w�͈̔͂̃A�h���X��Ԃ�
'
' ��F $A$1 -> $A$1:$A$50
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Public Function RowRange(sTopAddress As String) As Range
    Set RowRange = Range(Range(sTopAddress), Range(sTopAddress).End(xlDown))
End Function

'
' �����͈͎̔擾
'
' �w�肵���Z������ŉ��w�͈̔͂̃A�h���X��Ԃ�
'
' ��F $A$1 -> $A$1:$A$50
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Public Function RowRangeAddress(sTopAddress As String) As String

    RowRangeAddress = RowRange(sTopAddress).Address

End Function

'
' �s��͈͎̔擾
'
' �w�肵���Z������ŉ��w�A�ۉE�[�͈͂̃A�h���X��Ԃ�
'
' ��F $A$1 -> $A1$1:$F$50
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Public Function TableRange(sTopAddress As String) As Range

    Dim oRng As Range
    Set oRng = Range(Range(sTopAddress), Range(sTopAddress).End(xlDown))
    Set TableRange = Range(oRng, oRng.End(xlToRight))

End Function

'
' �s��͈͎̔擾
'
' �w�肵���Z������ŉ��w�A�ۉE�[�͈͂̃A�h���X��Ԃ�
'
' ��F $A$1 -> $A1$1:$F$50
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Public Function TableRangeAddress(sTopAddress As String) As String

    TableRangeAddress = TableRange(sTopAddress).Address

End Function

'
' �͈͂̍ō���ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Public Function GetAreaLeftColumn(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaLeftColumn = oRange.Column
End Function

'
' �͈͂̍ŉE��ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Public Function GetAreaRightColumn(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaRightColumn = oRange.Column + oRange.Columns.Count - 1
End Function

'
' �͈͂̍ŏ�s�ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Public Function GetAreaTopRow(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaTopRow = oRange.Row
End Function

'
' �͈͂̍ŉ��s�ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Public Function GetAreaBottomRow(sName As String) As Integer
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaBottomRow = oRange.Row + oRange.Rows.Count - 1
End Function

'
' ��Z������I�t�Z�b�g�ʒu�̒l�Z����Ԃ�
'
' vCell         IN      ��Z��
' nColumn       IN      ��ԍ�
'
Public Function GetOffset(vCell As Variant, nColumn As Integer, Optional nRow As Integer = 0) As Range
    Set GetOffset = vCell.Offset(nRow, nColumn - vCell.Column)
End Function

'
' ��Z�������̃I�t�Z�b�g�ʒu�̒l�Z����Ԃ�
'
' vCell         IN      ��Z��
' nColumn       IN      ��ԍ�
'
Public Function GetRowOffset(vCell As Variant, nRow As Integer, Optional nColumn As Integer = 0) As Range
    Set GetRowOffset = vCell.Offset(nRow - vCell.Row, nColumn)
End Function

'
' �͈͂���w�肵�����O�̗�ԍ���Ԃ�
' ���݂��Ȃ��ꍇ�̓[����Ԃ�
'
' sName      IN      �͈͖�
' sColName   IN      �J������
'
Public Function GetColIdx(sName As String, sColName As String) As Integer
    Dim nIndex As Integer
    nIndex = 1
    Dim oCell As Range
    For Each oCell In GetRange(sName).Rows(1).Columns
        If STrimAll(oCell.Value) = sColName Then
            GetColIdx = nIndex
            Exit Function
        End If
        nIndex = nIndex + 1
    Next oCell
    GetColIdx = 0
End Function

'
' �͈͂̃w�b�_�s�������L�[����擾����
'
' �͈͂̂P��ڂ̃w�b�_�s���������Ԃ�
'
' sName      IN      �͈͖�
'
Public Function GetAreaKeyData(sName As String) As Range
    Dim oRange As Range
    Set oRange = Range(sName)
    Set GetAreaKeyData = oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, 1).Rows()
End Function

'
' �͈͂̃L�[�񖼂�Ԃ�
'
' �͈͂̂P��ڂ̖��O��Ԃ�
'
' sName      IN      �͈͖�
'
Public Function GetAreaKeyName(sName As String) As Variant
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaKeyName = oRange.Cells(1, 1).Value
End Function

'
' �͈͂���w�肵���l�ɑΉ�����J�����l��Ԃ�
'
' vValue     IN      �����l
' sName      IN      �͈͖�
' sColName   IN      �J������
' bFlag      IN      VLookuUp�֐��̌����̌^�iFalse:���S��v�^True:��ԋ߂��f�[�^�j
'            OUT     ��v�����l
'
Public Function VLookupArea(vValue As Variant, sName As String, sColName As String, Optional bFlag As Boolean = False) As Variant
On Error GoTo ErrorHandler_VLookupArea
    
    VLookupArea = Application.WorksheetFunction.VLookup(vValue, GetRange(sName), GetColIdx(sName, sColName), bFlag)
    Exit Function

ErrorHandler_VLookupArea:
    MsgBox "�͈͖��F" & sName & vbCrLf & _
            "���ږ��F" & sColName & vbCrLf & _
            "�����l�F" & CStr(vValue) & vbCrLf & _
            "��������܂���B�V�[�g���m�F���Ă��������B", vbOKOnly
    End
End Function

'
' �����G���A�̂ǂ̕�������Ԃ�
'
' 1:�擪
' 2:����
' 3:����
'
' oCell     IN      �Z��
'
Public Function CheckMergeArea(oCell As Range)
    Dim nIdx As Integer
    nIdx = 1
    Dim nMax As Integer
    nMax = oCell.MergeArea.Rows.Count
    Dim vCell As Range
    For Each vCell In oCell.MergeArea.Rows
        If vCell.Address = oCell.Address Then
            If nIdx = 1 Then
                CheckMergeArea = 1
            ElseIf nIdx = nMax Then
                CheckMergeArea = 2
            Else
                CheckMergeArea = 3
            End If
        End If
        nIdx = nIdx + 1
    Next vCell
End Function

'
' �r��������
'
Public Sub SetBorder(oRange As Range)
    With oRange
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub


'
' ���W���[���̓Ǎ���
'
' oWorkBook  IN      WorkBook
' sPath      IN      �t�H���_�p�X
'
Public Sub ImportAll(oWorkBook As Workbook, sPath As String)
    On Error Resume Next
    
    Dim oFso        As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim sArModule() As String                   '// ���W���[���t�@�C���z��
    Dim sModule                                 '// ���W���[���t�@�C��
    Dim sExt        As String                   '// �g���q
    Dim iMsg                                    '// MsgBox�֐��߂�l
    
    iMsg = MsgBox("�����̃��W���[���͏㏑�����܂��B��낵���ł����H", vbOKCancel, "�㏑���m�F")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// �S���W���[���̃t�@�C���p�X���擾
    Dim FileList As Collection
    Set FileList = GetFiles(sPath, "\*")
        
    '// �S���W���[�������[�v
    For Each sModule In FileList
        '// �g���q���������Ŏ擾
        sExt = LCase(oFso.GetExtensionName(sModule))
        
        '// �g���q��cls�Afrm�Abas�̂����ꂩ�̏ꍇ
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// �������W���[�����폜
            Call oWorkBook.VBProject.VBComponents.Remove(oWorkBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// ���W���[����ǉ�
            Call oWorkBook.VBProject.VBComponents.Import(sModule)
            '// Import�m�F�p���O�o��
            Debug.Print sModule
        End If
    Next
End Sub

'
' ���W���[���̏o��
'
' oWorkBook  IN      WorkBook
' sPath      IN      �t�H���_�p�X
'
Public Sub ExportAll(oWorkBook As Workbook, sPath As String)
    On Error Resume Next
    
    Dim sFileName As String
    With ActiveWorkbook.VBProject
        Dim i As Integer
        For i = 1 To .VBComponents.Count
            Debug.Print "Type: " & .VBComponents(i).Type
            Debug.Print "Name: " & .VBComponents(i).Name
            If .VBComponents(i).Type = 1 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".bas"
                .VBComponents(i).Export sFileName
            ElseIf .VBComponents(i).Type = 2 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".cls"
                .VBComponents(i).Export sFileName
            End If
        Next i
    End With
End Sub

'
' �t�H���_��I������_�C�A���O��\��
'
' �����̓}�N���t�@�C���Ɠ����t�H���_���J��
'
' SelectDir     OUT     �t�H���_���t���p�X
'
Public Function SelectDir()
    ' �V�����t�@�C�����J��
    Dim FileSysObj As Object
    Dim sPathName As String
    Set FileSysObj = CreateObject("Scripting.FileSystemObject")
    sPathName = FileSysObj.GetParentFolderName(ActiveWorkbook.FullName)
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = sPathName + "\"
        .AllowMultiSelect = False
        If .Show = True Then
            SelectDir = .SelectedItems(1)
        Else
            MsgBox "�����𒆎~���܂��B"
            End
        End If
    End With

End Function

'
' ����̃t�H���_�̃t�@�C�����J��
'
' sPathName     IN      �t�H���_�p�X
' sExt          IN      �g���q�̎w��
' cFileList     OUT     �t�@�C�����X�g
'
Public Function GetFiles(sPathName As String, sExt As String) As Collection
    Dim sFile As String
    Dim cFileList As Collection
    Set cFileList = New Collection
    sFile = Dir(sPathName & sExt)
    Do While sFile <> ""
        cFileList.Add Item:=sFile
        sFile = Dir()
    Loop
    Set GetFiles = cFileList
End Function

'
' �v���O�����␳�p�F�Z���̏C��
'
' sRange     IN      �Z���̃A�h���X
' sValue     IN      �ύX����l
'
Public Sub ModCell(sRange As String, sValue As String)

    With Range(sRange)
        .Value = sValue
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = -16776961
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

End Sub

'
' �v���O�����␳�p�F�v���O�����ԍ��̃A�h���X��T��
'
' nProNo     IN      �v���O�����ԍ�
' sName      IN      �I�薼
' sColName   IN      �J�������i���[�XNo�A�g�A���[���j
'            OUT     �A�h���X������
'
Public Function SearchCell(nProNo As Integer, sName As String, sColName As String) As String

    Dim vCell As Range
    For Each vCell In GetAreaKeyData(�G���g���[�e�[�u�� & "[�v��No]")
        If GetOffset(vCell, Range(�G���g���[�e�[�u�� & "[�v��No]").Column).Value = nProNo And _
            GetOffset(vCell, Range(�G���g���[�e�[�u�� & "[�I�薼]").Column).Value = sName Then
            SearchCell = GetOffset(vCell, Range(�G���g���[�e�[�u�� & "[" & sColName & "]").Column).Address
            Exit Function
        End If
    Next vCell

End Function

'
' �v���O�����␳�p�F�v���O�����ԍ��̃A�h���X��T��
'
' nProNo     IN      �v���O�����ԍ�
' sTeam      IN      ������
' sClass     IN      �敪
' sColName   IN      �J�������iProg�����j
'            OUT     �A�h���X������
'
Public Function SearchRelayCell(nProNo As Integer, _
sTeam As String, sClass As String, sColName As String) As String

    Dim sName As String
    sName = "�v���O�����ԍ�" & Trim(CStr(nProNo))

    Dim nLane As Integer
    Dim nOrder As Integer
    Dim vCell As Range
    
    If IsNameExists(sName) Then
        Dim vProNo As Range
        For Each vProNo In GetRange(sName)
            If GetOffset(vProNo, Range("Prog����").Column).Value = sTeam And _
                GetOffset(vProNo, Range("Prog�敪").Column).Value = sClass Then
                SearchRelayCell = GetOffset(vProNo, Range(sColName).Column).Address
                Exit Function
            End If
        Next vProNo
    End If

End Function

'
' �v���O�������͗p�F�L�^��ʂ̎�ڑI��
'
' nProNo     IN      �v���O�����ԍ�
' nHeat      IN      �g
'
Public Sub SetRace(nProNo As Integer, nHeat As Integer)
    GetRange("�L�^��ʎ�ڔԍ�").Value = nProNo
    GetRange("�L�^��ʑg").Value = nHeat
End Sub

'
' �v���O�������͗p�F�L�^��ʂ̃^�C������
'
' nIndex        IN      ����
' nLean         IN      ���[��
' sTime         IN      ����
' sAdditional   IN      ���l
'
Public Sub SetLean(nIndex As Integer, nLean As Integer, sTime As String, Optional sAdditional As String = "")
    GetRange("�L�^��ʃ��[��").Rows(nIndex).Value = nLean
    GetRange("�L�^��ʃ^�C��").Rows(nIndex).Value = sTime
    If sAdditional <> "" Then
        GetRange("�L�^��ʔ��l").Rows(nIndex).Value = sAdditional
    End If
End Sub


