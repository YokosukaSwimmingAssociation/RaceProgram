Attribute VB_Name = "CommonModule"
' ��
Public Const �w����� As String = "���{��w�����j���Z���"
Public Const �}�X�^�[�Y��� As String = "���{��}�X�^�[�Y���"
Public Const �w�}��� As String = "�w���}�X�^�[�Y���"
Public Const �s����� As String = "���{��s���̈���"
Public Const �I�茠��� As String = "���{��I�茠���j���"

Public Const �G���g���[�V�[�g As String = "�G���g���[�ꗗ"
Public Const �G���g���[�e�[�u�� As String = "�G���g���[�e�[�u��"
Public Const �v���O�����V�[�g As String = "�v���O����"
Public Const �t�H�[�}�b�g�V�[�g As String = "�v���O�����t�H�[�}�b�g"

Public Const ���[�X��� As Integer = 7       ' �P���[�X�̐l��
Public Const �ő僌�[���ԍ� As Integer = 9     ' ���[���̍ő�ԍ�
Public Const �ŏ����[���ԍ� As Integer = 3     ' ���[���̍ŏ��ԍ�
Public Const ���ϕ����g�� As Integer = 3 ' ���ϕ��������ɂ���g��

Public Const �I�薼�u�����N As String = "�@�@�D�@�@�D�@�@�D"
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
' bFlag         IN      True�F�ی�^False�F����
'
Public Sub SheetProtect(bFlag As Boolean)

    If bFlag Then
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, _
            Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Else
        ActiveSheet.Unprotect
    End If

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
    ' ���{��I�茠���j���
    If sGameName = �I�茠��� Then
        GetMaster = "�I�茠��ڋ敪"
    ' ���{��s���̈���
    ElseIf sGameName = �s����� Then
        GetMaster = "�s����ڋ敪"
    Else
        ' �w�}���
        GetMaster = "�w�}��ڋ敪"
    End If

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
' sSheetName    IN      �V�[�g��
'
Public Sub SheetActivate(sSheetName As String)
    If IsSheetExists(sSheetName) Then
        Worksheets(sSheetName).Activate
    Else
        MsgBox "�u" & sSheetName & "�v�V�[�g�����݂��܂���B" & vbCrLf & _
                "�������t�@�C�������g�����������B", vbOKOnly
        End
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
Private Function med3(ByVal x As Long, ByVal y As Long, ByVal z As Long)
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
    For Each Nm In ActiveWorkbook.Names
        If Nm.Name = sName Then
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
    For Each vNm In ActiveWorkbook.Names
        If vNm.Name Like sRegStr Then
            vNm.Delete
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
Public Function GetOffset(vCell As Variant, nColumn As Integer) As Range
    Set GetOffset = vCell.Offset(0, nColumn - vCell.Column)
End Function

'
' �͈͂���w�肵�����O�̗�ԍ���Ԃ�
'
' sName      IN      �͈͖�
' sColName   IN      �J������
'
Public Function GetAreaColumnIndex(sName As String, sColName As String) As Integer
    Dim nIndex As Integer
    nIndex = 1
    For Each vCell In GetRange(sName).Rows(1).Columns
        If STrimAll(vCell.Value) = sColName Then
            GetAreaColumnIndex = nIndex
            Exit Function
        End If
        nIndex = nIndex + 1
    Next vCell
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
Public Function VLookupArea(vValue As Variant, sName As String, sColName As String, Optional bFlag As Boolean = False) As Range
    VLookupArea = Application.WorksheetFunction.VLookup(vValue, GetRange(sName), GetAreaColumnIndex(sName, sColName), bFlag)
End Function

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
        For i = 1 To .VBComponents.Count
            Debug.Print "Type: " & .VBComponents(i).Type
            Debug.Print "Name: " & .VBComponents(i).Name
            If .VBComponents(i).Type = 1 Then
                sFileName = sPath & "\\" & .VBComponents(i).Name & ".vbs"
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

    For Each vCell In GetAreaKeyData(�G���g���[�e�[�u�� & "[�v��No]")
        If GetOffset(vCell, Range(�G���g���[�e�[�u�� & "[�v��No]").Column).Value = nProNo And _
            GetOffset(vCell, Range(�G���g���[�e�[�u�� & "[�I�薼]").Column).Value = sName Then
            SearchCell = GetOffset(vCell, Range(�G���g���[�e�[�u�� & "[" & sColName & "]").Column).Address
            Exit Function
        End If
    Next vCell

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


