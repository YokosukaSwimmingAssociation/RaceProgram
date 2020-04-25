Attribute VB_Name = "CommonModule"
' ��
Public Const �w����� As String = "���{��w�����j���Z���"
Public Const �}�X�^�[�Y��� As String = "���{��}�X�^�[�Y���"
Public Const �w�}��� As String = "�w���}�X�^�[�Y���"
Public Const �s����� As String = "���{��s���̈���"
Public Const �I�茠��� As String = "���{��I�茠���j���"

Public Const S_ENTRY_SHEET_NAME As String = "�G���g���[�ꗗ"
Public Const S_ENTRY_TABLE_NAME As String = "�G���g���[�e�[�u��"
Public Const S_PROGRAM_SHEE_TNAME As String = "�v���O����"
Public Const S_PROGRAM_FORMAT_SHEET_NAME As String = "�v���O�����t�H�[�}�b�g"

Public Const N_NUMBER_OF_RACE As Integer = 7       ' �P���[�X�̐l��
Public Const N_MIN_NUMBER_OF_RACE As Integer = 3   ' ���[���̍ŏ��l��
Public Const N_MIN_NUMBER_OF_RACE2 As Integer = 4   ' ���[���̍ŏ��l��
Public Const N_MAX_LANE_OF_RACE As Integer = 9     ' ���[���̍ő�ԍ�
Public Const N_MIN_LANE_OF_RACE As Integer = 3     ' ���[���̍ŏ��ԍ�
Public Const N_AVERAGE_DEC_RACE As Integer = 3      ' ���ϕ��������ɂ���g��

Public Const S_BLANK_NAME As String = "�@�@�D�@�@�D�@�@�D"
Private Const ARRAYSIZE = 10000

Public Type RaceNumber
    nProNo As Integer
    nRance As Integer
End Type

'
' �󔒒u��
'
' �S�p�󔒁A�A�������󔒂𔼊p�󔒂P�ɕϊ�
'
' sStr          IN      ������
'
Public Function STrim(sStr)
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
Public Function STrimAll(sStr)
    STrimAll = Replace(Replace(sStr, "�@", ""), " ", "")
End Function


'
' �V�[�g�̑��݃`�F�b�N�t���A�N�e�B�x�[�g
'
' sSheetName    IN      �V�[�g��
'
Sub SheetActivate(sSheetName As String)
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
Public Function IsSheetExists(sSheetName As String)
    IsSheetExists = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = sSheetName Then
            IsSheetExists = True
        End If
    Next ws
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
Public Function IsNameExists(sName As String)
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
' �e�[�u���̑��݊m�F
'
' oWorkSheet        IN      ���[�N�V�[�g
' sTableName        IN      ���O
'
Public Function IsTableExists(oWorkSheet As Worksheet, sTableName As String)
    IsTableExists = False
    For Each lst In oWorkSheet.ListObjects
        If lst.Name = sTableName Then
            IsTableExists = True
            Exit For
        End If
    Next
End Function

'
' �e�[�u���̒�`
'
' oWorkSheet        IN      ���[�N�V�[�g
' sTableName        IN      ���O
' sRange            IN      �����W�͈�
'
Public Sub SetTable(oWorkSheet As Worksheet, sTableName As String, sRange As String)
    If IsTableExists(oWorkSheet, sTableName) Then
        oWorkSheet.ListObjects(sTableName).Unlist
    End If
    oWorkSheet.ListObjects.Add(xlSrcRange, Range(sRange), , xlYes).Name = sTableName
End Sub

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
' �C�x���g�̔����A��ʕ`���On/Off
'
' bFlag     IN      True�F�ĊJ�^False�F�}��
'
Sub EventChange(bFlag As Boolean)
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
' �^�C�g���o�[�̕ύX
'
' sTitle    IN      �^�C�g���o�[������
'
Sub SetTitleMenu(sTitle As String)
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
' ����s�͈͎̔擾
'
' �w�肵���Z������ŉE�͈̔͂̃A�h���X��Ԃ�
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Function ColumnRange(sTopAddress As String)

    Set ColumnRange = Range(Range(sTopAddress), _
                    Range(sTopAddress).End(xlToRight))

End Function

'
' ����s�͈͎̔擾
'
' �w�肵���Z������ŉE�͈̔͂̃A�h���X��Ԃ�
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Function ColumnRangeAddress(sTopAddress As String)

    ColumnRangeAddress = ColumnRange(sTopAddress).Address

End Function

'
' �����͈͎̔擾
'
' �w�肵���Z������ŉ��w�͈̔͂̃A�h���X��Ԃ�
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Function RowRange(sTopAddress As String)
    Set RowRange = Range(Range(sTopAddress), _
                    Range(sTopAddress).End(xlDown))
End Function

'
' �����͈͎̔擾
'
' �w�肵���Z������ŉ��w�͈̔͂̃A�h���X��Ԃ�
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Function RowRangeAddress(sTopAddress As String)

    RowRangeAddress = RowRange(sTopAddress).Address

End Function

'
' �s��͈͎̔擾
'
' �w�肵���Z������ŉ��w�A�ۉE�[�͈͂̃A�h���X��Ԃ�
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Function TableRange(sTopAddress As String)

    Dim oRng As Range
    Set oRng = Range(Range(sTopAddress), Range(sTopAddress).End(xlDown))
    Set TableRange = Range(oRng, oRng.End(xlToRight))

End Function

'
' �s��͈͎̔擾
'
' �w�肵���Z������ŉ��w�A�ۉE�[�͈͂̃A�h���X��Ԃ�
'
' sTopAddres IN      �擪�̃Z���A�h���X
'
Function TableRangeAddress(sTopAddress As String)

    TableRangeAddress = TableRange(sTopAddress).Address

End Function

'
' �������ڋ敪��Ԃ�
'
Function GetMaster(sGameName As String)
    ' ���{��I�茠���j���
    If sGameName = "���{��I�茠���j���" Then
        GetMaster = "�I�茠��ڋ敪"
    ' ���{��s���̈���
    ElseIf sGameName = "���{��s���̈���" Then
        GetMaster = "�s����ڋ敪"
    Else
        GetMaster = "�w�}��ڋ敪"
    End If

End Function

'
' �v���O�����␳�p�F�Z���̏C��
'
' sRange     IN      �Z���̃A�h���X
' sValue     IN      �ύX����l
'
Sub ModCell(sRange As String, sValue As String)

    Range(sRange).Select
    Range(sRange).Value = sValue
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

'
' �v���O�����␳�p�F�v���O�����ԍ��̃A�h���X��T��
'
' nProNo     IN      �v���O�����ԍ�
' sName      IN      �I�薼
' sColName   IN      �J�������i���[�XNo�A�g�A���[���j
'            OUT     �A�h���X������
'
Function SearchCell(nProNo As Integer, sName As String, sColName As String)

    For Each oCell In Range(Range("$A$2"), Range("$A$2").End(xlDown))
        If Cells(oCell.Row, Range(S_ENTRY_TABLE_NAME & "[�v��No]").Column).Value = nProNo And _
            Cells(oCell.Row, Range(S_ENTRY_TABLE_NAME & "[�I�薼]").Column).Value = sName Then
            SearchCell = Cells(oCell.Row, Range(S_ENTRY_TABLE_NAME & "[" & sColName & "]").Column).Address
            Exit Function
        End If
    Next oCell

End Function

'
' �v���O�������͗p�F�L�^��ʂ̎�ڑI��
'
' nProNo     IN      �v���O�����ԍ�
' nHeat      IN      �g
'
Sub SetRace(nProNo As Integer, nHeat As Integer)
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
Sub SetLean(nIndex As Integer, nLean As Integer, sTime As String, Optional sAdditional As String = "")
    GetRange("�L�^��ʃ��[��").Rows(nIndex).Value = nLean
    GetRange("�L�^��ʃ^�C��").Rows(nIndex).Value = sTime
    If sAdditional <> "" Then
        GetRange("�L�^��ʔ��l").Rows(nIndex).Value = sAdditional
    End If
End Sub


'
' �͈͂̍ō���ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Function GetAreaLeftColumn(sName As String)
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaLeftColumn = oRange.Column
End Function

'
' �͈͂̍ŉE��ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Function GetAreaRightColumn(sName As String)
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaRightColumn = oRange.Column + oRange.Columns.Count - 1
End Function

'
' �͈͂̍ŏ�s�ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Function GetAreaTopRow(sName As String)
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaTopRow = oRange.Row
End Function

'
' �͈͂̍ŉ��s�ԍ���Ԃ�
'
' sName      IN      �͈͖�
'
Function GetAreaBottomRow(sName As String)
    Dim oRange As Range
    Set oRange = Range(sName)
    GetAreaBottomRow = oRange.Row + oRange.Rows.Count - 1
End Function


'
' ��Z������I�t�Z�b�g�ʒu�̒l�Z����Ԃ�
'
' oCell         IN      ��Z��
' nColumn       IN      ��ԍ�
'
Function GetOffset(oCell As Variant, nColumn As Integer)
    Set GetOffset = oCell.Offset(0, nColumn - oCell.Column)
End Function

'
' �͈͂̃w�b�_�s����w�肵����ԍ���Ԃ�
'
' sName      IN      �͈͖�
' sColName   IN      �J������
'            OUT     ��̔ԍ�
'
Function GetAreaColumnIndex(sName As String, sColName As String)
    Dim nIndex As Integer
    nIndex = 1
    For Each oCell In GetRange(sName).Rows(1).Columns
        If Replace(Replace(oCell.Value, "�@", ""), " ", "") = sColName Then
            GetAreaColumnIndex = nIndex
            Exit Function
        End If
        nIndex = nIndex + 1
    Next
End Function

'
' �͈͂̃w�b�_�s�������L�[����擾����
'
' sName      IN      �͈͖�
'
Function GetAreaKeyData(sName As String)
    Dim oRange As Range
    Set oRange = GetRange(sName)
    Set GetAreaKeyData = oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, 1).Rows()
End Function

'
' �͈͂̃L�[�񖼂�Ԃ�
'
' sName      IN      �͈͖�
'
Function GetAreaKeyName(sName As String)
    Dim oRange As Range
    Set oRange = GetRange(sName)
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
Function VLookupArea(vValue As Variant, sName As String, sColName As String, Optional bFlag As Boolean = False)
    VLookupArea = Application.WorksheetFunction.VLookup(vValue, GetRange(sName), GetAreaColumnIndex(sName, sColName), bFlag)
End Function

'
' ���W���[���̓Ǎ���
'
' oWorkBook  IN      WorkBook
' sPath      IN      �t�H���_�p�X
'
Sub ImportAll(oWorkBook As Workbook, sPath As String)
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
Sub ExportAll(oWorkBook As Workbook, sPath As String)
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
Function SelectDir()
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
Function GetFiles(sPathName As String, sExt As String)
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
' �V�[�g�̕ی�
'
' bFlag         IN      True�F�ی�^False�F����
'
Sub SheetProtect(bFlag As Boolean)

    If bFlag Then
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, _
            Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Else
        ActiveSheet.Unprotect
    End If

End Sub
