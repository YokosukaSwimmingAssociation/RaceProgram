VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not (Intersect(Target, Range("��")) Is Nothing) Then
        If Range("��").Value = "���{��I�茠���j���" Then
            Range("�g�ŏ��l��").Value = 3
            Range("�g��������").Value = "������������"
        Else
            Range("�g�ŏ��l��").Value = 4
            Range("�g��������").Value = "�P������"
        End If
        Call �V�[�g��\��
    Else
        Exit Sub
    End If
End Sub
