VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not (Intersect(Target, Range("記録画面種目番号")) Is Nothing) Then
        Call 種目名読込み
        Call レース番号読込み
    ElseIf Not (Intersect(Target, Range("記録画面組")) Is Nothing) Then
        Call レース番号読込み
    ElseIf Not (Intersect(Target, Range("記録画面レーン")) Is Nothing) Then
        Call 選手名読込み(Target)
    ElseIf Not (Intersect(Target, Range("記録画面タイム")) Is Nothing) Then
        Call 大会記録判定(Target)
    ElseIf Not (Intersect(Target, Range("記録画面違反")) Is Nothing) Then
        Call 違反反映(Target)
    Else
        Exit Sub
    End If
End Sub
