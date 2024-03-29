VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not (Intersect(Target, Range("大会名")) Is Nothing) Then
        If Range("大会名").Value = 選手権大会 Then
            Range("大会組最少人数").Value = 3
            Range("大会組レース定員").Value = 7
            Range("大会組最小レーン番号").Value = 3
            Range("大会組合せ方式").Value = "混合分け方式"
        ElseIf Range("大会名").Value = 室内記録会 Then
            Range("大会組最少人数").Value = 3
            Range("大会組レース定員").Value = 5
            Range("大会組最小レーン番号").Value = 1
            Range("大会組合せ方式").Value = "単純方式"
        Else
            Range("大会組最少人数").Value = 4
            Range("大会組レース定員").Value = 7
            Range("大会組最小レーン番号").Value = 3
            Range("大会組合せ方式").Value = "単純方式"
        End If
        Call シート非表示
    Else
        Exit Sub
    End If
End Sub

