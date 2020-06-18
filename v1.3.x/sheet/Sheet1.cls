VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Range("A4:A8")) Is Nothing And ActiveCell = "" Then
        If Target.Count = 1 Then
            Query.Show
        End If
    End If
End Sub

