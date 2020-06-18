Attribute VB_Name = "exportVBAMod"
Sub Sheet1_°´Å¥5_Click()
    Dim code
    For Each code In ThisWorkbook.VBProject.VBComponents
    code.Export ThisWorkbook.Path & "\code" & "\" & code.Name & "." & Split("cls bas cls frm")(code.Type Mod 4)
    Next
End Sub
