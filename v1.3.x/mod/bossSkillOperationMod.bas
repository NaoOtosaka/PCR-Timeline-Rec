Attribute VB_Name = "bossSkillOperationMod"
Sub 按钮31_Click()
    If MsgBox("确定清空吗？（此操作不可逆）", vbYesNo, "请选择") = vbYes Then
        Range("B7:D17").Value = ""
    End If
End Sub
Sub BOSS信息_按钮34_Click()
    If MsgBox("确定清空吗？（此操作不可逆）", vbYesNo, "请选择") = vbYes Then
        Range("B7:D17").Value = ""
        Range("B2:B4").Value = ""
    End If
End Sub
