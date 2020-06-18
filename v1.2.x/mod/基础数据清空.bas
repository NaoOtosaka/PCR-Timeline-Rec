Attribute VB_Name = "基础数据清空"
Sub 按钮1_Click()
    If MsgBox("确定删除吗？（此操作不可逆）", vbYesNo, "请选择") = vbYes Then
        Range("C39:AP39, C41:AP41, C47:AP47, C49:AP49, C55:AP55, C57:AP57, C63:AP63, C65:AP65, C71:AP71, C73:AP73").Interior.ColorIndex = xlNone
        
        Range("C39:AP39, C41:AP41, C47:AP47, C49:AP49, C55:AP55, C57:AP57, C63:AP63, C65:AP65, C71:AP71, C73:AP73").Value = ""

        Range("C83:AP83, C85:AP85, C91:AP91, C93:AP93, C99:AP99, C101:AP101, C107:AP107, C109:AP109, C115:AP115, C117:AP117").Interior.ColorIndex = xlNone
        
        Range("C83:AP83, C85:AP85, C91:AP91, C93:AP93, C99:AP99, C101:AP101, C107:AP107, C109:AP109, C115:AP115, C117:AP117").Value = ""
        
        Range("C127:M127, C129:M129, C135:M135, C137:M137, C143:M143, C145:M145, C151:M151, C153:M153, C159:M159, C161:M161").Interior.ColorIndex = xlNone
        
        Range("C127:M127, C129:M129, C135:M135, C137:M137, C143:M143, C145:M145, C151:M151, C153:M153, C159:M159, C161:M161").Value = ""

        Call ubDataClear
    End If
End Sub
