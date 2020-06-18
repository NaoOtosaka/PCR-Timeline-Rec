Attribute VB_Name = "basicDataOperationMod"
Sub 按钮7_Click()
    Range("E11:AP15").Value = Range("E4:AP8").Value
End Sub
Sub 按钮8_Click()
    If MsgBox("确定清空吗？（此操作不可逆）", vbYesNo, "请选择") = vbYes Then
        Range("A4:AP8, E11:AP15").Value = ""
        Range("D18:S22, X18:AM22").Value = ""
    End If
End Sub
Sub 按钮10_Click()
    Dim rg As Range
    Dim locationR As Integer
    Dim locationC As Integer
    
    
    For Each rg In Sheets("本体").Range("D18:D22,X18:X22")
        Debug.Print rg.Value
        
        If rg.Value = "" Then
            GoTo con
        End If
        
        locationR = Range("技能介绍_可用[skillName]").Find(What:=rg.Value).Row
        locationC = Range("技能介绍_可用[skillName]").Find(What:=rg.Value).Column
        
        Debug.Print locationR
        
        Debug.Print locationC
        
        data = Sheets("处理").Cells(locationR, locationC + 1).Value
        
        buffTime = Sheets("处理").Cells(locationR, locationC + 2).Value
        
        Debug.Print data
        Debug.Print buffTime
        
        r = rg.Row
        c = rg.Column
        
        Debug.Print r
        Debug.Print c
        
        Sheets("本体").Cells(r, c + 3).Value = data
        Sheets("本体").Cells(r, c + 14).Value = buffTime
        
con:
    Next
    
End Sub
Sub 按钮29_Click()
    If MsgBox("本操作将备份工作薄至当前文件所在目录下", vbYesNo, "请选择") = vbYes Then
        ActiveWorkbook.Save
        d = Format(Now(), "yyyy-mm-dd_HH.mm.ss")
        s = InputBox("请输入保存文件名")
        ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\" & s & "_" & d & ".xlsm"
    End If
End Sub
Sub 按钮30_Click()
    If MsgBox("本操作将备份表至当前工作簿", vbYesNo, "请选择") = vbYes Then
        ActiveSheet.Select
        ActiveSheet.Copy after:=Sheets("更新记录")
        
        If MsgBox("是否需要重命名备份表", vbYesNo, "请选择") = vbYes Then
            newName = InputBox("请输入新名称")
            ActiveSheet.Name = "_" & newName
        End If
    End If
End Sub

