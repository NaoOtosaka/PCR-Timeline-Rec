Attribute VB_Name = "便捷填充_备份"
Sub 按钮5_Click()
    If IsEmpty(Range("AL19")) Then
    
        MsgBox "非buff技能！"
        
    Else
        '技能数据读取
        skillName = Range("X19").Value
        buffTime = Range("AL19").Value
        buffNum = Range("AM19").Value
        temp = 0
    
        '技能效果判定查询
            Rem 定位
        skillR = Sheets("技能").Range("E:E").Find(What:=skillName).Row
        skillC = Sheets("技能").Range("E:E").Find(What:=skillName).Column
        
        skillTag = Sheets("技能").Cells(skillR, skillC + 1)
        
        Debug.Print skillTag
     
            Rem 记忆
        If skillTag = "" Then
            If MsgBox("该技能是否为破甲技能？（后续使用时会记忆本次选择选项）", vbYesNo, "请选择") = vbYes Then
                Sheets("技能").Cells(skillR, skillC + 1).Value = 1
                skillTag = 1
            Else
                Sheets("技能").Cells(skillR, skillC + 1).Value = 0
                skillTag = 0
            End If
        End If
        
        '颜色设定
        If skillTag = 1 Then
            buffColor = 37
        Else
            buffColor = 39
        End If
        
        '时长获取
        startTime = InputBox("请输入开始时间")
        
        If startTime = "" Then
            MsgBox "未输入开始时间"
            End
        End If
        
        '时间轴坐标初始化
        locationR = Range("C36:AP36, C68:AP68, C100:M100").Find(What:=startTime).Row
        locationC = Range("C36:AP36, C68:AP68, C100:M100").Find(What:=startTime).Column
        
        Debug.Print locationR
        Debug.Print locationC
        
        '循环填充
        For i = 0 To buffTime - 1 Step 1
            '换行
            If (locationC + i) > 42 Then
                locationR = locationR + 32
                locationC = 3
                temp = i
            End If
            
            '填充
            Cells(locationR + 9, locationC + i - temp).Interior.ColorIndex = buffColor
            Cells(locationR + 9, locationC + i - temp) = buffNum
        Next i
    
    End If
    
End Sub
