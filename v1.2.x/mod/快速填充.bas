Attribute VB_Name = "快速填充"
Public Function quickFill(in_skillName, in_skillTime, in_startRow)
    
    timeStyle = Sheets("_Sheet1").Range("T14").Value
    
    '完整性校验
    If IsEmpty(Range(in_skillName)) Then
        MsgBox "请选择技能！"
        End
    End If
    
    
    '技能数据读取
    skillName = Range(in_skillName).Value
    temp = 0
    
    '技能类型判定
    If IsEmpty(Range(in_skillTime)) Then
        MsgBox "非buff技能！"
        buffTime = -1
    Else
    
        buffTime = Range(in_skillTime).Value
    
    End If


    '技能效果判定查询
        Rem 定位
    skillR = Sheets("技能").Range("E:E").Find(What:=skillName).Row
    skillC = Sheets("技能").Range("E:E").Find(What:=skillName).Column
    
    skillTag = Sheets("技能").Cells(skillR, skillC + 1)
    
    
    Rem Debug.Print skillTag
    
    
    '破甲判定
    If buffTime > 0 Then
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
    Else
        skillTag = 0
    End If
    
    
    '颜色设定
    If skillTag = 1 Then
        buffColor = 37
    Else
        buffColor = 39
    End If
    
    
    '时长获取
    If buffTime > 0 Then
        startTime = InputBox("请输入开始时间")
    Else
        startTime = InputBox("请输入开始时间(该技能为非buff技能)")
        '回归赋值
        buffTime = 1
    End If
    
    '风格判定
    If timeStyle Then
        If startTime > 90 Then
            MsgBox "输入不符合当前时间模式"
            End
        End If
    Else
        If startTime > 60 And startTime < 100 Then
            MsgBox "输入不符合当前时间模式"
            End
        End If
    End If
        
    '输入值处理
    If startTime = "" Then
        MsgBox "未输入开始时间"
        End
    Else
        startTime = Int(startTime)
    End If
    
    
    '时间轴坐标初始化
    If startTime >= 51 Then
        locationR = Range("C36:AP36").Find(What:=startTime).Row
        locationC = Range("C36:AP36").Find(What:=startTime).Column
    ElseIf startTime >= 11 Then
        locationR = Range("C80:AP80").Find(What:=startTime).Row
        locationC = Range("C80:AP80").Find(What:=startTime).Column
    Else
        locationR = Range("C124:M124").Find(What:=startTime).Row
        locationC = Range("C124:M124").Find(What:=startTime).Column
    End If
    
    
    Rem Debug.Print locationR
    Rem Debug.Print locationC
    
    
    '结尾时间判定
    If startTime < buffTime Then
        buffTime = startTime + 1
    End If
    
    
    '循环填充
    For i = 0 To buffTime - 1 Step 1
        '换行
        If (locationC + i) > 42 Then
            locationR = locationR + 44
            locationC = 3
            temp = i
        End If
        
        '填充
        Cells(locationR + in_startRow, locationC + i - temp).Interior.ColorIndex = buffColor
        
        If i = 0 Then
            Cells(locationR + in_startRow, locationC + i - temp) = Left(skillName, 2)
        Else
            Cells(locationR + in_startRow, locationC + i - temp) = ""
        End If
        
    Next i
        
End Function
