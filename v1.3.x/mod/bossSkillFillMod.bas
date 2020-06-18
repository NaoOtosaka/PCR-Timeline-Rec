Attribute VB_Name = "bossSkillFillMod"
Public Function bossQuickFill(in_skillType, in_skillTime)
    
    timeStyle = Sheets("_Sheet1").Range("T14").Value
    
    '完整性校验
    If IsEmpty(Range(in_skillType)) Then
        MsgBox "请补充动作信息"
        End
    End If
    
    
    '技能数据读取
    skillType = Range(in_skillType).Value
    temp = 0
    
    
    '技能类型判定
    If IsEmpty(Range(in_skillTime)) Or Range(in_skillTime).Value = 1 Then
        buffTime = -1
    Else
        buffTime = Range(in_skillTime).Value
    End If
    
    
    '颜色设定
    If buffTime = -1 Then
        buffColor = 46
        startRow = 1
    Else
        buffColor = 42
        startRow = 2
    End If
    
    
    '时长获取
    If buffTime > 1 Then
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
        locationR = Sheets("本体").Range("C36:AP36").Find(What:=startTime).Row
        locationC = Sheets("本体").Range("C36:AP36").Find(What:=startTime).Column
    ElseIf startTime >= 11 Then
        locationR = Sheets("本体").Range("C80:AP80").Find(What:=startTime).Row
        locationC = Sheets("本体").Range("C80:AP80").Find(What:=startTime).Column
    Else
        locationR = Sheets("本体").Range("C124:M124").Find(What:=startTime).Row
        locationC = Sheets("本体").Range("C124:M124").Find(What:=startTime).Column
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
        If (locationC + i - temp) > 42 Then
            locationR = locationR + 44
            locationC = 3
            temp = i
        End If
        
        '填充
        Sheets("本体").Cells(locationR + startRow, locationC + i - temp).Interior.ColorIndex = buffColor
        
        If i = 0 Then
            Sheets("本体").Cells(locationR + startRow, locationC + i - temp) = Left(skillType, 2)
        Else
            Sheets("本体").Cells(locationR + startRow, locationC + i - temp) = ""
        End If
        
    Next i
        
End Function

