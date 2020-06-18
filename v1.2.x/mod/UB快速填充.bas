Attribute VB_Name = "UB快速填充"
Public Function ubQuickFill(in_skillName, in_skillTime, in_timeArr, in_startRow)
    
    timeStyle = Sheets("_Sheet1").Range("T14").Value
    
    'ub数据读取
    skillName = Range(in_skillName).Value
    
    'ub类型判定
    If Range(in_skillTime) = "" Then
        Exit Function
    Else
        buffTime = Int(Range(in_skillTime).Value)
    End If

    '技能效果判定查询
        Rem 定位
    skillR = Sheets("UB").Range("B:B").Find(What:=skillName).Row
    skillC = Sheets("UB").Range("B:B").Find(What:=skillName).Column
    
    skillTag = Sheets("UB").Cells(skillR, skillC + 2)
    
    '破甲判定
    If buffTime > 0 Then
            Rem 记忆
        If skillTag = "" Then
            If MsgBox("【" & skillName & "】是否为破甲技能？（后续使用时会记忆本次选择选项）", vbYesNo, "请选择") = vbYes Then
                Sheets("UB").Cells(skillR, skillC + 2).Value = 1
                skillTag = 1
            Else
                Sheets("UB").Cells(skillR, skillC + 2).Value = 0
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
    
    Dim arr
    
    arr = in_timeArr
    
    For Each r In arr
        If r = "" Then
            Exit For
        End If
        
        temp = 0
        
        startTime = r
        
        '风格判定
        If timeStyle Then
            If startTime > 90 Then
                startTime = startTime - 40
            End If
        Else
            If startTime > 60 And startTime < 100 Then
                startTime = startTime + 40
            End If
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
        
    Next r

End Function


