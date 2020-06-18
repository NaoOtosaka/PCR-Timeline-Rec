Attribute VB_Name = "uBOverloadedGroupMod"
Public Function ubOverLoad(in_timeLinePos, in_skillName, in_skillTime, in_timeArr, in_startRow)
        'ub类型判定
    If Range(in_skillTime) = "" Then
        
        End
        
    Else
    
        in_timeLinePos.Interior.ColorIndex = xlNone
        
        in_timeLinePos.Value = ""
    
        Call uBQuickFill(in_skillName, in_skillTime, in_timeArr, in_startRow)
    
    End If
    
End Function

Sub 按钮37_Click()
    
    Call ubOverLoad(Range("C45:AP45, C89:AP89, C133:M133"), "C25", "I25", Range("E11:AP11"), 9)
        
End Sub

Sub 按钮38_Click()

    Call ubOverLoad(Range("C53:AP53, C97:AP97, C141:M141"), "K25", "Q25", Range("E12:AP12"), 17)

End Sub

Sub 按钮39_Click()

    Call ubOverLoad(Range("C61:AP61, C105:AP105, C149:M149"), "S25", "Y25", Range("E13:AP13"), 25)

End Sub

Sub 按钮40_Click()

    Call ubOverLoad(Range("C69:AP69, C113:AP113, C157:M157"), "AA25", "AG25", Range("E14:AP14"), 33)

End Sub

Sub 按钮41_Click()

    Call ubOverLoad(Range("C77:AP77, C121:AP121, C165:M165"), "AI25", "AO25", Range("E15:AP15"), 41)

End Sub

