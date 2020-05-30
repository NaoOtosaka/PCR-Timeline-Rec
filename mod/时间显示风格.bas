Attribute VB_Name = "时间显示风格"
Public timeStyle As Boolean 'true为秒数 false为直读
Sub 按钮43_Click()

    timeLine = Range("C36:AG36")
    
    If timeStyle Then
        Range("H35").Value = "直读模式"
        Rem 标签设定
        Tag = 130
        Rem 循环
        For i = 3 To 33
            Sheets("本体").Cells(36, i).Value = Tag
            Tag = Tag - 1
        Next i

        timeStyle = False
        
    Else
        Range("H35").Value = "秒数模式"
        Rem 标签设定
        Tag = 90
        Rem 循环
        For i = 3 To 33
            Sheets("本体").Cells(36, i).Value = Tag
            Tag = Tag - 1
        Next i
        
        timeStyle = True
    End If
End Sub

