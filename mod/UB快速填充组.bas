Attribute VB_Name = "UB快速填充组"
Sub 按钮34_Click()

    Call ubDataClear

    Call ubQuickFill("C25", "I25", Range("E11:AP11"), 9)

    Call ubQuickFill("K25", "Q25", Range("E12:AP12"), 17)
    
    Call ubQuickFill("S25", "Y25", Range("E13:AP13"), 25)
        
    Call ubQuickFill("AA25", "AG25", Range("E14:AP14"), 33)
    
    Call ubQuickFill("AI25", "AO25", Range("E15:AP15"), 41)
End Sub
