Attribute VB_Name = "uBQuickFillGroupMod"
Sub °´Å¥34_Click()

    Call uBDataClear

    Call uBQuickFill("C25", "I25", Range("E11:AP11"), 9)

    Call uBQuickFill("K25", "Q25", Range("E12:AP12"), 17)
    
    Call uBQuickFill("S25", "Y25", Range("E13:AP13"), 25)
        
    Call uBQuickFill("AA25", "AG25", Range("E14:AP14"), 33)
    
    Call uBQuickFill("AI25", "AO25", Range("E15:AP15"), 41)
End Sub
