Attribute VB_Name = "bossSkillClearMod"
Public Function bossSkillDataClear()

    Range("C37:AP38, C81:AP82, C125:AP126").Interior.ColorIndex = xlNone
        
    Range("C37:AP38, C81:AP82, C125:AP126").Value = ""

End Function

