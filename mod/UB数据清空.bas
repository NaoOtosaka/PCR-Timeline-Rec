Attribute VB_Name = "UBÊý¾ÝÇå¿Õ"
Public Function ubDataClear()

    Range("C45:AP45, C53:AP53, C61:AP61, C69:AP69, C77:AP77").Interior.ColorIndex = xlNone
        
    Range("C45:AP45, C53:AP53, C61:AP61, C69:AP69, C77:AP77").Value = ""
    
    Range("C89:AP89, C97:AP97, C105:AP105, C113:AP113, C121:AP121").Interior.ColorIndex = xlNone
        
    Range("C89:AP89, C97:AP97, C105:AP105, C113:AP113, C121:AP121").Value = ""
    
    Range("C133:M133, C141:M141, C149:M149, C157:M157, C165:M165").Interior.ColorIndex = xlNone
        
    Range("C133:M133, C141:M141, C149:M149, C157:M157, C165:M165").Value = ""

End Function
