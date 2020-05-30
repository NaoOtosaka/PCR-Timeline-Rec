Attribute VB_Name = "填充测试"
Sub Sheet1_按钮1_Click()
    Dim time As Integer
    buffColor = 37
    
    buffTime = Range("B1").Value
    buffNum = Range("C1").Value
    
    Range("B4").Value = time
    
    startTime = InputBox("请输入开始时间")
    
    locationR = Range("K4:X4").Find(What:=startTime).Row
    locationC = Range("K4:X4").Find(What:=startTime).Column
    
    Range("K2").Value = locationR
    Range("K3").Value = locationC
    
    For i = 0 To buffTime - 1 Step 1
        Cells(locationR + 1, locationC + i).Interior.ColorIndex = buffColor
        Cells(locationR + 1, locationC + i) = buffNum
    Next i
    
End Sub
