Attribute VB_Name = "createWorkMod"
Public Function arrNum(arr) As Boolean

    Dim i&
    
    On Error Resume Next
    
    i = UBound(arr)
    
    If Err = 0 Then arrNum = True
    
End Function
Sub 出轴区_按钮1_Click()
    Dim file As String, arr As Variant, i
    
    '将当前的数据读入数组
    arr = Range("a1").CurrentRegion
    
    If arrNum(arr) Then
    
        d = Format(Now(), "yyyy-mm-dd_HH.mm.ss")
        
        s = InputBox("请输入保存文件名")
        
        '定义文本文件的名称
        file = ThisWorkbook.Path & "\txt\" & s & "_" & d & ".txt"
        
        '判断是否存在同名文本文件，存在先行删除
        If Dir(file) <> "" Then Kill file
        
        '使用print语句将数组中所有数据写入文本文件
        Open file For Output As #1
        
        Print #1, "=========================================================="
        Print #1, "=                                          该作业生成于：";
        Print #1, Format(Now(), "yyyy-mm-dd_HH.mm.ss");
        Print #1, "                                                ="
        Print #1, "=========================================================="
        
        Print #1, "              BOSS名称：";
        Print #1, Sheets("BOSS信息").Range("B2").Value
        
        Print #1, "              BOSS位置：";
        Print #1, Sheets("BOSS信息").Range("B3").Value
        
        Print #1, "              备注：";
        Print #1, Sheets("BOSS信息").Range("B4").Value
        
        Print #1, "=========================================================="
        
        For i = 1 To UBound(arr)
        
            Print #1, Join(Application.Index(arr, i), " - ")
            
        Next
        
        '关闭文本文件
        Close #1
        
    End If
End Sub
Sub 出轴区_按钮2_Click()

    Sheets("出轴区").Range("A1:G65536").ClearContents
    
End Sub
Sub 出轴区_按钮3_Click()
Dim dic As Object
    
    Dim ubTimeArr As Variant

    Dim min As String, sec As String

    Set dic = CreateObject("Scripting.Dictionary")
    
    For i = 0 To 4
    
        For v = 0 To 33
    
            If Sheets("本体").Cells(i + 11, v + 5).Value = "" Then
            
                Exit For
            
            End If
    
            If Not dic.exists(Sheets("本体").Cells(i + 11, v + 5).Value) Then
            
                dic.Add Sheets("本体").Cells(i + 11, v + 5).Value, CreateObject("Scripting.Dictionary")
            
            End If
            
            Index = dic(Sheets("本体").Cells(i + 11, v + 5).Value).Count
            
            Debug.Print Sheets("本体").Cells(i + 11, 1).Value
            
            dic(Sheets("本体").Cells(i + 11, v + 5).Value).Add Index, Sheets("本体").Cells(i + 11, 1).Value
            
        Next v
    
    Next i
    
    dKeys = dic.Keys
    
    ubTimeArr = bubbleSort(dKeys)
    
    i = 1
    
    For Each ubTime In ubTimeArr
        
        ubTimeTemp = 0
        
        min = 0
        
        sec = 0
        
        If ubTime >= 60 Then
            
            min = "1"
            
            ubTimeTemp = ubTime - 60
            
            If ubTimeTemp < 10 Then
            
                sec = " 0" & Right(Str(ubTimeTemp), 1)
            
            Else
            
                sec = Str(ubTimeTemp)
            
            End If
            
        Else
        
            sec = Str(ubTime)
        
        End If
        

    
        Sheets("出轴区").Cells(i, 1) = min & ":" & sec
        
        v = 2
        
        For Each ubName In dic(ubTime).Items
        
            Sheets("出轴区").Cells(i, v) = "[ub]" & ubName
            
            v = v + 1
            
        Next
        
        i = i + 1
    
    Next
End Sub


