Attribute VB_Name = "createWorkMod"
Public Function arrNum(arr) As Boolean

    Dim i&
    
    On Error Resume Next
    
    i = UBound(arr)
    
    If Err = 0 Then arrNum = True
    
End Function
Sub 出轴区_按钮1_Click()
    Application.DisplayAlerts = False
    
    Dim file As String, arr As Variant, i
    
    '将当前的数据读入数组
    arr = Range("a1").CurrentRegion
    
    If arrNum(arr) Then
    
        d = Format(Now(), "yyyy-mm-dd_HH.mm.ss")
        
        s = InputBox("请输入保存文件名")
        
        '新建一个对话框对象
        Set FolderDialogObject = Application.FileDialog(msoFileDialogFolderPicker)
        
        '配置对话框
        With FolderDialogObject
        
            .Title = "请选择要查找的文件夹"
        
            .InitialFileName = ThisWorkbook.Path
        
        End With
        
        '显示对话框
        
        FolderDialogObject.Show
        
        '获取选择对话框选择的文件夹
        
        Set paths = FolderDialogObject.SelectedItems
        
        '错误抑制
        On Error GoTo pathErr
        
        '定义文本文件的名称
        file = paths(1) & "\" & s & "_" & d & ".txt"
        
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
    
    Application.DisplayAlerts = True
    Exit Sub
pathErr:
    MsgBox "未选择路径，本次导出已取消"
    Exit Sub
End Sub
Sub 出轴区_按钮2_Click()

    Sheets("出轴区").Range("A1:G65536").ClearContents
    
End Sub
Sub 出轴区_按钮3_Click()
Dim dic As Object

    Sheets("出轴区").Range("A1:G65536").ClearContents

    timeStyle = Sheets("_Sheet1").Range("T14").Value
    
    Dim ubTimeArr As Variant

    Dim min As String, sec As String

    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim workName As String
    
    workName = Sheets("出轴区").Range("I3").Value
    
    If workName = "" Then
        
        MsgBox "请选择工作表！"
        Exit Sub
        
    End If
    
    For i = 0 To 4
    
        For v = 0 To 33
    
            If Sheets(workName).Cells(i + 11, v + 5).Value = "" Then
            
                Exit For
            
            End If
    
            If Not dic.exists(Sheets(workName).Cells(i + 11, v + 5).Value) Then
            
                dic.Add Sheets(workName).Cells(i + 11, v + 5).Value, CreateObject("Scripting.Dictionary")
            
            End If
            
            Index = dic(Sheets(workName).Cells(i + 11, v + 5).Value).Count
            
            Debug.Print Sheets(workName).Cells(i + 11, 1).Value
            
            dic(Sheets(workName).Cells(i + 11, v + 5).Value).Add Index, Sheets(workName).Cells(i + 11, 1).Value
            
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
            
            If timeStyle Then
                ubTimeTemp = ubTime - 60
            Else
                ubTimeTemp = ubTime - 100
            End If
                    
            If ubTimeTemp < 10 Then
            
                sec = " 0" & Right(str(ubTimeTemp), 1)
            
            Else
            
                sec = str(ubTimeTemp)
            
            End If
            
        ElseIf ubTime >= 10 Then
        
            sec = str(ubTime)
        
        Else
        
            sec = " 0" & Right(str(ubTime), 1)
        
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
Sub 按钮6_Click()
    
    Dim ar(1 To 100, 1 To 1)
    
    Dim i As Long, j As Long
    
    Sheets("表").UsedRange.Offset(1).ClearContents
    
    For i = 1 To Sheets.Count
        
        If Sheets(i).Visible = xlSheetVisible Then
            
            If Sheets(i).Name <> "BOSS信息" And Sheets(i).Name <> "出轴区" And Sheets(i).Name <> "更新记录" Then
                
                j = j + 1
                
                Sheets("表").Cells(j, 1) = Sheets(i).Name
                
            End If
            
        End If
        
    Next i
    
End Sub


