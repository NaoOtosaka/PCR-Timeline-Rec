VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Query 
   Caption         =   "人物选择"
   ClientHeight    =   2628
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6240
   OleObjectBlob   =   "Query.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Myr&, arrsj, m%, lb
Private Sub CommandButton1_Click() '确定
    For i = 0 To 1
        ActiveCell.Offset(0, i) = Me.ListBox1.List(Me.ListBox1.ListIndex, i + 1)
    Next
    Unload Me
    ActiveCell.Offset(0, 3).Select
End Sub

Private Sub CommandButton2_Click() '退出
    Unload Me
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
    For i = 0 To 1
        ActiveCell.Offset(0, i) = Me.ListBox1.List(Me.ListBox1.ListIndex, i + 1)
    Next
    Unload Me
    ActiveCell.Offset(0, 3).Select
End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    For i = 0 To 1
        ActiveCell.Offset(0, i) = Me.ListBox1.List(Me.ListBox1.ListIndex, i + 1)
    Next
    Unload Me
    ActiveCell.Offset(0, 3).Select
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i&, j%, t%, myStr$, k&, N$, P$
    Dim LG As Boolean, arr1()
    On Error Resume Next
    Me.ListBox1.Clear
    myStr = UCase(Me.TextBox1.Value)
    For i = 1 To Len(myStr)
        If Asc(Mid$(myStr, i, 1)) < 0 Then LG = True: Exit For
    Next
    arr2 = Array("ID", "角色名", "位置")
    k = k + 1
    ReDim arr1(1 To 3, 1 To k)
    For i = 1 To 3
        arr1(i, k) = arr2(i - 1)
    Next
    For i = 1 To UBound(arrsj)
        s = arrsj(i, 1) & arrsj(i, 2) & arrsj(i, 3)
        N = ""
        If LG Then
            N = s
        Else
            For j = 1 To Len(s)
                P = Mid(s, j, 1)
                If Asc(P) < 0 Then N = N & PinYin(P) Else N = N & P
            Next
        End If
        If InStr(N, myStr) Then
            k = k + 1
            ReDim Preserve arr1(1 To 3, 1 To k)
            For t = 1 To 3
                arr1(t, k) = arrsj(i, t)
            Next
        End If
    Next i
   
    If t = 0 Then Exit Sub
    Me.ListBox1.List = Application.Transpose(arr1)
    Me.ListBox1.Selected(1) = True
End Sub
Private Sub UserForm_Initialize()

    arrsj = Sheets("人物").Range("A2:C" & Sheets("人物").[A65536].End(3).Row)
    
    Me.ListBox1.List = arrsj
    
    Me.ListBox1.Selected(0) = True
    
End Sub
