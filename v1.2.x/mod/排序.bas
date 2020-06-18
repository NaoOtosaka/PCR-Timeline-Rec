Attribute VB_Name = "≈≈–Ú"
Public Function bubbleSort(arr As Variant)

    For i = 0 To UBound(arr)
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                k = arr(i)
                arr(i) = arr(j)
                arr(j) = k
            End If
        Next
    Next
    
    bubbleSort = arr
    
End Function


