Attribute VB_Name = "sortMod"
Public Function bubbleSort(arr As Variant)

    For i = 0 To UBound(arr)
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                K = arr(i)
                arr(i) = arr(j)
                arr(j) = K
            End If
        Next
    Next
    
    bubbleSort = arr
    
End Function


