Attribute VB_Name = "modQuickSort"
Option Explicit

Public Sub defSortCollectionByIndex(col As Collection)
    Dim arr() As Variant

    ' Collection in Array kopieren
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        Set arr(i - 1) = col.item(i)
    Next i

    ' Array sortieren
    defQuickSort arr, 0, UBound(arr), False

    ' Collection neu aufbauen
    For i = col.Count To 1 Step -1
        col.Remove i
    Next i
    For i = 0 To UBound(arr)
        col.Add arr(i)
    Next i
End Sub

Public Sub defQuickSort(arr() As Variant, ByVal low As Long, ByVal high As Long, Optional ByVal isStringArray As Boolean = True)
    Dim pivotValue As Long
    Dim tempSwap As Variant
    Dim i As Long, j As Long
    
    If low < high Then
        pivotValue = funcGetValueForSorting(arr((low + high) \ 2), isStringArray)
        i = low
        j = high
        
        Do While i <= j
            Do While funcGetValueForSorting(arr(i), isStringArray) < pivotValue
                i = i + 1
            Loop
            Do While funcGetValueForSorting(arr(j), isStringArray) > pivotValue
                j = j - 1
            Loop
            If i <= j Then
                ' Tausche
                If isStringArray Then
                    tempSwap = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tempSwap
                Else
                    Set tempSwap = arr(i)
                    Set arr(i) = arr(j)
                    Set arr(j) = tempSwap
                End If
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If low < j Then defQuickSort arr, low, j, isStringArray
        If i < high Then defQuickSort arr, i, high, isStringArray
    End If
End Sub

Public Function funcGetValueForSorting(item As Variant, isStringArray As Boolean) As Long
    If isStringArray Then
        funcGetValueForSorting = CLng(Split(item, "_")(3))
    Else
        funcGetValueForSorting = item.Index
    End If
End Function
