Sub checkItemsList()

Dim valuesListString As String
Dim valueArr() As String
Dim c As Long
Dim i As Long

For c = 1 To 21781
  
    valueArr = Split(Cells(c, 2), ",")

    For i = LBound(valueArr) To UBound(valueArr)-1
        
        Dim curValue As String
        Dim afterValue As String
        curValue = Replace(valueArr(i), " ", "") 
        afterValue = Replace(valueArr(i+1), " ", "") 

        If curValue <> afterValue then
            ' this condiction is for select with more precision
            ' If (Left(curValue, 2) = 25 Or Left(curValue, 2)=26) Or (Left(afterValue, 2) = 25 Or Left(afterValue, 2)=26) then
                Debug.Print curValue
                Debug.Print afterValue
                Cells(c, 3) = "True"
                Exit For
            ' End If
        End If
    Next
Next

End Sub


Sub checkExample()
    Dim valuesListString As String
    Dim valueArr() As String

    valuesListString = "7366000000000, 7366000000000, 7366000000000, 7366000000000, 7366000000000, 7366000000000, 7366000000000, 7366000000000, 7366000000000, 7366000000000, 2573660000003, 2573660000003, 2573660000003, 2573660000003, 2573660000003, 2573660000003, 2573660000003, 2573660000003, 2573660000003, 2573660000003"

    valuesListString = "1150000000001, 1150000000001, 1150000000001, 1150000000001, 1150000000001, 1150000000001, 1150000000001, 1150000000001, 1150000000001, 1150000000001"

    valueArr = Split(valuesListString2, ",")

    For i = LBound(valueArr) To UBound(valueArr)
        If i = UBound(valueArr) Then
            Exit For
        End If
        If Replace(valueArr(i), " ", "") <> Replace(valueArr(i + 1), " ", "") Then
            counter = counter + 1
            Debug.Print valueArr(i)
            Debug.Print valueArr(i + 1)
            Debug.Print "false"
            Exit For
        End If
    Next

End Sub

