<%
  '...Sort an Array
  Function fSortArray(aArray)
    Dim vTemp, i, j
    '...only proceed if dealing with a valid array
    If (VarType(aArray) Or vbArray) = VarType(aArray) Then
      If Ubound(aArray) = 0 Then
        fSortArray = aArray
        Exit Function
      End If
      '...use Swap sort to find lowest in order
      For i = 0 To Ubound(aArray)
        For j = i+1 To Ubound(aArray)
          '...if less we swap values
          If aArray(j) < aArray(i) Then
            vTemp = aArray(i)
            aArray(i) = aArray(j)
            aArray(j) = vTemp
          End If
        Next
      Next
      fSortArray = aArray
    Else
      fSortArray = aArray
    End If
  End Function
  
%>  