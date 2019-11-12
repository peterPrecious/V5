<%
  Dim aFeatures(12)
  sRandomize 12
  sDisplay 12

  '...create an array of vMax random numbers
  Sub sRandomize (vMax)
    Dim bFeatures
    Redim bFeatures(vMax)
    
    '...initialize
    For i = 1 to vMax
      aFeatures(i) = 0
      bFeatures(i) = 0
    Next
    '...try 5 times to fill array
    For k = 1 to 5
      Randomize
      For i = 1 to vMax
       j = Int(vMax * Rnd + 1)
       '...if new number add to array
       If bFeatures(j) = 0 Then
         '...find first available spot
         For l = 1 to vMax
           If aFeatures(l) = 0 Then
             aFeatures(l) = j
             Exit For
           End If
         Next
         '...signify that number is taken 
         bFeatures(j) = 1  
       End If
      Next
  '  sDisplay
    Next
    '...If any numbers still not fill then just fill
    For i = 1 to vMax
      If bFeatures(i) = 0 Then
        '...find first available spot
        For l = 1 to vMax
          If aFeatures(l) = 0 Then
            aFeatures(l) = i
            Exit For
          End If
        Next
      End If
    Next
  ' sDisplay     
  End Sub

  
  Sub sDisplay (vMax)
    response.write "<p>" 
    for l = 1 to vMax
'     response.write "<br>" &   aFeatures(l) & " " & bFeatures(l)
      response.write "<br>" &   aFeatures(l)
    next
  End Sub
 
%>