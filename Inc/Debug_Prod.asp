<%
  '...display the product session variables if Debug
  Dim i, j
  saProd = Session("Prod")
  For i = 1 To Session("ProdMax")
    For j = 0 To 6
      Response.Write "<br><b><font color='ORANGE'>" & "Prod(" & j & ", " & i & ")" & " : " & saProd(j, i) & "</font></b>"
    Next
  Next
%>