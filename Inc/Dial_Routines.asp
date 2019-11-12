<%
  '...enable text to be formatted with lf/cf
  Function fHtmlList(vText)
    If fNoValue(vText) Then vText = ""
    fHtmlList = ""
    '...allow multi spaces
    vText = Replace(vText, "  ", "&nbsp;&nbsp;")
    For i = 1 to Len(vText)
      j = Mid (vText, i, 1)
      If ASC(j) = 10 Then 
        fHtmlList = fHtmlList & "<br>"
      ElseIf Asc(j) <> "13" Then
        fHtmlList = fHtmlList & j
      End If                        
    Next     
  End Function
  
%>