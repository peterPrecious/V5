<%
  '...avoids using &vLang=FR
  If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
    Response.Redirect "Default.asp?vLang=FR&" & Request.ServerVariables("QUERY_STRING")
  Else
    Response.Redirect "Default.asp?vLang=FR"
  End If
%>