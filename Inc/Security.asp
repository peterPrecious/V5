<%
  '...stop if not secure and no bypass
  If Not Session("Secure") And Not vBypassSecurity Then Response.Redirect "Timeout.asp?vPage=" & Request.ServerVariables("Path_Info")
%>