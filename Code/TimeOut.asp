<%
  '...ensure you're not in a frame
  Response.Write "<SCRIPT FOR=window EVENT=onload>"
  Response.Write "top.window.location.href='TimeOutOk.asp?vClose=Y'"
  Response.Write "</SCRIPT>"
%>

