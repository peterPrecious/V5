<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<% 
   If Len(Session("QueryString")) = 0 Then Response.Redirect "/V5/Default.asp?vGoto=/V5/Public/Cat_Default.asp"
   Session("TabActive") = True '...this ensures all lower frames do not use a top border but use the border from Cat_Header
   sGetCust svCustId '...this gets the content options for this customer
%>

<html>

<head>
  <% If Len(Session("MultiUserManual")) = 0 Then %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <script>
    // get the MultiUserManual repository path and put into a session variable - note nothing is returned
    var vWs = WebService("/V5/Repository/Documents/MultiUserManual/MultiUserManual_ws.asp", "")
  </script>
  <% End If %>  
  <title>:: Vubiz</title>
</head>

<frameset border="0" frameborder="0" framespacing="0" rows="76,*">
  <frame marginheight="0" marginwidth="0" name="tabs" src="Cat_Header.asp" target="main" scrolling="no" noresize>
  <frame marginheight="0" marginwidth="0" name="main" src="/V5/Code/Ecom2Start.asp?vMode=More&vContentOptions=<%=vContentOptions%>" target="_self" scrolling="auto" noresize>
</frameset>

</html>