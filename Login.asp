<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 
  '...This provides quick access to the FAQ section typically from SignOffOk.asp
  Dim vParms, vLang, vCust, vId

  vLang = fDefault(Request("vLang"), "EN")

  '...customer and user id passed through?
  vCust = Request("vCust")
  If Len(vCust) > 0 Then 
    vParms = Server.UrlEncode("?vCust=" & vCust)
    vId = Request("vId")
    If Len(vId) > 0 Then
      vParms = vParms & Server.UrlEncode("&vId=" & vId)
    End If    
  Else
    vParms = ""
  End If

 
  Response.Redirect "/Chaccess/SignIn/Default.asp?vLang=" & vLang & "&vCust=" & vCust & "&vId=" & vId
%>
