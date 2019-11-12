<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim oXmlDom
  Dim vCust, vId, vAction, vResponse

  If Len(Request.Form) > 0 Then  
    Set oXmlDom = CreateObject("msxml2.DOMDocument.4.0")
    oXmlDom.LoadXml Request.Form
    vAction = oXmlDom.selectSingleNode("/VUBIZ/WS").Attributes.getNamedItem("vAction").Text
    vCust   = oXmlDom.selectSingleNode("/VUBIZ/WS").Attributes.getNamedItem("vCust").Text
    vId     = oXmlDom.selectSingleNode("/VUBIZ/WS").Attributes.getNamedItem("vId").Text 
  End If

  If vAction = "GetCatalogue" Then 
    vResponse = fResponse
  Else
    Response.Redirect "WShelp.asp"
  End If

  '...return the xml string  
  Response.Clear
  Response.Buffer = True
  Response.Write vResponse
  Response.End

  '...getting xml string from a sql stored procedure
  Function fResponse
    sOpenDbBase
    Set oRsBase = oDbBase.Execute ("Ws_GetCatalogue")
    fResponse = oRsBase.Fields(0).Value
    sCloseDbBase
  End Function

%>