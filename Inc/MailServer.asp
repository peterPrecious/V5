<%
  Sub sMailServer

    '...this gets the IP address of the router in front of this workstation
    '   used in V5 for determining what mail server is active

    Dim oXml, vIp
    Set oXml = CreateObject("Microsoft.XMLHTTP")

    '...staging firewall will restrict this call
    On Error Resume Next
    oXml.Open "POST", "//vubiz.com/ws/wsMyIp.asp", False
    oXml.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    oXml.Send ""
    vIp = oXml.ResponseText
    On Error Goto 0

    '...peter home office
    If vIp = "66.135.96.87" Then                                 
      Session("MailServer") = "mail.netsurf.net"

    '...staging server (temp)  
    ElseIf Instr(Lcase(Request.ServerVariables("HTTP_HOST")), "staging2") > 0 Then                         
      Session("MailServer") = "smtp.cogeco.net; smtp.cogeco.ca"

    '...oakville office
    ElseIf vIp = "72.38.29.238"  Then                            
      Session("MailServer") = "smtp.cogeco.net; smtp.cogeco.ca"

    '...production server
    Else                                                         
      Session("MailServer") = "vubizmail.dades.ca"
    End If

    svMailServer = Session("MailServer") '...normally done in initialize.asp on each page

'   Response.Write vIp & "<br>" & Session("MailServer") & "<br>" & Request.ServerVariables("HTTP_HOST")

  End Sub 
%>  