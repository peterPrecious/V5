<%
  Option Explicit

  Server.ScriptTimeout = 60 * 00

  Dim oMsg, oMail 
  Set oMsg  = Server.CreateObject("FathMail.Message")
  Set oMail = Server.CreateObject("FathMail.SMTP")
  
  oMsg.Subject      = "Test message!"
  oMsg.TextHTML     = "Test of the message body...<br>and a beauty it is..."
  oMsg.Sender       = "Peter Bulloch <peter@bullochonline.com>"
  oMsg.Recipients   = "Peter Bulloch <peter@bullochonline.com>"

  oMail.LoginMethod = 2
  oMail.Username    = "info@vubiz.com"
  oMail.Password    = "vubizpass"
  oMail.ServerPort  = 8025
  oMail.ServerAddr  = "anywhere.exchserver.com"
  oMail.Send oMsg

  If Err.Number     = 0 Then
  	Response.Write "Ok"
  Else
  	Response.Write "Error: " & oMail.LastCommandResponse 
  End If
  oMail.Disconnect

  Set oMsg  = Nothing
  Set oMail = Nothing
%>