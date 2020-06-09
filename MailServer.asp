<%
  Sub sMailServer

    '...this determines what mail server is active depending on where we are running
    '...make the appopriate one active for the appropriate server
    
'   Session("MailServer") = "smtp.broadband.rogers.com"          '...peter home
    Session("MailServer") = "smtp.cogeco.net; smtp.cogeco.ca"    '...oakville
'   Session("MailServer") = "smtp.ultrahosting.com"              '...production

    svMailServer = Session("MailServer") '...normally done in initialize.asp on each page

  End Sub 
%>  