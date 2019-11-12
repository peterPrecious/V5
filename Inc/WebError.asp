<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 
  Dim vLang, vMsg
  vLang = Session("Lang")
  If Len(vLang) <> 2 Or Instr(" EN FR ES PT ", vLang) = 0 Then vLang = "EN"
  
  Select Case vLang
    Case "EN" : vMsg = "The document you requested does not seem to be available.<br>If you know you have the correct URL, please email <a href='mailto:support@vubiz.com'>support@vubiz.com</a>. Thank you."
    Case "FR" : vMsg = "Le document que vous avez demand� ne semble pas �tre disponible.<br> Si vous savez que vous avez l�URL correct, veuillez envoyer un courriel � <a href='mailto:support@vubiz.com'>support@vubiz.com</a>. Merci."
    Case "ES" : vMsg = "El documento que usted solicit� no parece estar disponible.<br> Si sabe que tiene el URL correcto, s�rvase enviar un mensaje por correo electr�nico � <a href='mailto:support@vubiz.com'>support@vubiz.com</a>. Gracias."
    Case "PT" : vMsg = "O documento que pediu n�o se encontra dispon�vel.<br> Se tiver a certeza de ter o URL correcto, � favor enviar mensagem electr�nica para: <a href='mailto:support@vubiz.com'>support@vubiz.com</a>. Obrigado."
  End Select
%>

<html><head><title></title></head>
<body text="#000080" vLink="#000080" aLink="#000080" link="#000080" bgColor="#ffffff" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">

<!--#include virtual = "V5/Inc/Shell_Top.asp"-->
<table cellSpacing="0" cellPadding="0" border="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td width="100%" align="center"><font face="Verdana" size="1"><%=vMsg%></font></td>
  </tr>
</table>
<!--#include virtual = "V5/Inc/Shell_Bottom.asp"-->
</body></html>