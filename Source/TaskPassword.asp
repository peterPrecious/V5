<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Password_Routines.asp"-->
<%
   Dim vPassword, vMsg
   vMsg = ""

   '...initialize count to 0 on first pass and go to the pass
   If Session("MyWorld_PasswordAttemps") = "" Then Session("MyWorld_PasswordAttemps") = 0

   Session("MyWorld_PasswordAttemps") = Session("MyWorld_PasswordAttemps") + 1

   If Session("MyWorld_PasswordAttemps") > 3 Then 
     vMsg = "<!--{{-->Access Denied<!--}}-->"
   ElseIf Session("MyWorld_PasswordAttemps") > 1 Then 
     vMsg = "<!--{{-->Access attemp:<!--}}-->" & " " & Session("MyWorld_PasswordAttemps")
   End If
   
   If Request("vHidden") = "y" Then
     vPassword  = fEncode(Ucase(Trim(Request("vPassword"))))
     If vPassword = Session("MyWorld_Password") Then
       Session("MyWorld_PasswordEntered") = vPassword
       Session("MyWorld_PasswordAttemps") = ""
       Response.Redirect Session("MyWorld_Url")
     End If
   End If
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <form method="POST" action="TaskPassword.asp">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
      <% If vMsg <> "<!--{{-->Access Denied<!--}}-->" Then %> <tr>
        <td colspan="2" align="center"><h1>Password.</h1><h2><br>
        <!--[[-->You are about to enter a secure section.&nbsp; Please enter your password to proceed.<!--]]--><br>&nbsp;</h2></td>
      </tr>
      <% End If %> <% If Len(vMsg) > 0 Then %> <tr>
        <td colspan="2" align="center"><br><%=vMsg%><br>&nbsp;</td>
      </tr>
      <% End If %> <% If vMsg <> "<!--{{-->Access Denied<!--}}-->" Then %> <tr>
        <th align="right" width="50%">Password :&nbsp; </th>
        <td width="50%">&nbsp;<input type="password" size="29" name="vPassword" maxlength="16"></td>
      </tr>
      <tr>
        <td align="center" colspan="2"><br><% If Session("MyWorld_PasswordAttemps") = 1 Then %> <a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <% End If %> <input border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" type="image"><br>&nbsp;</td>
      </tr>
      <% End If %>
    </table>
    <input type="hidden" name="vHidden" value="y">
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>