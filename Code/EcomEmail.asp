<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Fathmail.asp"-->

<%
  Dim vCust, vId, vSubject, vSender, vBody, vRecipients, vResult

  vCust = Ucase(fDefault(Request("vCust"), svCustId))
  vId   = Ucase(fDefault(Request("vId"), svMembId))

  If Request.Form.Count > 0 Then
    vSubject  = fFr("Welcome to Vubiz")
    vSender   = fPhraH(000434)
    vSender   = "Registration"
    
    sGetCust  vCust
    sGetMembById  Right(vCust, 4), vId
    
    vMemb_Email = "pbulloch@vubiz.com"    
    vCust_StartURL = fDefault(vCust_StartURL, "vubiz.com")

    vBody = "" _
          & "<p style='text-align: left; FONT-SIZE: 8.0pt; FONT-WEIGHT: normal; FONT-FAMILY: Verdana,Geneva,Arial,Helvetica; COLOR: #3977B6'>" & vbCrLf _
          & "<br><b>Thank you and Welcome!</b> <br><br>For your records... <br><br>" & vbCrLf _
          & "<ul style='text-align: left; FONT-SIZE: 8.0pt; FONT-WEIGHT: normal; FONT-FAMILY: Verdana,Geneva,Arial,Helvetica; COLOR: #3977B6'>" & vbCrLf _
          & "<li>Your Customer ID is: " & vCust & "</li>" & vbCrLf _
          & "<li>Your Password is: " & vId & "</li>" & vbCrLf _
          & "<li>To access your Learning go to: <a href='//vubiz.com'>//" & vCust_StartURL & "</a></li>" & vbCrLf _
          & "<li>Please Sign Off after every session;</li>" & vbCrLf _
          & "<li>Feel free to <a href='mailto:info@vubiz.com'>email us</a> if you have any questions</li" & vbCrLf _
          & "<ul>" & vbCrLf _
          & "</p>" & vbCrLf 
  
  
    vBody       = fFr(vBody)
    vRecipients = fFr(vMemb_FirstName) & " " & fFr(vMemb_LastName) & " <" & vMemb_Email & ">; "
    vResult     = fFathMail(vSubject, vBody, vSender, vRecipients)   
    
  End If


  Function fFr (vPhrase)
    fFr = vPhrase

    fFr = Replace(fFr, "à", "&#224;") 
    fFr = Replace(fFr, "ç", "&#231;") 
    fFr = Replace(fFr, "è", "&#232;") 
    fFr = Replace(fFr, "é", "&#233;") 
    fFr = Replace(fFr, "ê", "&#234;") 

    fFr = Replace(fFr, "À", "&#192;") 
    fFr = Replace(fFr, "Ç", "&#199;") 
    fFr = Replace(fFr, "È", "&#200;") 
    fFr = Replace(fFr, "É", "&#201;") 
    fFr = Replace(fFr, "Ê", "&#202;") 
  End Function


%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %> 

  <form method='POST' action='EcomEmail.asp'>
    <div align="center">
    <br>&nbsp;<table border="1" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <th colspan="2">
        <h1><br>CREATE TEST EMAIL <br><br>using first name, last name, email address and the start URL <br>for this learner (note, ensure the above fields are valid) ...</h1>
        </th>
      </tr>
      <tr>
        <td align="right">Customer ID : </td>
        <td><input type='text' name='vCust' size='11' value='<%=vCust%>' class="c2"></td>
      </tr>
      <tr>
        <td align="right">Password : </td>
        <td><input type='text' name='vId' size='43' value='<%=vId%>' class="c2"></td>
      </tr>
      <tr>
        <td align="center" colspan="2" height="50"><input type='submit' value='Submit' name='B1' class="button">&nbsp;&nbsp;&nbsp; <%=vResult%></td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

