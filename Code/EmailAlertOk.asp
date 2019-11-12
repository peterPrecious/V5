<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Urls_Routines.asp"-->
<!--#include virtual = "V5/Inc/Fathmail.asp"-->
<!--#include file = "EmailAlertBody.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <base target="_self">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

<%
  Dim vErr, vEmail_Notify, vEmail_Note, vEmail_TaskTitle, vEmail_Subject  
  
  vEmail_Notify = Request("vEmail_Notify")
    
  '...return if no one selected to notify 
  If fNoValue(vEmail_Notify) Then 
    Response.Redirect "EmailAlert.asp?vTskH_No=" & vTskH_No & "&vTskH_Id= " & vTskH_Id
  End If

  vEmail_Subject = Request("vEmail_Subject")
  vEmail_Note    = Request("vEmail_Note")
  vTskH_Id       = Request("vTskH_Id")
  vTskH_No       = Request("vTskH_No")

  sGetTskH svCustAcctId, vTskH_No  '...get the task title (plus parent title if level 2)
  vEmail_TaskTitle = vTskH_Title
  If vTskH_Level = 2 Then '...get parent task title
    vEmail_TaskTitle = fGetParentTitle (vTskH_AcctId, vTskH_Id, vTskH_Order) & " | " & vEmail_TaskTitle   
  End If
  
  sEmailAlert vEmail_Notify, vEmail_Subject, vEmail_Note, fPhraH(000161)
  
  '...if there's an error then don't write the last message
  If vErr = "" Then

  Server.Execute vShellHi 
%>

  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
    <tr>
      <td width="100%" align="center" class="c2"><br>
      <!--webbot bot='PurpleText' PREVIEW='Your email alerts were sent successfully'--><%=fPhra(000052)%>.<br><br>&nbsp;
      <!--webbot bot='PurpleText' PREVIEW='Remember to sign off if you are finished'--><%=fPhra(000206)%>.<br>&nbsp; </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

<% End If %>

</body>

</html>


