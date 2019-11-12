<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>

<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vModId, vPassGrade, vGrade, vDate, vMsg, oRsCheck, vMembNo

  vModId       = Request.QueryString("vModId")
  vPassGrade   = CSng(Request.QueryString("vPassGrade"))
  vMembNo      = Request.QueryString("vMembNo") '...we only pass the vMembNo when posting exams scores offline, else use svMembNo

  If Len(vPassGrade) = 0 Then vPassGrade = Cint(80)/100 
  If Len(vMembNo)    = 0 Then vMembNo    = svMembNo
  sGetMemb(vMembNo)

  vMsg         = ""

  '...did the learner pass the exam?
  sOpenDb
  vSql = "SELECT Logs_Posted, Logs_Item FROM Logs WITH (nolock) WHERE Logs_AcctId='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_MembNo=" & vMembNo & " AND Left(Logs_Item, 6) = '" & vModID & "'"
' sDebug
  Set oRsCheck = oDB.Execute(vSql)
  If oRsCheck.Eof Then
    vMsg = fPhraH(000043)    
  Else
    vGrade = Cint(Right(oRsCheck("Logs_Item"), 3)) / 100
    vDate  = fFormatDate(oRsCheck("Logs_Posted"))
'   sDebug "vGrade", vGrade
'   sDebug "vPassGrade", vPassGrade
    If  vGrade < vPassGrade Then
      vMsg = fPhraH(000076)    
    End If
  End If
  oRsCheck.Close
  Set oRsCheck = Nothing
  sCloseDB
  

  '...if no error message display cert
  If vMsg = "" Then 
    Session("CertMark")  = vGrade
    Session("CertName")  = vMemb_FirstName & " " & vMemb_LastName
    Session("CertDate")  = vDate 
    Session("CertId")    = vModId
    Session("CertTitle") = fExamTitle (vModId)
    Session("CertType")  = "Exam"
%>

    <html>
      <head>
        <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

        <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
        <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
        <title>Display the Certificate</title>
      </head>
    
      <!-- put this script in any page that must refresh the contents panel -->
      <script for="window" event="onload">
        <% Response.Write "  window.open('Certificate.asp','Certificate','toolbar=no,wIdth=650,height=425,left=100,top=100,status=no,scrollbars=no,resizable=no')"%>
        parent.frames["contents"].location.href = parent.frames["contents"].location.href;
      </script>
      
    <body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080" onload="showtime()">
    
    <% Server.Execute vShellHi %>
    <table border="1" wIdth="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
      <tr>
        <td wIdth="100%" align="center">
        <h1>
        <!--webbot bot='PurpleText' PREVIEW='Congratulations!'--><%=fPhra(000108)%></h1>
        <h2>
        <br>
        <!--webbot bot='PurpleText' PREVIEW='Your Certificate is now displayed in a separate window.'--><%=fPhra(000290)%><br>
        <!--webbot bot='PurpleText' PREVIEW='This Certificate may now be printed by pressing &lt;Ctrl&gt;+P simultaneously.'--><%=fPhra(000006)%>
        <br>&nbsp; <br><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp; </h2> </td>
      </tr>
    </table>

<% Else %>

    <html>
      <head>
        <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

        <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
        <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
        <title>Retrieve Certificate</title>
      </head>
    
      <body>
      <% Server.Execute vShellHi %>
      <table border="1" width="100%" cellpadding="5" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
        <tr>
          <td width="100%" align="center">
          <h1>
          <!--webbot bot='PurpleText' PREVIEW='Examination Certificate'--><%=fPhra(000133)%></h1>
          <br><%=vMsg%> <br><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp; </td>
        </tr>
    </table>

<% End If %>


<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</body></html></html>



