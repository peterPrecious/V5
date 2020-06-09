<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Urls_Routines.asp"-->

<%
  '...ensure valid from name
  If Len(Trim(svMembFirstName & " " & svMembLastName)) = 0 Then
    Response.Redirect "Error.asp?vErr=" & Server.UrlEncode(fPhraH(000283))
  End If
  
  '...generate a list members
  Dim vMembList, vOptionA, vOptionB, vOptionList, iMax, vTskH_Id, vTskH_No, vMembNo, vMembName
  vTskH_Id    = Request("vTskH_Id")
  vTskH_No    = Request("vTskH_No")

  '...FirstName LastName ~ No ~ Valid Email ~~
  vOptionA    = Split(fMemb_List, "~~") '...contains all the learners with valid emails
  vOptionList = "<table border='0' style='border-collapse: collapse' cellpadding='0'>"
  For i = 0 to uBound(vOptionA)
    vOptionB = Split(vOptionA(i), "~")          
    vOptionList = vOptionList _
                & " <tr> "_
                & "   <td> "_
                & "     <input type='checkbox' name='vEmail_Notify' value='" & vOptionB(1) & "'>" _
                & "   </td> "_
                & "   <td> "_
                &       vOptionB(0) & f10() _
                & "   </td> "_
                & "   <td> "_
                &       vOptionB(2) _
                & "   </td> "_
                & " </tr> " & vbCrLf
  Next
  vOptionList = vOptionList & "</table>"
  iMax = i  '...no of entries for the jScript
  
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <script language="JavaScript">
    function CheckAll(){
      for (var i=0;i<document.SelectMembers.elements.length;i++){
        var j = document.SelectMembers.elements[i];
        if (j.name == 'vEmail_Notify')
          j.checked = document.SelectMembers.checkall.checked;
      }
    }
  </script>  
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <%
    Server.Execute vShellHi 
  %>
  <form name="SelectMembers" method="POST" action="EmailAlertOK.asp">
    <input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>">
    <input type="hidden" name="vTskH_No" value="<%=vTskH_No%>">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3" cellspacing="0">
      <tr>
        <td width="100%" valign="top" align="center" colspan="2">
        <h1 align="left">Email Alert</h1>
        <h2 align="left">
        <!--webbot bot='PurpleText' PREVIEW='This allows you to send an email to the learner(s) selected below.&nbsp; The preformatted note will include a link to this site, a reference to the current task plus any additional message that you enter below.'--><%=fPhra(000384)%></h2>
        </td>
      </tr>
      <!-- If sending to one member then use hidden fields, else option list -->
      <% If vMembNo <> "" Then  %> 
      <input type="hidden" name="vEmail_Notify" value="<%=vMembNo%>">
      <tr>
        <td valign="top" colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <th valign="top" nowrap align="right" width="30%">
        <!--webbot bot='PurpleText' PREVIEW='Send alert to'--><%=fPhra(000236)%> :</th>
        <td width="70%"><%=vMembName%></td>
      </tr>
      <% Else %>
      <tr>
        <th valign="top" nowrap align="right" width="30%">
          <!--webbot bot='PurpleText' PREVIEW='Send alert to'--><%=fPhra(000236)%> :</th>
        <td width="70%">
          <%=vOptionList%>
          <% If iMax > 1 Then %>
          <br><br>          
          <input onclick="CheckAll();" type="checkbox" value="Check All" name="checkall">
          <!--webbot bot='PurpleText' PREVIEW='(all/none)'--><%=fPhra(000055)%> 
          <% End If %>
        </td>
      </tr>
      <% End If %>
      <tr>
        <th valign="top" nowrap align="right" width="30%">Email Subject :</th>
        <td width="70%"><input type="text" name="vEmail_Subject" size="60" value="<%=svCustTitle%> - Alert"> <br>Enter above how you would like the Subject line to appear in learner&#39;s email client.</td>
      </tr>
      <tr>
        <th valign="top" nowrap align="right" width="30%">Additional Message :</th>
        <td width="70%"><textarea rows="8" name="vEmail_Note" cols="46"></textarea></td>
      </tr>
      <tr>
        <td valign="top" align="center" colspan="2"><br><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>"><img border="0" src="../Images/Icons/World.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" name="I1" type="image">&nbsp;&nbsp; <br><br>
        <!--webbot bot='PurpleText' PREVIEW='Please allow a few minutes to send email(s)'--><%=fPhra(000214)%>.<br>&nbsp;</td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


