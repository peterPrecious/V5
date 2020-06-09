<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Actn.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%  
  Dim vTskH_Id, vTskH_No, vAction, vDays, vOptionList1, vOptionList2 

  vAction       = Request("vAction")
  vActn_No      = Request("vActn_No") 
  vTskH_Id      = Request("vTskH_Id")
  vTskH_No      = Request("vTskH_No")

  sExtractActn

  If vActn_Item <> "" And IsNumeric(vActn_Owner) Then
    If vAction = "Edit" Then
      sUpdateActn
    Else  
      sInsertActn
    End If  
    Response.Redirect "ActnList.asp?vTskH_No=" & vTskH_No
  Else
    If vAction = "Edit" Then
      sGetActn (vActn_No)
    End If 
  End If


  '...build drop down combo box of Members (who are eligible to enter vTskH_No - coming)
  Dim aMembs, aMemb, vSelected
  aMemb  = "" : j = "" : vSelected = ""
  aMembs = Split(fMemb_List, "~~")
  For i = 0 to uBound(aMembs)
    aMemb = Split(aMembs(i), "~")
    If Clng(aMemb(1)) = vActn_Owner Then vSelected = "selected"
    vOptionList1 = vOptionList1 & "<option " & vSelected & " value='" & aMemb(1) & "'>" & aMemb(0) & "</option>" & vbNewLine
  Next

  '...build drop down combo box of Due Dates
  vOptionList2 = ""
  For i = 0 to 60
    vOptionList2 = vOptionList2 & "<option value='" & fFormatSqlDate(DateAdd("d", i, now)) & "'" & fSelect(fFormatDate(vActn_Due), fFormatDate(DateAdd("d", i, now))) & ">" & fFormatDate(DateAdd("d", i, now)) & "</option>" & vbCrLf
  Next  
%>

<html>

<head>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>

</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <form method="POST" action="ActnAdd.asp">
    <input type="hidden" name="vAction" value="<%=vAction%>"><input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>"><input type="hidden" name="vTskH_No" value="<%=vTskH_No%>"><input type="hidden" name="vActn_No" value="<%=vActn_No%>">
    <table border="1" width="100%" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="0" cellspacing="0">
      <tr>
        <td width="100%" valign="top" colspan="2"><font face="Verdana" size="1"><b>Add an Action Item</b>: Enter your instructions and select the action item &quot;owner&quot;.&nbsp; Suggest an completion date then &quot;update&quot;.&nbsp; NOTE: the &quot;Completed&quot; feature is only used to change an existing status.&nbsp; It will default to &quot;No&quot;.<br></font></td>
      </tr>
      <tr>
        <td width="30%" valign="top" align="right"><font face="Verdana" size="1"><b>Instructions :&nbsp;&nbsp; </b></font></td>
        <td width="70%" valign="top">&nbsp;<textarea rows="7" cols="40" name="vActn_Item"><%=vActn_Item%></textarea></td>
      </tr>
      <tr>
        <td width="30%" align="right" valign="top"><font face="Verdana" size="1"><b>Owner :&nbsp;&nbsp; </b></font></td>
        <td width="70%" valign="top">&nbsp;<select size="1" name="vActn_Owner"><%=vOptionList1%></select> </td>
      </tr>
      <tr>
        <td width="30%" valign="top"><p align="right"><font face="Verdana" size="1"><b>Complete before :&nbsp;&nbsp; </b></font></p>
        </td>
        <td width="70%" valign="top"><font face="Verdana" size="1">&nbsp;<select size="1" name="vActn_Due"><%=vOptionList2%></select> (Today&#39;s date is <%=fFormatDate(Now)%>)</font></td>
      </tr>
      <% If vAction = "Edit" Then %> <tr>
        <td width="30%" align="right" valign="top"><font face="Verdana" size="1"><b>Completed? :&nbsp;&nbsp; </b></font></td>
        <td width="70%" valign="top"><input type="radio" value="0" <%=fcheck(0, fsqlboolean(vactn_completed))%> name="vActn_Completed"><font face="Verdana" size="1">No</font><br><input type="radio" value="1" <%=fcheck(1, fsqlboolean(vactn_completed))%> name="vActn_Completed"><font face="Verdana" size="1">Yes</font></td>
      </tr>
      <% Else %> <input type="hidden" name="vActn_Completed" value="0"><% End If %> <tr>
        <td width="100%" colspan="2" align="center" valign="top">&nbsp;&nbsp;&nbsp;&nbsp; <br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I2" type="image"></td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>