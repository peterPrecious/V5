<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Actn.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Dial_Routines.asp"-->

<% 
  Dim vTskH_Id, vTskH_No, vAction, vSort, vTitle 
  vTskH_Id = Request("vTskH_Id")
  vTskH_No = Request("vTskH_No")
  vActn_No = Request("vActn_No") 
  vSort    = Request("vSort")
  
  If Request("vAction") = "Delete"    Then sDeleteActn
  If Request("vAction") = "Completed" Then sCompletedActn
%>

<html>

<head>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <base target="_self">
</head>


<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td width="100%">
      <form method="POST" action="ActnList.asp">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
            <td colspan="2"><font face="Verdana" size="1"><b>Action Items</b> are directives issued to a colleague that remain on file for 90 days.&nbsp; Once issued that colleague becomes the owner of the Action Item which is completed when that owner clicks &quot;completed&quot;.&nbsp; To post a new Action Item click &quot;action item&quot;.&nbsp; To edit an exiting Action Item, &quot;goto&quot; that item and edit its contents.&nbsp; You should notify the owner of the new action item via the &quot;email alert&quot; feature (if configured).&nbsp; <br></font>&nbsp; </td>
          </tr>
          <tr>
            <td><b><font face="Verdana" size="1">
            <!--webbot bot='PurpleText' PREVIEW='Sort by'--><%=fPhra(000243)%> : </font></b><select size="1" name="vSort">
            <option value="Actn_Items">
            <!--webbot bot='PurpleText' PREVIEW='Action Item'--><%=fPhra(000061)%></option>
            <option value="Memb_LastName">
            <!--webbot bot='PurpleText' PREVIEW='Owner'--><%=fPhra(000210)%></option>
            <option value="Actn_Posted">
            <!--webbot bot='PurpleText' PREVIEW='Date Posted'--><%=fPhra(000114)%></option>
            <option value="Actn_Due" selected>
            <!--webbot bot='PurpleText' PREVIEW='Date Due'--><%=fPhra(000113)%></option>
            </select> <input border="0" src="../Images/Buttons/Go_<%=svLang%>.gif" name="I1" type="image"> </td>
            <td align="right"><a <%=fStatX%> href="ActnAdd.asp?vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>"><img border="0" src="../Images/Buttons/ActionItem_<%=svLang%>.gif" alt="&lt;--~0000--&gt;Generate an Action Item"></a> </td>
          </tr>
        </table>
      </form>
      </td>
    </tr>
    <tr>
      <td width="100%">
      <table border="1" width="100%" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top"><font face="Verdana" size="1"><b>
          <!--webbot bot='PurpleText' PREVIEW='Action Item'--><%=fPhra(000061)%> </b></font></td>
          <td valign="top"><font face="Verdana" size="1"><b>
          <!--webbot bot='PurpleText' PREVIEW='Owner'--><%=fPhra(000210)%></b></font></td>
          <td valign="top" align="center"><font face="Verdana" size="1"><b>
          <!--webbot bot='PurpleText' PREVIEW='Date Posted'--><%=fPhra(000114)%><br>
          <!--webbot bot='PurpleText' PREVIEW='Date Due'--><%=fPhra(000113)%></b></font></td>
          <td valign="top" align="center"><font face="Verdana" size="1"><b>
          <!--webbot bot='PurpleText' PREVIEW='Completed'--><%=fPhra(000107)%><br></b><img border="0" src="../Images/Icons/Check.gif"></font></td>
          <td valign="top">&nbsp;</td>
        </tr>
        <%
        If fNoValue(vSort) Then vSort = "Actn_Due"
        
        sGetActn_rs
        Do While Not oRs.Eof
          sReadActn
          vActn_Item = fHtmlList(Trim(vActn_Item))
          If Len(vActn_Item) > 80 Then
            vActn_Item = Left(vActn_Item, 80) & "  ..."
          End If
          vTitle = "Action Item"
      %> <tr>
          <td valign="top" width="50%"><font face="Verdana" size="1"><%=vActn_Item%></font>&nbsp; </td>
          <td valign="top"><b><font face="Verdana" size="1" color="#3977B6"><%=vMemb_FirstName & " " & vMemb_LastName%></font></b>&nbsp; </td>
          <td valign="top" align="center"><font face="Verdana" size="1" color="#3977B6"><%=fFormatDate(vActn_Posted)%></font> <font face="Verdana" size="1"><br><b><font color="#FF0000"><%=fFormatDate(vActn_Due)%></font></b></font> </td>
          <td valign="top" align="center"><% If vActn_Completed Then %> <img border="0" src="../Images/Icons/Check.gif"> <% Else %>&nbsp;&nbsp; <% End If %> </td>
          <td valign="top" align="right"><a <%=fStatX%> href="ActnAdd.asp?vActn_No=<%=vActn_No%>&vAction=Edit&vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>"><img border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif" alt="<%=Server.UrlEncode(fPhraH(000125))%>"></a><a <%=fStatX%> href="ActnList.asp?vActn_No=<%=vActn_No%>&vAction=Delete&vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif" alt="<%=Server.UrlEncode(fPhraH(000301))%>"></a><br><a <%=fStatX%> href="ActnList.asp?vActn_No=<%=vActn_No%>&vAction=Completed&vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>"><img border="0" src="../Images/Buttons/Completed_<%=svLang%>.gif"></a> </td>
        </tr>
        <%
          oRs.MoveNext
        Loop
        sCloseDB
      %>
      </table>
      </td>
    </tr>
    <tr>
      <td valign="top" align="center"><br><a <%=fStatX%> href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>"><img border="0" src="../Images/Icons/World.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

