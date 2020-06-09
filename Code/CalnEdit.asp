<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Caln.asp"-->

<%
  Dim vActive, vTskH_Id, vTskH_No

  vActive    = Request("vActive")
  vCaln_Date = Request("vCaln_Date")  
  vTskH_Id   = Request("vTskH_Id")
  vTskH_No   = Request("vTskH_No")
  
  '...From Calendar
  If fNoValue(vActive) Then 
    sGetCaln Cdate(vCaln_Date), vTskH_No
  Else '...from form    
    vCaln_Details = Request("vCaln_Details")  
    sInsertCaln
    Response.Redirect "CalnList.asp?vCaln_Date=" & vCaln_Date & "&vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
  End If
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>

  <form method="POST" action="CalnEdit.asp" target="_self">
    <input type="hidden" name="vActive" value="y"><input type="hidden" name="vCaln_Date" value="<%=vCaln_Date%>"><input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>"><input type="hidden" name="vTskH_No" value="<%=vTskH_No%>">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" valign="top" colspan="2"><font face="Verdana" size="1"><b>Add/Edit Calendar<br></b>Add or edit the calendar activity list - <font color="#FF0000">be careful you do not remove other entries</font>.<br>&nbsp;</font></td>
      </tr>
      <tr>
        <td align="right" width="30%" valign="top"><font face="Verdana" size="1"><b>Date :&nbsp; </b></font></td>
        <td width="70%" valign="top"><font face="Verdana" size="1"><%=vCaln_Date%> </font></td>
      </tr>
      <tr>
        <td align="right" width="30%" valign="top"><font face="Verdana" size="1"><b>&nbsp; Activities :&nbsp; </b></font></td>
        <td width="70%"><textarea rows="10" name="vCaln_Details" cols="47"><%=vCaln_Details%></textarea> </td>
      </tr>
      <tr>
        <td colspan="2" valign="top" align="center"><br>&nbsp;<a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I1" type="image"> <br>&nbsp;</td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

