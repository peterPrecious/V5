<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Grou.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<html>
<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

<script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

<% Server.Execute vShellHi %>

<%
  Dim vFunction, aPrograms

  vFunction = ""
  
  '...update tables
  If Request("vFunction") = "add" Then
    sExtractGrou
    sInsertGrou
  ElseIf Request.Form("vFunction") = "edit" Then
    sExtractGrou
    sUpdateGrou
  ElseIf Len(Request("vDelGrouID")) = 7 Then 
    vGrou_ID = Request("vDelGrouID")
    sDeleteGrou
  End If  
    
  If Len(Request("vHidden")) = 0 Or Request("vFunction") = "del" Then
%>

<form method="POST" action="GrouEdit.asp" target="_self">
  <p>&nbsp;</p>
  <table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td>
      <h1>Group Table</h1>
      <p>Either enter a new Group you wish to add (ie G1234EN) then click &quot;add&quot; OR click on an existing Group you wish to edit.</p></td>
      <td align="right" valign="bottom"><input type="text" name="vAddGrouID" size="13"> <input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="I1" type="image"> </td>
    </tr>
  </table>
  <input type="hidden" name="vHidden" value="Y">
</form>

<!---Edit List-->

<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
  <tr>
    <th nowrap align="left">Group Id</th>
    <th nowrap align="left">Title</th>
    <th nowrap align="left">Programs</th>
  </tr>
  <%
    '...read Grou
    sOpenDbBase2
    vSql = "Select * FROM Grou "
    Set oRsBase2 = oDbBase2.Execute(vSQL)    
    Do While Not oRsBase2.Eof 
      sReadGrou
    %>
  <tr>
    <td valign="top"><a href="GrouEdit.asp?vEditGrouID=<%=vGrou_ID%>&vHidden=n"><%=vGrou_ID%></a>&nbsp; </td>
    <td valign="top" nowrap><%=fLeft(vGrou_Title, 48)%>&nbsp; </td>
    <td valign="top">
      <%
        If Len(vGrou_Programs) > 0 Then 
          aPrograms = Split(vGrou_Programs, " ")
          For i = 0 to Ubound(aPrograms)
            vProg_Id = Left(aPrograms(i), 7)
      %> 
      <a target="_blank" href="ProgramEdit.asp?vEditProgId=<%=vProg_Id%>&vHidden=n"><%=vProg_Id%></a> 
      <%
	        Next
	      Else
	    %>
	     &nbsp; 
	    <%    
	      End If
      %> 

    </td>
  </tr>
  <%  
      oRsBase2.MoveNext
    Loop
    Set oRsBase2 = Nothing
    sCloseDbBase2    
  %>

</table>

<%
  Else
    If Len(Request.Form("vAddGrouID")) = 7 Then 
      vGrou_Id = Request.Form("vAddGrouID")
      vFunction = "add"
    ElseIf Len(Request.QueryString("vEditGrouID")) = 7 Then 
      vGrou_Id = Request.QueryString("vEditGrouID")
      vFunction = "edit"
    Else
       Response.Redirect "GrouEdit.asp"          
    End If

    '...get the values (even if trying to add)
    sGetGrou (vGrou_Id)         
%>

<form method="POST" action="GrouEdit.asp" target="_self">
  <input type="hidden" name="vFunction" value="<%=vFunction%>"><input type="hidden" name="vGrou_ID" value="<%=vGrou_ID%>"><p>&nbsp;</p>
  <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <th align="right" width="30%" valign="top" nowrap>Group Id : </th>
      <th width="70%" align="left"><%=vGrou_ID%></th>
    </tr>
    <tr>
      <th align="right" width="30%" valign="top" nowrap>Title : </th>
      <td width="70%"><input type="text" size="46" name="vGrou_Title" value="<%=vGrou_Title%>"></td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap width="30%">Active ? </th>
      <td valign="top" width="70%">
        <input type="radio" value="1" name="vGrou_Active" <%=fcheck(fsqlboolean(vgrou_active), 1)%>>Yes<br>
        <input type="radio" value="0" name="vGrou_Active" <%=fcheck(fsqlboolean(vgrou_active), 0)%>>No
      </td>
    </tr>
    <tr>
      <th align="right" width="30%" valign="top" nowrap>Description : </th>
      <td width="70%"><textarea rows="4" name="vGrou_Desc" cols="43"><%=vGrou_Desc%></textarea></td>
    </tr>
    <tr>
      <th align="right" width="30%" valign="top" nowrap>Requires : </th>
      <td width="70%"><textarea rows="4" name="vGrou_Requires" cols="43"><%=vGrou_Requires%></textarea></td>
    </tr>
    <tr>
      <th align="right" width="30%" valign="top" nowrap>Supplier : </th>
      <td width="70%"><input type="text" size="46" name="vGrou_Supplier" value="<%=vGrou_Supplier%>"></td>
    </tr>
    <tr>
      <th align="right" valign="top" nowrap width="30%">Program Strings : </th>
      <td valign="top" width="70%"><textarea rows="6" name="vGrou_Programs1" cols="43"><%=vGrou_Programs%></textarea><br>Click to access programs : <br>
      <%
        If Len(vGrou_Programs) > 0 Then 
          aPrograms = Split(vGrou_Programs, " ")
          For i = 0 to Ubound(aPrograms)
            vProg_Id = Left(aPrograms(i), 7)
      %> 
      <a target="_blank" href="ProgramEdit.asp?vEditProgId=<%=vProg_Id%>&vHidden=n"><%=vProg_Id%></a> 
      <%
	        Next
	      End If
      %> 
      <br>&nbsp; 
      </td>
    </tr>
    <tr>
      <td align="center" width="100%" valign="top" colspan="2" height="38"><br><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:jconfirm('GrouEdit.asp?vDelGrouID=<%=vGrou_ID%>&vFunction=del', '<!--webbot bot='PurpleText' PREVIEW='Ok to delete?'--><%=fPhra(000199)%>')"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a> <br>&nbsp;</td>
    </tr>
  </table>
</form>

<%
  End If  
%>

<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body></html>




