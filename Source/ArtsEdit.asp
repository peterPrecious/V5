<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Arts.asp"-->
<% Session("HostDb") = "V5_Vubz"  '...set since bypassing "signin" %>

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
  
  <%
    Dim vFunction
    vFunction = ""
    
    '...update tables
    If Request("vFunction") = "add" Then
      sExtractArts
      sInsertArts
    ElseIf Request.Form("vFunction") = "edit" Then
      sExtractArts
      sUpdateArts
    ElseIf Len(Request("vDelArtsNo")) > 0 Then 
      vArts_No = Request("vDelArtsNo")
      sDeleteArts
    End If  
      
    If Len(Request("vForm")) = 0 Or Request("vFunction") = "del" Then
  %>

  <form method="POST" action="ArtsEdit.asp" target="_self">
    <table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td>
        <h1>Articles</h1>
        <h2>Click &quot;add&quot; to enter a new article OR click on an existing Article No you wish to edit or delete OR click on the Title if you wish to view the Article.</h2>
        </td>
        <td align="right"><input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="I1" type="image"> </td>
      </tr>
    </table>
    <input type="Hidden" name="vForm" value="Y">
  </form>
  <!---Edit List-->
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" height="26">
    <tr>
      <th align="left" height="13">Article No</th>
      <th align="left" height="13">Title</th>
    </tr>
    <%
    '...read Arts
    sOpenDb
    vSql = "Select * FROM Arts "
    Set oRs = oDb.Execute(vSQL)    
    Do While Not oRs.Eof 
      sReadArts
    %>
    <tr>
      <td height="12"><a <%=fStatX%> href="ArtsEdit.asp?vEditArtsNo=<%=vArts_No%>&vForm=n"><%=vArts_No%></a>&nbsp; </td>
      <td height="12"><a <%=fStatX%> href="javascript:articles('<%=vArts_No%>')"><%=vArts_Title%></a>&nbsp; </td>
    </tr>
    <%  
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  %>
  </table>
  <%
  Else
    If Len(Request.Form("vAddArtsNo")) = 0 And Len(Request.Form("vForm")) > 0 Then 
      vArts_No = 0
      vFunction = "add"
    ElseIf Len(Request.QueryString("vEditArtsNo")) > 0 Then 
      vArts_No = Request.QueryString("vEditArtsNo")
      vFunction = "edit"
    Else
       Response.Redirect "ArtsEdit.asp"          
    End If

    '...get the values (even if trying to add)
    If vArts_No <> 0 Then sGetArts
    
%>
  <form method="POST" action="ArtsEdit.asp" target="_self">
    <input type="Hidden" name="vFunction" value="<%=vFunction%>"><input type="Hidden" name="vArts_No" value="<%=vArts_No%>"><p>&nbsp;</p>
    <table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" valign="top" colspan="2">
        <h2 align="left">Articles</h2>
        <p align="left">Enter all values for the article and &quot;paste&quot; in the actual article (in clean HTML).&nbsp; Click update when finished.</p><p align="left">NOTE THIS NEEDS TO BE MODIFIED TO HANDLE NEW NUMBERING SYSTEM&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;&gt;<br>&nbsp;</p></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Article No : </th>
        <td width="70%" valign="top">&nbsp;<%=vArts_No%></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Type : </th>
        <td width="70%"><input type="radio" value="T" name="vArts_Type" <%=fcheck("t", varts_type)%>>Title (only add title - enter in Caps)<br><input type="radio" value="A" name="vArts_Type" <%=fcheck("a", varts_type)%>>Article (currently only add Title and Cut/Paste HTML Article)</td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Title : </th>
        <td width="70%">&nbsp;<input type="text" name="vArts_Title" size="52" value="<%=vArts_Title%>"></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Keywords : </th>
        <td width="70%">&nbsp;<textarea rows="2" name="vArts_Keywords" cols="44"><%=vArts_Keywords%></textarea></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Description : </th>
        <td width="70%">&nbsp;<input type="text" name="vArts_Desc" size="52" value="<%=vArts_Desc%>"></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Author : </th>
        <td width="70%">&nbsp;<input type="text" name="vArts_Author" size="52" value="<%=vArts_Author%>"></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Article : </th>
        <td width="70%">&nbsp;<textarea rows="8" name="vArts_Article" cols="44"><%=vArts_Article%></textarea></td>
      </tr>
      <tr>
        <td align="center" width="100%" valign="top" colspan="2" height="38"><br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="javascript:jconfirm('ArtsEdit.asp?vDelArtsNo=<%=vArts_No%>&vFunction=del', '<!--{{-->Ok to delete?<!--}}-->')"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a> <p align="center"><a <%=fStatX%> href="ArtsEdit.asp">Article List</a></p><p align="center"><a <%=fStatX%> href="Menu.asp"><img border="0" src="../Images/Icons/Administration.gif" alt="Click here for the Menu"></a></p></td>
      </tr>
    </table>
  </form>
  <%
    End If  
  %> 

	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


