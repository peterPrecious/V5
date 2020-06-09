<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_TskD.asp"-->

<% 
  Dim vAction, vInactiveNote, vInactive, vTemplate, vUrl, vCollapse
  sTskH_ClearSession   '...clear sessions

  vTskH_AcctId = Request("vTskH_AcctId")
  vTskH_Id     = Request("vTskH_Id") 
  vAction      = Request("vAction")
  vInactive    = Request("vInactive") : If fNoValue(vInactive)   Then vInactive = 0
  vTemplate    = Request("vTemplate") : If fNoValue(vTemplate)   Then vTemplate = 1
  
  Select Case vAction
    Case "collapse"   : vTskH_Collapse = 0 : sCollapseTskH vTskH_AcctId, vTskH_Id, vTskH_Collapse
    Case "expand"     : vTskH_Collapse = 1 : sCollapseTskH vTskH_AcctId, vTskH_Id, vTskH_Collapse
    Case "inactivate" : sInActivateTskH vTskH_AcctId, vTskH_Id
    Case "activate"   : sActivateTskH   vTskH_AcctId, vTskH_Id
    Case "clone"      : sCloneTskH      vTskH_AcctId, vTskH_Id
    Case "template"   : sTemplateTskH   vTskH_AcctId, vTskH_Id
    Case "delete"     : sDeleteTskH_rs  vTskH_AcctId, vTskH_Id
    Case "shift"
      If Request("vId1") <> 0 And Request("vId2") <> 0 Then 
        sShiftTask vTskH_AcctId, Request("vId1"), Request("vId2")
      End If
  End Select
%>

<html>

<head>
  <title>TaskEdit1</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!---Task List-->
  <table class="table">
    <tr>
      <td colspan="3"><h1 align="center">Task Library</h1><div class="c3">This displays all the available tasks (active unless specified).&nbsp; Click on the Title to review/edit the task, or on <b>Add</b> to create a new (empty) task list.&nbsp; You can optionally view system templates - listed in green - which can be <b>Delete</b>d, <b>View</b>ed and/or <b>Clone</b>d then modified.&nbsp; Note: templates cannot be edited directly. To create a new task start by cloning the <b>Empty Task</b> template (ensure you have selected <b>Include Templates for Cloning</b>).&nbsp; Decide if you want to <b>Collapse</b>/<b>Expand</b> the initial task list when learners enter My Learning (default = <b>Collapse</b>).&nbsp; You can <b>create a template</b> from a task, but remember it becomes a public template - ie access to templates is NOT restricted to your account!</div>
      <form method="POST" action="TaskEdit1.asp">
        <div align="center">
          <center>
          <table class="table">
            <tr>
              <td class="c2"><input type="radio" value="1" name="vInactive" <%=fcheck(vinactive, 1)%>>Yes&nbsp;&nbsp; <input type="radio" value="0" name="vInactive" <%=fcheck(vinactive, 0)%>>No&nbsp;&nbsp;&nbsp;&nbsp; </td>
              <td><p class="c2">Include Inactive Items</td>
            </tr>
            <tr>
              <td class="c2"><input type="radio" value="1" name="vTemplate" <%=fcheck(vtemplate, 1)%>>Yes&nbsp;&nbsp; <input type="radio" value="0" name="vTemplate" <%=fcheck(vtemplate, 0)%>>No&nbsp;&nbsp; </td>
              <td><font color="#008000">Include Templates for Cloning</font></td>
            </tr>
            <tr>
              <td colspan="2" align="center" height="20"><p><input border="0" src="../Images/Buttons/Go_<%=svLang%>.gif" type="image" name="I1"></p>
              </td>
            </tr>
          </table>
          </center></div>

      </form>
      </td>
    </tr>
    <tr>
      <th align="left" nowrap bgcolor="#DDEEF9" height="20">Tasks</th>
      <th nowrap bgcolor="#DDEEF9" height="20">Order</th>
      <th nowrap bgcolor="#DDEEF9" height="20">Action</th>
    </tr>
    <%
    '...read all Task Ids then display level 0s
    '   all account templates are 9999, use 9990 to include templates
    If vTemplate = 0 Then
      sGetTskH_rs svCustAcctId, 9999
    Else
      sGetTskH_rs svCustAcctId, 9990
    End If
    
    Dim vIdPrev
    vIdPrev = 0
    
    Do While Not oRs.Eof 
      sReadTskH
      If vTskH_Level = 0 Then
        If Not (vInactive = 0 And Not vTskH_Active) Then 
          If vTskH_Active Then 
            vInactiveNote = ""
          Else
            vInactiveNote = "  (Inactive)  "
          End If        
          If vTskH_AcctId = "0000" Then 
    %> <tr>
      <td align="left"><font color="#008000"><%=fLeft(vTskH_Title, 64)%></font>&nbsp; </td>
      <td align="center">&nbsp;</td>
      <td align="center"><p class="c2"><a target="_self" href="TaskView.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>">View</a>&nbsp; <a href="TaskEdit1.asp?vTskH_AcctID=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=clone&vInactive=<%=vInactive%>">Clone</a>&nbsp; <a href="TaskEdit1.asp?vTskH_AcctID=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=delete&vInactive=<%=vInactive%>"><font color="#FF0000">Delete - Caution</font></a></td>
    </tr>
    <%    Else %> <tr>
      <td align="left"><a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>"><%=fLeft(vTskH_Title, 64)%></a><%="&nbsp;&nbsp;&nbsp;" & vInactiveNote%>&nbsp; </td>
      <td align="center">
      <!-- Shift Ids --><a href="TaskEdit1.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vId1=<%=vTskH_Id%>&vAction=shift&vId2=<%=vIdPrev%>"><img border="0" src="../Images/Icons/ArrowUp.gif" align="top"></a>&nbsp; <a href="TaskEdit1.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vId1=<%=vTskH_Id%>&vAction=shift&vId2=<%=fNextInAcctTskH_Id (vTskH_AcctId, vTskH_Id)%>"><img border="0" src="../Images/Icons/ArrowDown.gif" align="bottom"></a> </td>
      <td align="center"><p class="c2"><% If vTskH_Collapse Then %> <a href="TaskEdit1.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=collapse">Collapse</a>&nbsp; <% Else %> <a href="TaskEdit1.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=expand">Expand</a>&nbsp; <% End If %> <% If vTskH_Active Then %> <a href="TaskEdit1.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=inactivate&vInactive=<%=vInactive%>">Inactivate</a>&nbsp; <% Else %> <a href="TaskEdit1.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=activate&vInactive=<%=vInactive%>">Activate</a>&nbsp; <% End If %> <a target="_self" href="TaskView.asp?vTskH_AcctID=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vInactive=<%=vInactive%>">View</a>&nbsp; <a href="TaskEdit1.asp?vTskH_AcctID=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=clone&vInactive=<%=vInactive%>">Clone</a>&nbsp;
      <a href="TaskEdit1.asp?vTskH_AcctID=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=delete&vInactive=<%=vInactive%>"><font color="#FF0000">Delete - Caution</font></a>&nbsp; <a href="TaskEdit1.asp?vTskH_AcctID=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vAction=template&vInactive=<%=vInactive%>"><font color="#008000">Create Template</font></a> </td>
    </tr>
    <%  
          vIdPrev = vTskH_Id
          End If
        End If
      End If
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
  %> <tr>
      <td style="text-align:center; padding:30px;" colspan="3">
        <input onclick="location.href='javascript:history.back(1)'" type="button" value="Return" name="bReturn" id="bReturn"class="button070"></p> 
        <h2><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>">My Learning</a></h2>
      </td>
    </tr>
  </table>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>