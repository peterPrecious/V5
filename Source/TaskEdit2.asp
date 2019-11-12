<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_Keys.asp"-->

<% 
  Dim vLockedNote, vInactive, vSpaces, vSpace, vTitle, vOrderNo, vAction, vNextOrder, vLevel, vMoveCnt, vMoveCntSave, vTskHNo

  sTskH_ClearSession   '...clear sessions
  vInactive = 1        '...assume inactive

  If Len(Request("vTskH_AcctId")) > 0 Then vTskH_AcctId = Request("vTskH_AcctId")
  If Len(Request("vTskH_Id"))     > 0 Then vTskH_Id     = Request("vTskH_Id")
  If Len(Request("vInactive"))    > 0 Then vInactive    = Request("vInactive")
  
  
  If Len(vTskH_AcctId) = 0 Then vTskH_AcctId = svCustAcctId
  
  '...get delete values from form
  If Request.Form("vForm").Count = 1 Then

    '...see if changing the number of moves or deleting a line
    vMoveCnt = Request("vMoveCnt")
    If Len(vMoveCnt) = 0 Then
      vMoveCnt = Request("vMoveCntSave")
    End If

    '...process the delete or up/down buttons
    For Each vFld In Request.Form
      If Left(vFld, 5) = "vDel_" Then
        vTskH_No = Cint(Mid(vFld, 6))
        sDeleteTskH
      ElseIf Left(vFld, 4) = "vUp_" And Right(vFld, 2) = ".x" Then  
        vTskHNo = Cint(Mid(vFld, 5, Len(vFld)-6))
        For k = 1 to Cint(vMoveCnt)
          sShiftOrder vTskH_AcctId, vTskH_Id, vTskHNo, "up"
        Next
      ElseIf Left(vFld, 6) = "vDown_" And Right(vFld, 2) = ".x" Then  
        vTskHNo = Cint(Mid(vFld, 7, Len(vFld)-8))
        For k = 1 to Cint(vMoveCnt)
          sShiftOrder vTskH_AcctId, vTskH_Id, vTskHNo, "down"
        Next
      End If
    Next
   
   
  '...else get values from query string
  Else
    If Len(Request("vTskH_No"))     > 0 Then vTskH_No     = Cint(Request("vTskH_No"))
    If Len(Request("vAction"))      > 0 Then vAction      = Request("vAction")
    If Len(Request("vLevel"))       > 0 Then vLevel       = Cint(Request("vLevel"))
    If Len(Request("vNextOrder"))   > 0 Then vNextOrder   = Cint(Request("vNextOrder"))
  End If

  '...actions
  Select Case vAction
    Case "unlock"     : sUnlockTskH vTskH_No
    Case "lock"       : sLockTskH vTskH_No
    Case "reset"      : sLock vTskH_No, svMembNo
    Case "delete"     : sDeleteTskH
    Case "add"        : sAddTskH vTskH_Id, vTskH_No, vNextOrder, vLevel
    Case "clone"      : sCloneTskHbyNo vTskH_Id, vTskH_No
    Case "up", "down" : sShiftOrder vTskH_AcctId, vTskH_Id, vTskH_No, vAction  '...shift the sort order
  End Select

  '...flag any child records if task list order has changed
  sFlagTaskChildren vTskH_AcctId, vTskH_Id

  vSpaces = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"  
%>

<html>

<head>
  <title>TaskEdit2</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script> 
    function submitForm(theForm)
    {
      theForm.action = 'TaskEdit2.asp';
      theForm.submit();
    }
  </script>

</head>

<body>

  <% Server.Execute vShellHi %>
  <form method="POST" action="TaskEdit2.asp">

    <input type="hidden" name="vForm" value="y">
    <input type="hidden" name="vTskH_AcctId" value="<%=vTskH_AcctId%>">
    <input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>">
    <input type="hidden" name="vMoveCntSave" value="<%=fDefault(vMoveCntSave, 1)%>">

    <table class="table">
      <tr>
        <td colspan="7">
          <h1>Task List</h1>
          <p class="c2">This displays all the available items within the selected task.&nbsp; Click on the Items Title to edit the task, or on <b>Clone</b> to create a duplicate item after the current item - note that cloning does not copy any attached digital assets - these must be added in each time.&nbsp; Click on the <b>Assets</b> to add/edit any digital assets (content).&nbsp; Click <b>Delete Item</b> to remove the selected item (note that if you wish to delete a group of task items, delete the &quot;lower level&quot; items before you delete the &quot;upper level&quot; task.&nbsp; To delete a group of items, tick the items to be deleted then click the <b>Delete</b> button at the top of the list.&nbsp; NOTE THIS IS AN IRREVERSIBLE MOVE.&nbsp; This is because deleting a group &quot;header&quot; will NOT automatically delete the lower level items.)&nbsp; Clicking <b>Lock</b> makes the item unavailable on My Learning unless the item is programmatically <b>Unlock</b>ed.&nbsp; <b>Unlock</b> makes locked items available.&nbsp; <b>Reset</b> (if displayed) removes any keys this learner may have acquired thus relocking the task - use carefully.&nbsp; </p>

          <table style="width:400px; margin:20px auto 20px auto;">
            <tr>
              <td class="c2">Include Inactive Items :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
              <td class="c2"><input type="radio" value="1" name="vInactive" <%=fcheck(vinactive, 1)%> checked>Yes&nbsp;&nbsp; <input type="radio" value="0" name="vInactive" <%=fcheck(vinactive, 0)%>>No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
              <td class="c2"><input src="../Images/Buttons/Go_<%=svLang%>.gif" type="image" name="bGo"></td>
            </tr>
          </table>

        </td>
      </tr>

      <tr>
        <td class="rowshade">Items<br>(do not embed HTML tags)</td>
        <td class="rowshade">Order<br><span style="font-weight: 400">Move items up/down<br></span>
          <select size="1" name="vMoveCnt" onchange="javascript:submitForm(this.form)">
            <% For j = 1 To 99 %>
            <option <%=fselect(j, vmovecnt)%> value="<%=j%>"><%=j%></option>
            <% Next %>
          </select>
          <span style="font-weight: 400"><br>line(s)/click</span>
        </td>
        <td class="rowshade">Lang</td>
        <td class="rowshade" colspan="3">Action</td>
        <td class="rowshade"><font color="#FF0000">Carefully<br><input border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif" name="I1" type="image"> <br>all selected task items</font></td>
      </tr>

      <%
        vNextOrder = 0
    
        '...read Task
        sGetTskH_rs svCustAcctId, vTskH_Id
    
        Do While Not oRs.Eof 
          sReadTskH
    
          If vTskH_Order >= vNextOrder Then
            vNextOrder = vTskH_Order + 1
          End If
    
          Select Case vTskH_Level
            Case 0
              vSpace = ""
              vTitle = "<b>" & fIf(Len(Trim(vTskH_Title)) = 0 Or (Left(Trim(vTskH_Title), 1) = "<" And Right(Trim(vTskH_Title), 1) = ">"), "[No Title]", fLeft(vTskH_Title, 48)) & "</b>"
            Case 1
              vSpace = Left(vSpaces, 24)
              vTitle = "<b>" & fIf(Len(Trim(vTskH_Title)) = 0 Or (Left(Trim(vTskH_Title), 1) = "<" And Right(Trim(vTskH_Title), 1) = ">"), "[No Title]", fLeft(vTskH_Title, 48)) & "</b>"
            Case 2 
              vSpace = Left(vSpaces, 48)
              vTitle = fIf(Len(Trim(vTskH_Title)) = 0 Or (Left(Trim(vTskH_Title), 1) = "<" And Right(Trim(vTskH_Title), 1) = ">"), "[No Title]", fLeft(vTskH_Title, 48))
          End Select
    
          If vTskH_Level > 0 Then 
          
            If vTskH_Locked Then 
              vLockedNote = "  (Locked)  "
            Else
              vLockedNote = ""
            End If   
            
      %>
      <tr>
        <td>
          <%=vSpace%><a href="TaskEdit3.asp?vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>"><%=vTitle%></a>
          <%="&nbsp;&nbsp;&nbsp;" & vLockedNote%> 
        </td>
        <td style="text-align:center;">
          <!-- Shift Order 

          <a href="TaskEdit2.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vAction=up&vInactive=<%=vInactive%>"><img border="0" src="../Images/Icons/ArrowUp.gif"></a>&nbsp; 
          <a href="TaskEdit2.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vAction=down&vInactive=<%=vInactive%>"><img border="0" src="../Images/Icons/ArrowDown.gif"></a> | 
          -->

          <input border="0" src="../Images/Icons/ArrowUp.gif" name="vUp_<%=vTskH_No%>" width="18" height="22" type="image">
          <%=fDefault(vMoveCnt, 1)%>
          <input border="0" src="../Images/Icons/ArrowDown.gif" name="vDown_<%=vTskH_No%>" width="18" height="22" type="image">
        </td>


        <td style="text-align:center;"><%=vTskH_Lang%> </td>
        <td style="text-align:center;" class="c2">
          <% If vTskH_Level = 1 Then %>
          <% If vTskH_Locked Then %>
          <a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vInactive=<%=vInactive%>&vAction=unlock">Unlock</a>&nbsp; 
            <% Else %>
          <a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vInactive=<%=vInactive%>&vAction=lock">Lock</a>&nbsp; 
            <% End If %>

          <% If Not fIsLocked (vTskH_No, svMembNo) Then %>
          <a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vInactive=<%=vInactive%>&vAction=reset">Reset</a>
          <% Else %>
              &nbsp;&nbsp;&nbsp; 
            <% End If %>



          <% Else %>&nbsp;&nbsp;&nbsp; 
        <% End If %> </td>
        <td style="text-align:center;" class="c2"><% If vTskH_Level > 1 Then %> <a href="TaskEdit4.asp?vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>">Assets</a> <% Else %>&nbsp;&nbsp;&nbsp; <% End If %> </td>
        <td style="text-align:center;" class="c2"><a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vInactive=<%=vInactive%>&vAction=clone">Clone</a> </td>
        <td style="text-align:center;" class="c2"><a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vInactive=<%=vInactive%>&vAction=delete">Delete Item</a> or select
          <input type="checkbox" value="0" name="vDel_<%=vTskH_No%>"></td>
      </tr>
      <% Else %>
      <tr>
        <td colspan="7" height="30"><%=vSpace%><a href="TaskEdit3.asp?vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>"><%=fIf(Len(Trim(vTitle)) < 8 , "[No Title]", vTitle)%></a>&nbsp; </td>
      </tr>
      <% 
          End If
          oRs.MoveNext
        Loop
        Set oRs = Nothing
        sCloseDb    
      %>
      <tr>
        <td colspan="7" style="text-align:center;">&nbsp;<h2><a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vNextOrder=<%=vNextOrder%>&vInactive=<%=vInactive%>&vAction=add&vLevel=1">Add Major Item</a><%=f10%><a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskH_No%>&vNextOrder=<%=vNextOrder%>&vInactive=<%=vInactive%>&vAction=add&vLevel=2">Add Minor Item</a></h2>
          <input onclick="location.href = 'javascript:history.back(1)'" type="button" value="Return" name="bReturn" id="bReturn" class="button070"><h2><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>">My Learning</a><%=f10%><a href="TaskEdit1.asp">Task Library</a></h2>
        </td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
