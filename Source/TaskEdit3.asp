<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_TskD.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Password_Routines.asp"-->

<%  
  sTskH_ClearSession   '...clear sessions

  '...from TaskEdit3.asp
  If Request.QueryString("vTskH_Id").Count > 0 Then
    vTskH_Id = Request("vTskH_Id")
    vTskH_No = Request("vTskH_No")
    sGetTskH svCustAcctId, vTskH_No      

    '...password?
    If Not fNoValue(vTskH_Password) Then
      If Session("MyWorld_PasswordEntered") <> vTskH_Password Then
        Session("MyWorld_Password") = vTskH_Password
        Session("MyWorld_Url") = "TaskEdit3.asp?vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
        Response.Redirect "TaskPassword.asp?" & Session("MyWorld_Url")
      End If
    End If

  Else
    sExtractTskH  
  End If
 
  If Request("bUpdate").Count > 0 Then
    sUpdateTskH 
    Response.Redirect "TaskEdit2.asp?vTskH_Id=" & vTskH_Id
  End If

  Function fCheckMark(i)
    If i = True Then
      fCheckMark = "<img border='0' src='../Images/Icons/Check.gif'>"
    Else  
      fCheckMark = "<font face='Verdana' size='1'>&nbsp;</font>" 
    End If
  End Function

%>

<html>
<head>
  <title>TaskEdit3</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <div>
    <h1>Task Items</h1>
    <p class="c2">This allows you to edit/add a title and description to the item and filter access using the <b>Group</b>, <b>Dates</b> and <b>Levels</b> fields.&nbsp; You can add advanced <b>Services</b> (if available).&nbsp; Most importantly, this displays the digital assets (content list) that are attached to this item.&nbsp; Click <b>Edit Content List</b> to add or remove digital assets.</p>
    <br /><br />
  </div>


  <form method="POST" action="TaskEdit3.asp" target="_self">
    <input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>">
    <input type="hidden" name="vTskH_No" value="<%=vTskH_No%>">
    <input type="hidden" name="vTskH_Level" value="<%=vTskH_Level%>">
    <input type="hidden" name="vTskH_Order" value="<%=vTskH_Order%>">
    <input type="hidden" name="vTskH_Child" value="<%=fSqlBoolean(vTskH_Child)%>">
    <table class="table">
      <tr>
        <th>Title :</th>
        <td><input type="text" size="62" name="vTskH_Title" value="<%=vTskH_Title%>"><% If Len(Trim(vTskH_Title)) = 0 Then Response.Write "[No Title]" %>
        </td>
      </tr>
      <tr>
        <th>Description :</th>
        <td><textarea rows="6" name="vTskH_Desc" cols="80"><%=vTskH_Desc%></textarea></td>
      </tr>
      <% If vTskH_Level = 0 Then %>
      <tr>
        <th>Password :</th>
        <td><input type="password" name="vTskH_Password" size="13" value="<%=vTskH_Password%>" maxlength="16"><br>Passwords can only be established at this level.&nbsp; Leave empty if not required. Do not use spaces or single quote marks.&nbsp; Try not to forget password as it is hard to recover.</td>
      </tr>
      <% End If %>
      <tr>
        <th>Language :</th>
        <td>
          <input type="radio" value="XX" name="vTskH_Lang" <%=fcheck("xx", vtskh_lang)%>>XX - always display (default) <br>
          <input type="radio" value="EN" name="vTskH_Lang" <%=fcheck("en", vtskh_lang)%>>EN<br>
          <input type="radio" value="FR" name="vTskH_Lang" <%=fcheck("fr", vtskh_lang)%>>FR<br>
          <input type="radio" value="ES" name="vTskH_Lang" <%=fcheck("es", vtskh_lang)%>>ES
        </td>
      </tr>

      <!-- Only Level 1/2 offers Selection Criteria  -->
      <% If vTskH_Level <= 2 Then %>
      <tr>
        <th>Active :</th>
        <td><input type="checkbox" name="vTskH_Active" value="1" <%=fcheck(fsqlboolean(vtskh_active), 1)%>></td>
      </tr>
      <%   i = fCriteriaList (svCustAcctId, "TskH") %>
      <tr>
        <th>Access Group 1 :</th>
        <td>
          <select size="<%=vCriteriaListCnt+1 %>" name="vTskH_Criteria" multiple><%=i%></select><br>Select &quot;All&quot; or restrict learner to one or more Groups, subject to the Dates and Level restrictions below.&nbsp; Use Ctrl+Click to select multiple Groups.&nbsp; Group values are established in and extracted from the <a target="_blank" href="CritEdit.asp">Group Table</a>.&nbsp; </td>
      </tr>

      <tr>
        <th>Access Group 2 :</th>
        <td>
          <select size="1" name="vTskH_Group2">
            <% 
            For i = 0 to 24 
              Response.Write "<option " & fSelect(i, vTskH_Group2) & " value='" & i & "'>" & i & "</option>"
            Next 
            %>
          </select>
          <br>If assigned to learners, values can be between 1 and 16.
        </td>
      </tr>

      <tr>
        <th>Access Ids :</th>
        <td>
          <textarea rows="6" name="vTskH_AccessIds" cols="80"><%=vTskH_AccessIds%></textarea>
          <br>Enter Learner Ids (NOT Passwords) who can access this task, separated by spaces, ie &quot;pbulloch s_spade&quot;.&nbsp; If any values are entered, then ONLY these learners can access this level (other than Administrators) regardless how you set the Access Level (below).</td>
      </tr>

      <tr>
        <th>Access Dates :</th>
        <td>
          <input type="text" name="vTskH_DateStart" size="16" value="<%=fFormatDate(vTskH_DateStart)%>">Between Start date (ie Jan 15, 2003)<br>
          <input type="text" name="vTskH_DateEnd" size="16" value="<%=fFormatSqlDate(vTskH_DateEnd)%>">and End date (ie Dec 31, 2003)
        </td>
      </tr>
      <tr>
        <th>Access Level :</th>
        <td>
          <input type="radio" value="1" name="vTskH_AccessLevel" <%=fcheck(1, vtskh_accesslevel)%>>Learners only <%=fIf(svMembLevel = 5, " (and Administrators)", "") %><br>
          <input type="radio" value="2" name="vTskH_AccessLevel" <%=fcheck(2, vtskh_accesslevel)%>>Learners or higher<br>
          <input type="radio" value="3" name="vTskH_AccessLevel" <%=fcheck(3, vtskh_accesslevel)%>>Facilitators or higher<br>
          <input type="radio" value="4" name="vTskH_AccessLevel" <%=fcheck(4, vtskh_accesslevel)%>>Managers or higher<br>
          <input type="radio" value="5" name="vTskH_AccessLevel" <%=fcheck(5, vtskh_accesslevel)%>>Administrators only</td>
      </tr>

      <% Else %>

      <input type="hidden" name="vTskH_Criteria" value="<%=vTskH_Criteria%>">
      <input type="hidden" name="vTskH_DateStart" value="<%=vTskH_DateStart%>">
      <input type="hidden" name="vTskH_DateEnd" value="<%=vTskH_DateEnd%>">
      <input type="hidden" name="vTskH_AccessLevel" value="<%=vTskH_AccessLevel%>">

      <% End If %>


      <% If svCustLevel = 4 Then %>

      <tr>
        <th>Access Accounts :</th>
        <td>
          <input type="text" name="vTskH_CustIds" size="62" value="<%=vTskH_CustIds%>" maxlength="2000"><br>Only use this filter if more than one Customer Id shares a common Account Id, in which case each customer shares the same &quot;My Learning&quot;.&nbsp; To restrict this task to a specific Customer Id, enter the appropriate 8 char Customer Id(s), separated by a space, ie: &quot;ABCD1234 ABCD2345&quot;. Note, like all access fields, you will need to add this filter on every task you wish restricted.</td>
      </tr>

      <%   If vTskH_Level = 0 Then %>
      <tr>
        <th>General Notification :<br><font color="#808000"><span style="font-weight: 400">(not active)</span></font></th>
        <td>
          <textarea rows="6" name="vTskH_NotifyAll" cols="53"></textarea><br>Enter in as many languages as are appropriate, a notification that all learners will see <b>EVERY TIME</b> they access My Learning.&nbsp; Leave empty if no General Notification is used.</td>
      </tr>
      <tr>
        <th>Custom Notification :<br><font color="#808000"><span style="font-weight: 400">(not active)</span></font></th>
        <td>
          <input type="radio" name="vTskH_Notify" value="0">No, do not generate a Custom Notification<br>
          <input type="radio" name="vTskH_Notify" value="1">Yes, generate a Custom Notification for learners based on the script &quot;Notification.asp&quot; stored in the client's Repository/Tools folder by Vu.</td>
      </tr>
      <%   End If %>


      <% End If %>

      <% If vTskH_Level = 2 Then %>
      <tr>
        <td align="center" colspan="2">
          <h2>Digital Assets</h2>
          <% 
'        Response.Write "<br>TskH_No: " & vTskH_No
         sGetTskD_rs vTskH_No      
         If Not oRs2.Eof Then
          %>
          <table style="width:600px; margin:30px auto 30px auto">
            <tr>
              <td class="rowshade">Order</td>
              <td class="rowshade">Type</td>
              <td class="rowshade" style="text-align:left">Id</td>
              <td class="rowshade" style="text-align:left">Title</td>
              <td class="rowshade">Active</td>
            </tr>
            <%   
           Do While Not oRs2.Eof 
             sReadTskD
              
            '...display title for program and module
            If Len(vTskD_Title) = 0 Then
              If Left(vTskD_Type, 1) = "M" Then 
                vTskD_Title = fModsTitle(vTskD_Id)
              ElseIf Left(vTskD_Type, 1) = "P" Then 
                vTskD_Title = fProgTitle(vTskD_Id)
              End If
            End If
             
            %>
            <tr>
              <td style="text-align:center"><%=vTskD_Order%></td>
              <td style="text-align:center">&nbsp;<%=vTskD_Type%> </td>
              <td>&nbsp;<%=fLeft(vTskD_Id, 32)%> </td>
              <td>&nbsp;<%=fLeft(vTskD_Title, 32) %> </td>
              <td style="text-align:center">&nbsp;<%=fCheckMark(vTskD_Active)%> </td>
            </tr>
            <% 
             oRs2.MoveNext
           Loop
           Set oRs2 = Nothing
           sCloseDb2 
            %>
          </table>

          <% End If %>
          <p style="text-align:center"><a href="TaskEdit4.asp?vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>">Edit Digital Assets</a></p>
        </td>
      </tr>

      <% End If %>

      <tr>
        <td style="text-align:center; margin:30px;" colspan="2">
          <h2>&nbsp;</h2>
          <h2><a href="TaskEdit1.asp">Task Library</a><%=f10%>
            <a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>">Task List</a>
            <% If vTskH_Level = 3 Then %><%=f10%><a href="TaskEdit4.asp?vTskH_No=<%=vTskH_No%>&vTskH_Id=<%=vTskH_Id%>">Digital Assets</a><% End If %>
          </h2>
          <input onclick="location.href = 'javascript:history.back(1)'" type="button" value="Return" name="bReturn" id="bReturn" class="button070"><%=f10%>
          <input type="submit" value="Update" name="bUpdate" class="button">
          <h2><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>">My Learning</a></h2>
        </td>
      </tr>

    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>
</html>
