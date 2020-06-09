<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_TskD.asp"-->

<html>

<head>
  <title>TaskEdit4</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi
  
    Dim vOrderNo

    sTskH_ClearSession   '...clear sessions
    vTskH_Id = Request("vTskH_Id")  
    
    '...from TaskEdit3 or MyWorld
    If Request.QueryString("vTskH_Id").Count > 0 Then
      vTskD_No = Request("vTskH_No")  
    Else
      sExtractTskD  
    End If
  
    If Request("vAction") = "add" Then
      sInsertTskDEmpty
    End If
   
    If Request("bUpdate").Count > 0 Then
      sUpdateTskD 
    End If
  
    If Request("bDelete").Count > 0 Then
      sDeleteTskD 
    End If
  
    '...move up a record in the sort order
    If Len(Request("bUp.x")) > 0 Then
      sGetTskD_Rs vTskD_No
      i = 0
      Do While Not oRs2.Eof 
        sReadTskD
        i = i + 1
        If vOrderNo - 1 = i Then '...add 1 to current record order
          vTskD_Order = i + 1
          sUpdateTskDOrder
        ElseIf vOrderNo = i Then '...remove 1 from designated record order
          vTskD_Order = i - 1
          sUpdateTskDOrder
        End If
        oRs2.MoveNext
      Loop
    End If
    
    '...move up a record in the sort order
    If Len(Request("bDown.x")) > 0 Then
      sGetTskD_Rs vTskD_No
      i = 0
      Do While Not oRs2.Eof 
        sReadTskD
        i = i + 1
        If vOrderNo + 1 = i Then '...add 1 to current record order
          vTskD_Order = i - 1
          sUpdateTskDOrder
        ElseIf vOrderNo = i Then '...remove 1 from designated record order
          vTskD_Order = i + 1
          sUpdateTskDOrder
        End If
        oRs2.MoveNext
      Loop
    End If
  
  %>

  <div>
    <h1>Task Assets</h1>
    <p class="c2">This displays all the available content (task digital assets) within the selected task.&nbsp; Edit an asset then click <b>Update</b> to update it, or <b>Delete</b> to remove a specific asset. Click on <b>Add</b> (at the bottom) to add a new asset to the list.&nbsp; To change the order in which the assets are listed click on the order arrow(s).</p>
    <h2><a href="#" onclick="toggle('divNotes')">Instructions</a> </h2>
    <div id="divNotes" class="div" style="text-align: left; width: 600px; margin: auto;">
      <p>Asset Type:</p>
      <ul>
        <li>Links to Vubiz.com will replace whatever is in the current window - this is to avoid session conflicts.</li>
        <li>Documents (Private) are extracted from the Client&#39;s Repository folder (uploaded by the client).</li>
        <li>Documents (Client) are extracted from the Client&#39;s Repository/Tools folder (setup by Vu).</li>
        <li>Documents (Common) are extracted from the Common Repository/Tools folder&nbsp; (setup by Vu).</li>
      </ul>
      <p>Asset Id: Leave Asset Id empty for Ecommerce Programs, Modules (Members or Criteria) or Jobs from the Member table.&nbsp; Enter for all other selections.</p>
      <p>Title: Leave empty for Ecommerce Programs, Jobs, Programs or Module Assets where the Title will be extracted from the Module/Program table. If you enter a Module Id, it will display the Module Title.&nbsp; If you use Module Prerequisites, enter launch Module Id|Prereq ModId, ie: 1234EN|1233EN or P1234EN|1234EN|1233EN (note only enter Program Id once).&nbsp; You can turn On/Off enhanced functionality in the Vubuild player with coding like: P1234EN|1001EN|N|Y|Y (Test Button | Bookmark | Completion Button) - all default to N if none are entered, but enter all 3 if any one of these functions are required.</p>
      <p>Display: Digital assets are typically display In Screen display (ie in the same frame) but can be displayed in a 800x600 pop-up or full screen popup. Module display is defined in the Module Table. </p>
    </div>
  </div>


  <table class="table">
    <tr>
      <td class="rowshade" style="width:080px">Order</td>
      <td class="rowshade" style="width:600px" colspan="5">Task Details</td>
      <td class="rowshade">Active</td>
      <td class="rowshade">Action</td>
    </tr>
    <%   
       sGetTskD_rs vTskD_No
       vOrderNo = 0
       Do While Not oRs2.Eof 
         sReadTskD
         vOrderNo = vOrderNo + 1 
    %>
    <form method="POST" action="TaskEdit4.asp#Bottom" target="_self" name="f<%=vTskD_Key%>">
      <input type="hidden" name="vTskD_Key" value="<%=vTskD_Key%>">
      <input type="hidden" name="vTskD_No" value="<%=vTskD_No%>">
      <input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>"><input type="hidden" name="vTskD_Order" value="<%=vTskD_Order%>">
      <input type="hidden" name="vOrderNo" value="<%=vOrderNo%>">
      <tr>
        <td style="text-align:center;"><%=vTskD_Order%></td>
        <th style="width:100px">Asset Type : </th><td>
          <select size="1" name="vTskD_Type">
            <option value="X">Inactive</option>
            <option value="AM" <%=fselect("am", vtskd_type)%>>1 Accessible Module</option>
            <option value="PC" <%=fselect("pc", vtskd_type)%>>1 Program</option>
            <option value="M" <%=fselect("m",  vtskd_type)%>>1 Module (ie P1234EN|9876EN)</option>
            <option value="MT" <%=fselect("mt", vtskd_type)%>>1 Module with Test Prerequisite</option>
            <option value="MR" <%=fselect("mr", vtskd_type)%>>1 Module - Review</option>
            <option value="S" <%=fselect("s",  vtskd_type)%>>1 Module - Staging</option>
            <option value="PE" <%=fselect("pe", vtskd_type)%>>Ecommerce Programs</option>
            <option value="U" <%=fselect("u",  vtskd_type)%>>Programs (Member Table) Legacy</option>
            <option value="UC" <%=fselect("uc", vtskd_type)%>>Programs (Member Table) </option>
            <option value="O" <%=fselect("o",  vtskd_type)%>>Jobs (Member Table) Legacy</option>
            <option value="OC" <%=fselect("oc", vtskd_type)%>>Jobs (Member Table)</option>
            <option value="R" <%=fselect("r",  vtskd_type)%>>Jobs (Criteria/Job Table) Legacy</option>
            <option value="RC" <%=fselect("rc", vtskd_type)%>>Jobs (Criteria/Job Table)</option>
            <option value="J" <%=fselect("j",  vtskd_type)%>>Programs (Job Table) Legacy</option>
            <option value="JC" <%=fselect("jc", vtskd_type)%>>Programs (Job Table)</option>
            <option value="T" <%=fselect("t",  vtskd_type)%>>Test</option>
            <option value="10" <%=fselect("10", vtskd_type)%>>VuAssess</option>
            <option value="E" <%=fselect("e",  vtskd_type)%>>Exam</option>
            <option value="ET" <%=fselect("et", vtskd_type)%>>Exam with Test Prerequisite</option>
            <option value="V" <%=fselect("v",  vtskd_type)%>>Document (Private)</option>
            <option value="D" <%=fselect("d",  vtskd_type)%>>Document (Client)</option>
            <option value="C" <%=fselect("c",  vtskd_type)%>>Document (Common)</option>
            <option value="L" <%=fselect("l",  vtskd_type)%>>Link</option>
            <option value="Y" <%=fselect("y",  vtskd_type)%>>Title</option>
          </select>
          *
        </td>
        <th style="width:100px">Asset Id : </th><td style="white-space:nowrap"><input type="text" name="vTskD_Id" size="28" value="<%=vTskD_Id%>">**</td>
        <td rowspan="2"><input src="../Images/Icons/ArrowUp.gif" name="bUp" type="image"><br><input src="../Images/Icons/ArrowDown.gif" name="bDown" type="image"></td>
        <td rowspan="2" style="text-align:center"><input type="checkbox" name="vTskD_Active" value="1" <%=fcheck(1, fsqlboolean(vtskd_active))%>></td>
        <td rowspan="2">
          <input type="submit" value="Update" name="bUpdate" class="button070"><br>
          <input type="submit" value="Delete" name="bDelete" class="button070">
        </td>
      </tr>
      <%
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
        <th style="width:180px" colspan="2">Title : </th>
        <td><input type="text" name="vTskD_Title" size="37" value="<%=vTskD_Title%>"></td>
        <th style="width:100px">Display : </th>
        <td>
          <select size="1" name="vTskD_Window">
            <option value="0" <%=fselect(0, vtskd_window)%>>In Screen</option>
            <option value="1" <%=fselect(1, vtskd_window)%>>Popup (800x600)</option>
            <option value="9" <%=fselect(9, vtskd_window)%>>Popup FullScreen</option>
          </select>
          ***
        </td>
      </tr>
    </form>

    <tr>
      <td colspan="8"><hr></td>
    </tr>
    <% 
           oRs2.MoveNext
         Loop
         Set oRs2 = Nothing
         sCloseDb2 
    %>
  </table>


      <div align="center" valign="top" colspan="8">&nbsp;<h2><a name="Bottom">Action</a>:<%=f10%><a href="TaskEdit1.asp">Task Library</a><%=f10%><a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>">Task List</a><%=f10%><a href="TaskEdit3.asp?vTskH_Id=<%=vTskH_Id%>&vTskH_No=<%=vTskD_No%>">Edit Task Items</a></h2>
        <input onclick="location.href = 'javascript:history.back(1)'" type="button" value="Return" name="bReturn" id="bReturn" class="button070"><%=f10%><input onclick="  location.href = 'TaskEdit4.asp?vTskH_Id=<%=vTskH_Id%>&amp;vTskH_No=<%=vTskD_No%>&amp;vAction=add#Bottom'" type="button" value="Add" name="bAdd" class="button070">
        <h2><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>">My Learning</a></h2>
      </div>





  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  <style>

    xtd, xth {
      white-space:nowrap;
    }

  </style>


</body>

</html>


