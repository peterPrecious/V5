<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->

<%
   Response.Buffer = False

   Dim vCritTitle1, vCritTitle2, vCritTitle3, vCritTitle4, vCritTitle5, aCritTitles

   sExtractCrit   

   If Request.Form("bAdd").Count > 0 Then
     sInsertCrit (svCustAcctId)
   ElseIf Request.Form("bUpd").Count > 0 Then
     sUpdateCrit
   ElseIf Request.Form("bDel").Count > 0 Then
     sDeleteCrit
   ElseIf Request.Form("bTit").Count > 0 Then

     '...extract titles, must exist from left to right (remove ' and ~)
     If Len(Request("vCritTitle1")) > 0 Then
	     vCritTitle1 = Replace(Replace(Request("vCritTitle1"), "~", ""), "'", "")
       vCust_CritTitles = vCritTitle1

       If Len(Request("vCritTitle2")) > 0 Then
         vCritTitle2 = Replace(Replace(Request("vCritTitle2"), "~", ""), "'", "")
         vCust_CritTitles = vCust_CritTitles & "~" & vCritTitle2

         If Len(Request("vCritTitle3")) > 0 Then
           vCritTitle3 = Replace(Replace(Request("vCritTitle3"), "~", ""), "'", "")
           vCust_CritTitles = vCust_CritTitles & "~" & vCritTitle3

           If Len(Request("vCritTitle4")) > 0 Then
             vCritTitle4 = Replace(Replace(Request("vCritTitle4"), "~", ""), "'", "")
             vCust_CritTitles = vCust_CritTitles & "~" & vCritTitle4

             If Len(Request("vCritTitle5")) > 0 Then
               vCritTitle5 = Replace(Replace(Request("vCritTitle5"), "~", ""), "'", "")
               vCust_CritTitles = vCust_CritTitles & "~" & vCritTitle5
             End If  
           End If  
         End If  
       End If  
	     sUpdateCustCritTitles (svCustId)
     End If  
       
   Else

     sGetCust (svCustId)
     If Len(vCust_CritTitles) > 0 Then
       aCritTitles = Split(vCust_CritTitles, "~")
       If Ubound(aCritTitles)  = 4 Then vCritTitle5 = aCritTitles(4)
       If Ubound(aCritTitles) >= 3 Then vCritTitle4 = aCritTitles(3)
       If Ubound(aCritTitles) >= 2 Then vCritTitle3 = aCritTitles(2)
       If Ubound(aCritTitles) >= 1 Then vCritTitle2 = aCritTitles(1)
                                        vCritTitle1 = aCritTitles(0)
     End If
   End If
     
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %>

  <table width="100%" border="1" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td colspan="8"><h1 align="center">Group Table</h1><h2>This table allows you to organize learners into groups for reporting or for training by assigning one or more <b>Job IDs</b> to each Group.&nbsp; Enter the new Group Id and optional Job Ids then click <b>Add</b>.&nbsp; All entered values will be displayed beneath in Group order.&nbsp; You can edit any values in any row by clicking <b>Update</b>, or delete any row using <b>Delete</b>.&nbsp; Note: work on one row at a time and use the appropriate Action button.</h2>
<!--
      <h6 align="center">Note: Multiple Job IDs are not yet active - leave at Select or only select ONE.</h6>
-->      
      </td>
    </tr>
    <tr>
      <th nowrap bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">No</th>
      <th nowrap bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Group Id</th>
      <th nowrap bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Job Ids</th>
      <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" nowrap>Action</th>
    </tr>
    <tr>
      <form method="POST" action="CritEdit.asp" name="Crit">
        <td align="center" valign="top">New...</td>
        <td align="center" valign="top"><input type="text" size="16" name="vCrit_Id"></td>
        <td align="center" valign="top">
          <select size="4" name="vCrit_JobsId" multiple>
          <option>Select</option>
          <%=fJobsOptions ("")%>
          </select> 
        </td>
        <td align="center" nowrap valign="top"><input type="submit" value="Add" name="bAdd" class="button070"></td>
      </form>
    </tr>
    <%
      sGetCrit_rs svCustAcctId
      If Not oRs2.Eof Then
	  %>
    <tr>
      <td colspan="8" align="center" valign="top">&nbsp;</td>
    </tr>
    <%
  		End If    
      Do While Not oRs2.Eof 
        sReadCrit
	  %>
    <tr>
      <form method="POST" action="CritEdit.asp" name="Crit_<%=vCrit_No%>">
        <td align="center" valign="top"><%=vCrit_No%></td>
        <td align="center" valign="top"><input type="text" size="16" name="vCrit_Id" value="<%=vCrit_Id%>"></td>
        <td align="center" valign="top"><select size="4" name="vCrit_JobsId" multiple>
        <option>Select</option>
        <%=fJobsOptions (vCrit_JobsId)%></select></td>
        <td align="center" nowrap valign="top"><input type="submit" value="Update" name="bUpd" class="button070"> <input type="submit" value="Delete" name="bDel" class="button070"> </td>
        <input type="hidden" name="vCrit_No" value="<%=vCrit_No%>">
      </form>
    </tr>
    <%
        oRs2.MoveNext
      Loop
      Set oRs2 = Nothing
      sCloseDb2    
	  %>
    <tr>
      <td colspan="8" align="center"><h2><br><a <%=fstatx%> href="JobsEdit.asp">Jobs Table</a>&nbsp; |&nbsp; <a <%=fstatx%> href="SkilEdit.asp">Skills Table</a><br></h2></td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


