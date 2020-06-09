<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  '...First time in?  
  If Session("Completion_InitParms") = "" Then 
    Session("Completion_InitParms") = "Y"
    Response.Redirect "Completion.asp?vNext=Completion_RoleManager.asp" 
  End If

  Dim vMsg, vRoles, vJobs, aRoles, aJobs, vJobsCnt, vAction

  '...display all available roles that can be reported upon
  If Request.Form.Count > 0 Then
 
    vRoles = Ucase(Replace(Request("vRoles"), ",", "")) 
    aRoles = Split(vRoles)
    vJobs  = Ucase(Replace(Request("vJobs"), ",", "")) 
    aJobs  = Split(vJobs)
    vAction = fIf(Request("bAdd").Count > 0 , "Add", "Del")

    sUpdateRoles

    If Ubound(aJobs) = 0 Then
      vMsg = "Jobs ID: " & vJobs & " was successfully "
    Else
      vMsg = "Jobs IDs: " & vJobs & " were successfully "
    End If        
      
    If vAction = "Add" Then
      vMsg = vMsg & "assigned to "
    Else  
      vMsg = vMsg & "removed from "
    End If

    If Ubound(aRoles) = 0 Then
      vMsg = vMsg & "Role: " & vRoles
    Else  
      vMsg = vMsg & "Roles: " & vRoles
    End If

  End If 


  '...format Roles for SQL ... 
  '   ie generate ('TM', 'CA') from TM CA
  Function fSqlRoles (vRoles)
    fSqlRoles = "('" & Replace (vRoles, " ", "', '") & "')"
  End Function


  '...get all Active Jobs
  Function fJobsAll
    fJobsAll= ""
    vJobsCnt = 0
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      If vJobs_Active Then 
        vJobsCnt = vJobsCnt + 1
        fJobsAll = fJobsAll & "<option value=" & Chr(34) & vJobs_Id & Chr(34) & ">" & vJobs_Id & " : " & vJobs_Title & fIf(vJobs_Active, "", " (Inactive)") & "</option>" & vbCrLf
      End If  
      oRs3.MoveNext
    Loop      
    sCloseDb3           
    Set oRs3 = Nothing
  End Function


  Sub sUpdateRoles
    Dim aCritJobs, bOk
    '...get the jobs in all selected roles
    vSql = ""_
         & "SELECT "_
         & "  Crit_No, Crit_JobsId "_
         & "FROM "_
         & "  Crit WITH (NOLOCK) "_
         & "WHERE "_
         & "  Crit_AcctId = '" & svCustAcctId & "' "_
         & "  AND "_
         & "  RIGHT (Crit_Id, " & Session("Completion_RLlen") & ") IN " & fSqlRoles (vRoles) 
    sCompletion_Debug
    sOpenDb
    sOpenDb2
    Set oRs = oDb.Execute(vSql)


    Do While Not oRs.Eof
      bOk           = False   '...set to true if we need to update this record
      vCrit_No      = oRs("Crit_No")
      vCrit_JobsId  = oRs("Crit_JobsId")
      aCritJobs     = Split(vCrit_JobsId)      
      If vAction = "Del" Then      
        '...see if any selected jobs are available for removal
        For i = 0 To Ubound(aJobs)
          If Instr(vCrit_JobsId, aJobs(i)) > 0 Then
            vCrit_JobsId = Replace(vCrit_JobsId, aJobs(i), "")
            vCrit_JobsId = Replace(vCrit_JobsId, "  ", " ")
            bOk = True
          End If
        Next
      Else
        '...see if any selected jobs are available, else add
        For i = 0 To Ubound(aJobs)
          If Instr(vCrit_JobsId, aJobs(i)) = 0 Then
            vCrit_JobsId = vCrit_JobsId & " " & aJobs(i)
            bOk = True
          End If
        Next    
      End If      
      If bOk Then 
        vSql = "UPDATE Crit SET Crit_JobsId = '" & Trim(vCrit_JobsId) & "' WHERE Crit_No = " & vCrit_No
        oDb2.Execute(vSql)
        sCompletion_Debug
      End If
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
  End Sub

%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>	

    // turn on/off any group element starting with "group" value, ie "vProg" or "vProgP1234"
    function checkOnOff(theElement, group) {
      var i, j, theForm = theElement.form;
      j = group.length;
      for (i = 0; i < theForm.length; i++) {
        if (theForm[i].type == "checkbox" && theForm[i].id.substring(0, j) == group) {
          theForm[i].checked = theElement.checked;
	      }
	    }
    }

 
   function Validate(theForm) {


      var message_01 = "<%=fPhraH(001256)%>"
      var message_02 = "<%=fPhraH(000645)%>"
      var isOk = false
     
      if (theForm.vJobs.length != undefined) {
        for (i=0; i < theForm.vJobs.length; i++) {
          if (theForm.vJobs[i].selected == true) {
            isOk = true;
          }       
        }
        if (isOk == false) {
          alert(message_01);
          theForm.vJobs(0).focus();
          return (false);
        }
      }


      var isOk = false;
      for (i=0; i < theForm.length; i++) {
        if (theForm[i].type == "checkbox" && theForm[i].id.length == 6) {
          if (theForm[i].checked == true) {
            isOk = true;
          }       
        }  
      }
      if (isOk == false) {
        alert(message_02);
        return (false);
      }

 

     return (true);
    }
  </script>
  <title>Completion</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <form method="POST" onsubmit="return Validate(this)" name="fRole" id="fRole" action="Completion_RoleManager.asp">
      <table border="1" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" id="table1" width="90%">
        <tr>
          <th colspan="2" class="c1" height="90" valign="top">
          <h1><br> <%=svCustTitle%> Role Manager</h1>
          <p class="c2">This utility will Assign or Remove the selected Job ID(s) to/from ALL selected Roles <br>regardless of what <%=Session("Completion_L0Tit")%> they are in.</p>
          <p><%=fIf(vMsg <> "", "<p class='c5'>" & vMsg & "</p>", "")%></p>
          </th>
        </tr>
        <tr>
          <td align="left" colspan="2" class="c1" height="60">
          Step 1 : Select the appropriate Job ID(s) from the Active Job ID list below...</td>
       </tr>
        <tr>
          <th align="right" valign="top" nowrap width="25%">Job Ids :</th>
          <% i = fJobsAll%>
          <td valign="top" nowrap><select size="<%=vJobsCnt%>" name="vJobs" multiple class="c2"><%= i%></select></td>
        </tr>
        <tr>
          <td align="left" colspan="2" class="c1" height="60">Step 2 : Select the appropriate Role(s) to apply the selected Job ID(s)...</td>
        </tr>
        <tr>
          <th align="right" valign="top" nowrap width="25%">
          Roles :</th>
          <td valign="top" nowrap>          
            <%
              aRoles = Split(Trim(Trim(Session("Completion_Roles_HO")) & " " & Trim(Session("Completion_Roles_XX"))))
              For i = 0 To Ubound(aRoles)
            %>
            <input type="checkbox" name="vRoles" id="vRoles" value="<%=aRoles(i)%>" <%=fchecks(Session("Completion_RoleP"), aRoles(i))%>><%=fPhraId(fRole_Title(aRoles(i)))%><br> 
            <%
              Next
            %>  
            <br>
            <input type="checkbox" name="checkRoles" name="checkRoles" onclick="checkOnOff(this, 'vRole');" value="null">Select All/None           
          </td>
        </tr>
        <tr>
          <td align="left" colspan="2" class="c1" height="60">Step 2 : Click to either &quot;Assign&quot; or &quot;Remove&quot; the selected Job(s)...</td>
        </tr>
        <tr>
          <td colspan="2" align="center" height="79">
         	  <input onclick="return bconfirm('Ok to Assign?')" type="submit" value="Assign" name="bAdd" class="button085">
         	  <%=f10%>
         	  <input onclick="return bconfirm('Ok to Remove?')"  type="submit" value="Remove" name="bDel" class="button085"> 
          </td>
        </tr>
      </table>
    </form>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>

