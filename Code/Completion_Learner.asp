<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->

<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vMyGr, vMyL1, vMyL0, vMyRL, vMyCh, vRole, vRegion, vSelected, vMessage, vOk, vWhere, aJobs, vCnt, bOk

  '...First time in?  
  If Session("Completion_InitParms") = "" Then 
    Session("Completion_InitParms") = "Y"
    Response.Redirect "Completion.asp?vNext=Completion_Learner.asp" 
  End If

  If Session("Completion_Level") < 3 Then Response.Redirect "Error.asp?vErr=" & Server.UrlEncode("This service is only available to Facilitators and Managers.") & "&vReturn=n"

  '...First time in?  
  If Session("Completion_InitParms") = "" Then 
    Session("Completion_InitParms") = "Y"
    Response.Redirect "Completion.asp?vNext=Completion_Learner.asp" 
  End If

  '...determine rights of user (RRRR TTTT R)
  vMyGr = fCriteria (svMembCriteria)
  vMyL1 = Left(vMyGr, Session("Completion_L1len"))
  vMyL0 = Mid(vMyGr, Session("Completion_L0str"), Session("Completion_L0len"))
  vMyRL = Right(vMyGr, Session("Completion_RLlen"))
	vMyCh = fRole_Children(vMyRL)

  vMemb_Active   = fDefault(Request("vMemb_Active"), 1)
  vMemb_Criteria = fDefault(Request("vMemb_Criteria"), svMembCriteria)
  
  If Request.Form("bDelete").Count = 1 Then
    vMemb_No = Request("vMemb_No")
    sDeleteMemb
    vMemb_No = 0
    vMemb_Criteria = "0"
    vMessage = fPhraH(000653)        

  ElseIf Request.Form("bUpdate").Count = 1 Or Request.Form("bAdd").Count = 1 Then

    vMemb_Id            = Ucase(Request("vMemb_Id"))
    vMemb_Email         = Lcase(vMemb_Id)
    vMemb_No            = Request.Form("vMemb_No")
    vMemb_FirstName     = Trim(Request.Form("vMemb_FirstName"))
    vMemb_LastName      = Trim(Request.Form("vMemb_LastName"))
    vMemb_Active        = fDefault(Request.Form("vMemb_Active"), 1)
    vMemb_Criteria      = Request.Form("vMemb_Criteria")
    vMemb_Jobs          = Replace(Request.Form("vMemb_Jobs"), ", ", " ")
    If Instr(vMemb_Jobs, "XXXX") > 0 Then vMemb_Jobs = ""
    vMemb_Memo          = Replace(Request.Form("vMemb_Memo"), ", ", " ")
    If Instr(vMemb_Memo, "XXXX") > 0 Then vMemb_Memo = ""

		vMemb_Level = fIf(Len(fRole_Children(Right(fCriteria(vMemb_Criteria), Session("Completion_RLlen")))) > 0, 3, 2)

    bOk = True
    If vMemb_No = 0 Then
      If spMembExistsById (svCustAcctId, vMemb_Id) Then 
        bOk = False
      End If
    ElseIf vMemb_Id <> spMembIdByNo (vMemb_No) Then 
      If spMembExistsById (svCustAcctId, vMemb_Id) Then 
        bOk = False
      End If
    End If      

    If bOk = False Then
      vMessage = "There is already a learner on file with that Password<br>'" & vMemb_Id & "'"
    Else
      sUpdateMemb svCustAcctId
      vMessage = "Learner was added/modified successfully."
    End If
  
  ElseIf Request.QueryString("vMemb_No").Count = 1 Then
    vMemb_No = Request.QueryString("vMemb_No")
    If vMemb_No > 0 Then
      vSql = " SELECT * FROM Memb WITH (NOLOCK) INNER JOIN Crit WITH (NOLOCK) ON Memb.Memb_Criteria = Crit.Crit_No WHERE Memb_No = " & vMemb_No  
      sCompletion_Debug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      sReadMemb
      vCrit_No  = oRs("Crit_No")
      vCrit_Id  = oRs("Crit_Id")
      vRegion   = Left(vCrit_Id, Session("Completion_L1len"))
      vRole     = Right(vCrit_Id, Session("Completion_RLlen"))
      Set oRs = Nothing      
      sCloseDb
    Else
      vRegion  = ""
    End If

  End If


  Function fL0s (vMemb_Memo)
    Dim vSelected, vPrL1
    vPrL1  = ""
    fL0s = vbCrLf
    vSql = " SELECT Unit_L1, Unit_L1Title, Unit_L0, Unit_L0Title "_
         & " FROM V5_Comp.dbo.Unit WITH (NOLOCK) "_
         & " WHERE Unit_AcctId = '" & svCustAcctId & "' "_
         & " ORDER BY Unit_L1, Unit_L0 "  
    sCompletion_Debug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      '...add in all before you start a new region
      If oRs("Unit_L1") <> vPrL1 Then
        vSelected = "" : If Instr(vMemb_Memo, oRs("Unit_L1") & "|" & Session("Completion_L0all")) > 0 Then vSelected = " selected" 
        fL0s = fL0s & "<option class='c4' value='" & oRs("Unit_L1") & "|" & Session("Completion_L0all") & "'" & vSelected & ">" & oRs("Unit_L1")  & " (" & oRs("Unit_L1Title")  & ") - 0000 (All " & fIf(oRs("Unit_L1")="8030", "Departments", "L0s") & ")</option>" & vbCrLf
      End If
      vSelected = "" : If Instr(vMemb_Memo, oRs("Unit_L1") & "|" & oRs("Unit_L0")) > 0 Then vSelected = " selected" 
      fL0s = fL0s & "<option value='" & oRs("Unit_L1") & "|" & oRs("Unit_L0") & "'" & vSelected & ">" & oRs("Unit_L1")  & " (" & oRs("Unit_L1Title")  & ") - " &  oRs("Unit_L0") & " (" & oRs("Unit_L0Title") & ")</option>" & vbCrLf
      vPrL1 = oRs("Unit_L1")
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
  End Function


  Function fJobs (vMemb_Memo, vCrit_Jobs)
    Dim vSelected, vClass
    fJobs = vbCrLf
    vSql = " SELECT Jobs_Id, Jobs_Active, Jobs_Title FROM Jobs WITH (NOLOCK) WHERE (Jobs_AcctId = '" & svCustAcctId & "') AND (CHARINDEX(Jobs_Id, '" & vCrit_Jobs & "') = 0) ORDER BY Jobs_Id "
    sCompletion_Debug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      If oRs("Jobs_Active") Then vClass = "" Else vClass = " class='c6'"
      vSelected = "" : If Instr(vMemb_Memo, oRs("Jobs_Id")) > 0 Then vSelected = " selected" 
      fJobs = fJobs & "<option value=" & Chr(34) & oRs("Jobs_Id") & Chr(34) & vSelected & vClass & ">" & oRs("Jobs_Id") & " - " & oRs("Jobs_Title") & fIf(oRs("Jobs_Active"), "", " (Inactive)") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop      
    sCloseDb           
    Set oRs = Nothing
  End Function

%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>User Profiles</title>
  <script>

    function fValidate() {

			var temp;

			temp = document.fForm.vMemb_Id.value;
			if (temp.length < 3){
			  alert("Please enter a valid Learner ID.");
			  document.fForm.vMemb_Id.focus();
			  return (false);
			}
			
			temp = document.fForm.vMemb_FirstName.value;
			if (temp.length < 1){
			  alert("Please enter a First Name.");
			  document.fForm.vMemb_FirstName.focus();
			  return (false)
			}
			
			temp = document.fForm.vMemb_LastName.value;
			if (temp.length < 1){
			  alert("Please enter a LastName.");
			  document.fForm.vMemb_LastName.focus();
			  return (false);
			}
			
			if (document.fForm.vMemb_Criteria.selectedIndex < 0) {
			  alert("Please select a Location.");
			  document.fForm.vMemb_Criteria.focus();
			  return (false);
			}

			return (true);
    }
    
    
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <form method="POST" action="Completion_Learner.asp" target="_self" onsubmit="return fValidate()" id="fForm" name="fForm">

    <input type="hidden" name="vMemb_No" value="<%=fDefault(vMemb_No, 0)%>">
    <input type="hidden" name="vMemb_Level" value="<%=fDefault(vMemb_Level, 2)%>">

    <table class="table">
      <tr>
        <td colspan="2" style="text-align:center">
          <h1><br><!--webbot bot='PurpleText' PREVIEW='Learner Profile'--><%=fPhra(000371)%></h1>
          <% If Len(vMessage) > 0 Then %><h5><%=vMessage%></h5><% End If %>
        </td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Learner ID'--><%=fPhra(000411)%>:</th>
        <td><input type="text" size="30" id="vMemb_Id" name="vMemb_Id" value="<%=vMemb_Id%>"><% If vMemb_No > 0 Then %>&nbsp;(Vubiz No: <%=vMemb_No%>)<% End If %></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%>:</th>
        <td><input type="text" size="30" name="vMemb_FirstName" value="<%=vMemb_FirstName%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%>:</th>
        <td><input type="text" size="30" name="vMemb_LastName" value="<%=vMemb_LastName%>" maxlength="64"></td>
      </tr>

      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Active'--><%=fPhra(000063)%>:</th>
        <td>
          <input type="radio" name="vMemb_Active" value="0" <%=fcheck(0, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%>
          <input type="radio" name="vMemb_Active" value="1" <%=fcheck(1, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%>
        </td>
      </tr>

      <% If vMemb_No > 0 Then %>
      <tr>
        <th>Level :</th>
        <td><%=fIf(vMemb_Level = 3, "Facilitator", fIf(vMemb_Level = 4, "Manager", fIf(vMemb_Level = 5, "Administrator", "Learner")))%> (Note: a Facilitator belongs to a Role that has Children)</td>
      </tr>
      <% 
				End If

        i = fLocation(vMemb_Criteria)        
      %>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Location'--><%=fPhra(001325)%>:</th>
        <td><select  name="vMemb_Criteria" style="width: 400px" size="<%=fMin(vCnt, 30)%>"><%=i%></select></td>
      </tr>


      <!--  Do not access these fields until learner is added -->
      <% If vMemb_No > 0 Then %>

      <tr>
        <th>
          <!--webbot bot='PurpleText' PREVIEW='Core Learning Programs'--><%=fPhra(000804)%>
          :</th>
        <td>
          <%
          sGetCrit (vMemb_Criteria)
          aJobs = Split(vCrit_JobsId)
          For i = 0 To Ubound(aJobs)
            sGetJobs  (aJobs(i))
            If i > 0 Then Response.Write "<br>"
            Response.Write vJobs_Id & " - " & vJobs_Title & fIf(vJobs_Active = 0, " (Inactive)", "")       
          Next
          %>
        </td>
      </tr>

      <% If Session("Completion_Level") > 3  Then %>

      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Additional Learning Programs'--><%=fPhra(000805)%>:</th>
        <td>
          <!--webbot bot='PurpleText' PREVIEW='Note Job Streams in <span style="font-weight:bold" class="red">Red</span> are currently Inactive.'--><%=fPhra(001630)%><br />
          <select name="vMemb_Jobs" size="1" multiple style="width: 400px; height: 200px">
            <option <%=fIf(Len(vMemb_Jobs) = 0, "selected", "") %> value="XXXX">No Additional Learning</option>
            <%=fJobs(vMemb_Jobs, vCrit_JobsId)%>
          </select><br>
          <!--webbot bot='PurpleText' PREVIEW='Do not include &quot;No Additional Learning&quot; with other Programs.'--><%=fPhra(000807)%>
        </td>
      </tr>


      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Extend Report Card Access Rights'--><%=fPhra(000808)%>:<br>&nbsp;</th>
        <td>
          <select name="vMemb_Memo" size="1" multiple style="width: 400px; height: 200px">
            <option <%=fIf(Len(vMemb_Memo) = 0, "selected", "") %> value="XXXX"><!--webbot bot='PurpleText' PREVIEW='No Extended Rights'--><%=fPhra(000810)%></option>
            <%=fL0s (vMemb_Memo)%>
          </select><br>
          <!--webbot bot='PurpleText' PREVIEW='Do not include &quot;No Extended Rights&quot; with other Rights.'--><%=fPhra(000811)%>
        </td>
      </tr>

      <% End If %>
      <% End If %>

      <tr>
        <td style="text-align:center" colspan="2">&nbsp;<br><br>
          <% If vMemb_No = 0 Then %>
          <input type="submit" value="Add Learner" name="bAdd" class="button">
          <% Else %>
          <input type="submit" value="Update Learner's Profile" name="bUpdate" class="button">
          <%   If Session("Completion_Level") = 5 And vMemb_No > 0 Then %><br>
          <br><br>If you delete this learner, their history will be permanently lost!<br>
          <input type="submit" value="Delete Learner" name="bDelete" class="button"><br>
          <%   End If %>
          <% End If %>

          <table style="width:600px; text-align:center; margin:auto;">
            <tr>
              <td>
                <h2><a href="Completion_Learners.asp"><!--webbot bot='PurpleText' PREVIEW='Learner Report'--><%=fPhra(000367)%></a></h2>
              </td>
              <% If vMemb_No > 0 Then %>
              <td>
                <h2><a href="Completion_Learner.asp?vMemb_No=0"><!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%></a></h2>
              </td>
              <%   If Session("Completion_Level") > 3 Then %>
              <td>
                <h2><a href="Completion_LocationManager.asp?vMemb_No=<%=vMemb_No%>"><!--webbot bot='PurpleText' PREVIEW='Change Learner's Location'--><%=fPhra(000784)%></a></h2>
              </td>
              <%   End If %>
              <% End If %>
            </tr>
          </table>

        </td>
      </tr>


    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>


