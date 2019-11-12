<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  '...this allows managers to modify learning rights that arrived here via the learner report
  
  Dim vPassword, vRole, vRegion, vSelected, vMessage, vOk, vWhere, aJobs, vCnt

  '...first time in?  
  If Session("Completion_InitParms") = "" Then 
    Session("Completion_InitParms") = "Y"
    Response.Redirect "Completion.asp?vNext=Completion_Learner.asp" 
  End If
  
  vSql = "SELECT * FROM Memb WITH (NOLOCK) INNER JOIN Crit WITH (NOLOCK) ON Memb.Memb_Criteria = Crit.Crit_No WHERE Memb_No = " & Request("vMemb_No")  
  sCompletion_Debug
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  sReadMemb

  vCrit_Id     = oRs("Crit_Id")
  vCrit_JobsId = oRs("Crit_JobsId")
  vRegion      = Left(vCrit_Id, Session("Completion_L1len"))
  vRole        = Right(vCrit_Id,  Session("Completion_RL1len"))

  Set oRs = Nothing      
  sCloseDb
 

  If Request.Form("bUpdate").Count = 1 Then
    vMemb_Jobs         = Replace(Request.Form("vMemb_Jobs"), ", ", " ")
    If Instr(vMemb_Jobs, "XXXX") > 0 Then vMemb_Jobs = ""
    vMemb_Memo         = Replace(Request.Form("vMemb_Memo"), ", ", " ")
    If Instr(vMemb_Memo, "XXXX") > 0 Then vMemb_Memo = ""
    vMemb_Level = Request("vMemb_Level")
    sUpdateMemb  svCustAcctId
  End If


  Function fUnitIds (vMemb_Memo)
    Dim vSelected, vPrev_Region
    vPrev_Region  = ""
    fUnitIds = vbCrLf
    vSql = " SELECT V5_Comp.dbo.Unit.Unit_L1, V5_Comp.dbo.Unit.Unit_L1Title, V5_Comp.dbo.Unit.Unit_L0, V5_Comp.dbo.Unit.Unit_L0Title "_
         & "   FROM V5_Comp.dbo.Unit WITH (NOLOCK) "_
         & " WHERE V5_Comp.dbo.Unit.Unit_AcctId = '" & svCustAcctId & "'" _
         & " ORDER BY V5_Comp.dbo.Unit.Unit_L1, V5_Comp.dbo.Unit.Unit_L0 "  
    sCompletion_Debug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      '...add in all before you start a new region
      If oRs("Unit_L1") <> vPrev_Region Then
        If Instr(vMemb_Memo, oRs("Unit_L1") & "|" & Session("Completion_L0all")) > 0 Then 
          vSelected = " selected" 
        Else
          vSelected = ""
        End If        
        
        fUnitIds = fUnitIds _
                 & "<option value='" & oRs("Unit_L1") & "|" & Session("Completion_L0all") & "'" & vSelected & " class='c4'>" _
                 & oRs("Unit_L1")  & " (" & oRs("Unit_L1Title")  & ") - " & Session("Completion_L0all") & " (All " & Session("Completion_L0tits") & ")"_
                 & "</option>" & vbCrLf

      End If
      vSelected = "" : If Instr(vMemb_Memo, oRs("Unit_L1") & "|" & oRs("Unit_L0")) > 0 Then vSelected = " selected" 
      fUnitIds = fUnitIds & "<option value='" & oRs("Unit_L1") & "|" & oRs("Unit_L0") & "'" & vSelected & ">" & oRs("Unit_L1")  & " (" & oRs("Unit_L1Title")  & ") - " &  oRs("Unit_L0") & " (" & oRs("Unit_L0Title") & ")</option>" & vbCrLf
      vPrev_Region = oRs("Unit_L1")
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
  <title>Completion_LearnerRights</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Extend Learner Rights for: </h1>

  <table class="table">
    <tr>
      <th style="width: 50%">Name:</th>
      <td class="c3" style="width: 50%"><%=vMemb_FirstName & " " & vMemb_LastName%></td>
    </tr>
    <tr>
      <th style="width: 50%">Email:</th>
      <td class="c3" style="width: 50%"><%=vMemb_Email%></td>
    </tr>
    <tr>
      <th style="width: 50%">Learner Id:</th>
      <td class="c3" style="width: 50%"><%=vMemb_Id%></td>
    </tr>
    <tr>
      <th style="width: 50%">Group:</th>
      <td class="c3" style="width: 50%"><%=vCrit_Id%></td>
    </tr>
    <tr>
      <th style="width: 50%">Vubiz No:</th>
      <td class="c3" style="width: 50%"><%=vMemb_No%></td>
    </tr>
    <tr>
      <th style="width: 50%">Level:</th>
      <td class="c3" style="width: 50%"><%=fIf(vMemb_Level=3, "Facilitator", "Learner")%></td>
    </tr>
  </table>

  <% If Len(vMessage) > 0 Then Response.Write "<h5>" & vMessage & "</h5>" %>


  <form method="POST" action="Completion_LearnerRights.asp" target="_self" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <input type="hidden" name="vMemb_No" value="<%=vMemb_No%>">

    <table class="table">

      <tr>
        <th>Provide Reporting Rights? </th>
        <td>
          <input type="radio" value="2" <%=fCheck(vMemb_Level, "2")%> name="vMemb_Level">No&nbsp;&nbsp; (leave at the default Learner level)<br>
          <input type="radio" value="3" <%=fCheck(vMemb_Level, "3")%> name="vMemb_Level">Yes&nbsp; (set learner to Facilitator level)<br><br>Note: this employee is still part of the uploaded learner group.&nbsp; Managers, on the other hand, are separate users setup by Vubiz and are not part of the learner group.
        </td>
      </tr>

      <tr>
        <th>Core Job Stream(s) :</th>
        <td>
          <%
          aJobs = Split(vCrit_JobsId)
          For i = 0 To Ubound(aJobs)
            sGetJobs  (aJobs(i))
            If i > 0 Then Response.Write "<br>"
            Response.Write vJobs_Id & " - " & vJobs_Title & fIf(vJobs_Active = 0, " (Inactive)", "")       
          Next
          %>
        </td>
      </tr>

      <tr>
        <th>Extended Learning :</th>
        <td>Job Streams in <span style="color:red; font-weight:bold">Red</span> are currently Inactive.<br>Do not include <font color="#000000">No Extended Learning</font>if you are selecting other Jobs.<br>
          <select name="vMemb_Jobs" size="1" multiple style="width: 400px; height: 200px">
            <option class="black" <%=fIf(Len(vMemb_Jobs) = 0, "selected", "") %> value="XXXX">No Extended Learning</option>
            <%=fJobs(vMemb_Jobs, vCrit_JobsId)%>
          </select>
        </td>
      </tr>


      <tr>
        <th>Extended Access :</th>
        <td>Do not include <span style="color:red; font-weight:bold">No Extended Access</span> if you select other Locations.<br>
          <select name="vMemb_Memo" size="1" multiple style="width: 400px; height: 200px">
            <option class="black" <%=fIf(Len(vMemb_Memo) = 0, "selected", "") %> value="XXXX">No Extended Access</option>
            <%=fUnitIds (vMemb_Memo)%>
          </select></td>
      </tr>

      <tr>
        <td colspan="2" style="text-align:center"><br /><br /></td>
      </tr>

      <tr>
        <td colspan="2" style="text-align:center">
          <input type="submit" value="Update" name="bUpdate" class="button"><br><br><a href="Completion_Learners.asp">Learner Report</a>
        </td>
      </tr>

    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>
