<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vActive, vGlobal, vNext, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vFormat, vLearners
  Dim vLastValue, vDetails, vCurList, vRecnt, vWhere, aCrit, vEdit, vExtend, vRole

  vNext            = Request("vNext")
  vActive          = fDefault(Request("vActive"), "1")
  vGlobal          = fDefault(Request("vGlobal"), "0")
  vFind            = fDefault(Request("vFind"), "S")
  vFindId          = fUnQuote(Request("vFindId"))
  vFindFirstName   = fUnQuote(Request("vFindFirstName"))
  vFindLastName    = fUnQuote(Request("vFindLastName"))
  vFindEmail       = fNoQuote(Request("vFindEmail"))
  vFindCriteria    = Request("vFindCriteria")
  vFormat          = fDefault(Request("vFormat"), "o")
  vLearners        = fDefault(Request("vLearners"), "n")

  vDetails         = Request("vDetails") 
  vLastValue       = Request("vLastValue") 
  vCurList         = fDefault(Request("vCurList"), 0)
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Learner Report</title>
</head>

<body>

  <% Server.Execute vShellHi %>

  <div style="text-align: center; margin-bottom: 30px;">
    <h1>Learner Report</h1>
    <p>The Learner Report is sorted by <%=Session("Completion_L1tit")%> | <%=Session("Completion_L0tit")%> | Role | Last Name.</p>
  </div>

  <table class="table" style="width: 800px; margin: auto;">
    <tr>
      <td class="rowshade" style="width: 080px; text-align: center;"><%=Session("Completion_L1tit")%></td>
      <td class="rowshade" style="width: 100px; text-align: center;"><%=Session("Completion_L0tit")%></td>
      <td class="rowshade" style="width: 070px; text-align: center;">Role</td>
      <td class="rowshade" style="width: 100px; text-align: left;">Learner ID</td>
      <td class="rowshade" style="width: 100px; text-align: left;">First Name</td>
      <td class="rowshade" style="width: 100px; text-align: left;">Last Name</td>
      <td class="rowshade" style="width: 070px; text-align: center;">Active?</td>
      <td class="rowshade" style="width: 070px; text-align: center;">Action</td>
    </tr>
    <tr>
      <td colspan="8">&nbsp;</td>
    </tr>

    <%      
      vSql = " SELECT * FROM "_
           & "   Memb WITH (NOLOCK) INNER JOIN "_ 
           & "   Crit WITH (NOLOCK) ON TRY_CAST(Memb.Memb_Criteria AS INT) = Crit.Crit_No"_      
           & " WHERE "_
           & "   (Memb_AcctId = '" & svCustAcctId & "')"_ 
           & "   AND (Memb_Level <= " & Session("Completion_Level") & ")"_ 
           & "   AND (Memb_Internal = 0)"_ 
           & "   AND (ISNUMERIC(Memb_Criteria) = 1)"_ 
           &     fIf (Len(vLastValue)    > 0, " AND (ISNULL(Memb_LastName,'') + ISNULL(Memb_FirstName,'') + CAST(Memb_No AS VARCHAR(10)) >= '" & fUnquote(vLastValue) & "')","") _
           &     fIf (Len(vFindId)       > 0, " AND (Memb_Id       LIKE '" & vFindId         & "%')", "") _
           &     fIf (Len(vFindLastName) > 0, " AND (Memb_LastName LIKE '" & vFindLastName   & "%')", "") _   
           &     fIf (vFindCriteria <> "All", " AND (Crit_No IN (" & vFindCriteria & "))", "") _
           &     fIf (vActive <> "*", " AND (Memb_Active = " & vActive & ")", "") _
           & " ORDER BY "_
           & "   Crit_Id, ISNULL(Memb_LastName,'') + ISNULL(Memb_FirstName,'') + CAST(Memb_No AS varchar(10))"

      sCompletion_Debug 
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        sReadMemb
        vCrit_Id = oRs("Crit_Id")
        vCurList = vCurList + 1
        
				vEdit = "" : vExtend = ""
        If Session("Completion_Level") >= 3 And Session("Completion_EditLearners") Then
        	vEdit = "<a href='Completion_Learner.asp?vMemb_No=" & vMemb_No & "'>Edit</a>"
        End If
        If Session("Completion_Level") > 3 Then
        	vExtend = "<a href='Completion_LearnerRights.asp?vMemb_No=" & vMemb_No & "'>Extend</a>"
				End If        

    %>
    <tr>
      <td style="white-space: nowrap; text-align: center;"><%=Left(vCrit_Id, Session("Completion_L1len"))%> </td>
      <td style="white-space: nowrap; text-align: center;"><%=Mid(vCrit_Id, Session("Completion_L0str"), Session("Completion_L0len"))%> </td>
      <td style="white-space: nowrap; text-align: center;"><%=Right(vCrit_Id, Session("Completion_RLlen"))%> </td>
      <td style="white-space: nowrap; text-align: left"><%=vMemb_Id%></td>
      <td style="white-space: nowrap; text-align: left"><%=vMemb_FirstName%> </td>
      <td style="white-space: nowrap; text-align: left"><%=vMemb_LastName%> </td>
      <td style="white-space: nowrap; text-align: center"><%=fIf(vMemb_Active, "Y", "N")%></td>
      <td style="white-space: nowrap; text-align: center"><%= Trim(vEdit & " " & vExtend)%></td>
    </tr>
    <%
        oRs.MoveNext
        If Cint(vCurList) > 0 And Cint(vCurList) Mod 100 = 0 Then Exit Do
      Loop
      Set oRs = Nothing
      sCloseDb
    %>
  </table>

  <div>
    <form method="POST" action="Completion_Learners_o.asp">

      <input type="hidden" name="vLearners" value="<%=vLearners%>">
      <input type="hidden" name="vFormat" value="<%=vFormat%>">
      <input type="hidden" name="vLastValue" value="<%=vMemb_LastName & vMemb_FirstName & vMemb_No%>">
      <input type="hidden" name="vFindCriteria" value="<%=vFindCriteria%>">
      <input type="hidden" name="vFindEmail" value="<%=vFindEmail%>">
      <input type="hidden" name="vFindLastName" value="<%=vFindLastName%>">
      <input type="hidden" name="vFindFirstName" value="<%=vFindFirstName%>">
      <input type="hidden" name="vFind" value="<%=vFind%>">
      <input type="hidden" name="vActive" value="<%=vActive%>">
      <input type="hidden" name="vCurList" value="<%=vCurList%>">
      <input type="hidden" name="vNext" value="<%=vNext%>">

      <p style="text-align: center; margin: 40px;">
        <input type="button" onclick="location.href = '<%=fDefault(vNext, "javascript:history.back(1)")%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button085">
        <% If Cint(vCurList) > 0 And Cint(vCurList) Mod 100 = 0 Then '...If next group, get next starting value %>
        <input type="submit" name="bNext" value="<%=bNext%>" class="button085">
        <% End If %>
        <br><br>
        <a href="Completion_Learners.asp">Restart Report</a>
      </p>

    </form>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>


