<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
  vUnit_No = Request("vUnit_No")
  sGetUnit vUnit_No
%>

<html>

<head>
  <title>Completion_LocationManager_Rev</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--#include file = "Completion_LocationManager_Top.asp"-->

  <div style="margin-bottom: 30px;">
    <h1>Job Streams by Role</h1>
    <h2><%=vUnit_L1Title & " | " & vUnit_L0Title & "<br>(" & vUnit_L1 & "|" & vUnit_L0 & ")"%></h2>
  </div>


  <table style="width:500px; margin:0 auto 30px; auto">
    <% 
      Dim vRoleLen, vLocnLen
      i = ""
      vRoleLen = Session("Completion_RLlen")
      vLocnLen = Session("Completion_L1len") + Session("Completion_L0len") + 1
      vSql = ""_
            & " SELECT"_     
            & "   RIGHT(Cr.Crit_Id, " & Session("Completion_RLlen") & ") AS Role, Cr.Crit_JobsId AS Jobs "_
            & " FROM"_         
            & "   V5_Vubz.dbo.Crit AS Cr WITH (NOLOCK)                                                                                           INNER JOIN"_
            & "   V5_Comp.dbo.Unit AS Un WITH (NOLOCK)                                                                                           ON "_
            & "     Cr.Crit_AcctId = Un.Unit_AcctId                                                                                AND "_ 
            & "     LEFT(Cr.Crit_Id, " & Session("Completion_L1len") & ") = Un.Unit_L1                                             AND "_
            & "     SUBSTRING(Cr.Crit_Id, " & Session("Completion_L0str") & ", " & Session("Completion_L0len") & ") = Un.Unit_L0"_
            & " WHERE"_     
            & "   (Un.Unit_No = " & vUnit_No & ")"
      sCompletion_Debug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
    %>
    <tr>
      <td><%=oRs("Role")%></td>
      <td><%=oRs("Jobs")%></td>
    </tr>
    <%    
      oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
    %>
  </table>


  <input type="button" onclick="location.href = 'Completion_LocationManager_Mod.asp?vUnit_No=<%=vUnit_No%>'" value="Edit" name="bMod" class="button070">
  <% If Not bUnit_HO Then   '...can't move HO%>
  <input type="button" onclick="location.href = 'Completion_LocationManager_Mov.asp?vUnit_No=<%=vUnit_No%>'" value="Move" name="bMov" class="button070">
  <% End If %>
  <% If Session("Completion_Level") = 5 And fUnitDeleteOk (vUnit_L1, vUnit_L0) Then %>
  <input type="button" onclick="location.href = 'Completion_LocationManager_Del.asp?vUnit_No=<%=vUnit_No%>'" value="Delete" name="bDelete" class="button070">
  <% End If %>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>
