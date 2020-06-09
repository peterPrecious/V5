<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
  '...First time in?  
  If Session("Completion_InitParms") = "" Then 
    Session("Completion_InitParms") = "Y"
    Response.Redirect "Completion.asp?vNext=Completion_LocationManager.asp" 
  End If
%>

<html>

<head>
  <title>Completion_LocationManager</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--#include file = "Completion_LocationManager_Top.asp"-->

  <div style="margin-bottom: 30px;">
    <h1><%=Session("Completion_L0Tit")%> List</h1>
    <h2>To review and/or edit a location, click on the <%=Session("Completion_L0Tit")%> title.</h2>
  </div>

  <table style="width: 400px; margin: auto;">
    <tr>
      <td>
        <table style="width: 400px; margin: auto;">
          <%
            vSql = " SELECT * FROM V5_Comp.dbo.Unit AS Unit WITH (NOLOCK) "_ 
                 & " WHERE Unit_AcctId = '" & svCustAcctId & "' "_
                 & " ORDER BY Unit.Unit_HO DESC, Unit.Unit_L1, Unit.Unit_L0 "
            sCompletion_Debug
            sOpenDb
            Set oRs = oDb.Execute(vSql)
            Do While Not oRs.Eof
              sReadUnit
          %>
          <tr>
            <td><a href="Completion_LocationManager_Rev.asp?vUnit_No=<%=vUnit_No%>"><%=vUnit_L1 & " " & vUnit_L0 & " : " & vUnit_L1Title & " | " &  vUnit_L0Title%></a> </td>
          </tr>
          <%
              oRs.MoveNext
            Loop
            Set oRs = Nothing
            sCloseDb
          %>
        </table>

      </td>
    </tr>

  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>


