<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->

<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
  Dim aUnit
  vUnit_No = Request("vUnit_No")
  sGetUnit vUnit_No


  If Request.Form.Count > 0 Then

    sOpenDb2

    '...update the unit table
    aUnit= Split(Request("vNewL1"), "|")   
    vSql = "UPDATE "_
         & "  V5_Comp.dbo.Unit "_
         & "SET " _
         & "  Unit_L1        = '" & aUnit(0) & "', " _
         & "  Unit_L1Title   = '" & fUnQuote(aUnit(1)) & "' " _
         & "WHERE "_
         & "  Unit_No        =  " & vUnit_No
    sCompletion_Debug
    oDb2.Execute(vSql)

    '...update the criteria table
    vSql = " UPDATE "_
         & "   V5_Vubz.dbo.Crit "_
         & " SET "_
         & "   Crit_Id = '" & aUnit(0) & "' + SUBSTRING(Crit_Id, " & Session("Completion_L1len") + 1 & ", 99)"_
         & " WHERE "_
         & "   Crit_AcctId = '" & svCustAcctId & "' AND SUBSTRING(Crit_Id, " & Session("Completion_L0str") & ", " & Session("Completion_L0len") & ") = '" & vUnit_L0 & "'"
    sCompletion_Debug
    oDb2.Execute(vSql)

    sCloseDb2 

    Response.Redirect "Completion_LocationManager_Rev.asp?vUnit_No=" & vUnit_No

  End If    

%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

	</head>

	<body>

		<% Server.Execute vShellHi %>
		<form method="POST" action="Completion_LocationManager_Mov.asp">
			<div align="center">
				<table border="1" cellpadding="2" bordercolor="#DDEEF9" width="600" cellspacing="0" style="border-collapse: collapse">
					<tr>
						<td width="100%" align="center">

  						<br><b>Move...<br>
  						<%=vUnit_L1Title%> | <%=vUnit_L0Title%> (<%=vUnit_L1%> | <%=vUnit_L0%>)<br>
  						...to another <%=Session("Completion_L1Tit")%>.</b><p>Select a new <%=Session("Completion_L1Tit")%> then click Update<br>
						  <br>

  						<% i = fL1s ("Move")%> 
  						<select class="c2" name="vNewL1" size="<%=fDefault(vSelectNo , 7)%>"><%=i%></select>
  						
  						<br><br>
  						<input onclick="location.href='Completion_LocationManager_Rev.asp?vUnit_No=<%=vUnit_No%>'" type="button" value="Cancel" name="bCancel0" class="button070"> <input type="submit" value="Update" name="bRegion" class="button070"> <br>
  						<br><br><br>

						</td>
					</tr>
				</table>
			</div>
			<input type="hidden" name="vUnit_No" value="<%=vUnit_No%>">
		</form>
		<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
		<!--#include file = "Completion_Footer.asp"-->

	</body>

</html>