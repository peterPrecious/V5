<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->


<!--#include virtual = "V5/Inc/Document.asp"-->
<!--#include virtual = "V5/Inc/Base64.asp"-->

<%
	Dim vActive, vGlobal, vCustId, vNext, vEdit, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vLearners, vLevel
	Dim vLastValue, vDetails, vCurList, vRecnt, vWhere, aCrit, bG2CustEmail, bG2MembEmail, vCols

	Dim vGlobalCustId, vSource

	Select Case svMembLevel
		Case 3 : vLearners = "23"
		Case 4 : vLearners = "234"
		Case 5 : vLearners = "2345"
	End Select

	vNext            = Request("vNext")
	vEdit            = fDefault(Request("vEdit"), "User" & fGroup & ".asp")
	vCustId          = fDefault(Request("vCustId"), svCustId)
	vActive          = fDefault(Request("vActive"), "1")
	vGlobal          = fDefault(Request("vGlobal"), "0")
	vFind            = fDefault(Request("vFind"), "S")
	vFindId          = fUnQuote(Request("vFindId"))
	vFindFirstName   = fUnQuote(Request("vFindFirstName"))
	vFindLastName    = fUnQuote(Request("vFindLastName"))
	vFindEmail       = fNoQuote(Request("vFindEmail"))
	vFindMemo        = fUnQuote(Request("vFindMemo"))
	vFindCriteria    = Request("vFindCriteria")
	vFormat          = fDefault(Request("vFormat"), "o")
	vLearners        = fDefault(Request("vLearners"), vLearners)

	vDetails         = Request("vDetails") 
	vLastValue       = Request("vLastValue") 
	vCurList         = fDefault(Request("vCurList"), 0)

	vWhere = ""
	If svMembLevel < 4 Then vWhere = vWhere & " AND (Memb_Level < 4)"
	If vActive = "0" Then vWhere = vWhere & " AND (Memb_Active = 1)"

	'...If there is a Last Value
	If Len(vLastValue) > 0 Then
		vWhere = vWhere & " AND (ISNULL(Memb_LastName,'') + ISNULL(Memb_FirstName,'') + CAST(Memb_No AS VARCHAR(10)) >= '" & fUnquote(vLastValue) & "')"
	End If

	If vFind = "S" Then
		If Len(vFindId)        > 0 Then vWhere = vWhere & " AND (Memb_Id        LIKE '" & vFindId         & "%')"
		If Len(vFindFirstName) > 0 Then vWhere = vWhere & " AND (Memb_FirstName LIKE '" & vFindFirstName  & "%')"
		If Len(vFindLastName)  > 0 Then vWhere = vWhere & " AND (Memb_LastName  LIKE '" & vFindLastName   & "%')"
		If Len(vFindEmail)     > 0 Then vWhere = vWhere & " AND (Memb_Email     LIKE '" & vFindEmail      & "%')"
		If Len(vFindMemo)      > 0 Then vWhere = vWhere & " AND (Memb_Memo      LIKE '" & vFindMemo       & "%')"
	Else
		If Len(vFindId)        > 0 Then vWhere = vWhere & " AND (Memb_Id        LIKE '%" & vFindId        & "%')"
		If Len(vFindFirstName) > 0 Then vWhere = vWhere & " AND (Memb_FirstName LIKE '%" & vFindFirstName & "%')"
		If Len(vFindLastName)  > 0 Then vWhere = vWhere & " AND (Memb_LastName  LIKE '%" & vFindLastName  & "%')"
		If Len(vFindEmail)     > 0 Then vWhere = vWhere & " AND (Memb_Email     LIKE '%" & vFindEmail     & "%')"
		If Len(vFindMemo)      > 0 Then vWhere = vWhere & " AND (Memb_Memo      LIKE '%" & vFindMemo      & "%')"
	End If

	If Len(vFindCriteria)    > 2 Then '...criteria can be 129,330 or just 129
		vWhere = vWhere & " AND ("
		aCrit = Split(vFindCriteria, ",")
		For i = 0 To Ubound(aCrit)
			vWhere = vWhere & fIf(i = 0, "", " OR ") & "Memb_Criteria = '" & aCrit(i)  & "'"
		Next    
		vWhere = vWhere & ")"
	End If   

	vLevel = ""  

	If Instr(vLearners, "s") > 0 Then 
		vLevel = "2"
		vWhere = vWhere & " AND (Memb_Sponsor > 0)"
	End If

	If Instr(vLearners, "1") > 0 Then
		vLevel = "1"
	End If

	If Instr(vLearners, "2") > 0 Then     '...if sponsors (above) then we are already grabbing level 2 (learner)
		If vLevel = "" Then 
			vLevel = "2" 
		Else
			vLevel = vLevel & ",2"
		End If   
	End If

	If Instr(vLearners, "3") > 0 Then
		If vLevel = "" Then 
			vLevel = "3"
		Else
			vLevel = vLevel & ",3"
		End If   
	End If
	If Instr(vLearners, "4") > 0 Then
		If vLevel = "" Then 
			vLevel = "4"
		Else
			vLevel = vLevel & ",4"
		End If   
	End If
	If Instr(vLearners, "5") > 0 Then
		If vLevel = "" Then 
			vLevel = "5"
		Else
			vLevel = vLevel & ",5"
		End If   
	End If

	vWhere = vWhere & " AND (Memb_Level IN (" & vLevel & "))"
	vWhere = vWhere & " AND (Memb_Id NOT LIKE '" & vPasswordx & "%' )"

	'...determine if this is a G2 site for Email alerts (bG2CustEmail)
	sGetCust vCustId
	bG2CustEmail = fIf(fCustG2Ok(vCustId) And vCust_EcomG2alert, True, False)
	vCols = fIf(bG2CustEmail, 8, 7)

	'...this prefixes the password with the Customer AcctId 
	Function fGlobal
		If vGlobal = 1 Then
'     fGlobal = "(" & oRs("Cust_Id) & ") "
		Else
			fGlobal = ""
		End If 
	End Function


	Function fRights
		fRights = ""
		If vMemb_Auth    Then fRights = " A"
		If vMemb_MyWorld Then fRights = fRights & " M"
		If vMemb_LCMS    Then fRights = fRights & " L"
		If Len(Trim(fRights)) = 0 Then fRights = ""
	End Function

%>

<html>

<head>
	<title>Users_O</title>
	<meta charset="UTF-8">
	<% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
	<script src="/V5/Inc/jQuery.js"></script>
	<link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
	<script src="/V5/Inc/Functions.js"></script>
	<script src="/V5/Inc/WebService.js"></script>
	<script src="/V5/Inc/Launch.js"></script>
	<script>
		function resendEmails(vMembNo, vLang) {
			var vAlert01  = (vLang == "EN") ?"Thank you. This learner's Welcome Email\nwill be resent within the next two hours." : "Merci. Le courriel de Bienvenue de l’apprenant sera renvoyé dans les deux prochaines heures."
			var vAlert02  = (vLang == "EN") ?"We are unable to provide this service, please notify support@vubiz.com." : "Nous sommes incapables de fournir ce service, s'il vous plaît aviser support@vubiz.com."
			var vParam    = "vFunction=ResendEmail&vMemberID =" + vMembNo;
			var vWs       = WebService("Users_ws.asp", vParam);  
			if (vWs == "ok") {
				alert(vAlert01);
				return (false);
			} else {   
				alert(vAlert02);
				return;
			}
		}

		function history(vMembNo) {
			vUrl ="/Gold/vuSCORMadmin/Default.aspx?memberId=" + vMembNo
			vuwindow(vUrl,1000,600,10,10,"y","Y","Y")
		}
	</script>
	<style type="text/css">
		.impersonate {
			width: 16px;
			height: 16px;
		}
	</style>
</head>

<body>

	<% Server.Execute vShellHi %>

	<h1><!--[[-->Learner Report<!--]]--></h1>
	<p class="c3">
		<!--[[-->Note:&nbsp; <b>Expires</b> will be blank unless the Learner was setup manually or via e-commerce.&nbsp; The <b>Group|Rights</b> filters will be empty unless configured.&nbsp; Click on the<!--]]--> <b><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></b> <!--[[-->to modify the Learner&#39;s profile.<!--]]-->&nbsp; 
		<% If bG2CustEmail Then %>
			<!--[[-->If a learner did not receive their &quot;Welcome&quot; email alert, for whatever reason, click the &quot;Resend&quot; button beside that learner's name.<!--]]-->
		<% End If %>       
	</p>

		<% If svMembLevel = 5 Then %>
		<p class="c6">Clicking on the Impersonate Icon <img class="impersonate" src="../Images/Impersonate.png" /> will convert this session into that of the selected learner. Sign Off when done and you will be returned to this session at the Learner Report.</p>
		<% End If %>        


	<table class="table">
		<tr>
			<th class="rowShade" style="text-align:left; width:150px"><!--[[-->Group<!--]]--><% If svMembLevel = 5 Then %><br><!--[[-->Rights<!--]]--><% End If %></th>
			<th class="rowShade" style="text-align:left; width:200px"><!--[[-->Name<!--]]-->, <!--[[-->Organization<!--]]--><br><!--[[-->Email Address<!--]]--></th>
			<th class="rowShade" style="text-align:center; width:100px"><!--[[-->Active<!--]]-->?<%=fIf(svMembLevel=5, "<br />History", "")%></th>
			<th class="rowShade" style="text-align:left; width:200px"><%=fIf(vGlobal="1","(CustId) ","")%> <%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%><br><!--[[-->Memo<!--]]--></th>
			<th class="rowShade" style="text-align:center; width:100px"><!--[[-->First Visit<!--]]--><br /><!--[[-->Last Visit<!--]]--></th>
			<th class="rowShade" style="text-align:center; width:100px"><!--[[-->Expires<!--]]--></th>
			<th class="rowShade" style="text-align:center; width:100px"># <!--[[-->Site Visits<!--]]--><br><!--[[-->Total Mins<!--]]--><br /><!--[[-->Online<!--]]--></th>
			<% If bG2CustEmail Then %>
			<th class="rowShade" style="text-align:center; width:100px"><!--[[-->Resend<!--]]--><br><!--[[-->Email Alert<!--]]--></th>
			<% End If %>
		</tr>
		<tr>
			<td colspan="<%=vCols%>">&nbsp;</td>
		</tr>
		<%  

			sGetMemb_Rs vCust_AcctId, vWhere, vGlobal

			Do While Not oRs.Eof
				sReadMemb '... oRs("Memb_NoMins") was added May 2, 2017 to get a more accurate time online - overrides normal Memb_NoHours - uses function and returns minutes - from sGetMemb_Rs
 
				'...get vGlobalCustId for Impersonation
				If vGlobal = 1 Then vGlobalCustId = oRs("Cust_Id")

				'...determine if this is a G2 learner (bG2MembEmail) for Resend Button
				bG2MembEmail = fIf(bG2CustEmail And vMemb_EcomG2alert And vMemb_Level = 2, True, False)
				vMemb_Id = fDefault(vMemb_Id, "N/A")

				'...ensure you can only see users below your level
				j = ""
				If vMemb_Level = 3 Then
					j = "<b> * </b>"
				ElseIf vMemb_Level = 4 Then
					j = "<b> ** </b>"
				ElseIf vMemb_Level = 5 Then
					j = "<b> *** </b>"
				End If     

				'...display if sponsored
				k = ""
				If vMemb_Sponsor > 0 Then
					k = "(<a href='User" & fGroup & ".asp?vMembNo=" & vMemb_Sponsor & "&vNext=" & vNext & "'>" & "<!--{{-->View Sponsor<!--}}-->" & "</a>)"
				End If


				vCurList = vCurList + 1
		%>
		<tr>
			<td style="white-space:nowrap">
				<%=fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", Replace(fCriteria(vMemb_Criteria), " + ", "<br>"))%>
				<%=fIf(vMemb_Group2 = 0 , "", "  [" & vMemb_Group2 & "]")%>
				<% If svMembLevel = 5 Then %><br><%=fRights()%><% End If %>
			</td>
			<td style="white-space:nowrap">
				<%=fLeft(vMemb_FirstName, 16) & " " & fLeft(vMemb_LastName, 16) & fIf(Len(vMemb_Organization) > 0, ", " & vMemb_Organization, "") & "<br>" & vMemb_Email & ""%> 
			</td>
			<td style="text-align:center; white-space:nowrap">
			<% If svMembLevel < 5 Then %>
			<%=fYN(vMemb_Active)%>
			<% Else %>      
			<a href="#" onclick="history(<%=vMemb_No%>)"><%=fYN(vMemb_Active)%></a>
			<% End If %>
			</td>
			<td style="white-space:nowrap">

				<% If svMembLevel = 5 Then
						'...note that vGoto uses ~5 and ~6 - these are so they don't get converted until it returns back to the start page - note add ~6!important to override custReturnUrl
						 vSource = "/V5/Default.asp~3vCust~2" & svCustId & "~1vId~2" & svMembId & "~1vGoto~2Default.asp~6vPage~5Users.asp~6!important"
						 If vGlobal = 1 Then 
				%>
						<a target="_top" href="/V5/Default.asp?vCust=<%=vGlobalCustId%>&vId=<%=vMemb_Id%>&vSource=<%=vSource%>"><img class="impersonate" src="../Images/Impersonate.png" /></a>
				<%   Else %>
						 <a target="_top" href="/V5/Default.asp?vCust=<%=svCustId%>&vId=<%=vMemb_Id%>&vSource=<%=vSource%>"><img class="impersonate" src="../Images/Impersonate.png" /></a>
				<%   End If %>
				<% End If %>


				<% If (svMembManager Or svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And InStr(vMemb_Id, vPasswordx) = 0 Then %>  
					<% =fIf(vGlobal = 1, "(" + vGlobalCustId + ")", "")%>
					<a href="<%=vEdit%>?vMembNo=<%=vMemb_No%>&vCustId=<%=vCustId%>&vNext=<%=vNext%>"><%=vMemb_Id%></a>
					<%=j%>
					<%=k%>
				<% Else %>
					********&nbsp; 
				<% End If %> 


				<% ' temp link to V8 for CAAM6914 
					If svCustId = "CAAM6914" Then
						Dim parms : parms = "membId=" & vMemb_Id
						Dim url   : url = "/v8?profile=MORT&parms=" & fBase64(parms)
				%>
					<span style="float:right; height: 17px;"><a target="_blank" href='<%=url%>'>V8</a></span>
				<%
						End If
				%>

				<br><span style="color:#3977B6"><%=vMemb_Memo%></span> 
			</td>
			<td style="text-align:center; white-space:nowrap"><%=fFormatDate(vMemb_FirstVisit)%><br><%=fFormatDate(vMemb_LastVisit)%> </td>
			<td style="text-align:center; white-space:nowrap"><%=fFormatDate(fIf(bG2CustEmail, fIf(IsDate(vMemb_Expires), vMemb_Expires, vCust_Expires), vMemb_Expires))%></td>
			<td style="text-align:center; white-space:nowrap"><%=vMemb_NoVisits & "<br>" & oRs("Memb_NoMins")%></td>   
			<% If bG2CustEmail Then %>
			<td style="text-align:center; white-space:nowrap">      
				<% If bG2MembEmail And Len(Trim(vMemb_Email)) > 0 And Len(Trim(vMemb_Programs)) > 0 Then %>
				<input onclick="resendEmails(<%=vMemb_No %>, '<%=svLang%>')" type="button" value="<%=bResend%>" name="bResend" class="button">      
				<% End If %>     
			</td>
			<% End If %>     

		</tr>
		<%
				oRs.MoveNext
				If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then Exit Do
			Loop
			Set oRs = Nothing
			sCloseDb
		%>
		
	</table>


	<br /><br />

	<form method="POST" action="Users_O.asp">
	 
		<div style="margin:auto; text-align:center;">
			<% If Len(vNext) > 0 Then %>
				<input type="button" onclick="location.href='<%=vNext%>'" value="<%=bReturn%>" name="bReturn" id="bReturn"class="button100">&emsp;
			<% End If %>
				<input type="button" onclick="location.href='Users.asp?vGlobal=<%=vGlobal%>&vNext=<%=vNext%>&vLearners=<%=vLearners%>&vCustId=<%=vCustId%>'" value="<%=bRestart%>" name="bRestart" class="button100">&emsp;
			<% If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then '...If next group, get next starting value %>
				<input type="hidden" name="vNext"          value="<%=vNext%>">
				<input type="hidden" name="vEdit"          value="<%=vEdit%>">
				<input type="hidden" name="vCustId"        value="<%=vCustId%>">
				<input type="hidden" name="vCurList"       value="<%=vCurList%>">
				<input type="hidden" name="vGlobal"        value="<%=vGlobal%>">
				<input type="hidden" name="vActive"        value="<%=vActive%>">
				<input type="hidden" name="vFind"          value="<%=vFind%>">
				<input type="hidden" name="vFindFirstName" value="<%=vFindFirstName%>">
				<input type="hidden" name="vFindLastName"  value="<%=vFindLastName%>">
				<input type="hidden" name="vFindEmail"     value="<%=vFindEmail%>">
				<input type="hidden" name="vFindMemo"      value="<%=vFindMemo%>">
				<input type="hidden" name="vFindCriteria"  value="<%=vFindCriteria%>">
				<input type="hidden" name="vLastValue"     value="<%=vMemb_LastName & vMemb_FirstName & vMemb_No%>">
				<input type="hidden" name="vFormat"        value="<%=vFormat%>">
				<input type="hidden" name="vLearners"      value="<%=vLearners%>">
				<input type="submit" name="bNext"          value="<%=bNext%>" class="button100">
			<% End If %>

				<br /><br />
				<% If vCust_Id = svCustId and vCust_InsertLearners Then %>
				<a href="<%=vEdit%>?vMembNo=0&vNext=<%=vNext%>&vCustId=<%=vCustId%>"><!--[[-->Add a Learner<!--]]--></a>
				<% End If %>

				<p><%=vCust_Id & "  (" & vCust_Title & ")"%>

		</div>

	</form>

	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>