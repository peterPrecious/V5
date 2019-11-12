<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
	Dim vAcct, vUser, vLevl, vNext, vOutp, vCurN, vSqlS, vStrD, vEndD, vActv, vProg, vMods, vPass, vLNam, vMemo, vGrou, vSave, bHuge, bZero, vRows, vFrom

	'...we get here either from the RTE_History.asp ("menu") or from somewhere like the info page
	vFrom = fDefault(Request("vFrom"), "menu") 
	vAcct = svCustAcctId
	vUser = svMembNo
	vActv = fDefault(Request("vActv"), "*")
	vLevl = svMembLevel
	vGrou = fDefault(Replace(Request("vGrou"), ", ", ","), "")
	vCurN = fDefault(Request("vCurN"), 0)
	vOutp = fDefault(Request("vOutp"), "O")
	vSave = fDefault(Request("vSave"), fIf(svMembLevel = 5, "y", "n"))
	vStrD = Request("vStrD") : If Not IsDate(vStrD) Then vStrD = ""
	vEndD = Request("vEndD") : If Not IsDate(vEndD) Then vEndD = ""
	vProg = fDefault(Replace(Request("vProg"), ", ", " "), "")
	vMods = fDefault(Replace(Request("vMods"), ", ", " "), "")
	vPass = fNoQuote(Request("vPass"))
	vLNam = fUnQuote(Request("vLNam"))
	vMemo = fUnQuote(Request("vMemo"))
	
	'...create the report if this page's form has been filled out or if we got here from the info page
	If Request.Form.Count > 0 Or vFrom <> "menu" Then
		'...check if the report will return 5000 or more records
		vRows = spHistoryCount (vAcct, vUser, fActv(vActv), vLevl, fNullValue(vGrou), fNullValue(vStrD), fNullValue(vEndD), fNullValue(vPass), fNullValue(vLNam), fNullValue(vMemo), fNullValue(vProg), fNullValue(vMods))

		bHuge = False : bZero = False
		If vRows > 5000 Then 
			bHuge = True 
		ElseIf vRows = 0 Then 
			bZero = True 
		Else 
			'...generate all the report records to table LogsR     
			spHistory vAcct, vUser, fActv(vActv), vLevl, fNullValue(vGrou), fNullValue(vStrD), fNullValue(vEndD), fNullValue(vPass), fNullValue(vLNam), fNullValue(vMemo), fNullValue(vProg), fNullValue(vMods)
			vNext = "RTE_History_" & vOutp  & ".asp" _
						& "?vActv=" & vActv _
						& "&vCurN=" & vCurN _
						& "&vStrD=" & vStrD _
						& "&vEndD=" & vEndD _
						& "&vProg=" & vProg _
						& "&vMods=" & vMods _
						& "&vPass=" & vPass _
						& "&vLNam=" & vLNam _
						& "&vMemo=" & vMemo _
						& "&vGrou=" & vGrou _ 
						& "&vSave=" & vSave _
						& "&vFrom=" & vFrom 
		End If
	 End If

		'...get a count of the rows that will be produced, similiar to below
	Function spHistoryCount (vAcct, vUser, vActv, vLevl, vGrou, vStrD, vEndD, vPass, vLNam, vMemo, vProg, vMods)

		Dim oRs
		sOpenCmd
		With oCmd
			.CommandText = "spHistoryCount"
			.Parameters.Append .CreateParameter("@Acct",	adChar, 		adParamInput,        4, vAcct)
			.Parameters.Append .CreateParameter("@User",  adInteger, 	adParamInput,         , vUser)
			.Parameters.Append .CreateParameter("@Actv",	adBoolean, 	adParamInput,         , vActv)
			.Parameters.Append .CreateParameter("@Levl",  adTinyInt, 	adParamInput,         , vLevl)
			.Parameters.Append .CreateParameter("@Grou",	adVarChar, 	adParamInput,     2000, vGrou)
			.Parameters.Append .CreateParameter("@StrD",  adDBDate, 	adParamInput,         , vStrD)
			.Parameters.Append .CreateParameter("@EndD",  adDBDate, 	adParamInput,         , vEndD)
			.Parameters.Append .CreateParameter("@Pass",  adVarChar, 	adParamInput,      128, vPass)
			.Parameters.Append .CreateParameter("@LNam",  adVarChar, 	adParamInput,       64, vLNam)
			.Parameters.Append .CreateParameter("@Memo",  adVarChar, 	adParamInput,       64, vMemo)
			.Parameters.Append .CreateParameter("@Prog",  adVarChar, 	adParamInput,     2000, vProg)
			.Parameters.Append .CreateParameter("@Mods",  adVarChar, 	adParamInput,     2000, vMods)
		End With

		Set oRs = oCmd.Execute()
		If oRs.Eof Then 
			spHistoryCount = 0
		Else
			spHistoryCount = oRs("reportRows")
		End If
		Set oRs = Nothing
		Set oCmd = Nothing
		sCloseDb

	End Function


	'...create a table of log items for this selection
	Function spHistory (vAcct, vUser, vActv, vLevl, vGrou, vStrD, vEndD, vPass, vLNam, vMemo, vProg, vMods)
		sOpenCmd
		With oCmd
			.CommandText = "spHistory"
			.Parameters.Append .CreateParameter("@Acct",	adChar, 		adParamInput,        4, vAcct)
			.Parameters.Append .CreateParameter("@User",  adInteger, 	adParamInput,         , vUser)
			.Parameters.Append .CreateParameter("@Actv",	adBoolean, 	adParamInput,         , vActv)
			.Parameters.Append .CreateParameter("@Levl",  adTinyInt, 	adParamInput,         , vLevl)
			.Parameters.Append .CreateParameter("@Grou",	adVarChar, 	adParamInput,     2000, vGrou)
			.Parameters.Append .CreateParameter("@StrD",  adDBDate, 	adParamInput,         , vStrD)
			.Parameters.Append .CreateParameter("@EndD",  adDBDate, 	adParamInput,         , vEndD)
			.Parameters.Append .CreateParameter("@Pass",  adVarChar, 	adParamInput,      128, vPass)
			.Parameters.Append .CreateParameter("@LNam",  adVarChar, 	adParamInput,       64, vLNam)
			.Parameters.Append .CreateParameter("@Memo",  adVarChar, 	adParamInput,       64, vMemo)
			.Parameters.Append .CreateParameter("@Prog",  adVarChar, 	adParamInput,     2000, vProg)
			.Parameters.Append .CreateParameter("@Mods",  adVarChar, 	adParamInput,     2000, vMods)
		End With
		oCmd.Execute()
		Set oCmd = Nothing
		sCloseDb
	End Function


	'...Actv (Memb_Active) is either 0(false), 1(true) or *(both)
	'   set * to null so we can pass over a boolean bit
	Function fActv (i)
		fActv = fIf(i = "*", Null, i)
	End Function


	Function fGroups 
		Dim vSelected
		fGroups   = "" : vGroupCnt = 0
		If svMembLevel > 3 Then 
			vSql = "SELECT Crit_Id FROM Crit WHERE Crit_AcctId = '" & svCustAcctId & "' ORDER BY Crit_Id"
		Else
			vSql = " SELECT Crit_Id FROM"_
					 & "   Memb INNER JOIN"_ 
					 & "   Memb_Crit ON Memb.Memb_No = Memb_Crit.Memb_Crit_MembNo INNER JOIN"_
					 & "   Crit ON Memb_Crit.Memb_Crit_CritNo = Crit.Crit_No"_
					 & " WHERE"_
					 & "   Memb.Memb_No = " & svMembNo 
		End If
		sOpenDb    
		Set oRs = oDb.Execute(vSql)  
		fGroups = vbCrLf 
		Do While Not oRs.Eof 
			vCrit_Id = oRs("Crit_Id") // change any embedded commas to $ to avoid sql function conflict
			vSelected = fIf(Instr(vGrou, vCrit_Id) > 0, " Selected ", "")
			fGroups = fGroups & "<option" & vSelected & " value='" & Replace(vCrit_Id, ",", "$") & "'>" & vCrit_Id & "</option>" & vbCrLf
			vGroupCnt = vGroupCnt + 1
			oRs.MoveNext
		Loop
		Set oRs = Nothing
		sCloseDb    
		vGroupCnt = fIf(vGroupCnt > 50, 12, fIf(vGroupCnt > 8, 8, vGroupCnt))
	End Function
 
%>

<html>

<head>
	<title>RTE_History_F</title>

<!--	
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">-->
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"> 

	<link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
	<script src="/V5/Inc/jQuery.js"></script>
	<script src="/V5/Inc/jQueryC.js"></script>
	<script src="/V5/Inc/Functions.js"></script>
	<% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
	<script src="/V5/Inc/Calendar.js"></script>
	<script>
			// save parms into cookies for future visits 
			// if we have a vNext then go to next page
			$(document).ready ( 
				function () {
//        debugger;

					$("#divLoader").hide();

					var vSave = "<%=vSave%>";
					if (vSave == "n") $(".advanced").hide();

					var vProg = "<%=vProg%>";
					if (vProg.length > 0) getList("PI", "<%=vProg%>");

					var vMods = "<%=vMods%>";
					if (vMods.length > 0) getList("MI", "<%=vMods%>");
			
					var vNext = "<%=vNext%>";
			
			
					var options = { path: '/', expires: 365 };	
					if (vNext.length > 0) {    
						$.cookie("History_<%=svCustId%>_vActv", "<%=vActv%>", options);
						$.cookie("History_<%=svCustId%>_vCurN", "<%=vCurN%>", options);
						$.cookie("History_<%=svCustId%>_vOutp", "<%=vOutp%>", options);
						$.cookie("History_<%=svCustId%>_vStrD", "<%=vStrD%>", options);
						$.cookie("History_<%=svCustId%>_vEndD", "<%=vEndD%>", options);
						$.cookie("History_<%=svCustId%>_vProg", "<%=vProg%>", options);
						$.cookie("History_<%=svCustId%>_vMods", "<%=vMods%>", options);
						$.cookie("History_<%=svCustId%>_vPass", "<%=vPass%>", options);
						$.cookie("History_<%=svCustId%>_vLNam", "<%=vLNam%>", options);
						$.cookie("History_<%=svCustId%>_vGrou", "<%=vGrou%>", options);
						$.cookie("History_<%=svCustId%>_vSave", "<%=vSave%>", options);

						// hide the form - specifically for access via Home page but NOT for the Excel Report
						if (vNext.substring(0,13) != "RTE_History_X") {
							$("#divLoader").show();
							$("#divBody").hide();
						}

						location.href = vNext;
					};

					// high light filled fields	
					hiliteFields();

				}
			);
			<%
			
			Function jValue(i)
				If IsNull(i) Then 
					jValue = "''"
				Else
					jValue = "'" & i & "'"
				End If
			End Function

			 %>

			// highlights and fields with previous values
			function hiliteFields(){

				$("input:text").each(function() {
					if ($(this).val().length > 0) { $(this).css("background-color", "lightyellow"); };
				});

				if ($("#progList option:selected").length > 0) {
					$("#progList").css("background-color", "lightyellow"); 
				};

				if ($("#modsList option:selected").length > 0) {
					$("#modsList").css("background-color", "lightyellow"); 
				};

				if ($("#critList option:selected").length > 0) {
					$("#critList").css("background-color", "lightyellow"); 
				};


			}    

			// hide div features
			function advHide() {
				document.styleSheets("adv").cssText = ".adv {MARGIN:0px;DISPLAY:none}";
				document.styleSheets("alt").cssText = "";
				document.getElementById("vFeat").value = "N";
			}
	
			// show div features
			function advShow() {
				document.styleSheets("adv").cssText = "";
				document.styleSheets("alt").cssText = ".alt {MARGIN:0px;DISPLAY:none}";
				document.getElementById("vFeat").value = "Y";
			}
	
			function dateOk(id) {
				if (isDate(id.value)) {
					return (true);
				} else {
					alert("Please enter a valid date.");
					id.focus();
					return (false)
				}    
			}
	
			function dateBefore(prvId, curId) {
				if (Date.parse(prvId.value) > Date.parse(document.getElementById(curId).value)) {
					alert("Starting date is after your Ending Date.");
					prvId.focus();
					return (false)
				} else {
					return (true);
				}    
			}
	
			function dateLater(curId, prvId) {
				if (Date.parse(curId.value) < Date.parse(document.getElementById(prvId).value)) {
					alert("Ending date is before your Starting Date.");
					curId.focus();
					return (false)
				} else {
					return (true);
				}    
			}


			function getList(param, selected) {
				if (param == "PI" || param == "PT") {
					var vParam = "selected=<%=vProg%>&param=" + param;
					var vWs    = WebService("RTE_History_ws.asp", vParam);
					$("#cellProgList").html(vWs);
				} else if (param =="MI" || param == "MT") {
					var vParam = "selected=<%=vMods%>&param=" + param;
					var vWs    = WebService("RTE_History_ws.asp", vParam);
					$("#cellModsList").html(vWs);
				}
			}

	</script>
	<script id="preProcess">
		// use this when launching a form that requires pre processing (it must be setup with the precise divs)
		function preProcess(hide, show) {
	
			$("#huge").hide();

			<% If svMembLevel = 3 Then %>  
			// if FAC with more than one Group option, demand a selection 
			if ($("#critList")[0] != undefined && $("#critList")[0].length > 1 && $("#critList")[0].selectedIndex == -1) {
				alert("/*--{[--*/Please select at least one Group./*--]}--*/");
				return false;
			}
			<% End If  %>  

			var msg = "/*--{[--*/Processing.../*--]}--*/";
			$(hide).hide();
			$(show).text(msg);
			$("form").submit();

			return true;
		}
	</script>
</head>

<body>

	<% Server.Execute vShellHi %>

	<!-- this is used when accessing this report from the Info page, preventing display of the form -->
	<div id="divLoader" style="margin: 30px; text-align: center;">
		<h1><!--[[-->Generating your Learner Report Card<!--]]--></h1>
		<img src="../Images/Common/ProgressBar.gif" />
	</div>

	<!-- we want to hide this div when accessing this report from the Info page, preventing display of the form -->
	<div id="divBody">

		<h1><!--[[-->Learner Report Card<!--]]--></h1>

		<div style="text-align: center; font-weight: 300; font-size: 12px;">[<a href="#" class="green" onclick="$('.advanced').toggle()"><!--[[-->Advanced Search On/Off<!--]]--></a>]</div>

		<div class="advanced" style="margin-bottom:20px;">
			<h2 style="text-align: left">
				<!--[[-->This shows the learning activities for all selected learners sorted by Learner, Program ID and Module ID.<!--]]-->
				<!--[[-->If requested below, your selection values will be saved and shown in future visits with a pale yellow background<!--]]-->.
			</h2>
		</div>

		<% If bHuge Or bZero Then %>
		<div id="huge" style="border: 1px solid red; width: 500px; margin: auto; padding: 10px; background-color: lightyellow">
			<% If bHuge Then %>
			<!--[[-->Your request will return more than 5000 learner records.<br />Please narrow your selection.<!--]]-->
			<% Else %>
			<!--[[-->Your request will not return any learner records.<br />Please broaden your selection.<!--]]-->
			<% End If %>
		</div>
		<% End If %>


		<form method="POST" action="RTE_History_F.asp" id="fHistory" accept-charset="utf-8">
			<input type="hidden" value="<%=Request("vParmNo")%>" name="vParmNo">
			<input type="hidden" name="vFeat" id="vFeat" value="Y">
			<table class="table">

				<tr>
					<td colspan="3" class="c3">
						<!--[[-->Show learner activities<!--]]-->...</td>
				</tr>
				<tr class="advanced">
					<th>
						<!--[[-->for learners that are<!--]]-->:</th>
					<td colspan="2">
						<input type="radio" name="vActv" value="1" <%=fcheck("1", vactv)%>><!--[[-->Active<!--]]-->
						<input type="radio" name="vActv" value="0" <%=fcheck("0", vactv)%>><!--[[-->Inactive<!--]]-->
						<input type="radio" name="vActv" value="*" <%=fcheck("*", vactv)%>><!--[[-->Both active and inactive<!--]]--></td>
				</tr>
				<tr>
					<th>
						<!--[[-->between<!--]]-->:</th>
					<td>
						<input type="text" name="vStrD" id="vStrD" size="12" value="<%=vStrD%>" style="text-align: center">
						<a title="<!--[[-->Start at the beginning...>!--]]-->" class="debug" onclick="fillField('vStrD', '')" href="#">&#937;</a>
						<a href="javascript:show_calendar('vStrD','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>', 'MONTH DD YYYY');"><img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a>
						<!--[[-->and<!--]]-->&nbsp;&nbsp;&nbsp;&nbsp; 
						<input type="text" name="vEndD" id="vEndD" size="12" value="<%=vEndD%>" style="text-align: center">
						<a title="<!--[[-->End at today's date...<!--]]--" class="debug" onclick="fillField('vEndD',  '')" href="#">&#937;</a>
						<a href="javascript:show_calendar('vEndD','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>');">
							<img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a><br />
						<% p1 = fFormatSQLDate(Now()) %>
						<!--[[-->Enter dates in English format (Mmm d, yyyy), ie ^1, or by using the calendar icon.<br />Use &#937; to remove a date filter which can be very resource intensive. Note: invalid dates are ignored.<!--]]-->
					</td>
				</tr>
				<tr id="progList" class="advanced">
					<th>
						<!--[[-->for Programs<!--]]-->(<!--[[-->sort by<!--]]-->
						<a href="#" onclick="getList(&quot;PI&quot;, &quot;<%=vProg%>&quot;)"><!--[[-->ID<!--]]--></a> | <a href="#" onclick="getList(&quot;PT&quot;, &quot;<%=vProg%>&quot;)"><!--[[-->Title<!--]]--></a>) :
					</th>
					<td id="cellProgList" colspan="2">
						<!--[[-->To list specific Programs, select available Programs by&nbsp; Id or Title at left.<!--]]--><br />
						<!--[[-->Note: this list only shows Programs that have been accessed.<!--]]-->
					</td>
				</tr>
				<tr id="modsList" class="advanced">
					<th>
						<!--[[-->for Modules<!--]]-->
						(<!--[[-->sort by<!--]]-->
						<a href="#" onclick="getList(&quot;MI&quot;, &quot;<%=vMods%>&quot;)"><!--[[-->ID<!--]]--></a> | <a href="#" onclick="getList(&quot;MT&quot;, &quot;<%=vMods%>&quot;)"> <!--[[-->Title<!--]]--></a>) :
					</th>
					<td id="cellModsList" colspan="2">
						<!--[[-->To list specific Modules, select available Modules by&nbsp; Id or Title at left.<!--]]--><br />
						<!--[[-->Note: this list only shows Modules that have been accessed.<!--]]-->
					</td>
				</tr>
				<tr>
					<th>
						<!--[[-->whose<!--]]-->&nbsp;<%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%>&nbsp;<!--[[-->contains<!--]]-->:</th>
					<td colspan="2">
						<input type="text" name="vPass" id="vPass" size="20" value="<%=vPass%>"><br>
						<!--[[-->ie &#39;RA&#39; will list all learners whose<!--]]-->&nbsp;<%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%>&nbsp;<!--[[-->contains &#39;RA&#39;, like &#39;SARAH011&#39; or &#39;sarah011&#39;<!--]]--></td>
				</tr>
				<tr>
					<th>
						<!--[[-->whose Last Name starts with<!--]]-->:</th>
					<td colspan="2">
						<input type="text" name="vLNam" size="20" value="<%=vLNam%>"><br>
						<!--[[-->ie &#39;Smi&#39; will include all learners whose Last Name starts with &#39;Smi&#39;<!--]]--></td>
				</tr>
				<tr class="advanced">
					<th>
						<!--[[-->whose Memo fields contains<!--]]-->:</th>
					<td colspan="2">
						<input type="text" name="vMemo" size="20" value="<%=vMemo%>"><br>
						<!--[[-->ie &#39;ho&#39; will include all learners whose Memo Field contains &#39;ho&#39;<!--]]--></td>
				</tr>
				<% 
						Dim vGroupCnt : vGroupCnt = 0
						i = fGroups
						If vGroupCnt > 1 Then
				%>
				<tr>
					<th>
						<!--[[-->from Group(s)<!--]]-->:</th>
					<td colspan="2">
						<% If svMembLevel = 3 Then %>
						<!--[[-->Select Group.&nbsp; Use Ctrl+Enter for multiple selections<!--]]-->
						<% Else %>
						<!--[[-->Leave unselected for ALL.&nbsp; Use Ctrl+Enter for multiple selections<!--]]-->
						<% End If %>
						<br /><br />
						<select id="critList" size="<%=vGroupCnt%>" name="vGrou" multiple><%=i%></select>
					</td>
				</tr>
				<%  
					Else 
				%>
				<input type="hidden" name="vGrou" value="<%=fIf(svMembCriteria = "0", "", Replace(fCriteria (svMembCriteria), ",", "$"))%>">
				<tr class="advanced">
					<th>
						<!--[[-->from Group<!--]]-->:
					</th>
					<td colspan="2">
						<b><%=fCriteria (svMembCriteria)%></b>
					</td>
				</tr>
				<% 
					End If 
				%>
				<tr>
					<th>
						<!--[[-->Output format<!--]]-->:
					</th>
					<td colspan="2">
						<input type="radio" name="vOutp" value="O" <%=fcheck("o", voutp)%>>
						<!--[[-->Online<!--]]-->
						<input type="radio" name="vOutp" value="X" <%=fcheck("x", voutp)%>><!--[[-->Excel<!--]]-->
					</td>
				</tr>
				<tr class="advanced">
					<th>
						<!--[[-->Save selections<!--]]-->:</th>
					<td colspan="2">
						<input type="radio" name="vSave" value="y" <%=fcheck("y", vsave)%>>
						<!--[[-->Yes<!--]]-->
						<input type="radio" name="vSave" value="n" <%=fcheck("n", vsave)%>>
						<!--[[-->No<!--]]-->
						<br />
						<!--[[-->If &#39;Yes&#39; then the above selections will save for your next visit and the Advanced Search fields will be shown.<!--]]-->
						<% If svMembLevel = 5 Then %>
							&nbsp; Defaults to &#39;Yes&#39; for Administrators.
						<% End If %>
					</td>
				</tr>
				<tr>
					<td colspan="3" style="text-align: center; height: 100px; vertical-align: middle;">
						<!-- this submit button must be setup exactly as follows to render the proper message -->
						<div id="divShow" class="buttonAlert">
							<div id="divHide">
								<input type="button" id="Start" onclick="preProcess('#divHide', '#divShow')" class="button100" value="<%=bNext%>" name="bStart"></div>
						</div>
					</td>
				</tr>
			</table>
		</form>
	</div>
	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
