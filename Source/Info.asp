<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<!--#include virtual = "V5/Inc/Document.asp"-->
<!--#include virtual = "V5/Inc/Base64.asp"-->

<% 
	Dim vIntro, vParagraph
	vIntro = "<!--{{-->Welcome<!--}}-->" 
	If svSecure And Len(svMembFirstName) > 1 Then 
		vIntro = vIntro & " " & svMembFirstName
	End If 

	'...determine what tabs are available and sponsored learners
	sGetCust svCustId

	'...insert these into tab names
	p1 = fTabName(vCust_Tab1Name, "<!--{{-->Info Page<!--}}-->")
	p2 = fTabName(vCust_Tab2Name, "<!--{{-->My Learning<!--}}-->")
	p3 = fTabName(vCust_Tab3Name, "<!--{{-->My Content<!--}}-->")
	p4 = fTabName(vCust_Tab4Name, "<!--{{-->My Development<!--}}-->")
	p5 = fTabName(vCust_Tab5Name, "<!--{{-->More Content<!--}}-->")
	p6 = fTabName(vCust_Tab6Name, "<!--{{-->Administration<!--}}-->")
	p7 = fTabName(vCust_Tab7Name, "<!--{{-->Sign Off<!--}}-->")

	p8 = "<!--{{-->Id<!--}}-->"
	p9 = "<!--{{-->Password<!--}}-->"
	
	Function fTabName(vTabName, vOriginal)
		Dim aTabName
		If IsNull(vTabName) Or Len(vTabName) = 0 Then  
			fTabName = vOriginal
		Else
			aTabName = Split(vTabName, "|")
			If svLang = "EN" And Ubound(aTabName) >= 0 Then 
				If Len(aTabName(0)) > 0 Then fTabName = Trim(Left(aTabName(0) & Space(20), 20))
			End If
			If svLang = "FR" And Ubound(aTabName) >= 1 Then
				If Len(aTabName(1)) > 0 Then fTabName = Trim(Left(aTabName(1) & Space(20), 20))
			End If
			If svLang = "ES" And Ubound(aTabName) >= 2 Then
				If Len(aTabName(2)) > 0 Then fTabName = Trim(Left(aTabName(2) & Space(20), 20))
			End If
		End If
	End Function

		'...determine if sponsored learner
	sGetMemb svMembNo

	'...get to see if pc or mac (for bookmarking)
	sGetQueryString

	If Request.Form.Count > 0 Then 
		vMemb_No        = svMembNo
		vMemb_Pwd       = Ucase(Request.Form("vMemb_Pwd"))
		vMemb_FirstName = Request.Form("vMemb_FirstName")
		vMemb_LastName  = Request.Form("vMemb_LastName") 
		vMemb_Email     = Request.Form("vMemb_Email")
		vMemb_VuNews    = fDefault(Request.Form("vMemb_VuNews"), 0)
		sUpdateMemb_Profile 
 '  Response.Redirect "#MyProfile"
	End If

	'...this will put either support@vubiz.com or the customers email address on the Contact Us link at the bottom
	Function fContactUs
		Dim vEmail, vText
		If svCustEmail = "none" Then
			fContactUs = ""
		Else
			vEmail = fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")

'     Select Case svLang
'       Case "FR" : vText = "Communiquez avec nous"
'       Case "ES" : vText = "P&#243;ngase en contacto con nosotros"
'       Case Else : vText = "Contact Us"
'     End Select
'			fContactUs = "<a href='mailto:" & vEmail & "?subject=" & svCustId & " Issue'>" & vText & " (" & vEmail & ")</a>"

			'...added ERGP link to web override on Apr 19, 2018
			If Left(svCustId, 4) = "ERGP" Then
				fContactUs = "<a target='_blank' href='http://eg.onlinelearninghelp.com'>eg.onlinelearninghelp.com</a>"
			Else
				fContactUs = "<a href='mailto:" & vEmail & "?subject=" & svCustId & " Issue'>" & vEmail & "</a>"
			End If

		End If
	End Function

	'...used to pass to browser budgie (added Jan 2, 2018)
	Function fEmail
	 fEmail = fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")
	End Function

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
	<title>::Info</title>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
	<script type="text/javascript" src="/V5/Inc/jQuery.js"></script>
	<link href="/V5/Inc/Vubi2.css" rel="stylesheet" />
	<script type="text/javascript" src="/V5/Inc/Functions.js"></script>
	<% If vRightClickOff Then %><script type="text/javascript" src="/V5/Inc/RightClick.js"></script><% End If %>
	<script type="text/jscript">

	<% 
		Dim vAlert

		'...see if this user is flagged as online
'   Session("Breach") = True
		If svMembLevel < 5 And Session("Breach") Then 
'   If Session("Breach") Then 
			Session("Breach") = False
			vAlert = "<!--{{-->Warning! Your status shows that you are already online. This can happen if you did not 'Sign Off' after your last session, are signed in with more than one browser window or someone else is accessing your account! This can, at a minimum, cause loss of data integrity.\n\n[Breach Status ID:<!--}}-->" & svCustAcctId & "-" & Year(Now) & Month(Now) & Day(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & "]."  
	%>
			alert("<%=vAlert%>");
	<% 
		End If

		'...see if using staging
		If (Not svMembInternal) And svServer = "staging.vubiz.com" And Len(Session("Staging")) = 0 Then             
			Session("Staging") = True
			vAlert = "<!--{{-->Warning! You are on the Vubiz STAGING server which is \nused for functional review NOT real time learning.\nNo records are maintained!\n\n If you intented to use our real time service\n please visit //vubiz.com.<!--}}-->"  
	%>
			alert("<%=vAlert%>");
	<% 
		End If
	%>

		// render popup status (twig to show "Yes" for enabled rather than for Blocker On)
		$(function () {
			//$("#popupStatus").html(parent.popupBlockerOn ? jYN("n", "<%=svLang%>") : jYN("y", "<%=svLang%>"));

			$(".toggleAgent").on("click", function () {
				$(".browserAgent").toggle().html("<br />(" + navigator.userAgent.toLowerCase() + ")");
			});
		});

	</script>
	<style type="text/css">
		.greenLarge { color: green; font-size: 1.35em; margin-top: 20px; }
		.toggleAgent { cursor: pointer; color: orange; }
		.c2 { margin-top: 30px; }
	</style>
</head>

<body>

	<% Server.Execute vShellHi %>

	<div id="info" style="width: 800px; margin: auto; xtext-align: left;">

		<% 
			'...Show|Hide the alert at admin's request via url below: //vubiz.com/v5/alert.asp then set password(21122112 and alert:   vAlert=y | vAlert=n
			'...also ADD same message to TabsLive!

			' And svMembLevel = 5 

			If Application("Alert") = "y" Then
		%>
		<p class="c2">::&ensp;<!--[[-->NOTICE!<!--]]--></p>

		<div style="background-color: yellow; padding: 10px;">
			<% If svLang = "FR" Then %>
				 Ce service sera interrompu pour fin d’amélioration et ne sera pas disponible le samedi 23 mai de 6h00 à 9h00 HNE. Nous nous excusons des inconvénients causés.
			<% ElseIf svLang = "ES" Then %>        
				 Este servicio estará en mantenimiento de rutina y no estará disponible el sábado 23 de mayo 06 a.m.-09 a.m. EST. Nos disculpamos por cualquier inconveniente.
			<% Else %>
				 This service will be undergoing routine maintenance and will not be available on Saturday May 23rd from 6 am to 9 am EST. We apologize for any inconvenience.
			<% End If %>
		</div>

		<% End If %>

		<p class="c2">::&nbsp;&nbsp;<%=Trim(vIntro)%></p>

		<% If vCust_Tab2 Then %>
		<!--[[-->Click on the <b>^2</b> tab above to access your programs.<!--]]-->&nbsp;
		<% End If %>

		<% If vCust_Tab3 Then %>
		<!--[[-->Click on the <b>^3</b> tab above to access your free or purchased programs.<!--]]-->&nbsp;
		<% End If %>

		<% If vCust_Tab5 Then %>
		<!--[[-->To purchase e-learning programs, click <b>^5</b> to complete a secure e-commerce process.&nbsp;&nbsp; Any programs purchased will then appear under the <b>^3</b> tab above.<!--]]-->&nbsp;
		<% End If %>

		<p class="c2">::&nbsp; <!--[[-->Browser Readiness Test<!--]]--></p>
		<a target="_blank" href="/browser?email=<%=fEmail%>&lang=<%=svLang%>"><!--[[-->Click here<!--]]--></a>&nbsp;<!--[[-->to confirm your browser is configured properly for this service.<!--]]-->

<!--
		<% If svLang = "EN" Then %>
		<p class="c2">::&nbsp; Help using this service</p>
		<a target="_blank" href="../Public/21_FAQ.asp">Click here</a>&nbsp;if you have any questions about how this service works.
		<% End If %>

		<% If svLang = "FR" Then %>
		<p class="c2">::&nbsp; Problèmes liés aux navigateurs?</p>
		<p><a target="_blank" href="../Public/BrowserIssues_FR.htm">Cliquez ici</a> pour options de réglage de votre navigateur Web</p>
		<% End If %>
-->



		<% If svLang = "FR" Then %>
		<p class="c2">::&nbsp; Communiquez avec nous</p>
		<span style="text-align: center; margin-top: 20px;"><%=fContactUs%></span>
		<% Else %>
		<p class="c2">::&nbsp; Contact Us</p>
		<span style="text-align: center; margin-top: 20px;"><%=fContactUs%></span>
		<% End If%>



		<%  
			Dim parms, url
			parms = "" : url = ""

			If svCustId = "CAAM3001" And svMembLevel = 2 Then
				parms = "custId=CAAM3001&membNo=" & fMembNo("3001", svMembId)
				url = "/v8?profile=mpc&parms=" & fBase64(parms)
			End If

'     If svCustId = "CCHS2544" And svMembId = "279488-CDFC" Then
'       parms = "custId=CCHS2544&membNo=" & fMembNo("2544", svMembId)

			If svCustId = "CCHS2544" Then
				parms = "custId=CCHS2544&membNo=" & svMembNo
				url = "/v8?profile=ccohsDemo&parms=" & fBase64(parms)
			End If

			If LEFT(svCustId, 4) = "CCHS" Then
				parms = "custId=" & svCustId & "&membNo=" & svMembNo
				url = "/v8?profile=ccohsDemo&parms=" & fBase64(parms)
			End If



			If svCustId = "IAPA2859" And svMembLevel = 2  Then
				parms = "custId=IAPA2859&membNo=" & fMembNo("2859", svMembId)
				url = "/v8?profile=wsps&parms=" & fBase64(parms)
			End If

			If parms <> "" Then
		%>
		<p class="c2">::&nbsp; New Mobile/Touch Friendly Interface (V8)</p>
		<a target="_blank" href="<%=url%>">Click here</a>&nbsp;to access your Content
		<% 
			End If 
		%>








		<p class="c2">::&nbsp;&nbsp;<!--[[-->Important<!--]]--></p>
		<!--[[-->Please do NOT <b>Bookmark</b> or <b>Add to Favorites</b> the address that appears in your web browser when you are logged in. You must login with your user credentials each time you enter this service. Please click the Sign Off tab at the end of each visit.<!--]]-->

		<% If vCust_MaxSponsor > 0 And vMemb_Sponsor = 0 Then '...if accounts allow sponsors then ensure this link is not for a sponsored learner %>
		<p class="c2">::&nbsp;&nbsp;<!--[[-->My Sponsored Learners<!--]]--></p>
		<!--[[-->If you would like to offer members of your organization access to your content, click ...<!--]]--><a href="Sponsors.asp"><u><!--[[-->Sponsored Learners.<!--]]--></u></a>
		<% End If %>

		<% If vCust_Scheduler Then %>
		<p class="c2">::&nbsp;&nbsp;<!--[[-->My Scheduler<!--]]--></p>
		<p class="c2">
			<a href="Scheduler.asp"><u>
				<!--[[-->Click here<!--]]--></u></a><!--[[-->to view your calendar.<!--]]-->
		</p>
		<% End If %>

		<% If vCust_InfoEditProfile Then %>
		<p class="c2"><a <%=fStatX%> name="MyProfile"></a>::&nbsp;<!--[[-->My Profile<!--]]--></p>
		<!--[[-->Enter/edit your name and email address below. Your name will then appear on any certificates issued for successful completion of assessments or exams.<!--]]-->
		<script type="text/javascript">
			function Validate(theForm) 
			{

				//  only check password if used by this memeber, else ignore
				if (theForm.vPassword.value == "check") 
				{

					if (theForm.vMemb_Pwd.value == "")
					{
						alert("Please enter a value for the \"Password\" field.");
						theForm.vMemb_Pwd.focus();
						return (false);
					}
			
					if (theForm.vMemb_Pwd.value.length < 4)
					{
						alert("Please enter at least 4 characters in the \"Password\" field.");
						theForm.vMemb_Pwd.focus();
						return (false);
					}
			
					if (theForm.vMemb_Pwd.value.length > 64)
					{
						alert("Please enter at most 64 characters in the \"Password\" field.");
						theForm.vMemb_Pwd.focus();
						return (false);
					}
			
					var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-_-@.";
					var checkStr = theForm.vMemb_Pwd.value;
					var allValid = true;
					for (i = 0;  i < checkStr.length;  i++)
					{
						ch = checkStr.charAt(i);
						for (j = 0;  j < checkOK.length;  j++)
							if (ch == checkOK.charAt(j))
								break;
						if (j == checkOK.length)
						{
							allValid = false;
							break;
						}
					}
					if (!allValid)
					{
						alert("Please enter only letter, digit and \"_-@.\" characters in the \"Password\" field.");
						theForm.vMemb_Pwd.focus();
						return (false);
					}
				
				}
				return (true);
			}
		</script>

		<div style="text-align: center">
			<form method="POST" action="info.asp" onsubmit="return Validate(this)" name="fHome">
				<input type="hidden" name="fProfile" value="Y" />
				<table class="table" style="width: 350px; margin: 20px auto auto auto;">
					<tr>
						<th>
							<!--[[-->First Name<!--]]-->:</th>
						<td>
							<% If Request.QueryString("vAction") = "edit" Then %>
							<input type="text" name="vMemb_FirstName" size="19" value="<%=svMembFirstName%>" maxlength="32" />
							<% Else %>
							<%=svMembFirstName%>
							<% End If %> 
						</td>
					</tr>
					<tr>
						<th>
							<!--[[-->Last Name<!--]]-->:</th>
						<td>
							<% If Request.QueryString("vAction") = "edit" Then %><input type="text" name="vMemb_LastName" size="19" value="<%=svMembLastName%>" maxlength="64">
							<% Else %>
							<%=svMembLastName%>
							<% End If %> 
						</td>
					</tr>

					<% If vCust_Pwd And (svMembLevel = 2 Or svCustId = "CFIB1660") Then %>
					<tr>
						<th>
							<!--[[-->Password<!--]]-->:</th>
						<td>
							<% If Request.QueryString("vAction") = "edit" Then %>
							<input type="password" name="vMemb_Pwd" size="19" value="<%=svMembPwd%>" maxlength="64" />
							<% Else %>
							<%="****************"%>
							<% End If %> 
						</td>
					</tr>
					<input type="hidden" name="vPassword" value="check" />
					<% Else %>
					<input type="hidden" name="vPassword" value="ignore" />
					<% End If %>

					<tr>
						<th>
							<!--[[-->Email Address<!--]]-->:</th>
						<td>
							<% If Request.QueryString("vAction") = "edit" Then %>
							<input type="text" name="vMemb_Email" size="19" value="<%=svMembEmail%>" />
							<% Else %>
							<%=fDefault(svMembEmail, "...<i><font color='#FF0000'>[none]")%>
							<% End If %> 
						</td>
					</tr>

					<% If svLang = "EN" And vCust_VuNews Then %>
					<tr>
						<th>Send vuNews <b><a href="javascript:toggle('Div_VuNews');">?</a></b></th>
						<td>
							<% If Request.QueryString("vAction") = "edit" Then %>
							<input type="radio" value="1" name="vMemb_VuNews" <%=fcheck(fsqlboolean(vMemb_VuNews), 1)%> />Yes&nbsp; 
							<input type="radio" value="0" name="vMemb_VuNews" <%=fcheck(fsqlboolean(vMemb_VuNews), 0)%> />No&nbsp;
							<% Else %>
							<%=fIf(vMemb_VuNews, "Yes", "No")%>
							<% End If %>
							<%=f5%>
						</td>
					</tr>
					<tr>
						<th colspan="2">
							<div align="center" id="Div_VuNews" class="div">
								<table class="table">
									<tr>
										<td>vuNews is an online newsletter that we publish quarterly.&nbsp; If interested, click Edit, select Yes to Send vuNews and your email address will be added to our distribution list.&nbsp; You can discontinue the newsletter at any time.<h6>Be assured, your profile will NEVER be released to any third parties.</h6>
											<p style="text-align: center">Thank you!</p>
										</td>
									</tr>
								</table>
							</div>
						</th>
					</tr>
					<% End If %>

					<tr>
						<td colspan="2" style="height: 30px; text-align: right;">

							<% If Request.QueryString("vAction") = "edit" Then %>
							<input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button" />
							<% Else %>
							<input onclick="location.href = 'Info.asp?vAction=edit#MyProfile'" type="button" value="<%=bEdit%>" name="bEdit" class="button" />
							<% End If %>

						</td>
					</tr>

				</table>
			</form>
		</div>

		<% End If %>
		<p class="c2">::&nbsp;&nbsp;<!--[[-->My Status<!--]]--></p>

		<% 
'     If vCust_Level = 4 OR svMembLevel = 5 Then 
		%>

		<!--
			<a href="LearnerReportCard2.asp?vMemb_No=<%=svMembNo%>&vInfoPage=y">Click here</a>&nbsp;for your Report Card.<br>
		-->

		<% 
'     End If
		%>


		<% 
'     If vCust_Level = 2 OR svMembLevel = 5 Then 
		%>
		<a href="RTE_History_F.asp?vPass=<%=svMembId%>&vFrom=<%=svPage%>">
			<!--[[-->Click here<!--]]--></a>&nbsp;<!--[[-->for your Report Card<!--]]-->.
		<% 
'     End If 
		%>


		<%
			'...data is passed in as svBrowser and determined in the initial default.asp
			Dim aTools, vTouch, vHTML5, vFlash, vCookies, vPopup, vEcom
			If Len(svBrowser) > 0 Then
				svBrowser = svBrowser & "     "
				aTools = Split(Ucase(svBrowser), "|")
				vTouch   = fYN (aTools(0))
				vBrowser = aTools(1)
				vHTML5   = fYN (aTools(2)) 
				vFlash   = fIf(aTools(3) = "0", fYN(0), aTools(3)) 
				vCookies = fYN (aTools(4)) 
				vPopup   = fYN (aTools(5)) 
				vEcom    = aTools(6) 
				If vEcom = "Y" Or vEcom = "N" Then 
					vEcom = fYN (vEcom) 
				Else 
					vEcom = "n/a" 
				End If
			End If
		%>

		<div style="text-align: center">
			<table class="table" style="width: 350px; margin: 20px auto auto auto;">
				<tr>
					<th><!--[[-->Customer Id<!--]]-->:</th>
					<td width="50%"><%=svCustId %></td>
				</tr>
				<tr>
					<th><% If svCustPwd Then %><!--[[-->Id<!--]]--><% Else %><!--[[-->Password<!--]]--><% End If %> :</th>
					<td><%=fIf(svMembInternal, "**********", svMembId)%></td>
				</tr>

				<tr>
					<th>&nbsp;</th>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<th><!--[[-->Touch Screen<!--]]-->:</th>
					<td id="touchStatus" width="50%"><%=vTouch%></td>
				</tr>
				<tr>
					<th><!--[[-->Browser<!--]]-->:</th>
					<td><%=vBrowser%>&ensp;<span class="toggleAgent">o</span>&ensp;&ensp;<span class="browserAgent" style="display: none">(asdf asdf asdf asdf )</span></td>
				</tr>
				<tr>
					<th>HTML5 :</th>
					<td><%=vHTML5%></td>
				</tr>
				<tr>
					<th>Flash :</th>
					<td><%=vFlash%></td>
				</tr>
				<tr>
					<th><!--[[-->Cookies Enabled<!--]]--> :</th>
					<td><%=vCookies%></td>
				</tr>
				<tr>
					<th><!--[[-->Popups Enabled<!--]]-->:</th>
					<td id="popupStatus"><%=vPopup %></td>
				</tr>
				<tr>
					<th><!--[[-->Ecommerce Ready<!--]]--> :</th>
					<td><%=vEcom%></td>
				</tr>
				<tr>
					<th>&nbsp;</th>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<th><!--[[-->First Visit<!--]]-->:</th>
					<td><%=fFormatDate(svMembFirstVisit)%></td>
				</tr>
				<tr>
					<th><!--[[-->Last Visit<!--]]-->:</th>
					<td><%=fFormatDate(svMembLastVisit)%></td>
				</tr>
				<% If fIsGroup2 Then %>
				<tr>
					<th><!--[[-->Account Expires<!--]]-->:</th>
					<td><%=fFormatDate(vCust_Expires)%></td>
				</tr>
				<% 
				Else		
					If IsDate(svMembExpires) Then 
						If svMembExpires > Now Then 
				%>
				<tr>
					<th><!--[[-->Programs Expires<!--]]-->:</th>
					<td><%=fFormatDate(svMembExpires)%></td>
				</tr>
				<%
						End If
					End If
				End If 
				%>
			</table>
		</div>


	</div>

	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
