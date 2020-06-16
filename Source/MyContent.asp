<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->

<%

	Dim aProgs, aProg, vProgId, vProgCnt, vProgsOk
	Dim vAlign, vStatusTitle, vBg, vMode, vOk, vExpires, vSort, vCatlPrev, vCertUrl, vTitle, vAttempts, vErrMsg, vUrl, aMods, vLine, vFolder
	Dim vModId, vMaxAttempts, vPassGrade, vBestScore, vLastScore, vNext

	vProgId   = Request("vProgId")
	vSort     = fDefault(Request("vSort"), "n")
	vMode     = fIf(svMembLevel > 2, "All", "My")
	vCatlPrev = ""

	sPrintHeader

	'...get customer product string
	sGetCust svCustId

	'...get any user and ecom program 
	sGetMemb svMembno
	vEcom_Programs = fEcomPrograms (svCustId, svMembId)
	
	'...Use to see if any lines are printed, if not go to another page
	vProgCnt = 0

	'...keep a list of ok progs so dups don't appear
	vProgsOk = ""
	
	'...get the Catalogue info
	If vSort = "y" Then 
		sGetCatlByTitle_Rs svCustId
	Else
		sGetCatl_Rs svCustId
	End If
	Do While Not oRs2.Eof
	
		'...get catalogue info from catalogue table
		sReadCatl
		
		'...extract the program strings from the catalogue content string
		aProgs = Split(vCatl_Programs)
		
		'...process each program
		For j = 0 To Ubound(aProgs) '...aProgs(j): "P1001EN~50~79~23.5~90"
			aProg = Split(aProgs(j), "~") 
		
			'...get program info from the prog table
			sGetProg aProg(0)

			'...get pricing unless price is 9999  
			vProg_US       = aProg(1)
			vProg_CA       = aProg(2)
			vProg_Duration = aProg(4)
		
			vOk = False
		

'	stop

			'...if facilitator/manager/administrator show all programs that are chargeable except the inactive ones  
'     If vMode = "All" And vProg_US > 0 And vProg_US <> 9999 Then

			'...if facilitator/manager/administrator show all programs
			If vMode = "All" Then

				vOk = True
		
			'...else for users prog must be free, purchased (but not via group2) or put onto the member table
			ElseIf (vProg_US = 0 And vCust_MaxUsers >=0) Or (Instr(vEcom_Programs, vProg_Id) > 0 And vCust_MaxUsers >=0) Or Instr(vMemb_Programs, vProg_Id) > 0 Then
		
				 '...if free ensure there is no duration or not expired
				 If vProg_US = 0 Then
					 If vProg_Duration = 0 Then
						 vOk = True
					 Else
						 If Not IsDate(fFormatDate(svMembFirstVisit)) Then svMembFirstVisit = Now
						 If DateAdd("d", vProg_Duration, svMembFirstVisit) > Now Then vOk = True
					 End If
				 Else
					 vOk = True
				 End If
		
				 '...ensure prog only displayed once 
				 If vOk Then
					 If Instr(vProgsOk, vProg_Id) > 0 Then
						 vOk = False
					 Else
						 vProgsOk = vProgsOk & " " & vProg_Id
					 End If
				 End If
		
			End If
		
			If vOk Then
				vProgCnt = vProgCnt + 1

				'...determine expiry date of the content
				vExpires = ""

				'...if from the member record
				If Instr(vMemb_Programs, vProg_Id) > 0 Then 
										 
					'...if entered an expiry date
					If fFormatDate(vMemb_Expires) <> " " Then
						vExpires = vMemb_Expires

					'...else a duration  
					ElseIf vMemb_Duration > 0 Then
						vExpires = DateAdd("d", vMemb_Duration, svMembFirstVisit)
					End If
		
				'...if from the ecom record
				ElseIf Instr(vEcom_Programs, vProg_Id) > 0 Then 
					k = Instr(vEcom_Programs, vProg_Id) '...is this program in the ecom string?
					l = Instr(k, vEcom_Programs, "|") - 1 '...find the end of the pair
					If l = -1 Then l = Len(vEcom_Programs) '...else get the end of string
					vExpires = Mid(vEcom_Programs, k+8, l-k-7)            
		
				'...if free (then no expiry if duration=0 else expires after firstvisit plus duration)
				ElseIf vProg_US = 0 And vProg_Duration = 0 Then
					vExpires = ""

				'...if on customer record
				ElseIf fFormatDate(vCust_Expires)  <> " " Then
					vExpires = vCust_Expires

				'...else create the expiry date
				Else
					vExpires = DateAdd("d", vProg_Duration, svMembFirstVisit)
				End If
							 
				'...if there is an expiry on the customer file then nothing can excede this date
				If fFormatDate(vCust_Expires) <> " " Then
					If fFormatDate(vExpires) <> " " Then
						If DateDiff("d", vCust_Expires, vExpires) > 0 Then
							vExpires = vCust_Expires
						End If
					Else
						vExpires = vCust_Expires
					End If
				End If

'       If fFormatDate(vMemb_Expires) = " " Then vMemb_Expires = Now
	 
				If Instr(vProg_Title, "<") > 0 Then vProg_Title = Left(vProg_Title, Instr(vProg_Title, "<") -1 )
				If Instr(vProg_Title, "(") > 0 Then vProg_Title = Left(vProg_Title, Instr(vProg_Title, "(") -1 )
		
				'...print each program header info
				sPrintPrograms          
	 
			End If

		Next
		oRs2.MoveNext
	Loop
 
	'...no programs?
	If vProgCnt = 0 Then sNoPrograms

	sPrintFooter
	



	'...HTML _______________________________________________________________________
	Sub sPrintHeader
%>

<html>

<head>
	<title>My Content</title>  
	<meta charset="UTF-8">
	<link href="<%=svDomain%>/Inc/Vubiz_Original.css" type="text/css" rel="stylesheet">
	<% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
	<script src="/V5/Inc/Functions.js"></script>
	<script src="/V5/Inc/Launch.js"></script>
	<script src="/V5/Inc/jQuery.js"></script>
	<script>AC_FL_RunContent = 0;</script>
	<script src="/V5/Inc/AC_RunActiveContent.js" language="javascript"></script>
	<script>
		function jTitle (vTitle, vImage) {
			var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
			AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
		}

		if (location.hash !== "") window.location.hash=location.hash
		function fAlert() {
			var vPhrase = "/*--{[--*/You have no more attempts available for this assessment./*--]}--*/"
			alert(vPhrase);
		}

		/* when page gets focus (after its lost it, ie if !bodyFocus - set to false when window is launched */
		$(function() {
			$(window).focus (
				function() {
					if (!parent.bodyFocus) {
						parent.bodyFocus = true;
						location.reload();
					}
				}
			);
		}	);
	</script>
</head>

<body>

<% 
	Server.Execute vShellHi 
%>
	
	<table class="table">
		<tr>
			<td style="white-space:nowrap">
				<script>jTitle("/*--{[--*/My Content/*--]}--*/", 'SingleLicense.jpg')</script>
			</td>
			<td style="white-space:nowrap">
				<h1><!--[[-->All available courses are listed by category below. &nbsp;To navigate:<!--]]--></h1>
				<ol class="c2">
					<li><!--[[-->Click on a bolded <b>Program Title</b> to expand the program.<!--]]--></li>
					<li><!--[[-->Click on a black <font color="#000000">Module Title</font> to launch a module.<!--]]--></li>
					<li><!--[[-->Click on a green <font color="#008000">[Status]</font> link for details of your learning activities.<!--]]--></li>
					<li><!--[[-->Sort by<!--]]-->: <a href="MyContent.asp?vSort=n"><!--[[-->Default Order<!--]]--></a> | <a href="MyContent.asp?vSort=y"><!--[[-->Category Order<!--]]--></a>.</li>
				</ol>
			</td>
			<td style="text-align:center; white-space:nowrap">
				<% If svMembLevel = 5 Or svMembManager Then %><a class="c2" href="Patience.asp?vNext=RTE_MyContent.asp"><!--[[-->New view!<!--]]--></a><% End If %>
			</td>
		</tr>
	</table>

<% 
	End Sub 




	Sub sPrintPrograms        
%>
		<div align="right">
			<table class="table">
				<!-- Program Header Table -->
				<% 
					If vCatlPrev <> vCatl_Title Then
				%>
				<tr>
					<td align="Left"><% If vCatlPrev <> "" Then Response.Write "<br><br>" %> <h1><%=vCatl_Title%>asdfasdfasdf</h1></td>
				</tr>
				<%
						vCatlPrev = vCatl_Title
					End If 
				%>
				<tr>
					<td style="text-align:left">
					<div  style="text-align:right">
						<table style="width:95%; margin-left:5%">
							<tr>
								<td height="30"><a name="<%=vProg_Id%>"></a><p class="c1"><a target="_self" href="MyContent.asp?vLang=<%=svLang%>&vProgId=<%=vProg_Id%>&vSort=<%=vSort%>#<%=vProg_Id%>"><%=vProg_Title%></a></p></td>
							</tr>
							<%  If (vProgId = "" And vProgCnt = 1) Or vProgId = vProg_Id Then %>
							<tr>
								<td align="Left">
								<table border="0" width="100%" id="table14" cellspacing="0" cellpadding="2">
											
									<% If vMode = "My" And Len(vExpires) > 1 Then %>
									<tr>
										<td width="5%">&nbsp;</td>
										<td width="90%"><p class="c2"><b>
										<!--[[-->Expires<!--]]-->:</b> <%=fFormatDate(DateAdd("d", -1, vExpires))%><br></p></td>
									</tr>
									<% End If %> 
											
									<% If Len(vProg_Desc) > 0 Then %>
									<tr>
										<td width="5%">&nbsp;</td>
										<td width="90%"><p class="c2"><%=vProg_Desc%></p></td>
									</tr>
									<% End If %>


									<% If Len(Trim(vProg_Mods)) > 0 Then %>
									<tr>
										<td width="5%">&nbsp;</td>
										<td width="90%"><br><!--[[-->Estimated program length<!--]]--> : <%=vProg_Length%>&nbsp;<!--[[-->Hour(s)<!--]]-->.<br>&nbsp;</td>
									</tr>
									<tr>
										<td width="5%">&nbsp;</td>
										<td width="90%"><p class="c2"><b>
										<font color="#000000"><!--[[-->Modules<!--]]--></font> :</b></p></td>
									</tr>
									<% End If %>



								</table>
								</td>
							</tr>
							<%  End If 
									
									'...begin print module info if first program or selected program _____________________________________________________
									If (vProgId = "" And vProgCnt = 1) Or vProgId = vProg_Id Then 
							%>
							<tr>
								<td>
								<div>&nbsp;
									<table border="0" id="table13" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="2">

										<%
											'...vNext is where the RTE returns to on this page
											vNext = svPage
											vNext = Server.HtmlEncode(svPage & "?vProgId=" & vProg_Id & "&vSort=n#" & vProg_Id)
											vNext = Server.UrlEncode(svPage & "?vProgId=" & vProg_Id & "&vSort=n#" & vProg_Id)

											aMods = Split(Trim(vProg_Mods), " ")
											For vLine = 0 To Ubound(aMods)
												sGetMods aMods(vLine)
												If vMods_Active Then
													vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#EDF5FC' bordercolor='#FFFFFF'"   '...color ever other line        
										%>
										<tr>
											<td width="5%">&nbsp;</td>
											<td <%=vbg%> nowrap class="c8">

												<% If vMods_FullScreen Then %>
												<a class="c8" <%=fstatx%> href="javascript:fullScreen('<%=vProg_Id%>|<%=vMods_Id%>|<%=vProg_Test%>|<%=vProg_Bookmark%>|<%=vProg_CompletedButton%>')"><%=vMods_Title%></a> 
												<% ElseIf Ucase(vMods_Type) = "FX" OR Ucase(vMods_Type) = "XX" Then %>
												<a class="c8" href="/V5/LaunchObjects.asp?vModId=<%=vProg_Id%>|<%=vMods_Id%>&vNext=<%=vNext%>"><%=vMods_Title%></a>
												<% Else %>
												<a class="c8" <%=fstatx%> href="javascript:<%=vMods_Script%>('<%=vProg_Id%>|<%=vMods_Id%>|<%=vProg_Test%>|<%=vProg_Bookmark%>|<%=vProg_CompletedButton%>')"><%=vMods_Title%></a> 
												<% End If %>														

												<% If Len(Trim(vMods_AssessmentUrl)) > 0 Then %>
													[<a <class="c8" <%=fstatx%> href="javascript:<%=vMods_AssessmentScript%>('<%=vMods_AssessmentUrl%>','<%=vMods_Id%>')"><!--[[-->Assessment<!--]]--></a>] 
												<% End If %> 
											</td>
											<td <%=vbg%> nowrap width="30%" class="green">
												[<a href="javascript:SiteWindow('ModuleDescription.asp?vClose=Y&vModId=<%=vMods_Id%>')"><!--[[-->Description<!--]]--></a>]
												[<%=fModStatusLink (svMembNo, vProg_Id, vMods_Id)%>]

	<%

	'                         [<div id="< =vProg_Id & "|" & vMods_Id >">new status</div>]

	%>

											</td>
										</tr>
										<%  
												End If
											Next


											'...If VuAssess included via a launch module then get Launch Module Id so we can find the scores, etc
											If Len(vProg_Assessment) = 6 Then 
												sGetMods vProg_Assessment
											End If
											If Len(vProg_Assessment) = 6 And vMods_Active Then                         
												vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#EDF5FC' bordercolor='#FFFFFF'"
										%>      
										<tr>
											<td width="5%">&nbsp;</td>
											<td <%=vbg%> class="c8">

											<% 
												'...if learner has passed then generate a cert
												vBestScore = fBestScore (svMembNo, vProg_Assessment)                            
												If vBestScore/100 >= fIf(vProg_AssessmentScore = 0, .8, vProg_AssessmentScore) Then 
													vTitle   = fModsTitle(vProg_Assessment)


	'                              If Len(vProg_AssessmentCert) > 0 Then
	'                                vFolder = vProg_AssessmentCert
	'                              ElseIf Len(vCust_AssessmentCert) > 0 Then
	'                                vFolder = vCust_AssessmentCert
	'                              Else
	'                                vFolder = svLang
	'                              End If


													vLastScore = fLastPassed(vProg_Assessment, vProg_AssessmentScore)
											%> 
														<a <%=fstatx%> class="c8" href="javascript:fullScreen('<%=fCertificateUrl("", "", vBestScore, vLastScore, vProg_Assessment, vTitle, "", "", "", vProg_Id, "", "", "")%>')"><!--[[-->Examination<!--]]--></a>
											<% 
												Else
													'...determine how many attempts this learner has
													If vProg_AssessmentAttempts > 0 Then
														vAttempts = vProg_AssessmentAttempts
													ElseIf vCust_AssessmentAttempts > 0 Then
														vAttempts = vCust_AssessmentAttempts
													Else
														vAttempts = 3
													End If
																 
													If vAttempts = 99 Or fNoAttempts(svMembNo, vProg_Assessment) < vAttempts Then 
											%> 
														<% If vMods_FullScreen Then %>
														<a <%=fstatx%> href="javascript:fullScreen('<%=vProg_Id%>|<%=vMods_Id%>|<%=vProg_Test%>|<%=vProg_Bookmark%>|<%=vProg_CompletedButton%>')"><!--[[-->Examination<!--]]--></a> 
														<% ElseIf Ucase(vMods_Type) = "FX" Then %>
														<a class="c8" <%=fstatx%> href="/V5/LaunchObjects.asp?vModId=<%=vProg_Id%>|<%=vMods_Id%>&vNext=<%=vNext%>"><!--[[-->Examination<!--]]--></a>
														<% Else %>
														<a class="c8" <%=fstatx%> href="javascript:<%=vMods_Script%>('<%=vProg_Id%>|<%=vMods_Id%>|<%=vProg_Test%>|<%=vProg_Bookmark%>|<%=vProg_CompletedButton%>')"><!--[[-->Examination<!--]]--></a> 
														<% End If %>
											<% 
													Else 
											%> 
															<a <%=fstatx%> href="javascript:fAlert()"><!--[[-->Examination<!--]]--></a> 
											<% 
													End If 

												End If 
											%>
											</td>
											<td <%=vbg%> class="green" width="30%">[<%=fAssessmentStatus (svMembNo, vProg_Assessment)%>]</td>
										</tr>
										<%   

											'...platform exam included?
											ElseIf Len(vProg_Exam) > 1 Then  
												vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#EDF5FC' bordercolor='#FFFFFF'"
												Session("CertProg") = vProg_Id '...use this for custom platform cert (If needed)
										%>
										<tr>
											<td width="5%">&nbsp;</td>
											<td <%=vbg%> class="c8">

										<%
												'...if this does NOT use a platform custom cert then use VuAssess Custom Cert
												vOk = False
												If Not vProg_CustomCert Then
													'...grab parms from the exam string
													i = Instr(Lcase(vProg_Exam), "vmodid=")         : vModId       = Mid(vProg_Exam, i+7, 6)
													i = Instr(Lcase(vProg_Exam), "vmaxattempts=")   : vMaxAttempts = Mid(vProg_Exam, i+13, 1)
													i = Instr(Lcase(vProg_Exam), "vpassgrade=")     : vPassGrade   = Mid(vProg_Exam, i+11, 2) / 100
													'...if learner has passed then generate a vuCert
													vBestScore = fBestScore (svMembNo, vModId)                            
													If vBestScore/100 >= fIf(vPassGrade = 0, .8, vPassGrade) Then 
														vOk = True
														sOpenDbBase
														vSql = "Select * FROM TstH WHERE TstH_ID = '" & vModId & "'"
														Set oRs = oDbBase.Execute(vSql)    
														vTitle = oRs("TstH_Title")
														sCloseDbBase  


	'                                If Len(vProg_AssessmentCert) > 0 Then
	'                                 vFolder = vProg_AssessmentCert
	'                                ElseIf Len(vCust_AssessmentCert) > 0 Then
	'                                 vFolder = vCust_AssessmentCert
	'                                Else
	'                                 vFolder = svLang
	'                                End If


														vLastScore = fLastPassed(vModId, vPassGrade)
													End If
												End If                            
												If vOk Then   													
											%> 
													<a <%=fstatx%> class="c8" href="javascript:fullScreen('<%=fCertificateUrl("", "", vBestScore, vLastScore, vProg_Assessment, vTitle, svLang, "", "", vProg_Id, "", "", "")%>')"><!--[[-->Examination<!--]]--></a>
											<%
												'...if using a custom cert then let it follow former path
												Else  
											%>
													<a <%=fstatx%> href="javascript:examwindow('<%=vProg_Exam%>')"><!--[[-->Examination<!--]]--></a>
											<%
												End If
											%>

											</td>
											<td <%=vbg%> width="30%" class="green">[<%=fExamStatus (svMembNo, vProg_Exam)%>]</td>
										</tr>
										<%
											End If            
										%>  

	<!-- don't use vProg_Cert any more -->      
										<%
											If vProg_Cert <> 0 Then  
												vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#EDF5FC' bordercolor='#FFFFFF'"
										%>
										<tr>
											<td width="5%">&nbsp;</td>
											<td <%=vbg%> class="c8">
												<a <%=fstatx%> href="javascript:vuwindow('CertCompletion.asp?vProgId=<%=vProg_Id%>&vClose=Y',650,425,100,100,'no','no','no')"><!--[[-->Certificate of Completion<!--]]--></a>
											</td>
											<td <%=vbg%> class="c2">&nbsp;</td>
										</tr>
										<% End If %>


										<%
											'...generate a conditional certificate if a collection if a set of assessments has been passed (from previous Programs)
											If Len(vProg_AssessmentIds) > 0 Then
												If vProg_AssessmentScore = 0 Then vProg_AssessmentScore = .8
												vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#EDF5FC' bordercolor='#FFFFFF'"
										%>
										<tr>
											<td width="5%">&nbsp;</td>
											<td <%=vbg%> class="c8">
										<%  
												'...get last date all assessments were passed, else ""
												i = fLastPassed(vProg_AssessmentIds, vProg_AssessmentScore)
												If IsDate(i) Then
	'                             vUrl = "/V5/Assessments/CustomCerts/" & vProg_AssessmentCert & "/Default.htm?vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vLastScore=" & fFormatDate(i) & "&vMods_Id=0000EN"
	%>                              
													<a <%=fstatx%> class="c8" href="javascript:fullScreen('<%=fCertificateUrl("", "", vBestScore, i, vProg_Assessment, vTitle, "", "", "", vProg_Id, "", "", "")%>')"><!--[[-->Certificate of Completion<!--]]--></a>
	<%
												Else
													vUrl = "<!--{{-->In order to be granted a Certificate of Completion<br>you must achieve at least<!--}}-->" & "&nbsp;" & FormatPercent(vProg_AssessmentScore, 0) & "&nbsp;" & "<!--{{-->on all assessments.<!--}}-->"
													vUrl = "Error.asp?vErr=" & Server.UrlEncode(vUrl) & "&vClose=Y&vReturn=close"
	%>
												<a class="black" <%=fstatx%> href="#" onclick="window.open('<%=vUrl%>','Vubiz','toolbar=no,scrollbars=yes,width=750,height=500')"><!--[[-->Certificate of Completion<!--]]--></a>
	<%
												End If
	%>
											</td>
											<td <%=vbg%> class="c2">&nbsp;</td>
										</tr>
										<%
											End If
										%>
									</table>
								</div>
								</td>
							</tr>
				<%  
						End If 
						'...end print module info _____________________________________________________
				%>
						</table>
					</div>
					</td>
				</tr>
			</table>
		</div>
<% 
	End Sub




	Sub sPrintFooter
%>
	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

	</body>
</html>
<% 

	End Sub 




	Sub sNoPrograms
%>

<h6 style="text-align:center">
<!--[[-->There are no programs available.<!--]]--></h6>
<% 
	End Sub
%>