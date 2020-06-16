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
<!--#include virtual = "V5/Inc/RTE.asp"-->

<!--#include file = "RTE_Routines.asp"-->
<!--#include file = "RTE_ModuleStatusRoutines.asp"-->

<%
	Dim aProgs, aProg, vProgId, vProgCnt, vProgsOk, vRow
	Dim vMode, vOk, vExpires, vCatlPrev, vCertUrl, vTitle, vAttempts, vErrMsg, vUrl, aMods, vLine, vFolder, vCatlCnt
	Dim vPassScore, vBestScore, vLastScore
	Dim vLaunch, vCertificate, vNext, cLaunch, cCertificate
	
	cCertificate = fPhraH(000089)
	cLaunch = fPhraH(001419)
	vRow      = 0 '...used to assign a number to each row for highlighting
	vProgId   = Request("vProgId")
	vMode     = fIf(svMembLevel > 2, "All", "My")
	vCatlCnt  = 0
	vCatlPrev = ""
	vProgsOk  = ""  '...keep a list of ok progs so dups don't appear

	'...this gets the module features after spRTECatl has been run (a generic version of this is in Db_Mods.asp fModsFeatures)
	Function fFeatures   
		fFeatures = "&nbsp;"
		If RTE_ModsFeaAcc Then fFeatures = fFeatures & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaAcc.png'>&nbsp;"
		If RTE_ModsFeaAud Then fFeatures = fFeatures & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaAud.png'>&nbsp;"
		If RTE_ModsFeaMob Then fFeatures = fFeatures & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaMob.png'>&nbsp;"
		If RTE_ModsFeaHyb Then fFeatures = fFeatures & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaHyb.png'>&nbsp;"
		If RTE_ModsFeaVid Then fFeatures = fFeatures & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaVid.png'>&nbsp;"    
	End Function

	'...return the latest date ALL assessments were completed - else return null
	Function fCompletionDate(progMods)

		progMods = Trim(progMods)

		Dim aProgMods, progNo, modsNo, lastDate, thisDate, i, j
		fCompletionDate = null
		lastDate = null
		If (InStr(progMods, "|") = 0) Then Exit Function

		aProgMods = Split(progMods)

		For i = 0 To Ubound(aProgMods)
			j = Split(aProgMods(i), "|") 
			If (Ubound(j) <> 1) Then fCompletionDate = null : Exit Function
			progNo = fProgNoById (j(0))
			modsNo = fModsNoById (j(1))
			thisDate = fRTEmoduleCompletionDate(vMemb_No, progNo, modsNo)
			If (Not IsDate(thisDate)) Then fCompletionDate = null : Exit Function
			If (Not IsDate(lastDate)) Then 
				lastDate = thisDate
			Else
				lastDate = fMaxDate(thisDate, lastDate)
			End If
		Next

		fCompletionDate = lastDate

	End Function

%>

<html>

<head>
	<title>My Content</title>
	<meta charset="UTF-8">
	<link href="//code.jquery.com/ui/1.9.2/themes/base/jquery-ui.css" rel="stylesheet">
	<script src="/V5/Inc/jQuery.js"></script>
	<script src="/V5/Inc/jQuery.draggable.js"></script>
	<link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
	<script src="/V5/Inc/Functions.js"></script>
	<% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
	<script src="/V5/Inc/Launch.js"></script>
	<script>AC_FL_RunContent = 0;</script>
	<script type="text/javascript" src="/V5/Inc/AC_RunActiveContent.js"></script>
	<script>    

			function jTitle(vTitle, vImage) {
				var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
				AC_FL_RunContent('codebase', '//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0', 'name', 'flashVars', 'width', '265', 'height', '85', 'align', 'middle', 'id', 'flashVars', 'src', '/V5/Images/Titles/VuTitles', 'FlashVars', vParm, 'quality', 'high', 'bgcolor', '#ffffff', 'allowscriptaccess', 'sameDomain', 'allowfullscreen', 'false', 'pluginspage', '///go/getflashplayer', 'movie', '/V5/Images/Titles/VuTitles');
			};

			/* this contains the Y axis of the click where we position the divOuter */
			var pageY = 20; 

			// determine vertical offset after scrolling so we can position the info window at pageY
			function determinePageY(){
				if (self.pageYOffset) {
					pageY = self.pageYOffset;  // modern browsers
				} else if (document.documentElement && document.documentElement.scrollTop) {
					pageY = document.documentElement.scrollTop; // Explorer 6 Strict
				} else if (document.body) {
					pageY = document.body.scrollTop; // all other Explorers
				};
				pageY = pageY + 20;  
			};

			window.onscroll = function () {
				determinePageY();
			};

			// this is the row that uses the editors - needed for highlighting
			var editorRowEle;

			/*
				 high light a rows background with white, lightYellow or Yellow
				 if Yellow then dont wipe out with normal white/lightYellow
				 to turn off a Yellow background use next routine
			*/
			function hiLiteRow(ele, cl)   {
				var curClass = $(ele)[0].className;
				if (curClass != "bgAlert") {
					$(ele)[0].className = cl;
				} else {
					editorRowEle = ele;
				}
			}

			function noLiteRow(ele, cl) {
				if (ele != undefined) {
					$(ele)[0].className = cl;
				}
			}

			$(function () {

				// make the div draggable
				$("#divOuter").draggable();

				// Disable caching of AJAX responses
				$.ajaxSetup({ cache: false });

				// capture the postion of the mouse click so we know where to position the info window (replaced above)
				determinePageY();

				//  when page gets focus after its lost it, ie if !bodyFocus - set to false when window is launched
				$(window).focus(
					function () {
						if (!parent.bodyFocus) {
							parent.bodyFocus = true;
							location.reload();
						}
					}
				);
			});


			// generic div/window  for all functions
			function openWindow(parm1, parm2, parm3, parm4) {        
				closeWindow();                                  // reset window        
				divOff("flash")                                 // turn off since we cannot render on top of it

				if (reNumeric.test(parm3) == null) {
					pageY = parm3;                                // RTE_ModsStats launches a cert and we want to have it rendered at the same Y position as the original click, so if we get a value in parm3 we use that for the click/y top position
				}

				$("#divOuter")[0].style.width  = 700 + "px";    // define default width  (modified for session editor)
				$("#divOuter")[0].style.height = 500 + "px";    // define default height (modified for session editor)
				$("#tabOuter")[0].style.height = 450 + "px"; 

				switch (parm1) {
					case "Mods" : $("#divInner").load("Module.asp?vMods_Id=" + parm2); divOn("divInner"); break;                                                      // program description
					case "P"    : $("#divInner").load("RTE_ProgDesc.asp?vProgId=" + parm2); divOn("divInner"); break;                                                 // program description
					case "M"    : $("#divInner").load("RTE_ModsDesc.asp?vModsId=" + parm3); divOn("divInner"); break;                                                 // module description
					case "S"    : $("#divInner").load("RTE_ModsStat.asp?vProgNo=" + parm2 + "&vModsId=" + parm3 + "&vProgId=" + parm4); divOn("divInner"); break;     // status 1 (new with expired history)
					case "C"    : 
						$("iframe")[0].src = parm2; 
						divOn("ifrInner"); 
						$("#divOuter")[0].style.width  = 800 + "px"; 
						$("#divOuter")[0].style.height = 600 + "px"; 
						$("#tabOuter")[0].style.height = 550 + "px"; 
						break;                                                                               // certificate service
					case "Edit" :                                                                                                                                     // session editor service
						$("iframe")[0].src = "/Gold/vuSCORMAdmin/sessionEdit.aspx?sessionID=" + parm2; 
						divOn("ifrInner"); 
						$("#divOuter")[0].style.width  = 400 + "px"; 
						$("#divOuter")[0].style.height = 475 + "px"; 
						$("#tabOuter")[0].style.height = 425 + "px"; 
						break;   
					case "Hist" :                                                                                                                                     // session editor service
						$("iframe")[0].src = "/Gold/vuSCORMAdmin/Default.aspx?memberID=" + parm2; 
						divOn("ifrInner"); 
						$("#divOuter")[0].style.width = 1000 + "px"; 
						$("#divOuter")[0].style.height = 800 + "px"; 
						$("#tabOuter")[0].style.height = 750 + "px"; 
						break;   
				}
				$("#divOuter")[0].style.top = pageY + "px";          // set top to pageY (computed onload/onscroll)

				divOn("divOuter");
			}

			function closeWindow() {
				divOff("divOuter");
				divOff("divInner");
				divOff("ifrInner");
				divOn("flash");
				noLiteRow(editorRowEle, 'bgOff');
			}  

	</script>
	<style>
		.button070 { margin: 2px; }
		.bgOn { background-color: #DDEEF9; }

		td { vertical-align: middle; }
		#divOuter { position: absolute; padding: 0px; width: 700px; height: 500px; left: 15px; background-color: #FFFFFF; z-index: 99; }
		td, th { border: 0; }
	</style>
</head>

<body>

	<% Server.Execute vShellHi %>

	<table class="table">
		<tr>
			<td>
				<div id="flash">
					<img src="../Images/Ecom/MyContent_<%=svLang %>.png" />
				</div>
			</td>
			<td style="text-align: right; width: 300px;">
				<table>
					<tr>
						<th style="text-align: center; padding-bottom: 10px;" colspan="6">
							<!--webbot bot='PurpleText' PREVIEW='Content Features'--><%=fPhra(001413)%>&nbsp; <span style="font-weight: 400">(<!--webbot bot='PurpleText' PREVIEW='mouseover'--><%=fPhra(001424)%>)</span></th>
					</tr>
					<tr>
						<td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaAcc.png" title="<!--webbot bot='PurpleText' PREVIEW='Includes compatibility with most screen readers and closed captioning (WCAG Level AA).'--><%=fPhra(001442)%>"></a> </td>
						<td>
							<!--webbot bot='PurpleText' PREVIEW='Accessible'--><%=fPhra(001415)%></td>
						<td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaHyb.png" title="<!--webbot bot='PurpleText' PREVIEW='Content available in Flash or HTML.'--><%=fPhra(001628)%>"></a> </td>
						<td>
							<!--webbot bot='PurpleText' PREVIEW='Hybrid'--><%=fPhra(001613)%></td>
						<td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaMob.png" title="<!--webbot bot='PurpleText' PREVIEW='Tablet friendly.'--><%=fPhra(001641)%>"></a> </td>
						<td>
							<!--webbot bot='PurpleText' PREVIEW='Mobile'--><%=fPhra(001416)%></td>
					</tr>
					<tr>
						<td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaAud.png" title="<!--webbot bot='PurpleText' PREVIEW='Requires headphones or speaker to hear audio.'--><%=fPhra(001443)%>"></a> </td>
						<td>
							<!--webbot bot='PurpleText' PREVIEW='Audio'--><%=fPhra(001417)%></td>
						<td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaVid.png" title="<!--webbot bot='PurpleText' PREVIEW='Contains or streams video content.'--><%=fPhra(001445)%>"></a> </td>
						<td>
							<!--webbot bot='PurpleText' PREVIEW='Video'--><%=fPhra(001418)%></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<%  

		sGetCust svCustId                                         '...get customer product string
		sGetMemb svMembNo                                         '...get user data
		vEcom_Programs = fEcomPrograms (svCustId, svMembId)       '...get any purchased programs
		sGetCatl_Rs svCustId                                      '...get catalogue for this customer

			 
		Do While Not oRs2.Eof                                     '...go through catalogue
			sReadCatl
			vCatlCnt = vCatlCnt + 1

			aProgs = Split(vCatl_Programs)                          '...extract the program strings from the catalogue content string

			'...process each program
			For j = 0 To Ubound(aProgs)                             '...aProgs(j): "P1001EN~50~79~23.5~90"
				aProg = Split(aProgs(j), "~")     
				sGetProg aProg(0)                                     '...get program info from the prog table (looking for freebies)
				vProg_US       = aProg(1)                             '...get pricing unless price is 9999  
				vProg_CA       = aProg(2)
				vProg_Duration = aProg(4)        
				vOk = False    


		'if (aProg(0) = "P3281EN") Then stop


				'...if facilitator/manager/administrator show all programs
'         If vMode = "All" And Len(Trim(vProg_Mods)) > 0 Then
				If vMode = "All" Then  '...added Aug 2012 to allow special certs that have no modules - just used to generate a certificate
					vOk = True
		
				'...else program must be free, purchased (but not via group2) or put onto the member table
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
			
					If Instr(vProg_Title, "<") > 0 Then vProg_Title = Left(vProg_Title, Instr(vProg_Title, "<") -1 )
					If Instr(vProg_Title, "(") > 0 Then vProg_Title = Left(vProg_Title, Instr(vProg_Title, "(") -1 )
				
	%>

	<table class="table">
		<% 
					If vCatlPrev <> vCatl_Title Then
		%>
		<tr>
			<td colspan="2" style="height: 30px; border-top: 1px solid #008000;">
				<p class="c2"><%=vCatl_Title%></p>
			</td>
		</tr>
		<%
						vCatlPrev = vCatl_Title
					End If 

					vCertificate = "&nbsp;"
					If Len(vProg_AssessmentIds) > 0 Then
						'...check for program completion certificate (via RTE NOT LMS) - re-added Apr 22, 2016 - values must be P1234EN|1234EN P54332EN|9876EN else returns null
						 vLastScore = fCompletionDate(vProg_AssessmentIds)
						 If IsDate(vLastScore) Then 
							 vUrl = fCertificateUrl("", "", "", vLastScore, "", vProg_Title, "", "", "", vProg_Id, "", "", "")
							 vCertificate = "<input onclick=""openWindow('C', '" & vUrl & "')"" type=""button"" value=""Certificate"" name=""B5"" class=""button070"">"
						 Else  
							 vCertificate = fPhraH(001776)
							 vCertificate = "<input onclick=""alert('" & vCertificate & "')"" type=""button"" value=""" & cCertificate & """ name=""B5"" class=""button070"">"
						 End If
					End If

					vRow = vRow + 1
		%>
		<tr>

			<td>&ensp;&ensp;</td>
			<td>
				<table class="table">
					<tr id="R_<%=vRow %>" onmouseover="hiLiteRow(this, 'bgOn')" onmouseout="hiLiteRow(this, 'bgOff')">
						<td class="underline" style="width: 665px">
							<%=vProg_Title%>
							<a href="javascript:hiLiteRow($('#R_<%=vRow %>'), 'bgAlert'); openWindow('P', '<%=vProg_Id%>', '')">
								<!--webbot bot='PurpleText' PREVIEW='Description'--><%=fPhra(000118)%></a>
							<% If vMode <> "All" Then %>|<!--webbot bot='PurpleText' PREVIEW='Expires'--><%=fPhra(000137)%>: <%=fFormatDate(vExpires)%><% End If %>
						</td>
						<td class="underline" style="max-width: 255px;">&nbsp;</td>
						<td class="underline" style="width: 80px; white-space: nowrap;"><%=vCertificate%></td>
					</tr>
					<tr>
						<td colspan="3">
							<div style="text-align: right;">
								<a name="<%=vProg_Id%>"></a>
								<table class="table">
									<%
										RTE_IsLaunchModule = False '...this helps in RTE code to twig Status values
										aMods = Split(Trim(vProg_Mods), " ")
										For vLine = 0 To Ubound(aMods)

											sGetMods(aMods(vLine))
											If vMods_Active Then

												vRow = vRow + 1
									
												'... get the session details for this module  
												sRTEsessionCore svMembNo, vProg_No, aMods(vLine), vProg_Id

												'...create the launch URL - first for FX/XX/H (unless fullscreen, use popup)
												If (Not RTE_ModsFullScreen) And (RTE_ModsType = "Z" Or RTE_ModsType = "FX" Or RTE_ModsType = "XX" Or RTE_ModsType = "H") Then 
													vNext = Server.UrlEncode(svPage & "?vProgId=" & RTE_ProgId)
													vUrl = "location.href='/V5/LaunchObjects.asp?vModId=" & RTE_ProgId & "|" & RTE_ModsId & "&vNext=" & vNext & "'"
												Else 
													RTE_ModsScript = fIf(RTE_ModsFullScreen, "fullScreen", RTE_ModsScript)
													vUrl = "javascript:" & RTE_ModsScript & "('" & RTE_ProgId & "|" & RTE_ModsId & "|" & vProg_Test & "|" & vProg_Bookmark & "|" & vProg_CompletedButton & "')"
												End If 

												vLaunch = "<input onclick=""" & vUrl & """ type=""button"" value=""" & cLaunch & """ name=""B5"" class=""button070"">"
												vCertificate = "&nbsp;"

												If RTE_ModsVuCert And RTE_Status = fPhraH(000107) Then
													vUrl = fCertificateUrl("", "", RTE_BestScore, RTE_CompletedDate, RTE_ModsId, RTE_ModsTitle, "", "", "", RTE_ProgId, "", RTE_SessionId, "")
													vCertificate = "<input onclick=""hiLiteRow($('#R_" & vRow & "'), 'bgAlert'); openWindow('C', '" & vUrl & "')"" type=""button"" value=""" & cCertificate & """ name=""B5"" class=""button070"">"
												End If                          
'stop																								
									%>
									<tr id="R_<%=vRow %>" onmouseover="hiLiteRow(this, 'bgOn')" onmouseout="hiLiteRow(this, 'bgOff')">
										<td class="underline" style="width: 665px"><%=f5%><%=RTE_ModsTitle%> <a href="javascript:hiLiteRow($('#R_<%=vRow %>'), 'bgAlert'); openWindow('M', '<%=RTE_ProgId%>', '<%=RTE_ModsId%>')">
											<!--webbot bot='PurpleText' PREVIEW='Description'--><%=fPhra(000118)%></a><% If svMembLevel > 4 Then %> <span class="green"><a class="green" target="_blank" href="Module.asp?vMods_Id=<%=RTE_ModsId%>"><%=RTE_ModsType%></a></span><% End If %><%=fFeatures%></td>
										<td class="underline" style="width: 175px; white-space: nowrap;"><a href="javascript:hiLiteRow($('#R_<%=vRow %>'), 'bgAlert'); openWindow('S', '<%=vProg_No%>',   '<%=RTE_ModsId%>', '<%=RTE_ProgId%>')"><%=RTE_Status%></a></td>
										<td class="underline" style="width: 080px; white-space: nowrap; text-align: center"><%=vLaunch%></td>
										<td class="underline" style="width: 080px; white-space: nowrap; text-align: center"><%=vCertificate%></td>
									</tr>
									<%  
										End If
									Next

									'...If VuAssess included via a launch module then get Launch Module Id so we can find the scores, etc
									If Len(vProg_Assessment) >= 6 Then
										vRow = vRow + 1                                       
										RTE_IsLaunchModule = True '...this helps in RTE code to twig Status values
										'... get the session details for this assessment
										sRTEsessionCore svMembNo, vProg_No, vProg_Assessment, vProg_Id

										If RTE_Status <> fPhraH(000107) Then
											'...create the launch URL - first for FX/XX (unless fullscreen, use popup)
											If (Not RTE_ModsFullScreen) And (RTE_ModsType = "Z" Or RTE_ModsType = "FX" Or RTE_ModsType = "XX") Then 
												vNext = Server.UrlEncode(svPage & "?vProgId=" & RTE_ProgId)
												vUrl = "location.href='/V5/LaunchObjects.asp?vModId=" & RTE_ProgId & "|" & RTE_ModsId & "&vNext=" & vNext & "'"
											Else 
												RTE_ModsScript = fIf(RTE_ModsFullScreen, "fullScreen", RTE_ModsScript)
												vUrl = "javascript:" & RTE_ModsScript & "('" & RTE_ProgId & "|" & RTE_ModsId & "|" & vProg_Test & "|" & vProg_Bookmark & "|" & vProg_CompletedButton & "')"
											End If 
											vLaunch = "<input onclick=""" & vUrl & """ type=""button"" value=""" & cLaunch & """ name=""B5"" class=""button070"">"
											vCertificate = "&nbsp;"
										Else
											vLaunch = "&nbsp;"
											vUrl = fCertificateUrl("", "", RTE_BestScore, RTE_CompletedDate, RTE_ModsId, RTE_ModsTitle, "", "", "", RTE_ProgId, "", "","")
											vCertificate = "<input onclick=""hiLiteRow($('#R_" & vRow & "'), 'bgAlert'); openWindow('C', '" & vUrl & "')"" type=""button"" value=""" & cCertificate & """ name=""B5"" class=""button070"">"
										End If  

									%>
									<tr id="R_<%=vRow %>" onmouseover="hiLiteRow(this, 'bgOn')" onmouseout="hiLiteRow(this, 'bgOff')">
										<td class="underline" style="width: 665px"><%=f5%><!--webbot bot='PurpleText' PREVIEW='Examination'--><%=fPhra(000132)%><% If svMembLevel > 4 Then %> <span class="green"><a class="green" target="_blank" href="Module.asp?vMods_Id=<%=RTE_ModsId%>"><%=RTE_ModsType%></a></span><% End If %><%=fFeatures%> </td>
										<td class="underline" style="width: 175px; white-space: nowrap;"><a href="javascript:hiLiteRow($('#R_<%=vRow %>'), 'bgAlert'); openWindow('S',  '<%=vProg_No%>', '<%=RTE_ModsId%>', '<%=RTE_ProgId%>')"><%=RTE_Status%></a></td>
										<td class="underline" style="width: 080px; white-space: nowrap; text-align: center"><%=vLaunch%></td>
										<td class="underline" style="width: 080px; white-space: nowrap; text-align: center"><%=vCertificate%></td>
									</tr>
									<%
									End If            
									%>
								</table>
							</div>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<%   
			End If  
			Next
			oRs2.MoveNext
			Loop  
	%>

	<!-- This is a general purpose window with embedded shell (in RTE_History_0/RTE_MyContent) -->
	<div id="divOuter" class="ui-widget-content div" style="background-color: white; text-align: center; border: 1px solid navy; width: 700px; height: 500px;">

		<input style="float: right; margin: 5px; color: red; font-weight: bold;" type="button" onclick="closeWindow()" value="X" name="bClose" class="button">

		<table id="tabOuter" class="shell" style="width: 95%; height: 450px; margin: 20px auto;">
			<tr>
				<td class="shellTopLeft"></td>
				<td class="shellTop"></td>
				<td class="shellTopRight"></td>
			</tr>
			<tr>
				<td class="shellLeft"></td>
				<td style="text-align: center;">

					<div class="div" id="divInner" style="padding: 10px; background-color: #FFFFFF;"></div>
					<iframe class="div" id="ifrInner" name="ifrInner" style="width: 100%; height: 400px; border: 0"></iframe>

				</td>
				<td class="shellRight"></td>
			</tr>
			<tr>
				<td class="shellBottomLeft"></td>
				<td class="shellBottom"></td>
				<td class="shellBottomRight"></td>
			</tr>
		</table>

	</div>

	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


