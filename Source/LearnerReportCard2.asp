<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/AssessmentReset.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->
<!--#include file = "LearnerReportCard_Functions.asp"-->

<%
  Dim vNext, vEdit, vCustId, vFind, vFindId, vFindFailing, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vFindActive, vFindCompleted, vFindBookmarks, vStrDate, vEndDate, vFormat, vTitle
  Dim vCardLast, vDateLast, vProgLast, aMods, vMods, aParms, aTimeSpent, vModCnt
  Dim vTimeSpent, vBookmark, aBookmark, aBest, vBestScore, vBestDate, vNoAttempts, vExam_Id, vDiv, bDiv, bInfoPage, vTSno, vBMno
  Dim vPurchaser, vPurchased, vExpires, vUrl
  Dim vScoreDate, vScoreId, vScoreValue, vPassScore
  Dim vModId, isExam

  '...learners can enter this page from the info page (Home.asp) so need to ensure we have the defaults setup properly
  vCurList       = Request("vCurList")
  vMaxList       = Request("vMaxList")
  vStrDate       = fDefault(Request("vStrDate"), "Jan 1, 2000")
  vEndDate       = fDefault(Request("vEndDate"), fFormatSqlDate(DateAdd("d", 1, Now)))  
  vNext          = Request("vNext")
  vEdit          = fDefault(Request("vEdit"), "User" & fGroup & ".asp")
  vCustId        = fDefault(Request("vCustId"), svCustId)

  vFind          = Request("vFind")
  vFindId        = Request("vFindId")
  vFindFailing   = Request("vFindFailing")
  vFindCompleted = Request("vFindCompleted")
  vFindBookmarks = Request("vFindBookmarks")
  vFindFirstName = Request("vFindFirstName")
  vFindLastName  = Request("vFindLastName")
  vFindEmail     = Request("vFindEmail")
  vFindCriteria  = Request("vFindCriteria")
  vFindActive    = Request("vFindActive")
  vFormat        = Request("vFormat")

  '...debug
  vUrl  = "" _
        & "<br>vCurList: "       & vCurList _
        & "<br>vMaxList: "       & vMaxList _
        & "<br>vStrDate: "       & vStrDate _
        & "<br>vEndDate: "       & vEndDate _
        & "<br>vNext: "          & vNext _
        & "<br>vEdit: "          & vEdit _
        & "<br>vCustId: "        & vCustId _
        & "<br>vFind: "          & vFind _
        & "<br>vFindId: "        & vFindId _
        & "<br>vFindFailing: "   & vFindFailing _
        & "<br>vFindCompleted: " & vFindCompleted _
        & "<br>vFindBookmarks: " & vFindBookmarks _
        & "<br>vFindFirstName: " & vFindFirstName _
        & "<br>vFindLastName: "  & vFindLastName _
        & "<br>vFindEmail: "     & vFindEmail _
        & "<br>vFindCriteria: "  & vFindCriteria _
        & "<br>vFindActive: "    & vFindActive _
        & "<br>vFindCriteria: "  & vFindCriteria _
        & "<br>vFindActive: "    & vFindActive _
        & "<br>vFormat: "        & vFormat
' Response.Write vUrl

  sGetCust vCustId '...need to get vCust_AssessmentScore (mastery)

  vMemb_No = Request("vMemb_No")

  bDiv = False
  bInfoPage = fIf(Request("vInfoPage")="y", True, False)
  vTSno = 0 '...use this to uniquely identify the TimeSpent Divs for JS

  '...if resetting 
  If Request("vReset").Count > 0 Then 
    aParms     = Split(Request("vReset"), "|")
    vMemb_No   = aParms(0)
    vDiv       = aParms(4)
    sAssessmentReset vCust_AcctId, aParms(0), aParms(1), aParms(2), aParms(3)
  End If

  Dim vAssessmentAttempts, vAssessmentScore

  Sub sAssessmentValues (vCustId, vProgId)    
    '...get max no of attempts and mastery score, else use default values
    vAssessmentAttempts = 3
    vAssessmentScore    = 80
    '   any by Account?
    sGetCust vCustId
    If Not vCust_Eof Then
      If vCust_AssessmentAttempts > 0 Then vAssessmentAttempts = vCust_AssessmentAttempts
      If vCust_AssessmentScore    > 0 Then vAssessmentScore    = vCust_AssessmentScore
    End If
    '   any by Program?
    If Len(vProgId) = 7 And vProgId <> "P0000XX" Then 
    	sGetProg (vProgId)
	    If Not vProg_Eof Then
        If vProg_AssessmentAttempts > 0 Then vAssessmentAttempts = vProg_AssessmentAttempts
        If vProg_AssessmentScore    > 0 Then vAssessmentScore    = vProg_AssessmentScore
      End If
    End If
  End Sub

  Function fModLen (i)
    '...determine module length from log item, ie no longer just 6, but typically up to "_" 

'    fModLen = 




  End Function

%>

<html>

<head>
  <title>LearnerReportCard2</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Launch.js"></script>
  <script>

    function open1Div(theDiv) {
      document.getElementById(theDiv).style.display = "block";
    }

    function hide1Div(theDiv) {
      document.getElementById(theDiv).style.display = "none";
    } 


    function openDivs() {
      var divs = document.getElementsByTagName("div");
      var j = divs.length;
      for (i=0; i < j; i++) {
        if (divs[i].id.substring(0, 5) == 'Div_P') {
          document.getElementById(divs[i].id).style.display = "block";
        }
      }  
    }

   function hideDivs() {
      var divs = document.getElementsByTagName("div");
      var j = divs.length;
      for (i=0; i < j; i++) {
        if (divs[i].id.substring(0, 5) == 'Div_P') {
          document.getElementById(divs[i].id).style.display = "none";
        }
      }  
    }


    function initTSelements() {
      var divs = document.getElementsByTagName("div");
      var j = divs.length;
      for (i=0; i < j; i++) {
        if (divs[i].id.substring(0, 10) == 'Div_TSedit') {
          document.getElementById(divs[i].id).style.display = "none";
        }
        if (divs[i].id.substring(0, 10) == 'Div_TSdisp') {
          document.getElementById(divs[i].id).style.display = "block";
        }
      }  
    }

    function initBMelements() {
      var divs = document.getElementsByTagName("div");
      var j = divs.length;
      for (i=0; i < j; i++) {
        if (divs[i].id.substring(0, 10) == 'Div_BMedit') {
          document.getElementById(divs[i].id).style.display = "none";
        }
        if (divs[i].id.substring(0, 10) == 'Div_BMdisp') {
          document.getElementById(divs[i].id).style.display = "block";
        }
      }  
    }

    function initSCelements() {
      var divs = document.getElementsByTagName("div");
      var j = divs.length;
      for (i=0; i < j; i++) {
        if (divs[i].id.substring(0, 10) == 'Div_SCform') {
          document.getElementById(divs[i].id).style.display = "none";
        }
        if (divs[i].id.substring(0, 10) == 'Div_SClink') {
          document.getElementById(divs[i].id).style.display = "block";
        }
      }  
    }

    function setSCprog(vProgId) {
      document.getElementById("vScoreProg").value = vProgId;
    }


    function jUpdateTS(theDivNo, vLogsNo) {
      var vLogsItem = document.getElementById('Txt_TS_' + theDivNo).value;
      var vParam = "vFunction=updateTS&vLogsNo=" + vLogsNo + "&vLogsItem=" + vLogsItem
      var vWs = WebService("LearnerReportCard_ws.asp", vParam)
      
      // put the new TS value into the link
      var innerHtml = "  <a href=\"javascript:open1Div('Div_TSedit_" + theDivNo + "'); hide1Div('Div_TSdisp_" + theDivNo + "');\">" + vLogsItem + "</a>"
  	  document.getElementById("Div_TSdisp_" + theDivNo).innerHTML = innerHtml; 
     
      hide1Div('Div_TSedit_' + theDivNo); 
      open1Div('Div_TSdisp_' + theDivNo);
    }


    function jUpdateBM(theDivNo, vLogsNo, vMembNo, vProgNo, vModsNo) {
      var vLogsItem = document.getElementById('Txt_BM_' + theDivNo).value;

      var vParam = "vFunction=updateBM&vLogsNo=" + vLogsNo + "&vLogsItem=" + vLogsItem
//    alert(vParam);
      var vWs = WebService("LearnerReportCard_ws.asp", vParam)
      
      // put the new BM value into the link
      var innerHtml = "  <a href=\"javascript:open1Div('Div_BMedit_" + theDivNo + "'); hide1Div('Div_BMdisp_" + theDivNo + "');\">" + vLogsItem + "</a>"
  	  document.getElementById("Div_BMdisp_" + theDivNo).innerHTML = innerHtml; 

      // call RTE to change Bookmark [public void SessionSetBookmark(int MemberID, int ModuleID, int ProgramID, string Bookmark)]
      var vParam = "MemberID=" + vMembNo + "&ProgramID=" + vProgNo + "&ModuleID=" + vModsNo + "&Bookmark=" + vLogsItem;
//    alert(vParam);
      var vWs    = WebService("/Gold/vuSCORM/Service.asmx/SessionSetBookmark", vParam);

      hide1Div('Div_BMedit_' + theDivNo); 
      open1Div('Div_BMdisp_' + theDivNo);
    }


    function Validate(theForm) {
      var vScoreDate  = theForm.vScoreDate.value;
      var vScoreProg  = theForm.vScoreProg.value;
      var vScoreId    = theForm.vScoreId.value;
      var vScoreValue = theForm.vScoreValue.value;

      var vParam      = "vFunction=addSC&vScoreDate=" + escape(vScoreDate) + "&vScoreProg=" + vScoreProg + "&vScoreId=" + vScoreId + "&vScoreValue=" + vScoreValue + "&vCust_Id=<%=vCust_Id%>&vMemb_No=<%=vMemb_No%>"
      var vWs         = WebService("LearnerReportCard_ws.asp", vParam)

      if (vWs == "inv date") {
        alert ("Please enter a valid Date.");
        theForm.vScoreDate.focus();
        return (false);
      }
      if (vWs == "inv prog") {
        alert ("Please enter a valid Program Id.");
        theForm.vScoreProg.focus();
        return (false);
      }
      if (vWs == "inv cust") {
        alert ("Cannot process - contact Systems. [Invalid Customer Id]");
        theForm.vScoreId.focus();
        return (false);
      }
      if (vWs == "inv mods") {
        alert ("Please enter a valid Assessment Id\nthat is a member of the above Program.");
        theForm.vScoreId.focus();
        return (false);
      }
      if (vWs == "inv val") {
        alert ("Please enter a valid Score from 1-100.");
        theForm.vScoreValue.focus();
        return (false);
      }
      return (true);
    }



    function jReset(vMembNo, vProgNo, vModsNo, vUrl, vMsg) {
      if (confirm(vMsg)) {
        // call the RTE to ensure these values are reset there as well
        var vParam = "MemberID=" + vMembNo + "&ProgramID=" + vProgNo + "&ModuleID=" + vModsNo;
        var vWs    = WebService("/Gold/vuSCORM/Service.asmx/SessionDelete", vParam);
        // then run the routine to erase them on the platform
        location.href = vUrl;
      }  
      else {
        return (false);
      }
    }  



  </script>
</head>

<body id="Vubix" onload="initTSelements(); initBMelements(); initSCelements();">

  <% Server.Execute vShellHi %>

  <h1><!--[[-->Learner Report Card<!--]]-->&nbsp;<!--[[-->for<!--]]-->&nbsp;<%=fMembName (vMemb_No)%></h1>

  <div class="c2" style="width:90%; margin:auto;">
    <!--[[-->This displays the complete learning activities sorted by Program Title. A Program must have been accessed for it to appear on the Report Card. For example, if 3 Programs were assigned or purchased but only the first was started, then only the first Program will appear here. Click on the Program Id to view that Program's activities (the first Program is always listed).<!--]]-->&nbsp; 
    <% If svMembLevel > 2 Then %><!--[[-->Click on a Best Score that represents a passing grade to yield a certificate.<!--]]--><% End If %>&nbsp; 
    <% If svMembLevel > 3 Then %>You can reset a non-passing assessment score (ie <span style="color:red">DELETE</span> all log entries for that assessment for the Learner) by clicking on the Reset link to the right.&nbsp; NOTE: You cannot reset passing scores. The link will only appear for non-passing Best Scores. Should you wish to reset the Time Spent in any module, click on the Time Spent link and enter a new value. If you wish to add a score for this Learner in a Program that appears in the Report Card below, click on the link below.<% End If %>
  </div>
  
  <div style="margin:20px; text-align:center;">
    <a class="c3" href="javascript:openDivs()"><!--[[-->Show All Details<!--]]--></a> <%=f10%>
    <a class="c3" href="javascript:hideDivs()"><!--[[-->Hide All Details<!--]]--></a>
  </div>    


  <% If Not bInfoPage And svMembLevel > 3 Then %>
	<div id="Div_SCform">			
    <a name="AddScore"></a>
    <form onsubmit="return Validate(this)" name="fForm" method="POST" action="LearnerReportCard2.asp">
      <table class="table">
        <tr>
          <td colspan="2"><h2>Add a Score</h2></td>
        </tr>
        <tr>
          <th style="white-space:nowrap">Date :</th>
          <td><input type="text" size="12" name="vScoreDate" maxlength="12"> In English format, ie <%=fFormatSqlDate(Now)%></td>
        </tr>
        <tr>
          <th style="white-space:nowrap">Program Id :</th>
          <td><input type="text" size="12" name="vScoreProg" id="vScoreProg" maxlength="7"> Do not change</td>
        </tr>
        <tr>
          <th style="white-space:nowrap">Assessment Id :</th>
          <td><input type="text" size="12" name="vScoreId" maxlength="6"> Assessment Id, ie 1234EN</td>
        </tr>
        <tr>
          <th style="white-space:nowrap">Score :</th>
          <td><input type="text" name="vScoreValue" size="4" maxlength="3">% (1-100) </td>
        </tr>
        <tr>
          <td colspan="2"style="text-align:center; padding:20px;">
            Click Return if you do NOT wish to add a score else click Update.<br>Note: you can reset/delete this value using the Reset Score feature below.
            <br /><br />
            <input type="button" value="Return" name="bReturn" id="bReturn"  class="button100" onclick="hide1Div('Div_SCform'); open1Div('Div_SClink');">
            <%=f10%>
            <input type="submit" value="Update" name="bAddScore" class="button100">                
          </td>
        </tr>
      </table>
      <input type="hidden" name="vMemb_No"       value="<%=vMemb_No%>">
      <input type="hidden" name="vStrDate"       value="<%=vStrDate%>">
      <input type="hidden" name="vEndDate"       value="<%=vEndDate%>">
      <input type="hidden" name="vCurList"       value="<%=vCurList%>">
      <input type="hidden" name="vNext"          value="<%=vNext%>">
      <input type="hidden" name="vEdit"          value="<%=vEdit%>">
      <input type="hidden" name="vCustId"        value="<%=vCustId%>">
      <input type="hidden" name="vFind"          value="<%=vFind%>">
      <input type="hidden" name="vFindId"        value="<%=vFindId%>">
      <input type="hidden" name="vFindFailing"   value="<%=vFindFailing%>">
      <input type="hidden" name="vFindCompleted" value="<%=vFindCompleted%>">
      <input type="hidden" name="vFindBookmarks" value="<%=vFindBookmarks%>">
      <input type="hidden" name="vFindFirstName" value="<%=vFindFirstName%>">
      <input type="hidden" name="vFindLastName"  value="<%=vFindLastName%>">
      <input type="hidden" name="vFindEmail"     value="<%=vFindEmail%>">
      <input type="hidden" name="vFindCriteria"  value="<%=vFindCriteria%>">
      <input type="hidden" name="vFindActive"    value="<%=vFindActive%>">
      <input type="hidden" name="vFormat"        value="<%=vFormat%>">    
            
    </form>
  </div>     
  <% End If %>


  <%
'   Response.Write "<br>vMemb_No " & vMemb_No
'   Response.Write "<br>svLang " & svLang
'   Response.Write "<br>vStrDate " & vStrDate & " - " & cDate(vStrDate)
'   Response.Write "<br>vEndDate " & vEndDate & " - " & cDate(vEndDate)

    sOpenCmd
    With oCmd
      .CommandText = "spLearnerReportCard"
      .Parameters.Append .CreateParameter("@Memb_No", adInteger, adParamInput,    , vMemb_No)
      .Parameters.Append .CreateParameter("@Lang",    adVarChar, adParamInput,   2, svLang)
      .Parameters.Append .CreateParameter("@StrDate", adDate,    adParamInput,    , vStrDate)
      .Parameters.Append .CreateParameter("@EndDate", adDate,    adParamInput,    , fFormatSqlDate(DateAdd("d", 1, vEndDate)))
    End With
    Set oRs = oCmd.Execute()
  
    If Not oRs.Eof Then
      sOpenDb2
      bDiv = True

      '...read until either eof or end of group
      Do While Not oRs.Eof     

        '..what is the passing grade
        sGetProg (oRs("Prog_Id"))
        If vProg_AssessmentScore = .01 Then
          vPassScore = 0
        ElseIf vProg_AssessmentScore > .01 Then
          vPassScore = vProg_AssessmentScore    
        ElseIf vCust_AssessmentScore = .01 Then
          vPassScore = 0
        ElseIf vCust_AssessmentScore > .01 Then
          vPassScore = vCust_AssessmentScore    
        Else
          vPassScore = .8
        End If
        vPassScore = vPassScore * 100
  
        '...grab the first program to turn on that DIV (unless aready set by reset above)
        If vDiv = "" Then vDiv = oRs("Prog_Id") 
  %>

  <table class="table">
    <tr>
      <th class="rowshade" style="width:10%; text-align:left;">&nbsp;<a href="javascript:toggle('Div_<%=oRs("Prog_Id")%>')"><%=oRs("Prog_Id")%></a> </th>
      <th class="rowshade" style="width:60%; text-align:left;white-space:nowrap" colspan="2"><% = oRs("Prog_Title") %></th>
      <th class="rowshade" style="width:30%; text-align:right;white-space:nowrap">
		    <% If Not bInfoPage And svMembLevel > 3 Then %>
        <a class="red" onclick="open1Div('Div_SCform'); setSCprog('<%=oRs("Prog_Id")%>')" href="#AddScore">Add a Score</a>&nbsp;&nbsp;
        <% End If %>
      </th>
    </tr>
    <%
    		vSql = " SELECT Ecom_Issued, Ecom_Expires, Ecom_CardName, Ecom_FirstName, Ecom_LastName"_
    		     & " FROM Ecom WITH (NOLOCK) "_
    		     & " WHERE (Ecom_MembNo = " & vMemb_No & ") AND (Ecom_Programs = '" & oRs("Prog_Id") & "') "
'   		sDebug
    		Set oRs2 = oDb2.Execute(vSql)
    		If Not oRs2.Eof Then
    		  vPurchaser = oRs2("Ecom_CardName")  : If Len(vPurchaser) = 0 Then vPurchaser = oRs2("Ecom_FirstName") & " " & oRs2("Ecom_LastName")
    		  vPurchased = oRs2("Ecom_Issued")  
    		  vExpires   = oRs2("Ecom_Expires")
    %>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;<!--[[-->Purchaser:<!--]]--> <%=vPurchaser%></td>
      <td colspan="2"><!--[[-->Purchased:<!--]]--> <%=fFormatDate(vPurchased)%></td>
      <td><!--[[-->Expires:<!--]]--> <%=fFormatDate(vExpires)%></td>
    </tr>
    <%
    		End If
    		Set oRs2 = Nothing
    %>
  </table>

  <div class="div" id="Div_<%=oRs("Prog_Id")%>" style="padding-top: 5px; padding-bottom: 20px">
    <table class="table">
      <tr>
        <th style="width:40%; white-space:nowrap; text-align:left;" colspan="2" ><!--[[-->Module/Assessment<!--]]--></th>
        <th style="width:10%; white-space:nowrap; text-align:center;"><!--[[-->Time<!--]]--><br><!--[[-->Spent<!--]]--><br><span style="font-weight: 400">(<!--[[-->Mins<!--]]-->)</span></th>
        <th style="width:10%; white-space:nowrap; text-align:center;"><!--[[-->Book<!--]]--><br><!--[[-->Mark<!--]]--></th>
        <th style="width:10%; white-space:nowrap; text-align:center;">#<br><!--[[-->Attempts<!--]]--></th>
        <th style="width:10%; white-space:nowrap; text-align:center;"><!--[[-->Best<!--]]--> <br><!--[[-->Score<!--]]--> %<br><span style="font-weight: 400">(<!--[[-->Cert<!--]]-->)</span></th>
        <th style="width:10%; white-space:nowrap; text-align:center;"><!--[[-->Best<!--]]--> <br><!--[[-->Score<!--]]--> <br>Date</th>
        <th style="width:10%; white-space:nowrap; text-align:center;"><% If ((svMembLevel = 4 And vCust_ResetLearners) Or svMembManager Or svMembLevel = 5) Then %>&nbsp;Reset&nbsp;<br>&nbsp;Score&nbsp;<% End If %>        </th>
      </tr>
      <%
        vMods = Trim(oRs("Prog_Mods"))
        If Len(oRs("Prog_Exam")) > 2 Then 
          vMods = vMods & " " & Mid(oRs("Prog_Exam"), 22, 6) & "_EXAM" '...add on _EXAM so we know not to display the modid - actually an exam id (it will be 11 characters long)
        End If
        If Len(oRs("Prog_Assessment")) > 2 Then 
          If Instr(vMods, oRs("Prog_Assessment")) = 0 Then        
            vMods = vMods & " " & oRs("Prog_Assessment") 
          End If
        End If

        aMods = Split(vMods)
        For vModCnt = 0 To Ubound(aMods)

          isExam = fIf(Instr("_EXAM", aMods(vModCnt)) > 0, True, False)
          vModId = Replace(aMods(vModCnt), "_EXAM", "")

'         sGetMods (Left(aMods(vModCnt), 6)) '...this was added to determine if there's a cert for the launch module
          sGetMods (vModId) '...this was added to determine if there's a cert for the launch module

'         If Len(aMods(vModCnt)) = 11 Then 
          If isExam Then 
            vTimeSpent  = 0
'           vTitle = fExamTitle(Left(aMods(vModCnt), 6))
            vTitle = fExamTitle(vModId)
          Else  
'           j = spLogsTimeSpent (vMemb_No, oRs("Prog_Id"), Left(aMods(vModCnt), 6), vStrDate, vEndDate)
            j = spLogsTimeSpent (vMemb_No, oRs("Prog_Id"), vModId, vStrDate, vEndDate)
            aTimeSpent  = Split(j, "|")
            vLogs_No    = Clng(aTimeSpent(0))
            vTimeSpent  = Clng(aTimeSpent(1))         
            vTitle      = vMods_Title
          End If
          
'         aBest = Split(spLogsBestValues(vMemb_No, Left(aMods(vModCnt), 6), vStrDate, vEndDate), "|")
          aBest = Split(spLogsBestValues(vMemb_No, vModId, vStrDate, vEndDate), "|")
          If Ubound(aBest) = 0 Then
            vBestScore = 0
            vBestDate = ""
          Else
            vBestScore = Cint(aBest(0))
            vBestDate  = aBest(1)
          End If

'         vNoAttempts = spLogsAttempts(vMemb_No, Left(aMods(vModCnt), 6), vStrDate, vEndDate)   
          vNoAttempts = spLogsAttempts(vMemb_No, vModId, vStrDate, vEndDate)   
      %>

      <tr>
        <td>
          <%' =fIf(IsExamLen(aMods(vModCnt)) < 11, aMods(vModCnt), "Exam")%>
          <% =fIf(Not IsExam, vModId, "Exam")%>
        </td>
        <td>
          <%'=fIf(Len(aMods(vModCnt)) < 11, fModsTitle(aMods(vModCnt)), "[Note: this exam (" & Left(aMods(vModCnt), 6) & ") may also appear below.]")%>
          <% =fIf(Not IsExam, fModsTitle(vModId), "[Note: this exam (" & Left(aMods(vModCnt), 6) & ") may also appear below.]")%>
        </td>
        <td style="text-align:center;">
          <% If vTimeSpent > 0 Then %> 
            <% If svMembLevel < 4 Then %>
            <%   =vTimeSpent%>
            <% Else %>
            <%   vTSno = vTSno + 1%>            
              <div id="Div_TSedit_<%=vTSno%>">
                <form>
                  <input type="text" id="Txt_TS_<%=vTSno%>" size="2" name="vTimeSpent" value="<%=vTimeSpent%>" style="text-align: right"> 
                  <input onclick="jUpdateTS(<%=vTSno%>, <%=vLogs_No%>);" type="button" value="Reset" id="bUpdateTS" class="button">
                </form>
              </div>
              <div id="Div_TSdisp_<%=vTSno%>">
                <a href="javascript:open1Div('Div_TSedit_<%=vTSno%>'); hide1Div('Div_TSdisp_<%=vTSno%>');"><%=vTimeSpent%></a>
              </div>
            <% End If %>          
          <% End If %>         
        </td>
        <td style="text-align:center;">
          <%  
            If Len(aMods(vModCnt)) = 6 Then
              aBookmark = Split(spLogsBookmark(vMemb_No, aMods(vModCnt), vStrDate, vEndDate), "|")
              vLogs_No  = aBookmark(0)
              vBookmark = aBookmark(1)            
              If IsNumeric(vBookmark) Then 
              	vBookmark = Cint(vBookmark)
              Else
                vBookmark = 0
              End If
              If vBookmark > 0 Then  
                If svMembLevel < 4 Or Not IsNumeric(vBookmark) Then
                  Response.Write vBookmark
                Else
                  vBMno = vBMno + 1
          %>            
              <div id="Div_BMedit_<%=vBMno%>">
                <form>
                  <input type="text" id="Txt_BM_<%=vBMno%>" size="2" name="vBookmark" value="<%=vBookmark%>" style="text-align: right" maxlength="3"> 
                  <input onclick="jUpdateBM(<%=vBMno%>, <%=vLogs_No%>, <%=oRs("Memb_No")%>, <%=vProg_No%>, <%=vMods_No%>);" type="button" value="Reset" id="bUpdateTS" class="button">
                </form>
              </div>
              <div id="Div_BMdisp_<%=vBMno%>">
                <a href="javascript:open1Div('Div_BMedit_<%=vBMno%>'); hide1Div('Div_BMdisp_<%=vBMno%>');"><%=vBookmark%></a>
              </div>
          <%      
                End If
              End If
            End If 
          %>  
        </td>
        <td style="text-align:center;"><%=fIf(vNoAttempts=0, "", vNoAttempts)%></td>
        <td style="text-align:center;">  
          <% 
             '...if there's a score then output something
	         	 If vBestScore > 0 Then 
	         	   '...if it's not a passing score, or we do not generate a cert, then just display score 
               If vBestScore < vPassScore Or Not vMods_VuCert Then
                 Response.Write vBestScore
	         	   '...display the certificate
               Else 
          %>     
               <a href="#" onclick="fullScreen('<%=fCertificateUrl(oRs("Memb_FirstName"), oRs("Memb_LastName"), vBestScore, vBestDate, Left(aMods(vModCnt), 6), vTitle, "", "", "", oRs("Prog_Id"), "", "", "")%>')"><%=vBestScore%></a> 
          <% 
          		End If 
          	'...if score is negative display zero (-1 forces 0)	
            ElseIf vBestScore < 0 Then
               Response.Write "0"
            '...else display space
            Else
               Response.Write "&nbsp;"
            End If 
          %>
        </td>
        <td style="white-space:nowrap; text-align:center;"><%=fFormatDate(vBestDate)%></td>
        <td style="white-space:nowrap; text-align:center;">
          <% 
            '...only allow reset of failed scores (even if no score is showing), ie do not allow if the best score is >= passing score
            If ((svMembLevel = 4 And vCust_ResetLearners) Or svMembManager Or svMembLevel = 5) Then 
          %>
            <input onclick="jReset(<%=oRs("Memb_No")%>, <%=vProg_No%>, <%=vMods_No%>, 'LearnerReportCard2.asp?vStrDate=<%=vStrDate%>&vEndDate=<%=vEndDate%>&vCurList=<%=vCurList%>&vNext=<%=vNext%>&vCustId=<%=vCustId%>&vFind=<%=vFind%>&vFindId=<%=vFindId%>&vFind=<%=vFind%>&vFindFailing=<%=vFindFailing%>&vFindCompleted=<%=vFindCompleted%>&vFindFirstName=<%=fjUnQuote(vFindFirstName)%>&vFindLastName=<%=fjUnQuote(vFindLastName)%>&vFindEmail=<%=vFindEmail%>&vFindCriteria=<%=vFindCriteria%>&vFindActive=<%=vFindActive%>&vFormat=<%=vFormat%>&vReset=<%=oRs("Memb_No") & "|" & fjUnQuote(oRs("Memb_FirstName")) & "|" & fjUnQuote(oRs("Memb_LastName")) & "|" & Left(aMods(vModCnt), 6) & "|" & oRs("Prog_Id")%>','<!--[[-->Ok to Reset/Delete this Score?  You cannot reverse this action.<!--]]-->')" type="button" value="Reset" name="bReset" class="button">
          <%  
            End If 
          %>
        </td>  
      </tr>
      <%
        Next
      %>
    </table>
  </div>
  <%    
        oRs.MoveNext
      Loop 

      Set oRs = Nothing
      Set oCmd = Nothing

      sCloseDb
      sCloseDb2

    End If 
  %>

  <%

      '...any legacy exams?   
      '...this section was not modified for long mod Ids - assumption being they are all old mods Id
      vSql = " SELECT DISTINCT"_ 
           & "   Logs.Logs_MembNo, "_
           & "   MAX(Logs.Logs_Posted) AS [Best Date], "_
           & "   LEFT(Logs.Logs_Item, 6) AS [Exam Id], "_
           & "   MAX(RIGHT(Logs.Logs_Item, 3)) AS [Best Score], "_
           & "   V5_Base.dbo.TstH.TstH_Title AS [Exam Title], "_
           & "   MAX(SUBSTRING(Logs.Logs_Item, 8, 1)) AS [Attempts], "_
           & "   Memb.Memb_FirstName, "_
           & "   Memb.Memb_LastName, "_
           & "   Memb.Memb_No "_
  
           & " FROM"_  
           & "   Logs                  WITH (NOLOCK) LEFT OUTER JOIN "_ 
           & "   Catl_Prog             WITH (NOLOCK) ON Logs.Logs_AcctId = Catl_Prog.Catl_Prog_AcctId AND LEFT(Logs.Logs_Item, 6) <> Catl_Prog.Catl_Prog_ExamId INNER JOIN "_ 
           & "   V5_Base.dbo.TstH      WITH (NOLOCK) ON LEFT(Logs.Logs_Item, 6) = V5_Base.dbo.TstH.TstH_Id INNER JOIN "_
           & "   Memb                  WITH (NOLOCK) ON Logs.Logs_MembNo = Memb.Memb_No "_
  
           & " WHERE"_ 
           & "   (LEN(Logs.Logs_Item) = 12) AND (RIGHT(Logs.Logs_Item, 3) <> '000') AND (Logs.Logs_Type = 'T') AND (Logs.Logs_AcctId = '" & svCustAcctId & "') AND (Logs.Logs_Membno = " & vMemb_No & ") "_
           & " GROUP BY "_ 
           & "   Logs.Logs_MembNo, LEFT(Logs.Logs_Item, 6), V5_Base.dbo.TstH.TstH_Title, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_No "_
           & " ORDER BY "_ 
           & "   V5_Base.dbo.TstH.TstH_Title "
'     sDebug     
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then
  %>
  <table class="table">
    <tr>
      <th style="text-align:left" class="rowshade">&nbsp;<a href="javascript:toggle('Div_P0000XX')">P0000XX</a>&nbsp;</th>
      <th style="text-align:left" class="rowshade">Note: the following are from the older format Exams that may also appear in a program above.&nbsp;</th>
    </tr>
  </table>
  <div class="div" id="Div_P0000XX" style="padding-top: 5px; padding-bottom: 20px">
    <table style="border-collapse: collapse" border="1" bordercolor="#DDEEF9" width="90%">
      <tr>
        <th colspan="4" style="white-space:nowrap">
        <!--[[-->Exam<!--]]--></th>
        <th style="white-space:nowrap">#<br>
        <!--[[-->Attempts<!--]]--></th>
        <th style="white-space:nowrap">
        <!--[[-->Best<!--]]--><br><!--[[-->Score<!--]]--> %<br><span style="font-weight: 400">(<!--[[-->Cert<!--]]-->)</span></th>
        <th style="white-space:nowrap">
        <!--[[-->Best<br>Score<!--]]--><br>Date</th>
        <th style="white-space:nowrap">
          <% If ((svMembLevel = 4 And vCust_ResetLearners) Or svMembManager Or svMembLevel = 5) Then %>&nbsp;Reset&nbsp;<br>&nbsp;Score&nbsp;<% End If %></th>
      </tr>
      <%
        Do While Not oRs.Eof
          bDiv = True
          vBestScore = Cint(oRs("Best Score"))
      %>
      <tr>
        <td colspan="3">&nbsp;<%=oRs("Exam Id")%>&nbsp;</td>
        <td width="50%"><%=oRs("Exam Title")%></td>
        <td><%=oRs("Attempts")%></td>
        <td style="white-space:nowrap">

          <% 
	         	 If vBestScore > 0 Then 
	         	   '...if it's not a passing score, or we do not generate a cert, then just display score 
               If vBestScore < vPassScore Or Not vMods_VuCert Then
                 Response.Write vBestScore
	         	   '...display the certificate
               Else 
          %>     
            <a <%=fstatx%> href="javascript:fullScreen('<%=fCertificateUrl(oRs("Memb_FirstName"), oRs("Memb_LastName"), vBestScore, oRs("Best Date"), oRs("Exam Id"), oRs("Exam Title"), "", "", "", oRs("Exam Id"), "", "", "")%>')"><%=vBestScore%></a>
          <% 
          		End If 
            ElseIf vBestScore < 0 Then
               Response.Write "0"
            Else
               Response.Write "&nbsp;"
            End If 
          %>
  
       </td>
        <td style="white-space:nowrap"><%=fFormatDate(oRs("Best Date"))%></td>
        <td>
        
          <% If ((svMembLevel = 4 And vCust_ResetLearners) Or svMembManager Or svMembLevel = 5) And oRs("Attempts") > 0  Then %>
            <a href="javascript:jconfirm('LearnerReportCard2.asp?vStrDate=<%=vStrDate%>&vEndDate=<%=vEndDate%>&vCurList=<%=vCurList%>&vNext=<%=vNext%>&vCustId=<%=vCustId%>&vFind=<%=vFind%>&vFindId=<%=vFindId%>&vFind=<%=vFind%>&vFindFailing=<%=vFindFailing%>&vFindCompleted=<%=vFindCompleted%>&vFindFirstName=<%=vFindFirstName%>&vFindLastName=<%=vFindLastName%>&vFindEmail=<%=vFindEmail%>&vFindCriteria=<%=vFindCriteria%>&vFindActive=<%=vFindActive%>&vFormat=<%=vFormat%>&vReset=<%=oRs("Memb_No") & "|" & oRs("Memb_FirstName") & "|" & oRs("Memb_LastName") & "|" & oRs("Exam Id") & "|P0000XX" %>','Ok to Reset/Delete this Score?  You cannot reverse this action.')">Reset</a> 
          <% End If %>
        </td>
      </tr>
  <%    
        oRs.MoveNext
      Loop 
      Set oRs = Nothing
      sCloseDb
  %>
    </table>
  </div>

  <%  
    End If
  %>  



  <!-- This will open either the first or a Reset Div and turn off the Details button, unless there are no log items on file -->

  <% If bDiv Then %>
  <script>
//  open1Div('Div_<%=vDiv%>');
  </script>
  <% Else %> 
  <p class="red"><!--[[-->There are no learning activities on file.<!--]]--></p>
  <% End If %>

  <% If Not bInfoPage Then %>

    <div style="text-align:center; margin:20px;">       
    <% If Len(vNext) > 0 Then %>
    <input type="button" onclick="location.href='<%=vNext%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button100"><%=f10%>
    <% End If %>    
    <% vCurList = 0 %>
    <% 
       vUrl = "LearnerReportCard1.asp?" _
            & "vStrDate="        & vStrDate _
            & "&vEndDate="       & vEndDate _
            & "&vCurList="       & vCurList _
            & "&vNext="          & vNext _
            & "&vEdit="          & vEdit _
            & "&vCustId="        & vCustId _
            & "&vFind="          & vFind _
            & "&vFindId="        & vFindId _
            & "&vFindFailing="   & vFindFailing _
            & "&vFindCompleted=" & vFindCompleted _
            & "&vFindBookmarks=" & vFindBookmarks _
            & "&vFindFirstName=" & fjUnQuote(vFindFirstName) _
            & "&vFindLastName="  & fjUnQuote(vFindLastName) _
            & "&vFindEmail="     & vFindEmail _
            & "&vFindCriteria="  & vFindCriteria _
            & "&vFindActive="    & vFindActive _
            & "&vFormat="        & vFormat         
        ' Response.Write "<p align='left'>" & Replace(vUrl, "&", "<br>&") & "</p>"     
    %>
    <input type="button" id="bReturn" onclick="location.href='<%=vUrl%>'" value="<%=bBack%>" name="bBack" class="button100"><%=f10%>
    <% 
       vUrl = "LearnerReportCard.asp?" _
            & "vStrDate="        & vStrDate _
            & "&vEndDate="       & vEndDate _
            & "&vCurList="       & vCurList _
            & "&vNext="          & vNext _
            & "&vEdit="          & vEdit _
            & "&vCustId="        & vCustId _
            & "&vFind="          & vFind _
            & "&vFindId="        & vFindId _
            & "&vFindFailing="   & vFindFailing _
            & "&vFindCompleted=" & vFindCompleted _
            & "&vFindBookmarks=" & vFindBookmarks _
            & "&vFindFirstName=" & fjUnQuote(vFindFirstName) _
            & "&vFindLastName="  & fjUnQuote(vFindLastName) _
            & "&vFindEmail="     & vFindEmail _
            & "&vFindCriteria="  & vFindCriteria _
            & "&vFindActive="    & vFindActive _
            & "&vFormat="        & vFormat         
        ' Response.Write "<p align='left'>" & Replace(vUrl, "&", "<br>&") & "</p>"     
    %>
    <input type="button" id="bRestart" onclick="location.href='<%=vUrl%>'" value="<%=bRestart%>" name="bRestart" class="button100">
    </div>

    <h3><%=vCust_Id & "  (" & vCust_Title & ")"%></h3>

  <% End If %>
    
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>