<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->

<%
  Dim vNext, vActv, vCurN, vOutp, vStrD, vEndD, vProg, vMods, vAssg, vPass, vLNam, vMemo, vGrou, vSave, vFrom, vDifNo
  Dim vCurMemb, vCurProg, vCurMods, vCurSess, vCurObjs
  Dim vPrvMemb, vPrvProg, vPrvMods, vPrvSess, vPrvObjs
  Dim vActive, vBg, vPassword, vDate, vCompleted, vUrl
  Dim vGroupId, vMembNo, vMembId, vMembActive, vMembLevel, vFirstName, vLastName, vEmail, vMembMemo, vProgNo, vProgId, vProgTitle, vModsNo, vModsId, vModsTitle, vModsVuCert, vTimeSpent, vScore, vLastDate, vExpired, vRecurs, vSessionId, vObjectiveId, vEcom, vMultiSCO

  vFrom = Request("vFrom") 
  vActv = Request("vActv") 
  vCurN = fDefault(Request("vCurN"), 0)
  vOutp = Request("vOutp") 
  vStrD = Request("vStrD") 
  vEndD = Request("vEndD") 
  vProg = Request("vProg") 
  vMods = Request("vMods") 
  vAssg = Request("vAssg") 
  vPass = Request("vPass")
  vLNam = Request("vLNam")
  vMemo = Request("vMemo")
  vGrou = Request("vGrou")
  vSave = Request("vSave")

  '...if "back" then determine new starting point
  If Request("bBack").Count > 0 And vCurN > 100 Then 
    If vCurN  Mod 100 = 0 Then
      vCurN = vCurN - 200
    Else
      vCurN = Cint(vCurN / 100) * 100 - 100
    End If
  End If
          
  Function fScore(vVal)
    If IsNumeric(vScore) Then 
      fScore = vVal
    Else
      fScore = " "
    End If
  End Function
	

  Function fFirstMemb(vVal)
    If vCurMemb = vPrvMemb Then
      fFirstMemb = " "
    Else
      fFirstMemb = vVal
			vPrvProg = ""
			vPrvMods = ""
    End If
  End Function

  Function fFirstProg(vVal)
    If vCurProg = vPrvProg Then
      fFirstProg = " "
    Else
      fFirstProg = vVal
    End If
  End Function

  Function fFirstMods(vVal)
'   If vCurMods = vPrvMods Then
    If vCurMods = vPrvMods And vCurProg = vPrvProg Then
      fFirstMods = " "
    Else
      fFirstMods = vVal
    End If
  End Function

  Function fFirstSess(vVal)
    If vCurSess = vPrvSess Then
      fFirstSess = " "
    Else
      fFirstSess = vVal
    End If
  End Function

  Function fFirstObjs(vVal)
    If IsNull(vCurObjs) Or vCurObjs = vPrvObjs Then
      fFirstObjs = " "
    Else
      fFirstObjs = vVal
    End If
  End Function

  Function fActive()
    If vCurMemb <> vPrvMemb Then
	    fActive = fYN (vActive)
    Else
      fActive = " "
    End If
  End Function

  Function fCompleted()	        
    fCompleted = fYN (vCompleted)
  End Function

  Function fExpired(vExpired, vRecurs)
    If vRecurs = 0 Then
      fExpired = "na"
    ElseIf IsDate(vExpired) Then  
      fExpired= fYN (True)
    Else
      fExpired= fYN (False)
    End If
  End Function

%>

<html>

<head>
  <title>RTE_History_O</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="//code.jquery.com/jquery-1.9.1.js"></script>
  <script src="//code.jquery.com/ui/1.10.2/jquery-ui.js"></script>
  <script src="/V5/Inc/jQuery.draggable.js"></script>
  <script src="/V5/Inc/jQueryC.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script> 
    function jRestart() {
      location.href = "<%=fIf(vFrom = "menu", "RTE_History.asp", vFrom) %>";
    }

    // store row no where last window was opened (-1 mean no open window)
    var lastRowNo  = -1;

    // underline a row when launching an action item
    function underLineRed(rowNo) {        
      $("#R_" + rowNo)[0].className = "underLineRed"; //set new red line
      lastRowNo = rowNo; // store new row so it can be removed
    }

    function underLineOff() {
      if (lastRowNo != -1 ) {
        $("#R_" + lastRowNo)[0].className = "underLineOff";
        lastRowNo = -1;
      }
    }

    function jTitle(vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase', '//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0', 'name', 'flashVars', 'width', '265', 'height', '85', 'align', 'middle', 'id', 'flashVars', 'src', '/V5/Images/Titles/VuTitles', 'FlashVars', vParm, 'quality', 'high', 'bgcolor', '#ffffff', 'allowscriptaccess', 'sameDomain', 'allowfullscreen', 'false', 'pluginspage', '///go/getflashplayer', 'movie', '/V5/Images/Titles/VuTitles');
    }

      // this contains the Y axis of the click where we position the divOuter
      var pageY = 20; 


    $(function () {

      $("#divNotes").hide();

      // default focus to either Restart or Next (not Back)

      // make the div draggable
      $("#divOuter").draggable();

      // Disable caching of AJAX responses
      $.ajaxSetup({ cache: false });

      // Set divNotes to state last save
      $("#divNotes").css("display", $.cookie("History_<%=svCustId%>_divNotes_cssDisplay"));

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
      }

      window.onscroll = function () {
        determinePageY();
      }

      //  when page gets focus (after it's lost it, ie if !bodyFocus - set to false when window is launched
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
    function openWindow(rowNo, parm1, parm2, parm3, parm4, parm5) {        
      closeWindow();                                  // close previous / reset window        
      underLineRed(rowNo);                            // underline current line
      divOff("flash")                                 // turn off since we cannot render on top of it

			if (reNumeric.test(parm3) == null) {
				pageY = parm3;                                // RTE_ModsStats launches a cert and we want to have it rendered at the same Y position as the original click, so if we get a value in parm3 we use that for the click/y top position
			}

      $("#divOuter")[0].style.width = 800 + "px";     // define default width  (may be modified below editor)
      $("#divOuter")[0].style.height = 600 + "px";    // define default height (may be modified below)

      switch (parm1) {
        case "Mods" : $("#divInner").load("Module.asp?vMods_Id="      + parm2); divOn("divInner"); break;                                                                       // program description
        case "P"    : $("#divInner").load("RTE_ProgDesc.asp?vProgId=" + parm2); divOn("divInner"); break;                                                                       // program description
        case "M"    : $("#divInner").load("RTE_ModsDesc.asp?vModsId=" + parm3); divOn("divInner"); break;                                                                       // module description
        case "S"    : $("#divInner").load("RTE_ModsStat.asp?vProgNo=" + parm2 + "&vModsId=" + parm3 + "&vProgId=" + parm4); divOn("divInner"); break;                           // status
        case "Sync" : $("#divInner").load("RTE_Sync.asp?vSessionId="  + parm2 + "&vMembId=" + parm3 + "&vProgId=" + parm4 + "&vModsId=" + parm5); divOn("divInner"); break;     // sync                                               
        case "Ecom" : $("#divInner").load("RTE_Ecom.asp?vProgId="     + parm2 + "&vMembNo=" + parm3 ); divOn("divInner"); break;                          
        case "Cert" : 
          $("iframe")[0].src = parm2; 
          divOn("ifrInner"); 
          $("#divOuter")[0].style.width  = 800 + "px"; 
          $("#divOuter")[0].style.height = 600 + "px"; 
          $("#tabOuter")[0].style.height = 550 + "px"; 
          break;                                                                                                                                                                // certificate service
        case "Edit" :                                                                                                                                                           // session editor service
          $("iframe")[0].src = "/Gold/vuSCORMAdmin/SessionQuickEdit.aspx?MembNo=<%=svMembNo%>&SessionID=" + parm2 + "&memberID=" + parm3 + "&moduleID=" + parm4 + "&programID=" + parm5; 
          divOn("ifrInner"); 
          $("#divOuter")[0].style.width = 330 + "px"; 
          $("#divOuter")[0].style.height = 550 + "px"; 
          $("#tabOuter")[0].style.height = 500 + "px"; 
          $(".actionAlert").show();
          break;   
        case "Hist" :                                                                                                                                                           // session editor service
          $("iframe")[0].src = "/Gold/vuSCORMAdmin/Default.aspx?MembNo=<%=svMembNo%>&memberID=" + parm2; 
          divOn("ifrInner"); 
          $("#divOuter")[0].style.width = "90%"; 
          $("#divOuter")[0].style.height = 800 + "px"; 
          $("#tabOuter")[0].style.height = 750 + "px"; 
          break;   
        case "MultiSCO" :                                                                                                                                                       // session multisco report
          $("iframe")[0].src = "/Gold/vuclientreporting/Reportexport.aspx?AccountID=<%=Right(svCustId, 4)%>&reportfile=repLearnerMultiSCODetailsExport.frx&Type=CSV&SessionID=" + parm2; 
          divOn("ifrInner"); 
          $("#divOuter")[0].style.width = "90%"; 
          $("#divOuter")[0].style.height = 800 + "px"; 
          $("#tabOuter")[0].style.height = 750 + "px"; 
          break;   
      }

      $("#divOuter")[0].style.left = "50px";               // set left
      $("#divOuter")[0].style.top = pageY + "px";          // set top to where last click occurred
      $("#divOuter")[0].style.position = "absolute";       // make absolute

      divOn("divOuter");
    }

    function closeWindow() {
      divOff("divOuter");
      divOff("divInner");
      divOff("ifrInner");
      divOn("flash");
      underLineOff();                            
    }  


    $(function () { toggleSave("divNotes")});

    // save toggle state for when page reloads
    function toggleSave(ele) {
      $("#"+ele).toggle();
      $.cookie("History_<%=svCustId%>_divNotes_cssDisplay", $("#"+ele).css("display"));
      if ($("#divNotes")[0].style.display == "block") {
        $("#controlNotes").text("/*--{[--*/Hide Notes/*--]}--*/");
      } else {
        $("#controlNotes").text("/*--{[--*/Show Notes/*--]}--*/");
      }
    }

    // expand or contract test
    var less = <% If svMembLevel=5 Then %>true<% Else %>false<% End If %>;
    $(function () {renderExpandable()})
    $("#expandable").click(function () {renderExpandable()})

    function renderExpandable() {
      if (less) {
        $(".expandable").addClass("less");
        $("#expandable").text("/*--{[--*/More Text/*--]}--*/");
      } else {
        $(".expandable").removeClass("less");
        $("#expandable").text("/*--{[--*/Less Text/*--]}--*/");
      }
      less = !less;
    }
  </script>
  <style>
    .table tr:hover { background-color: #DDEEF9; }
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

  <h1 style="text-align: center"><!--[[-->Learner Report Card<!--]]--></h1>
  <div style="text-align: center">
    [ <a href="#" class="green" onclick="toggleSave('divNotes')" id="controlNotes"><!--[[-->Details On/Off<!--]]--></a> ] 
    [ <a href="#" class="green" onclick="renderExpandable()" id="expandable"></a> ]
  </div>
  <div id="divNotes" style="text-align: left; margin: 30px;">
    <ul>
      <li><!--[[-->All assessment <b>Score</b>s are shown with the latest score first followed by previous attempts.<!--]]--></li>
      <li><!--[[--><b>Time Spent</b> is the total time in minutes expended while reviewing a module or taking an assessment.<br />Note however, that if the assessment is recurring, ie it must be taken every 6 months, then once completed the Time Spent is reset.<!--]]--></li>
      <li><!--[[-->A <b>Closed</b> (&quot;Yes&quot;) session means that a recurring assessment has been completed. If it is not recurring it will show &quot;na&quot;.<!--]]--></li>
      <li><!--[[-->A <b>Completed</b> (&quot;Yes&quot;) session is when an assessment is passed or content is deemed completed.<!--]]--></li>
      <li><!--[[-->Larger text fields, shown in italics, can be expanded or reduced in size by clicking the &quot;More Text&quot;/&quot;Less Text&quot; link at the top of the page.<!--]]--></li>
      <li><!--[[-->You can see any Learner&#39;s Profile by clicking on their<!--]]--> <%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%><!--[[-->. Note that any changes made to a learner profile will not be reflected on this report unless it is restarted.<!--]]--></li>
      <li><!--[[-->Actions:<!--]]-->
        <ul>
          <li><!--[[--><b>&quot;Cert&quot;</b> will generate a certificate for a passed assessment (appears once per module).<!--]]--></li>
          <% If svMembLevel > 3 Then %>
          <li><b>&quot;Ecom&quot;</b> will bring up any ecommerce details (appears once per program).</li>
          <li><b>&quot;Hist&quot;</b> will bring up the complete details of the learner&#39;s history (appears once per learner).</li>
          <li><b>&quot;Edit&quot;</b> will bring up a line Editor allowing you to modify the learner&#39;s selected session history. <strong>NOTE: you must RESTART this report to view any changes..</strong></li>
          <% End If %>
          <% If svMembLevel > 4 Then %>
          <li><b>&quot;Sync&quot; </b>deletes learner&#39;s LMS FX logs items and replaces them with RTE FX session items (appears once per learner). Note: the Sync occurs as soon as you click &quot;Sync&quot;.</li>
          <% End If %>
        </ul>
      </li>
    </ul>
  </div>
  <br />

  <table class="table">
    <tr>
      <th class="rowshade" style="width: 50px; text-align: center;">
        <!--[[-->Row<!--]]--></th>
      <th class="rowshade" style="text-align: left;">
        <!--[[-->Group<!--]]--></th>
      <th class="rowshade" style="text-align: left;">
        <!--[[-->Name<!--]]--></th>
      <th class="rowshade" style="text-align: left;"><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Active<!--]]--></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Program<!--]]--></th>
      <th class="rowshade" style="text-align: left; width: 200px; white-space: nowrap;">
        <!--[[-->Title<!--]]--></th>
      <th class="rowshade" style="text-align: left;">
        <!--[[-->Module<!--]]--></th>
      <th class="rowshade" style="text-align: left; width: 200px; white-space: nowrap;">
        <!--[[-->Title<!--]]--></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Time<!--]]--><br>
        <!--[[-->Spent<!--]]--></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Score<!--]]--></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Date<!--]]--></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Completed<!--]]--></th>
      <th class="rowshade" style="text-align: center;">
        <!--[[-->Closed<!--]]--></th>
      <th class="rowshade" style="text-align: center;" colspan="6">
        <!--- after edit restart -->
        <div class="actionAlert bgYellow" style="padding: 10px; display: none;">
          <input type="button" onclick="jRestart()" class="button200" style="color: red; font-weight: 600;" value="Restart to view changes!" name="bRestart">
        </div>
        <!--- normal restart -->
        <input type="button" onclick="jRestart()" class="button100" style="text-align: center;" value="<%=fIf(vFrom="menu", bRestart, bReturn) %>" name="bRestart">
      </th>
    </tr>

    <% 
	    vSql = "SELECT DISTINCT * FROM LogsR WHERE UserNo = " & svMembNo & " AND RowNo > " & vCurN
	    sOpenDb
	    Set oRs = oDb.Execute(vSql)
	
	    vPrvMemb = ""
	    vPrvProg = ""
	    vPrvMods = ""
	    vPrvSess = ""
	    vPrvObjs = ""
	
	    '...read until either eof or end of group
	    Do While Not oRs.Eof
	
	      vGroupId       = oRs("GroupId")
	      vMembNo        = oRs("MembNo")
	      vMembId        = oRs("MembId")
	      vMembMemo      = oRs("MembMemo")
	      vMembActive		 = oRs("MembActive")
	      vMembLevel 		 = oRs("MembLevel")
	      vFirstName     = oRs("FirstName")
	      vLastName      = oRs("LastName")
	      vEmail         = oRs("Email")
	
	      vProgNo        = oRs("ProgNo")
	      vProgId        = oRs("ProgId")
	      vProgTitle     = oRs("ProgTitle")
	      vModsNo        = oRs("ModsNo")
	      vModsId        = oRs("ModsId")
	      vModsTitle     = oRs("ModsTitle")
	      vModsVuCert    = oRs("ModsVuCert")
	
	      vTimeSpent     = oRs("TimeSpent")
	      vScore    		 = oRs("Score")
	      vLastDate      = oRs("LastDate")
	      vExpired       = oRs("Expired")
	      vCompleted     = oRs("Completed")
	      vRecurs				 = oRs("Recurs")
	      vSessionId		 = oRs("SessionId")
	      vObjectiveId	 = oRs("ObjectiveId")
	      vEcom	         = oRs("Ecom")
	      vMultiSCO      = oRs("MultiSCO")
	        
	      vCurMemb 			 = vMembId        
	      vCurProg 			 = vProgId        
	      vCurMods  	   = vModsId        
	      vCurSess  	   = fOkValue(vSessionId)
	      vCurObjs  	   = fOkValue(vObjectiveId)

        If varType(vTimeSpent) = vbNull Then 
          vTimeSpent = ""
        ElseIf Clng(vTimeSpent) = 0 Then 
          vTimeSpent = ""
        End If


	      If IsNull(vCompleted) Then vCompleted = True
	
	      vNext = Server.UrlEncode("RTE_History_F.asp?vCurN=" & Request("vCurN"))  '...get the starting value of the page so it returns to the identical page

	      If svMembInternal Or svMembManager Then
	        vPassword = "<a  href='User" & fGroup & ".asp?vMembNo=" & vMembNo & "&vNext=" & vNext & "'>" & fLeft(vMembId, 16) & "</a>" & fIf(vMembLevel = 3, " *", "") & fIf(vMembLevel = 4, " **", "")
	      ElseIf svMembLevel > vMembLevel Then
	        vPassword = vMembId & fIf(vMembLevel = 3, " *", "") & fIf(vMembLevel = 4, " **", "")
        Else
          vPassword = "**********"
	      End If

        If vCurMemb <> vPrvMemb Or vCurMods <> vPrvMods Then
'           vBg = fIf (vDifNo Mod 2 = 1, " style='background-color: #B0CAF0'", "")
          vBg = fIf (vDifNo Mod 2 = 1, "#B0CAF0", "white")
          vDifNo = vDifNo + 1
        End If 

    %>


    <tr id="R_<%=vCurN %>">
      <td style="width: 50px; text-align: center"><%=vCurN%></td>
      <td style="text-align: left">
        <div class="expandable"><%=fFirstMemb(vGroupId)%></div>
      </td>
      <td style="text-align: left">
        <div class="expandable"><%=fFirstMemb(vFirstName & " " & vLastName)%></div>
      </td>
      <td style="text-align: left">
        <div class="expandable"><%=fFirstMemb(vPassword)%></div>
      </td>
      <td style="text-align: center; white-space: nowrap;"><%=fFirstMemb(fYN(vMembActive))%></td>
      <td style="text-align: center; white-space: nowrap;"><%=fFirstProg(vProgId)%></td>
      <td style="text-align: left;">
        <div class="expandable"><%=fFirstProg(Trim(vProgTitle))%></div>
      </td>
      <td style="text-align: left; white-space: nowrap;"><%=fFirstMods(vModsId & " " & fModsType(vModsId))%></td>
      <td style="text-align: left;">
        <div class="expandable"><%=fFirstMods(Trim(vModsTitle))%></div>
      </td>
      <td style="text-align: center; white-space: nowrap;"><%=fFirstSess(vTimeSpent)%></td>
      <td style="text-align: center; white-space: nowrap;"><%=fScore(vScore)%></td>
      <td style="text-align: center; white-space: nowrap;"><%=fFormatDate(vLastDate)%></td>
      <td style="text-align: center; white-space: nowrap;"><%=fFirstSess(fCompleted())%></td>
      <td style="text-align: center; white-space: nowrap;"><%=fFirstSess(fExpired(vExpired, vRecurs))%></td>

      <!-- render unique session buttons -->
      <td style="text-align: center; white-space: nowrap;">
        <%
 		    If IsNumeric(vCurSess) Then
          If vCurMods <> vPrvMods Or vCurProg <> vPrvProg Then
            If vModsVuCert And vCompleted Then 
              vUrl = fCertificateUrl(vFirstName, vLastName, vScore, vLastDate, vModsId, vModsTitle, "", "", "", vProgId, "", vSessionId, vEmail)
        %>
        <input class="button040" title="Certificate for the Assessment" type="button" onclick="openWindow(<%=vCurN %>, 'Cert', '<%=vUrl%>')" value="Cert">
        <%
            End If
          End If
        End If 
        %>
      </td>

      <td style="text-align: center;">
        <%
 		    If svMembLevel > 3 Then
          If vCurMods <> vPrvMods Or vCurProg <> vPrvProg Then
   		      If vEcom Then
        %>
        <input class="button040" title="Ecommerce Details for this Program" type="button" onclick="openWindow(<%=vCurN %>, 'Ecom', '<%=vProgId%>', <%=vMembNo%>)" value="Ecom" style="padding: 1px 0 1px 0;">
        <%
            End If 
          End If 
        End If 
        %>
      </td>

      <td style="text-align: center;">
        <%
 		    If svMembLevel > 3 Then
          If vCurMods <> vPrvMods Or vCurProg <> vPrvProg Then
        %>
        <input class="button040" title="Full Editor for the Learner" type="button" onclick="openWindow(<%=vCurN %>, 'Hist', <%=vMembNo%>)" value="Hist">
        <%
          End If 
        End If 
        %>
      </td>

      <td style="text-align: center;">
        <%
 		    If svMembLevel > 3 Then
          If vCurMods <> vPrvMods Or vCurProg <> vPrvProg Then
        %>
        <input class="button040" title="Mini Editor for this Assessment Attempt" type="button" onclick="openWindow(<%=vCurN %>, 'Edit', '<%=vSessionId%>', <%=vMembNo%>, <%=vModsNo%>, <%=vProgNo%>)" value="Edit">
        <%
          End If 
        End If 
        %>
      </td>

      <td style="text-align: center;">
        <%
        If svMembLevel > 4 And IsNumeric(vCurSess) And (vCurMods <> vPrvMods Or vCurProg <> vPrvProg) Then
        %>
        <input class="button040" title="Sync this RTE session data with LMS session data" type="button" onclick="openWindow(<%=vCurN %>, 'Sync', <%=vSessionId%>, '<%=vMembId%>', '<%=vProgId%>', '<%=vModsId%>')" value="Sync">
        <%
        End If 
        %>
      </td>


      <td style="text-align: center;">
        <% 
        If svMembLevel > 2 And vMultiSCO Then
        %>
        <input class="button040" title="Show details for this MultiSCO session" type="button" onclick="openWindow(<%=vCurN %>, 'MultiSCO', <%=vSessionId%>)" value="More">
        <%
        End If 
        %>
      </td>

    </tr>
    <%
      vScore = ""
      vTimeSpent = ""
      vCurN = vCurN + 1

      vPrvMemb = vMembId
      vPrvProg = vProgId
      vPrvMods = vModsId
      vPrvSess = vSessionId
        
      If Cint(vCurN) Mod 100 = 0 Then Exit Do
      oRs.MoveNext
    Loop 

      
    %>
    <tr>
      <td colspan="20" style="text-align: center">
        <br />
        <div class="actionAlert bgYellow" style="padding: 10px; display: none;">
          <input type="button" onclick="jRestart()" class="button200" style="color: red; font-weight: 600;" value="Restart to view changes!" name="bRestart1">
        </div>
      </td>
    </tr>
    <tr>
      <td colspan="20" style="text-align: center">
        <form method="POST" action="RTE_History_O.asp">

          <!-- These form elements are for posting to this page to review next/back -->
          <input type="hidden" name="vFrom" value="<%=vFrom%>"><input type="hidden" name="vCurN" value="<%=vCurN%>"><input type="hidden" name="vStrD" value="<%=vStrD%>"><input type="hidden" name="vEndD" value="<%=vEndD%>"><input type="hidden" name="vProg" value="<%=vProg%>"><input type="hidden" name="vMods" value="<%=vMods%>"><input type="hidden" name="vAssg" value="<%=vAssg%>"><input type="hidden" name="vPass" value="<%=vPass%>"><input type="hidden" name="vLNam" value="<%=vLNam%>"><input type="hidden" name="vGrou" value="<%=vGrou%>"><input type="hidden" name="vOutp" value="<%=vOutp%>"><input type="hidden" name="vSave" value="<%=vSave%>">

          <input type="button" onclick="jRestart()" class="button100" value="<%=fIf(vFrom="menu", bRestart, bReturn) %>" name="bRestart">
          <%
          '...if NOT EOF then render the NEXT start with next value
          If Cint(vCurN) > 0 And Cint(vCurN) Mod 100 = 0 Then  
          %>
          <input type="submit" class="button100" value="<%=bNext %>" id="bNext" name="bNext">
          <%
  	      End If 
          '...for BACK reduce vCurN by 200 when processing form
          If Cint(vCurN) > 100 Then  
          %>
          <input type="submit" class="button100" value="<%=bBack %>" id="bBack" name="bBack">
          <%
    	      End If 
          %>
        </form>
      </td>
    </tr>
  </table>

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
          <iframe class="div" id="ifrInner" name="ifrInner" style="width: 100%; height: 100%; border: 0"></iframe>

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

  <style>
    #tabOuter tr td { padding: 0; border: 0; text-align: center; }
  </style>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
