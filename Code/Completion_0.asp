<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  '...you can only get from this page via Completion.asp which initializes all variables (unless we are responding to the form post)
  Dim vPrevL1, vAll_L1s, vAll_L0s, vSelected, vRoleChecked, vProgChecked, vModsChecked, aProgs, aProg, vWhere, aLocs, aLoc, vMyLocation, vMyRole, aRoles, vMyLocNoRole
  Dim vProgPrev, vProgId, vProgTitle, vModsId, vModsTitle, bModsAct, vModsCnt, vStrDate, vEndDate, vModsNo, vMemoLen

  sGetMemb svMembNo
  If vMemb_Level = 5 Or vMemb_Level = 4 Then svMembCriteria = "0"
  vMemoLen = Session("Completion_L0len") + Session("Completion_L1len")   '... if there is a value in the memo field (extended location) it must be greater than size since there's a pipe between them, ie ABC|DEFG

  '...display all available roles that can be reported upon
  If Request.Form.Count = 0 Then
    
    '...determine rights of user (RRRR TTTT RR)
    vMyLocation   = fCriteria (svMembCriteria)
    vMyRole       = Right(vMyLocation, Session("Completion_RLlen"))

    '...if facilitator get my location without role (cannot be crit = 0)
    If svMembLevel = 3 Then  
      If svMembCriteria = "0" Or Not IsNumeric(svMembCriteria) Then
        Response.Redirect "Error.asp?vErr=You must be assigned to one group (location) to use this service.&vReturn=n"
      Else
        vMyLocNoRole  = Left(vMyLocation, Len(vMyLocation) - Session("Completion_RLlen")) '...used for Fac reports, ie "SK SASKA " (note trailing space)
      End If
    End If

    '...grab available list of roles for admins and superwankers or those with extended rights
'   If Session("Completion_Level") > 4 Or Len(vMemb_Memo) > 8 Then
    If Session("Completion_Level") > 3 Or Len(vMemb_Memo) > vMemoLen Then
      Session("Completion_Roles")      = fDefault(Session("Completion_Roles"), fRole_All())

    '...grab available list of roles and crits for facs (their role plus children)
    ElseIf Session("Completion_Level") = 3 Then
      Session("Completion_Roles")      = fDefault(Session("Completion_Roles"), fRole_Children(vMyRole) & "," & vMyRole)
      
      Session("Completion_LXval") = ""
      aRoles = Split(Session("Completion_Roles"), ",")
      For i = 0 To Ubound(aRoles)
        Session("Completion_LXval")  = Session("Completion_LXval") & "'" & vMyLocNoRole & aRoles(i) & "',"
      Next
      Session("Completion_LXval") = Left(Session("Completion_LXval"), Len(Session("Completion_LXval")) -1)<!--  -->

    '...grab list of roles for others
    Else
      Session("Completion_Roles")      = fOkValue(fDefault(Session("Completion_Roles"), fRole_Children(vMyRole)))
      If Len(Session("Completion_Roles")) = 0 Then 
        Response.Redirect "Error.asp?vErr=You have not been assigned any Roles to manage.&vReturn=n"
      End If
    End If

    Session("Completion_Active")       = fDefault(Session("Completion_Active"), "Y")
    Session("Completion_Completed")    = fDefault(Session("Completion_Completed"), "X")

    Session("Completion_StrDate")      = fFormatSqlDate(fDefault(Session("Completion_StrDate"), "Jan 01, 2000"))
    Session("Completion_EndDate")      = fFormatSqlDate(fDefault(Session("Completion_EndDate"), Now()))

  Else  


    '...capture all selected roles 
    Session("Completion_RoleP") = Ucase(Replace(Request("vRoles"), ", ", ",")) 
    Session("Completion_RoleD") = ""
    aRoles = Split(Session("Completion_RoleP"), ",")
    For i = 0 To Ubound(aRoles)
      Session("Completion_RoleD") = Session("Completion_RoleD") & fPhraId(fRole_Title(aRoles(i))) & "<br>" 
    Next

    If (Len(Session("Completion_RoleD")) > 4) Then 
      Session("Completion_RoleD") = Left(Session("Completion_RoleD"), Len(Session("Completion_RoleD"))-4)
    End If

    '...capture all selected dates
    Session("Completion_StrDate")       = fFormatSqlDate(fDefault(Request("vStrDate"), Session("Completion_StrDate")))
    Session("Completion_EndDate")       = fFormatSqlDate(fDefault(Request("vEndDate"), Session("Completion_EndDate")))

    '...this clears out any previous report parameters that this learner has created
    sResetReport

    '...capture all selected programs|modules
    Session("Completion_Programs") = ""
    Session("Completion_ProgramD") = ""
    
    '...capture selections for rendering on/off checkbox
    Session("Completion_selProgIds") = ""
    Session("Completion_selModsIds") = ""

    vProgPrev = ""
     '...store selected progs
    For Each vFld In Request.Form
      If Left(vFld, 6) = "vProgs" And Len(vFld) = 11 Then
        Session("Completion_selProgIds") = Session("Completion_selProgIds") & vFld & " "
      End If
    Next  

    For i = 1 to Request("vModsCnt")
      For Each vFld In Request.Form
        If Left(vFld, 10) = "vProg_" & Right("0000" & i, 4) Then
          aProg  = Split(Request(vFld), "|")
          '...store selected mods
          Session("Completion_selModsIds") = Session("Completion_selModsIds") & vFld & " "
          '...flag selected prog/mods 
          sCreateReport aProg(0), aProg(2)
          If vProgPrev <> aProg(0) Then
            Session("Completion_Programs") = Session("Completion_Programs") & "," & aProg(0)
          End If
            Session("Completion_Programs") = Session("Completion_Programs") & "|" & aProg(2)
          If vProgPrev <> aProg(0) Then
            Session("Completion_ProgramD") = Session("Completion_ProgramD") & aProg(0) & " : " & aProg(1) & "<br>"
          End If
            Session("Completion_ProgramD") = Session("Completion_ProgramD") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & aProg(2) & " : " & aProg(3) & "<br>"
            vProgPrev = aProg(0)
          Exit For
        End If
      Next
    Next


    Session("Completion_checkRoles") = fDefault(Request("checkRoles"), "off")
    Session("Completion_checkProgs") = fDefault(Request("checkProgs"), "off")

    Session("Completion_Programs")   = Mid(Session("Completion_Programs"), 2) '...kill first comma
    Session("Completion_ProgramD")   = Left(Session("Completion_ProgramD"), Len(Session("Completion_ProgramD"))-4)

    Session("Completion_Locations")  = Request("vLocations")
    Session("Completion_LocationD")  = Request("vLocations")
    Session("Completion_L1val")      = Left(Session("Completion_Locations"), Session("Completion_L1len"))
    Session("Completion_L0val")      = Mid(Session("Completion_Locations"),  Session("Completion_L1len") + 2, Session("Completion_L0len"))

    '   complete report tables
    sEndReport

    Session("Completion_Active") = Ucase(Request("vActive"))
    Select Case Session("Completion_Active")
    	Case "Y"  : Session("Completion_ActiveD") = fPhraH(000024)
      Case "N"  : Session("Completion_ActiveD") = fPhraH(000189)
      Case Else : Session("Completion_ActiveD") = fPhraH(000602)
    End Select

    Session("Completion_Completed") = fDefault(Ucase(Request("vCompleted")), "X")
    Select Case Session("Completion_Completed")
    	Case "Y"  : Session("Completion_CompletedD") = fPhraH(000024)
      Case "N"  : Session("Completion_CompletedD") = fPhraH(000189)
      Case Else : Session("Completion_CompletedD") = fPhraH(000602)
    End Select

    For Each i In Session.Contents
      If Instr(i, "Completion_") > 0 Then
        Response.Write "<br>" & i & " : " & Session(i)
      End If
    Next

    '...Check if all L1s ie "0000|0000" (for example)
    If Instr(Session("Completion_Locations"), Session("Completion_L1all") & "|" & Session("Completion_L0all")) > 0 Then 
  	  Response.Redirect "Patience.asp?vNext=Completion_1.asp"
    '...Check if all units are in L0 ie "1234|0000" (for example)
    ElseIf Instr(Session("Completion_Locations"), "|" & Session("Completion_L0titall")) > 0 Then 
      Response.Redirect "Completion_4.asp"
    '...Assume and individual region/unit ie "1234|6565" (for example)
    Else
  	  Response.Redirect "Patience.asp?vNext=Completion_5.asp"
    End If

  End If 
%>

<html>

<head>
  <title>Completion_0</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>	

    // turn on/off any group element starting with "group" value, ie "vProg" or "vProgP1234"
    function checkOnOff(theElement, group) {
      var i, j, theForm = theElement.form;
      j = group.length;
      for (i = 0; i < theForm.length; i++) {
        if (theForm[i].type == "checkbox" && theForm[i].id.substring(0, j) == group) {
          theForm[i].checked = theElement.checked;
	      }
	    }
    }

 
   function Validate(theForm) {
      var message_01 = "<%=fPhraH(000645)%>"
      var message_02 = "<%=fPhraH(000646)%>"
      var isChecked = false
     
      if (theForm.vRoles.length != undefined) {
        for (i=0; i < theForm.vRoles.length; i++) {
          if (theForm.vRoles[i].checked == true) {
            isChecked = true;
          }       
        }
        if (isChecked == false) {
          alert(message_01);
          theForm.vRoles(0).focus();
          return (false);
        }
      }


      var isChecked = false;
      for (i=0; i < theForm.length; i++) {
        if (theForm[i].type == "checkbox" && theForm[i].id.length == 15) {
          if (theForm[i].checked == true) {
            isChecked = true;
          }       
        }  
      }
      if (isChecked == false) {
        alert(message_02);
        return (false);
      }

		 theForm.bContinue.disabled = true;
		 document.getElementById("spanMessage").innerHTML = "This can take several minutes.  Please be patient.";
     return (true);
    }
  </script>
  <style>
    th, td { padding: 10px; }
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

    <div>
      <h1><!--webbot bot='PurpleText' PREVIEW='Completion Report'--><%=fPhra(000863)%></h1>
      <h2><!--webbot bot='PurpleText' PREVIEW='Selecting multiple Roles and multiple Programs will require several minutes of processing.'--><%=fPhra(001372)%></h2>
      <h2><!--webbot bot='PurpleText' PREVIEW='Learning Completion Rates'--><%=fPhra(000603)%> - <!--webbot bot='PurpleText' PREVIEW='Selection Criteria'--><%=fPhra(000610)%><br><!--webbot bot='PurpleText' PREVIEW='For all Active Learners'--><%=fPhra(000873)%></h2>
    </div>

    <form method="POST" onsubmit="return Validate(this)" name="Completion_0" action="Completion_0.asp">
      <table class="table">
        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Roles'--><%=fPhra(000615)%> :</th>
          <td>          
            <%
              '...display available roles for normal folk
              aRoles = Split(Session("Completion_Roles"), ",")
              For i = 0 To Ubound(aRoles)
            %>
              <input type="checkbox" name="vRoles" id="vRoles" value="<%=aRoles(i)%>" <%=fchecks(Session("Completion_RoleP"), aRoles(i))%>><%=fPhraId(fRole_Title(aRoles(i)))%><br> 
            <%
              Next
            %>
            <br>
            <input type="checkbox" name="checkRoles" onclick="checkOnOff(this, 'vRole');" value="on" <%=fCheck("on", Session("Completion_checkRoles"))%>><!--webbot bot='PurpleText' PREVIEW='Select All/None'--><%=fPhra(000761)%>           
          </td>
        </tr>

        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Programs | Modules'--><%=fPhra(001238)%> :</th>
          <td>
          <%
            '...Display all programs for admins or those with extended content stored in memo field
            If Session("Completion_Level") > 3 Or Len(vMemb_Memo) > vMemoLen Then
              vSql = " SELECT DISTINCT" _   
                   & "   RepC.RepC_ProgId AS [ProgId],    " _
                   & "   Prog.Prog_Title1 AS [ProgTitle], " _
                   & "   RepC.RepC_ModsId AS [ModsId],    " _
                   & "   Mods.Mods_Title  AS [ModsTitle]  " _
                   & " FROM "_
                   & "   V5_Comp.dbo.RepC AS RepC WITH (NOLOCK) INNER JOIN "_
                   & "   V5_Base.dbo.Prog AS Prog WITH (NOLOCK) ON RepC.RepC_ProgId + '" & svLang & "' = Prog.Prog_Id INNER JOIN "_
                   & "   V5_Base.dbo.Mods AS Mods WITH (NOLOCK) ON RepC.RepC_ModsId + '" & svLang & "' = Mods.Mods_Id "_
                   & " WHERE "_
                   & "   RepC.RepC_UserNo = " & svMembNo _ 
                   & " ORDER BY "_
                   & "   Prog.Prog_Title1, Mods_Title "

            ElseIf Session("Completion_Level") = 3 Then

              vSql = "SELECT DISTINCT" _ 
                   & "  Rc.RepC_ProgId AS [ProgId],     " _ 
                   & "  Pr.Prog_Title1 AS [ProgTitle],  " _ 
                   & "  Rc.RepC_ModsId AS [ModsId],     " _ 
                   & "  Mo.Mods_Title  AS [ModsTitle]   " _
                   & "FROM "_         
                   & "  V5_Vubz.dbo.Crit AS Cr      WITH (NOLOCK)                                                                            INNER JOIN "_
                   & "  V5_Vubz.dbo.Crit_Jobs AS CJ WITH (NOLOCK) ON Cr.Crit_No = CJ.Crit_Jobs_CritNo                                                 INNER JOIN "_
                   & "  V5_Vubz.dbo.Jobs_Prog AS JP WITH (NOLOCK) ON CJ.Crit_Jobs_JobsNo = JP.Jobs_Prog_JobsNo                                        INNER JOIN "_
                   & "  V5_Comp.dbo.RepC AS Rc      WITH (NOLOCK) ON Cr.Crit_AcctId = Rc.RepC_AcctId AND LEFT(JP.Jobs_Prog_ProgId, 5) = Rc.RepC_ProgId     INNER JOIN "_
                   & "  V5_Base.dbo.Prog AS Pr      WITH (NOLOCK) ON Rc.RepC_ProgId + '" & svLang & "' = Pr.Prog_Id                                        INNER JOIN "_
                   & "  V5_Base.dbo.Mods AS Mo      WITH (NOLOCK) ON Rc.RepC_ModsId + '" & svLang & "' = Mo.Mods_Id "_ 
                   & "WHERE "_     
                   & "  (Cr.Crit_AcctId = '" & svCustAcctId & "') AND "_ 
                   & "  (Cr.Crit_Id IN (" & Session("Completion_LXval") & ")) "_
                   & "ORDER BY "_
                   & "  Rc.RepC_ProgId, Rc.RepC_ModsId "


            '...Display only programs available to this user
            Else    
              vSql = "SELECT DISTINCT" _   
                   & "  vMod2.ProgId AS [ProgId],         " _
                   & "  Prog.Prog_Title1 AS [ProgTitle],  " _
                   & "  vMod2.ModsId AS [ModsId],         " _
                   & "  Mods.Mods_Title  AS [ModsTitle]   " _
                   & "FROM "_
                   & "  V5_Comp.dbo.vMod2 AS vMod2 WITH (NOLOCK) INNER JOIN "_
                   & "  V5_Base.dbo.Prog  AS Prog  WITH (NOLOCK) ON vMod2.ProgId + '" & svLang & "' = Prog.Prog_Id INNER JOIN "_
                   & "  V5_Base.dbo.Mods  AS Mods  WITH (NOLOCK) ON vMod2.ModsId + '" & svLang & "' = Mods.Mods_Id "_
                   & "WHERE "_
                   & "  (vMod2.MembNo = " & svMembNo & ") AND "_
                   & "  (CHARINDEX(RIGHT(vMod2.CritId, " & Session("Completion_RLlen") & "), '" & Session("Completion_RoleP") & "') > 0) "_                   
                   & "ORDER BY "_
                   & "  Prog.Prog_Title1 "
            End If        

            sCompletion_Debug

            '...Grab modules for each program (need to incorporate into reports)
            sOpenDb
            Set oRs = oDb.Execute(vSql)
            vModsCnt = 0
            bModsAct = False
            vProgPrev = ""

            Do While Not oRs.Eof
              vProgId     = oRs("ProgId")
              vProgTitle  = oRs("ProgTitle")
              vModsId     = oRs("ModsId") 
              vModsTitle  = oRs("ModsTitle")              
					  
              '...Display the Progam Id and Title
              If vProgId <> vProgPrev Then

                If bModsAct Then '...if we did have modules then close the DIV
                  Response.Write "</div>"
                  bModsAct = False
                End If 

                vProgChecked  = fIf(Session("Completion_checkProgs") = "on" Or Instr(Session("Completion_selProgIds"), vProgId) > 0, " checked", "")              
%>              
                <input type="checkbox" name="vProgs<%=vProgId%>" id="vProgs<%=vProgId%>" onclick="checkOnOff(this, 'vProg<%=vProgId%>')" value="<%=vProgId%>" <%=vProgChecked%>>
                <a href="#" onclick="toggle('div_<%=vProgId%>')"><%= fLeft(vProgTitle, 80)%></a><br> 
<%                
                vProgPrev = vProgId
              End If

              '...Display the Module Id and Title
              If Not bModsAct Then '...if this is the first module then put within a DIV
                Response.Write "<div id='div_" & vProgId & "' class='div' style='margin-left: 20'>"
                bModsAct = True
              End If              
              
              vModsCnt      = vModsCnt + 1
              vModsNo       = "vProg_" & Right("0000" & vModsCnt, 4)
              vModsChecked  = fIf(Instr(Session("Completion_selModsIds"), vModsNo) > 0, " checked", "")              
%>            
               <input type="checkbox" name="<%=vModsNo%>" id="vProg<%=vProgId & "|" & Left(vModsId, 4)%>" value="<%=vProgId & "|" & vProgTitle & "|" & Left(vModsId, 4) & "|" & vModsTitle%>" <%=vModsChecked%>><%=vModsTitle%><br>
<%            
              oRs.MoveNext
            Loop

            If bModsAct Then Response.Write "</div>"  '...if we did have modules then close the DIV

            Set oRs = Nothing
            sCloseDb
          %>

          <br>
          <input type="checkbox" name="checkProgs" onclick="checkOnOff(this, 'vProg');" value="on" <%=fCheck("on", Session("Completion_checkProgs"))%>><!--webbot bot='PurpleText' PREVIEW='Select All/None'--><%=fPhra(000761)%><br><br>
          <!--webbot bot='PurpleText' PREVIEW='Click Program Title to refine Selection.'--><%=fPhra(001373)%><br>
          <!--webbot bot='PurpleText' PREVIEW='Report will show all selected Modules'--><%=fPhra(001374)%></td>

        </tr>
        <input type="hidden" name="vActive" value="y">

        <%
            '...temp do not display the completion options  
            i = 0 : if i = 1 Then
        %>

        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Learning Completed'--><%=fPhra(000656)%>&nbsp;:</th>
          <td>
          	<input type="radio" value="Y" name="vCompleted" <%=fchecks(session("Completion_completed"), "y")%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%> 
          	<input type="radio" value="N" name="vCompleted" <%=fchecks(session("Completion_completed"), "n")%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%> 
          	<input type="radio" value="X" name="vCompleted" <%=fchecks(session("Completion_completed"), "x")%>><!--webbot bot='PurpleText' PREVIEW='All'--><%=fPhra(000602)%> 
         </td>
        </tr>
        <%
            end if
        %> 


        <% 

          p1 = fFormatSqlDate (Now)
          If vMyLocation = "All" Or vMyLocation = "Tous" Then
            Session("Completion_L1val") = ""
            Session("Completion_L0val") = ""
          
            '...temp do not display the date options  
            i = 0 : if i = 1 Then
        %>
        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Completed on or after'--><%=fPhra(001283)%> :</th>
          <td>
            <input type="text" name="vStrDate" size="13" value="<%=Session("Completion_StrDate")%>"> <!--webbot bot='PurpleText' PREVIEW='Use English Date format'--><%=fPhra(001375)%>.&nbsp; <!--webbot bot='PurpleText' PREVIEW='Defaults to Jan 01, 2000.'--><%=fPhra(001376)%>
          </td>
        </tr>
        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Completed on or before'--><%=fPhra(001285)%> :</th>
          <td>
            <input type="text" name="vEndDate" size="13" value="<%=Session("Completion_EndDate")%>"> <!--webbot bot='PurpleText' PREVIEW='Use English Date format'--><%=fPhra(001375)%>.&nbsp; <!--webbot bot='PurpleText' PREVIEW='Defaults to today, ^1.'--><%=fPhra(001377)%>
          </td>
        </tr>
        
        <%
            end if
        %>    
            
        <tr>
          <th><%=Session("Completion_L1tit") & " | " & Session("Completion_L0tit")%>&nbsp;:</th>
          <td>
            <select name="vLocations" size="1">
          <%
              vSelected = ""

              vAll_L1s = "All " & Session("Completion_L1tits")
              vAll_L0s = "All " & Session("Completion_L0tits")
              vPrevL1  = ""

              vSql = " SELECT "_     
                   & "   Unit_L1, "_  
                   & "   Unit_L1Title, "_  
                   & "   Unit_L0, "_  
                   & "   Unit_L0Title "_
                   & " FROM "_         
                   & "   V5_Comp.dbo.Unit WITH (NOLOCK) "_
                   & " WHERE "_     
                   & "   (Unit_AcctId = '" & svCustAcctId & "') AND "_
                   & "   (Unit_Active = 1) "_
                   & " ORDER BY "_ 
                   & "   Unit_L1, Unit_L0 "

              sCompletion_Debug
              sOpenDb
              Set oRs = oDb.Execute(vSql)

              If Session("Completion_Locations") = Session("Completion_L1all") & "|" & Session("Completion_L0all") Then vSelected = "selected" Else vSelected = ""
              Response.Write vbCrLf & "<option value='" & Session("Completion_L1all") & "|" & Session("Completion_L0all") & "'" & vSelected & " >" & vAll_L1s & "</option>"

              Do While Not oRs.Eof
                vUnit_L1       = oRs("Unit_L1")
                vUnit_L0       = oRs("Unit_L0")
                vUnit_L1Title  = oRs("Unit_L1Title")
                vUnit_L0Title  = oRs("Unit_L0Title")

                If vUnit_L1 <> vPrevL1 Then
                  If Session("Completion_Locations") = vUnit_L1 & "|" & Session("Completion_L0all") Then vSelected = " selected" Else vSelected = ""
                  Response.Write vbCrLf & "<option value='" & vUnit_L1 & "|" & Session("Completion_L0all") & "'" & vSelected & " >" & vUnit_L1 & " (" & vUnit_L1Title & ")   | " & vAll_L0s & "</option>"
                End If

                If Session("Completion_Locations") = vUnit_L1 & "|" & vUnit_L0 Then vSelected = " selected" Else vSelected = ""
                Response.Write vbCrLf & "<option value='" & vUnit_L1 & "|" & vUnit_L0 & "'" & vSelected & " >" & vUnit_L1 & " | " & vUnit_L0 & " (" & vUnit_L0Title & ")</option>"

                vPrevL1 = vUnit_L1

                oRs.MoveNext
              Loop
              Set oRs = Nothing
              sCloseDb
            %>
            </select> 
          </td>
        </tr>


        <%
          '...minimum value must be like: 1234|5432, multiples like: 1234|5432 1234|5433 
          ElseIf Len(vMemb_Memo) > vMemoLen Then
            aLocs = Split(vMemb_Memo)
        %>
        <tr>
          <th><%=Session("Completion_L1tit") & " | " & Session("Completion_L0tit")%>&nbsp;:</th>
          <td>
            <!-- normal location -->
            <input type="radio" name="vLocations" value="<%=vMyLocation%>">
            <% = Left(vMyLocation, Session("Completion_L1len")) & "&nbsp;|&nbsp;" & Mid(vMyLocation, Session("Completion_L0str"), Session("Completion_L0len")) %> 
            <% =" (" & fL1Title(Left(vMyLocation, Session("Completion_L1len"))) & "&nbsp;|&nbsp;" & fL0Title(Mid(vMyLocation, Session("Completion_L0str"), Session("Completion_L0len"))) & ")"%>

            <!-- extended location(s) -->
            <% If Len(vMemb_Memo) > vMemoLen Then %>
                 <hr style="border:1px solid #DDEEF9">
            <%   For i = 0 To Ubound(aLocs) 
                   aLoc = Replace(aLocs(i), "|", " ") & " " & vMyRole
            %>
                   <input type="radio" name="vLocations" value="<%=aLoc%>"><% = Left(aLocs(i), Session("Completion_L1len")) & "&nbsp;|&nbsp;" & Mid(aLocs(i), Session("Completion_L0str"), Session("Completion_L0len")) %>
            <%     =" (" & fL1Title(Left(aLocs(i), Session("Completion_L1len"))) & "&nbsp;|&nbsp;" & fL0Title(Mid(aLocs(i), Session("Completion_L0str"), Session("Completion_L0len"))) & ")"%><br>
            <%   Next   
               End If 
            %>
          </td>
          <%
            Else
              Session("Completion_L1val") = Left(vMyLocation, Session("Completion_L1len"))
              Session("Completion_L0val") = Mid(vMyLocation, Session("Completion_L1str"), Session("Completion_L1len"))
				  %> 
				  <input type="hidden" value="<%=vMyLocation%>" name="vLocations">
        </tr>


        <% If Session("Completion_Level") <> 3 Then %>
        <tr>
          <th><%=Session("Completion_L1tit") & " | " & Session("Completion_L0tit") %>&nbsp;:</th>
          <td>&nbsp;
            <% = Left(vMyLocation, Session("Completion_L1len")) & " | " & Mid(vMyLocation, Session("Completion_L1str"), Session("Completion_L1len")) %> 
          </td>
        </tr>
        <% End If %>

        <%
          End If
        %>

        <tr>
          <td style="text-align:center; margin:40px;" colspan="2"><br>
		       	<input type="submit" value="<%=bContinue%>" name="bContinue" id="bContinue" class="button">
		       	<br><br><span id="spanMessage" class="c5">Please only click button once!</span>
         	</td>
        </tr>


      </table>
      <input type="hidden" name="vModsCnt" value="<%=vModsCnt%>">

    </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>

