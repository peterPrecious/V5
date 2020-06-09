<%
  Dim vAction, vContent, aTskHs, aTskHList, vAccess, vIcon, vAlt, vTitle, vLink, vAlert, vAdmin, vUrl, vTarget, aProg, aMods, vMods, vRepository, vLocked, vPrereqId, vPrereq, vFolder, vAttempts, vUrlTitle1, vUrlTitle2
  Dim bLinkOk '...use this to bypass normal links and use advanced DIVs
  Dim bWs : bWs = True '...use to get program data from web service

  Dim vService1, vService2, vService3, vService4, vService5, vService6, vBg2, vBg4, vBg5, vBg6
  '...Tree array determine what node is 0:closed or 1:open
  Dim aTree, vTree, vTree_Prev, vLevel_Prev 
  '...Selection Criteria
  Dim vDateStart, vDateEnd, vTaskFilterOk, vPassword

  vTaskFilterOk = True
  vLocked       = False
  vPassword     = ""
  vAccess       = 0


  '...fMyWorld2 creates the html for MyWorld2 tasklist using "vContent"
  Function fMyWorld2 (vAcctId, vId, vActionX)


    sGetTskH_Rs vAcctId, vId    
    vContent = ""
    sLevel_99 '...dummy row to ensure column spacing is ok
    vAction = vActionX '...ensure all subs can access this variable
    vTree_Prev = 0        

    '...if not template and not the first pass then get values from session variable
    If vAction <> "template" And  Len(Session(svMembNo & "-" & vId & "-Tree")) > 0 Then

      If Len(Session(svMembNo & "-" & vId & "-Tree")) > 0 Then
        vTree = Session(svMembNo & "-" & vId & "-Tree")
      End If

      aTree = Split(vTree, " ")
      '...else initialize the array as having all nodes expanded/collapsed depending on file
    Else    
      Redim aTree(0)
      Do While Not oRs.Eof
        sReadTskH
        If Ubound(aTree) < vTskH_Order Then Redim Preserve aTree(vTskH_Order)  
        If vTskH_Collapse Then
          aTree(vTskH_Order) = 1
        Else
          aTree(vTskH_Order) = 0 
        End If
        oRs.MoveNext
      Loop
    End If

    '...if user is trying top open/close a task then toggle the status (999/998)
    If Len(Request("vToggle")) > 0 Then
      i = Request("vToggle")
      If i = 999 Then '...expand all
        For i = 0 to Ubound(aTree) : aTree(i) = 1 : Next 
      ElseIf i = 998 Then '...collapse all
        For i = 0 to Ubound(aTree) : aTree(i) = 0 : Next 
      Else
        If aTree(i) = 0 Then
          aTree(i) = 1
        Else    
          aTree(i) = 0
        End If
      End If
    End If

    '...read the task list
    If Not vTskH_Eof Then 
      oRs.MoveFirst

      Do While Not oRs.Eof
        sReadTskH

        '...see if level is filtered
        If vTskH_Level < 3 Then vTaskFilterOk = fTaskFilterOk
        
        If vTskH_Active And vTaskFilterOk Then

          '...get the current access state (0:closed, 1:open)
          vAccess = aTree(vTskH_Order)
         
            '...build the output line from current task if
          '   a level one (open or closed) 
          '   a level two if prev level 1 was open     
          '   a level two if template 
          If vTskH_Level < 2 Or (vTskH_Level = 2 And (vTree_Prev = 1 or vAction = "template")) Then
            Select Case vTskH_Level        
              Case 0
                sLevel_0
              Case 1
                sLevel_1
                vTree_Prev = vAccess
              Case 2 
                If Not vLocked Then
                  If fTskD_ActiveRs (vTskH_No) Then vTskH_Child = True
                  sLevel_2
                End If
            End Select
          End If

        End If
        oRs.MoveNext
        vLevel_Prev = vTskH_Level

      Loop
      sCloseDB

      '...display this line to ensure grid is ok      
      sLevel_99          

      '...Store tree in a session variable
      vTree = Join(aTree, " ")
      Session(svMembNo & "-" & vId & "-Tree") = vTree
      
      '...Store tree in a cookie for 1 day
      Response.Cookies(svMembNo & "-" & vId & "-Tree") = vTree
      Response.Cookies(svMembNo & "-" & vId & "-Tree").Expires = fFormatSqlDate(DateAdd("d", 1, Now))

    End If
    fMyWorld2 = vContent
  End Function


  Sub sLevel_0
    vContent = vContent & "<tr>" & vbCrLf
    vContent = vContent & "  <td colspan='3'><h1>" & vTskH_Title & "</h1><p class='c2'>" & vTskH_Desc & vPassword & "</p></td>" & vbCrLf
    vContent = vContent & "</tr>" & vbCrLf
  End Sub

  Sub sLevel_1
    vContent = vContent & "<tr>" & vbCrLf
    vContent = vContent & "  <td style='padding-top:10px;'>" & fAccess & " </td>" & vbCrLf
    vContent = vContent & "  <td colspan='2' class='c2' style=margin-top:0;'>" & vTskH_Title & "<br>" & vTskH_Desc & vPassword & "</td>" & vbCrLf
    vContent = vContent & "</tr>" & vbCrLf
  End Sub

  Sub sLevel_2
    Dim vOk
    vContent = vContent & "<tr>" & vbCrLf
    vContent = vContent & "  <td></td>" & vbCrLf
    vContent = vContent & "  <td style='padding-top:10px;'>" & fAccess & " </td>" & vbCrLf

    If Len(Trim(vTskH_Title)) > 0 Then
    vContent = vContent & "  <td class='c3'>" & vTskH_Title & "<br>" & vTskH_Desc & vPassword & "</td>" & vbCrLf
    Else
    vContent = vContent & "  <td>" & vTskH_Desc & vPassword & "</td>" & vbCrLf
    End If

    vContent = vContent & "</tr>" & vbCrLf

    '...get the tasks if vAccess = 1 And Not Eof And Active And Not locked
    vOk = False 
    If vAccess = 1 Then vOk = True

    If vOk And vLocked Then vOk = False
    If vOk And Not fIsLocked (vTskH_No, svMembNo) Then vOk = True

    If vAction = "template" Then vOk = True

    If vOk Then
      sGetTskD_Rs vTskH_No
      Do While Not oRs2.Eof 
        sReadTskD

        If vTskD_Active Then
          vTarget = ""
          vPreReq = True

          Select Case vTskD_Type

            Case "M", "MR", "MT", "AM", "10"  '...Module Live/Review/Test Prereq/Accessible/VuAssess

              If vTskD_Type = "10" Then
                vIcon  = "checkmark.gif"
              Else
                vIcon  = "bookclosed.gif"
              End If

              If vTskD_Type = "M" Or  vTskD_Type = "MT" Or vTskD_Type = "AM" Then
                vAlt   = Server.HtmlEncode("<!--{{-->Learning Module<!--}}-->")
              ElseIf vTskD_Type = "10" Then  
                vAlt   = Server.HtmlEncode("<!--{{-->Assessment<!--}}-->")
              Else  
                vAlt   = Server.HtmlEncode("<!--{{-->Module Review<!--}}-->")
              End If

              '...get player script - check if using simple or full mod id
              If vTskD_Type = "M" Or vTskD_Type = "MR" Or vTskD_Type = "AM"  Or vTskD_Type = "10" Then 

                If Len(vTskD_Id) = 6 Then '...just module so add dummy program
                  sGetProg ("P0000XX")
                  sGetMods (vTskD_Id)
                ElseIf Left(vTskD_Id, 1) <> "P" Then '...module plus test/bookmark
                  sGetMods Left(vTskD_Id, 6)
                Else
                  vProg_Id = Left(vTskD_Id, 7)
                  sGetProg vProg_Id
                  sGetMods (Mid(vTskD_Id, 9, 6))
                End If 

              ElseIf vTskD_Type = "MT" Then 
                '...get player script - check if using simple or full mod id
                If Len(vTskD_Id) = 13 Then '...just two modules
                  sGetMods Left(vTskD_Id, 6)
                  '...see if test for second module and if so change icon to lock plus link
                  vPrereqId = Right(vTskD_Id, 6)
                ElseIf Left(vTskD_Id, 1) <> "P" Then '...module plus test/bookmark
                  sGetMods Left(vTskD_Id, 6)
                  vPrereqId = Mid(vTskD_Id, Instr(vTskD_Id, "|")+1, 6)
                Else
                  vProg_Id = Left(vTskD_Id, 7)
                  sGetMods (Mid(vTskD_Id, 9, 6))
                  vPrereqId = Mid(vTskD_Id, InStrRev(vTskD_Id, "|")+1, 6)
                End If 

                '...now check if prereq ok
                If fBestTestGrade(svMembNo, vPreReqId) >= 80 Then
                  vPreReq = True
                Else
                  vPreReq = False
                  vIcon  = "lock.gif"
                End If
                
                '...strip off vPreReq for Url below
                vTskD_Id = Left(vTskD_Id, InStrRev(vTskD_Id, "|")-1)

              End If

              '...if review, add vReview=Y to modid
              If vTskD_Type = "MR" Then
                vTskD_Id = vTskD_Id & "&vReview=Y"
              End If

              '...vTskD_Id may contain full module address, ie: P1001EN|0002EN|N|Y
							'... H mods added Apr 18, 2018 
'             If (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z") And Not vMods_FullScreen Then
              If (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z" Or Ucase(vMods_Type) = "H") And Not vMods_FullScreen Then
                vUrl   = "/V5/LaunchObjects.asp?vModId=" & vProg_Id & "|" & vMods_Id & "&vNext=" & svPage
 '            ElseIf (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z") And vMods_FullScreen Then
              ElseIf (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z" Or Ucase(vMods_Type) = "H") And vMods_FullScreen Then
                vUrl   = "javascript:fullScreen('" & vTskD_Id & "')"
              Else
                vUrl   = "javascript:" & vMods_Script & "('" & vTskD_Id & "')"
              End If

              vTitle = vTskD_Title 
              If fNoValue(vTitle) Then vTitle = vMods_Title

              '...add link to the module description, etc unless no prereq
              If vPreReq Then
                vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch the Module<!--}}-->")
                vTitle = "<p><a " & fStatX & " href=""" & vUrl & """ title='" & vUrlTitle1 & "'>" & vTitle & "</a> "
              End If

              '...add exam status to vuAssess else add description and status
              If vTskD_Type = "10" Then
              
                vTitle = fVuAssessLink(vMods_Id, vTskD_Title, vTskD_Title)

              '...remaining assets
              Else
                vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to view Description<!--}}-->")
                vTitle = vTitle & " [<a " & fStatX & " href=""javascript:SiteWindow('ModuleDescription.asp?vClose=Y&vModId=" & vMods_Id & "')"" title='" & vUrlTitle1 & "'>" & "<!--{{-->Description<!--}}-->" & "</a>]"
                vTitle = vTitle & "&nbsp;<span class='green'>[" & fModStatusLink (svMembNo, vProg_Id, vMods_Id) & "]</span>"

              End If

              sLevel_3


            Case "S" '...Module Staging
              vIcon  = "bookclosed.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Learning Module<!--}}-->")
              '...get player script - check if using simple or full mod id
              If Len(vTskD_Id) = 6 Then '...just module
                sGetMods (vTskD_Id)
              ElseIf Left(vTskD_Id, 1) <> "P" Then '...module plus test/bookmark
                sGetMods Left(vTskD_Id, 6)
              Else
                vProg_Id = Left(vTskD_Id, 7)
                sGetMods (Mid(vTskD_Id, 9, 6))
              End If 
              '...vTskD_Id may contain ful module address, ie: P1001EN|0002EN|n|Y

              If Lcase(vMods_Script) = "zmodulewindow" Then
                vUrl   = "javascript:zmodulestaging('" & vTskD_Id & "')"
              ElseIf Lcase(vMods_Script) = "fmodulewindow" Then
                vUrl   = "javascript:fmodulestaging('" & vTskD_Id & "')"
              Else 
                vUrl = ""
              End If

              vTitle = vTskD_Title 
              If fNoValue(vTitle) Then vTitle = vMods_Title

              '...add link to the module description, etc
              vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch the Module<!--}}-->")
              vUrlTitle2 = Server.HtmlEncode("<!--{{-->Click here to view Description<!--}}-->")
              vTitle = "<a " & fStatX & " href=" & Chr(34) & vUrl & Chr(34) & " title='" & vUrlTitle1 & "'>" & vTitle & "</a> "
              vTitle = vTitle & "[<a " & fStatX & " href=""javascript:SiteWindow('ModuleDescription.asp?vClose=Y&vModId=" & vMods_Id & "')"" title='" & vUrlTitle2 & "'>" & "<!--{{-->Description<!--}}-->" & "</a>]"
              vTitle = vTitle & "&nbsp;<span class='green'>[" & fModStatusLink (svMembNo, vProg_Id, vMods_Id) & "]</span>"

              sLevel_3

            '...on May 2008 we only use the compressed options for next group
            '   even if they select uncompressed - and all use smart links
            Case "P","PC","U","UC","O","OC","R","RC","J","JC"  

              '...Program/Modules (either from Member|Progs/Member|Jobs/Criteria/Job)
              '   P: get Progs from my learning
              '   U: get Progs from user table
              '   O: get Jobs from user table
              '   R: get Jobs from criteria table
              '   J: get Jobs from my learning
             
              If Left(vTskD_Type, 1) = "P" Then 
                aProg = Split(Trim(vTskD_Id))

              '...get program from member record - put there by the skills training program or manually
              ElseIf Left(vTskD_Type, 1) = "U" Then
                aProg = Split(Trim(vMemb_Programs))
              
              '...get jobs programs from member record (note there are different formats, such as J1234EN|P4500EN... and J1234EN J1235EN...
              ElseIf Left(vTskD_Type, 1) = "O" Then 
                '...clean up Job codes, from: "J0000EN|P1138EN , J0003EN|P4701EN" to: J0000EN|P1138EN J0003EN|P4701EN
                vMemb_Jobs = Replace (vMemb_Jobs, ",", "")
                vMemb_Jobs = Trim(Replace (vMemb_Jobs, "  ", " "))
                aProg = Split(Trim(vMemb_Jobs), " ")              
                '...create table of programs
                k = "" 

                For i = 0 To Ubound(aProg)
                  If Left(aProg(i), 1) = "J" Then
                    sGetJobs Left(aProg(i), 7) '...just use job id in case format is: J1234EN|P1234EN (ignore the program part)
                    If vJobs_Active Then
                      k = k & Replace(vJobs_Mods, "XX", svLang) & " "
                    End If
                  End If
                Next
                aProg = Split(Trim(k))

              Else
                If Left(vTskD_Type, 1) = "R" Then
                  sGetJobsByMemb
                Else
                  sGetJobs vTskD_Id
                End If
                aProg = Split(Trim(Replace(Ucase(vJobs_Mods), "XX", svLang)), " ")  '...modified Jan 2008 to allow Jobs to contain Mods (actually programs) with XX which will be replaced with the language of the learner
              End If

              vIcon  = "offline.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Learning Program<!--}}-->")

              If IsArray(aProg) Then
                For k = 0 To Ubound(aProg) 
                  If bWs Then
                    sWsLink aProg(k)  
                  Else
                    sCreateLink aProg(k)  
                  End If
                Next
              End If
              

            '...grab active programs from the ecom table
            Case "PE"
              vIcon  = "offline.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Acquired Program<!--}}-->")
              aProg = Split(fEcomProgram2 (svCustId, svMembId))
              For k = 0 To Ubound(aProg) 
                sCreateLink aProg(k)  
              Next

            Case "L" '...Link
              vIcon = "form.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Program/Document<!--}}-->")
              '...if there's some parms add "&" else add "?"
              If Instr(vTskD_Id, "?") > 0 Then
                vTskD_Id = vTskD_Id & "&"
              Else
                vTskD_Id = vTskD_Id & "?"
              End If

              '...if Vubiz link, then send to new window so URLs do not conflict - ie each window contains it's own link within Vubiz
              If Instr(Lcase(vTskD_Id), "vubiz.com") Then 
                vUrl = vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
                vTitle = "<p><a target='_top' href='" & vUrl & "'>" & vTskD_Title & "</a> "
  
              '...if non Vubiz link
              Else
                If vTskD_Window = 1 Then
                  vUrl = "javascript:SiteWindow('" & vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No & "&MembNo=" & svMembNo & "')"                                 
                ElseIf vTskD_Window = 9 Then
                  vUrl = "javascript:fullScreen('" & vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No & "&MembNo=" & svMembNo & "')"                                 
                Else
                  vUrl = vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No & "&MembNo=" & svMembNo
                End If
                vTitle = "<p><a " & fStatX & " href=" & vUrl & ">" & vTskD_Title & "</a> "                
              End If

              sLevel_3
            
            Case "V", "D", "C"  '...Documents: Private, Client, Common
              vIcon  = "document.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Repository Document<!--}}-->")
              '...if there's some parms add "&" else add "?"
              If Instr(vTskD_Id, "?") > 0 Then
                vTskD_Id = vTskD_Id & "&"
              Else
                vTskD_Id = vTskD_Id & "?"
              End If
              '...build up the Url
              vUrl = "../Repository/"
              Select Case vTskD_Type
                  Case "C" : vUrl = vUrl & svHostDb & "/" & "0000"       & "/Tools/" & vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
                  Case "V" : vUrl = vurl & svHostDb & "/" & svCustAcctId & "/"       & vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
                  Case "D" : vurl = vUrl & svHostDb & "/" & svCustAcctId & "/Tools/" & vTskD_Id & "vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
              End Select
              If vTskD_Window = 1 Then 
                vUrl= "javascript:SiteWindow('" & vUrl &  "')"                                 
              ElseIf vTskD_Window = 9 Then 
                vUrl= "javascript:fullScreen('" & vUrl &  "')"                                
              End If 
'             vTitle = vTskD_Title
              vTitle = "<p><a " & fStatX & " href=" & vUrl & ">" & vTskD_Title & "</a> "                
              sLevel_3

            Case "T" '...Test
              vIcon = "checkmark.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Self Assessment<!--}}-->")
              vUrl = vTskD_Id
              vTitle = vTskD_Title
              sLevel_3


            Case "E", "ET" '...Exam and Exam with Prerequisite
              vIcon  = "checkmark.gif"

              If vTskD_Type = "ET" Then 
                '...get player script - check if using simple or full mod id
                sGetMods (Right(vTskD_Id, 6))
                vPrereqId = Mid(vTskD_Id, InStrRev(vTskD_Id, "|")+1, 6)
                '...now check if prereq ok
                If fBestTestGrade(svMembNo, vPreReqId) >= 80 Then
                  vPreReq = True
                Else
                  vPreReq = False
                  vIcon  = "lock.gif"
                End If                
                '...strip off vPreReq for Url below
                vTskD_Id = Left(vTskD_Id, InStrRev(vTskD_Id, "|")-1)
              End If


              vAlt   = Server.HtmlEncode("<!--{{-->Examination<!--}}-->")
              vUrl   = "javascript:examwindow('" & vTskD_Id & "')"
              vTitle = "<p>" & vTskD_Title & "</p>"

              sLevel_3


            Case "Y" '...Title
              vIcon  = "line.gif"
              vAlt   = Server.HtmlEncode("<!--{{-->Title<!--}}-->")
              vUrl   = ""
              vTitle = vTskD_Title
              sLevel_3

          End Select
        
        End If

        oRs2.MoveNext
      Loop
      
    End If
  End Sub
  
  Sub sLevel_3

    vContent = vContent & "<tr>" & vbCrLf
    vContent = vContent & "  <td style='text-align:center;'></td>" & vbCrLf

    '...make icon inactive if a template view
    If vIcon = "line.gif" Then '...title, bold
      vContent = vContent & "  <td style='text-align:center; padding-top:10px;'><img border='0' src='../Images/Icons/" & vIcon & "' alt='" & vAlt & "'></td>" & vbCrLf
      vContent = vContent & "  <td>" & vTitle & "</td>" & vbCrLf
    ElseIf vIcon = "lock.gif" Then '...title, bold
      vContent = vContent & "  <td style='text-align:center; padding-top:10px;'><img border='0' src='../Images/Icons/" & vIcon & "' alt='" & vAlt & "'></td>" & vbCrLf
      vContent = vContent & "  <td>" & vTitle & "</td>" & vbCrLf
    ElseIf vAction <> "template" And vUrl <> "" Then
      vContent = vContent & "  <td style='text-align:center; padding-top:10px;'><a " & fStatX & " href=" & vUrl & vTarget & "><img border='0' src='../Images/Icons/" & vIcon & "' alt='" & vAlt & "'></a></td>" & vbCrLf
      vContent = vContent & "  <td>" & vTitle & "</td>" & vbCrLf
    Else
      vContent = vContent & "  <td style='text-align:center; padding-top:10px;'><img border='0' src='../Images/Icons/" & vIcon & "' alt='" & vAlt & "'></td>" & vbCrLf
      vContent = vContent & "  <td>" & vTitle & "</td>" & vbCrLf
    End If
    vContent = vContent & "</tr>" & vbCrLf

  End Sub

  Sub sLevel_99
    vContent = vContent & "<tr>" & vbCrLf
    vContent = vContent & "  <td></td>" & vbCrLf
    vContent = vContent & "  <td> </td>" & vbCrLf
    vContent = vContent & "  <td> </td>" & vbCrLf
    vContent = vContent & "</tr>" & vbCrLf
  End Sub

  Function fAlert
    fAlert = "&nbsp;"
    If i = 2 Then
      fAlert = "<img border='0' src='../Images/Icons/bang.gif'>"
    End If
  End Function


  Function fAccess
    vLocked = False
    fAccess = "&nbsp;"
    If vAction <> "template" Then
      If vTskH_Locked And fIsLocked (vTskH_No, svMembNo) Then
        fAccess = "<img border='0' src='../Images/Icons/lock.gif' alt='" & Server.HtmlEncode("<!--{{-->Locked Task<!--}}-->") & "' width='18' height='22'>"
        vLocked = True
      ElseIf Not vTskH_Child Then
        fAccess = "&nbsp;"
      ElseIf vAccess = 0 Then
'...anchor bug in XP SP2 - removed
'       fAccess = "<a name='" &  vTskH_Order & "' href='MyWorld2Redirect.asp?vTskH_AcctId=" & vTskH_AcctId & "&vTskH_Id=" & vTskH_Id & "&vToggle=" & vTskH_Order & "#" &  vTskH_Order & "'><img border='0' src='../Images/Common/VuPlus.gif'  alt='" & Server.HtmlEncode("<!--{{-->Expand Node<!--}}-->") & "'></a>"
        fAccess = "<a name='" &  vTskH_Order & "' href='MyWorld2Redirect.asp?vTskH_AcctId=" & vTskH_AcctId & "&vTskH_Id=" & vTskH_Id & "&vToggle=" & vTskH_Order & "'><img border='0' src='../Images/Common/VuPlus.gif'  alt='" & Server.HtmlEncode("<!--{{-->Expand Node<!--}}-->") & "'></a>"
      Else
'...anchor bug in XP SP2 - removed
'       fAccess = "<a name='" &  vTskH_Order & "' href='MyWorld2Redirect.asp?vTskH_AcctId=" & vTskH_AcctId & "&vTskH_Id=" & vTskH_Id & "&vToggle=" & vTskH_Order & "#" &  vTskH_Order & "'><img border='0' src='../Images/Common/VuMinus.gif' alt='" & Server.HtmlEncode("<!--{{-->Collapse Node<!--}}-->") & "'></a>"
        fAccess = "<a name='" &  vTskH_Order & "' href='MyWorld2Redirect.asp?vTskH_AcctId=" & vTskH_AcctId & "&vTskH_Id=" & vTskH_Id & "&vToggle=" & vTskH_Order & "'><img border='0' src='../Images/Common/VuMinus.gif' alt='" & Server.HtmlEncode("<!--{{-->Collapse Node<!--}}-->") & "'></a>"
      End If
    End If
  End Function


  Sub sWsLink (vProgId)
    vUrl   = ""
    sGetProg vProgId
    vTitle = vbCrLf & "<span class='c4'><a href='javascript:getProgramData(""" & vProgId & """, """ & svMembId & """)'><b>" & vProg_Title & "</b></a></span>"
    vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch the Module<!--}}-->")
    vTitle = vTitle & vbCrLf & "    <div class='div' id='div_" & vProgId & "'></div>" & vbCrLf & "  "
    sLevel_3
  End Sub


  '...This creates a smart link for each program whereby modules are hidden/displayed
  Sub sCreateLink (vProgId)

    Dim aMods, vExamId, i, j, k
    
    sGetProg vProgId
    aMods = Split(vProg_Mods)
    vUrl   = ""
    vTitle = "<span class='c4'><a href='javascript:toggle(""div_" & vProgId & """)'><b>" & vProg_Title & "</b></a></span>" & vbCrLf
    vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch the Module<!--}}-->")
    '...embed the modules in a div that can be hidden/displayed
    vTitle = vTitle & "<div class='div' id='div_" & vProgId & "'>" & vbCrLf
    vTitle = vTitle & "  <table style='BORDER-COLLAPSE: collapse' bordercolor='#ddeef9' cellpadding='2' border='0'>" & vbCrLf
    
    For k = 0 to Ubound(aMods)

      sGetMods aMods(k)

      vUrl = vProgId & "|" & vMods_Id  & "|" & vProg_Test & "|" & vProg_Bookmark & "|" & vProg_CompletedButton

			'... H modules (plus Z) added Apr 18, 2018
'     If (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX") And Not vMods_FullScreen Then
      If (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z" Or Ucase(vMods_Type) = "H") And Not vMods_FullScreen Then
        vUrl   = "/V5/LaunchObjects.asp?vModId=" & vUrl & "&vNext=" & svPage
'     ElseIf (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX") And vMods_FullScreen Then
      ElseIf (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z" Or Ucase(vMods_Type) = "H") And vMods_FullScreen Then
        vUrl   = "javascript:fullScreen('" & vUrl & "')"
      Else
        vUrl   = "javascript:" & vMods_Script & "('" & vUrl & "')"
      End If

      vTitle = vTitle & "    <tr>" & vbCrLf
      vTitle = vTitle & "      <td><p><a " & fStatX & " href=""" & vUrl & """ title='" & vUrlTitle1 & "'>" & vMods_Title & "</a>"
      If fModsDesc (vMods_Id) Then
        vTitle = vTitle & " [<a " & fStatX & " href=""javascript:SiteWindow('ModuleDescription.asp?vClose=Y&vModId=" & vMods_Id & "')"" title='" & vUrlTitle1 & "'>" & "<!--{{-->Description<!--}}-->" & "</a>]"
      End If
      vTitle = vTitle & "&nbsp;<span class='green'>[" & fModStatusLink (svMembNo, vProg_Id, vMods_Id) & "]</span></td>" & vbCrLf
      vTitle = vTitle & "    </tr>" & vbCrLf
    Next

    '...assessment included?
    If Len(vProg_Assessment) > 0 Then  
      sGetMods (vProg_Assessment)
      vTitle = vTitle & "    <tr>" & vbCrLf
      vTitle = vTitle & "      <td>"
      vTitle = vTitle &          fVuAssessLink (vMods_Id, "<!--{{-->Examination<!--}}-->", vMods_Title)
      vTitle = vTitle & "      </td>"
      vTitle = vTitle & "    </tr>" & vbCrLf

    '...platform exam included?
    ElseIf Lcase(vProg_Exam) <> "n" Then  
      Session("CertProg") = vProg_Id
      vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch examination<!--}}-->")
      vTitle = vTitle & "    <tr>" & vbCrLf
      vTitle = vTitle & "      <td><a " & fStatX & " href=""javascript:examwindow('" & vProg_Exam & "')"" title='" & vUrlTitle1 & "'>Examination</a>"
      vExamId = Mid(vProg_Exam, 22, 6)
      If fExamOk(vExamId) Then
      vTitle = vTitle & "        &nbsp;<span class='green'>[" & fAssessmentStatus (svMembNo, vExamId) & "]</span>"
      End If

      vTitle = vTitle & "    </tr>" & vbCrLf
    End If

    vTitle = vTitle & "  </table>" & vbCrLf
    vTitle = vTitle & "</div>" & vbCrLf

    '...reset this as this will be used for the icon
    vUrl   = "'javascript:toggle(""div_" & vProgId & """)'"

    sLevel_3

  
  End Sub   

%>
