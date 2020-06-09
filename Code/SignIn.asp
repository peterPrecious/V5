<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<!-- include virtual = "V5/MailServer/MailServer.asp"-->

<%

  Dim vCustOk, vOk, vError, vLogo, vCurrVisit 
  
  '...if session expired...
  If Len(Session("Lang")) = 0 Then Session("Lang") = "EN"

  vLogo = ""
  Session("Secure") = False
  sGetQueryString

  '...get MailServer from above include
'  sMailServer

  '...Values from language page?
  If Request.QueryString("vPage") = "language" Then  
    vLang = Request.QueryString("vLang")
    sPutQueryString
  End If

  '...Values from signin form on home page?
  If Request.Form("fSignin") = "Y" Then  
    vCust = fUnquote(Ucase(Trim(Request.Form("vCust"))))
    If fNoValue(vCust) Then vCust = ""
    vId   = fUnquote(Ucase(Trim(Request.Form("vId"))))
    If fNoValue(vId) Then vId   = ""     
    '...update the parms
    sPutQueryString

  '...Values passed via QueryString
  ElseIf Request.Form("vQueryString").Count = 1 Then
    '...now get the values from the public signin form
    Session("QueryString") = Request.Form("vQueryString")
    sGetQueryString
    vCust = fUnquote(Ucase(Trim(Request.Form("vCust"))))
    If fNoValue(vCust) Then vCust = ""
    vId   = fUnquote(Ucase(Trim(Request.Form("vId"))))
    If fNoValue(vId) Then vId     = ""     
    vPwd   = fUnquote(Ucase(Trim(Request.Form("vPwd"))))
    If fNoValue(vPwd) Then vPwd   = ""     

    '...update the parms
    sPutQueryString
  Else
    sGetQueryString
  End If


  '...use VU House accounts if no fields entered and update querystring
  If vCust = "" Then
    If vLang = "EN" Then
      vCust = "VUBZ5678"
      sPutQueryString
    ElseIf vLang = "FR" Then
      vCust = "VUBZ2275"
      sPutQueryString
    ElseIf vLang = "ES" Then
      vCust = "VUBZ2294"
      sPutQueryString
    End If
  End If

  vCustOk               = False
  vOk                   = False
  vMemb_Id              = vId
  
  Session("CustBanner") = "Vubz.jpg"
  Session("HostDb")     = vHostDb
  Session("CustId")     = vCust
  Session("CustAcctId") = Right(vCust, 4) '...temp until we check cust id
  Session("MembId")     = vId

  '...ensure customer id is always 8 chars long and the last 4 are numeric
  If Len(vCust) <> 8 Then
    vError = fPhraH(000592)
    Response.Redirect "SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError)
  
'...modified May 17, 2016 to allow in alpha  
'  ElseIf Not IsNumeric(Right(vCust, 4)) Then
'    vError = fPhraH(000592)
'    Response.Redirect "SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError)
 
   End If

  If fCustOk Then vCustOk = True

  '...valid customer?
  If Not vCustOk Then

    '...any access issues?
    If Len(vError) > 0 Then
      Response.Redirect "SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError)
    End If    

    '...must include a password
    If vMemb_Id = "" Then 
      vError = fPhraH(000253)
      Response.Redirect "SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError)
    End If    

    '...no such account
    vError = fPhraH(000253)
    Response.Redirect "SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError)
  End If

  '...If valid customer but no password
  If vCustOk And vId = "" Then
    '...see if there's a action directive for ecom customers

    Select Case vAction
      Case "OFFERINGS"    : Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("Ecom2Start.asp?vMode=More&vContentOptions=") & vContentOptions & "&vTab=4"
      Case "ORDER"        : Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("Ecom2Start.asp?vMode=More&vContentOptions=") & vContentOptions & "&vTab=4"
      Case "ORDERSINGLE"  : Session("TabActive") = True '...we need this for new public site
                            Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("Ecom2Default.asp?vEcom_Media=Online") & "&vTab=4"
      Case "ORDERGROUP2"  : Session("TabActive") = True '...we need this for new public site
                            Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("Ecom2Default.asp?vEcom_Media=Group2") & "&vTab=4"
      Case "ORDERSOLO"    : Response.Redirect "Ecom2Start.asp?vClose=Y&vMode=More&vContentOptions=" & vContentOptions
'     Case "PROMO"        : Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("EcomCdPromo.asp?vDiscount=10") & "&vTab=4"
      Case "QORDERSINGLE" : Response.Redirect "Ecom2BypassSingle.asp"
      Case "QORDERGROUP"  : Response.Redirect "Ecom2BypassGroup.asp"
      Case "QORDERGROUP2" : Response.Redirect "Ecom2BypassGroup2.asp"
'     Case "QORDERPRODS"  : Response.Redirect "Ecom2BypassProds.asp"
      Case "SIGNIN"       : Response.Redirect "/ChAccess/Vubiz/Default.asp?vCust=" & vCust & "&vLang=" & svLang
    End Select  

    If Len(vGoto) > 0 Then
      '...note when passing parameters in the vGoto statement, encode as follows:
      '   note that Server.Url Encode gets wiped out when using get/putquerystring
      vGoto = Replace (vGoto, "~1", "&")
      vGoto = Replace (vGoto, "~2", "=")
      vGoto = Replace (vGoto, "~3", "?")
      Response.Redirect vGoto
    End If

    '...all other customers
    Response.Redirect "Default.asp"
  End If


  '...if Cust Id And Memb Id Ok then what action?
  If vCustOk And fMembOk Then 

    '...this allows manager to bypass ecom (used by boreal college)
    If vAction = "ECOMBYPASS" Then
      Session("EcomBypass") = fMembEcomBypass(Session("CustAcctId"), vMemb_Memo)  '...allows this manager to bypass Ecommerce
      Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("Ecom2Start.asp?vMode=More&vContentOptions=") & vContentOptions & "&vTab=2"
    End If

    '...if ENROLL, then return with success
    If vAction = "ENROLL" Then
      Response.Redirect "Error.asp?vClose=y&vErr=" & Server.UrlEncode(fPhraH(000023)) & "&vReturn=" & vSource

    '...channel, qModId, special launch code or regular entry
    ElseIf vAction = "ORDER" Then
      Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("Ecom2Start.asp?vMode=More&vContentOptions=") & vContentOptions & "&vTab=2"

    ElseIf vAction = "MYWORLD" Then
      Response.Redirect "Default.asp?vPage=" & Server.UrlEncode("MyWorld.asp?vTskH_Id=" & vTskH_Id) & "&vTab=1"

    ElseIf vMemb_Level < 3 And Len(vCust_ContentLaunch) > 0 Then
      Response.Redirect vCust_ContentLaunch

    ElseIf Len(vGoto) > 0 Then
      '...note when passing parameters in the vGoto statement, encode as follows:
      '   note that Server.Url Encode gets wiped out when using get/putquerystring
      vGoto = Replace (vGoto, "~1", "&")
      vGoto = Replace (vGoto, "~2", "=")
      vGoto = Replace (vGoto, "~3", "?")

      '...these were added Oct 24 2017 to handle impersonation - rare
      vGoto = Replace (vGoto, "~4", "&")
      vGoto = Replace (vGoto, "~5", "=")
      vGoto = Replace (vGoto, "~6", "?")

      Response.Redirect vGoto

    ElseIf Len(vQmodId) > 0 Then

      '...set session variable so the closing of the mod will log off the customer
      Session("ModAutoClose")=True
      '...note when passing parameters via vQmodId statement, 
      '   encode as follows: vQModId=0123EN~32 which converts to vQModId=0123EN&vPageNo=32
      '   encode as follows: vQModId=0123EN>NN which converts to vQModId=0123EN&vTest=N&vBookmark=N (note default is NN so use NY, YN, YY)
      '   encode as follows: vQModId=0123EN^   which converts to vQModId=0123EN&showtree=1 (only need a single ^)
      '   encode as follows: vQModId=0123EN*   which converts to vQModId=0123EN&vBuild=Y (only need a single ^)
      '   encode as follows: vQModId=0123EN_09-24 which converts to vQModId=0123EN&psi=09&pei=24
      '   note that Server.Url Encode gets wiped out when using get/putquerystring

      vQModId = Replace (vQModId, "~", "&vPageNo=")
      vQModId = Replace (vQModId, "^", "&showtree=1")
      vQModId = Replace (vQModId, "*", "&vBuild=Y")

      vQModId = Replace (vQModId, "_", "&psi=")
      vQModId = Replace (vQModId, "-", "&pei=")
  
      '...set default to no test and no bookmark (these are for older F modules)
      If Instr(vQModId, ">") = 0 Then 
        vQModId = vQModId & "&vTest=N&vBookmark=N"
      Else
        vQModId = Replace (vQModId, ">NN", "&vTest=N&vBookmark=N")
        vQModId = Replace (vQModId, ">NY", "&vTest=N&vBookmark=Y")
        vQModId = Replace (vQModId, ">YN", "&vTest=Y&vBookmark=N")
        vQModId = Replace (vQModId, ">YY", "&vTest=Y&vBookmark=Y")
      End If

      Response.Redirect "../LaunchObjects.asp?vModId=" & vQModId '...this was the original quick launch but now it is handled by the main launch routine

    Else
      Response.Redirect "Default.asp"

    End If
    
  Else

    '...Else go back to signin error screen
    Session("CustId")     = ""
    Session("CustAcctId") = ""
    Session("HostDb")     = ""
    If fNoValue(Session("CustBanner")) Then Session("CustBanner") = "Vubz.jpg"
    If fNoValue(vError) Then vError = fPhraH(000253)
    Response.Redirect "SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError)
  End If


  '_________________________________________________________________________________

  Function fMembOk
    Dim i, j, k, vExpiresDate

    '...assume member access is not valid
    fMembOk = False

    '...Setup Customer DataBase and get Cust Info
    svCustAcctId  = Session("CustAcctId")
    svHostDb      = Session("HostDb")

    '...Get Memb info
    vMemb_Eof = True
    vSql = "SELECT * FROM Memb WITH (nolock) WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id = '" & vId & "'" 
    sOpenDb
    Set oRs = oDb.Execute(vSql)

    '...member on file?
    If Not oRs.Eof Then 
      vMemb_Eof = False
      sReadMemb
    End If
    Set oRs = Nothing    
    sCloseDb

    '...if on file
    If Not vMemb_Eof Then

      '...restrictions apply to users (extended to facs/mgrs Jan 26)
      If vMemb_Level < 5 And vMemb_Active = 0 Then
        vError = fPhraH(001648)
        Exit Function
      End If     

      '...If facilitator + bypass checking 
      If vMemb_Level > 2 Then
        sUpdateMembInfo
        fMembOk = True
        Exit Function
      End If

      '...are there size restrictions on the ID?
      If vCust_IDsSize > 0 And Len(vId) <> vCust_IDsSize And vMemb_Sponsor = 0 Then
        vError = fPhraH(001214)
        Exit Function
      End If

      '...ensure member entered via landing page or with vAction = "ECOMBYPASS" unless responding to an Email Alert
      If vCust_Auto And Len(vSource) = 0 And vAction <> "MYWORLD" And vAction <> "ECOMBYPASS" Then
        vError = fPhraH(000266)
        Exit Function
      End If

      '...if need pwd but are coming in via SSO as a learner then don't look for password
      If vCust_Pwd And vMemb_Level < 3 AND vAction = "SSO" And Len(vPwd) = 0 Then
        sUpdateMembInfo
        fMembOk = True
        Exit Function
      End If
      
      '...if need pwd ensure correct action
      If vCust_Pwd And vAction <> "MYWORLD" Then
        '...user can't ENROLL with an existing Id
        If vAction = "ENROLL" Then
          vError = fPhraH(000038)
          Exit Function
        ElseIf vAction <> "SIGNIN" THen
          vError = fPhraH(000037)
          Exit Function
        End If
        '...ensure pwd is valid
        If vPwd <> vMemb_Pwd Then
          vError = fPhraH(000034)
          Exit Function
        End If
      End If      

      '...if auto enroll or vAction = "ECOMBYPASS" And there is not a password then update
      If (vCust_Auto Or vAction="ECOMBYPASS") And Not vCust_Pwd Then

        If Len(vFirstName) > 0 Then vMemb_FirstName = fUnQuote(vFirstName)
        If Len(vLastName)  > 0 Then vMemb_LastName  = fUnQuote(vLastName)
        If Len(vEmail)     > 0 Then vMemb_Email     = fUnQuote(vEmail)
        If Len(vMemo)      > 0 Then vMemb_Memo      = fUnQuote(vMemo)
        '...when passing in content for auto-enroll, content must start with either P or J and be separated by spaces
        If Len(vTraining)  > 0 Then 
          If Left(vTraining, 1) = "P" Then
            vMemb_Programs  = vTraining
          ElseIf Left(vTraining, 1) = "J" Then    
            vMemb_Jobs      = vTraining
          End If
        Else  
          vTraining = "X"
        End If
        '...if there is was a vCriteria passed in then get vMemb_Criteria value if we are being passed a valid Group 1 value (vCrit_Id), else use "0"
        If Len(vCriteria)  > 0 Then 
          vMemb_Criteria  = fSigninCriteria(vCust_AcctId, vCriteria)
        End If

        vSql = ""
        If Len(vFirstName) > 0      Then vSql = vSql & " Memb_FirstName  = '" & fUnquote(vMemb_FirstName) & "', " 
        If Len(vLastName)  > 0      Then vSql = vSql & " Memb_LastName   = '" & fUnquote(vMemb_LastName)  & "', " 
        If Len(vEmail)     > 0      Then vSql = vSql & " Memb_Email      = '" & fUnquote(vMemb_Email)     & "', " 
        If Len(vMemo)      > 0      Then vSql = vSql & " Memb_Memo       = '" & fUnquote(vMemb_Memo)      & "', " 
        If Len(vCriteria)  > 0      Then vSql = vSql & " Memb_Criteria   = '" & vMemb_Criteria            & "', " 
        If Left(vTraining, 1) = "J" Then vSql = vSql & " Memb_Jobs       = '" & vMemb_Jobs                & "', " 
        If Left(vTraining, 1) = "P" Then vSql = vSql & " Memb_Programs   = '" & vMemb_Programs            & "', " 

        '...string trailing commas 
        If Len(vSql) > 0 Then 
          vSql = "UPDATE Memb SET" & Left(vSql, Len(vSql)-2) & " WHERE Memb_No   =  " & vMemb_No
          sOpenDb 
      '   sDebug
          oDb.Execute(vSql)
          sCloseDb
          sUpdateMemb_Session
        End If
      End If

    '...if not on file
    Else

      '...ensure member entered via landing page unless responding to an Email Alert
      If vCust_Auto And Len(vSource) = 0 And vAction <> "MYWORLD" And vAction <> "ECOMBYPASS" Then
        vError = fPhraH(000523)
        Exit Function
      End If


      '...if need pwd ensure correct action (other than MYWORLD)
      If vCust_Pwd Then
        '...user can't SIGNIN without an existing Id
        If vAction = "SIGNIN" Then
          vError = fPhraH(000026)
          Exit Function
        ElseIf vAction <> "ENROLL" THen
          vError = fPhraH(000037)
          Exit Function
        End If
      End If

      '...if not auto signing...
      If Not vCust_Auto And vAction <> "ECOMBYPASS" Then 
        vOk = False
        Exit Function
      End If

      '...If Auto members or "ENROLL" Password members, add normal with any parms
      If Not vCust_Pwd Or (vCust_Pwd And vAction = "ENROLL") Then

        vMemb_FirstName = fUnQuote(vFirstName)
        vMemb_LastName  = fUnQuote(vLastName)
        vMemb_Email     = fUnQuote(vEmail)
        vMemb_AcctId    = vCust_AcctId
        vMemb_Id        = vId
        vMemb_Pwd       = fNoQuote(vPwd)
        vMemb_Memo      = Replace(fUnquote(vMemo), "~1", "&")
        vMemb_Level     = 2
        vMemb_Active    = True
        vMemb_Criteria  = fSigninCriteria(vCust_AcctId, fUnquote(vCriteria))  '...get vMemb_Criteria value if we are being passed a valid Group 1 value (vCrit_Id), else use "0"
        vMemb_Jobs      = fIf(Left(vTraining, 1)="J", vTraining, Null)
        vMemb_Programs  = fIf(Left(vTraining, 1)="P", vTraining, Null)

        sAddMemb vMemb_AcctId         '...add new member
      End If

    End If


    '...get vProg_MaxHours (at end of the vCust_Program string - 4th - remember there may be multiple programs) 
    '   can't remember why we do this???  were getting program values that were not proper strings
    vProg_MaxHours = 0
'... commented out Jun 1, 2016 as not sure why we need this
'    i = Split(Trim(vCust_Programs), " ")
'    For j = 0 to uBound(i)
'      k = Split(i(j), "~")
'      vProg_MaxHours = vProg_MaxHours + k(3)
'    Next     
     Session("CustMaxHours") = vProg_MaxHours




    '...see if there's an expiry date
    vExpiresDate = True '...assume there's a valid expiry date
    If Not IsDate(vMemb_Expires) Then 
      vExpiresDate = False
    ElseIf Year(vMemb_Expires) < 2000 Then
      vExpiresDate = False
    End If

    '...if no expirey date see if duration has expired
    If Not vExpiresDate Then
      If vMemb_Duration > 0 Then
        If Not IsDate(fFormatDate(vMemb_FirstVisit)) Then '...if no first date, make it today
          If Now > DateAdd("d", vMemb_Duration, Now) Then 
            vError = fPhraH(000254)
            Exit Function
          End If
        ElseIf Now > DateAdd("d", vMemb_Duration, vMemb_FirstVisit) Then 
          vError = fPhraH(000254)
          Exit Function
        End If
      End If
    End If

    '...if there is an expirey date on the member file, then ensure not expired
    If IsDate(vMemb_Expires) Then
      If Year(vMemb_Expires) > 2000 Then
        If Now > vMemb_Expires Then 
          vError = fPhraH(000254)
          Exit Function
        Else
          vExpiresDate = True '...valid vExpires date
        End If
      End If
    End If

    '...Ok if you get this far
    sUpdateMembInfo
    fMembOk = True

  End Function
  

  '_________________________________________________________________________________

  Sub sUpdateMembInfo
  

    '...update session info
    Session("MembNo")           = vMemb_No
    Session("MembPwd")          = vMemb_Pwd
    Session("MembFirstName")    = vMemb_FirstName
    Session("MembLastName")     = vMemb_LastName
    Session("MembEmail")        = vMemb_Email
    Session("MembLevel")        = vMemb_Level
    Session("MembNoVisits")     = vMemb_NoVisits + 1
    Session("MembNoHours")      = vMemb_NoHours / 60
    Session("MembFirstVisit")   = vMemb_FirstVisit
    Session("MembLastVisit")    = vMemb_LastVisit
    Session("MembExpires")      = vMemb_Expires
    Session("MembCriteria")     = vMemb_Criteria
    Session("MembManager")      = vMemb_Manager
    Session("MembInternal")     = vMemb_Internal
    Session("CurrVisit")        = Now
    Session("Secure")           = True
    Session("Breach")           = fIf(vMemb_Online, True, False)

    svMembNo = Session("MembNo")

    '... update Access information record using formatted date/time
    '... (re)added browser Jul 10, 2017 (needed to increase size of memb_browser field)

    vSql = "UPDATE Memb SET"_
         & "   Memb_Browser    = '" & vBrowser & "'," _
         & "   Memb_NoVisits   =  " & vMemb_NoVisits + 1 & "," _
         & "   Memb_FirstVisit = '" & fFormatSqlDateTime(fDefault(vMemb_FirstVisit, Now)) & "'," _ 
         & "   Memb_LastVisit  = '" & fFormatSqlDateTime(Now) & "'," _ 
         & "   Memb_Online     = 1 " _ 
         & "WHERE Memb_No      =  " & vMemb_No
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb

  End Sub


  
  '_________________________________________________________________________________

  Function fCustOk

    fCustOk = False
    sGetCust (vCust)
    If Not vFileOk Then Exit Function

    '...InActive
    If Not vCust_Active And vId <> vPassword5 Then 
      vError = fPhraH(000524)
      Exit Function      
    End If

    '...Expired?
    If IsDate(vCust_Expires) Then
      If Year(vCust_Expires) > 2000 Then
        If Now > vCust_Expires Then 
          If vId <> vPassword5 And vId <> "CCHSSALES" Then 
            sGetMembById Right(vCust, 4), vId
            If Not vMemb_Manager Then
              vError = fPhraH(000525)
              Exit Function
            End If
          End If
        End If
      End If
    End If

    If fNoValue(vCust_Cluster) Then
      vCust_Cluster = "C0001"
    End If
    
    Session("CustId")           = vCust
    Session("CustAcctId")       = Right("0000" & vCust_AcctId, 4) '...this ensure 3 digit account ids remain as 4, ie 0138

    Session("CustBanner")       = vCust_Banner
    Session("CustUrl")          = vCust_Url
    Session("CustEmail")        = vCust_Email
    Session("CustTitle")        = vCust_Title
    Session("CustCluster")      = vCust_Cluster
    Session("CustFreeHours")    = vCust_FreeHours
    Session("CustFreeDays")     = vCust_FreeDays
    Session("CustAuto")         = vCust_Auto
    Session("CustIssueIds")     = vCust_IssueIds
    Session("CustActivateIds")  = vCust_ActivateIds
    Session("CustIdsSize")      = vCust_IdsSize
    Session("CustExpires")      = vCust_Expires
    Session("CustCluster")      = vCust_Cluster
    Session("CustLevel")        = vCust_Level
    Session("CustPwd")          = vCust_Pwd

    '...determine return address, if not on customer table use vSource from LP, else leave empty
    '...Oct 24 2017 added the vSource value !important whereby "vSource" overrides "CustReturnUrl" 
    '...this was used in user_o.asp to impersonate
    If Instr(vSource, "!important") > 0 Then
      Session("CustReturnUrl")    = vSource
    ElseIf Len(vCust_ReturnUrl) > 0 Then
      Session("CustReturnUrl")    = vCust_ReturnUrl
    ElseIf Len(vSource) > 0 Then
      Session("CustReturnUrl")    = vSource
    End If

    '...increase session timeout (increased to 60 * 6 to jive with launchobject.asp - by PB on Jan 8, 2016)
'   Session.Timeout = 40
    Session.Timeout = 60 * 6

    fCustOk = True

  End Function
  
 
  Function fCheck
    Dim i, j, k, vTemp
    Const cAlpha = "ABCDEFGHXY"
    vTemp = vMemb_No * 4141
    vTemp = vMemb_No * 4141 + svCustId
    vTemp = Right("0000" & vTemp, 4)
    fCheck = "": 
    For i = 1 To 4
      j = mid(vTemp,i,1)   
      k = mid(cAlpha, j+1, 1)
      fCheck = fCheck & k
    Next
  End Function  

%>


