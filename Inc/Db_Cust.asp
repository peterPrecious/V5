<%
  '____ Cust  ________________________________________________________________________

  Dim vCust_No, vCust_Placeholder, vCust_Id, vCust_AcctId, vCust_ParentId, vCust_Title, vCust_Lang, vCust_Agent, vCust_MaxSponsor, vCust_IssueIds, vCust_ResetStatus, vCust_IssueIdsTemplate, vCust_IssueIdsMemo, vCust_ActivateIds, vCust_IdsSize, vCust_FreeHours, vCust_FreeDays
  Dim vCust_Auto, vCust_Groups, vCust_Programs, vCust_CdPrograms, vCust_ContentOnline, vCust_ContentGroup, vCust_ContentGroup2, vCust_ContentProds, vCust_ContentCDs, vCust_Active, vCust_Desc, vCust_Added, vCust_Expires
  Dim vCust_Email, vCust_Cluster, vCust_CertLogoVubiz, vCust_CertLogoCust, vCust_CertEmailAlert, vCust_AssessmentAttempts, vCust_AssessmentScore, vCust_AssessmentCert, vCust_Level, vCust_ContentLaunch, vCust_Survey, vCust_NoCert, vCust_CustomCert
  Dim vCust_Tab1, vCust_Tab2, vCust_Tab3, vCust_Tab4, vCust_Tab5, vCust_Tab6, vCust_Tab7, vCust_Tab4Type 
  Dim vCust_Tab1Name, vCust_Tab2Name, vCust_Tab3Name, vCust_Tab4Name, vCust_Tab5Name, vCust_Tab6Name, vCust_Tab7Name
  Dim vCust_Auth, vCust_MyWorldLaunch, vCust_MaxUsers, vCust_Pwd, vCust_CritTitles, vCust_SeedLogs, vCust_Modified
  Dim vCust_EcomCurrency, vCust_EcomGroupLicense, vCust_EcomGroupSeat, vCust_EcomGroup2Rates, vCust_EcomSplit, vCust_EcomDiscOptions, vCust_EcomDisc, vCust_EcomDiscSplitCust, vCust_EcomDiscSplitVubz, vCust_EcomDiscSplitOwnr, vCust_EcomDiscMinUS, vCust_EcomDiscMinCA, vCust_EcomDiscMinQty, vCust_EcomDiscLimit, vCust_EcomDiscOriginal, vCust_EcomDiscPrograms, vCust_EcomRepurPrograms, vCust_EcomRepurDisc, vCust_EcomRepurPeriod, vCust_EcomCorpRate, vCust_EcomCorpDuration, vCust_EcomCorpProgram, vCust_EcomSeller, vCust_EcomOwner, vCust_EcomG2alert, vCust_CorpAlert
  Dim vCust_Resources, vCust_ResourcesMaxSponsor, vCust_VuNews, vCust_Scheduler, vCust_EcomReports, vCust_InfoEditProfile
  Dim vCust_InsertLearners, vCust_UpdateLearners, vCust_DeleteLearners, vCust_ResetLearners
  Dim vCust_Note1, vCust_Note2, vCust_Note3, vCust_Note4, vCust_Note5
  Dim vCust_ChannelParent, vCust_ChannelV8, vCust_ChannelNop
  Dim vCust_ChannelReportsTo, vCust_ChannelGuests, vCust_CatalogueMaster, vCust_CatalogueSibling
  Dim vCust_Banner, vCust_Url, vCust_StartUrl, vCust_ReturnUrl, vCust_Completion

  Dim vCust_EcomConfirmation, vCust_EcomEmailBody, vCust_EcomEmailAddress '...note last field is not used   

  '...place db values here in case we need to split up for various languages  
  Dim vCust_Banner_Original, vCust_Url_Original, vCust_StartUrl_Original, vCust_ReturnUrl_Original  

  Dim vCust_Eof
  Dim vContentOptions '...used to determine ecom options (used by tabs and vAction=Order and Ecom2Start.asp)


  '...Get Customer RecordSet 
  Sub sGetCust_Rs   
    vSql  = "SELECT * FROM Cust"
    sOpenDb
    Set oRs  = oDb.Execute(vSql)
  End Sub  


  Sub sGetCust_Rs_AcctId
    vSql  = "SELECT * FROM Cust ORDER BY Cust_AcctId"
    sOpenDb
    Set oRs  = oDb.Execute(vSql)
  End Sub  


  '...Get Cust Recordset
  Sub sGetCust (vCustId)
    vCust_Eof  = True
    If Len(Session("HostDb"))  = 0 Then Response.Redirect "Timeout.asp?vPage=" & Request.ServerVariables("Path_Info")
    vSql  = "SELECT * FROM Cust WHERE Cust_Id = '" & vCustId & "'"
'   sDebug
    sOpenDb    
    Set oRs  = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadCust
      vCust_Eof  = False
    End If
    Set oRs  = Nothing
    sCloseDb    
  End Sub


  Function fCustPrograms (vFunction)
    '...setups process contents
    Dim aProgs, aProg, vProg
    fCustPrograms   = ""   
    '...Store Programs Info (format: id~Title~us$~ca$~maxhrs~desc~mods~duration)
    aProgs  = Split(Trim(vCust_Programs), " " )
    For i  = 0 to Ubound(aProgs)
      aProg  = Split(aProgs(i), "~")      
      vProg_Id        = aProg(0)
      If vFunction  = "All" Or vFunction  = vProg_Id Or Left(vFunction, 1)  = "G" Then
        vProg_US        = aProg(1)
        vProg_CA        = aProg(2)
        vProg_MaxHours  = aProg(3)
        vProg_Duration  = aProg(4)
        sGetProg vProg_Id
        fCustPrograms   = fCustPrograms & "~~" & vProg_Id & "~" & vProg_Title & "~" & vProg_US & "~" & vProg_CA & "~" & vProg_MaxHours & "~" & vProg_Desc & "~" & vProg_Mods & "~" & vProg_Duration
      End If
    Next
  End Function

 
  Sub sReadCust
    vCust_No                  = oRs("Cust_No")
    vCust_Placeholder         = fIf(oRs("Cust_Placeholder"), 1, 0)
    vCust_Id                  = oRs("Cust_Id")
    vCust_AcctId              = oRs("Cust_AcctId")
    vCust_ParentId            = oRs("Cust_ParentId")
    vCust_Title               = oRs("Cust_Title")
    vCust_Lang                = oRs("Cust_Lang")
    vCust_Agent               = oRs("Cust_Agent")
    vCust_MaxSponsor          = oRs("Cust_MaxSponsor")
    vCust_IssueIds            = oRs("Cust_IssueIds")
    vCust_ResetStatus         = oRs("Cust_ResetStatus")
    vCust_IssueIdsTemplate    = oRs("Cust_IssueIdsTemplate")
    vCust_IssueIdsMemo        = oRs("Cust_IssueIdsMemo")
    vCust_ActivateIds         = oRs("Cust_ActivateIds")
    vCust_IdsSize             = oRs("Cust_IdsSize")
    vCust_FreeHours           = oRs("Cust_FreeHours")
    vCust_FreeDays            = oRs("Cust_FreeDays")
    vCust_Auto                = oRs("Cust_Auto")

    vCust_Groups              = fUnquote(oRs("Cust_Groups"))
    vCust_Programs            = fUnquote(oRs("Cust_Programs"))
    vCust_CdPrograms          = fUnquote(oRs("Cust_CdPrograms"))

    vCust_Active              = oRs("Cust_Active")
    vCust_Desc                = oRs("Cust_Desc")
    vCust_Email               = oRs("Cust_Email")
    vCust_Added               = oRs("Cust_Added")
    vCust_Expires             = oRs("Cust_Expires")
    vCust_Modified            = oRs("Cust_Modified")
    vCust_Level               = oRs("Cust_Level")
    vCust_ContentLaunch       = oRs("Cust_ContentLaunch")
    vCust_Cluster             = oRs("Cust_Cluster")
    vCust_Survey              = oRs("Cust_Survey")
    vCust_NoCert              = oRs("Cust_NoCert")
    vCust_CustomCert          = oRs("Cust_CustomCert")
    vCust_CertLogoVubiz       = oRs("Cust_CertLogoVubiz")
    vCust_CertLogoCust        = oRs("Cust_CertLogoCust")
    vCust_CertEmailAlert      = oRs("Cust_CertEmailAlert")

    vCust_AssessmentAttempts  = oRs("Cust_AssessmentAttempts")
    vCust_AssessmentScore     = oRs("Cust_AssessmentScore")
    vCust_AssessmentCert      = oRs("Cust_AssessmentCert")

    vCust_Tab1                = oRs("Cust_Tab1")
    vCust_Tab2                = oRs("Cust_Tab2")
    vCust_Tab3                = oRs("Cust_Tab3")
    vCust_Tab4                = oRs("Cust_Tab4")
    vCust_Tab5                = oRs("Cust_Tab5")
    vCust_Tab6                = oRs("Cust_Tab6")
    vCust_Tab7                = oRs("Cust_Tab7")

    vCust_Tab4Type            = oRs("Cust_Tab4Type")

    vCust_Tab1Name            = oRs("Cust_Tab1Name")
    vCust_Tab2Name            = oRs("Cust_Tab2Name")
    vCust_Tab3Name            = oRs("Cust_Tab3Name")
    vCust_Tab4Name            = oRs("Cust_Tab4Name")
    vCust_Tab5Name            = oRs("Cust_Tab5Name")
    vCust_Tab6Name            = oRs("Cust_Tab6Name")
    vCust_Tab7Name            = oRs("Cust_Tab7Name")

    vCust_InfoEditProfile     = oRs("Cust_InfoEditProfile")

    vCust_ContentOnline       = oRs("Cust_ContentOnline")
    vCust_ContentGroup        = oRs("Cust_ContentGroup")
    vCust_ContentGroup2       = oRs("Cust_ContentGroup2")
    vCust_ContentCDs          = oRs("Cust_ContentCDs")
    vCust_ContentProds        = oRs("Cust_ContentProds")

    vCust_EcomReports         = oRs("Cust_EcomReports")
    vCust_EcomSeller          = oRs("Cust_EcomSeller")
    vCust_EcomOwner           = oRs("Cust_EcomOwner")
    vCust_EcomCurrency        = oRs("Cust_EcomCurrency")
    vCust_EcomCorpRate        = oRs("Cust_EcomCorpRate")
    vCust_EcomCorpDuration    = oRs("Cust_EcomCorpDuration")
    vCust_EcomCorpProgram     = oRs("Cust_EcomCorpProgram")
    vCust_EcomGroupLicense    = oRs("Cust_EcomGroupLicense")
    vCust_EcomGroupSeat       = oRs("Cust_EcomGroupSeat")
    vCust_EcomGroup2Rates     = oRs("Cust_EcomGroup2Rates")
    vCust_EcomSplit           = oRs("Cust_EcomSplit")
    vCust_EcomDiscOptions     = oRs("Cust_EcomDiscOptions")
    vCust_EcomDisc            = oRs("Cust_EcomDisc")
    vCust_EcomDiscSplitCust   = oRs("Cust_EcomDiscSplitCust")
    vCust_EcomDiscSplitVubz   = oRs("Cust_EcomDiscSplitVubz")
    vCust_EcomDiscSplitOwnr   = oRs("Cust_EcomDiscSplitOwnr")
    vCust_EcomDiscMinUS       = oRs("Cust_EcomDiscMinUS")
    vCust_EcomDiscMinCA       = oRs("Cust_EcomDiscMinCA")
    vCust_EcomDiscMinQty      = oRs("Cust_EcomDiscMinQty")
    vCust_EcomDiscLimit       = oRs("Cust_EcomDiscLimit")
    vCust_EcomDiscOriginal    = oRs("Cust_EcomDiscOriginal")
    vCust_EcomDiscPrograms    = oRs("Cust_EcomDiscPrograms")
    vCust_EcomRepurPrograms   = oRs("Cust_EcomRepurPrograms")
    vCust_EcomRepurDisc       = oRs("Cust_EcomRepurDisc")
    vCust_EcomRepurPeriod     = oRs("Cust_EcomRepurPeriod")

    vCust_EcomConfirmation    = oRs("Cust_EcomConfirmation")
    vCust_EcomEmailAddress    = oRs("Cust_EcomEmailAddress") '...not used
    vCust_EcomEmailBody       = oRs("Cust_EcomEmailBody")

    vCust_EcomG2alert         = oRs("Cust_EcomG2alert")
    vCust_CorpAlert           = oRs("Cust_CorpAlert")

    vCust_MyWorldLaunch       = oRs("Cust_MyWorldLaunch")
    vCust_MaxUsers            = oRs("Cust_MaxUsers")
    vCust_Auth                = oRs("Cust_Auth")
    vCust_Pwd                 = oRs("Cust_Pwd")
    vCust_CritTitles          = oRs("Cust_CritTitles")   
    vCust_Resources           = oRs("Cust_Resources")
    vCust_ResourcesMaxSponsor = oRs("Cust_ResourcesMaxSponsor")
    vCust_InsertLearners      = oRs("Cust_InsertLearners")
    vCust_UpdateLearners      = oRs("Cust_UpdateLearners")
    vCust_DeleteLearners      = oRs("Cust_DeleteLearners")
    vCust_ResetLearners       = oRs("Cust_ResetLearners")
    vCust_SeedLogs            = oRs("Cust_SeedLogs")

    vCust_Completion          = oRs("Cust_Completion")

    vCust_Note1               = oRs("Cust_Note1")
    vCust_Note2               = oRs("Cust_Note2")
    vCust_Note3               = oRs("Cust_Note3")
    vCust_Note4               = oRs("Cust_Note4")
    vCust_Note5               = oRs("Cust_Note5")

    vCust_CatalogueMaster     = oRs("Cust_CatalogueMaster")
    vCust_CatalogueSibling    = oRs("Cust_CatalogueSibling")
    vCust_ChannelParent       = oRs("Cust_ChannelParent")
    vCust_ChannelV8           = oRs("Cust_ChannelV8")
    vCust_ChannelNop          = oRs("Cust_ChannelNop")
    vCust_ChannelReportsTo    = oRs("Cust_ChannelReportsTo")
    vCust_ChannelGuests       = oRs("Cust_ChannelGuests")

    vCust_Banner_Original     = oRs("Cust_Banner")
    vCust_Url_Original        = oRs("Cust_Url")
    vCust_StartUrl_Original   = oRs("Cust_StartUrl")
    vCust_ReturnUrl_Original  = oRs("Cust_ReturnUrl")

    vCust_Banner              = fMyLang(vCust_Banner_Original)
    vCust_Url                 = fMyLang(vCust_Url_Original)
    vCust_StartUrl            = fMyLang(vCust_StartUrl_Original)
    vCust_ReturnUrl           = fMyLang(vCust_ReturnUrl_Original)

    If svLang  = "EN" Then
      vCust_VuNews            = oRs("Cust_VuNews")   
    Else
      vCust_VuNews            = 0
    End If
    vCust_Scheduler           = oRs("Cust_Scheduler")   

    '...determine what ecommerce options are available (vContentOptions)
    '...if a generated group2 site then turn off all options except the 4th (AddOn2)

    If Len(Trim(vCust_ParentId)) = 4  Then 
      vContentOptions  = "NNNY"
    Else
      vContentOptions  = "NNNN" '...default options, ie: single/group/group2/addon2
      If vCust_ContentOnline            Then vContentOptions  = "Y" & Right(vContentOptions, 3)
      If vCust_ContentGroup             Then vContentOptions  = Left(vContentOptions, 1) & "Y" & Right(vContentOptions, 2)
      If vCust_ContentGroup2            Then vContentOptions  = Left(vContentOptions, 2) & "Y" & Right(vContentOptions, 1)
    End If

    '...update the Sssion Variable from input unless we have a vSource with priority
    If Instr(Session("CustReturnUrl"), "!important") = 0 Then
      If Len(vCust_ReturnUrl) > 0 Then  
        Session("CustReturnUrl") = vCust_ReturnUrl
        svCustReturnUrl          = vCust_ReturnUrl
      End If
    End If


  End Sub

  Function fMyLang(vFld)
    Dim aFld    
    fMyLang = vFld
    If Len(Trim(vFld)) > 0 Then 
      aFld = Split(vFld, "|")
      If Ubound(aFld) = 1 Then 
        If svLang = "FR" Then 
          fMyLang = aFld(1)
        Else
          fMyLang = aFld(0)
        End If
      ElseIf Ubound(aFld) = 2 Then 
        If svLang = "ES" Then 
          fMyLang = aFld(2)
        ElseIf svLang = "FR" Then 
          fMyLang = aFld(1)
        Else
          fMyLang = aFld(0)
        End If
      End If
    End If
  End Function



  Sub sExtractCust

    '...Cust_Placeholder was added Jan 27, 2016 so Customers.asp could add a placeholder record to ensure the Cust_Id is not duplicated
    '...only used in Customer.asp on "Add"

    vCust_Id                  = Request.Form("vCust_Id")
    vCust_Placeholder         = Request.Form("vCust_Placeholder")
    vCust_AcctId              = Request.Form("vCust_AcctId")
    vCust_ParentId            = Request.Form("vCust_ParentId")
    vCust_Title               = Request.Form("vCust_Title")

    vCust_Lang                = Request.Form("vCust_Lang")
    vCust_Agent               = Request.Form("vCust_Agent")
    vCust_MaxSponsor          = Request.Form("vCust_MaxSponsor")
    vCust_IssueIds            = Request.Form("vCust_IssueIds")
    vCust_ResetStatus         = Request.Form("vCust_ResetStatus")
    vCust_IssueIdsTemplate    = fDefault(Request.Form("vCust_IssueIdsTemplate"), 0)
    vCust_IssueIdsMemo        = Request.Form("vCust_IssueIdsMemo")
    vCust_ActivateIds         = Request.Form("vCust_ActivateIds")
    vCust_IdsSize             = Request.Form("vCust_IdsSize")
    vCust_FreeHours           = fDefault(Request.Form("vCust_FreeHours"), 0)
    vCust_FreeDays            = fDefault(Request.Form("vCust_FreeDays"), 0)
    vCust_Auto                = Request.Form("vCust_Auto")

    vCust_Groups              = fNoQuote(Request.Form("vCust_Groups"))
'   vCust_Programs            = fCustProgram(Trim(Replace(fNoQuote(Request.Form("vCust_Programs")), "  ", " ")))
    vCust_CdPrograms          = Trim(Replace(fNoQuote(Request.Form("vCust_CdPrograms")), "  ", " "))

    vCust_Active              = Request.Form("vCust_Active")
    vCust_Desc                = Request.Form("vCust_Desc")
    vCust_Email               = Request.Form("vCust_Email")
    vCust_Added               = fFormatDate(Request.Form("vCust_Added"))
    vCust_Expires             = fFormatDate(Request.Form("vCust_Expires"))
    vCust_Survey              = Request.Form("vCust_Survey")
    vCust_NoCert              = Request.Form("vCust_NoCert")
    vCust_CustomCert          = Request.Form("vCust_CustomCert")
    vCust_CertLogoVubiz       = Request.Form("vCust_CertLogoVubiz")
    vCust_CertLogoCust        = Request.Form("vCust_CertLogoCust")
    vCust_CertEmailAlert      = Request.Form("vCust_CertEmailAlert")

    vCust_AssessmentAttempts  = fDefault(Request.Form("vCust_AssessmentAttempts"), 0)
    vCust_AssessmentScore     = fDefault(Request.Form("vCust_AssessmentScore"), 0)
    vCust_AssessmentCert      = Request.Form("vCust_AssessmentCert")

    vCust_Level               = Request.Form("vCust_Level")
    vCust_ContentLaunch       = Request.Form("vCust_ContentLaunch")
    vCust_Cluster             = Ucase(Request.Form("vCust_Cluster"))

    vCust_Tab1                = fDefault(Request.Form("vCust_Tab1"), 0)
    vCust_Tab2                = fDefault(Request.Form("vCust_Tab2"), 0)
    vCust_Tab3                = fDefault(Request.Form("vCust_Tab3"), 0)
    vCust_Tab4                = fDefault(Request.Form("vCust_Tab4"), 0)
    vCust_Tab5                = fDefault(Request.Form("vCust_Tab5"), 0)
    vCust_Tab6                = fDefault(Request.Form("vCust_Tab6"), 0)
    vCust_Tab7                = fDefault(Request.Form("vCust_Tab7"), 0)

    vCust_Tab4Type            = Request.Form("vCust_Tab4Type")
    vCust_Tab1Name            = Request.Form("vCust_Tab1Name")
    vCust_Tab2Name            = Request.Form("vCust_Tab2Name")
    vCust_Tab3Name            = Request.Form("vCust_Tab3Name")
    vCust_Tab4Name            = Request.Form("vCust_Tab4Name")
    vCust_Tab5Name            = Request.Form("vCust_Tab5Name")
    vCust_Tab6Name            = Request.Form("vCust_Tab6Name")
    vCust_Tab7Name            = Request.Form("vCust_Tab7Name")

    vCust_InfoEditProfile     = Request.Form("vCust_InfoEditProfile")

    vCust_ContentOnline       = Request.Form("vCust_ContentOnline")
    vCust_ContentGroup        = Request.Form("vCust_ContentGroup")
    vCust_ContentGroup2       = Request.Form("vCust_ContentGroup2")
    vCust_ContentCDs          = Request.Form("vCust_ContentCDs")
    vCust_ContentProds        = Request.Form("vCust_ContentProds")

    vCust_EcomSeller          = fDefault(Request.Form("vCust_EcomSeller"), 0)
    vCust_EcomOwner           = fDefault(Request.Form("vCust_EcomOwner"), 0)

    vCust_EcomCurrency        = fDefault(Request.Form("vCust_EcomCurrency"), "CA")
    vCust_EcomCorpRate        = fDefault(Request.Form("vCust_EcomCorpRate"),0)
    vCust_EcomCorpDuration    = fDefault(Request.Form("vCust_EcomCorpDuration"),0)
    vCust_EcomCorpProgram     = Request.Form("vCust_EcomCorpProgram")
    vCust_EcomGroupLicense    = Request.Form("vCust_EcomGroupLicense")
    vCust_EcomGroupSeat       = Request.Form("vCust_EcomGroupSeat")
    vCust_EcomGroup2Rates     = Request.Form("vCust_EcomGroup2Rates")
    vCust_EcomSplit           = Request.Form("vCust_EcomSplit")
    vCust_EcomDiscOptions     = Request.Form("vCust_EcomDiscOptions")
    vCust_EcomDisc            = Request.Form("vCust_EcomDisc")
    vCust_EcomDiscSplitCust   = Request.Form("vCust_EcomDiscSplitCust")
    vCust_EcomDiscSplitVubz   = Request.Form("vCust_EcomDiscSplitVubz")
    vCust_EcomDiscSplitOwnr   = Request.Form("vCust_EcomDiscSplitOwnr")
    vCust_EcomDiscMinUS       = Request.Form("vCust_EcomDiscMinUS")
    vCust_EcomDiscMinCA       = Request.Form("vCust_EcomDiscMinCA")
    vCust_EcomDiscMinQty      = Request.Form("vCust_EcomDiscMinQty")
    vCust_EcomDiscLimit       = Request.Form("vCust_EcomDiscLimit")
    vCust_EcomDiscOriginal    = Request.Form("vCust_EcomDiscOriginal")
    vCust_EcomDiscPrograms    = Request.Form("vCust_EcomDiscPrograms")
    vCust_EcomRepurDisc       = Request.Form("vCust_EcomRepurDisc")
    vCust_EcomRepurPeriod     = Request.Form("vCust_EcomRepurPeriod")
    vCust_EcomRepurPrograms   = Request.Form("vCust_EcomRepurPrograms")

    vCust_EcomConfirmation    = Request.Form("vCust_EcomConfirmation")
    vCust_EcomEmailAddress    = Request.Form("vCust_EcomEmailAddress")
    vCust_EcomEmailBody       = Request.Form("vCust_EcomEmailBody")

    vCust_EcomG2alert         = fDefault(Request.Form("vCust_EcomG2alert"), 0)

    vCust_CorpAlert           = fDefault(Request.Form("vCust_CorpAlert"), 0)
    vCust_MyWorldLaunch       = Request.Form("vCust_MyWorldLaunch")
    vCust_MaxUsers            = Request.Form("vCust_MaxUsers")
    vCust_Auth                = fDefault(Request.Form("vCust_Auth"), 0)
    vCust_Pwd                 = fDefault(Request.Form("vCust_Pwd"), 0)
    vCust_Resources           = Request.Form("vCust_Resources")
    vCust_ResourcesMaxSponsor = fDefault(Request.Form("vCust_ResourcesMaxSponsor"), 500)
    vCust_VuNews              = fDefault(Request.Form("vCust_VuNews"), 0)
    vCust_Scheduler           = fDefault(Request.Form("vCust_Scheduler"), 0)
    vCust_SeedLogs            = Trim(Ucase(Request.Form("vCust_SeedLogs")))


    vCust_Completion          = fDefault(Request.Form("vCust_Completion"), 0)

    vCust_Banner              = Request.Form("vCust_Banner")
    vCust_Url                 = Request.Form("vCust_Url")
    vCust_StartUrl            = fNoQuote(Server.HtmlEncode(Trim(Request.Form("vCust_StartUrl"))))
    vCust_ReturnUrl           = fNoQuote(Server.HtmlEncode(Trim(Request.Form("vCust_ReturnUrl"))))

    vCust_Note1               = fUnquote(Request.Form("vCust_Note1"))
    vCust_Note2               = fUnquote(Request.Form("vCust_Note2"))
    vCust_Note3               = fUnquote(Request.Form("vCust_Note3"))
    vCust_Note4               = fUnquote(Request.Form("vCust_Note4"))
    vCust_Note5               = fUnquote(Request.Form("vCust_Note5"))

    vCust_CatalogueMaster     = fDefault(Request.Form("vCust_CatalogueMaster"), 0)
    vCust_CatalogueSibling    = fDefault(Request.Form("vCust_CatalogueSibling"), 0)
    vCust_ChannelParent       = fDefault(Request.Form("vCust_ChannelParent"), 0)
    vCust_ChannelV8           = fDefault(Request.Form("vCust_ChannelV8"), 0)
    vCust_ChannelNop          = fDefault(Request.Form("vCust_ChannelNop"), 0)
    vCust_ChannelReportsTo    = fDefault(Request.Form("vCust_ChannelReportsTo"), 0)
    vCust_ChannelGuests       = fDefault(Request.Form("vCust_ChannelGuests"), 0)

    vCust_InsertLearners      = fDefault(Request.Form("vCust_InsertLearners"), 1)
    vCust_UpdateLearners      = fDefault(Request.Form("vCust_UpdateLearners"), 1)
    vCust_DeleteLearners      = fDefault(Request.Form("vCust_DeleteLearners"), 1)
    vCust_ResetLearners       = fDefault(Request.Form("vCust_ResetLearners"), 1)
 
    If fNoValue(vCust_AcctId)             Then vCust_AcctId             = Right(vCust_Id, 4)
    If fNoValue(vCust_Auto)               Then vCust_Auto               = 0
    If fNoValue(vCust_ActivateIds)        Then vCust_ActivateIds        = 0
    If fNoValue(vCust_IssueIds)           Then vCust_IssueIds           = 0
    If fNoValue(vCust_ResetStatus)        Then vCust_ResetStatus        = 0
    If fNoValue(vCust_IssueIdsMemo)       Then vCust_IssueIdsMemo       = 0
    If fNoValue(vCust_Active)             Then vCust_Active             = 1
    If fNoValue(vCust_Lang)               Then vCust_Lang               = "EN"
    If fNoValue(vCust_Agent)              Then vCust_Agent              = "VUBZ"
    If fNoValue(vCust_IdsSize)            Then vCust_IdsSize            = 0
    If fNoValue(vCust_MaxSponsor)         Then vCust_MaxSponsor         = 0
    If fNoValue(vCust_Level)              Then vCust_Level              = 2

    If fNoValue(vCust_ContentOnline)      Then vCust_ContentOnline      = 0
    If fNoValue(vCust_ContentGroup)       Then vCust_ContentGroup       = 0
    If fNoValue(vCust_ContentGroup2)      Then vCust_ContentGroup2      = 0
    If fNoValue(vCust_ContentProds)       Then vCust_ContentProds       = 0
    If fNoValue(vCust_ContentCDs)         Then vCust_ContentCDs         = 0

    If fNoValue(vCust_Survey)             Then vCust_Survey             = 0
    If fNoValue(vCust_NoCert)             Then vCust_NoCert             = 0
    If fNoValue(vCust_CustomCert)         Then vCust_CustomCert         = 0


    If fNoValue(vCust_InfoEditProfile)    Then vCust_InfoEditProfile    = 1
    If fNoValue(vCust_Cluster)            Then vCust_Cluster            = "C0001"

    If fNoValue(vCust_EcomGroupLicense)   Then vCust_EcomGroupLicense   = 0
    If fNoValue(vCust_EcomGroupSeat)      Then vCust_EcomGroupSeat      = 0
    If fNoValue(vCust_EcomGroup2Rates)    Then vCust_EcomGroup2Rates    = "5|20~10|30~25|40~50|50~200|60"

    If fNoValue(vCust_EcomSplit)          Then vCust_EcomSplit          = 0
    If fNoValue(vCust_EcomDiscSplitCust)  Then vCust_EcomDiscSplitCust  = 100
    If fNoValue(vCust_EcomDiscSplitVubz)  Then vCust_EcomDiscSplitVubz  = 0
    If fNoValue(vCust_EcomDiscSplitOwnr)  Then vCust_EcomDiscSplitOwnr  = 0

    If fNoValue(vCust_EcomDiscOptions)    Then vCust_EcomDiscOptions    = 0
    If fNoValue(vCust_EcomDisc)           Then vCust_EcomDisc           = 0
    If fNoValue(vCust_EcomDiscMinUS)      Then vCust_EcomDiscMinUS      = 0
    If fNoValue(vCust_EcomDiscMinCA)      Then vCust_EcomDiscMinCA      = 0

    If fNoValue(vCust_EcomDiscMinQty)     Then vCust_EcomDiscMinQty     = 0
    If fNoValue(vCust_EcomDiscLimit)      Then vCust_EcomDiscLimit      = 0
    If fNoValue(vCust_EcomDiscOriginal)   Then vCust_EcomDiscOriginal   = 0

    If fNoValue(vCust_EcomRepurDisc)      Then vCust_EcomRepurDisc      = 0
    If fNoValue(vCust_EcomRepurPeriod)    Then vCust_EcomRepurPeriod    = 0
    If fNoValue(vCust_EcomRepurPrograms)  Then vCust_EcomRepurPrograms  = 0
    If fNoValue(vCust_EcomConfirmation)   Then vCust_EcomConfirmation   = 0
    If fNoValue(vCust_MaxUsers)           Then vCust_MaxUsers           = 0

  End Sub


  '...Ensure the program string has current max hours
  Function fCustProgram (vCust_Programs)
    Dim vUpdate, aProgs, aProg
    vUpdate  = False
    aProgs  = Split(Trim(vCust_Programs), " ")
    For i  = 0 To Ubound(aProgs)      
      aProg  = Split(aProgs(i), "~")
      sGetProg aProg(0)
      If aProg(3) <> Cstr(vProg_Length) Then
        vUpdate  = True
        aProg(3)  = vProg_Length
        aProgs(i)  = Join(aProg, "~")
      End If
    Next    
    If vUpdate Then
      fCustProgram  = Join(aProgs, " ")
    Else
      fCustProgram  = vCust_Programs
    End If
  End Function
  

  Sub sInsertCust
    vFileOk  = False

    '...added last check on 2015-08-11
    If Len(vCust_AcctId) <> 4 Or Not IsNumeric(vCust_AcctId) Or Not fIsOkCustId(vCust_Id) Then
      vFileOk  = False
      Err.Description = fDefault(Err.Description, "The Customer Id is using values that are not unique.")
      Exit Sub
    End If    

    '...change to ANSI_WARNINGS ON from OFF Nov 19, 2015 for computed columns
    vSql  = "SET ANSI_WARNINGS ON " _                                                                                  
          & "INSERT INTO Cust (" _
          & "  Cust_Id,"_
          & "  Cust_AcctId,"_
          & "  Cust_ParentId,"_ 
          & "  Cust_Title,"_ 
          & "  Cust_Lang,"_ 
          & "  Cust_Agent,"_ 
          & "  Cust_MaxSponsor,"_ 
          & "  Cust_IssueIds,"_ 
          & "  Cust_ResetStatus,"_ 
          & "  Cust_IssueIdsTemplate,"_ 
          & "  Cust_IssueIdsMemo,"_ 
          & "  Cust_ActivateIds,"_ 
          & "  Cust_IdsSize,"_ 
          & "  Cust_FreeHours,"_ 
          & "  Cust_FreeDays,"_ 
          & "  Cust_Auto,"_ 
          & "  Cust_Groups,"_ 
          & "  Cust_Programs,"_ 
          & "  Cust_CdPrograms,"_ 
          & "  Cust_ContentOnline,"_ 
          & "  Cust_ContentGroup,"_ 
          & "  Cust_ContentGroup2,"_ 
          & "  Cust_ContentProds,"_ 
          & "  Cust_ContentCDs,"_ 
          & "  Cust_Active,"_ 
          & "  Cust_Desc,"_ 
          & "  Cust_Email,"_ 
          & "  Cust_Expires,"_ 
          & "  Cust_Cluster,"_ 
          & "  Cust_Level,"_ 
          & "  Cust_ContentLaunch,"_ 
          & "  Cust_Survey,"_ 
          & "  Cust_NoCert,"_ 
          & "  Cust_CustomCert,"_ 
          & "  Cust_CertLogoVubiz,"_ 
          & "  Cust_CertLogoCust,"_ 
          & "  Cust_CertEmailAlert,"_ 
          & "  Cust_AssessmentAttempts,"_ 
          & "  Cust_AssessmentScore,"_ 
          & "  Cust_AssessmentCert,"_ 
          & "  Cust_Tab1,"_ 
          & "  Cust_Tab2,"_ 
          & "  Cust_Tab3,"_ 
          & "  Cust_Tab4,"_ 
          & "  Cust_Tab5,"_ 
          & "  Cust_Tab6,"_ 
          & "  Cust_Tab7,"_ 
          & "  Cust_Tab4Type,"_ 
          & "  Cust_Tab1Name,"_ 
          & "  Cust_Tab2Name,"_ 
          & "  Cust_Tab3Name,"_ 
          & "  Cust_Tab4Name,"_ 
          & "  Cust_Tab5Name,"_ 
          & "  Cust_Tab6Name,"_ 
          & "  Cust_Tab7Name,"_ 
          & "  Cust_InfoEditProfile,"_ 
          & "  Cust_EcomCurrency,"_ 
          & "  Cust_EcomCorpRate,"_ 
          & "  Cust_EcomCorpDuration,"_ 
          & "  Cust_EcomCorpProgram,"_ 
          & "  Cust_EcomGroupLicense,"_ 
          & "  Cust_EcomGroupSeat,"_ 
          & "  Cust_EcomGroup2Rates,"_ 
          & "  Cust_EcomSplit,"_ 
          & "  Cust_EcomDiscOptions,"_ 
          & "  Cust_EcomDisc,"_ 
          & "  Cust_EcomDiscSplitCust,"_ 
          & "  Cust_EcomDiscSplitVubz,"_ 
          & "  Cust_EcomDiscSplitOwnr,"_ 
          & "  Cust_EcomDiscMinUS,"_ 
          & "  Cust_EcomDiscMinCA,"_ 
          & "  Cust_EcomDiscMinQty,"_ 
          & "  Cust_EcomDiscLimit,"_ 
          & "  Cust_EcomDiscOriginal,"_ 
          & "  Cust_EcomDiscPrograms,"_ 
          & "  Cust_EcomRepurPrograms,"_ 
          & "  Cust_EcomRepurDisc,"_ 
          & "  Cust_EcomRepurPeriod,"_ 

          & "  Cust_EcomConfirmation,"_ 
          & "  Cust_EcomEmailAddress,"_ 
          & "  Cust_EcomEmailBody,"_ 

          & "  Cust_EcomSeller,"_ 
          & "  Cust_EcomOwner,"_ 
          & "  Cust_EcomG2alert,"_ 

          & "  Cust_CorpAlert,"_ 
          & "  Cust_MyWorldLaunch,"_ 
          & "  Cust_MaxUsers,"_ 
          & "  Cust_Auth,"_ 
          & "  Cust_Pwd,"_ 
          & "  Cust_Resources,"_ 
          & "  Cust_ResourcesMaxSponsor,"_ 
          & "  Cust_VuNews,"_ 
          & "  Cust_Scheduler,"_ 
          & "  Cust_SeedLogs,"_

          & "  Cust_Completion,"_ 

          & "  Cust_Banner,"_ 
          & "  Cust_Url,"_ 
          & "  Cust_StartUrl,"_ 
          & "  Cust_ReturnUrl,"_ 

          & "  Cust_Note1,"_
          & "  Cust_Note2,"_
          & "  Cust_Note3,"_
          & "  Cust_Note4,"_
          & "  Cust_Note5,"_

          & "  Cust_CatalogueMaster,"_
          & "  Cust_CatalogueSibling,"_
          & "  Cust_ChannelParent,"_
          & "  Cust_ChannelV8,"_
          & "  Cust_ChannelNop,"_
          & "  Cust_ChannelReportsTo,"_
          & "  Cust_ChannelGuests,"_

          & "  Cust_InsertLearners,"_
          & "  Cust_UpdateLearners,"_
          & "  Cust_DeleteLearners,"_
          & "  Cust_ResetLearners"_
          & ") "_

          & "VALUES ("_
          & "  '" & vCust_Id & "',"_ 
          & "  '" & vCust_AcctId & "',"_ 
          & "  '" & vCust_ParentId & "',"_ 
          & "  '" & fUnQuote(vCust_Title) & "',"_ 
          & "  '" & vCust_Lang & "',"_ 
          & "  '" & vCust_Agent   & "',"_ 
          & "   " & vCust_MaxSponsor & ","_ 
          & "   " & vCust_IssueIds & ","_ 
          & "   " & vCust_ResetStatus & ","_ 
          & "   " & vCust_IssueIdsMemo & ","_ 
          & "  '" & vCust_IssueIdsTemplate & "',"_ 
          & "   " & vCust_ActivateIds & ","_ 
          & "   " & vCust_IdsSize & ","_ 
          & "   " & vCust_FreeHours & ","_ 
          & "   " & vCust_FreeDays & ","_ 
          & "   " & vCust_Auto & ","_ 
          & "  '" & fUnquote(vCust_Groups) & "',"_ 
          & "  '" & fUnquote(vCust_Programs) & "',"_ 
          & "  '" & fUnquote(vCust_CdPrograms)   & "',"_ 
          & "   " & vCust_ContentOnline & ","_ 
          & "   " & vCust_ContentGroup & ","_ 
          & "   " & vCust_ContentGroup2 & ","_ 
          & "   " & vCust_ContentProds & ","_ 
          & "   " & vCust_ContentCDs & ","_ 
          & "   " & vCust_Active & ","_ 
          & "  '" & fUnquote(vCust_Desc) & "',"_ 
          & "  '" & vCust_Email & "',"_ 
          & "  '" & vCust_Expires & "',"_ 
          & "  '" & vCust_Cluster & "',"_ 
          & "   " & vCust_Level & ","_  
          & "  '" & vCust_ContentLaunch & "',"_  
          & "   " & vCust_Survey & ","_  
          & "   " & vCust_NoCert & ","_  
          & "   " & vCust_CustomCert & ","_ 
          & "  '" & vCust_CertLogoVubiz & "',"_ 
          & "  '" & vCust_CertLogoCust & "',"_ 
          & "  '" & vCust_CertEmailAlert & "',"_ 
          & "   " & vCust_AssessmentAttempts & ","_ 
          & "   " & vCust_AssessmentScore & ","_ 
          & "  '" & vCust_AssessmentCert & "',"_ 
          & "   " & vCust_Tab1 & ","_ 
          & "   " & vCust_Tab2 & ","_ 
          & "   " & vCust_Tab3 & ","_ 
          & "   " & vCust_Tab4 & ","_ 
          & "   " & vCust_Tab5 & ","_ 
          & "   " & vCust_Tab6 & ","_ 
          & "   " & vCust_Tab7 & ","_  
          & "  '" & vCust_Tab4Type & "',"_ 
          & "  '" & vCust_Tab1Name & "',"_ 
          & "  '" & vCust_Tab2Name & "',"_ 
          & "  '" & vCust_Tab3Name & "',"_ 
          & "  '" & vCust_Tab4Name & "',"_ 
          & "  '" & vCust_Tab5Name & "',"_ 
          & "  '" & vCust_Tab6Name & "',"_ 
          & "  '" & vCust_Tab7Name & "',"_ 
          & "  '" & vCust_InfoEditProfile & "',"_ 
          & "  '" & vCust_EcomCurrency & "',"_ 
          & "   " & vCust_EcomGroupLicense & ","_ 
          & "   " & vCust_EcomCorpRate & ","_ 
          & "   " & vCust_EcomCorpDuration & ","_ 
          & "  '" & vCust_EcomCorpProgram & "',"_ 
          & "   " & vCust_EcomGroupSeat & ","_ 
          & "  '" & vCust_EcomGroup2Rates & "',"_ 
          & "   " & vCust_EcomSplit & ","_ 
          & "   " & vCust_EcomDiscOptions & ","_ 
          & "   " & vCust_EcomDisc & ","_ 
          & "   " & vCust_EcomDiscSplitCust & ","_ 
          & "   " & vCust_EcomDiscSplitVubz & ","_ 
          & "   " & vCust_EcomDiscSplitOwnr & ","_ 
          & "   " & vCust_EcomDiscMinUS & ","_ 
          & "   " & vCust_EcomDiscMinCA & ","_ 
          & "   " & vCust_EcomDiscMinQty & ","_ 
          & "   " & vCust_EcomDiscLimit & ","_ 
          & "   " & vCust_EcomDiscOriginal & ","_ 
          & "  '" & vCust_EcomDiscPrograms & "',"_ 
          & "   " & vCust_EcomRepurPrograms & ","_ 
          & "   " & vCust_EcomRepurDisc & ","_ 
          & "   " & vCust_EcomRepurPeriod & ","_ 
          & "   " & vCust_EcomConfirmation & ","_ 
          & "  '" & vCust_EcomEmailAddress & "',"_ 
          & "  '" & vCust_EcomEmailBody & "',"_ 
          & "   " & vCust_EcomSeller & ","_ 
          & "   " & vCust_EcomOwner & ","_ 
          & "   " & vCust_EcomG2alert & ","_ 
          & "   " & vCust_CorpAlert & ","_ 
          & "  '" & vCust_MyWorldLaunch & "',"_ 
          & "   " & vCust_MaxUsers & ","_ 
          & "   " & vCust_Auth & ","_ 
          & "   " & vCust_Pwd & ","_ 
          & "  '" & vCust_Resources & "',"_ 
          & "   " & vCust_ResourcesMaxSponsor & ","_ 
          & "   " & vCust_VuNews & ","_ 
          & "   " & vCust_Scheduler & ","_ 
          & "  '" & vCust_SeedLogs & "',"_

          & "   " & vCust_Completion & ","_ 

          & "  '" & vCust_Banner & "',"_ 
          & "  '" & vCust_Url & "',"_ 
          & "  '" & fNoQuote(vCust_StartUrl) & "',"_ 
          & "  '" & fNoQuote(vCust_ReturnUrl) & "',"_ 

          & "  '" & vCust_Note1 & "',"_
          & "  '" & vCust_Note2 & "',"_
          & "  '" & vCust_Note3 & "',"_
          & "  '" & vCust_Note4 & "',"_
          & "  '" & vCust_Note5 & "',"_

          & "   " & vCust_CatalogueMaster & " ,"_
          & "   " & vCust_CatalogueSibling & " ,"_
          & "   " & vCust_ChannelParent & " ,"_
          & "   " & vCust_ChannelV8 & " ,"_
          & "   " & vCust_ChannelNop & " ,"_
          & "   " & vCust_ChannelReportsTo & " ,"_
          & "   " & vCust_ChannelGuests & " ,"_

          & "   " & vCust_InsertLearners & ","_
          & "   " & vCust_UpdateLearners & ","_
          & "   " & vCust_DeleteLearners & ","_
          & "   " & vCust_ResetLearners  & " "_
          & "  )"

'   sDebug
    sOpenDb
    On Error Resume Next
    oDb.Execute(vSql)
    sCloseDb

    vFileOk = False
    If Err.Number  = 0 Or Err.Number  = "" Then 
      sAddInternalMemb vCust_AcctId                           '...add internal learners    
      If vCust_Level = 4 Then sSetupRepository vCust_AcctId   '...add repository for corporate accounts  
    End If    
  End Sub


  Function fCustMaxUsers(vCustId)
    vSql  = "SELECT Cust_MaxUsers FROM Cust WHERE Cust_Id  = '" & vCustId & "'"
    sOpenDb
    Set oRs  = oDb.Execute(vSql)
'   sDebug
    If oRs.Eof Then 
      fCustMaxUsers = 0
    Else
      fCustMaxUsers = oRs("Cust_MaxUsers")
    End If
    sCloseDb
  End Function


  Function fCustExpires(vCustId)
    vSql  = "SELECT Cust_Expires FROM Cust WHERE Cust_Id  = '" & vCustId & "'"
    sOpenDb3
    Set oRs3  = oDb3.Execute(vSql)
'   sDebug
    If oRs3.Eof Then 
      fCustExpires = 0
    Else
      fCustExpires = oRs3("Cust_Expires")
    End If
    sCloseDb3
  End Function  


  Sub sUpdateCustMaxUsers(vCustId, vLearners)
    vSql  = "UPDATE Cust SET Cust_MaxUsers = Cust_MaxUsers + " & vLearners & " WHERE Cust_Id  = '" & vCustId & "'"
'   sDebug
    sOpenDb
    On Error Resume Next
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sUpdateCustTitle(vCustId, vTitle)
    vSql  = "UPDATE Cust SET Cust_Title = '" & fUnQuote(vTitle) & "' WHERE Cust_Id  = '" & vCustId & "'"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub
  
  
  Sub sUpdateCustDateAdded(vCustId, vDate)
    vSql  = "UPDATE Cust SET Cust_Added = '" & fFormatSqlDate(vDate) & "' WHERE Cust_Id  = '" & vCustId & "'"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub  
    

  Sub sUpdateCust
  
    '...added last check on 2015-08-11 - modified to handle placeholders
'    If Not IsNumeric(vCust_AcctId) Or Not fIsOkCustId(vCust_Id) Then
'      vFileOk  = False
'      Err.Description = fDefault(Err.Description, "The Customer Id is using values that are not unique.")
'      Exit Sub
'    End If   
  
     '...change to ANSI_WARNINGS ON from OFF Nov 19, 2015 for computed columns
    vSql  = "SET ANSI_WARNINGS ON "
    vSql  = vSql & "UPDATE Cust SET"
    vSql  = vSql & " Cust_Placeholder          = 0                                   , " 
    vSql  = vSql & " Cust_AcctId               = '" & fUnquote(vCust_AcctId)     & "', " 
    vSql  = vSql & " Cust_ParentId             = '" & fUnquote(vCust_ParentId)   & "', " 
    vSql  = vSql & " Cust_Title                = '" & fUnQuote(vCust_Title)      & "', " 
    vSql  = vSql & " Cust_Lang                 = '" & vCust_Lang                 & "', " 
    vSql  = vSql & " Cust_Agent                = '" & vCust_Agent                & "', " 
    vSql  = vSql & " Cust_MaxSponsor           =  " & vCust_MaxSponsor           & " , " 
    vSql  = vSql & " Cust_IssueIds             =  " & vCust_IssueIds             & " , " 
    vSql  = vSql & " Cust_ResetStatus          =  " & vCust_ResetStatus          & " , " 
    vSql  = vSql & " Cust_IssueIdsTemplate     = '" & vCust_IssueIdsTemplate     & "', " 
    vSql  = vSql & " Cust_IssueIdsMemo         =  " & vCust_IssueIdsMemo         & " , " 
    vSql  = vSql & " Cust_ActivateIds          =  " & vCust_ActivateIds          & " , " 
    vSql  = vSql & " Cust_IdsSize              =  " & vCust_IdsSize              & " , " 
    vSql  = vSql & " Cust_FreeHours            =  " & vCust_FreeHours            & " , " 
    vSql  = vSql & " Cust_FreeDays             =  " & vCust_FreeDays             & " , " 
    vSql  = vSql & " Cust_Auto                 =  " & vCust_Auto                 & " , " 
    vSql  = vSql & " Cust_Groups               = '" & fUnquote(vCust_Groups)     & "', " 
    vSql  = vSql & " Cust_Programs             = '" & fUnquote(vCust_Programs)   & "', " 
    vSql  = vSql & " Cust_CdPrograms           = '" & fUnquote(vCust_CdPrograms) & "', " 
    vSql  = vSql & " Cust_ContentOnline        =  " & vCust_ContentOnline        & " , " 
    vSql  = vSql & " Cust_ContentGroup         =  " & vCust_ContentGroup         & " , " 
    vSql  = vSql & " Cust_ContentGroup2        =  " & vCust_ContentGroup2        & " , " 
    vSql  = vSql & " Cust_ContentCDs           =  " & vCust_ContentCDs           & " , " 
    vSql  = vSql & " Cust_ContentProds         =  " & vCust_ContentProds         & " , " 
    vSql  = vSql & " Cust_Active               =  " & vCust_Active               & " , " 
    vSql  = vSql & " Cust_Desc                 = '" & fUnquote(vCust_Desc)       & "', " 
    vSql  = vSql & " Cust_Email                = '" & vCust_Email                & "', " 
    vSql  = vSql & " Cust_Added                = '" & vCust_Added                & "', " 
    vSql  = vSql & " Cust_Expires              = '" & vCust_Expires              & "', " 
    vSql  = vSql & " Cust_Cluster              = '" & vCust_Cluster              & "', " 
    vSql  = vSql & " Cust_Survey               =  " & vCust_Survey               & " , " 
    vSql  = vSql & " Cust_NoCert               =  " & vCust_NoCert               & " , " 
    vSql  = vSql & " Cust_CustomCert           =  " & vCust_CustomCert           & " , " 
    vSql  = vSql & " Cust_CertLogoVubiz        = '" & vCust_CertLogoVubiz        & "', " 
    vSql  = vSql & " Cust_CertLogoCust         = '" & vCust_CertLogoCust         & "', " 
    vSql  = vSql & " Cust_CertEmailAlert       = '" & vCust_CertEmailAlert       & "', " 
    vSql  = vSql & " Cust_AssessmentAttempts   =  " & vCust_AssessmentAttempts   & " , " 
    vSql  = vSql & " Cust_AssessmentScore      =  " & vCust_AssessmentScore      & " , " 
    vSql  = vSql & " Cust_AssessmentCert       = '" & vCust_AssessmentCert       & "', " 
    vSql  = vSql & " Cust_Level                =  " & vCust_Level                & " , " 
    vSql  = vSql & " Cust_ContentLaunch        = '" & vCust_ContentLaunch        & "', " 
    vSql  = vSql & " Cust_Tab1                 =  " & vCust_Tab1                 & " , " 
    vSql  = vSql & " Cust_Tab2                 =  " & vCust_Tab2                 & " , " 
    vSql  = vSql & " Cust_Tab3                 =  " & vCust_Tab3                 & " , " 
    vSql  = vSql & " Cust_Tab4                 =  " & vCust_Tab4                 & " , " 
    vSql  = vSql & " Cust_Tab5                 =  " & vCust_Tab5                 & " , " 
    vSql  = vSql & " Cust_Tab6                 =  " & vCust_Tab6                 & " , " 
    vSql  = vSql & " Cust_Tab7                 =  " & vCust_Tab7                 & " , " 
    vSql  = vSql & " Cust_Tab4Type             = '" & vCust_Tab4Type             & "', " 
    vSql  = vSql & " Cust_Tab1Name             = '" & vCust_Tab1Name             & "', " 
    vSql  = vSql & " Cust_Tab2Name             = '" & vCust_Tab2Name             & "', " 
    vSql  = vSql & " Cust_Tab3Name             = '" & vCust_Tab3Name             & "', " 
    vSql  = vSql & " Cust_Tab4Name             = '" & vCust_Tab4Name             & "', " 
    vSql  = vSql & " Cust_Tab5Name             = '" & vCust_Tab5Name             & "', " 
    vSql  = vSql & " Cust_Tab6Name             = '" & vCust_Tab6Name             & "', " 
    vSql  = vSql & " Cust_Tab7Name             = '" & vCust_Tab7Name             & "', " 
    vSql  = vSql & " Cust_InfoEditProfile      =  " & vCust_InfoEditProfile      & " , " 

    vSql  = vSql & " Cust_EcomSeller           =  " & fDefault(vCust_EcomSeller, 0) & " , " 
    vSql  = vSql & " Cust_EcomOwner            =  " & fDefault(vCust_EcomOwner, 0)  & " , " 
    vSql  = vSql & " Cust_EcomCurrency         = '" & fDefault(vCust_EcomCurrency, "CA") & "', " 
    vSql  = vSql & " Cust_EcomCorpRate         =  " & vCust_EcomCorpRate         & " , " 
    vSql  = vSql & " Cust_EcomCorpDuration     =  " & vCust_EcomCorpDuration     & " , " 
    vSql  = vSql & " Cust_EcomCorpProgram      = '" & vCust_EcomCorpProgram      & "', " 
    vSql  = vSql & " Cust_EcomGroupLicense     =  " & vCust_EcomGroupLicense     & " , " 
    vSql  = vSql & " Cust_EcomGroupSeat        =  " & vCust_EcomGroupSeat        & " , " 
    vSql  = vSql & " Cust_EcomGroup2Rates      = '" & vCust_EcomGroup2Rates      & "', " 
    vSql  = vSql & " Cust_EcomSplit            =  " & vCust_EcomSplit            & " , " 
    vSql  = vSql & " Cust_EcomDiscOptions      =  " & vCust_EcomDiscOptions      & " , " 
    vSql  = vSql & " Cust_EcomDisc             =  " & vCust_EcomDisc             & " , " 
    vSql  = vSql & " Cust_EcomDiscSplitCust    =  " & vCust_EcomDiscSplitCust    & " , " 
    vSql  = vSql & " Cust_EcomDiscSplitVubz    =  " & vCust_EcomDiscSplitVubz    & " , " 
    vSql  = vSql & " Cust_EcomDiscSplitOwnr    =  " & vCust_EcomDiscSplitOwnr    & " , " 
    vSql  = vSql & " Cust_EcomDiscMinUS        =  " & vCust_EcomDiscMinUS        & " , " 
    vSql  = vSql & " Cust_EcomDiscMinCA        =  " & vCust_EcomDiscMinCA        & " , " 
    vSql  = vSql & " Cust_EcomDiscMinQty       =  " & vCust_EcomDiscMinQty       & " , " 
    vSql  = vSql & " Cust_EcomDiscLimit        =  " & vCust_EcomDiscLimit        & " , " 
    vSql  = vSql & " Cust_EcomDiscOriginal     =  " & vCust_EcomDiscOriginal     & " , " 
    vSql  = vSql & " Cust_EcomDiscPrograms     = '" & vCust_EcomDiscPrograms     & "', " 
    vSql  = vSql & " Cust_EcomRepurDisc        =  " & vCust_EcomRepurDisc        & " , " 
    vSql  = vSql & " Cust_EcomRepurPeriod      =  " & vCust_EcomRepurPeriod      & " , " 
    vSql  = vSql & " Cust_EcomRepurPrograms    =  " & vCust_EcomRepurPrograms    & " , " 
    vSql  = vSql & " Cust_EcomConfirmation     =  " & vCust_EcomConfirmation     & " , " 
    vSql  = vSql & " Cust_EcomEmailAddress     = '" & vCust_EcomEmailAddress     & "', " 
    vSql  = vSql & " Cust_EcomEmailBody        = '" & vCust_EcomEmailBody        & "', " 
    vSql  = vSql & " Cust_EcomG2alert          =  " & vCust_EcomG2alert          & " , " 
    vSql  = vSql & " Cust_CorpAlert            =  " & vCust_CorpAlert            & " , " 

    vSql  = vSql & " Cust_MyWorldLaunch        = '" & vCust_MyWorldLaunch        & "', " 
    vSql  = vSql & " Cust_MaxUsers             =  " & vCust_MaxUsers             & " , " 
    vSql  = vSql & " Cust_Auth                 =  " & vCust_Auth                 & " , " 
    vSql  = vSql & " Cust_Pwd                  =  " & vCust_Pwd                  & " , " 
    vSql  = vSql & " Cust_Resources            = '" & vCust_Resources            & "', " 
    vSql  = vSql & " Cust_ResourcesMaxSponsor  =  " & vCust_ResourcesMaxSponsor  & " , " 
    vSql  = vSql & " Cust_VuNews               =  " & vCust_VuNews               & " , " 
    vSql  = vSql & " Cust_Scheduler            =  " & vCust_Scheduler            & " , " 
    vSql  = vSql & " Cust_SeedLogs             = '" & vCust_SeedLogs             & "', " 

    vSql  = vSql & " Cust_Completion           =  " & vCust_Completion           & " , " 

    vSql  = vSql & " Cust_Banner               = '" & vCust_Banner               & "', " 
    vSql  = vSql & " Cust_Url                  = '" & vCust_Url                  & "', " 
    vSql  = vSql & " Cust_StartUrl             = '" & fNoQuote(vCust_StartUrl)   & "', " 
    vSql  = vSql & " Cust_ReturnUrl            = '" & fNoQuote(vCust_ReturnUrl)  & "', " 

    vSql  = vSql & " Cust_Note1                = '" & vCust_Note1                & "', " 
    vSql  = vSql & " Cust_Note2                = '" & vCust_Note2                & "', " 
    vSql  = vSql & " Cust_Note3                = '" & vCust_Note3                & "', " 
    vSql  = vSql & " Cust_Note4                = '" & vCust_Note4                & "', " 
    vSql  = vSql & " Cust_Note5                = '" & vCust_Note5                & "', " 

    vSql  = vSql & " Cust_CatalogueMaster      =  " & vCust_CatalogueMaster      & " , " 
    vSql  = vSql & " Cust_CatalogueSibling     =  " & vCust_CatalogueSibling     & " , " 
    vSql  = vSql & " Cust_ChannelParent        =  " & vCust_ChannelParent        & " , " 
    vSql  = vSql & " Cust_ChannelV8            =  " & vCust_ChannelV8            & " , " 
    vSql  = vSql & " Cust_ChannelNop           =  " & vCust_ChannelNop           & " , " 
    vSql  = vSql & " Cust_ChannelReportsTo     =  " & vCust_ChannelReportsTo     & " , " 
    vSql  = vSql & " Cust_ChannelGuests        =  " & vCust_ChannelGuests        & " , " 

    vSql  = vSql & " Cust_InsertLearners       =  " & vCust_InsertLearners       & " , "
    vSql  = vSql & " Cust_UpdateLearners       =  " & vCust_UpdateLearners       & " , "
    vSql  = vSql & " Cust_DeleteLearners       =  " & vCust_DeleteLearners       & " , "
    vSql  = vSql & " Cust_ResetLearners        =  " & vCust_ResetLearners        & "   "

    vSql  = vSql & " WHERE Cust_Id  = '" & vCust_Id                              & "'  "
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb

      '... add internal learners/repository if first time - ie placeholder = true
    If vCust_Placeholder = "1" Then 
      sAddInternalMemb vCust_AcctId                           '...add internal learners    
      If vCust_Level = 4 Then sSetupRepository vCust_AcctId   '...add repository for corporate accounts  
    End If    
  
  End Sub


  Sub sUpdateCustPrograms 
    vSql  = "UPDATE Cust SET"
    vSql  = vSql & " Cust_Programs  = '" & vCust_Programs  & "'  " 
    vSql  = vSql & " WHERE Cust_Id  = '" & vCust_Id        & "'  "
'   sDebug
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub


  Sub sUpdateCustCritTitles (vCust_Id)
    vSql  = "UPDATE Cust SET Cust_CritTitles  = '" & Trim(vCust_CritTitles) & "' WHERE Cust_Id  = '" & vCust_Id & "'  "
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub
  

  '...get all Cust
  Function fCustOptions
    Dim oRs
    fCustOptions  = ""
    sOpenDb
    vSql  = "Select * FROM Cust "
    Set oRs  = oDb.Execute(vSql)    
    Do While Not oRs.EOF 
      fCustOptions  = fCustOptions & "<option>" & oRs("Cust_Id") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop
    sCloseDb           
  End Function


  '...setup a repository folder
  Sub sSetupRepository(vAcctId)
    Dim oFs, vPath
    '...this is the folder for the repository
    vPath  = Lcase(Server.MapPath("\V5\Repository\" & svHostDb & "\" & vAcctId)) & "\"
 '  sDebug "vPath", vPath
    Set oFs  = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Not oFs.FolderExists(vPath) Then
      oFs.CreateFolder vPath
    End If
    If Err <> 0 Then 
      Response.Write "<p><center><b><font face='Verdana' size='1' color='#FF0000'>Error: The Repository:<br>" & vPath & "<br>could NOT be setup!  Please inform Systems.</font></b></center></p>"
    End If
  End Sub
  
  
  '...delete a repository folder
  Sub sDeleteRepository(vAcctId)
    Dim oFs, vPath
    '...this is the folder for the repository
    vPath  = Lcase(Server.MapPath("\V5\Repository\" & svHostDb & "\" & vAcctId)) & "\"
    vPath  = Lcase(Server.MapPath("\V5\Repository\" & svHostDb & "\" & vAcctId))
 '  sDebug "vPath", vPath
    Set oFs  = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If oFs.FolderExists(vPath) Then
      oFs.DeleteFolder(vPath)
    End If
    If Err <> 0 Then 
      Response.Write "<p><center><span class='c5'>Error: The Repository:<br>" & vPath & "<br>could NOT be deleted!  Please inform Systems.</span></center></p>"
    End If
  End Sub  
  


  '...Get Prog Duration
  Function fCustProgDuration (vCustId, vProgId)
    Dim aProgs, aProg, i
    fCustProgDuration  = 0
    vSql  = "SELECT Cust_Programs FROM Cust WHERE Cust_Id = '" & vCustId & "'"
    sOpenDb2    
    Set oRs2  = oDb2.Execute(vSql)
    vCust_Programs  = oRs2("Cust_Programs")
    Set oRs2  = Nothing
    sCloseDb2
    aProgs  = Split(Trim(vCust_Programs), " ")
    For i  = 0 to uBound(aProgs)
      If vProgId  = Left(aProgs(i), 7) Then
        aProg  = Split(aProgs(i), "~")
        fCustProgDuration  = aProg(4)
        Exit Function         
      End If
    Next
  End Function  


  '...Get Prog Amount
  Function fCustProgAmount (vCustId, vProgId, vCurrency)
    Dim aProgs, aProg, i
    fCustProgAmount  = 0
    vSql  = "SELECT Cust_Programs FROM Cust WHERE Cust_Id = '" & vCustId & "'"
    sOpenDb2    
    Set oRs2  = oDb2.Execute(vSql)
    vCust_Programs  = oRs2("Cust_Programs")
    Set oRs2  = Nothing
    sCloseDb2
    aProgs  = Split(Trim(vCust_Programs), " ")
    For i  = 0 to uBound(aProgs)
      If vProgId  = Left(aProgs(i), 7) Then
        aProg  = Split(aProgs(i), "~")
        If vCurrency  = "CA" Then
          fCustProgAmount  = aProg(2)
        Else
          fCustProgAmount  = aProg(1)       
        End If
        Exit Function         
      End If
    Next
  End Function
 

  Sub sDeleteCust '...delete entire account using vCust_AcctId
    sOpenDb    

    oDb.Execute("DELETE FROM Cust WHERE Cust_AcctId     = '" & vCust_AcctId & "'")
    oDb.Execute("DELETE FROM Keys WHERE Keys_AcctId     = '" & vCust_AcctId & "'")
    oDb.Execute("DELETE FROM Logs WHERE Logs_AcctId     = '" & vCust_AcctId & "'")
    oDb.Execute("DELETE FROM Memb WHERE Memb_AcctId     = '" & vCust_AcctId & "'")
    oDb.Execute("DELETE FROM Catl WHERE Catl_CustId     = '" & vCust_Id & "'")
    oDb.Execute("DELETE FROM Ecom WHERE Ecom_CustId     = '" & vCust_Id & "'")         '...delete Ecom PARENT records added Nov 19, 2015 to stagingweb
    oDb.Execute("DELETE FROM Ecom WHERE Ecom_NewAcctId  = '" & vCust_AcctId & "'")     '...delete Ecom CHILD  records added Nov 19, 2015 to stagingweb

    '...delete Task assets
    vSql  = "SELECT TskH_No FROM TskH WHERE TskH_AcctId  = '" & vCust_AcctId & "'"
    Set oRs  = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      vSql  = "DELETE FROM TskD WHERE TskD_No  = " & oRs("TskH_No")
      oDb.Execute(vSql)
      oRs.MoveNext
    Loop
    Set oRs  = Nothing 
    oDb.Execute("DELETE FROM TskH WHERE TskH_AcctId = '" & vCust_AcctId & "'")
    sCloseDb
    sDeleteRepository(vCust_AcctId)
  End Sub


  Sub sDeleteLinkedCust '...delete just customer record using vCust_Id
    sOpenDb    
    oDb.Execute("DELETE FROM Cust WHERE Cust_Id = '" & vCust_Id & "'")
    sCloseDb    
  End Sub


  Function fHasLinkedCust(vCust_AcctId)
    fHasLinkedCust = False
    vSql  = "SELECT Cust_Id FROM Cust WHERE Cust_AcctId  = '" & vCust_AcctId & "' AND RIGHT(Cust_Id, 4) <> '" & vCust_AcctId & "'"
    sOpenDb4 
    Set oRs4  = oDb4.Execute(vSql)
    Do While Not oRs4.Eof 
      fHasLinkedCust = True
      oRs4.MoveNext
    Loop
    Set oRs4  = Nothing
    sCloseDb4  
  End Function


  Function fHasChildCust(vCust_AcctId)
    fHasChildCust = False
    vSql  = "SELECT Cust_Id FROM Cust WHERE Cust_ParentId  = '" & vCust_AcctId & "'"
    sOpenDb4 
    Set oRs4  = oDb4.Execute(vSql)
    Do While Not oRs4.Eof 
      fHasChildCust = True
      oRs4.MoveNext
    Loop
    Set oRs4  = Nothing
    sCloseDb4  
  End Function


  Function fIsOkCustId(vCust_Id)
    fIsOkCustId = True
    vSql  = "SELECT Cust_Id FROM Cust WHERE SUBSTRING(Cust_Id, 5, 99)  = '" & Mid(vCust_Id, 5) & "'"
    sOpenDb4 
    Set oRs4  = oDb4.Execute(vSql)
    Do While Not oRs4.Eof 
      fIsOkCustId = False
      oRs4.MoveNext
    Loop
    Set oRs4  = Nothing
    sCloseDb4  
  End Function


  Function fIsCorporate
    '...for ecommerce (Ecom2Start.asp) determine if Corporate or Channel
    fIsCorporate = False
    sOpenDb2 
    vSql  = "SELECT Cust_Level FROM Cust WHERE Cust_Id  = '" & svCustId & "'"
    Set oRs2  = oDb2.Execute(vSql)
    If Not oRs2.Eof Then
      If oRs2("Cust_Level") = 4 Then 
        fIsCorporate = True
      End If
    End If
    Set oRs2  = Nothing
    sCloseDb2
  End Function

  '...is current user in a G2 site
  Function fIsGroup2
    fIsGroup2 = False
    sOpenDb2 
    vSql  = "SELECT Cust.Cust_AcctId FROM Cust INNER JOIN Ecom ON Cust.Cust_AcctId = Ecom.Ecom_NewAcctId WHERE (Ecom.Ecom_Media = 'Group2') AND (Cust.Cust_AcctId = '" & svCustAcctId & "')"
    Set oRs2  = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fIsGroup2 = True
    Set oRs2  = Nothing
    sCloseDb2
  End Function


  '...is specified site a G2 site (almost same as above)
  Function fCustG2Ok (vCustId)
    fCustG2Ok = False
    vSql = "SELECT Cust.Cust_Id FROM Cust INNER JOIN Ecom ON Cust.Cust_AcctId = Ecom.Ecom_NewAcctId AND Ecom.Ecom_Media = 'Group2' WHERE (Cust.Cust_Id = '" & vCustId & "')"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fCustG2Ok = True
    Set oRs = Nothing
    sCloseDb    
  End Function


  '...is current user in a G2 site
  Function fIsV8
    fIsV8 = False
    sOpenDb2 
    vSql  = "SELECT Cust.Cust_AcctId FROM Cust WHERE (Cust.Cust_AcctId = '" & svCustAcctId & "') AND (Cust.Cust_ChannelV8 = 1) "
    Set oRs2  = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fIsV8 = True
    Set oRs2  = Nothing
    sCloseDb2
  End Function


  '...is current user from NOP
  Function fIsNop
    fIsNop = False
    sOpenDb2 
    vSql  = "SELECT Cust.Cust_AcctId FROM Cust WHERE (Cust.Cust_AcctId = '" & svCustAcctId & "') AND (Cust.Cust_ChannelNop = 1) "
    Set oRs2  = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fIsNop = True
    Set oRs2  = Nothing
    sCloseDb2
  End Function

  
  '...is current user from a parent account
  Function fIsParent
    fIsParent = False
    vSql = "SELECT Cust_Id FROM Cust WHERE (Cust_Id = '" & svCustId & "' AND Cust_Level = 2 AND LEN(Cust_ParentId) = 0)"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fIsParent = True
    Set oRs = Nothing
    sCloseDb    
  End Function



  Sub sCloneCust (vCustId, vNewAcctId, vMaxUsers, vFacilitator, vPrograms, vExpires)
    '...clone the current customer record into the next highest available acctid
    '   note: extract just the ordered program strings fCustNewProgram (vCustId, vPrograms)
    '   note: auto enroll is always turned off, needs to be turned on by vu staff if needed

    '...change to ANSI_WARNINGS ON from OFF Nov 19, 2015 for computed columns
    vSql  = "SET ANSI_WARNINGS ON "_
          & "INSERT INTO Cust ("_
          &   "Cust_Id, "_
          &   "Cust_AcctId, "_
          &   "Cust_ParentId, "_ 
          &   "Cust_Title, "_
          &   "Cust_Lang, "_
          &   "Cust_Agent, "_
          &   "Cust_MaxSponsor, "_
          &   "Cust_IssueIds, "_
          &   "Cust_ResetStatus, "_
          &   "Cust_IssueIdsTemplate, "_
          &   "Cust_IssueIdsMemo, "_
          &   "Cust_ActivateIds, "_
          &   "Cust_IdsSize, "_
          &   "Cust_FreeHours, "_
          &   "Cust_FreeDays, "_
          &   "Cust_Auto, "_
          &   "Cust_Groups, "_
          &   "Cust_Programs, "_
          &   "Cust_ContentOnline, "_
          &   "Cust_ContentGroup, "_
          &   "Cust_ContentProds, "_
          &   "Cust_ContentCDs, "_
          &   "Cust_Active, "_
          &   "Cust_Desc, "_
          &   "Cust_Email, "_
          &   "Cust_Expires, "_
          &   "Cust_Cluster, "_
          &   "Cust_Level, "_
          &   "Cust_ContentLaunch, "_
          &   "Cust_Survey, "_
          &   "Cust_NoCert, "_
          &   "Cust_CustomCert, "_
          &   "Cust_CertLogoVubiz, "_
          &   "Cust_CertLogoCust, "_
          &   "Cust_CertEmailAlert, "_
          &   "Cust_AssessmentAttempts, "_
          &   "Cust_AssessmentScore, "_
          &   "Cust_AssessmentCert, "_
          &   "Cust_Tab1, "_
          &   "Cust_Tab2, "_
          &   "Cust_Tab3, "_
          &   "Cust_Tab4, "_
          &   "Cust_Tab5, "_
          &   "Cust_Tab6, "_
          &   "Cust_Tab7, "_
          &   "Cust_Tab4Type, "_
          &   "Cust_Tab1Name, "_
          &   "Cust_Tab2Name, "_
          &   "Cust_Tab3Name, "_
          &   "Cust_Tab4Name, "_
          &   "Cust_Tab5Name, "_
          &   "Cust_Tab6Name, "_
          &   "Cust_Tab7Name, "_
          &   "Cust_InfoEditProfile, "_
          &   "Cust_EcomCurrency, "_
          &   "Cust_EcomCorpRate, "_
          &   "Cust_EcomCorpDuration, "_
          &   "Cust_EcomCorpProgram, "_
          &   "Cust_EcomGroupLicense, "_
          &   "Cust_EcomGroupSeat, "_
          &   "Cust_EcomSplit, "_
          &   "Cust_EcomDiscOptions, "_
          &   "Cust_EcomDisc, "_
          &   "Cust_EcomDiscSplitCust, "_
          &   "Cust_EcomDiscSplitVubz, "_
          &   "Cust_EcomDiscSplitOwnr, "_
          &   "Cust_EcomDiscMinUS, "_
          &   "Cust_EcomDiscMinCA, "_
          &   "Cust_EcomDiscMinQty, "_
          &   "Cust_EcomDiscLimit, "_
          &   "Cust_EcomDiscOriginal, "_
          &   "Cust_EcomDiscPrograms, "_
          &   "Cust_EcomRepurPrograms, "_
          &   "Cust_EcomRepurDisc, "_
          &   "Cust_EcomRepurPeriod, "_
          &   "Cust_EcomConfirmation, "_
          &   "Cust_EcomEmailAddress, "_
          &   "Cust_EcomEmailBody, "_
          &   "Cust_EcomG2alert, "_
          &   "Cust_EcomGroup2Rates, "_
          &   "Cust_CorpAlert, "_
          &   "Cust_MyWorldLaunch, "_
          &   "Cust_MaxUsers, "_
          &   "Cust_Auth, "_
          &   "Cust_Pwd, "_
          &   "Cust_Resources, "_
          &   "Cust_ResourcesMaxSponsor, "_
          &   "Cust_VuNews, "_
          &   "Cust_Scheduler, "_
          &   "Cust_SeedLogs, "_
          &   "Cust_InsertLearners, "_
          &   "Cust_UpdateLearners, "_
          &   "Cust_DeleteLearners, "_
          &   "Cust_ResetLearners, "_

          &   "Cust_Banner, "_
          &   "Cust_Url, "_
          &   "Cust_StartUrl, "_
          &   "Cust_ReturnUrl, "_


          &   "Cust_Completion "_
          & ") "_
          
          & "(SELECT "_

          &   "'" & Left(vCustId,4) & vNewAcctId & "' AS Cust_Id, "_
          &   "'" & vNewAcctId & "' AS Cust_AcctId, "_
          &   "'" & Right(vEcom_CustId, 4) & "' AS Cust_ParentId, "_
          &   "Cust_Title, "_
          &   "Cust_Lang, "_
          &   "Cust_Agent, "_
          &   "Cust_MaxSponsor, "_
          &   "Cust_IssueIds, "_
          &   "Cust_ResetStatus, "_
          &   "Cust_IssueIdsTemplate, "_
          &   "Cust_IssueIdsMemo, "_
          &   "Cust_ActivateIds, "_
          &   "Cust_IdsSize, "_
          &   "Cust_FreeHours, "_
          &   "Cust_FreeDays, "_
          &   "0 AS Cust_Auto, "_
          &   "Cust_Groups, "_
          &   "'" & fCustNewProgram (vCustId, vPrograms) & "' AS Cust_Programs, "_
          &   "Cust_ContentOnline, "_
          &   "Cust_ContentGroup, "_
          &   "Cust_ContentProds, "_
          &   "Cust_ContentCDs, "_
          &   "Cust_Active, "_
          &   "Cust_Desc, "_
          &   "Cust_Email, "_
          &   "'" & vExpires & "', "_
          &   "'C0001' AS Cust_Cluster, "_
          &   "Cust_Level, "_
          &   "Cust_ContentLaunch, "_
          &   "Cust_Survey, "_
          &   "Cust_NoCert, "_
          &   "Cust_CustomCert, "_
          &   "Cust_CertLogoVubiz, "_
          &   "Cust_CertLogoCust, "_
          &   "Cust_CertEmailAlert, "_
          &   "Cust_AssessmentAttempts, "_
          &   "Cust_AssessmentScore, "_
          &   "Cust_AssessmentCert, "_
          &   "Cust_Tab1, "_
          &   "Cust_Tab2, "_
          &   "Cust_Tab3, "_
          &   "Cust_Tab4, "_
          &   "Cust_Tab5, "_
          &   "Cust_Tab6, "_
          &   "Cust_Tab7, "_
          &   "Cust_Tab4Type, "_
          &   "Cust_Tab1Name, "_
          &   "Cust_Tab2Name, "_
          &   "Cust_Tab3Name, "_
          &   "Cust_Tab4Name, "_
          &   "Cust_Tab5Name, "_
          &   "Cust_Tab6Name, "_
          &   "Cust_Tab7Name, "_
          &   "Cust_InfoEditProfile, "_
          &   "Cust_EcomCurrency, "_
          &   "Cust_EcomCorpRate, "_
          &   "Cust_EcomCorpDuration, "_
          &   "Cust_EcomCorpProgram, "_
          &   "Cust_EcomGroupLicense, "_
          &   "Cust_EcomGroupSeat, "_
          &   "Cust_EcomSplit, "_
          &   "Cust_EcomDiscOptions, "_
          &   "Cust_EcomDisc, "_
          &   "Cust_EcomDiscSplitCust, "_
          &   "Cust_EcomDiscSplitVubz, "_
          &   "Cust_EcomDiscSplitOwnr, "_
          &   "Cust_EcomDiscMinUS, "_
          &   "Cust_EcomDiscMinCA, "_
          &   "Cust_EcomDiscMinQty, "_
          &   "Cust_EcomDiscLimit, "_
          &   "Cust_EcomDiscOriginal, "_
          &   "Cust_EcomDiscPrograms, "_
          &   "Cust_EcomRepurPrograms, "_
          &   "Cust_EcomRepurDisc, "_
          &   "Cust_EcomRepurPeriod, "_
          &   "Cust_EcomConfirmation, "_
          &   "Cust_EcomEmailAddress, "_
          &   "Cust_EcomEmailBody, "_
          &   "0, "_
          &   "Cust_EcomGroup2Rates, "_
          &   "Cust_CorpAlert, "_
          &   "Cust_MyWorldLaunch, "_
          &   "" & vMaxUsers & " AS Cust_Maxusers, "_
          &   "Cust_Auth, "_
          &   "Cust_Pwd, "_
          &   "Cust_Resources, "_
          &   "Cust_ResourcesMaxSponsor, "_
          &   "Cust_VuNews, "_
          &   "Cust_Scheduler, "_
          &   "Cust_SeedLogs, "_
          &   "Cust_InsertLearners, "_
          &   "Cust_UpdateLearners, "_
          &   "Cust_DeleteLearners, "_
          &   "Cust_ResetLearners, "_

          &   "Cust_Banner, "_
          &   "Cust_Url, "_
          &   "Cust_StartUrl, "_
          &   "Cust_ReturnUrl, "_

          &   "Cust_Completion "_

          & "FROM Cust WHERE Cust_Id  = '" & vCustId & "')"

'   sDebug
    sOpenDb2 
    oDb2.Execute(vSql)
    sCloseDb2    
    
  End Sub


  '...Extract Matching Programs (for group sales)
  '...Used in sCloneCust above
  '...Disabled Feb 5, 2018 as no longer used - just for previous build users
  Function fCustNewProgram (vCustId, vPrograms)
    fCustNewProgram = ""
  End Function

'    Dim aProgs, aProg, i
'    '...get the program string from the customer account
'    vSql  = "SELECT Cust_Programs FROM Cust WHERE Cust_Id = '" & vCustId & "'"
'    sOpenDb2    
'    Set oRs2  = oDb2.Execute(vSql)
'    If Not oRs2.Eof THen
'      vCust_Programs  = oRs2("Cust_Programs")
'      Set oRs2  = Nothing
'    End If
'    sCloseDb2

'    aProgs  = Split(Trim(vCust_Programs), " ")
'    For i  = 0 to uBound(aProgs)
'      If Instr(vPrograms, Left(aProgs(i), 7)) > 0  Then
'        aProg  = Split(aProgs(i), "~")
'        aProg(1)  = 0   '...no ecommerce allowed
'        aProg(2)  = 0   '...no ecommerce allowed
'        aProg(4)  = 365 '...allow 365 days
'        fCustNewProgram  = fCustNewProgram & Join(aProg, "~") & " "
'      End If
'    Next
'   '...remove trailing blank 
'    fCustNewProgram  = Trim(fCustNewProgram)
'  End Function
 

  '...Get Cust Expires
  Function fCustExpires
    fCustExpires = ""
    vSql  = "SELECT Cust_Expires FROM Cust WHERE Cust_Id = '" & svCustId & "'"
    sOpenDb2    
    Set oRs2  = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fCustExpires = oRs2("Cust_Expires")
    Set oRs2  = Nothing
    sCloseDb2    
  End Function


  '...Get G1 Cust Expires
  Function fCustG1Expires (vCustId)
    fCustG1Expires = ""
    vSql  = "SELECT Cust_Expires FROM Cust WHERE Cust_Id = '" & vCustId & "'"
    sOpenDb2    
    Set oRs2  = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fCustG1Expires = oRs2("Cust_Expires")
    Set oRs2  = Nothing
    sCloseDb2    
  End Function


  Sub sUpdateCustExpires (vCustAcctId, vCustExpires)
    vSql  = "UPDATE Cust SET Cust_Expires = '" & fFormatSqlDate(vCustExpires) & "' WHERE Cust_AcctId = '" & vCustAcctId & "'"
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub

  
  Function fCustParentG2alertOk (vCustAcctId)
    fCustParentG2alertOk = False
    vSql  = "SELECT Cust_1.Cust_EcomG2Alert AS [ParentAlert] FROM dbo.Cust INNER JOIN dbo.Cust AS Cust_1 ON dbo.Cust.Cust_ParentId = Cust_1.Cust_AcctId WHERE Cust.Cust_AcctId = '" & vCustAcctId & "'"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      fCustParentG2alertOk = oRs2("ParentAlert")
    End If
    Set oRs2 = Nothing
    sCloseDb2      
  End Function


  Sub sUpdateCustG2alert (vCustAcctId, vCustG2alert)
    vSql  = "UPDATE Cust SET Cust_EcomG2alert = " & vCustG2alert & " WHERE Cust_AcctId = '" & vCustAcctId & "'"
'   sDebug
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub
  

  '...Is Cust Ok
  Function fCustOk (vCustId)
    fCustOk = False
    vSql = "SELECT Cust_Id FROM Cust WHERE Cust_Id= '" & vCustId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fCustOk = True
    Set oRs = Nothing
    sCloseDb    
  End Function


  '...Get Acct Id from Cust Id
  Function fCustAcctId (vCustId)
    Dim vSql
    fCustAcctId = ""
    vSql = "SELECT Cust_AcctId FROM Cust WHERE Cust_Id = '" & vCustId & "'"
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vSql)
    If Not oRs4.Eof Then fCustAcctId = oRs4("Cust_AcctId")
    Set oRs4 = Nothing
    sCloseDb4    
  End Function


  '...this grabs the next AcctId below and prefaces it with CUST
  Function fNextCustId
    fNextCustId = Left(svCustId, 4) & fNextAcctId
  End Function




 '...use for testing only then kill this and use/rename _new (below)
  Function fNextAcctId_prev ()
    fNextAcctId = ""
'   vSql = "SELECT TOP 1 RIGHT(A.Cust_Id, 4) + 1 AS Next "_
'        & "FROM Cust AS A LEFT OUTER JOIN Cust AS B ON RIGHT(A.Cust_Id, 4) + 1 = RIGHT(B.Cust_Id, 4) "_
'        & "WHERE (RIGHT(B.Cust_Id, 4) IS NULL) "_
'        & "ORDER BY RIGHT(A.Cust_Id, 4) "
    vSql = "SELECT TOP 1 CAST(RIGHT(a.Cust_Id, 4) AS INT) + 1 AS Next "_
         & "FROM Cust AS a LEFT OUTER JOIN Cust AS b ON CAST(RIGHT(a.Cust_Id, 4) AS INT) + 1 = CAST(RIGHT(b.Cust_Id, 4) AS INT) "_
         & "WHERE (b.Cust_Id IS NULL) "_
         & "ORDER BY Next"
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vSql)
    If Not oRs4.Eof Then 
      fNextAcctId = Right("0000" & oRs4("Next"), 4)
    End If
    Set oRs4 = Nothing
    sCloseDb4   
  End Function  



  '...get next available acctId from existing customer ids (this allows for negative numbers but NOT alpha
  Function fNextAcctId_temp ()

    fNextAcctId = ""

'   vSql = "SELECT TOP 1 RIGHT(A.Cust_Id, 4) + 1 AS Next "_
'        & "FROM Cust AS A LEFT OUTER JOIN Cust AS B ON RIGHT(A.Cust_Id, 4) + 1 = RIGHT(B.Cust_Id, 4) "_
'        & "WHERE (RIGHT(B.Cust_Id, 4) IS NULL) "_
'        & "ORDER BY RIGHT(A.Cust_Id, 4) "

'   vSql = "SELECT TOP 1 CAST(RIGHT(a.Cust_Id, 4) AS INT) + 1 AS Next "_
'        & "FROM Cust AS a LEFT OUTER JOIN Cust AS b ON CAST(RIGHT(a.Cust_Id, 4) AS INT) + 1 = CAST(RIGHT(b.Cust_Id, 4) AS INT) "_
'        & "WHERE (b.Cust_Id IS NULL) "_
'        & "ORDER BY Next"

    ' updated May 25, 2016 to ensure we don't go above 9999 - also good with negative numbers that should start with -999 
    vSql = "SELECT        TOP 1	CAST(RIGHT(c1.Cust_Id, 4) AS INT) + 1 AS nextId "_
         & "FROM					Cust AS c1 LEFT OUTER JOIN Cust AS c2 ON CAST(RIGHT(c1.Cust_Id, 4) AS INT) + 1 = CAST(RIGHT(c2.Cust_Id, 4) AS INT) "_
         & "WHERE					(c2.Cust_Id IS NULL) AND CAST(RIGHT(c1.Cust_Id, 4) AS INT) < 9999 AND CAST(RIGHT(c1.Cust_Id, 4) AS INT) <> -1 "_
         & "ORDER	BY      nextId"

    Dim nextNo, nextId
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vSql)
    If Not oRs4.Eof Then 
      nextId = oRs4("nextId")
      nextNo = Int(nextId)
      If (nextNo < 0) Then
        fNextAcctId = "-" & RIGHT("000" & MID(nextId, 2), 3) '...don't include the minus sign
      Else
        fNextAcctId = RIGHT("0000" & nextId, 4)
      End If
    End If
    Set oRs4 = Nothing
    sCloseDb4   
  End Function  



  '...get next available acctId from existing customer ids (this is configured to only return alpha)
  Function fNextAcctId()
    Dim oRs
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp6nextAlpha"
      .Parameters.Append .CreateParameter("@nextId", adVarChar, adParamOutput, 4)
    End With
    oCmdApp.Execute()
    fNextAcctId = oCmdApp.Parameters("@nextId").Value
    Set oCmdApp = Nothing
    sCloseDbApp
  End Function


%>