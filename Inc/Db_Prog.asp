<%
  '...the vProg_US_Memo, CA_ and _Duration_Memo are only memo fields used to send partners who want to do their own Ecommerce

  Dim vProg_Id, vProg_No, vProg_Mods, vProg_Scos, vProg_Title1, vProg_Title2, vProg_Promo, vProg_US_Memo, vProg_CA_Memo, vProg_Length, vProg_Duration_Memo, vProg_Desc, vProg_GrouId, vProg_Memo, vProg_Bookmark, vProg_CompletedButton, vProg_Test, vProg_LogTestResults, vProg_ResetStatus, vProg_Exam, vProg_Retired 
	Dim vProg_Owner, vProg_EcomSplitOwner1, vProg_EcomSplitOwner2, vProg_EcomGroupLicense, vProg_EcomGroupSeat, vProg_TaxExempt, vProg_Discounts
	Dim vProg_Assessment, vProg_AssessmentAttempts, vProg_AssessmentCert, vProg_AssessmentIds, vProg_AssessmentScore
	Dim vProg_Cert, vProg_CertTimeSpent, vProg_CertTestScore, vProg_CertTestAttempts, vProg_CustomCert, vProg_CertSimple
  Dim vProg_Nasba_Cpe '...added Sep 11th for certificates

  Dim vProg_Title '...not on db / this is the title that is either the general title (Title1) or the Variation assiged to a particular set of customer ids

  '...note these values are retrieved from vCust_Programs string, not the prog table
  Dim vProg_US, vProg_CA, vProg_MaxHours, vProg_Duration 

  Dim vProg_Eof, vProg_Ok


  '...Get Prog Record
  Sub sGetProg (vProgId)
    vProg_Eof = True
    vSql = "SELECT * FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadProg
      vProg_Eof = False
    End If
    Set oRsBase = Nothing
    sCloseDbBase
  End Sub


  Sub sGetProgByNo (vProgNo)
    vProg_Eof = True
    vSql = "SELECT * FROM Prog WHERE Prog_No= " & vProgNo
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadProg
      vProg_Eof = False
    End If
    Set oRsBase = Nothing
    sCloseDbBase
  End Sub


  '...Get Prog Title
  Function fProgTitle (vProgId)
    Dim oRsBase
    fProgTitle = ""
    vSql = "SELECT Prog_Title1, Prog_Title2 FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      vProg_Title1 = oRsBase("Prog_Title1")
      vProg_Title2 = oRsBase("Prog_Title2")
      sProgTitle
      fProgTitle = vProg_Title
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Get Clean Prog Title - strip off any trailing HTML
  Function fProgTitleClean (vProgId)
  	Dim i
    fProgTitleClean = fProgTitle (vProgId)
    i = Instr(fProgTitleClean, "<")
    If i > 0 Then fProgTitleClean = Left(fProgTitleClean, i-1)
  End Function



  '...Get Prog Mods
  Function fProgMods (vProgId)
    Dim oRsBase
    fProgMods = ""
    vSql = "SELECT Prog_Mods FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      fProgMods = Trim(oRsBase("Prog_Mods"))
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Get Prog Group
  Function fProgGroup (vProgId)
    Dim oRs
    fProgGroup = ""
    vSql = "SELECT Prog_GrouId FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRs = oDbBase.Execute(vSql)
    If Not oRs.Eof Then 
      fProgGroup = oRs("Prog_GrouId")
    End If
    Set oRs = Nothing
    sCloseDbBase    
  End Function

 

  '...Get Prog Recordset
  Sub sGetProg_Rs
    vSql = "SELECT * FROM Prog "
    sOpenDbBase    
    Set oRs = oDbBase.Execute(vSql)
  End Sub


  Sub sReadProg
    vProg_Id                 = oRsBase("Prog_Id")
    vProg_No                 = oRsBase("Prog_No")
    vProg_Mods               = oRsBase("Prog_Mods")
    vProg_Scos               = oRsBase("Prog_Scos")
    vProg_Title1             = oRsBase("Prog_Title1")
    vProg_Title2             = oRsBase("Prog_Title2")
    vProg_Promo              = oRsBase("Prog_Promo")
    vProg_US_Memo            = oRsBase("Prog_US_Memo")
    vProg_CA_Memo            = oRsBase("Prog_CA_Memo")
    vProg_Length             = oRsBase("Prog_Length")  : If vProg_Length = 0 Then vProg_Length = ""
    vProg_Duration_Memo      = oRsBase("Prog_Duration_Memo")
    vProg_Memo               = oRsBase("Prog_Memo") 
    vProg_Desc               = oRsBase("Prog_Desc")    : if Len(vProg_Desc) = 0 Then vProg_Desc = ""
    vProg_GrouId             = oRsBase("Prog_GrouId")
    vProg_Bookmark           = oRsBase("Prog_Bookmark") 
    vProg_CompletedButton    = oRsBase("Prog_CompletedButton")    
    vProg_LogTestResults     = oRsBase("Prog_LogTestResults") 
    vProg_ResetStatus        = oRsBase("Prog_ResetStatus") 
    vProg_Exam               = oRsBase("Prog_Exam") 
    vProg_Retired            = oRsBase("Prog_Retired") 
    vProg_Assessment         = oRsBase("Prog_Assessment") 
    vProg_AssessmentAttempts = oRsBase("Prog_AssessmentAttempts")
    vProg_AssessmentCert     = oRsBase("Prog_AssessmentCert") 
    vProg_AssessmentIds      = oRsBase("Prog_AssessmentIds") 
    vProg_AssessmentScore    = oRsBase("Prog_AssessmentScore") 
    vProg_Test               = oRsBase("Prog_Test") 
    vProg_Owner              = oRsBase("Prog_Owner") 
    vProg_EcomSplitOwner1    = oRsBase("Prog_EcomSplitOwner1") 
    vProg_EcomSplitOwner2    = oRsBase("Prog_EcomSplitOwner2") 
    vProg_EcomGroupLicense   = oRsBase("Prog_EcomGroupLicense")
    vProg_EcomGroupSeat      = oRsBase("Prog_EcomGroupSeat")
    vProg_TaxExempt          = oRsBase("Prog_TaxExempt")
    vProg_Cert               = oRsBase("Prog_Cert")
    vProg_CertTimeSpent      = oRsBase("Prog_CertTimeSpent") 
    vProg_CertTestScore      = oRsBase("Prog_CertTestScore") 
    vProg_CertTestAttempts   = oRsBase("Prog_CertTestAttempts")    
    vProg_CertSimple		     = oRsBase("Prog_CertSimple")    
    vProg_CustomCert         = oRsBase("Prog_CustomCert")    
    vProg_Discounts          = oRsBase("Prog_Discounts")    
    vProg_Nasba_Cpe          = oRsBase("Prog_Nasba_Cpe")    

    sProgTitle
  End Sub


  Sub sReadProgEcom
    vProg_Id                 = oRs("Prog_Id")
    vProg_Title1             = oRs("Prog_Title1")
    vProg_Title2             = oRs("Prog_Title2")
    vProg_Owner              = oRs("Prog_Owner") 
    vProg_EcomSplitOwner1    = oRs("Prog_EcomSplitOwner1") 
    vProg_EcomSplitOwner2    = oRs("Prog_EcomSplitOwner2") 
    vProg_EcomGroupLicense   = oRs("Prog_EcomGroupLicense")
    vProg_EcomGroupSeat      = oRs("Prog_EcomGroupSeat")
    vProg_TaxExempt          = oRs("Prog_TaxExempt")
    vProg_Discounts          = oRs("Prog_Discounts")    
    sProgTitle
  End Sub


  '...This will return vProg_Title either from Title1 or Title2 (ie customer variant)
  Sub sProgTitle
    Dim i, j
    i = Instr(vProg_Title2, svCustId) 
    If i > 0 Then
      vProg_Title = Trim(Mid(vProg_Title2, i+8))
      j = Instr(vProg_Title, "~")
      If j > 0 Then 
        vProg_Title = Trim(Left(vProg_Title, j-1))
      End If
    Else
      vProg_Title = vProg_Title1
    End If
  End Sub


  Sub sExtractProg
    vProg_Id                 = Ucase(Request.Form("vProg_Id"))
    vProg_Title1             = fUnquote(Request.Form("vProg_Title1"))
    vProg_Title2             = fUnquote(Request.Form("vProg_Title2"))
    vProg_Promo              = fUnquote(Request.Form("vProg_Promo"))
    vProg_Mods               = Trim(Request.Form("vProg_Mods"))
    vProg_Scos               = Trim(Request.Form("vProg_Scos"))
    vProg_Desc               = fUnquote(Request.Form("vProg_Desc"))
    vProg_US_Memo            = fDefault(Request.Form("vProg_US_Memo"), 0)
    vProg_CA_Memo            = fDefault(Request.Form("vProg_CA_Memo"), 0)
    vProg_Duration_Memo      = fDefault(Request.Form("vProg_Duration_Memo"), 0)
    vProg_GrouId             = Request.Form("vProg_GrouId")
    vProg_Length             = fProgLength
    vProg_Memo               = fUnquote(Request.Form("vProg_Memo"))
    vProg_Bookmark           = fDefault(Request.Form("vProg_Bookmark"), "Y")
    vProg_CompletedButton    = fDefault(Request.Form("vProg_CompletedButton"), "N")    
    vProg_LogTestResults     = fDefault(Request.Form("vProg_LogTestResults"), "N")
    vProg_ResetStatus        = fDefault(Request.Form("vProg_ResetStatus"), 0)
    vProg_Exam               = Request.Form("vProg_Exam")
    vProg_Retired            = Request.Form("vProg_Retired")
    vProg_Assessment         = Request.Form("vProg_Assessment")
    vProg_AssessmentAttempts = Request.Form("vProg_AssessmentAttempts")
    vProg_AssessmentCert     = Request.Form("vProg_AssessmentCert")
    vProg_AssessmentIds      = Request.Form("vProg_AssessmentIds")
    vProg_AssessmentScore    = Request.Form("vProg_AssessmentScore")
    vProg_Test               = fDefault(Request.Form("vProg_Test"), "N")
    vProg_Owner              = Ucase(Request.Form("vProg_Owner"))
    vProg_EcomSplitOwner1    = fDefault(Request.Form("vProg_EcomSplitOwner1"), 0)
    vProg_EcomSplitOwner2    = fDefault(Request.Form("vProg_EcomSplitOwner2"), 0)
    vProg_EcomGroupLicense   = fDefault(Request.Form("vProg_EcomGroupLicense"), 0)
    vProg_EcomGroupSeat      = fDefault(Request.Form("vProg_EcomGroupSeat"), 0)
    vProg_TaxExempt          = fDefault(Request.Form("vProg_TaxExempt"), 0)
    vProg_Cert               = fDefault(Request.Form("vProg_Cert"), 0)
    vProg_CertTimeSpent      = fDefault(Request.Form("vProg_CertTimeSpent"), 0)
    vProg_CertTestScore      = fDefault(Request.Form("vProg_CertTestScore"), 0) 
    vProg_CertTestAttempts   = fDefault(Request.Form("vProg_CertTestAttempts"), 0)    
    vProg_CertSimple         = fDefault(Request.Form("vProg_CertSimple"), 0)    
    vProg_CustomCert         = fDefault(Request.Form("vProg_CustomCert"), 0)    
    vProg_Discounts          = fDefault(Request.Form("vProg_Discounts"), "Y")    
    vProg_Nasba_Cpe          = Request.Form("vProg_Nasba_Cpe")
  End Sub

  
  Sub sInsertProg
    vProg_Ok = False
    vSql = "INSERT INTO Prog "
    vSql = vSql & "(Prog_Id, Prog_Title1, Prog_Title2, Prog_Promo, Prog_Mods, Prog_Scos, Prog_Desc, Prog_US_Memo, Prog_CA_Memo, Prog_Length, Prog_Duration_Memo, Prog_GrouId, Prog_Memo, Prog_Bookmark, Prog_CompletedButton, Prog_LogTestResults, Prog_ResetStatus, Prog_Exam, Prog_Retired, Prog_Assessment, Prog_AssessmentAttempts, Prog_AssessmentCert, Prog_AssessmentIds, Prog_AssessmentScore, Prog_Test, Prog_Owner, Prog_EcomSplitOwner1, Prog_EcomSplitOwner2, Prog_EcomGroupLicense, Prog_EcomGroupSeat, Prog_TaxExempt, Prog_Cert, Prog_CertTimeSpent, Prog_CertTestScore, Prog_CertTestAttempts, Prog_CertSimple, Prog_CustomCert, Prog_Discounts, Prog_Nasba_Cpe)"
    vSql = vSql & " VALUES ('" & vProg_Id & "', '" & vProg_Title1 & "', '" & vProg_Title2 & "', '" & vProg_Promo & "', '" & vProg_Mods & "', '" & vProg_Scos & "', '" & vProg_Desc & "', " & vProg_US_Memo & ", " & vProg_CA_Memo & ", " & vProg_Length & ", " & vProg_Duration_Memo & ", '" & vProg_GrouId & "', '" & vProg_Memo & "', '" & vProg_Bookmark & "', '" & vProg_CompletedButton & "', '" & vProg_LogTestResults & "', " & vProg_ResetStatus & ", '" & vProg_Exam & "', " & fSqlBoolean(vProg_Retired) & ", '" & vProg_Assessment & "', " & vProg_AssessmentAttempts & ", '" & vProg_AssessmentCert & "', '" & vProg_AssessmentIds & "',  " & vProg_AssessmentScore & ", '" & vProg_Test & "', '" & vProg_Owner & "', " & vProg_EcomSplitOwner1 & ", " & vProg_EcomSplitOwner2 & ", " & vProg_EcomGroupLicense & ", " & vProg_EcomGroupSeat & ", " & vProg_TaxExempt & ", " & vProg_Cert & ", " & vProg_CertTimeSpent & ", " & vProg_CertTestScore & ", " & vProg_CertTestAttempts & ", " & vProg_CertSimple & ", " & vProg_CustomCert & ", '" & vProg_Discounts & "', '" & vProg_Nasba_Cpe & "')"
'   sDebug
    sOpenDbBase
    On Error Resume Next
    oDbBase.Execute(vSql)
    sCloseDbBase
    If Err <> 0 Then Exit Sub 
    vProg_Ok = True
'   sUpdateCustProgLength vProg_Id, vProg_Length
'   spProgModsUpdate vProg_Id (now handled by trigger)
  End Sub


  Sub sUpdateProg
    vSql = "UPDATE Prog SET"
    vSql = vSql & " Prog_Title1             = '" & vProg_Title1             & "', " 
    vSql = vSql & " Prog_Title2             = '" & vProg_Title2             & "', " 
    vSql = vSql & " Prog_Promo              = '" & vProg_Promo              & "', " 
    vSql = vSql & " Prog_Mods               = '" & vProg_Mods               & "', " 
    vSql = vSql & " Prog_Scos               = '" & vProg_Scos               & "', " 
    vSql = vSql & " Prog_Desc               = '" & vProg_Desc               & "', " 
    vSql = vSql & " Prog_US_Memo            =  " & vProg_US_Memo            & " , "
    vSql = vSql & " Prog_CA_Memo            =  " & vProg_CA_Memo            & " , "
    vSql = vSql & " Prog_Length             =  " & vProg_Length             & " , "
    vSql = vSql & " Prog_Duration_Memo      =  " & vProg_Duration_Memo      & " , "
    vSql = vSql & " Prog_GrouId             = '" & vProg_GrouId             & "', "
    vSql = vSql & " Prog_Memo               = '" & vProg_Memo               & "', "
    vSql = vSql & " Prog_Bookmark           = '" & vProg_Bookmark           & "', "
    vSql = vSql & " Prog_CompletedButton    = '" & vProg_CompletedButton    & "', "
    vSql = vSql & " Prog_LogTestResults     = '" & vProg_LogTestResults     & "', "
    vSql = vSql & " Prog_ResetStatus        =  " & vProg_ResetStatus        & " , "
    vSql = vSql & " Prog_Exam               = '" & vProg_Exam               & "', "
    vSql = vSql & " Prog_Retired            =  " & fSqlBoolean(vProg_Retired) & " , "
    vSql = vSql & " Prog_Assessment         = '" & vProg_Assessment         & "', "
    vSql = vSql & " Prog_AssessmentAttempts = '" & vProg_AssessmentAttempts & "', "
    vSql = vSql & " Prog_AssessmentCert     = '" & vProg_AssessmentCert     & "', "
    vSql = vSql & " Prog_AssessmentIds      = '" & vProg_AssessmentIds      & "', "
    vSql = vSql & " Prog_AssessmentScore    =  " & vProg_AssessmentScore    & " , "
    vSql = vSql & " Prog_Test               = '" & vProg_Test               & "', "
    vSql = vSql & " Prog_Owner              = '" & vProg_Owner              & "', "
    vSql = vSql & " Prog_EcomSplitOwner1    =  " & vProg_EcomSplitOwner1    & " , "
    vSql = vSql & " Prog_EcomSplitOwner2    =  " & vProg_EcomSplitOwner2    & " , "
    vSql = vSql & " Prog_EcomGroupLicense   =  " & vProg_EcomGroupLicense   & " , " 
    vSql = vSql & " Prog_EcomGroupSeat      =  " & vProg_EcomGroupSeat      & " , " 
    vSql = vSql & " Prog_TaxExempt          =  " & vProg_TaxExempt          & " , " 
    vSql = vSql & " Prog_Cert               =  " & vProg_Cert               & " , " 
    vSql = vSql & " Prog_CertTimeSpent      =  " & vProg_CertTimeSpent      & " , " 
    vSql = vSql & " Prog_CertTestScore      =  " & vProg_CertTestScore      & " , " 
    vSql = vSql & " Prog_CertTestAttempts   =  " & vProg_CertTestAttempts   & " , " 
    vSql = vSql & " Prog_CertSimple         =  " & vProg_CertSimple         & " , " 
    vSql = vSql & " Prog_CustomCert         =  " & vProg_CustomCert         & " , " 
    vSql = vSql & " Prog_Discounts          = '" & vProg_Discounts          & "', " 
    vSql = vSql & " Prog_Nasba_Cpe          = '" & vProg_Nasba_Cpe          & "'  " 

    vSql = vSql & " WHERE Prog_Id           = '" & vProg_Id                 & "'  "

'   sDebug
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase

'   sUpdateCustProgLength vProg_Id, vProg_Length
'   now handled by Prog trigger

  End Sub
  

  Sub sDeleteProg
    vSql = "DELETE FROM Prog WHERE Prog_Id = '" & vProg_Id & "'"
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase
'   spProgModsUpdate vProg_Id (now handled by trigger)
  End Sub

  '...get all Program module lengths
  Function fProgLength
    fProgLength = 0
    Dim aMods
    aMods = Split(Trim(vProg_Mods), " ")
    For i = 0 to uBound(aMods)
      fProgLength = fProgLength + fProgModsLength (aMods(i))  
    Next
  End Function

 
  '...Get Mods Length
  Function fProgModsLength (vModsId)
    fProgModsLength = 0
    vSql = "SELECT Mods_Length FROM Mods WHERE Mods_Id= '" & vModsId & "' AND Mods_Active = 1"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fProgModsLength = oRsBase("Mods_Length")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Get Assessment Attempts (Used in RTE_ModsStat)
  Function fProgAttempts (vProgId)
    fProgAttempts = 0
    vSql = "SELECT Prog_AssessmentAttempts FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fProgAttempts = oRsBase("Prog_AssessmentAttempts")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function

 
  '...get all Prog
  Function fProgOptions
    Dim oRs
    fProgOptions = ""
    sOpenDbBase
    vSql = "Select * FROM Prog"
    Set oRs = oDbBase.Execute(vSql)    
    Do While Not oRs.EOF 
      fProgOptions = fProgOptions & "<option>" & oRs("Prog_Id") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop      
    sCloseDbBase           
  End Function    


  '...Tax Exempt?
  Function fProgTaxExempt (vProgId)
    vProg_Id = fIf(Len(vProgId) > 7, Right(vProgId, 7), vProgId)
    vSql = "SELECT Prog_TaxExempt FROM Prog WHERE Prog_Id= '" & vProg_Id & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      fProgTaxExempt = oRsBase("Prog_TaxExempt")
    Else
      fProgTaxExempt = False
    End If      
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Discounts Ok?
  Function fProgDiscountsOk (vProgId)
    fProgDiscountsOk = False
    vProg_Id = fIf(Len(vProgId) > 7, Right(vProgId, 7), vProgId)
    vSql = "SELECT Prog_Discounts FROM Prog WHERE Prog_Id= '" & vProg_Id & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      If oRsBase("Prog_Discounts") = "Y" Then fProgDiscountsOk = True
    End If      
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...clone a program
  Function fCloneProg (vProgId, vCloneId)
    Dim vOk
    vOk = False
    fCloneProg = False

    '...ensure vProgId exists and vCloneId is valid and does NOT exist
    If fProgOk(vProgId) Then
      If Len(vCloneId) = 7 Then
        If IsNumeric(Mid(vCloneId, 2, 4)) Then
          If Left(vCloneId, 1) = "P" Then
            If Instr("EN FR ES PT", Right(vCloneId, 2)) > 0 Then
              If Not fProgOk(vCloneId) Then
                vOk = True
              End If
            End If
          End If
        End If
      End If
    End If
      
    If vOk Then
      vSql = "SET ANSI_WARNINGS ON " _
           & "INSERT INTO Prog " _
           & "(Prog_Id, Prog_Title1, Prog_Title2, Prog_Promo, Prog_Mods, Prog_Scos, Prog_Desc, Prog_US_Memo, Prog_CA_Memo, Prog_Length, Prog_Duration_Memo, Prog_GrouId, Prog_Memo, Prog_Bookmark, Prog_CompletedButton, Prog_LogTestResults, Prog_ResetStatus, Prog_Exam, Prog_Retired, Prog_Assessment, Prog_AssessmentAttempts, Prog_AssessmentCert, Prog_AssessmentIds, Prog_AssessmentScore, Prog_Test, Prog_Owner, Prog_EcomSplitOwner1, Prog_EcomSplitOwner2, Prog_EcomGroupLicense, Prog_EcomGroupSeat, Prog_TaxExempt, Prog_Cert, Prog_CertTimeSpent, Prog_CertTestScore, Prog_CertTestAttempts, Prog_CertSimple, Prog_CustomCert, Prog_Discounts, Prog_Nasba_Cpe) " _
           & "(SELECT '" & vCloneId & "' AS Prog_Id, Prog_Title1, Prog_Title2, Prog_Promo, Prog_Mods, Prog_Scos, Prog_Desc, Prog_US_Memo, Prog_CA_Memo, Prog_Length, Prog_Duration_Memo, Prog_GrouId, Prog_Memo, Prog_Bookmark, Prog_CompletedButton, Prog_LogTestResults, Prog_ResetStatus, Prog_Exam, Prog_Retired, Prog_Assessment, Prog_AssessmentAttempts, Prog_AssessmentCert, Prog_AssessmentIds, Prog_AssessmentScore, Prog_Test, Prog_Owner, Prog_EcomSplitOwner1, Prog_EcomSplitOwner2, Prog_EcomGroupLicense, Prog_EcomGroupSeat, Prog_TaxExempt, Prog_Cert, Prog_CertTimeSpent, Prog_CertTestScore, Prog_CertTestAttempts, Prog_CertSimple, Prog_CustomCert, Prog_Discounts, Prog_Nasba_Cpe FROM Prog WHERE Prog_Id  = '" & vProgId & "')"
'     sDebug
      sOpenDbBase 
      oDbBase.Execute(vSql)
      sCloseDbBase    
      fCloneProg = True
    End If
  End Function


  '...Get ProgNo using the ProgId (used for RTE)
  Function fProgNoById (vProgId)
    fProgNoById = 0
    vSql = "SELECT Prog_No FROM Prog WHERE Prog_Id = '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fProgNoById = oRsBase("Prog_No")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Is Prog Ok
  Function fProgOk (vProgId)
    Dim oRsBase
    fProgOk = False
    vSql = "SELECT Prog_Title1 FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      fProgOk = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function

  '...Get Next Available Prog Id
  Function fProgNext ()
    Dim oRsBase
    fProgNext = ""
    vSql = "SELECT Prog_Title1 FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      fProgOk = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  Function fProgXX (vProgId)
    fProgXX = vProgId
    If Len(fProgXX) = 7 Then
      If Ucase(Right(fProgXX, 2)) = "XX" Then
        fProgXX = Left(fProgXX, 5) & svLang
      End If
    End If
  End Function
  
%>