<%
  Dim vMods_Id, vMods_No, vMods_Format, vMods_Title, vMods_Active, vMods_VuCert, vMods_Desc, vMods_Outline, vMods_Goals, vMods_Length, vMods_Url, vMods_Script, vMods_AssessmentUrl, vMods_AssessmentScript, vMods_SkillSet, vMods_PreviewMax
  Dim vMods_AccessOk, vMods_AccessNo, vMods_Type, vMods_Player, vMods_Width, vMods_Height, vMods_FullScreen, vMods_Fluid, vMods_Memo, vMods_ParentId, vMods_Competency, vMods_Completion
  Dim vMods_Reviewed, vMods_FeaAcc, vMods_FeaAud, vMods_FeaMob, vMods_FeaHyb, vMods_FeaVid 

  Dim vMods_Eof, vMods_Langs  '...this is created whenever a module is pulled up to show what other languages are available with this MOD, ie "EN ES FR PT"
  Dim vMods_Features          '...this is created whenever a module is pulled up to show the features available
  
  '____ Mods ________________________________________________________________________

  '...Get Mods Record by ID
  Sub sGetMods (vModId)
    vSql = "SELECT * FROM Mods WHERE Mods_Id= '" & vModId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadMods
      vMods_Eof = False
    Else
      vMods_Eof = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Sub


  '...Get Mods Record by No
  Sub sGetModsByNo (vModsNo)
    vSql = "SELECT * FROM Mods WHERE Mods_No= " & vModsNo
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      sReadMods
      vMods_Eof = False
    Else
      vMods_Eof = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Sub


  '...Get Mods Title (if Preview at the beginning then use last 6 characters)
  Function fModsTitle (vModId)
    fModsTitle = ""
'   vSql = "SELECT Mods_Title FROM Mods WHERE Mods_Id= '" & fIf(Len(vModId)>6, Right(vModId, 6), vModId) & "'"
    vSql = "SELECT Mods_Title FROM Mods WHERE Mods_Id= '" & fIf(Instr("0123456789", Left(vModId, 1)) = 0, Mid(vModId, 2), vModId) & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fModsTitle = oRsBase("Mods_Title")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Get Mods Length
  Function fModsLength (vModId)
    fModsLength = 0
    vSql = "SELECT Mods_Length FROM Mods WHERE Mods_Id= '" & vModId & "' And Mods_Active = 1"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fModsLength = oRsBase("Mods_Length")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  Sub sReadMods
    vMods_Id               = Ucase(oRsBase("Mods_Id"))
    vMods_No               = oRsBase("Mods_No")
    vMods_AccessOk         = oRsBase("Mods_AccessOk")
    vMods_AccessNo         = oRsBase("Mods_AccessNo")
    vMods_Format           = oRsBase("Mods_Format")
    vMods_Title            = oRsBase("Mods_Title")
    vMods_ParentId         = Ucase(oRsBase("Mods_ParentId"))
    vMods_Active           = oRsBase("Mods_Active")
    vMods_VuCert           = oRsBase("Mods_VuCert")
    vMods_Desc             = oRsBase("Mods_Desc")
    vMods_Outline          = oRsBase("Mods_Outline")
    vMods_Goals            = oRsBase("Mods_Goals")
    vMods_Length           = Csng(oRsBase("Mods_Length"))
    vMods_Type             = oRsBase("Mods_Type")
    vMods_Player           = Cint(oRsBase("Mods_Player"))
    vMods_Url              = oRsBase("Mods_Url")
    vMods_Script           = oRsBase("Mods_Script")
    vMods_AssessmentUrl    = oRsBase("Mods_AssessmentUrl")
    vMods_AssessmentScript = oRsBase("Mods_AssessmentScript")
    vMods_SkillSet         = oRsBase("Mods_SkillSet")
    vMods_Competency       = oRsBase("Mods_Competency")
    vMods_Completion       = oRsBase("Mods_Completion")
    vMods_Reviewed         = oRsBase("Mods_Reviewed")

    vMods_FeaAcc           = oRsBase("Mods_FeaAcc")
    vMods_FeaAud           = oRsBase("Mods_FeaAud")
    vMods_FeaMob           = oRsBase("Mods_FeaMob")
    vMods_FeaHyb           = oRsBase("Mods_FeaHyb")
    vMods_FeaVid           = oRsBase("Mods_FeaVid")

    vMods_PreviewMax       = oRsBase("Mods_PreviewMax")
    vMods_Memo             = oRsBase("Mods_Memo")
    vMods_Width            = oRsBase("Mods_Width")
    vMods_Height           = oRsBase("Mods_Height")
    vMods_FullScreen       = oRsBase("Mods_FullScreen")
    vMods_Fluid            = oRsBase("Mods_Fluid")

    sModsLangs (Left(vMods_Id, 4)) '...grab the languages available for this mod    
    sModsFeatures
   
  End Sub
  
  
  '...Get Mods Features to render on lists
  Function sModsFeatures
    vMods_Features = "&nbsp;"
    If vMods_FeaAcc Then vMods_Features = vMods_Features & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaAcc.png'>&nbsp;"
    If vMods_FeaAud Then vMods_Features = vMods_Features & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaAud.png'>&nbsp;"
    If vMods_FeaMob Then vMods_Features = vMods_Features & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaMob.png'>&nbsp;"
    If vMods_FeaHyb Then vMods_Features = vMods_Features & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaHyb.png'>&nbsp;"
    If vMods_FeaVid Then vMods_Features = vMods_Features & "<img align='absbottom' border='0' src='../Images/RTE/ModsFeaVid.png'>&nbsp;"   
  End Function  


  Sub sExtractMods
    vMods_Id               = Ucase(Request.Form("vMods_Id"))
    vMods_AccessOk         = Ucase(Request.Form("vMods_AccessOk"))
    vMods_AccessNo         = Ucase(Request.Form("vMods_AccessNo"))
    vMods_Format           = fDefault(Request.Form("vMods_Format"), 0)
    vMods_Title            = Request.Form("vMods_Title")
    vMods_ParentId         = Ucase(Request.Form("vMods_ParentId"))
    vMods_Active           = Request.Form("vMods_Active")
    vMods_VuCert           = Request.Form("vMods_VuCert")
    vMods_Desc             = Request.Form("vMods_Desc")
    vMods_Outline          = Request.Form("vMods_Outline")
    vMods_Goals            = Request.Form("vMods_Goals")
    vMods_Length           = Csng(fDefault(Request.Form("vMods_Length"), 1.0))
    vMods_Type             = Ucase(Request.Form("vMods_Type"))
    vMods_Player           = fDefault(Request.Form("vMods_Player"), 0)
    vMods_Url              = Lcase(Request.Form("vMods_Url"))
    vMods_Script           = Request.Form("vMods_Script")
    vMods_AssessmentUrl    = Request.Form("vMods_AssessmentUrl")
    vMods_AssessmentScript = Request.Form("vMods_AssessmentScript")
    vMods_SkillSet         = Ucase(Request.Form("vMods_SkillSet"))
    vMods_Competency       = Request.Form("vMods_Competency")
    vMods_Completion       = fDefault(Request.Form("vMods_Completion"), 1)
    vMods_Reviewed         = fDefault(Request.Form("vMods_Reviewed"), 1)

    vMods_FeaAcc           = fDefault(Request.Form("vMods_FeaAcc"), 0)
    vMods_FeaAud           = fDefault(Request.Form("vMods_FeaAud"), 0)
    vMods_FeaMob           = fDefault(Request.Form("vMods_FeaMob"), 0)
    vMods_FeaHyb           = fDefault(Request.Form("vMods_FeaHyb"), 0)
    vMods_FeaVid           = fDefault(Request.Form("vMods_FeaVid"), 0)

    vMods_PreviewMax       = fDefault(Request.Form("vMods_PreviewMax"), 0)
    vMods_Width            = fDefault(Request.Form("vMods_Width"), 0)
    vMods_Height           = fDefault(Request.Form("vMods_Height"), 0)
    vMods_FullScreen       = fDefault(Request.Form("vMods_FullScreen"), 0)
    vMods_Fluid            = fDefault(Request.Form("vMods_Fluid"), 0)
    vMods_Memo             = Request.Form("vMods_Memo")
  End Sub
  

  Sub spModsAlterById
    sOpenCmdBase
    With oCmdBase
      .CommandText = "spModsAlterById"     
      .Parameters.Append .CreateParameter("@Mods_Id",	              adVarChar,  adParamInput,    7, vMods_Id)
      .Parameters.Append .CreateParameter("@Mods_No",	              adInteger,  adParamInput,     , vMods_No)
'     .Parameters.Append .CreateParameter("@Mods_NextId",           adVarChar,  adParamInput,    6, vMods_NextId)
'     .Parameters.Append .CreateParameter("@Mods_AccessOk",	        adVarChar,  adParamInput, 2000, vMods_AccessOk)
'     .Parameters.Append .CreateParameter("@Mods_AccessNo",	        adVarChar,  adParamInput, 2000, vMods_AccessNo)
'     .Parameters.Append .CreateParameter("@Mods_ParentId",	        adVarChar,  adParamInput,    6, vMods_ParentId)
      .Parameters.Append .CreateParameter("@Mods_Format",	          adTinyInt,  adParamInput,     , vMods_Format)
      .Parameters.Append .CreateParameter("@Mods_Title",	          adVarChar,  adParamInput, 2000, vMods_Title)
      .Parameters.Append .CreateParameter("@Mods_Active",	          adBoolean,  adParamInput,     , vMods_Active)
      .Parameters.Append .CreateParameter("@Mods_Desc",	            adVarChar,  adParamInput, 8000, vMods_Desc)
      .Parameters.Append .CreateParameter("@Mods_Outline",	        adVarChar,  adParamInput, 8000, vMods_Outline)
      .Parameters.Append .CreateParameter("@Mods_Goals",	          adVarChar,  adParamInput, 8000, vMods_Goals)
      .Parameters.Append .CreateParameter("@Mods_Length",	          adSingle,   adParamInput,     , vMods_Length)
      .Parameters.Append .CreateParameter("@Mods_SkillSet",	        adVarChar,  adParamInput, 8000, vMods_SkillSet)
      .Parameters.Append .CreateParameter("@Mods_Competency",	      adVarChar,  adParamInput,   50, vMods_Competency)
      .Parameters.Append .CreateParameter("@Mods_Completion",	      adBoolean,  adParamInput,     , vMods_Completion)

      .Parameters.Append .CreateParameter("@Mods_Reviewed",	        adBoolean,  adParamInput,     , vMods_Reviewed)

      .Parameters.Append .CreateParameter("@Mods_FeaAcc",	          adBoolean,  adParamInput,     , vMods_FeaAcc)
      .Parameters.Append .CreateParameter("@Mods_FeaAud",	          adBoolean,  adParamInput,     , vMods_FeaAud)
      .Parameters.Append .CreateParameter("@Mods_FeaMob",	          adBoolean,  adParamInput,     , vMods_FeaMob)
      .Parameters.Append .CreateParameter("@Mods_FeaHyb",	          adBoolean,  adParamInput,     , vMods_FeaHyb)
      .Parameters.Append .CreateParameter("@Mods_FeaVid",	          adBoolean,  adParamInput,     , vMods_FeaVid)

      .Parameters.Append .CreateParameter("@Mods_Type",	            adVarChar,  adParamInput,    2, vMods_Type)
      .Parameters.Append .CreateParameter("@Mods_Player",           adSmallInt, adParamInput,     , vMods_Player)
      .Parameters.Append .CreateParameter("@Mods_Url",	            adVarChar,  adParamInput, 1000, vMods_Url)
      .Parameters.Append .CreateParameter("@Mods_Script",	          adVarChar,  adParamInput,  125, vMods_Script)
      .Parameters.Append .CreateParameter("@Mods_AssessmentUrl",    adVarChar,  adParamInput, 1000, vMods_AssessmentUrl)
      .Parameters.Append .CreateParameter("@Mods_AssessmentScript", adVarChar,  adParamInput, 1000, vMods_AssessmentScript)
'     .Parameters.Append .CreateParameter("@Mods_CCF",	            adVarChar,  adParamInput,   50, vMods_CCF)
      .Parameters.Append .CreateParameter("@Mods_PreviewMax",	      adInteger,  adParamInput,     , vMods_PreviewMax)
      .Parameters.Append .CreateParameter("@Mods_VuCert",	          adBoolean,  adParamInput,     , vMods_VuCert)
      .Parameters.Append .CreateParameter("@Mods_Width",	          adSmallInt, adParamInput,     , vMods_Width)
      .Parameters.Append .CreateParameter("@Mods_Height",	          adSmallInt, adParamInput,     , vMods_Height)
      .Parameters.Append .CreateParameter("@Mods_FullScreen",	      adBoolean,  adParamInput,     , vMods_FullScreen)
      .Parameters.Append .CreateParameter("@Mods_Fluid",	          adBoolean,  adParamInput,     , vMods_Fluid)
      .Parameters.Append .CreateParameter("@Mods_Memo",	            adVarChar,  adParamInput, 2000, vMods_Memo)  
    End With
    oCmdBase.Execute()
    Set oCmdBase = Nothing
    sCloseDbBase
  End Sub


  Sub sInsertMods
    vSql = "INSERT INTO Mods "
    vSql = vSql & "(Mods_Id, Mods_Format, Mods_Title, Mods_Active, Mods_AccessOk, Mods_AccessNo, Mods_ParentId, Mods_VuCert, Mods_Desc, Mods_Outline, Mods_Goals, Mods_Length, Mods_SkillSet, Mods_Competency, Mods_Completion, Mods_Reviewed, Mods_FeaAcc, Mods_FeaAud, Mods_FeaMob, Mods_FeaHyb, Mods_FeaVid, Mods_Type, Mods_Player, Mods_Url, Mods_PreviewMax, Mods_Script, Mods_Width, Mods_Height, Mods_FullScreen, Mods_Fluid, Mods_Memo)"
    vSql = vSql & " VALUES ('" & vMods_Id & "', '" & fUnquote(vMods_Title) & "', " & vMods_Active & ", '" & vMods_AccessOk & "', '" & vMods_AccessNo & "', '" & vMods_ParentID & "', " & vMods_VuCert & ", '" & fUnquote(vMods_Desc) & "', '" & fUnquote(vMods_Outline) & "', '" & fUnquote(vMods_Goals) & "', " & vMods_Length & ", '" & fUnquote(vMods_SkillSet) & "', '" & fUnquote(vMods_Competency) & "', " & fSqlBoolean(vMods_Completion) & ", " & fSqlBoolean(vMods_Reviewed) & ", " & fSqlBoolean(vMods_FeaAcc) & ", " & fSqlBoolean(vMods_FeaAud) & ", " & fSqlBoolean(vMods_FeaMob) & ", " & fSqlBoolean(vMods_FeaHyb) & ", " & fSqlBoolean(vMods_Vid) & ", '" & vMods_Type & "', " & vMods_Player & ", '" & vMods_Url & "', " & vMods_PreviewMax & ", '" & vMods_Script & "', " & vMods_Width & ", " & vMods_Height & ", " & fSqlBoolean(vMods_FullScreen) & ", " & fSqlBoolean(vMods_Fluid) & ", '" & vMods_Memo & "')" 
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub


  Sub sUpdateMods
    vSql = "UPDATE Mods SET"
    vSql = vSql & " Mods_Format     = '" & vMods_Format                	& " , " 
    vSql = vSql & " Mods_Title      = '" & fUnquote(vMods_Title)      	& "', " 
    vSql = vSql & " Mods_Active     =  " & vMods_Active               	& " , " 
    vSql = vSql & " Mods_AccessOk   = '" & vMods_AccessOk             	& "', " 
    vSql = vSql & " Mods_AccessNo   = '" & vMods_AccessNo             	& "', " 
    vSql = vSql & " Mods_ParentId   = '" & vMods_ParentId             	& "', " 
    vSql = vSql & " Mods_VuCert     = "  & vMods_VuCert               	& ",  " 
    vSql = vSql & " Mods_Desc       = '" & fUnquote(vMods_Desc)       	& "', " 
    vSql = vSql & " Mods_Outline    = '" & fUnquote(vMods_Outline)    	& "', " 
    vSql = vSql & " Mods_Goals      = '" & fUnquote(vMods_Goals)      	& "', " 
    vSql = vSql & " Mods_Skillset   = '" & fUnquote(vMods_Skillset)   	& "', " 
    vSql = vSql & " Mods_Competency = '" & fUnquote(vMods_Competency) 	& "', " 
    vSql = vSql & " Mods_Completion =  " & fSqlBoolean(vMods_Completion)	& " , " 

    vSql = vSql & " Mods_Reviewed   =  " & fSqlBoolean(vMods_Reviewed)	& " , " 

    vSql = vSql & " Mods_FeaAcc     =  " & fSqlBoolean(vMods_FeaAcc)	  & " , " 
    vSql = vSql & " Mods_FeaAud     =  " & fSqlBoolean(vMods_FeaAud)	  & " , " 
    vSql = vSql & " Mods_FeaMob     =  " & fSqlBoolean(vMods_FeaMob)	  & " , " 
    vSql = vSql & " Mods_FeaHyb     =  " & fSqlBoolean(vMods_FeaHyb)	  & " , " 
    vSql = vSql & " Mods_FeaVid     =  " & fSqlBoolean(vMods_FeaVid)	  & " , " 

    vSql = vSql & " Mods_Type       = '" & vMods_Type                 	& "', " 
    vSql = vSql & " Mods_Player     =  " & vMods_Player                 & " , " 
    vSql = vSql & " Mods_Url        = '" & fUnquote(vMods_Url)        	& "', " 
    vSql = vSql & " Mods_PreviewMax =  " & vMods_PreviewMax           	& " , " 
    vSql = vSql & " Mods_Script     = '" & vMods_Script               	& "', " 
    vSql = vSql & " Mods_Length     =  " & vMods_Length               	& " , " 

    vSql = vSql & " Mods_Memo       = '" & fUnquote(vMods_Memo)       	& "', " 
    vSql = vSql & " Mods_Width      =  " & vMods_Width                	& " , " 
    vSql = vSql & " Mods_Height     =  " & vMods_Height               	& " , " 
    vSql = vSql & " Mods_FullScreen =  " & vMods_FullScreen           	& " , " 
    vSql = vSql & " Mods_Fluid      =  " & vMods_Fluid                  & "   " 

    vSql = vSql & " WHERE Mods_Id   = '" & vMods_Id                   	& "'  "

'   sDebug
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub


  Sub sUpdateModsActive (vModsId, vModsActive)
    vSql = "UPDATE Mods SET"
    vSql = vSql & " Mods_Active     = "  & vModsActive              & "   " 
    vSql = vSql & " WHERE Mods_Id   = '" & vModsId                  & "'  "
'   sDebug
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub
  
  
  Sub sDeleteMods
    vSql = "DELETE Mods WHERE Mods_Id = '" & vMods_Id & "'"
    sOpenDbBase
    oDbBase.Execute(vSql)
    sCloseDbBase
  End Sub

  
  '...get all Mods
  Function fModsOptions
    Dim oRsBase
    fModsOptions = ""
    sOpenDbBase
    vSql = "Select * FROM Mods"
    Set oRsBase = oDbBase.Execute(vSql)    
    Do While Not oRsBase.EOF 
      fModsOptions = fModsOptions & "<option>" & oRsBase("Mods_Id") & "</option>" & vbCRLF
      oRsBase.MoveNext
    Loop      
    sCloseDbBase           
  End Function


  Function fPreviewMax
    fPreviewMax = ""
    If vMods_PreviewMax > 0 and vMods_PreviewMax < 27 Then
      fPreviewMax = Chr(64 + vMods_PreviewMax)
    End If    
  End Function


  '...clone a module
  Function fCloneMods (vModsId, vCloneId)
    Dim vOk
    vOk = False
    fCloneMods = False

    '...ensure vModsId exists and vCloneId is valid and does NOT exist
    If fModOk(vModsId) Then
      If Len(vCloneId) = 6 Then
        If IsNumeric(Left(vCloneId, 4)) Then
          If Instr("EN FR ES PT", Right(vCloneId, 2)) > 0 Then
'           If Not fModOk(vCloneId) Then
              vOk = True
'           End If
          End If
        End If
      End If
    End If
      
    If vOk Then
      '...delete if on file so we can reinsert
      vMods_Id = vCloneId
      sDeleteMods
    
      vSql = "SET ANSI_WARNINGS ON " _
           & "INSERT INTO Mods " _
           & "(Mods_Id, Mods_Format, Mods_Title, Mods_ParentId, Mods_Active, Mods_AccessOk, Mods_AccessNo, Mods_VuCert, Mods_Desc, Mods_Outline, Mods_Goals, Mods_Length, Mods_Type, Mods_Player, Mods_Url, Mods_Script, Mods_AssessmentUrl, Mods_AssessmentScript, Mods_SkillSet, Mods_PreviewMax, Mods_Memo, Mods_Completion, Mods_Reviewed, Mods_FeaAcc, Mods_FeaAud, Mods_FeaMob, Mods_FeaHyb, Mods_FeaVid) " _
           & "(SELECT '" & vCloneId & "' AS Mods_Id, Mods_Format, Mods_Title, Mods_ParentId, Mods_Active, Mods_AccessOk, Mods_AccessNo, Mods_VuCert, Mods_Desc, Mods_Outline, Mods_Goals, Mods_Length, Mods_Type, Mods_Player, Mods_Url, Mods_Script, Mods_AssessmentUrl, Mods_AssessmentScript, Mods_SkillSet, Mods_PreviewMax, Mods_Memo, Mods_Completion, Mods_Reviewed, Mods_FeaAcc, Mods_FeaAud, Mods_FeaMob, Mods_FeaHyb, Mods_FeaVid FROM Mods WHERE Mods_Id  = '" & vModsId & "')"
'     sDebug
      sOpenDbBase 
      oDbBase.Execute(vSql)
      sCloseDbBase    
      fCloneMods = True
    End If
  End Function


  '...Is Module Ok
  Function fModOk (vModsId)
    Dim oRsBase
    fModOk = False
    vSql = "SELECT Mods_Id FROM Mods WHERE Mods_Id= '" & vModsId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      fModOk = True
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Module Status (0:non doesn't exist, 1:active, 2:inactive, 3, active/no completion)
  Function fModsStatus (vModsId)
    Dim oRsBase
    vSql = "SELECT Mods_Active, Mods_Completion FROM Mods WHERE Mods_Id= '" & vModsId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    fModsStatus = 0
    If Not oRsBase.Eof Then 
      If oRsBase("Mods_Active") = 0 Then
        fModsStatus = 2
      ElseIf oRsBase("Mods_Completion") = 0 Then
        fModsStatus = 3
      Else
        fModsStatus = 1
      End If
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...get all Mods
  Function fModsOk (vMods)
    Dim oRsBase, aMods
    fModsOk = ""
    sOpenDbBase
    aMods = Split(vMods)
    For i = 0 To Ubound(aMods)
      vSql = "Select Mods_Id FROM Mods WHERE Mods_Id = '" & aMods(i) & "'"
      Set oRsBase = oDbBase.Execute(vSql)    
      If oRsBase.Eof Then fModsOk = fModsOk & " " & aMods(i) 
    Next
    sCloseDbBase           
  End Function

  '...Does Mod use VuCerts?
  Function fModsVuCert (vModId)
    fModsVuCert= False
    vSql = "SELECT Mods_VuCert FROM Mods WHERE Mods_Id= '" & vModId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If oRsBase.Eof Then 
      fModsVuCert= False
    Else
      fModsVuCert= oRsBase("Mods_VuCert") '...true or false
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Does Mod use ProgramCerts (typically CCHS), if so return the prog id that certifcate.asp needs
  Function fProgCert (vModId)
    fProgCert  = False
    vSql = ""
    vSql = vSql & "SELECT Prog.Prog_Id, Prog.Prog_CustomCert FROM Mods "
    vSql = vSql & "INNER JOIN Prog ON CHARINDEX(Mods.Mods_ID, Prog.Prog_Mods) > 0 "
    vSql = vSql & "INNER JOIN V5_Vubz.dbo.Catl ON CHARINDEX(Prog.Prog_Id, V5_Vubz.dbo.Catl.Catl_Programs) > 0 "
    vSql = vSql & "WHERE (Mods.Mods_ID = '" & vModId & "') AND (V5_Vubz.dbo.Catl.Catl_CustId = '" & svCustId & "') "

    fProgCert = ""
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      If oRsBase("Prog_CustomCert") Then
        fProgCert = oRsBase("Prog_Id")
      End If
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function
  

  '...get mods type (from ScoreModule.asp)
  Function fModsType (vModId)
    vSql = "SELECT Mods_Type FROM Mods WHERE Mods_Id = '" & vModId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fModsType = oRsBase("Mods_Type")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...Any description? used in myworldcode.asp
  Function fModsDesc (vModId)
    fModsDesc = False
    vSql = "SELECT Mods_Desc FROM Mods WHERE Mods_Id = '" & vModId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof And Len(Trim(oRsBase("Mods_Desc"))) > 0 Then fModsDesc = True
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function

  
  '...Get Mods No using the ModsId (used for RTE)
  Function fModsNoById (vModId)
    fModsNoById = 0
    vSql = "SELECT Mods_No FROM Mods WHERE Mods_Id = '" & vModId & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then fModsNoById = oRsBase("Mods_No")
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function

  
  '...see if ExamId is ok
  Function fExamOk (vExamId)
    Dim oRsBase
    fExamOk = True
    sOpenDbBase
    vSql = "Select TstH_Id FROM TstH WHERE TstH_Id = '" & vExamId & "'"
    Set oRsBase = oDbBase.Execute(vSql)    
    If oRsBase.Eof Then fExamOk = False
    sCloseDbBase           
  End Function  


  '...see which langs are offered with this mod, ie 1234 (do not include lang, ie do not pass 1234EN)
  Sub sModsLangs (vModLeft)
    Dim oRsBase2
    vMods_Langs = ""
    sOpenDbBase2
    vSql = "SELECT RIGHT(Mods_Id, 2) AS Mods_Lang FROM Mods WHERE (Mods_Id LIKE '" & vModLeft & "%')"
    Set oRsBase2 = oDbBase2.Execute(vSql)    
    Do While Not oRsBase2.Eof
      vMods_Langs = vMods_Langs & oRsBase2("Mods_Lang") & " "
      oRsBase2.MoveNext
    Loop
    sCloseDbBase2           
  End Sub
 
 
  '...show programs containing this module
  Function spProgByMods (vModsId)
    spProgByMods = ""
    sOpenCmdBase
    With oCmdBase
      .CommandText = "spProgByMods"     
      .Parameters.Append .CreateParameter("@Mods_Id",	adVarChar,  adParamInput, 7, vModsId)
    End With
    Set oRsBase = oCmdBase.Execute()
    Do While Not oRsBase.Eof
      spProgByMods = spProgByMods & " " & "<a target='_blank' href='Program.asp?vEditProgID=" & oRsBase("Prog_Id") & "'>" & oRsBase("Prog_Id") & "</a>"
      oRsBase.MoveNext
    Loop
    Set oRsBase = Nothing
    Set oCmdBase = Nothing
    sCloseDbBase
  End Function 


  '...show programs containing this sco (parent of a multi sco)
  Function spProgByScos (vModsId)
    spProgByScos = ""
    sOpenCmdBase
    With oCmdBase
      .CommandText = "spProgByScos"     
      .Parameters.Append .CreateParameter("@Scos_Id",	adVarChar,  adParamInput, 7, vModsId)
    End With
    Set oRsBase = oCmdBase.Execute()
    Do While Not oRsBase.Eof
      spProgByScos = spProgByScos & " " & "<a target='_blank' href='ProgramEdit.asp?vEditProgID=" & oRsBase("Prog_Id") & "'>" & oRsBase("Prog_Id") & "</a>"
      oRsBase.MoveNext
    Loop
    Set oRsBase = Nothing
    Set oCmdBase = Nothing
    sCloseDbBase
  End Function 


  '...update the program length of any programs containing a modified module (especially when a module is inactivated)
  Sub spUpdateProgLengths (vModsId)
    sOpenCmdBase
    With oCmdBase
      .CommandText = "spUpdateProgLengths"     
      .Parameters.Append .CreateParameter("@Mods_Id",	adChar,  adParamInput, 7, vModsId)
    End With
    oCmdBase.Execute()
    Set oCmdBase = Nothing
    sCloseDbBase
  End Sub

 
%>