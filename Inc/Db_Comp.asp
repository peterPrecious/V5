<%
  Dim vComp_No, vComp_AcctId, vComp_Number, vComp_Lang, vComp_Title, vComp_Active, vComp_BO1, vComp_BO2, vComp_BO3, vComp_BO4
  Dim vComp_AlteredOn, vComp_AlteredBy
  
  Dim vComp_Desc1, vComp_Desc1_L1, vComp_Desc1_L2, vComp_Desc1_L3, vComp_Desc1_L4, vComp_Desc1_L5, vComp_Desc1_L6, vComp_Desc1_L7, vComp_Desc1_L8 
  Dim vComp_Desc2, vComp_Desc2_L1, vComp_Desc2_L2, vComp_Desc2_L3, vComp_Desc2_L4, vComp_Desc2_L5, vComp_Desc2_L6, vComp_Desc2_L7, vComp_Desc2_L8  
  Dim vComp_Desc3, vComp_Desc3_L1, vComp_Desc3_L2, vComp_Desc3_L3, vComp_Desc3_L4, vComp_Desc3_L5, vComp_Desc3_L6, vComp_Desc3_L7, vComp_Desc3_L8  
  Dim vComp_Desc4, vComp_Desc4_L1, vComp_Desc4_L2, vComp_Desc4_L3, vComp_Desc4_L4, vComp_Desc4_L5, vComp_Desc4_L6, vComp_Desc4_L7, vComp_Desc4_L8  

  Dim vComp_Mods1, vComp_Clss1, vComp_Actv1
  Dim vComp_Mods2, vComp_Clss2, vComp_Actv2
  Dim vComp_Mods3, vComp_Clss3, vComp_Actv3
  Dim vComp_Mods4, vComp_Clss4, vComp_Actv4

  Dim vComp_Desc, vComp_Eof
  
  Dim vCompSize : vCompSize = 4 '...this sets the number of rows in the last drop down called
  


  '____ Comp  ________________________________________________________________________

  '...returns a full recordset of competencies
  Sub spCompSelectAll (vAcctId, vLang, vLevel)
    sOpenCmdGap
    With oCmdGap
      .CommandText = "spCompSelectAll"
      .Parameters.Append .CreateParameter("@Comp_AcctId", adVarChar, adParamInput,  4, vAcctId)
      .Parameters.Append .CreateParameter("@Comp_Lang",		adVarChar, adParamInput,  2, vLang)
      .Parameters.Append .CreateParameter("@Level",       adTinyInt, adParamInput,   , vLevel)
    End With
    Set oRs = oCmdGap.Execute()
  End Sub


  '...get competency by Number and Level
  Sub spCompSelectByNumber (vAcctId, vLang, vNumber, vLevel)
    sOpenCmdGap
    With oCmdGap
      .CommandText = "spCompSelectByNumber"
      .Parameters.Append .CreateParameter("@Comp_AcctId", adVarChar, adParamInput,  4, vAcctId)
      .Parameters.Append .CreateParameter("@Comp_Lang",		adVarChar, adParamInput,  2, vLang)
      .Parameters.Append .CreateParameter("@Comp_Number",	adTinyInt, adParamInput,   , vNumber)
      .Parameters.Append .CreateParameter("@Level",       adTinyInt, adParamInput,   , vLevel)
    End With
    Set oRs = oCmdGap.Execute()
    vComp_Number = vNumber
    vComp_Title  = oRs("Comp_Title")
    vComp_Desc   = oRs("Comp_Desc")
    Set oRs = Nothing      
    Set oCmdGap = Nothing
    sCloseDbGap    
  End Sub


  Sub sReadCompByLevel
    vComp_Number      = oRs("Comp_Number")
    vComp_Title   		= oRs("Comp_Title")
    vComp_BO1         = oRs("Comp_BO1")
    vComp_BO2         = oRs("Comp_BO2")
    vComp_BO3         = oRs("Comp_BO3")
    vComp_BO4         = oRs("Comp_BO4")
    vComp_Desc      	= oRs("Comp_Desc")
  End Sub


  '...Get Items for dropdown
  Function spCompOptionsActv (vAcctId, vMembNo, vGapYear, vLang, vNumber, vLevel)
    spCompOptionsActv = ""
    vCompSize = 0
    sOpenCmdGap
    With oCmdGap
      .CommandText = "spCompOptionsActv"
      .Parameters.Append .CreateParameter("@AcctId",  	adVarChar,  	adParamInput,  4, vAcctId)
      .Parameters.Append .CreateParameter("@MembNo",  	adInteger,  	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@Year",	  	adSmallInt, 	adParamInput,   , vGapYear)
      .Parameters.Append .CreateParameter("@Lang",			adVarChar, 	adParamInput,  2, vLang)
      .Parameters.Append .CreateParameter("@Number",		adTinyInt, 	adParamInput,   , vNumber)
      .Parameters.Append .CreateParameter("@Level",    	adTinyInt, 	adParamInput,   , vLevel)
    End With
    Set oRs = oCmdGap.Execute()
    Do While Not oRs.Eof 
      spCompOptionsActv = spCompOptionsActv & "<option " & fIf(IsNull(oRs("Selected")), "", "selected") & " value='" & oRs("Actv_Id") & "'>" & oRs("Actv_Title") & fHighlight & "</option>" & vbCrLf
      vCompSize = vCompSize + 1
      oRs.MoveNext
    Loop      
    Set oRs = Nothing      
    Set oCmdGap = Nothing
    sCloseDbGap  
  End Function 

  Function spCompOptionsMods (vAcctId, vMembNo, vGapYear, vLang, vNumber, vLevel)
    Dim vCompleted, bOk    
    spCompOptionsMods = ""
    vCompSize = 0
    sOpenCmdGap
    With oCmdGap
      .CommandText = "spCompOptionsMods"
      .Parameters.Append .CreateParameter("@AcctId",  	adVarChar,  	adParamInput,  4, vAcctId)
      .Parameters.Append .CreateParameter("@MembNo",  	adInteger,  	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@Year",	  	adSmallInt, 	adParamInput,   , vGapYear)
      .Parameters.Append .CreateParameter("@Lang",			adVarChar, 	adParamInput,  2, vLang)
      .Parameters.Append .CreateParameter("@Number",		adTinyInt, 	adParamInput,   , vNumber)
      .Parameters.Append .CreateParameter("@Level",    	adTinyInt, 	adParamInput,   , vLevel)
    End With
    Set oRs = oCmdGap.Execute()
    Do While Not oRs.Eof 
      '...only include modules that have not been completed within past 18 months
      bOk = False
      vCompleted = fCompleted (vMembNo, oRs("Mods_Id"))
      If vCompleted = "" Then 
        bOk = True
      ElseIf IsDate(vCompleted) Then
        If DateDiff("m", vCompleted, Now()) > 18 Then
          bOk = True
        End If
      End If
      If bOk Then 
        spCompOptionsMods = spCompOptionsMods & "<option " & fIf(IsNull(oRs("Selected")), "", "selected") & " value='" & oRs("Mods_Id") & "'>" & oRs("Mods_Title") & fHighlight & "</option>" & vbCrLf
      End If        
      vCompSize = vCompSize + 1
      oRs.MoveNext
    Loop      
    Set oRs = Nothing      
    Set oCmdGap = Nothing
    sCloseDbGap  
  End Function 

  Function spCompOptionsClss (vAcctId, vMembNo, vGapYear, vLang, vNumber, vLevel)
    spCompOptionsClss = ""
    vCompSize = 0
    sOpenCmdGap
    With oCmdGap
      .CommandText = "spCompOptionsClss"
      .Parameters.Append .CreateParameter("@AcctId",  	adVarChar,  	adParamInput,  4, vAcctId)
      .Parameters.Append .CreateParameter("@MembNo",  	adInteger,  	adParamInput,   , vMembNo)
      .Parameters.Append .CreateParameter("@Year",	  	adSmallInt, 	adParamInput,   , vGapYear)
      .Parameters.Append .CreateParameter("@Lang",			adVarChar, 	adParamInput,  2, vLang)
      .Parameters.Append .CreateParameter("@Number",		adTinyInt, 	adParamInput,   , vNumber)
      .Parameters.Append .CreateParameter("@Level",    	adTinyInt, 	adParamInput,   , vLevel)
    End With
    Set oRs = oCmdGap.Execute()
    Do While Not oRs.Eof 
      spCompOptionsClss = spCompOptionsClss & "<option " & fIf(IsNull(oRs("Selected")), "", "selected") & " value='" & oRs("Clss_Id") & "'>" & oRs("Clss_Title") & fHighlight & "</option>" & vbCrLf
      oRs.MoveNext
      vCompSize = vCompSize + 1
    Loop      
    Set oRs = Nothing      
    Set oCmdGap = Nothing
    sCloseDbGap  
  End Function 


  '...Get Titles for dropdown (Modules.asp)  ' ie: spCompTitles ("2818", "EN", vMods_Competency) 
  Function spCompTitles (vAcctId, vLang, vCompNos)
    spCompTitles = ""
    sOpenCmdGap
    With oCmdGap
      .CommandText = "spCompTitles"
      .Parameters.Append .CreateParameter("@AcctId",  	adVarChar,  	adParamInput,  4, vAcctId)
      .Parameters.Append .CreateParameter("@Lang",			adVarChar, 	adParamInput,  2, vLang)
    End With
    Set oRs = oCmdGap.Execute()
    Do While Not oRs.Eof 
      spCompTitles = spCompTitles & "<option " & fSelect (oRs("ComH_Number"), vMods_Competency) & " value='" & oRs("ComH_Number") & "'>" & oRs("ComH_Title") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop      
    Set oRs = Nothing      
    Set oCmdGap = Nothing
    sCloseDbGap  
  End Function 

  '...this creates a highlighted {Selected] value to help hight the selected value when printing
  Function fHighLight
    fHighLight = fIf(IsNull(oRs("Selected")), "", fIf(svLang = "FR", "&nbsp;[Sélectionné]", "&nbsp;[Selected]"))
  End Function

%>