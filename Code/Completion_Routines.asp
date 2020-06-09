<%
  Dim vCompletion_Debug
  Dim vUnit_No, vUnit_AcctId, vUnit_L0, vUnit_L0Title, vUnit_L1, vUnit_L1Title, vUnit_L2Id, vUnit_L2Title, vUnit_L3Id, vUnit_L3Title, vUnit_HO, vUnit_Active
  Dim bUnit_Eof, bUnit_HO, bUnit_Active

  Dim vRole_No, vRole_AcctId, vRole_Id, vRole_Title, vRole_Children
  Dim bRole_Eof, bParm_Eof

  '...determine Completion Level 
  Session("Completion_Level") = svMembLevel + fIf(svMembManager, 1, 0)

  '...initialize the beginning of each session (if requested), for each person that launched this session
  '   should only be set to "Y" once per session
  If Session("Completion_InitParms") = "Y" Then 

    '...grab the titles used in this account whenever page is loaded - IN THE SELECTED LANGUAGE
    '   and the zone offsets
    bParm_Eof = True
    vSql = "SELECT * FROM V5_Comp.dbo.Parm WITH (NOLOCK) WHERE Parm_AcctId = '" & svCustAcctId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)

    If Not oRs.Eof Then 

      Session("Completion_L9tit")  = fPhraId(oRs("Parm_L9tit"))
      Session("Completion_L3tit")  = fPhraId(oRs("Parm_L3tit"))
      Session("Completion_L3tits") = fPhraId(oRs("Parm_L3tits"))
      Session("Completion_L2tit")  = fPhraId(oRs("Parm_L2tit"))
      Session("Completion_L2tits") = fPhraId(oRs("Parm_L2tits"))
      Session("Completion_L1tit")  = fPhraId(oRs("Parm_L1tit"))
      Session("Completion_L1tits") = fPhraId(oRs("Parm_L1tits"))
      Session("Completion_L0tit")  = fPhraId(oRs("Parm_L0tit"))
      Session("Completion_L0tits") = fPhraId(oRs("Parm_L0tits"))
  
      Session("Completion_LearnerId") = fPhraId(oRs("Parm_LearnerId"))
  
      Session("Completion_L3str")  = oRs("Parm_L3str")
      Session("Completion_L2str")  = oRs("Parm_L2str")
      Session("Completion_L1str")  = oRs("Parm_L1str")
      Session("Completion_L0str")  = oRs("Parm_L0str")
      Session("Completion_RLstr")  = oRs("Parm_RLstr")
  
      Session("Completion_L3len")  = oRs("Parm_L3len")
      Session("Completion_L2len")  = oRs("Parm_L2len")
      Session("Completion_L1len")  = oRs("Parm_L1len")
      Session("Completion_L0len")  = oRs("Parm_L0len")
      Session("Completion_RLlen")  = oRs("Parm_RLlen")

      Session("Completion_L3all")  = fIf(Session("Completion_L3len") > 0 , String(Session("Completion_L3len"), "0"), "")
      Session("Completion_L2all")  = fIf(Session("Completion_L2len") > 0 , String(Session("Completion_L2len"), "0"), "")
      Session("Completion_L1all")  = fIf(Session("Completion_L1len") > 0 , String(Session("Completion_L1len"), "0"), "")
      Session("Completion_L0all")  = fIf(Session("Completion_L0len") > 0 , String(Session("Completion_L0len"), "0"), "")

      Session("Completion_EditLearners") = oRs("Parm_EditLearners")
  
      bParm_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb    


    '...get roles in HO and other (XX) for Utililties [ plus the count of HO and XX roles ]
    Session("Completion_Roles_HOcnt") = 0
    Session("Completion_Roles_XXcnt") = 0

    vSql = "SELECT Role_Id, Role_HO "_
         & " FROM "_
         & "  V5_Comp.dbo.Role WITH (nolock) "_
         & "WHERE "_
         & "  (Role_AcctId = '" & svCustAcctId & "') "_
         & "ORDER BY "_ 
         & "  Role_HO, Role_Order "

    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      If oRs("Role_HO") Then
        Session("Completion_Roles_HO") = Trim(Session("Completion_Roles_HO") & " " & oRs("Role_Id"))
        Session("Completion_Roles_HOcnt") = Session("Completion_Roles_HOcnt") + 1 
      Else
        Session("Completion_Roles_XX") = Trim(Session("Completion_Roles_XX") & " " & oRs("Role_Id"))
        Session("Completion_Roles_XXcnt") = Session("Completion_Roles_XXcnt") + 1 
      End If
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    

    Session("Completion_InitParms") = "N"

  End If  


  '...initialize Content?
  If Session("Completion_InitContent") = "Y" Then
    '...create a table of all available content from which the learner can make their selection
    vSql = " "_
         & "DELETE V5_Comp.dbo.RepC WHERE RepC_UserNo = " & svMembNo & " "_
         & "INSERT INTO V5_Comp.dbo.RepC (  "_
         & "	RepC_AcctId,    "_
         & "	RepC_UserNo,    "_
         & "	RepC_ProgId,    "_
         & "	RepC_ModsId     "_    
         & "  )  "_ 
         & "  SELECT "_    
         & "    AcctId,       "_
         &      svMembNo & ", " _
         & "    ProgId,       "_
         & "    ModsId        "_
         & "  FROM "_
         & "    V5_Comp.dbo.vMod1 WITH (NOLOCK) "_
         & "  WHERE "_
         & "    AcctId = '" & svCustAcctId & "'"
    sCompletion_Debug
    sOpenDb    
    oDb.Execute(vSql)           
    sCloseDb   

    Session("Completion_InitContent") = "N"
  End If


  '...testing for level 5 - will show SQL 
  If svMembLevel = 5 Or Session("Completion_Level") > 4 Or Lcase(svHost) = "localhost/v5" Then
    If Request.QueryString("vCompletion_Debug") = "y" Then 
      Session("Completion_Debug") = True
    ElseIf Request("vCompletion_Debug") = "n" Then 
      Session("Completion_Debug") = False
    End If
    If Request.QueryString("vCompletion_Initialize") = "y" Then 
      sData_Initialize
    End If
  End If


  '...get L1 Title (ie Region/State, etc)
  Function fL1Title (vL1)
    fL1Title = ""
    vSql = "SELECT DISTINCT Unit_L1Title FROM V5_Comp.dbo.Unit WITH (NOLOCK) WHERE Unit_AcctId = '" & svCustAcctId & "' AND Unit_L1 = '" & vL1 & "'"
    sCompletion_Debug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fL1Title = oRs2("Unit_L1Title")
    sCloseDb2    
  End Function


  '...get L0 Title (ie Theatre/Store/Dept)
  Function fL0Title (vL0)
    fL0Title = ""
    vSql = "SELECT Unit_L0Title FROM V5_Comp.dbo.Unit WITH (NOLOCK) WHERE Unit_AcctId = '" & svCustAcctId & "' AND Unit_L0 = '" & vL0 & "'"
    sCompletion_Debug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fL0Title = oRs2("Unit_L0Title")
    sCloseDb2    
  End Function


  '...shows latest vSql statement
  Sub sCompletion_Debug 
    If Session("Completion_Debug") Then 
      On Error Resume Next
      If Err.Number = 0 Then 
        vCompletion_Debug = vCompletion_Debug & "<br><b><font color='ORANGE'>" & vSql & "</font></b><br>"
      End If
      On Error GoTo 0
    End If
  End Sub


  Sub sGetUnit (vUnitNo)
    bUnit_Eof = True
    vSql = "SELECT * FROM V5_Comp.dbo.Unit WITH (NOLOCK) WHERE Unit_No = " & vUnitNo
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadUnit
      bUnit_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb    
  End Sub
  

  Sub sGetUnitByL0 (vL0)
    bUnit_Eof = True
    vSql = "SELECT * FROM V5_Comp.dbo.Unit WITH (NOLOCK) WHERE Unit_L0 = '" & vL0 & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadUnit
      bUnit_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb    
  End Sub
  
  Function fUnitExists (vL1, vL0)
    fUnitExists = True
    vSql = "SELECT Unit_No FROM V5_Comp.dbo.Unit WITH (NOLOCK) WHERE Unit_AcctId = '" & svCustAcctId & "' AND Unit_L1 = '" & vL1 & "' OR Unit_L0 = '" & vL0 & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then fUnitExists = False
    Set oRs = Nothing
    sCloseDb    
  End Function  


  Function fLocnExists (vL0)
    fLocnExists = True
    vSql = "SELECT TOP 1 Unit_No FROM V5_Comp.dbo.Unit WITH (NOLOCK) WHERE Unit_AcctId = '" & svCustAcctId & "' AND Unit_L0 = '" & vL0 & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then fLocnExists = False
    Set oRs = Nothing
    sCloseDb    
  End Function  

  

  Sub sReadUnit
    vUnit_No        = oRs("Unit_No")
    vUnit_L0        = oRs("Unit_L0")
    vUnit_L0Title   = oRs("Unit_L0Title")
    vUnit_L1        = oRs("Unit_L1")
    vUnit_L1Title   = oRs("Unit_L1Title")
    bUnit_HO        = oRs("Unit_HO")
    bUnit_Active    = oRs("Unit_Active")
    vUnit_HO        = fIf(bUnit_HO, 1, 0)
    vUnit_Active    = fIf(bUnit_Active, 1, 0)
  End Sub


  Sub sInsertUnit (vL1, vL1Title, vL0, vL0Title)
    vSql = ""                                _
         & " INSERT INTO V5_Comp.dbo.Unit"   _
         & " ("                              _ 
         & " 	Unit_AcctId,"                  _
         & " 	Unit_L1,"                      _
         & " 	Unit_L1Title,"                 _
         & " 	Unit_L0,"                      _
         & " 	Unit_L0Title"                  _
         &  ")"                              _ 
         & " VALUES"                         _
         & " ('"                             _ 
         &     svCustAcctId           & "', '"_
         &     vL1                    & "', '"_
         &     vL1Title               & "', '"_
         &     vL0                    & "', '"_
         &     vL0Title               &    "'"_
         &  ")"                                
    Do While Instr(vSql, "  ") > 0 : vSql = Replace(vSql, "  ", " ") : Loop '...strip spaces
    sCompletion_Debug
    sOpenDb    
    oDb.Execute(vSql)           
    sCloseDb  
  End Sub


  Sub sDeleteUnit (vUnitNo)
    vSql = "DELETE V5_Comp.dbo.Unit WHERE Unit_No = " & vUnitNo
    sCompletion_Debug
    sOpenDb    
    oDb.Execute(vSql)           
    sCloseDb  
  End Sub


  '...get my Role title
  Function fRole_Title(vRole)
    bRole_Eof = True
    fRole_Title = ""
    vSql = "SELECT Role_Title FROM V5_Comp.dbo.Role WITH (nolock) WHERE Role_AcctId = '" & svCustAcctId & "' AND Role_Id = '" & vRole & "' ORDER BY Role_Order"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      fRole_Title = Trim(oRs("Role_Title"))
      bRole_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb    
  End Function

  
  '...get Children Roles for the current Role (ie what Roles can this person see in the report)
  Function fRole_Children(vRole)
    bRole_Eof = True
    fRole_Children = ""
    vSql = "SELECT Role_Children FROM V5_Comp.dbo.Role WITH (NOLOCK) WHERE Role_AcctId = '" & svCustAcctId & "' AND Role_Id = '" & vRole & "' ORDER BY Role_Order"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      fRole_Children = fOkValue(oRs("Role_Children"))
      fRole_Children = Replace(fRole_Children, " ", "")
      bRole_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb    
  End Function


  '...get all Roles for the account (typically for admins, etc)
  Function fRole_All()
    bRole_Eof = True
    fRole_All = ""
    vSql = "SELECT Role_Id FROM V5_Comp.dbo.Role WITH (NOLOCK) WHERE Role_AcctId = '" & svCustAcctId & "' ORDER BY Role_Order"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      fRole_All = fRole_All & oRs("Role_Id") & ","
      bRole_Eof = False
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb    
    '...strip off trailing comma
    fRole_All = fIf(Len(fRole_All) > 1, Left(fRole_All, Len(fRole_All)-1), "")
  End Function


  '...clean out the report tables
  Sub sResetReport
    sOpenDb

    vSql = " DELETE V5_Comp.dbo.RepM WHERE RepM_UserNo = " & svMembNo
    sCompletion_Debug
    oDb.Execute(vSql)

    vSql = " DELETE V5_Comp.dbo.RepL WHERE RepL_UserNo = " & svMembNo
    sCompletion_Debug
    oDb.Execute(vSql)

    vSql = " DELETE V5_Comp.dbo.Rcnt WHERE Rcnt_UserNo = " & svMembNo
    sCompletion_Debug
    oDb.Execute(vSql)

    vSql = " UPDATE V5_Comp.dbo.RepC SET RepC_Selected = 0 WHERE (RepC_UserNo = " & svMembNo & ")"
    sCompletion_Debug
    oDb.Execute(vSql)

    sCloseDb
  End Sub


  '...flag the selected program / module (one at a time)
  Sub sCreateReport (vProgId, vModsId)
    vSql = " UPDATE V5_Comp.dbo.RepC "_
         & "   SET RepC_Selected = 1 "_ 
         & "   WHERE "_
         & "     (RepC_UserNo =  " & svMembNo & ")   AND "_
         & "     (RepC_ProgId = '" & vProgId  & "')  AND "_
         & "     (RepC_ModsId = '" & vModsId  & "') "
    sCompletion_Debug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  '...when finished create the learner/content tables
  Sub sEndReport
    Dim vRole : vRole = "('" & Replace(Session("Completion_RoleP"), ",", "', '") & "')"


    '...create a table of valid learners/programs/modules - union that with any extended learning options
    vSql = " "_
         & "DELETE V5_Comp.dbo.RepM WHERE RepM_UserNo = " & svMembNo & " "_

         & "INSERT INTO V5_Comp.dbo.RepM "_
         & "  ( "_
         & "	RepM_AcctId,"_
         & "	RepM_UserNo,"_
         & "	RepM_MembNo,"_    
         & "	RepM_ProgId,"_    
         & "	RepM_ModsId"_    
         & "  ) "_ 

         & "  SELECT DISTINCT '"_      
         &      svCustAcctId & "'                         AS RepM_AcctId, " _  
         &      svMembNo & "                              AS RepM_UserNo, " _  
         & "    Me.Memb_No										            AS RepM_MembNo, " _   
         & "    LEFT(JP.Jobs_Prog_ProgId, 5)				      AS RepM_ProgId, " _  
         & "    LEFT(PM.Prog_Mods_ModsId,4)		            AS RepM_ModsId  " _  

         & "  FROM " _           
         & "    V5_Vubz.dbo.Memb					AS Me WITH (NOLOCK)		                                                  INNER JOIN " _
         & "    V5_Vubz.dbo.Crit_Jobs			AS CJ WITH (NOLOCK)		ON Me.Memb_Criteria	    = CJ.Crit_Jobs_CritNo     INNER JOIN " _
         & "    V5_Vubz.dbo.Jobs_Prog			AS JP WITH (NOLOCK)		ON CJ.Crit_Jobs_JobsNo	= JP.Jobs_Prog_JobsNo     INNER JOIN " _
         & "    V5_Base.dbo.Prog_Mods		  AS PM WITH (NOLOCK)		ON JP.Jobs_Prog_ProgId	= PM.Prog_Mods_ProgId                " _

         & "  WHERE " _     
         & "    (Me.Memb_AcctId = '" & svCustAcctId & "') AND " _ 
         & "    (ISNUMERIC(Me.Memb_Criteria) = 1)         AND " _
         & "    Me.Memb_Active = 1                        AND " _
         & "    Me.Memb_Internal = 0                          " _

         & "  UNION " _     
  
         & "  SELECT DISTINCT '"_      
         &      svCustAcctId & "'                         AS RepM_AcctId, " _  
         &      svMembNo & "                              AS RepM_UserNo, " _  
         & "    Me.Memb_No										            AS RepM_MembNo, " _   
         & "    LEFT(JP.Jobs_Prog_ProgId, 5)              AS RepM_ProgId, " _     
         & "    LEFT(PM.Prog_Mods_ModsId, 4)              AS RepM_ProgId  " _    

         & "  FROM " _           
         & "    V5_Vubz.dbo.Memb					AS	Me WITH (NOLOCK)	                                                                                                   INNER JOIN " _
         & "    V5_Vubz.dbo.Jobs_Prog			AS	JP WITH (NOLOCK)	ON Me.Memb_AcctId       = JP.Jobs_Prog_AcctId AND CHARINDEX(JP.Jobs_Prog_JobsId, Me.Memb_Jobs) > 0 INNER JOIN " _
         & "    V5_Base.dbo.Prog_Mods			AS	PM WITH (NOLOCK)	ON JP.Jobs_Prog_ProgId  = PM.Prog_Mods_ProgId " _

         & "  WHERE " _     
         & "    (Me.Memb_AcctId = '" & svCustAcctId & "') AND " _ 
         & "    (ISNUMERIC(Me.Memb_Criteria) = 1)         AND " _
         & "    Me.Memb_Active = 1                        AND " _
         & "    Me.Memb_Internal = 0                          "

'        & "    V5_Vubz.dbo.Jobs_Prog			AS	JP WITH (NOLOCK)	ON Me.Memb_AcctId       = JP.Jobs_Prog_AcctId AND Me.Memb_Jobs = JP.Jobs_Prog_JobsId    INNER JOIN " _


    sCompletion_Debug
    sOpenDb
    oDb.CommandTimeout = 120      
    oDb.Execute(vSql)           
    sCloseDb  




    '  create table of all learners in selected roles and locations who access selected content
      vSql = " "_
         & " DELETE V5_Comp.dbo.RepL WHERE RepL_UserNo = " & svMembNo & " "_

         & " INSERT INTO V5_Comp.dbo.RepL "_

         & " 	 SELECT DISTINCT '"_ 
         &       svCustAcctId & "' 																							AS AcctId, " _
         &       svMembNo & " 																									AS UserNo, " _
         & "     NULL AS L3, "_
         & "     NULL AS L2, "_
         & "     SUBSTRING(Cr.Crit_Id, " & Session("Completion_L1str") & ", " & Session("Completion_L1len") & ") 	AS L1, "_
         & "     SUBSTRING(Cr.Crit_Id, " & Session("Completion_L0str") & ", " & Session("Completion_L0len") & ") 	AS L0, "_
         & "     SUBSTRING(Cr.Crit_Id, " & Session("Completion_RLstr") & ", " & Session("Completion_RLlen") & ") 	AS RL, "_
         & "     Me.Memb_No 																										AS MembNo, "_
         & "     Me.Memb_Id 																										AS MembId, "_
         & "     Me.Memb_FirstName 																							AS FirstName, "_
         & "     Me.Memb_LastName 																							AS LastName "_

         & "   FROM "_
         & "     V5_Vubz.dbo.Memb AS Me WITH (NOLOCK)                                                 INNER JOIN "_ 
         & "     V5_Vubz.dbo.Crit AS Cr WITH (NOLOCK) ON  Me.Memb_Criteria  = Cr.Crit_No              INNER JOIN "_
         & "     V5_Comp.dbo.RepC AS Rc WITH (NOLOCK) ON  Me.Memb_AcctId    = Rc.RepC_AcctId          INNER JOIN "_
         & "     V5_Comp.dbo.RepM AS Rm WITH (NOLOCK) ON  Me.Memb_No        = Rm.RepM_MembNo          AND "_
         & "                                              Rc.RepC_AcctId    = Rm.RepM_AcctId          AND "_ 
         & "                                              Rc.RepC_ProgId    = Rm.RepM_ProgId          AND "_
         & "                                              Rc.RepC_ModsId    = Rm.RepM_ModsId              "_

         & "   WHERE "_    
         & "     (Me.Memb_AcctId = '" & svCustAcctId & "') 														AND "_ 
         & "     (Me.Memb_Level IN (2,3,4)) 																					AND "_
         & "     (Me.Memb_Active = 1) 																								AND "_ 
         & "     (Me.Memb_Internal = 0) 																							AND "_ 
         & "     (ISNUMERIC(Me.Memb_Criteria) = 1) 																		AND "_ 
         & "     (SUBSTRING(Cr.Crit_Id, " & Session("Completion_RLstr") & ", " & Session("Completion_RLlen") & ") IN " & vRole & ") 	AND "_
         & "     (Rc.RepC_Selected = 1) "

    If Session("Completion_L1val") <> Session("Completion_L1all") Then
      vSql = vSql & " AND (SUBSTRING(Cr.Crit_Id, " & Session("Completion_L1str") & ", " & Session("Completion_L1len") & ") = '" & Session("Completion_L1val") & "')"
    End If

    If Session("Completion_L0val") <> Session("Completion_L0all") Then
      vSql = vSql & " AND (SUBSTRING(Cr.Crit_Id, " & Session("Completion_L0str") & ", " & Session("Completion_L0len") & ") = '" & Session("Completion_L0val") & "')"
    End If

    sCompletion_Debug
    sOpenDb
    oDb.CommandTimeout = 120 
    oDb.Execute(vSql)
    sCloseDb

    '   create a summary table containing the number of mods in each selected program to determine % completion
    vSql = " "_
         & " INSERT INTO V5_Comp.dbo.Rcnt "_

         & "   SELECT "_     
         & "     RepC_AcctId, "_ 
         & "     RepC_UserNo, "_ 
         & "     RepC_ProgId, "_ 
         & "     COUNT(RepC_ModsId), "_
         & "     MAX (CASE "_ 
         & "       WHEN Prog.Prog_AssessmentScore > 0 THEN Prog.Prog_AssessmentScore * 100 "_ 
         & "       WHEN Cust.Cust_AssessmentScore > 0 THEN Cust.Cust_AssessmentScore * 100 "_
         & "       ELSE 80 END) "_
         & "   FROM "_         
         & "     V5_Comp.dbo.RepC AS RepC WITH (NOLOCK) INNER JOIN "_
         & "   	 V5_Vubz.dbo.Cust AS Cust WITH (NOLOCK) ON RepC.RepC_AcctId = Cust.Cust_AcctId INNER JOIN "_
         & "   	 V5_Base.dbo.Prog AS Prog WITH (NOLOCK) ON RepC.RepC_ProgId + 'EN' = Prog.Prog_Id "_
         & "   WHERE "_     
         & "   	 (RepC.RepC_UserNo = " & svMembNo & ") AND "_
         & "   	 (RepC.RepC_Selected = 1) "_
         & "   GROUP BY "_ 
         & "   	 RepC_AcctId, "_ 
         & "   	 RepC_UserNo, "_ 
         & "   	 RepC_ProgId "

    sCompletion_Debug
    sOpenDb
    oDb.CommandTimeout = 120 
    oDb.Execute(vSql)
    sCloseDb


    '...create a table of valid scores for selected programs 
    vSql = " "_
         & "DELETE V5_Comp.dbo.RepS WHERE RepS_UserNo = " & svMembNo & " "_

         & "INSERT INTO V5_Comp.dbo.RepS ( "_
         & "	RepS_AcctId,"_
         & "	RepS_UserNo,"_
         & "	RepS_MembNo,"_    
         & "	RepS_ProgId,"_    
         & "	RepS_ModsId,"_    
         & "	RepS_NoAttempts,"_    
         & "	RepS_BestDate,"_    
         & "	RepS_BestScore"_    
         & "  )"_ 

         & "  SELECT '"_      
         &      svCustAcctId & "'                         AS AcctId," _  
         &      svMembNo & "                              AS UserNo," _  
         & "    Me.Memb_No															  AS MembNo,"_ 
         & "    Rm.RepM_ProgId														AS ProgId,"_ 
         & "    Rm.RepM_ModsId														AS ModsId,"_ 
         & "    COUNT(Lo.Logs_Item)												AS NoAttempts,"_ 
         & "    MAX(Lo.Logs_Posted)												AS BestDate,"_ 
         & "    MAX(CAST(RIGHT(Lo.Logs_Item, 3) AS INT))	AS BestScore"_
         & "  FROM"_         
         & "    V5_Vubz.dbo.Memb													AS Me	WITH (NOLOCK) 											                            INNER JOIN "_
         & "    V5_Comp.dbo.RepM                  			  AS Rm WITH (NOLOCK) ON Me.Memb_No = Rm.RepM_MembNo							      INNER JOIN "_
         & "    V5_Base.dbo.Mods													AS Mo WITH (NOLOCK) ON Rm.RepM_ModsId + 'EN' = Mo.Mods_Id						  INNER JOIN "_
         & "    V5_Vubz.dbo.Logs													AS Lo WITH (NOLOCK) ON Rm.RepM_MembNo = Lo.Logs_MembNo                AND        "_ 
         & "                                                                     Rm.RepM_ModsId = LEFT(Lo.Logs_Item, 4)         INNER JOIN "_
         & "    V5_Comp.dbo.Rcnt                          AS Rc WITH (NOLOCK) ON Rm.RepM_ProgId = Rc.Rcnt_ProgId                           "_ 
         & "  WHERE"_     
         & "    (Me.Memb_AcctId = '" & svCustAcctId & "') AND"_ 
         & "    (Lo.Logs_AcctId = '" & svCustAcctId & "') AND"_ 
         & "    (Rc.Rcnt_AcctId = '" & svCustAcctId & "') AND"_ 
         & "    (Mo.Mods_Active = 1) 											AND"_ 
         & "    (Mo.Mods_Completion = 1) 									AND"_ 
         & "    (Lo.Logs_Type = 'T')											AND"_
         & "    (Rc.Rcnt_UserNo = " & svMembNo & ")          "_
         & "  GROUP BY"_ 
         & "    Me.Memb_No,"_ 
         & "    Rm.RepM_ProgId,"_ 
         & "    Rm.RepM_ModsId,"_ 
         & "    LEFT(Lo.Logs_Item, 4)"
  
    sCompletion_Debug
    sOpenDb    
    oDb.CommandTimeout = 120 
    oDb.Execute(vSql)           
    sCloseDb  


    '   update RepS with completion status
    vSql = "UPDATE V5_Comp.dbo.RepS"_
         & "  SET RepS_Completed = CASE WHEN Rs.Reps_BestScore >= Rc.Rcnt_Mastery THEN 1 ELSE 0 END"_
         & "  FROM V5_Comp.dbo.RepS AS Rs INNER JOIN V5_Comp.dbo.Rcnt AS Rc ON Rs.RepS_AcctId = Rc.Rcnt_AcctId AND Rs.RepS_ProgId = Rc.Rcnt_ProgId" 
    sCompletion_Debug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb         
  End Sub
  


  Function fLocation (vMembCriteria)
  	Dim vSelect, vCritNo, vCritId
    fLocation = vbCrLf
    vSql = " SELECT "_    
         & "   Crit.Crit_No, Crit.Crit_Id, Unit.Unit_L1, Unit.Unit_L1Title, Unit.Unit_L0, Unit.Unit_L0Title "_
         & " FROM "_  
         & "   V5_Vubz.dbo.Crit Crit WITH (NOLOCK) INNER JOIN "_ 
         & "   V5_Comp.dbo.Unit Unit WITH (NOLOCK) ON LEFT(Crit.Crit_Id, " & Session("Completion_L1len") & ") = Unit.Unit_L1 AND SUBSTRING(Crit.Crit_Id, " & Session("Completion_L0str") & ", " & Session("Completion_L0len") & ") = Unit.Unit_L0 "_
         & " WHERE "_    
         & "   			(Crit.Crit_AcctId = '" & svCustAcctId & "')	"_         
         & "   	AND	(Unit.Unit_Active = 1) 	"         
		If Session("Completion_Level") = 3 Then 
    vSql = vSql _ 
         & "   AND	(Unit.Unit_L1 = '" & vMyL1 & "')							"_     
         & "   AND	(Unit.Unit_L0 = '" & vMyL0 & "')							"_     
         & "   AND	(CHARINDEX(RIGHT(Crit.Crit_Id, " & Session("Completion_RLlen") & "), '" & vMyCh & "') > 0)"
		End If
    vSql = vSql _ 
         & " ORDER BY "_
         & "   Crit.Crit_Id "  

    sCompletion_Debug

    vCnt = 0
    vSelect = ""

    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vCritNo       = oRs("Crit_No")
      vCritId       = oRs("Crit_Id")
      vUnit_L1      = oRs("Unit_L1")
      vUnit_L0      = oRs("Unit_L0")
      vUnit_L1Title = oRs("Unit_L1Title")
      vUnit_L0Title = oRs("Unit_L0Title")

'     ... someone put a manager into multiple groups (should not have any groups) so it bombed, these are now changed to 0 as of 2015-08-11
'     vSelect = fIf(cInt(vMembCriteria) <> vCritNo, "", " selected ")
      vSelect = fIf(fPureInt(vMembCriteria) <> vCritNo, "", " selected ")

      fLocation = fLocation & "<option value='" & vCritNo & "'" & vSelect & ">" & vUnit_L1 & " (" & vUnit_L1Title & ")  - " & vUnit_L0 & " (" & vUnit_L0Title & ") - " & Right(vCritId, 2) & "</option>" & vbCrLf
      vCnt = vCnt + 1
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
  End Function

%>

