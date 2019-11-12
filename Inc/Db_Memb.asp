<%
  Dim vMemb_AcctId, vMemb_Id, vMemb_Pwd, vMemb_No, vMemb_FirstName, vMemb_LastName, vMemb_Email, vMemb_Level,  vMemb_FirstVisit, vMemb_LastVisit, vMemb_NoVisits, vMemb_Cust, vMemb_Criteria, vMemb_Group2, vMemb_Group3, vMemb_JobsNo, vMemb_Skills, vMemb_Memo, vMemb_Organization, vMemb_NoHours, vMemb_Expires, vMemb_Online
  Dim vMemb_Active, vMemb_Internal, vMemb_Browser, vMemb_Programs, vMemb_ProgramsAdded, vMemb_EcomG2Alert, vMemb_Duration, vMemb_Jobs, vMemb_MaxSponsor, vMemb_Sponsor, vMemb_VuNews
  Dim vMemb_Auth, vMemb_MyWorld, vMemb_LCMS, vMemb_Ecom, vMemb_Channel, vMemb_VuBuild, vMemb_Manager, vMemb_AlteredOn, vMemb_AlteredBy
  Dim vMemb_Guid, vMemb_LastAssignedBy
  Dim vMemb_Eof


  '____ Memb  ________________________________________________________________________

  '...Returns an learner list: FirstName, LastName, No, Email (ActnAdd.asp EmailAlert.asp)
  Function fMemb_List
    fMemb_List = ""
    vSql = " " _
         & " SELECT "_ 
         & "   Memb_No, Memb_FirstName, Memb_LastName, Memb_Email "_
         & " FROM "_ 
         & "   Memb WITH (NOLOCK) "_ 
         & " WHERE "_ 
         & "   Memb_AcctId = '" & svCustAcctId & "' AND "_ 
         & "   Memb_Active = 1 AND "_ 
         & "   Len(Memb_Email) > 3 AND "_ 
         & "   Memb_Level <= " & svMembLevel _
         &     fIf(svMembCriteria <> "0", " AND CHARINDEX(Memb_Criteria, '" & svMembCriteria & "') > 0", "") _ 
         & " ORDER BY "_ 
         & "   Memb_FirstName, Memb_LastName"
'   sDebug

    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      fMemb_List = fMemb_List & oRs("Memb_FirstName") & " " & oRs("Memb_LastName") & "~" & oRs("Memb_No")  & "~" & oRs("Memb_Email") & "~~"
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
    '...strip trailing "~~"
    If Len(fMemb_List) > 3 Then fMemb_List = Left(fMemb_List, Len(fMemb_List)-2)
  End Function


  '...Get Member RecordSet for Users.asp/UsersOk.asp/VuAssess_Examinees
  '   currently only sorting by "last"
  '...on May 2, 2017 added the fTotalMins function which is more accurate then the older Memb_NoHours (actually minutes)
  Sub sGetMemb_Rs (vCustId, vWhere, vGlobal)


    '...big admin can see all learners
    If vGlobal = 1 And svMembLevel = 5 Then
'     vSql = "SELECT Memb.*, dbo.fTotalMins(Memb_No) AS Memb_NoMins FROM Memb WITH (NOLOCK) "
      vSql = "SELECT Memb.*, dbo.fTotalMins(Memb_No) AS Memb_NoMins, Cust_Id FROM Memb WITH (NOLOCK) INNER JOIN Cust WITH (NOLOCK) ON Memb.Memb_AcctId = Cust.Cust_AcctId AND Cust.Cust_Active = 1 "
      If Len(vWhere) > 0 Then vSql = vSql & "WHERE" & Mid(vWhere, 5) '...skip the " AND "
    '...big mgr can see all learners in their account set


    ElseIf vGlobal = 1 And svMembManager Then
      vSql = "SELECT Memb.*, dbo.fTotalMins(Memb_No) AS Memb_NoMins, Cust_Id FROM Memb WITH (NOLOCK) INNER JOIN Cust WITH (NOLOCK) ON Memb.Memb_AcctId = Cust.Cust_AcctId "
      vSql = vSql & "WHERE (LEFT(Cust.Cust_Id, 4) = '" & Left(svCustId, 4) & "')" & vWhere


    Else
      vSql = "SELECT Memb.*, dbo.fTotalMins(Memb_No) AS Memb_NoMins FROM Memb WITH (NOLOCK) WHERE (Memb_AcctId = '" & vCustId & "')" & vWhere
    End If
    vSql = vSql & " ORDER BY ISNULL(Memb_LastName,'') + ISNULL(Memb_FirstName,'') + CAST(Memb_No AS varchar(10))"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
  End Sub

  Sub sExtractMemb
    vMemb_No            = Request.Form("vMemb_No")
    vMemb_Id            = Ucase(Trim(Request.Form("vMemb_Id")))
    vMemb_Pwd           = Ucase(Trim(Request.Form("vMemb_Pwd")))
    vMemb_FirstName     = fUnquote(Request.Form("vMemb_FirstName"))
    vMemb_LastName      = fUnquote(Request.Form("vMemb_LastName"))
    vMemb_Email         = Trim(fUnquote(Request.Form("vMemb_Email")))
    vMemb_Active        = Request.Form("vMemb_Active")
    vMemb_Internal      = Request.Form("vMemb_Internal")
    vMemb_FirstVisit    = Request.Form("vMemb_FirstVisit")
    vMemb_Expires       = Request.Form("vMemb_Expires")
    vMemb_Level         = fDefault(Request.Form("vMemb_Level"), 2)
    vMemb_Auth          = Request.Form("vMemb_Auth")
    vMemb_MyWorld       = Request.Form("vMemb_MyWorld")
    vMemb_LCMS          = Request.Form("vMemb_LCMS")
    vMemb_Ecom          = Request.Form("vMemb_Ecom")
    vMemb_Channel       = Request.Form("vMemb_Channel")
    vMemb_VuBuild       = Request.Form("vMemb_VuBuild")
    vMemb_Manager       = Request.Form("vMemb_Manager")
    vMemb_Programs      = Ucase(Trim(Request.Form("vMemb_Programs")))
    vMemb_ProgramsAdded = Request.Form("vMemb_ProgramsAdded")
    vMemb_EcomG2alert   = fDefault(Request.Form("vMemb_EcomG2alert"), 1)
    vMemb_Duration      = fDefault(Request.Form("vMemb_Duration"),0)
    vMemb_Jobs          = Ucase(Trim(Request.Form("vMemb_Jobs")))
    vMemb_Memo          = fUnquote(Request.Form("vMemb_Memo"))
    vMemb_Organization  = fUnquote(Request.Form("vMemb_Organization"))
    vMemb_VuNews        = fDefault(Request.Form("vMemb_VuNews"),0)
    vMemb_Criteria      = fDefault(Trim(Replace(Request.Form("vMemb_Criteria"), ",", "")), "0")
    vMemb_Group2        = fDefault(Request.Form("vMemb_Group2"), 0)

    '...remove "All" if with other groups, ie "0 234 1234" becomes "234 1234"
    If Len(vMemb_Criteria) > 2 Then
      If Left(vMemb_Criteria, 2) = "0 " Then
        vMemb_Criteria = Mid(vMemb_Criteria, 3)
      End If
    End If

  End Sub


	'...note: this assumes the vMembId has been trimmed and ucased before here
	Function fMembIdOk (vMembId)	
		fMembIdOk = False
		If Len(vMembId) < 4 Then Exit Function
		Dim i, j
'		Const k = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_.-@!#$%^&*()+"
' 	Const k = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@$%^*()_-{}[];<>,.:"  '...extended values (removed + option on Sep 26, 2016 but not posted live until Nov 2016
		Const k = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@$%^*()_-{}[];,.:"  '...extended values (removed < and > option on Feb 08, 2017

		For i = 1 To Len(vMembId)
			j = Mid(vMembId, i, 1)
			If Instr(k, j) = 0 Then Exit Function
		Next		
		fMembIdOk = True				
	End Function

  '...get the current fields from the current record in the record set
  Sub sReadMemb
    vMemb_AcctId        = oRs("Memb_AcctId")
    vMemb_Id            = oRs("Memb_Id")
    vMemb_Pwd           = oRs("Memb_Pwd")
    vMemb_No            = oRs("Memb_No")
    vMemb_FirstName     = oRs("Memb_FirstName")
    vMemb_LastName      = oRs("Memb_LastName")
    vMemb_Email         = Trim(oRs("Memb_Email"))
    vMemb_Level         = oRs("Memb_Level")

    vMemb_Criteria      = oRs("Memb_Criteria")
    vMemb_Group2        = oRs("Memb_Group2")
    vMemb_Group3        = oRs("Memb_Group3")
    vMemb_JobsNo        = oRs("Memb_JobsNo")
    vMemb_Skills        = oRs("Memb_Skills")
    vMemb_Memo          = fOkValue(Trim(oRs("Memb_Memo")))
    vMemb_Organization  = fOkValue(Trim(oRs("Memb_Organization")))
    vMemb_FirstVisit    = oRs("Memb_FirstVisit")
    vMemb_LastVisit     = oRs("Memb_LastVisit")
    vMemb_NoVisits      = oRs("Memb_NoVisits")
    vMemb_NoHours       = oRs("Memb_NoHours")
    vMemb_Expires       = oRs("Memb_Expires")   
    vMemb_Online        = oRs("Memb_Online")   
    vMemb_Active        = oRs("Memb_Active")   
    vMemb_Internal      = oRs("Memb_Internal")   
    vMemb_Browser       = oRs("Memb_Browser")   
    vMemb_Programs      = fOkValue(oRs("Memb_Programs"))

    vMemb_ProgramsAdded = oRs("Memb_ProgramsAdded")
    vMemb_EcomG2alert   = oRs("Memb_EcomG2alert")

    vMemb_Duration      = oRs("Memb_Duration")
    vMemb_Jobs          = fOkValue(oRs("Memb_Jobs"))
    vMemb_Sponsor       = oRs("Memb_Sponsor")
    vMemb_MaxSponsor    = oRs("Memb_MaxSponsor")
    vMemb_VuNews        = oRs("Memb_VuNews")

		If oRs("Memb_Level") < 4 Then
			vMemb_Auth         = False
			vMemb_MyWorld      = False
			vMemb_LCMS         = False
			vMemb_Ecom         = False
			vMemb_Channel      = False
			vMemb_VuBuild      = False
			vMemb_Manager      = False
		Else
			vMemb_Auth         = oRs("Memb_Auth")
			vMemb_MyWorld      = oRs("Memb_MyWorld")
			vMemb_LCMS         = oRs("Memb_LCMS")
			vMemb_Ecom         = oRs("Memb_Ecom")
			vMemb_Channel      = oRs("Memb_Channel")
			vMemb_VuBuild      = oRs("Memb_VuBuild")
			vMemb_Manager      = oRs("Memb_Manager")
		End If

		vMemb_Guid           = oRs("Memb_Guid")
		vMemb_LastAssignedBy = oRs("Memb_LastAssignedBy")

  End Sub


  Sub sGetMemb (vMembNo)
    If IsNumeric(fOkValue(vMembNo)) Then
      vMemb_Eof = True
      vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE Memb_No = " & vMembNo
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then 
        sReadMemb
        vMemb_Eof = False
      End If
      Set oRs = Nothing      
      sCloseDb
    Else 
      vMemb_Eof = False
    End If
  End Sub      


  Sub sGetMembById (vAcctId, vMembId)
    vMemb_Eof = True
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub      


  Sub sGetMembByPwd (vAcctId, vMembId, vMembPwd)
    vMemb_Eof = True
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "' AND Memb_Pwd = '" & vMembPwd & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub      


  Sub sGetMembByOrganization (vAcctId, vMembOrganization)
    vMemb_Eof = True
    vSql = "SELECT TOP 1 * FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Organization = '" & vMembOrganization & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub      


  Sub sGetMembByEmail (vAcctId, vMembEmail)
    vMemb_Eof = True
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Email = '" & vMembEmail & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub      


  Sub sGetMembByIdAndLastName (vAcctId, vMembId, vMembLastName)
    vMemb_Eof = True
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "' AND UPPER(Memb_LastName) = '" & Ucase(fUnQuote(vMembLastName)) & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub


  Sub sGetMembByIdAndMemo (vAcctId, vMembId, vMembMemo)
    vMemb_Eof = True
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "' AND UPPER(Memb_Memo) = '" & Ucase(fUnQuote(vMembMemo)) & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    End If
    Set oRs = Nothing      
    sCloseDb
  End Sub


  '...note must only have one email on this account to be valid
  Function fMembIdByEmail (vAcctId, vMembEmail)
    sOpenDb
    vSql = "SELECT Top 1 Memb_Id FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "'"
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      fMembIdByEmail = "AcctId"
    Else
      vSql = "SELECT Count(Memb_Id) AS [Count] FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Email = '" & vMembEmail & "'"
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then 
        If oRs("Count") = 0 Then
          fMembIdByEmail = "None"        
        ElseIf oRs("Count") > 1 Then
          fMembIdByEmail = "Multiple"
        Else
          vSql = "SELECT Memb_Id FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Email = '" & vMembEmail & "'"
          Set oRs = oDb.Execute(vSql)
          fMembIdByEmail = oRs("Memb_Id")       
        End If       
      End If
    End If
    Set oRs = Nothing      
    sCloseDb
  End Function     

  
  Sub sUpdateMemb_Profile
    vSql = "UPDATE Memb SET"
    vSql = vSql & " Memb_FirstName  = '" & fUnquote(vMemb_FirstName)  & "', " 
    vSql = vSql & " Memb_LastName   = '" & fUnquote(vMemb_LastName)   & "', " 
    vSql = vSql & " Memb_Pwd        = '" & fUnquote(vMemb_Pwd)        & "', " 
    vSql = vSql & " Memb_Email      = '" & fUnquote(vMemb_Email)      & "', " 
    vSql = vSql & " Memb_VuNews     =  " & fSqlBoolean(vMemb_VuNews) & "   " 
    vSql = vSql & " WHERE Memb_No   =  " & vMemb_No
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
    sUpdateMemb_Session
  End Sub

 
  '...using when job no, note deletes any programs that might have been previously selected
  Sub sUpdateMembJobsNo (vMemb_No, vJobsNo)
    vSql = "UPDATE Memb SET Memb_Programs = '', Memb_JobsNo = " & vJobsNo & " WHERE Memb_No   =  " & vMemb_No
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


   '...using when skills ratings
   Sub sUpdateMembSkills (vMemb_No, vMembSkills)
    vSql = "UPDATE Memb SET Memb_Skills = '" & vMembSkills & "' WHERE Memb_No   =  " & vMemb_No
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

 
  Sub sUpdateMembPrograms (vMemb_No, vPrograms)
    vSql = "UPDATE Memb SET Memb_Programs = '" & vPrograms & "' WHERE Memb_No   =  " & vMemb_No
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sUpdateMembMemo(vMemb_No, vMemb_Memo)
    vSql = "UPDATE Memb SET Memb_Memo= '" & vMemb_Memo & "' WHERE Memb_No   =  " & vMemb_No
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sUpdateMemb_Session
   '...update session and local variables if current user  
    If Clng(fOkValue(vMemb_No)) = svMembNo Then
      If Len(vMemb_FirstVisit) = 0 Then vMemb_FirstVisit = svMembFirstVisit '...edit on "home.asp" does not update this value
      Session("MembFirstName")  = vMemb_FirstName
      Session("MembLastName")   = vMemb_LastName
      Session("MembPwd")        = vMemb_Pwd
      Session("MembEmail")      = vMemb_Email
      Session("MembFirstVisit") = vMemb_FirstVisit
      Session("MembCriteria")   = vMemb_Criteria
      svMembFirstName           = vMemb_FirstName
      svMembLastName            = vMemb_LastName
      svMembEmail               = vMemb_Email    
      svMembFirstVisit          = vMemb_FirstVisit
      svMembCriteria            = vMemb_Criteria
    End If
  End Sub

  
  Sub sDeleteMemb
    sOpenDb
    vSql = "DELETE FROM Memb WHERE Memb_No = " & vMemb_No
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sInactivateMemb
    sOpenDb
    vSql = "UPDATE Memb SET Memb_Active = 0 WHERE Memb_No = " & vMemb_No
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...this returns a space separated list of Accounts containing this ID
  Function fMembIdAll (vMembId)
    sOpenDb3
    vSql = "SELECT Memb_AcctId FROM Memb WITH (NOLOCK) WHERE Memb_Id = '" & vMembId & "'"
    Set oRs3 = oDb3.Execute(vSql)
    While Not oRs3.Eof
      fMembIdAll = fMembIdAll & " " & oRs3("Memb_AcctId")
      oRs3.MoveNext
    Wend
    Set oRs3 = Nothing      
    sCloseDb3
  End Function


  '...this deletes this member id from ALL accounts
  Sub sDeleteMembAllById (vMembId)
    sOpenDb
    vSql = "DELETE FROM Memb WHERE Memb_Id = '" & vMembId & "'"
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...this inactivate this member id from ALL accounts
  Sub sInactivateMembAllById (vMembId)
    sOpenDb
    vSql = "UPDATE Memb SET Memb_Active = 0 WHERE Memb_Id = '" & vMembId & "'"
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Function fAllMembCount
    sOpenDb2
    vSql = "SELECT COUNT (Memb_No) AS MembCount FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id NOT LIKE '" & vPasswordx & "%' AND Memb_Id NOT LIKE '" & Left(svCustId, 4) & "SALES'"
    Set oRs2 = oDb2.Execute(vSql)
'   sDebug
    fAllMembCount = oRs2("MembCount")
    Set oRs2 = Nothing
    sCloseDb2
  End Function


  Function fMembCount (vAcctId)
    sOpenDb2
    vSql = "SELECT COUNT (Memb_No) AS MembCount FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Level = 2"
    Set oRs2 = oDb2.Execute(vSql)
'   sDebug
    fMembCount = oRs2("MembCount")
    Set oRs2 = Nothing
    sCloseDb2
  End Function



  
  Function fMembName (vMembNo)
    fMembName = ""
    If Len(vMembNo) > 0 Then
      If VarType(vMembNo) = vbString And IsNumeric(vMembNo) Then 
        vMembNo = cLng(vMembNo)
      ElseIf Vartype(vMembNo) <> vbLong Then
        Exit Function
      End If
    Else
      Exit Function
    End If  
    vSql = "SELECT Memb_FirstName, Memb_LastName FROM Memb WITH (NOLOCK) WHERE Memb_No = " & vMembNo
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fMembName = oRs2("Memb_FirstName") & " " & oRs2("Memb_LastName")
    Set oRs2 = Nothing      
    sCloseDb2
  End Function   



  Function fMembDropdown (vMembNo) 
    Dim vCurrentNo, vSelected, vMembName, vOk
    If Len(vMembNo) > 0 Then
      If VarType(vMembNo) = vbString And IsNumeric(vMembNo) Then 
        vMembNo = cLng(vMembNo)
      ElseIf Vartype(vMembNo) <> vbLong Then
        vMembNo = 0
      End If
    Else
      vMembNo = 0
    End If     
    fMembDropDown = vbCrLf
    vSql = "SELECT Memb_No, Memb_FirstName, Memb_LastName, Memb_Criteria FROM Memb WITH (NOLOCK) Where Memb_AcctId = " & svCustAcctId
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vMemb_No       = oRs("Memb_No")
      vMembName      = oRs("Memb_FirstName") & " " & oRs("Memb_LastName")
      vMemb_Criteria = oRs("Memb_Criteria")
      If Len(Trim(vMembName)) > 0 Then 
        vOk = True
      Else
        vOk = False
      End If      
      '...ensure current user has criteria access 
      If svMembCriteria <> "0" And vMemb_Criteria <> svMembCriteria Then
        vOk = False
      End If
      If vOk Then
        If vMemb_No = vMembNo Then
          vSelected = " SELECTED" 
        Else
          vSelected = ""
        End If
        i = "          <option value=" & Chr(34) & vMemb_No & Chr(34) & vSelected & ">" & vMembName & "</option>" & vbCrLf
        fMembDropdown = fMembDropdown & i
      End If
      oRs.MoveNext	        
    Loop
    Set oRs = Nothing      
    sCloseDb
'    '...save the current Member No
'    vMembNo = vCurrentNo
  End Function   



  Function fMembFacsDropdown (vMembNo)
    Dim vSelected
    vMembNo = fPureInt(vMembNo)
    fMembFacsDropdown = "" 
    vSql = "SELECT Memb_No, Memb_Id, Memb_FirstName, Memb_LastName FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = " & svCustAcctId & " AND Memb_Level = 3 AND Memb_Internal = 0 ANd Memb_Active = 1"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vSelected = fIf(oRs("Memb_No") = vMembNo, " SELECTED", "") 
      fMembFacsDropdown = fMembFacsDropdown & "          <option value=" & Chr(34) & oRs("Memb_No") & Chr(34) & vSelected & ">" & oRs("Memb_FirstName") & " " & oRs("Memb_LastName") & " [" & oRs("Memb_Id") & "]</option>" & vbCrLf
      oRs.MoveNext	        
    Loop
    Set oRs = Nothing      
    sCloseDb
  End Function   


  Function fMembRegister (svCustAcctId)
    fMembRegister = ""
    '...all fields submitted
    If fNoValue(vMemb_Email) Or fNoValue(vMemb_FirstName) Or fNoValue(vMemb_LastName) Then
      vMsg = "All fields must be entered correctly."
      Exit Function
    End If
    '...valid email address          
    i = Instr(vMemb_Email, "@")
    If i < 2 Or i > Len(vMemb_Email) - 2 Then
      vMsg = "Invalid Email Address."
      Exit Function
    End If
    '...ensure there's no one using that email
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE (Memb_AcctId = '" & svCustAcctId & "' AND Memb_Email = '" & vMemb_Email & "')" 
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then
      sCloseDb
      vMsg = "That Email Address is already on file."
      Exit Function
    End If
    sCloseDb

    sAddMemb svCustAcctId

    If vFileOk Then 
      fMembRegister = vMemb_Id
      vMsg = ""
    Else
      vMsg = "Unable to issue you a password.  Please email us for help."
    End If

  End Function  

  
  '...this will redirect group facilitators to UserGroup.asp rather than User.asp for limited functionality
  Function fGroup
    fGroup = ""
    vSql = "SELECT TOP 1 Ecom.Ecom_Media FROM Cust INNER JOIN Ecom ON Cust.Cust_AcctId = Ecom.Ecom_NewAcctId WHERE (Ecom_Archived IS NULL) AND (Cust.Cust_AcctId = '" & svCustAcctId & "')"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then
      If oRs2("Ecom_Media") = "Group2" Then fGroup = "Group"
    End If
    Set oRs2 = Nothing
    sCloseDb2
  End Function


  Function fMembEcomBypass (vAcctId, vMembId)
    fMembEcomBypass = False
    vSql = "SELECT Memb_Ecom FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof And oRs("Memb_Ecom") Then fMembEcomBypass = True
    Set oRs = Nothing      
    sCloseDb
  End Function


  '...Get MembNo, return 0 if not on file
  Function fMembNo (vCustId, vMembId)
    fMembNo = 0
    vSql = "SELECT Memb_No FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vCustId & "' AND Memb_Id= '" & vMembId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      fMembNo = oRs("Memb_No")
    End If
    Set oRs = Nothing
    sCloseDb    
  End Function


   '...add internal passwords when creating/cloning a new site (Db_Cust, DB_QCust)
  Sub sAddInternalMemb (vAcctId)

    Dim vAcctIdsave : vAcctIdsave = vAcctId      '...save as this might get wiped out in the sMemb_Empty 
    sMemb_Empty                                  '...clean out all fields
    vMemb_AcctId    = vAcctIdsave
    vMemb_Internal  = 1
    vMemb_FirstName = "Vubiz"
    vMemb_LastName  = "Internal"    
    vMemb_Pwd       = "VUB!Z"
    vMemb_No = 0 : vMemb_Id = vPassword5 : vMemb_Level = 5 : sAddMemb vAcctId
    vMemb_No = 0 : vMemb_Id = vPassword4 : vMemb_Level = 4 : sAddMemb vAcctId
    vMemb_No = 0 : vMemb_Id = vPassword3 : vMemb_Level = 3 : sAddMemb vAcctId
    vMemb_No = 0 : vMemb_Id = vPassword2 : vMemb_Level = 2 : sAddMemb vAcctId
  End Sub



  '..............  Sponsors  ........................

  Sub sExtractSponsors
    vMemb_FirstName    = fUnquote(Request("vFirstName"))
    vMemb_LastName     = fUnquote(Request("vLastName"))
    vMemb_Email        = fUnquote(Request("vEmail"))
  End Sub


  Sub sAddSponsors (vMembFirstName, vMembLastName, vMembEmail, vSponsorNo)    
    fNextMembNo (svCustAcctId)
    vMemb_AcctId     = svCustAcctId
    vMemb_FirstName  = vMembFirstName
    vMemb_LastName   = vMembLastName
    vMemb_Email      = vMembEmail
    vMemb_Sponsor    = vSponsorNo
    vMemb_Expires    = fFormatSqlDate(Now + 90)
    vMemb_FirstVisit = Now
    vMemb_NoVisits   = 0
    sUpdateMemb svCustAcctId
  End Sub


  Sub sGetSponsors (vSponsorNo)
    vSql = "SELECT * FROM Memb WITH (NOLOCK) WHERE (Memb_AcctId = '" & svCustAcctId & "') AND Memb_Sponsor = " & vSponsorNo & " ORDER BY Memb_FirstVisit DESC"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
  End Sub


  Sub sInactivateSponsor (vMembNo)
    vSql = "UPDATE Memb SET Memb_Active = 0 WHERE Memb_No   =  " & vMembNo
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sActivateSponsor (vMembNo)
    vSql = "UPDATE Memb SET Memb_Active = 1 WHERE Memb_No   =  " & vMembNo
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sExtendSponsor (vMembNo, vMembExpires)
    vSql = "UPDATE Memb SET Memb_Expires = '" & vMembExpires & "' WHERE Memb_No   =  " & vMembNo
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sMaxSponsor (vMembNo, vMaxSponsor)
    vSql = "UPDATE Memb SET Memb_MaxSponsor = " & vMaxSponsor & " WHERE Memb_No =  " & vMembNo
    sOpenDb 
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Function fNoSponsors (vMembNo)
    fNoSponsors = 0
    vSql = "SELECT COUNT(*) As [MaxSponsors] FROM Memb WITH (NOLOCK) WHERE Memb_Sponsor =  " & vMembNo & " AND Memb_Expires > '" & NOW & "'"
    sOpenDb 
'   sDebug
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fNoSponsors = Cint(oRs("MaxSponsors"))
    sCloseDb
    Set oRs = Nothing    
  End Function


  '...returns sponsored learners' details for User.asp
  Function fSponsoredList (vMembSponsor)
    fSponsoredList = ""
    vSql = "SELECT Memb_No, Memb_FirstName, Memb_LastName, Memb_Expires, Memb_Active FROM Memb WITH (NOLOCK) WHERE Memb_Sponsor = " & vMembSponsor
    sOpenDb 
'   sDebug
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      fSponsoredList = fSponsoredList & oRs("Memb_No") & "|" & oRs("Memb_FirstName") & "|" & oRs("Memb_LastName") & "|" & oRs("Memb_Expires") & "|" & oRs("Memb_Active") & "~"
      oRs.MoveNext
    Loop 
    If Len(fSponsoredList) > 0 Then fSponsoredList = Left(fSponsoredList, Len(fSponsoredList) - 1)
    sCloseDb
    Set oRs = Nothing    
  End Function


  '...returns sponsor details for User.asp
  Function fSponsorList (vMembSponsor)
    fSponsorList = ""
    vSql = " SELECT Memb_1.Memb_FirstName AS Sponsor_FirstName, Memb_1.Memb_LastName AS Sponsor_LastName" _
         & " FROM Memb WITH (NOLOCK) INNER JOIN Memb Memb_1 WITH (NOLOCK) ON Memb.Memb_Sponsor = Memb_1.Memb_No" _
         & " WHERE (Memb.Memb_Sponsor = " & vMembSponsor & ")"
    sOpenDb 
'   sDebug
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fSponsorList = oRs("Sponsor_FirstName") & "|" & oRs("Sponsor_LastName") 
    sCloseDb
    Set oRs = Nothing    
  End Function


  Function fNextSponsorDate (vMembNo)
    fNextSponsorDate = " "
    vSql = "SELECT MIN(Memb_Expires) AS [NextDate] FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Sponsor =  " & vMembNo & " AND Memb_Expires > '" & NOW & "'"
    sOpenDb 
'   sDebug
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fNextSponsorDate = oRs("NextDate")
    sCloseDb
    Set oRs = Nothing    
  End Function


  Function fModifiedBy (vMembNo) 
    sGetMemb vMembNo
    If vMemb_Eof Then
      fModifiedBy = "[User is not on file]"
    Else
      If vMemb_Level < svMemb_Level Then 
        fModifiedBy = vFirstName & " " & vLastName & "(" & vMemb_Id & ")"
      Else
        fModifiedBy = vFirstName & " " & vLastName
      End If    
    End If
  End Function 


  '...this is used to generate a vubiz id when there is no memb novMemb_No
  Function fNextMembNo (vAcctId)
    vMemb_Id      = "Temp_" & Now()
    fNextMembNo   = spMembNoById (vAcctId, vMemb_Id, svMembNo)
    vMemb_Id      = fMemb_Id(vAcctId, vMemb_No)
    sUpdateMembByNo vAcctId
  End Function
  
  '...this is used to insert a group id returning the memb no
  Function fNextMembNo2 (vAcctId, vId)
    vMemb_Id      = vId
    fNextMembNo2  = spMembNoById (vAcctId, vMemb_Id, svMembNo)
    sUpdateMembByNo vAcctId
  End Function

  '...this adds/updates memb when there is a vMemb_No
  Sub sAddMemb (vAcctId)
    sUpdateMemb vAcctId
  End Sub


  '...if we have a valid vMemb_No then sUpdateMembByNo
  '   else if we have a valid vMemb_Id then update 
  '   else create a Memb_Id 
  Sub sUpdateMemb (vAcctId)
    Dim bOk : bOk = False
    If IsNumeric(fOkValue(vMemb_No)) Then 
      bOk = True
    End If
    If bOk And Clng(vMemb_No) > 0 Then 
      sUpdateMembByNo vAcctId
      Exit Sub
    End If
    If Len(fOkValue(vMemb_Id)) = 0 Then
      fNextMembNo (vAcctId)
      Exit Sub
    Else
      vMemb_No = spMembNoById (vAcctId, vMemb_Id, svMembNo)
      sUpdateMembByNo vAcctId
    End If
  End Sub


  Sub sUpdateMembByNo (vAcctId)
    vMemb_AcctId  = vAcctId
    sTableUpdate "V5_Vubz", "Memb", vMemb_No
    sUpdateMemb_Session
  End Sub


  '...this grabs the recordset for this id
  '   if not on file it will create an empty rs and return all fields for the update routine (in initialize.asp)
  Function spMembNoById (vAcctId, vMembId, vMembAlteredBy)
    Dim oRs
    vMembAlteredBy = fDefault(vMembAlteredBy, 0)
    sOpenCmd
    With oCmd
      .CommandText = "spMembNoById"
      .Parameters.Append .CreateParameter("@Memb_AcctId",    		 adVarChar, adParamInput,   6, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Id",        		 adVarChar, adParamInput, 128, vMembId)
      .Parameters.Append .CreateParameter("@Memb_AlteredBy", 		 adInteger, adParamInput,   0, vMembAlteredBy)
    End With
    Set oRs = oCmd.Execute()
    spMembNoById = oRs("Memb_No")
    vMemb_No = spMembNoById
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...this grabs the id for the memb_no
  '   if not on file it will create an empty rs and return all fields for the update routine (in initialize.asp)
  Function spMembIdByNo (vMembNo)
    Dim oRs
    sOpenCmd
    With oCmd
      .CommandText = "spMembIdByNo"
      .Parameters.Append .CreateParameter("@Memb_No", adInteger, adParamInput, 0, vMembNo)
    End With
    Set oRs = oCmd.Execute()
    spMembIdByNo = oRs("Memb_Id")
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...this returns all fields for this recordset which is then used for the update routine (db_Update.asp)
  Function spMembByNo (vMembNo)
    Dim oRs
    sOpenCmd
    With oCmd
      .CommandText = "spMembByNo"
      .Parameters.Append .CreateParameter("@Memb_No",        		 adInteger, adParamInput,   0, vMembNo)
    End With
    Set oRs = oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...this returns a temp GUID for access to Portal (Added Dec 12 2017)
  Function sp5getMembGuidTemp (membGuid)
    Dim oRsApp
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp5getMembGuidTemp"
      .Parameters.Append .CreateParameter("@membGuid", adVarChar, adParamInput, 38, membGuid)
    End With
    Set oRsApp = oCmdApp.Execute()
    sp5getMembGuidTemp = oRsApp("guidTemp")
    Set oCmdApp = Nothing
    sCloseDbApp
  End Function


  '...Quick count the Active/Inactive/All users within an acct
  Sub spMembCount (vAcctId, i, j, k)
    Dim oRs
    sOpenCmd
    With oCmd
      .CommandText = "spMembCount"
      .Parameters.Append .CreateParameter("@Memb_AcctId",    		adVarChar, adParamInput, 04, vAcctId)
    End With
    Set oRs = oCmd.Execute()
    i = oRs("Active")
    j = oRs("InActive")
    k = oRs("All")
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...Count the Active/Inactive/All users within an acct including just learners
  Sub spMembCountAll (vAcctId, act2, ina2, act3, ina3, act4, ina4)
    Dim oRs
    sOpenCmd
    With oCmd
      .CommandText = "spMembCountAll"
      .Parameters.Append .CreateParameter("@Memb_AcctId",    		adVarChar, adParamInput, 04, vAcctId)
    End With
    Set oRs = oCmd.Execute()
    ina2 = oRs("ina2")
    act2 = oRs("act2")
    ina3 = oRs("ina3")
    act3 = oRs("act3")
    ina4 = oRs("ina4")
    act4 = oRs("act4")
    Set oRs  = Nothing
    Set oCmd = Nothing
    sCloseDb
  End Sub


  '...See if this user is on file (returns true or false)
  Function spMembExistsByNo (vNo)
    sOpenCmd
    With oCmd
      .CommandText = "spMembExistsByNo"
      .Parameters.Append .CreateParameter("RETURN_VALUE",   adInteger, adParamReturnValue,   , Null)
      .Parameters.Append .CreateParameter("@Memb_No",  	    adInteger, adParamInput,        0, vNo)
    End With
    oCmd.Execute()
    spMembExistsByNo = oCmd.Parameters(0)
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...See if this user id is on file (returns true or false)
  Function spMembExistsById (vAcctId, vId)
    sOpenCmd
    With oCmd
      .CommandText = "spMembExistsById"
      .Parameters.Append .CreateParameter("RETURN_VALUE",   adInteger, adParamReturnValue,   , Null)
      .Parameters.Append .CreateParameter("@Memb_AcctId",  	adVarChar, adParamInput,        4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Id",    		adVarChar, adParamInput,      128, vId)
    End With
    oCmd.Execute()
    spMembExistsById = oCmd.Parameters(0)
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...See if this user id/pwd is on file (returns true or false)
  Function spMembExistsByIdPwd (vAcctId, vId, vPwd)
    sOpenCmd
    With oCmd
      .CommandText = "spMembExistsByIdPwd"
      .Parameters.Append .CreateParameter("RETURN_VALUE",   adInteger, adParamReturnValue,   , Null)
      .Parameters.Append .CreateParameter("@Memb_AcctId",  	adVarChar, adParamInput,        4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Id",    		adVarChar, adParamInput,      128, vId)
      .Parameters.Append .CreateParameter("@Memb_Pwd",    	adVarChar, adParamInput,      128, vPwd)
    End With
    oCmd.Execute()
    spMembExistsByIdPwd = oCmd.Parameters(0)
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...See if this user name is on file (returns true or false)
  Function spMembExistsByName (vAcctId, vFirstName, vLastName)
    sOpenCmd
    With oCmd
      .CommandText = "spMembExistsByName"
      .Parameters.Append .CreateParameter("RETURN_VALUE",     adInteger, adParamReturnValue,   , Null)
      .Parameters.Append .CreateParameter("@Memb_AcctId",  	  adVarChar, adParamInput,        4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_FirstName",  adVarChar, adParamInput,       32, vFirstName)
      .Parameters.Append .CreateParameter("@Memb_LastName",   adVarChar, adParamInput,       64, vLastName)
    End With
    oCmd.Execute()
    spMembExistsByName = oCmd.Parameters(0)
    Set oCmd = Nothing
    sCloseDb
  End Function


  '...Create / Update Learners via Upload Advanced
  '   update by acctid/id and assume ACTIVE = 1
  '   NOTE: you need to open and close before using this function
  Function spMembUpload (vAcctId, vId, vCriteria, vFirstName, vLastName, vEmail, vPwd, vPrograms, vMemo, vJobs, vGroup3)
  	sOpenCmd
    With oCmd
      .CommandText = "spMembUpload "
      .Parameters.Append .CreateParameter("@Memb_AcctId",  	  adVarChar, adParamInput,        4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Id",  	  		adVarChar, adParamInput,      128, vId)
      .Parameters.Append .CreateParameter("@Memb_Criteria",  	adInteger, adParamInput,         , vCriteria)
      .Parameters.Append .CreateParameter("@Memb_FirstName",  adVarChar, adParamInput,       32, vFirstName)
      .Parameters.Append .CreateParameter("@Memb_LastName",   adVarChar, adParamInput,       64, vLastName)
      .Parameters.Append .CreateParameter("@Memb_Email",  	  adVarChar, adParamInput,      128, vEmail)
      .Parameters.Append .CreateParameter("@Memb_Pwd",  	  	adVarChar, adParamInput,       64, vPwd)
      .Parameters.Append .CreateParameter("@Memb_Programs",  	adVarChar, adParamInput,     8000, vPrograms)
      .Parameters.Append .CreateParameter("@Memb_Memo",  	  	adVarChar, adParamInput,      512, vMemo)
      .Parameters.Append .CreateParameter("@Memb_Jobs",  	  	adVarChar, adParamInput,     8000, vJobs)
      .Parameters.Append .CreateParameter("@Memb_Group3",     adInteger, adParamInput,         , vGroup3)
      .Parameters.Append .CreateParameter("@Memb_AlteredBy",  adInteger, adParamInput,         , svMembNo)
    End With
    oCmd.Execute()
	  Set oCmd = Nothing
	  sCloseDb
  End Function



  Sub spMembByOrganization (vAcctId, vOrganization)
    sOpenCmd
    With oCmd
      .CommandText = "spMembByOrganization"
      .Parameters.Append .CreateParameter("@Memb_AcctId",    		adVarChar, adParamInput,   4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Organization", adVarChar, adParamInput, 128, vOrganization)
    End With
    Set oRs   = oCmd.Execute()
    If Not oRs.Eof Then 
      sReadMemb
      vMemb_Eof = False
    Else
      vMemb_Eof = True
    End If
    Set oRs   = Nothing
    Set oCmd  = Nothing
    sCloseDb
  End Sub


   Sub spMembInactivate (vAcctId, vDaysOk)
    sOpenCmd
    With oCmd
      .CommandText = "spMembInactivate"
      .Parameters.Append .CreateParameter("@Memb_AcctId", adVarChar, adParamInput,   4, vAcctId)
      .Parameters.Append .CreateParameter("@DaysOk",    	adInteger, adParamInput,    , vDaysOk)
    End With
    oCmd.Execute()
    Set oCmd  = Nothing
    sCloseDb
  End Sub


   Sub spMembActiveById (vAcctId, vMembId, vActive, vAlteredBy)
    sOpenCmd
    With oCmd
      .CommandText = "spMembActiveById"
      .Parameters.Append .CreateParameter("@Memb_AcctId",    		adVarChar, adParamInput,   4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Id",        		adVarChar, adParamInput, 128, vMembId)
      .Parameters.Append .CreateParameter("@Memb_Active", 			adBoolean, adParamInput,   0, vActive)
      .Parameters.Append .CreateParameter("@Memb_AlteredBy", 		adInteger, adParamInput,   0, vAlteredBy)
    End With
    oCmd.Execute()
    Set oCmd  = Nothing
    sCloseDb
  End Sub


   Sub spMembDeleteLearners (vAcctId)
    sOpenCmd
    With oCmd
      .CommandText = "spMembDeleteLearners"
      .Parameters.Append .CreateParameter("@Memb_AcctId",    		adVarChar, adParamInput,   4, vAcctId)
    End With
    oCmd.Execute()
    Set oCmd  = Nothing
    sCloseDb
  End Sub
  
  '...Create a unique Learner ID
  Function fMemb_Id (vAcctId, vMembNo)
    fMemb_Id = Right(100000000 + vMembNo, Len(vMembNo)) & "-" & fSecurityCode(vAcctId, vMembNo)
  End Function

  '...Extract MembNo from a unique Learner ID (used in Ecom6GeneratedId.asp)
  Function fMemb_No (vMembId)
    Dim aNo
    aNo = Split(vMembId, "-")
    If Ubound(aNo) = 0 Then
      fMemb_No = 0
    Else
      fMemb_No = fPureInt(aNo(0))
    End If   
  End Function

  '...Generate the security code (also in EcomGeneratedId.asp) 
  '   stopped using vAcctId June 7th 2016 after introducing negative and alpha acctids - just use membNo
  Function fSecurityCode (vAcctId, vMembNo)
    Dim vTemp, i, j, k
    Const cAlpha = "ABCDEFGHXY"
'   vTemp = vMembNo * 4141
'   vTemp = vMembNo * 4141 + vAcctId
    vTemp = vMembNo * 2112
    vTemp = Right("0000" & vTemp, 4)
    fSecurityCode = ""
    For i = 1 To 4
      j = mid(vTemp, i, 1)   
      k = mid(cAlpha, j+1, 1)
      fSecurityCode = fSecurityCode & k
    Next
  End Function

  '...use this for new table update before adding a new file so old data is not includedd in new updates
  Sub sMemb_Empty
    vMemb_AcctId        = Empty
    vMemb_Id            = Empty
    vMemb_Pwd           = Empty
    vMemb_No            = Empty
    vMemb_FirstName     = Empty
    vMemb_LastName      = Empty
    vMemb_Email         = Empty
    vMemb_Level         = Empty
    vMemb_FirstVisit    = Empty
    vMemb_LastVisit     = Empty
    vMemb_NoVisits      = Empty
    vMemb_Cust          = Empty
    vMemb_Criteria      = Empty
    vMemb_Group2        = Empty
    vMemb_JobsNo        = Empty
    vMemb_Skills        = Empty
    vMemb_Memo          = Empty
    vMemb_Organization  = Empty
    vMemb_NoHours       = Empty
    vMemb_Expires       = Empty
    vMemb_Active        = Empty
    vMemb_Internal      = Empty
    vMemb_Browser       = Empty
    vMemb_Programs      = Empty
    vMemb_ProgramsAdded = Empty
    vMemb_EcomG2Alert   = Empty
    vMemb_Duration      = Empty
    vMemb_Jobs          = Empty
    vMemb_MaxSponsor    = Empty
    vMemb_Sponsor       = Empty
    vMemb_VuNews        = Empty
    vMemb_Auth          = Empty
    vMemb_MyWorld       = Empty
    vMemb_LCMS          = Empty
    vMemb_Ecom          = Empty
    vMemb_Channel       = Empty
    vMemb_VuBuild       = Empty
    vMemb_Manager       = Empty
    vMemb_Guid          = Empty
    vMemb_LastAssignedBy= Empty
  
    vMemb_AlteredOn     = Empty
    vMemb_AlteredBy     = Empty
  End Sub


  '...this is used to sheild IDs for levels = or greater than current  
  Function fMembId (MembId, MembLevel)
    If MembLevel < svMembLevel Or svMembLevel = 5 Then 
      fMembId = fIf(IsNumeric(MembId) AND Left(MembId, 1) = "0", "'" & MembId, MembId)
    Else
      fMembId = "******"
    End If  
  End Function


  '...Get MembNo using the MembId (used for RTE)
  Function fMembNoById (vAcctId, vMembId)
    fMembNoById = 0
    vSql = "SELECT Memb_No FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "'" 
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fMembNoById = oRs("Memb_No")
    Set oRs = Nothing
    sCloseDb
  End Function


  '...Get Active Fac MembNo using the MembId (used for Upload2)
  Function fFacMembNoById (vAcctId, vMembId)
    fFacMembNoById = 0
    vSql = "SELECT Memb_No FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "' AND Memb_Level = 3 AND Memb_Internal = 0 AND Memb_Active = 1" 
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fFacMembNoById = oRs("Memb_No")
    Set oRs = Nothing
    sCloseDb
  End Function


  '...Get Active Fac MembNo using the MembId (used for Upload2)
  Function fV8Fields (vAcctId, vMembId)
    fV8Fields = ""
    vSql = "SELECT Memb_Parent, Memb_Catalogue FROM Memb WHERE Memb_AcctId = '" & vAcctId & "' AND Memb_Id = '" & vMembId & "'" 
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fV8Fields = oRs("Memb_Parent") & " | " & oRs("Memb_Catalogue")
    Set oRs = Nothing
    sCloseDb
  End Function

%>