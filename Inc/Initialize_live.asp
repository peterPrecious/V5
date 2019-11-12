<%
  '...stop if not secure and no bypass then timeout (note long url in case called from repository)
  If Not Session("Secure") And Not vBypassSecurity Then 
    Response.Redirect "/V5/Code/TimeOut.asp"
  End If

  '...force a longer/shorter session.timeout
  Dim vTempSession
  vTempSession = fDefault(Request.QueryString("vSession"), 0)
  If IsNumeric(vTempSession) Then
    If i > 0 Then Session.Timeout = Request.QueryString("vSession")
  End If

  '...declare local DB variables
  Dim oDb, oDb2, oDb3, oDb4, oDb5, oDbBase, oDbBase2, oDbGap, oDbGap2, oDbRTE
  Dim oRs, oRs2, oRs3, oRs4, oRs5, oRsBase, oRsBase2, oRsGap, oRsGap2, oRsRTE
  Dim oCmd, oCmdBase, oCmdGap, oCmdGap2, oCmdRTE
 
  '...declare local variables used for session variables
  Dim svServer, svHostDb, svHostDbPwd, svMailServer, svHost, svSQL, svSSL, svDomain, svFontNo
  Dim svCustId, svCustAcctId, svCustBanner, svCustUrl, svCustEmail, svCustProcess, svCustTitle, svCustReturnUrl, svCustCluster, svCustLevel, svCustPwd
  Dim svCustFreeHours, svCustFreeDays, svCustAuto, svCustIssueIds, svCustActivateIds, svCustIdsSize, svCustExpires, svCustMaxHours, svCustEcomDisc
  Dim svMembId, svMembPwd, svMembNo, svMembFirstName, svMembLastName, svMembLevel, svMembFirstVisit, svMembLastVisit, svMembNoVisits, svMembNoHours, svMembExpires, svMembEmail, svMembCriteria, svMembManager, svMembGap, svMembInternal
  Dim svSecure, svCurrVisit, svLang, svProcesses, svPrograms, svProgram, svMultiUserManual, svPage, svTranslate, svEcomBypass, svBrowser, svProdNo, svProdMax

  Dim bDebug, vDebugOk, vDebug, vDebugStop, vFld, vValue, vSql, vFileOk, vFileDesc, vShellHi, vShellLo, vShellRootHi, vShellRootLo, vRightClickOff, vPage, vMembEmail

  '...data types for ado stored procs
  Const adCmdStoredProc = &H0004
  Const adParamInput = &H0001
  Const adParamOutput = &H0002

  '---- CursorTypeEnum Values ----
  Const adOpenForwardOnly = 0
  Const adOpenKeyset = 1
  Const adOpenDynamic = 2
  Const adOpenStatic = 3

  '---- LockTypeEnum Values ----
  Const adLockReadOnly = 1
  Const adLockPessimistic = 2
  Const adLockOptimistic = 3
  Const adLockBatchOptimistic = 4

  '---- CursorLocationEnum Values ----
  Const adUseServer = 2
  Const adUseClient = 3
  Const adParamReturnValue = 4
  Const adEmpty = 0

  Const adTinyInt = 16
  Const adSmallInt = 2
  Const adInteger = 3
  Const adBigInt = 20
  Const adUnsignedTinyInt = 17
  Const adUnsignedSmallInt = 18
  Const adUnsignedInt = 19
  Const adUnsignedBigInt = 21
  Const adSingle = 4
  Const adDouble = 5
  Const adCurrency = 6
  Const adDecimal = 14
  Const adNumeric = 131
  Const adBoolean = 11
  Const adError = 10
  Const adUserDefined = 132
  Const adVariant = 12
  Const adIDispatch = 9
  Const adIUnknown = 13
  Const adGUID = 72
  Const adDate = 7
  Const adDBDate = 133
  Const adDBTime = 134
  Const adDBTimeStamp = 135
  Const adBSTR = 8
  Const adChar = 129
  Const adVarChar = 200
  Const adLongVarChar = 201
  Const adWChar = 130
  Const adVarWChar = 202
  Const adLongVarWChar = 203
  Const adBinary = 128
  Const adVarBinary = 204
  Const adLongVarBinary = 205
  Const adChapter = 136
  Const adFileTime = 64
  Const adPropVariant = 138
  Const adVarNumeric = 139
  Const adArray = &H2000

	'...browser type
  If Len(Request("vBrowser")) > 0 Then 
    Session("Browser") = Request("vBrowser")
  End If

  '...disable Right Click for regular users
  If Instr(Lcase(Session("Host")), "localhost") > 0 Or Instr(Lcase(Session("Host")), "peter") > 0 Or Instr(Lcase(Session("Host")), "staging.") > 0  Or Session("MembLevel") = 5 Then
    vRightClickOff = False
  Else
    vRightClickOff = True
  End If

  '...get current page name for tracking and the translation engine
  Session("Page") = Request.ServerVariables("Script_Name") 
  Session("Page") = Mid(Session("Page"), InstrRev(Session("Page"), "/") + 1 )
  
  '...admins turning on/off the alert?
  If Request("vAlert") = "y" Or Request("vAlert") = "n"  Then 
    If Session("MembLevel") = 5 Then Application("Alert") = Request("vAlert")
  End If
  
  '...changing language?
  If Len(Request("vLang")) > 0 Then 
'   If Request("vLang") <> svLang Then Stop '...don't change language unless
    Session("Lang") = Request("vLang")
  End If
  If Session("Lang") = "" Then Session("Lang") = "EN"

  '...changing translation options?
  If Request("vTranslate") = "y" Then 
    Session("Translate") = True
  ElseIf Request("vTranslate") = "n" Then 
    Session("Translate") = False
  End If
  
  '...these files are executed to display a "joint" or "solo" shell in each page with a gui
  If Ucase(Request("vClose")) <> "Y" And Ucase(vClose) <> "Y" Then 
    vShellHi = "/V5/Inc/Shell_Hi.asp"
  Else
    vShellHi = "/V5/Inc/Shell_HiSolo.asp"
  End If  
  vShellLo   = "/V5/Inc/Shell_Lo.asp" 
 
  '...Determine host info and db info
  Session("SQL")            = "SQL01,1400"                                                      : svSQL       = Session("SQL")
  Session("SSL")            = fIf(Lcase(Request.ServerVariables("HTTPS")) = "on", True, False)  : svSSL       = Session("SSL")
  Session("Server")         = Lcase(Request.ServerVariables("HTTP_HOST"))                       : svServer    = Session("Server")
  Session("Host")           = Lcase(Request.ServerVariables("HTTP_HOST") & "/V5")               : svHost      = Session("Host")
  If Len(Session("HostDb")) = 0 Then Session("HostDb") = "V5_Vubz"                              : svHostDb    = Session("HostDb")
  Session("Domain")         = fIf(svSSL, "https://", "//") & svHost                        : svDomain    = Session("Domain")
  Session("HostDbPwd")      = "vudb2112mississauga"                                             : svHostDbPwd = Session("HostDbPwd") 

             
  vDebugStop = False
 
  '...go through the session variables and put into local variables and debug If required
' vDebugOk = False: If Lcase(Request.QueryString("vdebug"))="vudebug" AND Session("MembLevel") = 5 Then vDebugOk = True  
  vDebugOk = False: If Lcase(Request.QueryString("vdebug"))="vudebug"                              Then vDebugOk = True
  
  If vDebugOk Then vDebug = "<font color='orange'>"

  For Each vFld In Session.Contents
    vValue = Session(vFld)
    Select Case Lcase(vFld)    
      Case Lcase("CustId")              :svCustId                 = vValue
      Case Lcase("CustTitle")           :svCustTitle              = vValue
      Case Lcase("CustReturnUrl")       :svCustReturnUrl          = vValue
      Case Lcase("CustAcctId")          :svCustAcctId             = vValue
      Case Lcase("CustBanner")          :svCustBanner             = vValue
      Case Lcase("CustUrl")             :svCustUrl                = vValue
      Case Lcase("CustEmail")           :svCustEmail              = vValue
      Case Lcase("CustFreeHours")       :svCustFreeHours          = vValue
      Case Lcase("CustFreeDays")        :svCustFreeDays           = vValue
      Case Lcase("CustAuto")            :svCustAuto               = vValue
      Case Lcase("CustIssueIds")        :svCustIssueIds           = vValue
      Case Lcase("CustActivateIds")     :svCustActivateIds        = vValue
      Case Lcase("CustIdsSize")         :svCustIdsSize            = vValue
      Case Lcase("CustExpires")         :svCustExpires            = vValue
      Case Lcase("CustMaxHours")        :svCustMaxHours           = vValue
      Case Lcase("CustEcomDisc")        :svCustEcomDisc           = vValue
      Case Lcase("CustCluster")         :svCustCluster            = vValue
      Case Lcase("CustLevel")           :svCustLevel              = vValue
      Case Lcase("CustPwd")             :svCustPwd                = vValue
      Case Lcase("MembId")              :svMembId                 = vValue
      Case Lcase("MembPwd")             :svMembPwd                = vValue
      Case Lcase("MembNo")              :svMembNo                 = vValue
      Case Lcase("MembFirstName")       :svMembFirstName          = vValue
      Case Lcase("MembLastName")        :svMembLastName           = vValue
      Case Lcase("MembEmail")           :svMembEmail              = vValue
      Case Lcase("MembLevel")           :svMembLevel              = vValue
      Case Lcase("MembNoVisits")        :svMembNoVisits           = vValue
      Case Lcase("MembNoHours")         :svMembNoHours            = vValue
      Case Lcase("MembFirstVisit")      :svMembFirstVisit         = vValue
      Case Lcase("MembLastVisit")       :svMembLastVisit          = vValue
      Case Lcase("MembExpires")         :svMembExpires            = vValue    
      Case Lcase("MembCriteria")        :svMembCriteria           = vValue    
      Case Lcase("MembManager")         :svMembManager            = vValue
      Case Lcase("MembGap")             :svMembGap                = vValue
      Case Lcase("MembInternal")        :svMembInternal           = vValue

      Case Lcase("Browser")             :svBrowser                = vValue
      Case Lcase("CurrVisit")           :svCurrVisit              = vValue
      Case Lcase("EcomBypass")          :svEcomBypass             = vValue
      Case Lcase("Host")                :svHost                   = vValue
      Case Lcase("HostDb")              :svHostDb                 = vValue
      Case Lcase("HostDbPwd")           :svHostDbPwd              = vValue
      Case Lcase("Server")              :svServer                 = vValue
      Case Lcase("FontNo")              :svFontNo                 = vValue
      
      Case Lcase("MailServer")          :svMailServer             = vValue
      Case Lcase("MultiUserManual")     :svMultiUserManual        = vValue

      Case Lcase("Secure")              :svSecure                 = vValue
      Case Lcase("Lang")                :svLang                   = vValue
      Case Lcase("Page")                :svPage                   = vValue
      Case Lcase("Processes")           :svProcesses              = vValue
      Case Lcase("Programs")            :svPrograms               = vValue
      Case Lcase("Program")             :svProgram                = vValue
      Case Lcase("ProdNo")              :svProdNo                 = vValue
      Case Lcase("ProdMax")             :svProdMax                = vValue
      Case Lcase("Translate")           :svTranslate              = vValue

     End Select
    
    '...display the session variables if Debug
    If vDebugOk Then 
      If vFld = "Prod" Then 
        If Session("ProdNo") > 0 Then  
          Server.Execute "/V5/Inc/Debug_Prod.asp"
        End If
      Else
        sDebug vFld, vValue
      End If
    End If

  Next

  If vDebugStop Then Stop '...dummy statement for debugging

  '---DB Functions--------------------------------------------------------------------------------


  Sub sOpenDb
    On Error Resume Next
    vFileOk = False
    Set oDb = Server.CreateObject("ADODB.Connection")
    oDb.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
		oDb.CommandTimeout=100
    oDb.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDb
    oDb.Close
    Set oDb = Nothing
  End Sub


  Sub sOpenDb2
    On Error Resume Next
    vFileOk = False
    Set oDb2 = Server.CreateObject("ADODB.Connection")
    oDb2.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
		oDb2.CommandTimeout=100
    oDb2.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDb2
    oDb2.Close
    Set oDb2 = Nothing
  End Sub


  Sub sOpenDb3
    On Error Resume Next
    vFileOk = False
    Set oDb3 = Server.CreateObject("ADODB.Connection")
    oDb3.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
		oDb3.CommandTimeout=100
    oDb3.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDb3
    oDb3.Close
    Set oDb3 = Nothing
  End Sub


  Sub sOpenDb4
    On Error Resume Next
    vFileOk = False
    Set oDb4 = Server.CreateObject("ADODB.Connection")
    oDb4.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
		oDb4.CommandTimeout=100
    oDb4.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub


  Sub sCloseDb4
    oDb4.Close
    Set oDb4 = Nothing
  End Sub
  
  
  Sub sOpenDb5
    On Error Resume Next
    vFileOk = False
    Set oDb5 = Server.CreateObject("ADODB.Connection")
    oDb5.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
		oDb5.CommandTimeout=100
    oDb5.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDb5
    oDb5.Close
    Set oDb5 = Nothing
  End Sub


  Sub sOpenDbGap
    On Error Resume Next
    vFileOk = False
    Set oDbGap = Server.CreateObject("ADODB.Connection")
    oDbGap.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Gap;Data Source=" & svSQL
    oDbGap.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub


  Sub sCloseDbGap
    oDbGap.Close
    Set oDbGap = Nothing
  End Sub


  Sub sOpenDbGap2
    On Error Resume Next
    vFileOk = False
    Set oDbGap2 = Server.CreateObject("ADODB.Connection")
    oDbGap2.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Gap;Data Source=" & svSQL
    oDbGap2.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDbGap2
    oDbGap2.Close
    Set oDbGap2 = Nothing
  End Sub

  Sub sOpenDbBase
    On Error Resume Next
    vFileOk = False
    Set oDbBase = Server.CreateObject("ADODB.Connection")
    oDbBase.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Base;Data Source=" & svSQL
    oDbBase.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDbBase
    oDbBase.Close
    Set oDbBase = Nothing
  End Sub  


  Sub sOpenDbBase2
    On Error Resume Next
    vFileOk = False
    On Error Resume Next
    Set oDbBase2 = Server.CreateObject("ADODB.Connection")
    oDbBase2.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Base;Data Source=" & svSQL
    oDbBase2.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub


  Sub sCloseDbBase2
    oDbBase2.Close
    Set oDbBase2 = Nothing
  End Sub


  Sub sOpenCmd
    sOpenDb
    Set oCmd = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oDb
    oCmd.CommandType = adCmdStoredProc
    oCmd.CommandTimeout = 60000
  End Sub


  Sub sOpenCmdBase
    sOpenDbBase
    Set oCmdBase = Server.CreateObject("ADODB.Command")
    Set oCmdBase.ActiveConnection = oDbBase
    oCmdBase.CommandType = adCmdStoredProc
    oCmdBase.CommandTimeout = 60000
  End Sub


  Sub sOpenCmdGap
    sOpenDbGap
    Set oCmdGap = Server.CreateObject("ADODB.Command")
    Set oCmdGap.ActiveConnection = oDbGap
    oCmdGap.CommandType = adCmdStoredProc
    oCmdGap.CommandTimeout = 60000
  End Sub


  Sub sOpenCmdRTE
    sOpenDbRTE
    Set oCmdRTE = Server.CreateObject("ADODB.Command")
    Set oCmdRTE.ActiveConnection = oDbRTE
    oCmdRTE.CommandType = adCmdStoredProc
    oCmdRTE.CommandTimeout = 60000
  End Sub




  Sub sOpenCmdGap2
    sOpenDbGap2
    Set oCmdGap2 = Server.CreateObject("ADODB.Command")
    Set oCmdGap2.ActiveConnection = oDbGap2
    oCmdGap2.CommandType = adCmdStoredProc
    oCmdGap2.CommandTimeout = 60000
  End Sub


  Sub sOpenDbRTE
    On Error Resume Next
    vFileOk = False
    Set oDbRTE = Server.CreateObject("ADODB.Connection")
    oDbRTE.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
		oDbRTE.CommandTimeout=100
    oDbRTE.Open
    If Err.Number = 0 or Err.Number = "" Then vFileOk = True
  End Sub

  Sub sCloseDbRTE
    oDbRTE.Close
    Set oDbRTE = Nothing
  End Sub


  '---Other Functions--------------------------------------------------------------------------------

  '...used in Ecom2Catalogue.asp, Ecom2Programs.asp, Ecom3Programs.asp
  Function fPromo (i)
    fPromo = fIf(Len(i) > 2, "<br><i><font color='red'>" & i & "</font></i></br>", "")
  End Function


  '...is value null, empty or ""
  Function fNoValue (vTemp)
    fNoValue = False
    If VarType (vTemp) = vbEmpty Or VarType (vTemp) = vbNull Or vTemp = "" Then fNoValue = True  
  End Function


  '...set Value to "" if either null or empty 
  Function fOkValue (vTemp)
    If fNoValue(vTemp) Then 
      fOkValue = ""
    Else
      fOkValue = vTemp
    End If
  End Function


  '...set Value to Null if OK Value = ""
  Function fNullValue (vTemp)
    If fOkValue(vTemp) = "" Then 
      fNullValue = Null
    Else
      fNullValue = vTemp
    End If
  End Function  


  '...if i is a clean integer, ie 234 rather than -2,34 then return as a long integer, otherwise return 0
  Function fPureInt(i)
    fPureInt = Clng(0)
    If IsNumeric(fOkValue(i)) Then
      If Int(Abs(i)) = Csng(i) Then
        If i > 0 Then
          fPureInt = Clng(i)
        End If
      End If
    End If
  End Function  


  '...if i is numeric, return the positive integer (rounded up) value else return zero
  Function fOkInt(i)
    fOkInt = Clng(0)
    If IsNumeric(fOkValue(i)) Then
      fOkInt = Round(i + .5)    
      fOkInt = Abs(Clng(fOkInt))
    End If
  End Function  


  Sub sDebug
    If Lcase(svServer) = "localhost" Or svMembLevel = 5 Then
      On Error Resume Next
'     Response.Write "<br><b><font color='ORANGE'>(" & Now & ") " & vSql & "</font></b><br>"
      Response.Write "<br><b><font color='ORANGE'>" & vSql & "</font></b><br>"
      If Err.Number > 0 Then 
        Response.Write "<br><b><font color='ORANGE'>" & "[unprintable]" & "</font></b><br>"
      End If
      Response.Flush
      On Error GoTo 0
    End If
  End Sub

  
  '...replace single and smart quotes with two single quotes so we don't screw up SQL
  Function fUnQuote (vTemp)
    If fNoValue(vTemp) Then
      fUnquote = ""
    Else
      fUnquote = vTemp
      
      fUnquote = Replace(fUnquote, "“", "'")
      fUnquote = Replace(fUnquote, "”", "'")
      fUnquote = Replace(fUnquote, "‘", "'")
      fUnquote = Replace(fUnquote, "’", "'")
      fUnquote = Replace(fUnquote, """", "'")
      fUnquote = Replace(fUnquote, "''", "'")

      fUnquote = Trim(Replace(fUnquote, "'", "''"))
    End If
  End Function
  

  '...replace single quotes with \' so we don't screw up Java Script
  Function fjUnQuote (vTemp)
    If fNoValue(vTemp) Then
      fjUnquote = ""
    Else
      fjUnquote = Replace(vTemp, "'", "\'")
    End If
  End Function 


  '...remove single or double quotes
  Function fNoQuote (vTemp)
    If fNoValue(vTemp) Then
      fNoQuote = ""
    Else
      fNoQuote = Replace(vTemp,"'"," ")
      fNoQuote = Trim(Replace(fNoQuote, Chr(34), " "))
    End If
  End Function


  '...replace single quotes with html quotes so we don't screw up html in content.asp
  Function fHtmlUnquote (vTemp)
    If fNoValue(vTemp) Then
      fHtmlUnquote = ""
    Else
      fHtmlUnquote = Replace(vTemp,"'","&#39;")
    End If
  End Function


  '...Server.UrlEncode but replace "+" with "%20" 
  Function fUrlEncode(i)
    fUrlEncode = Server.UrlEncode (fOkValue(i))
    fUrlEncode = Replace(fUrlEncode, "+", "%20")
  End Function 
  

 
  '...set check boxes or radio buttons (or drop down combos)
  Function fCheck (vFld, vValue)
    fCheck = "" : If Ucase(vValue) = Ucase(vFld) Then fCheck = "checked"
  End Function


  '...set check boxes or radio buttons (or drop down combos) - will select if within a group of items
  Function fChecks (vFld, vValue)
    fChecks = "" : If Instr(Ucase(fOkValue(vFld)), Ucase(vValue)) > 0 Then fChecks = "checked"
  End Function

  
  Function fSelect (vFld, vValue)
    fSelect = "" : If Ucase(vValue) = Ucase(vFld) Then fSelect = "selected"
  End Function

  Function fSelects (vFld, vValue)
    fSelects = "" : If Instr(Ucase(fOkValue(vFld)), Ucase(vValue)) > 0 Then fSelects = "selected"
  End Function


  '...is boolean if 1 or 0
  Function fBoolean(i)
    fBoolean = False
    If IsNumeric(i) Then
	    If i = 0 Or i = 1 Then fBoolean = True
	  End If
  End Function
  

  '...Everything is False unless True, 1, Y, y
  Function fBooleanPlus(i)
    fBooleanPlus = True
    If VarType(i) = vbBoolean Then 
      If i Then Exit Function
    End If
    If fBoolean(i) Then 
      If i = 1 Then Exit Function
    End If
    If i = "y" Or i = "Y" Then Exit Function
    fBooleanPlus = False
  End Function

  
  '...convert true/false to 0/1 for sql
  Function fSqlBoolean (i)
    fSqlBoolean = 0
    If VarType(i) = vbBoolean Then
      If i = True Then fSqlBoolean = 1
    ElseIf VarType(i) = vbString Then
      If Lcase(i) = "true" Or i = "1" Then fSqlBoolean = 1
    ElseIf VarType(i) = vbInteger Then
      If i = 1 Then fSqlBoolean = 1
    End If
  End Function


  '...as above but leave as "" if not boolean
  Function fSqlBooleanPlus (i)
    If Vartype(i) = vbBoolean Then
      fSqlBooleanPlus = fSqlBoolean(i)
    Else
      fSqlBooleanPlus = ""
    End If
  End Function


  Function fFormatCurrency (i, j)
    If svLang = "FR" Then
      fFormatCurrency = Replace(FormatNumber(i, j), ".", ",")
    Else
      fFormatCurrency = FormatCurrency(i, j)
    End If
  End Function


  Function fFormatDate (i)
    Dim aMonth
'   If i = "" Then fFormatDate = "" : Exit Function '...if they clear out the date leave empty
    fFormatDate = " "
    If Not IsDate (i) Then Exit Function
    If Year(i) < 2000 Then Exit Function
    Select Case svLang
      Case "FR" : aMonth = Split ("janv. févr. mars avril mai juin juillet août sept. oct. nov. déc.", " ") : fFormatDate = Day(i) & " " & aMonth(Month(i) -1) & " " & Year(i)                 
      Case "ES" : aMonth = Split ("ene. feb. mar. abr. may. jun. jul. ago. sept. oct. nov. dic.", " ")      : fFormatDate = Day(i) & " " & aMonth(Month(i) -1) & " " & Year(i)
      Case Else : aMonth = Split ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", " ")                   : fFormatDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
    End Select
  End Function


  Function fFormatDateTime (i)
    Dim aMonth
    fFormatDateTime = " "
    If Not IsDate (i) Then Exit Function
    If Year(i) < 2000 Then Exit Function
    Select Case svLang
      Case "FR" : aMonth = Split ("janv. févr. mars avril mai juin juillet août sept. oct. nov. déc.", " ") : fFormatDateTime = Day(i) & " " & aMonth(Month(i) -1) & " " & Year(i)                 
      Case "ES" : aMonth = Split ("ene. feb. mar. abr. may. jun. jul. ago. sept. oct. nov. dic.", " ")      : fFormatDateTime = Day(i) & " " & aMonth(Month(i) -1) & " " & Year(i)
      Case Else : aMonth = Split ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", " ")                   : fFormatDateTime = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
    End Select
    If svLang = "EN" Then
      fFormatDateTime = fFormatDateTime & " " & Right("00" & Hour(i), 2) & ":" & Right("00" & Minute(i), 2)
    End If    
  End Function


  '...use this for internal date management, like SQL and must be in EN
  Function fFormatSqlDate (i)
    Dim aMonth
    fFormatSqlDate = " "
    If Not IsDate (i) Then Exit Function
    If Year(i) < 2000 Then Exit Function
    aMonth = Split ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", " ")
    fFormatSqlDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
  End Function  
  

  Function fFormatSqlDateTime (i)
    Dim aMonth
    fFormatSqlDateTime = " "
    If Not IsDate (i) Then Exit Function
    If Year(i) < 2000 Then Exit Function
    aMonth = Split ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", " ")                   
    fFormatSqlDateTime = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
    fFormatSqlDateTime = fFormatSqlDateTime & " " & Right("00" & Hour(i), 2) & ":" & Right("00" & Minute(i), 2)
  End Function



  Function fFormatMonth (i)
    Dim aMonth
    fFormatMonth = " "
    If Not IsNumeric(i) Then Exit Function
    If i < 1 or i > 12 Then Exit Function
    Select Case svLang
      Case "FR" : aMonth = Split ("janv. fév. mars avr. mai juin juill. août sept. oct. nov. déc.", " ") : fFormatMonth = aMonth(i -1)
      Case "ES" : aMonth = Split ("ene. feb. mar. abr. may. jun. jul. ago. sept. oct. nov. dic.", " ")   : fFormatMonth = aMonth(i -1)
      Case Else : aMonth = Split ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", " ")                : fFormatMonth = aMonth(i -1)
    End Select
  End Function


  '...returns larger (most recent) date
  Function fMaxDate(i, j)
    If IsDate(i) And IsDate(j) Then
      If DateDiff("d", i, j) < 0 Then
        fMaxDate = j
      End If  
    ElseIf IsDate(i) Then
      fMaxDate = i
    ElseIf IsDate(j) Then
      fMaxDate = j 
    Else
      fMaxDate = ""      
    End If    
  End Function


  Function fFormatDecimals (i)
    If svLang = "FR" Then
      fFormatDecimals = Replace(i, ".", ",")
    Else
      fFormatDecimals = i
    End If
  End Function


  '...these values return "" if i = 0
  Function fFormatPercent (i, j)
    fFormatPercent = fIf(i > 0, FormatPercent (i, j), "")
  End Function
  Function fFormatNumber (i, j)
    fFormatNumber = fIf(i > 0, FormatNumber (i, j), "")
  End Function


  '...if field is wider than j chars, truncate to j
  Function fLeft(i, j)
    i = fOkValue(i)
    If Len(i) > j Then
      fLeft = Left(i, j) & " ..."
'     fLeft = Left(i, j) & "<a title='" & i & "' href='#'><span class='green'>...</span></a>"
    Else
      fLeft = i    
    End If    
  End Function


  '...if field is wider than j chars, truncate to j but allow mouseover (be careful)
  Function fSmartLeft(i, j)
    i = fOkValue(i)
    If Len(i) > j Then
      fSmartLeft= Left(i, j) & "<a title='" & i & "' href='#'>...</a>"
    Else
      fSmartLeft= i    
    End If    
  End Function


  '...return j if i has no value
  Function fDefault(i, j)
    If fNoValue(i) Then
      fDefault = j
    Else
      fDefault = i
    End If  
  End Function


  '...return j greater than i
  Function fMax(i, j)
    If fNoValue(i) Then i = 0
    If fNoValue(j) Then j = 0
    If j > i Then
      fMax = j
    Else
      fMax = i
    End If  
  End Function

  
  '...return j lesser than i
  Function fMin(i, j)
    If fNoValue(i) Then i = 0
    If fNoValue(j) Then j = 0
    If j < i Then
      fMin = j
    Else
      fMin = i
    End If  
  End Function  
  
  
  '...IIf Function
  Function fIf(i, j, k)
    fIf = k
    If Vartype(i) = 11 Then     
      If i Then fIf = j
    End If
  End Function

  '...returns True/False if i is a valid email 
  Function fIsEmail(i)
    Dim regEx
    Set regEx = New RegExp
    '...original from functions.js:    
    '   var reEmail = new RegExp( /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,6})+$/ ); // original basic email edit
    '             regEx.Pattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,6})+$"
    '... this should match what's in functions.js but that was causing me issues
    regEx.Pattern = "\b(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@"_
                  & "((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."_
                  & "([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|"_
                  & "([a-zA-Z]+[\w-]+\.)+[a-zA-Z]{2,4})\b"
    regEx.Global = true
    fIsEmail = regEx.Test(i)
  End Function

  
  '...Clear status line on links
  Function fStatX 
    If vRightClickOff Then 
      fStatX = "onmouseover=""javascript:window.status=' ';return true"" onmousedown=""javascript:window.status=' ';return true"" onmouseout=""javascript:window.status=' ';return true"""
    Else
      fStatX = ""
    End If
  End Function


  Function f5
    f5 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  End Function 

  Function f10
    f10 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  End Function 


	'...this generates a maxLength event for textareas (plus an onblur event to handle pasting of characters)
	Function fMaxLength(i)
		fMaxLength = " onkeypress=" & Chr(34) & "return fMax('" & svLang & "', this, " & i & ")" & Chr(34) _
		           & " onblur=" & Chr(34) & "return fMax('" & svLang & "', this, " & i & ")" & Chr(34)
	End Function 


	'...this generates a minLength event for any field 
	Function fMinLength(i)
		fMinLength = " onblur=" & Chr(34) & "return fMin('" & svLang & "', this, " & i & ")" & Chr(34) 
	End Function 
  
  
  '...this retrieves the original value (typically done by the browser but used in Web Services)
  Function fUrlDecode(sConvert)
    Dim aSplit, sOutput, I
    If IsNull(sConvert) Or Len(sConvert) = 0 Then
       fUrlDecode = ""
       Exit Function
    End If	
    '... convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")	
    '... next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")	
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If	
    fUrlDecode = sOutput
  End Function  

  
  '...this creates an attractive button with class vType and text of vText - vText should be passed in using
  '   the various button values found in Db_Phra.asp, ie fButton("Next", bNext)

  Function fButton(vType, vText)
    fButton = "<a id='butShell' class='butShell' href='javascript:void(0)'><span class='butIcon butXX'></span><input class='butInput' type='submit' name='butXX' id='butXX' value='YY' /></a>"
    fButton = Replace(fButton, "XX", vType)
    fButton = Replace(fButton, "YY", vText)
  End Function

	'... the outer span allows text when the inner span is disabled (see Activity_adv_f.asp).
  Function fButtonS(vType, vText, vJava)
    fButtonS = "<span id='divMessage'><span id='divXX'><a id='butShell' class='butShell' href='javascript:void(0)'><span class='butIcon butXX'></span><input class='butInput' type='button' onclick='JS' name='butXX' id='butXX' value='YY' /></a></span ></span >"
    fButtonS = Replace(fButtonS, "XX", vType)
    fButtonS = Replace(fButtonS, "YY", vText)
    fButtonS = Replace(fButtonS, "JS", vJava)
  End Function




  '...this is a general purpose SQL UPDATE routine
  '   that retrieves all fields in the recordset in any table in any databse that
  '   contains a primary key whose name is vTable & "_No" (ie "Memb_No")

  '...it can be called from any page but that page must contain the appropriate
  '   "V5\Inc\Db_Xxxx.asp" include statement for the appropriate table
  '   so we can check all table values to see if they are not empty, and if they 
  '   they have any value (other than primary key values), they will be updated
  '...note that values that are null or set to "" will also be updated

  Sub sTableUpdate(vDb, vTable, vKey)
    Dim oDb, oRs, vValue, bTrack, vFldNo
    bTrack = False
    vFileDesc = ""
    On Error Resume Next
    vFileOk = False
    Set oDb = Server.CreateObject("ADODB.Connection")
    oDb.ConnectionString = "Provider=SQLOLEDB.1;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & vDb & ";Data Source=" & svSQL
    oDb.Open
    If IsEmpty(vKey) Then
      vFileOk   = False
      vFileDesc = "The primary key does not exist."
    ElseIf Not IsNumeric(vKey) Then
      vFileOk   = False
      vFileDesc = "The primary key is not numeric."
    ElseIf Not IsNumeric(vKey) Then
      vFileOk   = False
      vFileDesc = "The primary key must be greater than zero."
    ElseIf Err.Description <> "" Then 
      vFileOk   = False
      vFileDesc = "That Database does not exist."
    Else
      Set oRs = CreateObject("ADODB.Recordset")      
      With oRs
        Set .ActiveConnection = oDb
        .CursorType       = adOpenStatic
        .CursorLocation   = adUseClient
        .LockType         = adLockOptimistic  
        '...ensure there is a stored proc available
        On Error Resume Next
        .Open "SELECT * FROM " & vTable & " WHERE " & vTable & "_No = " & vKey
        If Err.Description <> "" Then
          vFileOk   = False
          vFileDesc = "That table does not exist."
        Else
          On Error Goto 0  
          '...ensure that there is a record with that member no
          If oRs.Eof Then
            vFileOk   = False
            vFileDesc = "That Learner is not on file"
          Else
            If bTrack Then    
              vFileDesc = "<table border='1' cellspacing='0' style='border-collapse: collapse; font-family: Calibri; font-size: 10pt' bordercolor='#00FFFF' cellpadding='5'>"
              vFileDesc = vFileDesc & "<tr><td><b>Field</b></td><td><b>Before</b></td><td><b>After</b></td></tr>"           
            End If
            For vFldNo = 0 To oRs.Fields.Count - 1
              vFld = "v" & oRs.Fields(vFldNo).Name                
              If Not IsEmpty(Eval(vFld)) And (oRs.Fields(vFldNo).Properties("KEYCOLUMN").Value = False Or oRs.Fields(vFldNo).Properties("ISAUTOINCREMENT").Value = False) Then
                If bTrack Then vFileDesc = vFileDesc & "<tr><td>" & vFld & "</td><td>" & oRs.Fields(vFldNo).Value & "</td><td>" & Eval(vFld) & "</td>"
                vValue    = Eval(vFld)
                '...remove any double quotes put in for regular sql (if there's a value)
                If Len(fOkValue(vValue)) > 0 Then
                  vValue    = Replace(vValue, "''", "'")
                End If

                '...ignore GUIDs (only used on create)
                If oRs.Fields(vFldNo).Type <> adGuid Then

                  '...is this a date field?
                  If oRs.Fields(vFldNo).Type = 7 Or oRs.Fields(vFldNo).Type = 133 Or oRs.Fields(vFldNo).Type = 135 Then
                    '...if Date = 0 then set field to Null 
                    If vValue = "" Then
                      .Fields(vFldNo).Value = Null
                    '...if valid Date then update field
                    ElseIf IsDate(vValue) Then
                      .Fields(vFldNo).Value = vValue
                    '...otherwise use whatever default is in place
                    End If
                  Else
                    .Fields(vFldNo).Value = vValue
                  End If

                End If

              End If    
              If bTrack Then vFileDesc = vFileDesc & "</tr>"    
            Next
            .Update    
          End If   
        End If       
        .Close
        vFileOk   = True
      End With    
      If bTrack Then vFileDesc = vFileDesc & "</table>"  
      Set oRs  = Nothing
    End If
    oDb.Close
    Set oDb = Nothing
    If bTrack Then Response.Write vFileDesc
  End Sub
  

  '... use for excelwriter to render dates as YYYY.MM.DD - configure as strings
  Function fExcelDate (vDate)
    If IsDate(vDate) Then
      fExcelDate = Year(vDate) & "." & Right("00" & Month(vDate), 2) & "." & Right("00" & Day(vDate), 2)
    Else
      fExcelDate =""
    End If 
  End Function  


  Function fYN (i)
    If fBooleanPlus (i) Then 
      Select Case svLang
        Case "ES"  fYn = "sí"
        Case "FR"  fYn = "oui"
        Case Else  fYn = "Yes"
      End Select
    Else
      Select Case svLang
        Case "ES"  fYn = "¡no"
        Case "FR"  fYn = "non"
        Case Else  fYn = "No"
      End Select
    End If
  End Function
  

%>