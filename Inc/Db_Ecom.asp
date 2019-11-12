<%

  Dim vEcom_No, vEcom_AcctId, vEcom_CustId, vEcom_Id, vEcom_MembNo, vEcom_CatlNo, vEcom_Programs, vEcom_Prices, vEcom_Taxes, vEcom_Issued, vEcom_Expires, vEcom_Amount, vEcom_Currency, vEcom_Lang, vEcom_FirstName, vEcom_LastName, vEcom_CardName, vEcom_Address, vEcom_City, vEcom_Postal, vEcom_Province, vEcom_Country, vEcom_Phone, vEcom_Email, vEcom_Quantity, vEcom_Seats, vEcom_NewAcctId, vEcom_Media, vEcom_Label, vEcom_OrderNo, vEcom_Shipping, vEcom_Memo, vEcom_Organization, vEcom_Source, vEcom_InternetSecure, vEcom_Adjustment, vEcom_Agent, vEcom_Archived
  Dim vEcom_OrderId, vEcom_LineId
  Dim vEcom_Eof, vEcom_Ok

  '____ Ecom  ________________________________________________________________________

  Sub sAddEcom ()
    '...modified Mar 01, 2017 to unquote organization field
    '...modified Aug 24, 2017 to unquote the other fields
    '...modified Dec 14, 2017 to unquote the other fields
    If Instr("EN FR ES", vEcom_Lang) = 0 Then vEcom_Lang = "EN"
    vSql = "INSERT INTO Ecom (Ecom_CustId, Ecom_AcctId, Ecom_Id, Ecom_MembNo, Ecom_CatlNo, Ecom_Programs, Ecom_Prices, Ecom_Taxes, Ecom_Issued, Ecom_Expires, Ecom_Amount, Ecom_Currency, Ecom_Lang, Ecom_FirstName, Ecom_LastName, Ecom_CardName, Ecom_Address, Ecom_City, Ecom_Postal, Ecom_Province, Ecom_Country, Ecom_Phone, Ecom_Email, Ecom_Quantity, Ecom_NewAcctId, Ecom_Media, Ecom_Label, Ecom_OrderNo, Ecom_Shipping, Ecom_Memo, Ecom_Organization, Ecom_Source, Ecom_InternetSecure, Ecom_Agent)"
    vSql = vSql & " VALUES ('" & vEcom_CustId & "', '" & vEcom_AcctId & "', '" & vEcom_Id & "', " & vEcom_MembNo & ", " & fDefault(vEcom_CatlNo, 0) & ", '" & fIf(Len(vEcom_Programs) > 7, Right(vEcom_Programs, 7), vEcom_Programs) & "', " & vEcom_Prices & ", " & vEcom_Taxes & ", '" & vEcom_Issued & "', '" & vEcom_Expires & "', " & vEcom_Amount & ", '" & vEcom_Currency & "', '" & vEcom_Lang & "', '" & fUnquote(vEcom_FirstName) & "', '" & fUnquote(vEcom_LastName) & "', '" & fUnquote(vEcom_CardName) & "', '" & fUnquote(vEcom_Address) & "', '" & fUnquote(vEcom_City) & "', '" & vEcom_Postal & "', '" & vEcom_Province & "', '" & fUnquote(vEcom_Country) & "', '" & vEcom_Phone & "', '" & vEcom_Email & "', " & vEcom_Quantity & ", '" & vEcom_NewAcctId & "', '" & vEcom_Media & "', '" & vEcom_Label & "', '" & vEcom_OrderNo & "', " & vEcom_Shipping & ", '" & fUnquote(vEcom_Memo) & "', '" & fUnquote(vEcom_Organization) & "', '" & fDefault(vEcom_Source,"E") & "', '" & vEcom_InternetSecure & "', '" & vEcom_Agent & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  '...Get Ecom programs that are NOT expired - note: there can be multiple orders from the same member
  '   note the commented lines are from when we passed thru the expire date
  Function fEcomPrograms (vCustId, vMembId)
    Dim aPrograms
    fEcomPrograms = ""
    vSql = "SELECT Ecom_Programs, Ecom_Expires FROM Ecom WHERE Ecom_Archived IS NULL AND Ecom_CustId = '" & vCustId & "' AND Ecom_Id= '" & vMembId & "' AND Ecom_Expires >= '" & fFormatSqlDate(Now) & "'"
'   sDebug
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      '...break up program string (if more than one program) and add issue date beside each from which the expired date can be determined in Content.asp
      aPrograms = Split(Trim(oRs("Ecom_Programs"))," ")
      For i = 0 to Ubound(aPrograms)      
        fEcomPrograms = fEcomPrograms & aPrograms(i) & "~" & oRs("Ecom_Expires") & "|" 
      Next
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb
    '...strip the trailing pipe
    If Instr(fEcomPrograms, "|") > 0 Then 
      fEcomPrograms = Left(fEcomPrograms, Len(fEcomPrograms)-1)
    End If
  End Function



  '...Get Ecom programs that are NOT expired but do not return the expiry dates (used in MyWorldCode.asp)
  Function fEcomProgram2 (vCustId, vMembId)
    Dim aPrograms, i
    fEcomProgram2 = ""
    vSql = "SELECT Ecom_Programs FROM Ecom WHERE Ecom_Archived IS NULL AND Ecom_CustId = '" & vCustId & "' AND Ecom_Id= '" & vMembId & "' AND Ecom_Expires >= '" & fFormatSqlDate(Now) & "'"
'   sDebug
    sOpenDb3
    Set oRs3 = oDb3.Execute(vSql)
    Do While Not oRs3.Eof
      '...break up program string (if more than one program)
      aPrograms = Split(Trim(oRs3("Ecom_Programs"))," ")
      For i = 0 to Ubound(aPrograms)      
        fEcomProgram2 = fEcomProgram2 & aPrograms(i) & " " 
      Next
      oRs3.MoveNext
    Loop
    Set oRs3 = Nothing
    sCloseDb3
    fEcomProgram2 = Trim(fEcomProgram2)
  End Function



  '...Get MembEcom Recordset
  Sub sGetEcom_Rs (vEcomCustId, vEcomMembId)
    vSql = "SELECT * FROM Ecom WHERE Ecom_Archived IS NULL AND Ecom_CustId = '" & vEcomCustId & "' AND Ecom_Id= '" & vEcomMembId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
  End Sub


  '...Get MembEcom Recordset by NewAcctId
  Sub sGetEcomByNewAcctId_Rs (vEcomNewAcctId)
    vSql = "SELECT * FROM Ecom WHERE Ecom_Archived IS NULL AND Ecom_NewAcctId = '" & vEcomNewAcctId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
  End Sub


  '...Get CustId by NewAcctId
  Function fEcomCustId (vEcomNewAcctId)
    vSql = "SELECT DISTINCT Ecom_CustId FROM Ecom WHERE Ecom_Archived IS NULL AND Ecom_NewAcctId = '" & vEcomNewAcctId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      fEcomCustId = ""
    Else
      fEcomCustId = oRs("Ecom_CustId")
    End If
    Set oRs = Nothing
    sCloseDb  
  End Function


  '...Get Ecom Recordset by No
  Sub sGetEcom
    vSql = "SELECT * FROM Ecom WHERE Ecom_No = " & vEcom_No
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadEcom
      vEcom_Eof = False
    Else
      vEcom_Eof = True
    End If
    Set oRs = Nothing
    sCloseDb    
  End Sub


  Sub sReadEcom
    vEcom_No             = oRs("Ecom_No")
    vEcom_AcctId         = oRs("Ecom_AcctId")
    vEcom_CustId         = oRs("Ecom_CustId")
    vEcom_Id             = oRs("Ecom_Id")
    vEcom_MembNo         = oRs("Ecom_MembNo")
    vEcom_Programs       = oRs("Ecom_Programs")
    vEcom_CatlNo         = oRs("Ecom_CatlNo")
    vEcom_Prices         = oRs("Ecom_Prices")
    vEcom_Taxes          = oRs("Ecom_Taxes")
    vEcom_Issued         = oRs("Ecom_Issued")
    vEcom_Expires        = oRs("Ecom_Expires")
    vEcom_Amount         = oRs("Ecom_Amount")
    vEcom_Currency       = oRs("Ecom_Currency")
    vEcom_Lang           = oRs("Ecom_Lang")
    vEcom_FirstName      = oRs("Ecom_FirstName")
    vEcom_LastName       = oRs("Ecom_LastName")
    vEcom_CardName       = oRs("Ecom_CardName")
    vEcom_Address        = oRs("Ecom_Address")
    vEcom_City           = oRs("Ecom_City")
    vEcom_Postal         = oRs("Ecom_Postal")
    vEcom_Province       = oRs("Ecom_Province")
    vEcom_Country        = oRs("Ecom_Country")
    vEcom_Phone          = oRs("Ecom_Phone")
    vEcom_Email          = oRs("Ecom_Email")
    vEcom_Quantity       = oRs("Ecom_Quantity")
    vEcom_NewAcctId      = oRs("Ecom_NewAcctId")
    vEcom_Media          = oRs("Ecom_Media")
    vEcom_Label          = oRs("Ecom_Label")
    vEcom_OrderNo        = oRs("Ecom_OrderNo")
    vEcom_Shipping       = oRs("Ecom_Shipping")
    vEcom_Memo           = oRs("Ecom_Memo")
    vEcom_Organization   = oRs("Ecom_Organization")
    vEcom_Source         = oRs("Ecom_Source")
    vEcom_InternetSecure = oRs("Ecom_InternetSecure")
    vEcom_Adjustment     = oRs("Ecom_Adjustment")
    vEcom_Agent          = oRs("Ecom_Agent")
    vEcom_Archived       = oRs("Ecom_Archived")
    vEcom_OrderId        = oRs("Ecom_OrderId")
    vEcom_LineId         = oRs("Ecom_LineId")
  End Sub

  Sub sExtractEcom
    vEcom_No             = Request.Form("vEcom_No")
    vEcom_MembNo         = Request.Form("vEcom_MembNo")
    vEcom_AcctId         = Request.Form("vEcom_AcctId")
    vEcom_CustId         = Request.Form("vEcom_CustId")
    vEcom_Id             = Request.Form("vEcom_Id")
    vEcom_Programs       = Request.Form("vEcom_Programs")
    vEcom_CatlNo         = Request.Form("vEcom_CatlNo")
    vEcom_Prices         = Request.Form("vEcom_Prices")
    vEcom_Taxes          = Request.Form("vEcom_Taxes")
    vEcom_Issued         = Request.Form("vEcom_Issued")
    vEcom_Expires        = Request.Form("vEcom_Expires")
    vEcom_Amount         = Request.Form("vEcom_Amount")
    vEcom_Currency       = Request.Form("vEcom_Currency")
    vEcom_Lang           = Request.Form("vEcom_Lang")
    vEcom_FirstName      = fUnquote(Request.Form("vEcom_FirstName"))
    vEcom_LastName       = fUnquote(Request.Form("vEcom_LastName"))
    vEcom_CardName       = fUnquote(Request.Form("vEcom_CardName"))
    vEcom_Address        = fUnquote(Request.Form("vEcom_Address"))
    vEcom_City           = fUnquote(Request.Form("vEcom_City"))
    vEcom_Postal         = fUnquote(Request.Form("vEcom_Postal"))
    vEcom_Province       = Request.Form("vEcom_Province")
    vEcom_Country        = Request.Form("vEcom_Country")
    vEcom_Phone          = Request.Form("vEcom_Phone")
    vEcom_Email          = Request.Form("vEcom_Email")
    vEcom_Quantity       = Request.Form("vEcom_Quantity")
    vEcom_NewAcctId      = Request.Form("vEcom_NewAcctId")
    vEcom_Media          = Request.Form("vEcom_Media")
    vEcom_Label          = Request.Form("vEcom_Label")
    vEcom_OrderNo        = Request.Form("vEcom_OrderNo")
    vEcom_OrderId        = Request.Form("vEcom_OrderId")
    vEcom_LineId         = Request.Form("vEcom_LineId")
    vEcom_Shipping       = Request.Form("vEcom_Shipping")
    vEcom_Memo           = fUnquote(Request.Form("vEcom_Memo"))
    vEcom_Organization   = fUnquote(Request.Form("vEcom_Organization"))
    vEcom_Source         = fDefault(Request.Form("vEcom_Source"), "E")
    If fNoValue(vEcom_Quantity) Then
      vEcom_Quantity     = 1
    ElseIf Not IsNumeric(vEcom_Quantity) Then   
      vEcom_Quantity     = 1
    End If
    vEcom_InternetSecure = Request.Form("vEcom_InternetSecure")
    vEcom_Adjustment     = fDefault(Request.Form("vEcom_Adjustment"), 0)
    vEcom_Agent          = Request.Form("vEcom_Agent")
  End Sub
  
  Sub sInsertEcom
    vSql = "SET ANSI_WARNINGS ON "
    vSql = vSql & "INSERT INTO Ecom "
    vSql = vSql & "(Ecom_CustId, Ecom_AcctId, Ecom_Id, Ecom_MembNo, Ecom_Programs, Ecom_Prices, Ecom_Taxes, Ecom_Issued, Ecom_Expires, Ecom_Amount, Ecom_Currency, Ecom_Lang, Ecom_FirstName, Ecom_LastName, Ecom_CardName, Ecom_Address, Ecom_City, Ecom_Postal, Ecom_Province, Ecom_Country, Ecom_Phone, Ecom_Email, Ecom_Quantity, Ecom_NewAcctId, Ecom_Media, Ecom_Label, Ecom_OrderNo, Ecom_Shipping, Ecom_Memo, Ecom_Organization, Ecom_Source, Ecom_InternetSecure, Ecom_Adjustment, Ecom_Agent, Ecom_OrderId, Ecom_LineId)"
    vSql = vSql & " VALUES ('" & vEcom_CustId & "', '" & vEcom_AcctId & "', '" & vEcom_Id & "', " & vEcom_MembNo & ", '" & vEcom_Programs & "', " & vEcom_Prices & ", " & vEcom_Taxes & ", '" & vEcom_Issued & "', '" & vEcom_Expires & "', " & vEcom_Amount & ", '" & vEcom_Currency & "', '" & vEcom_Lang & "', '" & vEcom_FirstName & "', '" & vEcom_LastName & "', '" & vEcom_CardName & "', '" & vEcom_Address & "', '" & vEcom_City & "', '" & vEcom_Postal & "', '" & vEcom_Province & "', '" & vEcom_Country & "', '" & vEcom_Phone & "', '" & vEcom_Email & "', " & vEcom_Quantity & ", '" & vEcom_NewAcctId & "', '" & vEcom_Media & "', '" & vEcom_Label & "', '" & vEcom_OrderNo & "', " & fDefault(vEcom_Shipping,0) & ", '" & vEcom_Memo & "', '" & vEcom_Organization & "', '" & vEcom_Source & "', '" & vEcom_InternetSecure & "', " & fSqlBoolean (vEcom_Adjustment) & ", '" & vEcom_Agent & "', '" & vEcom_OrderId & "', '" & vEcom_LineId & "')"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    '...get last ecom_no added
    vSql = " SELECT TOP 1 Ecom_No FROM Ecom ORDER BY Ecom_No DESC"
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then vEcom_No = oRs("Ecom_No")
    sCloseDb
  End Sub

  Sub sUpdateEcom
    vSql = "SET ANSI_WARNINGS ON "
    vSql = vSql & "UPDATE Ecom SET"
    vSql = vSql & " Ecom_CustId          = '" & vEcom_CustId                     & "', " 
    vSql = vSql & " Ecom_AcctId          = '" & vEcom_AcctId                     & "', " 
    vSql = vSql & " Ecom_Id              = '" & vEcom_Id                         & "', " 
    vSql = vSql & " Ecom_MembNo          =  " & vEcom_MembNo                     & " , " 
    vSql = vSql & " Ecom_Programs        = '" & vEcom_Programs                   & "', " 
    vSql = vSql & " Ecom_CatlNo          =  " & fDefault(vEcom_CatlNo, 0)        & " , " 
    vSql = vSql & " Ecom_Prices          =  " & vEcom_Prices                     & " , " 
    vSql = vSql & " Ecom_Taxes           =  " & vEcom_Taxes                      & " , " 
    vSql = vSql & " Ecom_Issued          = '" & vEcom_Issued                     & "', " 
    vSql = vSql & " Ecom_Expires         = '" & vEcom_Expires                    & "', " 
    vSql = vSql & " Ecom_Amount          =  " & vEcom_Amount                     & " , " 
    vSql = vSql & " Ecom_Currency        = '" & vEcom_Currency                   & "', " 
    vSql = vSql & " Ecom_Lang            = '" & vEcom_Lang                       & "', " 
    vSql = vSql & " Ecom_FirstName       = '" & vEcom_FirstName                  & "', " 
    vSql = vSql & " Ecom_LastName        = '" & vEcom_LastName                   & "', " 
    vSql = vSql & " Ecom_CardName        = '" & vEcom_CardName                   & "', " 
    vSql = vSql & " Ecom_Address         = '" & vEcom_Address                    & "', " 
    vSql = vSql & " Ecom_City            = '" & vEcom_City                       & "', " 
    vSql = vSql & " Ecom_Postal          = '" & vEcom_Postal                     & "', " 
    vSql = vSql & " Ecom_Province        = '" & vEcom_Province                   & "', " 
    vSql = vSql & " Ecom_Country         = '" & vEcom_Country                    & "', " 
    vSql = vSql & " Ecom_Phone           = '" & vEcom_Phone                      & "', " 
    vSql = vSql & " Ecom_Email           = '" & vEcom_Email                      & "', " 
    vSql = vSql & " Ecom_Quantity        =  " & vEcom_Quantity                   & " , " 
    vSql = vSql & " Ecom_NewAcctId       = '" & vEcom_NewAcctId                  & "', " 
    vSql = vSql & " Ecom_Media           = '" & vEcom_Media                      & "', " 
    vSql = vSql & " Ecom_Label           = '" & vEcom_Label                      & "', " 
    vSql = vSql & " Ecom_Orderno         = '" & vEcom_Orderno                    & "', " 
    vSql = vSql & " Ecom_OrderId         = '" & vEcom_OrderId                    & "', " 
    vSql = vSql & " Ecom_LineId          = '" & vEcom_LineId                     & "', " 
    vSql = vSql & " Ecom_Shipping        =  " & fDefault(vEcom_Shipping, 0)      & " , " 
    vSql = vSql & " Ecom_Memo            = '" & vEcom_Memo                       & "', " 
    vSql = vSql & " Ecom_Organization    = '" & vEcom_Organization               & "', " 
    vSql = vSql & " Ecom_Source          = '" & vEcom_Source                     & "', " 
    vSql = vSql & " Ecom_InternetSecure  = '" & vEcom_InternetSecure             & "', " 
    vSql = vSql & " Ecom_Adjustment      =  " & fDefault(vEcom_Adjustment, 0)    & " , " 
    vSql = vSql & " Ecom_Agent           = '" & vEcom_Agent                      & "'  " 

    vSql = vSql & " WHERE Ecom_No        =  " & vEcom_No  
   'sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Function fIsNewEcom
    fIsNewEcom = True
    If Len(vEcom_InternetSecure) = 0 Then Exit Function
    vSql = "SELECT * FROM Ecom WHERE Ecom_InternetSecure = '" & vEcom_InternetSecure & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      fIsNewEcom = False
    End If
    Set oRs = Nothing
    sCloseDb  
  End Function

  Sub sUpdateEcomLabel
    vSql = "UPDATE Ecom SET Ecom_Label = '" & fUnquote(vEcom_Label) & "' WHERE Ecom_OrderNo = '" & vEcom_OrderNo & "'"
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Sub sDeleteEcom
    sOpenDb
    vSql = "DELETE FROM Ecom WHERE Ecom_No= " & vEcom_No
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  '...note, this has been disabled
  Sub sDeleteEcomByMembNo
    sOpenDb
    vSql = "DELETE FROM Ecom WHERE Ecom_MembNo= " & vMemb_No
'   sDebug
'   oDb.Execute(vSql)
    sCloseDb
  End Sub

  Function fRepurchased(vProg_Id)
    fRepurchased = False
    If fNoValue(svMembNo) Then Exit Function '...if from public site there is no user
    '...determine if repurchase option is for this program or any program during a specific period
    If vCust_EcomRepurPrograms Then '...all programs
      If vCust_EcomRepurPeriod = 0 Then
        vSql = "SELECT * FROM Ecom WHERE (Ecom_MembNo = " & svMembNo & ")"
      Else
        vSql = "SELECT * FROM Ecom WHERE (Ecom_MembNo = " & svMembNo & ") AND (DATEDIFF([day], Ecom_Issued, GETDATE()) < " & vCust_EcomRepurPeriod & ") "
      End If
    Else '...this programs
      If vCust_EcomRepurPeriod = 0 Then
        vSql = "SELECT * FROM Ecom WHERE (Ecom_MembNo = " & svMembNo & ") AND (CHARINDEX('" & vProg_Id & "', Ecom_Programs) > 0)"
      Else
        vSql = "SELECT * FROM Ecom WHERE (Ecom_MembNo = " & svMembNo & ") AND (DATEDIFF([day], Ecom_Issued, GETDATE()) < " & vCust_EcomRepurPeriod & ") AND (CHARINDEX('" & vProg_Id & "', Ecom_Programs) > 0)"
      End If
    End If
'   sDebug              
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fRepurchased = True
    sCloseDb2
  End Function
  
  
  
  
  '...this creates the dropdown for the number of programs ordered and used when the group 2 ecom site was purchased
  Function fEcomGroupProgs (vMembPrograms)
    Dim vCount, vOrdered, vAssigned
    vCount = 0
    fEcomGroupProgs = ""

    sOpenDb
    sOpenDb2

    '...in case adjustments were made, take the sum of the Quantity per Program 
'   vSql = "SELECT Ecom_Programs, SUM(Ecom_Quantity) AS Ecom_Quantity FROM Ecom WHERE Ecom_NewAcctId = '" & svCustAcctId & "' AND Ecom_Quantity > 0 GROUP BY Ecom_Programs ORDER BY Ecom_Programs"
    vSql = "SELECT Ecom_Programs, SUM(CASE WHEN Ecom_Amount < 0 THEN Ecom_Quantity * -1 ELSE Ecom_Quantity END) AS Ecom_Quantity FROM Ecom WHERE Ecom_NewAcctId = '" & svCustAcctId & "' AND Ecom_Archived IS NULL GROUP BY Ecom_Programs ORDER BY Ecom_Programs"
'   sDebug
    
    Set oRs = oDb.Execute(vSql)

    Do While Not oRs.Eof
        
      vEcom_Programs = oRs("Ecom_Programs")

      vSql =        " SELECT COUNT(Memb.Memb_Programs) AS Assigned FROM Memb "
      vSql = vsql & " WHERE (Memb.Memb_AcctId = '" & svCustAcctId & "') "
      vSql = vsql & " AND (CHARINDEX('" & vEcom_Programs & "', Memb.Memb_Programs) > 0)"

      Set oRs2 = oDb2.Execute(vSql)
      vAssigned = oRs2("Assigned")

      vEcom_Quantity = oRs("Ecom_Quantity") - vAssigned
'     vEcom_Expires  = oRs("Ecom_Expires")


      '...only display programs if not already assigned
      If Instr(vMembPrograms, vEcom_Programs) = 0 And vEcom_Quantity > 0 Then
        vCount = vCount + 1
        i = "            <option value=" & Chr(34) & vEcom_Programs & Chr(34) & ">" & vEcom_Programs & "  (available: " & Right("0000" & vEcom_Quantity, 4) & ") - " & fProgTitle (vEcom_Programs) & "</option>" & vbCrLf
        fEcomGroupProgs = fEcomGroupProgs & i
      End If
      oRs.MoveNext	        
    Loop
    Set oRs  = Nothing      
    Set oRs2 = Nothing      
    sCloseDb    
    sCloseDb2    
    
    If vCount > 0 Then
      fEcomGroupProgs = vbCrLf & "<select  style='width: 100%' name='vPrograms' multiple size='" & vCount & "'>" & fEcomGroupProgs & "          </select>"
      fEcomGroupProgs = fEcomGroupProgs  & vbCrLf 
    End If
  End Function 

%>