<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<% 
  Server.ScriptTimeout = 60 * 20

  Dim bUpdateOnline, bUpdateGroup2, bTest, bCommit
  Dim vMaxUsers, vTotal_Amount, vStatus, vProgramCnt, vGroupCustId, vGroupId, vCnt

  Dim aCatlNo, aPrograms, aPrice, aQuantity, aPrices, aTaxes, aAmount
  Dim vPrice, vGST, vPST, vHST, aGST, aPST, aHST
  Dim vTot_Quantity, vTot_Prices, vTot_GST, vTot_PST, vTot_HST, vTot_Amount 
  Dim vTmp_Quantity, vTmp_GST, vTmp_PST, vTmp_HST, vTmp_Amount      '...temp fields for totals

  '...is this a test? if so display results locally rather that to the web service
  bTest = fIf(Lcase(Request("vTest")) = "y", True, False)

  '...is this a validate or commit?
  bCommit = fIf(Lcase(Request("vAction")) = "v", False, True)

  '...valid entry?
' If (Request("vEcom_Media") <> "Online") Then
'   vStatus = "499 Group and AddOn Transactions are temporarily suspended."
' ElseIf Len(Request("vEcom_Media")) = 0 Then

  If Len(Request("vEcom_Media")) = 0 Then
    vStatus = "468 Service was accessed without a Transaction Type."
  ElseIf Instr("Online Group2 AddOn2", Request("vEcom_Media")) = 0 Then
    vStatus = "469 Service was accessed with an incorrect Transaction Type."
  Else
    vStatus = fPost '...get status from application
  End If

  '...return status info, display is testing
  If vStatus <> "" Then '...send error message
    If bTest Then
      Response.Write vStatus & "<br>"
    Else
      Response.Status = vStatus
    End If
  Else '...send success plus values
    '...return value to web service

    If bCommit Then
      vStatus = "200 Successfully Posted"
    Else      
      vStatus = "200 Transaction Validated"
      '...put in dummy data for validate only
      vEcom_CustId    = "XXXX0000"
      vEcom_NewAcctId = "0000"
      vEcom_Id        = "Password"
      vEcom_Expires   = "Jan 01, 2000"      
    End If

    If bTest Then
      Response.Write vStatus & "<br>"
    Else
      Response.Status = vStatus
    End If
  

    If vEcom_Media = "Group2" And Not bUpdateGroup2 Then
      Response.AddHeader "Account Id", Left(vEcom_CustId, 4) & vEcom_NewAcctId
      Response.AddHeader "Password", vEcom_Id
      Response.AddHeader "Expiry Date", vEcom_Expires
      If bTest Then 
        Response.Write """Account Id""" & ",""" & Left(vEcom_CustId, 4) & vEcom_NewAcctId & """<br>"
        Response.Write """Password""" & ",""" & vEcom_Id & """<br>"
        Response.Write """Expiry Date""" & ",""" & vEcom_Expires & """<br>"
      End If  
    ElseIf vEcom_Media = "Group2" And bUpdateGroup2 Then
      Response.AddHeader "Expiry Date", vEcom_Expires
      If bTest Then 
        Response.Write """Expiry Date""" & ",""" & vEcom_Expires & """<br>"
      End If  
    '...return a password for individual sales unless one is passed through (bUpdateOnline = True)
    ElseIf Not bUpdateOnline And Not bUpdateGroup2 Then
      Response.AddHeader "Password", vEcom_Id
      If bTest Then Response.Write """Password""" & ",""" & vEcom_Id & """<br>"
      '...currently returning 90 days from today - may need to modify if they sell courses that expire in other than 90 days (ie 365 days) 
      Response.AddHeader "Expiry Date", vEcom_Expires
      If bTest Then 
        Response.Write """Expiry Date""" & ",""" & vEcom_Expires & """<br>"
      End If  
    ElseIf bUpdateOnline And Not bUpdateGroup2 Then
      '...currently returning 90 days from today - may need to modify if they sell courses that expire in other than 90 days (ie 365 days) 
      Response.AddHeader "Expiry Date", vEcom_Expires
      If bTest Then 
        Response.Write """Expiry Date""" & ",""" & vEcom_Expires & """<br>"
      End If  
    End If  

  End If


  '...this is called by the web service - note that certain values are collections and need to be separated by pipes
  Function fPost

    fPost = ""  

    '...when live data comes from a form post ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '...when testing it comes via the URL     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'  If Request.QueryString.Count = 0 Then fPost = "444 The service did not receive any QueryString values to process." : Exit Function
'  For Each vFld In Request.QueryString  '...use for testing with big URL (From Bryan)

   If Request.Form.Count = 0 Then fPost = "444 The service did not receive any Form values to process." : Exit Function
   For Each vFld In Request.Form

      vValue = fUnQuote(Request(vFld))
//    vValue = fUrlDecode(vValue) '...added Mar 1, 2016 to handle Bryan's WS which does a URLencode but uses forms - doesn't work with Accents

      Select Case vFld
        Case "vEcom_CustId"         : vEcom_CustId         = Ucase(vValue)
        Case "vEcom_Id"             : vEcom_Id             = Ucase(vValue) '...if used then update existing online account

        Case "vGroupCustId"         : vGroupCustId         = Ucase(vValue) '...if used then update existing group account
        Case "vGroupId"             : vGroupId             = Ucase(vValue) '...if used then update existing group account

        Case "vEcom_Programs"       : vEcom_Programs       = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vPrice"               : vPrice               = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vEcom_Quantity"       : vEcom_Quantity       = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vEcom_Prices"         : vEcom_Prices         = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vGST"                 : vGST                 = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vPST"                 : vPST                 = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vHST"                 : vHST                 = Replace(Replace(vValue, ", ", ","), ",", "|")
        Case "vEcom_Amount"         : vEcom_Amount         = Replace(Replace(vValue, ", ", ","), ",", "|")

        Case "vTot_Quantity"        : vTot_Quantity        = Clng(fDefault(vValue, 0))
        Case "vTot_Prices"          : vTot_Prices          = Ccur(fDefault(vValue, 0))
        Case "vTot_GST"             : vTot_GST             = Ccur(fDefault(vValue, 0))
        Case "vTot_PST"             : vTot_PST             = Ccur(fDefault(vValue, 0))
        Case "vTot_HST"             : vTot_HST             = Ccur(fDefault(vValue, 0))
        Case "vTot_Amount"          : vTot_Amount          = Ccur(fDefault(vValue, 0))

        Case "vEcom_Memo"	          : vEcom_Memo           = Replace(Replace(vValue, ", ", ","), ",", "|")

        Case "vEcom_Currency"       : vEcom_Currency       = vValue
        Case "vEcom_Lang"           : vEcom_Lang           = vValue
        Case "vEcom_Media"          : vEcom_Media          = vValue

        Case "vEcom_FirstName"      : vEcom_FirstName      = vValue
        Case "vEcom_LastName"       : vEcom_LastName       = vValue
        Case "vEcom_Email"          : vEcom_Email          = vValue

        Case "vEcom_CardName"       : vEcom_CardName       = vValue
        Case "vEcom_Address"        : vEcom_Address        = vValue
        Case "vEcom_City"           : vEcom_City           = vValue
        Case "vEcom_Postal"         : vEcom_Postal         = vValue
        Case "vEcom_Province"       : vEcom_Province       = vValue
        Case "vEcom_Country"        : vEcom_Country        = vValue
        Case "vEcom_Phone"          : vEcom_Phone          = vValue
        Case "vEcom_Organization"		: vEcom_Organization   = vValue

        Case "vEcom_Source"		      : vEcom_Source         = fDefault(vValue, "C")  // CCOHS, etc will not send anything thus assume "C", NOP will send "E"

      End Select  
    Next  

    vEcom_Shipping       = 0  
    vEcom_Issued         = fFormatSqlDate (Now)

    vMemb_FirstName			 = fUnquote(vEcom_FirstName)
    vMemb_LastName       = fUnquote(vEcom_LastName)
    vMemb_Email          = vEcom_Email
    vMemb_Organization   = vEcom_Organization

    '...valid Customer?
    If Len(vEcom_CustId) <> 8 Then 
      fPost = "461 No Customer ID has been sent." : Exit Function
    Else
      sGetCust vEcom_CustId
      If vCust_Eof Then
        fPost = "462 Customer ID is invalid." : Exit Function
      End If  
    End If

    '...confirm receipt of at least one program
    i = Trim(Replace(vEcom_Programs, "|", ""))
    If Len(i) < 7 Then
      fPost = "467 No Programs have been selected." : Exit Function
    End If
  
    vMaxUsers            = -1    ' (no contraints for group2 but need this for My Content)  
    vEcom_AcctId         = Right(vEcom_CustId, 4)
    vProgramCnt          = -1    '...this is number of programs that were received (base 0)

    bUpdateOnline        = fIf(Len(vEcom_Id) > 0, True, False)

    '...check the Group2 AddOn2 values
    If vEcom_Media = "AddOn2" Then
      If Len(vGroupCustId) > 0 and Len(vGroupId) > 0 Then

        If Not fCustOk(vGroupCustId)   Then fPost = "470 You are trying to Add On to a Group site that does not exist."       : Exit Function
        If vGroupCustId = vEcom_CustId Then fPost = "445 You are trying to Add On to a Parent site rather than a Group site." : Exit Function
        If Not fCustG2Ok(vGroupCustId) Then fPost = "446 You are trying to Add On to a Group site that has not been purchased." : Exit Function

        '...if the GroupId is not on file then it might have been changed, so see if you can use the one submitted, which should be the original works
        sGetMembById Right(vGroupCustId, 4), vGroupId

        If vMemb_Eof                   Then fPost = "471 You are trying to Add On to a Group site with a password that does not exist or is not assigned to a facilitator." : Exit Function   
        If vMemb_Level <> 3            Then fPost = "471 You are trying to Add On to a Group site with a password that does not exist or is not assigned to a facilitator." : Exit Function   

        bUpdateGroup2      = True
        vEcom_MembNo       = vMemb_No
        vEcom_Id           = vMemb_Id
        vEcom_FirstName    = fUnquote(vMemb_FirstName)
        vEcom_LastName     = fUnquote(vMemb_LastName)
        vEcom_Email        = vMemb_Email
        vEcom_NewAcctId    = Right(vGroupCustId, 4)
        vEcom_Expires      = fFormatSqlDate (Now + 365)
        vEcom_Media        = "Group2"     '...rename back

      Else
        fPost = "472 You are trying to Add On to a Group site without including the Group CustId and/or the Group Password." : Exit Function
      End If
    Else
      bUpdateGroup2        = False
    End If
   
    '...get the original catalogue value based on the programs selected (normally passed in via vubiz ecommerce)
    aPrograms = Split(vEcom_Programs, "|")
    vEcom_CatlNo = ""
    j = 0 '... temp CatlNo
    For vCnt = 0 To Ubound(aPrograms)
      If Len(aPrograms(vCnt)) = 7 Then
        If fProgOk (aPrograms(vCnt)) Then
          j = fCatlNo(vEcom_CustId, aPrograms(vCnt))
          If j = 0 Then 
            fPost = "466 Line item " & vCnt + 1 & " contained a Program Id (" & aPrograms(vCnt) & ") that is not in the current catalogue." : Exit Function
          End If
          vEcom_CatlNo = vEcom_CatlNo & fIf(vCnt > 0, "|" & j, j)
          vProgramCnt = vCnt
        Else
          fPost = "473 Line item " & vCnt + 1 & " contained a invalid Program Id (" & aPrograms(vCnt) & ")." : Exit Function
        End If
      Else
        Exit For
      End If
    Next


    '...if there is only one program ordered, ensure there are not multiple other values
    If vProgramCnt = 0 Then    
      If Instr(vEcom_Programs, "|") > 0 Then vEcom_Programs = Left(vEcom_Programs, Instr(vEcom_Programs, "|") - 1)        
      If Instr(vPrice, "|")         > 0 Then vPrice         = Ccur(Left(vPrice, Instr(vPrice, "|") - 1))			            Else vPrice	= Ccur(vPrice)
      If Instr(vEcom_Quantity, "|") > 0 Then vEcom_Quantity = Clng(Left(vEcom_Quantity, Instr(vEcom_Quantity, "|") - 1))	Else vEcom_Quantity	= Clng(vEcom_Quantity)
      If Instr(vEcom_Prices, "|")   > 0 Then vEcom_Prices   = Ccur(Left(vEcom_Prices, Instr(vEcom_Prices, "|") - 1))		  Else vEcom_Prices	= Ccur(vEcom_Prices)
      If Instr(vEcom_Amount, "|")   > 0 Then vEcom_Amount   = Ccur(Left(vEcom_Amount, Instr(vEcom_Amount, "|") - 1))		  Else vEcom_Amount	= Ccur(vEcom_Amount)
      If Instr(vGST, "|")           > 0 Then vGST           = Ccur(Left(vGST, Instr(vGST, "|") - 1))				              Else vGST	= Ccur(vGST)
      If Instr(vPST, "|")           > 0 Then vPST           = Ccur(Left(vPST, Instr(vPST, "|") - 1))				              Else vPST	= Ccur(vPST)
      If Instr(vHST, "|")           > 0 Then vHST           = Ccur(Left(vHST, Instr(vHST, "|") - 1))				              Else vHST	= Ccur(vHST)
    End If

    '...for individual sale, create a user Id
    If vEcom_Media = "Online" Then    
     
      If Len(vEcom_Id) > 0 Then
        sGetMembById vEcom_AcctId, vEcom_Id
        '...if not on file then add member (assuming auto-enroll is allowed)
        If vMemb_Eof Then
          If vCust_Auto Then
            sMemb_Empty
            vMemb_Eof            = True
            vMemb_FirstName			 = fUnquote(vEcom_FirstName)
            vMemb_LastName       = fUnquote(vEcom_LastName)
            vMemb_Email          = vEcom_Email
            vMemb_Organization   = vEcom_Organization
            vMemb_AcctId         = vEcom_AcctId
            vMemb_No             = 0
            vMemb_Id             = vEcom_Id
            sUpdateMemb vEcom_AcctId

          Else
            fPost = "486 Password is not on file." : Exit Function
          End If

        Else  

          If Len(vEcom_FirstName)    > 0 Then vMemb_FirstName    = fUnquote(vEcom_FirstName)
          If Len(vEcom_LastName)     > 0 Then vMemb_LastName     = fUnquote(vEcom_LastName)
          If Len(vEcom_Email)        > 0 Then vMemb_Email        = vEcom_Email
          If Len(vEcom_Organization) > 0 Then vMemb_Organization = vEcom_Organization          
          sUpdateMemb vEcom_AcctId

        End If
        vEcom_MembNo             = vMemb_No
      Else
        vMemb_Eof                = True
        vMemb_AcctId             = vEcom_AcctId
        If bCommit Then 
          vEcom_MembNo           = fNextMembNo (vEcom_AcctId)
        End If        
        vEcom_Id                 = vMemb_Id
      End If  
'     vEcom_Expires              = fFormatSqlDate (Now + 90)  

    '...if Group2 (but NOT an AddOn) create new customer, catalogue and add facilitator and vu team plus the local support manager
    ElseIf Not bUpdateGroup2 Then

      '...generate new customer account id passing current id to be cloned, plus the max no users
      If bCommit Then
        vEcom_NewAcctId = fNextAcctId
        vMemb_AcctId    = vEcom_NewAcctId
        vEcom_MembNo    = fNextMembNo (vEcom_NewAcctId)
        vEcom_Id        = vMemb_Id
        vEcom_Expires   = fFormatSqlDate (Now + 365)
      End If

      '...create the new customer record
      If bCommit Then
        sCloneCust vEcom_CustId, vEcom_NewAcctId, vMaxUsers, vEcom_Id, vEcom_Programs, vEcom_Expires  
      End If

      '...update the member table with the current member (facilitator), plus a manager and admnistrator
      vMemb_AcctId    = vEcom_NewAcctId
      vMemb_Expires   = vEcom_Expires '...all learners can access the content until the site expires
      vMemb_No = 0 :  vMemb_Id        = vEcom_Id      : vMemb_Level = 3     : If bCommit Then sAddMemb vMemb_AcctId

      If bCommit Then
        sAddInternalMemb vMemb_AcctId '...add internals
        vMemb_Internal = 0 : vMemb_No = 0 :  vMemb_Id = Left(vEcom_CustId, 4) & "_SALES" : vMemb_Level = 4 : vMemb_Manager = 1 : vMemb_Ecom = 1 : sAddMemb vMemb_AcctId   '...add a client manager
      End If
     
    ElseIf bUpdateGroup2 Then

      '...update the expiry date on the customer table
      vEcom_Expires   = fFormatSqlDate (Now + 365)
      If bCommit Then
        sUpdateCustExpires vEcom_NewAcctId, vEcom_Expires
      End If
          
    End If


  
    '...update Ecom table with each program ordered (each program might have a different expiry date)
    '   multiple programs are separated by pipes as are prices, etc
    If vProgramCnt = 0 Then
      
      vEcom_Taxes    = Ccur(vGST) + Ccur(vPST) + Ccur(vHST)
      If vEcom_Prices <> vPrice * vEcom_Quantity Then fPost = "474 Line item 1 is not extended properly." : Exit Function
      If vEcom_Amount <> vEcom_Prices + vEcom_Taxes Then fPost = "475 Line item 1 is not totalled properly." : Exit Function
      If vEcom_Media = "Online" And vEcom_Quantity  <> 1 Then fPost = "476 Line item 1 must have a quantity of 1." : Exit Function

      If vEcom_Media  = "Online" Then
        vEcom_Expires = fFormatSqlDate (Now + fCatlExpires(vEcom_CustId, vEcom_Programs))  
      End If        

      If bCommit Then 
        sAddEcom
      End If
  
    Else
  
      aCatlNo   = Split(vEcom_CatlNo, "|")  
      aPrograms = Split(vEcom_Programs, "|")
      aPrice    = Split(vPrice, "|")
      aQuantity = Split(vEcom_Quantity, "|")
      aPrices   = Split(vEcom_Prices, "|")
      aGST      = Split(vGST, "|")
      aPST      = Split(vPST, "|")
      aHST      = Split(vHST, "|")
      aAmount   = Split(vEcom_Amount, "|")
  
      '...loop through to confirm all totals and programs are ok - not use vProgramCnt (generated at top) to avoid empty values
      For vCnt = 0 To vProgramCnt
        vEcom_Programs = aPrograms(vCnt)
        vPrice         = Ccur(aPrice(vCnt))
        vEcom_Quantity = Clng(aQuantity(vCnt))
        vEcom_Prices   = Ccur(aPrices(vCnt))
        vEcom_Taxes    = Ccur(aGST(vCnt)) + Ccur(aPST(vCnt)) + Ccur(aHST(vCnt))
        vEcom_Amount   = Ccur(aAmount(vCnt))
        vEcom_CatlNo   = Clng(aCatlNo(vCnt))

        If vEcom_Media = "Online" And vEcom_Quantity  <> 1 Then fPost = "477 Line item " & vCnt + 1 & " must have a quantity of 1." : Exit Function

        vTmp_Quantity  = vTmp_Quantity + vEcom_Quantity
        vTmp_GST       = vTmp_GST + Ccur(aGST(vCnt))
        vTmp_PST       = vTmp_PST + Ccur(aPST(vCnt))
        vTmp_HST       = vTmp_HST + Ccur(aHST(vCnt))
        vTmp_Amount    = vTmp_Amount + vEcom_Amount

        If vEcom_Prices <> vPrice * vEcom_Quantity Then fPost = "478 Line item " & vCnt + 1 & " is not extended properly." : Exit Function
        If vEcom_Amount <> vEcom_Prices + vEcom_Taxes Then fPost = "479 Line item " & vCnt + 1 & " is not totalled properly." : Exit Function
        If vEcom_CatlNo = 0 Then fPost = "480 Program " & vEcom_Programs & " in line item " & vCnt + 1 & " is not in the catalogue."
      Next  

      If vTmp_Quantity <> vTot_Quantity Then fPost = "481 Quantities are not totalled properly." : Exit Function
      If vTmp_GST      <> vTot_GST Then fPost = "482 GST is not total properly." : Exit Function
      If vTmp_PST      <> vTot_PST Then fPost = "483 PST is not total properly." : Exit Function
      If vTmp_HST      <> vTot_HST Then fPost = "484 HST is not total properly." : Exit Function
      If vTmp_Amount   <> vTot_Amount Then fPost = "485 Extensions are not totalled properly." : Exit Function

      '...post line items into the ecom table
      For vCnt = 0 To vProgramCnt
        vEcom_Programs  = aPrograms(vCnt)
        vEcom_Quantity  = aQuantity(vCnt)
        vEcom_Prices    = Ccur(aPrices(vCnt))
        vEcom_Taxes     = Ccur(aGST(vCnt)) + Ccur(aPST(vCnt)) + Ccur(aHST(vCnt))
        vEcom_Amount    = Ccur(aAmount(vCnt))
        vEcom_CatlNo    = fCatlNo(vEcom_CustId, vEcom_Programs) '...store the catl no (find which group (catl no) this programs belongs to)
        If vEcom_Media  = "Online" Then
          vEcom_Expires = fFormatSqlDate (Now + fCatlExpires(vEcom_CustId, vEcom_Programs))  
        End If        
        If bCommit Then
          sAddEcom
        End If
      Next
  
    End If 
    
    '...prepare values for Member table
    vMemb_No            = vEcom_MembNo
    vMemb_AcctId        = fDefault(vEcom_NewAcctId, vEcom_AcctId)
    vMemb_Level         = fDefault(vMemb_Level, 2) '...if new member else pick up previous level
    vMemb_Id            = vEcom_Id
    vMemb_FirstName     = vMemb_FirstName
    vMemb_LastName      = vMemb_LastName
    vMemb_Email         = vMemb_Email
    vMemb_Organization  = vEcom_Organization
  
    '...add to member table unless member ordered a course previously
    '   if group ecom, then the memember (facilitator), plus the manager and administrator are added after we build the new customer record
  
    '...finally, create the catalogue for the group sales and/or addons
    '... for some reason we called AddOn2 Group2 - coincidentally works ok here ????????????????

    If bCommit Then
      If vEcom_Media = "Online" Then
        sUpdateMemb vMemb_AcctId
      ElseIf vEcom_Media = "Group2" Then
        sRecreateCatl vEcom_CustId, Left(vEcom_CustId, 4) & vEcom_NewAcctId
      ElseIf vEcom_Media = "AddOn2" Then
        sRecreateCatl vEcom_CustId, vGroupCustId
      End If
    End If

  End Function  

%>