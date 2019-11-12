<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Debug_Routines.asp"-->

<% 
  Dim vEmailDebug 
  vEmailDebug = ""

  '...set since bypassing "signin"
  Session("HostDb") = "V5_Vubz"

  '...get order info from InternetSecure form
  For Each vFld In Request.Form
    vValue = fUnQuote(Request.Form(vFld))
    Select Case vFld
      Case "vEcom_CustId"         : vEcom_CustId         = vValue
      Case "vEcom_AcctId"         : vEcom_AcctId         = vValue
      Case "vEcom_Id"             : vEcom_Id             = vValue
      Case "vEcom_Programs"       : vEcom_Programs       = vValue
      Case "vEcom_Prices"         : vEcom_Prices         = vValue
      Case "vEcom_Taxes"          : vEcom_Taxes          = vValue
      Case "vEcom_Expires"        : vEcom_Expires        = vValue
      Case "vEcom_Amount"         : vEcom_Amount         = vValue
      Case "vEcom_Currency"       : vEcom_Currency       = vValue
      Case "vEcom_Lang"           : vEcom_Lang           = vValue
      Case "vEcom_FirstName"      : vEcom_FirstName      = vValue
      Case "vEcom_LastName"       : vEcom_LastName       = vValue
      Case "vEcom_Quantity"       : vEcom_Quantity       = vValue
      Case "vEcom_Media"          : vEcom_Media          = vValue
      Case "vEcom_OrderNo"        : vEcom_OrderNo        = vValue
      Case "vEcom_Shipping"       : vEcom_Shipping       = vValue
      Case "receiptnumber"        : vEcom_InternetSecure = vValue
      Case "xxxName"              : vEcom_CardName       = vValue
      Case "xxxAddress"           : vEcom_Address        = vValue
      Case "xxxCity"              : vEcom_City           = vValue
      Case "xxxPostal"            : vEcom_Postal         = vValue
      Case "xxxProvince"          : vEcom_Province       = vValue
      Case "xxxCountry"           : vEcom_Country        = vValue
      Case "xxxPhone"             : vEcom_Phone          = vValue
      Case "xxxEmail"             : vEcom_Email          = vValue
      Case "vMemb_FirstName"      : vMemb_FirstName      = vValue
      Case "vMemb_LastName"       : vMemb_LastName       = vValue
      Case "vMemb_Email"          : vMemb_Email          = vValue
    End Select

    '...build form values for email tracking/debugging (stripping out any br tags)
    vEmailDebug = vEmailDebug & vFld & " = " & Replace(vValue, "<br>", " ") & "<br>"

  Next  
   
  '...this sends out a email to allison lee showing all parameters that were received from internet secure
  '   note, for further analysis this can be used to track parameters sent to internetsecure (ecomcheckout.asp)  
  '   don't email tests on supporting servers or if email is peterbulloch
  If Lcase(svHost) <> "localhost/v5" And Lcase(svHost) <> "peter/v5" Then
    sDebugByEmail "Ecom: " & vEcom_CustId & "-" & vEcom_Id , vEmailDebug    
  End If

  '...double issue
  If fIsNewEcom = False Then 
    Response.Redirect "EcomError.asp?vMsg=Dup"   
  End If

  '...invalid response from InternetSecure of user screwing around
  If Session("EcomIssued") = True Or Len(vEcom_Amount) = 0 Or Len(vEcom_Programs) = 0 Or Len(vEcom_CustId) = 0 Or Len(vEcom_Media) = 0 Then 
    Response.Redirect "EcomError.asp?vMsg=Err"   
  End If


  '...If individual online sale, always generate Ids, then check...
  '   if Member Id pass thru, see if already on file
  If vEcom_Media <> "CDs" Then
  
    If vEcom_Quantity = 1 Then   

      If Len(vEcom_Id) > 0 Then
        sGetMembById vEcom_AcctId, vEcom_Id
        If vMemb_Eof Then
          vEcom_MembNo = fNextMembNo (vEcom_AcctId)
        Else
          vEcom_MembNo = vMemb_No
        End If
      Else
        vMemb_Eof = True
        vEcom_MembNo = fNextMembNo (vEcom_AcctId)
        vEcom_Id = Right(10000000 + vEcom_MembNo, 7) & "-" & fSecurityCode(vEcom_CustId, vEcom_MembNo)
      End If

    '...if group facilitator
    Else

      '...generate new customer account id passing current id to be cloned, plus the max no users
      vEcom_NewAcctId = fNextAcctId
      vEcom_MembNo    = fNextMembNo (vEcom_NewAcctId)
      vEcom_Id        = Right(10000000 + vEcom_MembNo, 7) & "-" & fSecurityCode(vEcom_NewAcctId, vEcom_MembNo)

      '...create the new customer record
      sCloneCust vEcom_CustId, vEcom_NewAcctId, vEcom_Quantity, vEcom_Id, vEcom_Programs

      '...update the member table with the current member (facilitator), plus a manager and admnistrator
      vMemb_AcctId    = vEcom_NewAcctId
      vMemb_FirstName = vMemb_FirstName
      vMemb_LastName  = vMemb_LastName
      vMemb_Email     = vMemb_Email
      vMemb_No = 0 :  vMemb_Id        = vEcom_Id      : vMemb_Level = 3  :sAddMemb vMemb_AcctId
      '...clear out details for next two members
      vMemb_FirstName = "" : vMemb_LastName  = "" : vMemb_Email     = ""
      vMemb_No = 0 :  vMemb_Id        = vPassword3    : vMemb_Level = 3  :sAddMemb vMemb_AcctId
      vMemb_No = 0 :  vMemb_Id        = vPassword4    : vMemb_Level = 4  :sAddMemb vMemb_AcctId
      vMemb_No = 0 :  vMemb_Id        = vPassword5    : vMemb_Level = 5  :sAddMemb vMemb_AcctId

    End If

  Else
    '...force to 0 for CDs
    vEcom_MembNo = 0
  End If

  vEcom_Issued   = fFormatSqlDate (Now)

  '...update Ecom audit table for each program since each program might have a different expires date
  '   multiple programs are separated by spaces as are prices and expires
  If Instr(vEcom_Programs, "|") = 0 Then
    sAddEcom vMemb_AcctId

  Else
    Dim aPrograms, aPrices, aTaxes, aExpires, aAmounts, vTotalPrices, vTotalAmount
    aPrograms = Split(vEcom_Programs, "|")
    aPrices   = Split(vEcom_Prices, "|")
    aTaxes    = Split(vEcom_Taxes, "|")
    aExpires  = Split(vEcom_Expires, "|")

    '...get total program prices so the invoice total can be proportionately split into separate values (basically same unless tax added)
    vTotalPrices = 0
    vTotalAmount = vEcom_Amount    


    '...use this section for older sales that do not capture taxes
    If Len(vEcom_Taxes) = 0 Then

      For i = 0 To Ubound(aPrograms)
        vTotalPrices = vTotalPrices + aPrices(i)
      Next
      '...get the values for the ecom table
      For i = 0 To Ubound(aPrograms)
        vEcom_Programs = aPrograms(i)
        vEcom_Prices   = aPrices(i)
        vEcom_Expires  = aExpires(i)
        '...split the total proportionately
        vEcom_Amount   = vTotalAMount * vEcom_Prices / vTotalPrices      
        sAddEcom
      Next

    Else

      '...get the values for the ecom table
      For i = 0 To Ubound(aPrograms)
        vEcom_Programs = aPrograms(i)
        vEcom_Prices   = Ccur(aPrices(i))
        vEcom_Taxes    = Ccur(aTaxes(i))
        vEcom_Expires  = aExpires(i)
        vEcom_Amount   = vEcom_Prices + vEcom_Taxes
        '...if shipping CDs, put shipping on first program
        If i = 0 And IsNumeric(vEcom_Shipping) Then
          vEcom_Amount = vEcom_Amount + vEcom_Shipping
        End If
        If i > 0 And IsNumeric(vEcom_Shipping) Then
          vEcom_Shipping = 0
        End If
        
        vTotalPrices = vTotalPrices + aPrices(i)
        sAddEcom
      Next

    End If

  End If

  '...Store in Session variables so not visible on url
  If vEcom_Quantity = 1 Then
    Session("EcomCust")   = vEcom_CustId
  '...display new customer account
  Else
    Session("EcomCust")   = Left(vEcom_CustId,4) & vEcom_NewAcctId
  End If
  Session("EcomId")     = vEcom_Id
  Session("EcomIssued") = True
  
  '...add to Member Table
  vMemb_No        = vEcom_MembNo
  vMemb_AcctId    = fDefault(vEcom_NewAcctId, vEcom_AcctId)
  vMemb_FirstName = vMemb_FirstName
  vMemb_LastName  = vMemb_LastName
  vMemb_Email     = vMemb_Email
  vMemb_Level     = fDefault(vMemb_Level, 2) '...if new member else pick up previous level
  vMemb_Id        = vEcom_Id

  '...add to member table unless member ordered a course previously
  '   if group ecom, then the memember (facilitator), plus the manager and administrator are added after we build the new customer record
  If vEcom_Quantity = 1 Then
    If vMemb_Eof Then
      sAddMemb vMemb_AcctId
    Else
      sUpdateMemb
    End If
  End If

  '...display details to the customer
  If vEcom_Media = "CDs" Then
      Response.Redirect "EcomDisplayCd.asp?vEcom_CustId=" & vEcom_CustId & "&vEcom_FirstName=" & vEcom_FirstName & "&vEcom_LastName=" & vEcom_LastName & "&vEcom_Email=" & vEcom_Email & "&vEcom_Address=" & vEcom_Address & "&vEcom_City=" & vEcom_City& "&vEcom_Postal=" & vEcom_Postal & "&vEcom_Province=" & vEcom_Province & "&vEcom_Country=" & vEcom_Country & "&vEcom_OrderNo=" & vEcom_OrderNo & "&vClose=Y"
  Else
    If vEcom_Quantity = 1 Then
      Response.Redirect "EcomDisplayId.asp?vEcom_CustId=" & vEcom_CustId & "&vEcom_FirstName=" & vEcom_FirstName & "&vEcom_LastName=" & vEcom_LastName & "&vEcom_Email=" & vEcom_Email & "&vClose=Y"
    Else
      Response.Redirect "EcomDisplayIds.asp?vEcom_CustId=" & Left(vEcom_CustId,4) & vEcom_NewAcctId & "&vEcom_FirstName=" & vEcom_FirstName & "&vEcom_LastName=" & vEcom_LastName & "&vEcom_Email=" & vEcom_Email & "&vEcom_Quantity=" & vEcom_Quantity & "&vClose=Y"
    End If  
  End If
  
  '...Generate the security code (also in Db_Memb.asp) ... not used?
  Function fSecurityCodex (vCustNo, vMembNo)
    Dim vTemp
    Const cAlpha = "ABCDEFGHXY"
    vTemp = vMembNo * 4141
    vTemp = vMembNo * 4141 + Right(vCustNo, 4)
    vTemp = Right("0000" & vTemp, 4)
    fSecurityCode = ""
    For i = 1 To 4
      j = mid(vTemp,i,1)   
      k = mid(cAlpha, j+1, 1)
      fSecurityCode = fSecurityCode & k
    Next
  End Function

%>

