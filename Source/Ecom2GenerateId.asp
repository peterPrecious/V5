<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Debug_Routines.asp"-->

<!--#include virtual = "V5/Inc/Elavon.asp"-->

<% 
  '...this generates the files for an online sale or group sale
  Dim vSource, vTest, aCatlNo, aPrograms, aPrices, aTaxes, vExpires, aExpires, aAmounts, vTotalPrices, vTotalAmount
  
  '...determine if need to generate a new member Id
  Session("PassThru") = False

  If Len(Session("BypassEcom")) > 0 Then
    '...if bypassing Ecom then get SQL order No
    sGetSqlForm (Cint(Session("BypassEcom")))
    vEcom_InternetSecure = "Bypass"
    vEcom_OrderNo        = Session("BypassEcom")
  Else
    '...get the form values that were stored in SQL using the GUID
    If Request("ssl_result_message") <> "APPROVAL" Then
      Response.Redirect "EcomError.asp?vMsg=" & Server.UrlEncode("Elavon/Concierge Ecommerce Transaction NOT Approved.")   
    Else
      vEcom_InternetSecure = Request("ssl_txn_id")
      vEcom_OrderNo        = Request("ssl_invoice_number")
      sGetSqlForm(vEcom_OrderNo)
    End If
  End If

  svCustAcctId = vEcom_AcctId

  '...if Member Id pass thru, see if already on file
  If Len(vEcom_Id) > 0 Then
    spMembNoById vEcom_AcctId, vEcom_Id, svMembNo
    vEcom_MembNo   = vMemb_No
    sGetMemb (vEcom_MembNo) '... this is new to elavon
  Else
    vEcom_MembNo   = fNextMembNo (vEcom_AcctId)
    vEcom_Id       = vMemb_Id
  End If

  vEcom_Issued     = fFormatSqlDate (Now)

  '...update Ecom audit table for each program since each program might have a different expires date
  '   multiple programs are separated by spaces as are prices and expires

  aCatlNo   = Split(vEcom_CatlNo, "|")
  aPrograms = Split(vEcom_Programs, "|")
  aPrices   = Split(vEcom_Prices, "|")
  aTaxes    = Split(vEcom_Taxes, "|")
  aExpires  = Split(vEcom_Expires, "|")

  '...get total program prices so the invoice total can be proportionately split into separate values (basically same unless tax added)
  vTotalPrices = 0
  vTotalAmount = vEcom_Amount    

  '...post values into ecom table
  For i = 0 To Ubound(aPrograms)
    vEcom_CatlNo   = aCatlNo(i)
    vEcom_Programs = aPrograms(i)
    vEcom_Prices   = Ccur(aPrices(i))
    vEcom_Taxes    = Ccur(aTaxes(i))
    vEcom_Expires  = aExpires(i)
    vEcom_Amount   = vEcom_Prices + vEcom_Taxes
    vTotalPrices   = vTotalPrices + aPrices(i)
    sAddEcom
  Next  

  '...Store in Session variables to display next
  Session("EcomCust") = vEcom_CustId
  Session("EcomId")     = vEcom_Id
  Session("EcomIssued") = True

  '...add to Member Table
  vMemb_No        = vEcom_MembNo
  vMemb_AcctId    = vEcom_NewAcctId
  vMemb_Id        = vEcom_Id
  vMemb_Level     = fDefault(vMemb_Level, 2) '...if new member else pick up previous level
  
  '...add/update to member table unless member ordered a course previously
  sAddMemb vMemb_AcctId

  '...display details to the customer
  Response.Redirect "EcomDisplayId.asp?vClose=Y&vSource=" & vSource

%>