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
  '...this generates the files for an online sales or group2/addon2 sale
  Dim vSource, vTest, aCatlNo, aPrograms, aPrices, aQuantity, aTaxes, vExpires, aExpires, aAmounts, vTotalPrices, vTotalAmount, vMaxUsers
  
    '...determine if test or live
  vTest = fDefault(Request("vEcom_Test"), "y")

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

  '...get this record (mainly for the vCust_ParentId for AddOn2)
  sGetCust svCustId
  
  '...this goes on the learner/facilitator record
  aExpires        = Split(vEcom_Expires, "|")
  vExpires        = aExpires(0)
   
  '...get max users based on the quantity of seats sold (this is also split later down)
  aQuantity = Split(vEcom_Quantity, "|")
  vMaxUsers = 0
  For i = 0 To Ubound(aQuantity)
    vMaxUsers = vMaxUsers + aQuantity(i)
  Next
  vMaxUsers = vMaxUsers * -1       '...Group2 goes by seats not users so negate


  '...create new account and catalogue for Group2 
  If vEcom_Media = "Group2" Then   

    '...generate new customer account id passing current id to be cloned, plus the max no users
    vEcom_NewAcctId = fNextAcctId
    vMemb_AcctId    = vEcom_NewAcctId
    vEcom_MembNo    = fNextMembNo (vEcom_NewAcctId)
    vEcom_Id        = vMemb_Id

    '...create the new customer record
    sCloneCust vEcom_CustId, vEcom_NewAcctId, vMaxUsers, vEcom_Id, vEcom_Programs, vExpires
  
    '...create the new catalgue
    sCreateCatl vEcom_CustId, Left(vEcom_CustId, 4) & vEcom_NewAcctId, vEcom_CatlNo, vEcom_Programs
      
    '...update the ecom table - below (do this after updating the Cust and Catl)
    sUpdateGroupEcom

    '...update the member table with the current member (facilitator), plus a manager and admnistrator
    vMemb_AcctId    = vEcom_NewAcctId
    vMemb_Expires   = vExpires '...all learners can access the content until the site expires
    vMemb_No = 0 :  vMemb_Id        = vEcom_Id      : vMemb_Level = 3     : sAddMemb vMemb_AcctId

    '...add internals
    sAddInternalMemb vMemb_AcctId

    '...add a client manager  (added vMemb_Pwd Mar 6 2016 to better handle NOP and Portal access)
'   vMemb_Internal = 0 : vMemb_No = 0 :  vMemb_Id = Left(vEcom_CustId, 4) & "_SALES" : vMemb_Level = 4 : vMemb_Manager = 1 : vMemb_Ecom = 1 : sAddMemb vMemb_AcctId
    vMemb_Internal = 0 : vMemb_No = 0 :  vMemb_Id = Left(vEcom_CustId, 4) & "_SALES" : vMemb_Level = 4 : vMemb_Manager = 1 : vMemb_Ecom = 1 : vMemb_Pwd = "BIGMGR" : sAddMemb vMemb_AcctId

  '...for AddOn2 just update Ecom/Cust/Catl
  Else

    vEcom_CustId    = Left(vCust_Id, 4) & vCust_ParentId
    vEcom_AcctId    = vCust_ParentId
    vEcom_MembNo    = svMembNo
    vEcom_NewAcctId = vCust_AcctId

    '...update the ecom table - below (do this after updating the Cust and Catl)
    sUpdateGroupEcom

    '...update the customer record
    sUpdateCustMaxUsers vCust_Id, vMaxUsers
    sUpdateCustExpires  vCust_AcctId, DateAdd("yyyy", 1, Now())

    '...update the catalogue
    sRecreateCatl Left(vCust_Id, 4) & vCust_ParentId, vCust_Id

  End If



  '...Store in Session variables so not visible on url
  If vEcom_Media = "Group2" Then
    Session("EcomCust")   = Left(vEcom_CustId, 4) & vEcom_NewAcctId
  ElseIf vEcom_Media = "AddOn2" Then
    Session("EcomCust")   = svCustId
  End If

  Session("EcomId")     = vEcom_Id
  Session("EcomIssued") = True

  Response.Redirect "Ecom3DisplayIds.asp?vClose=Y"


  '...this creates the dropdown for the number of programs ordered and used when the group 2 ecom site was purchased
  Function fEcomGroupProgs (vMembPrograms)
    Dim vCount, vOrdered, vAssigned
    vCount = 0
    fEcomGroupProgs = ""

    sOpenDb
    sOpenDb2

    '...in case adjustments were made, take the sum of the Quantity per Program 
    vSql = "SELECT Ecom_Programs, SUM(Ecom_Quantity) AS Ecom_Quantity FROM Ecom WHERE Ecom_NewAcctId = '" & svCustAcctId & "' AND Ecom_Quantity > 0 GROUP BY Ecom_Programs ORDER BY Ecom_Programs"
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


      '...only display programs if not already assigned
      If Instr(vMembPrograms, vEcom_Programs) = 0 And vEcom_Quantity > 0 Then
        vCount = vCount + 1
        i = "            <option value=" & Chr(34) & vEcom_Programs & Chr(34) & ">" & vEcom_Programs & "  (available: " & Right("000" & vEcom_Quantity, 3) & ") - " & fLeft(fProgTitle (vEcom_Programs), 48) & "</option>" & vbCrLf
        fEcomGroupProgs = fEcomGroupProgs & i
      End If
      oRs.MoveNext	        
    Loop
    Set oRs  = Nothing      
    Set oRs2 = Nothing      
    sCloseDb    
    sCloseDb2    
    
    If vCount > 0 Then
      fEcomGroupProgs = vbCrLf & "<select name='vPrograms' multiple size='" & vCount & "'>" & fEcomGroupProgs & "          </select>"
      fEcomGroupProgs = fEcomGroupProgs  & vbCrLf 
    End If
  End Function   




  '...called when there is a collection of programs and catalogue values to be updated (a bit different than the one used in Ecom2GenerateId.asp)
  Sub sUpdateGroupEcom ()

    vEcom_Issued   = fFormatSqlDate (Now)
  
    '...update Ecom audit table for each program since each program might have a different expiry date
    '   multiple programs are separated by spaces as are prices and expires
    If Instr(vEcom_Programs, "|") = 0 Then
      
      sAddEcom
  
    Else
  
      aCatlNo   = Split(vEcom_CatlNo, "|")
      aPrograms = Split(vEcom_Programs, "|")
      aPrices   = Split(vEcom_Prices, "|")
      aQuantity = Split(vEcom_Quantity, "|")
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
          vEcom_Quantity = aQuantity(i)
          vEcom_CatlNo   = aCatlNo(i)
          vEcom_Programs = aPrograms(i)
          vEcom_Prices   = aPrices(i)
          vEcom_Expires  = aExpires(i)
          '...split the total proportionately
          vEcom_Amount   = vTotalAMount * vEcom_Prices / vTotalPrices      
          sAddEcom
        Next
  
      '...general ecom record for all modern sales, one per product but only put quantity (seat) on first product
      Else
  
        '...get the values for the ecom table
        For i = 0 To Ubound(aPrograms)
          vEcom_Quantity = aQuantity(i)
          vEcom_CatlNo   = aCatlNo(i)
          vEcom_Programs = aPrograms(i)
          vEcom_Prices   = Ccur(aPrices(i))
          vEcom_Taxes    = Ccur(aTaxes(i))
          vEcom_Expires  = aExpires(i)
          vEcom_Amount   = vEcom_Prices + vEcom_Taxes
          vTotalPrices   = vTotalPrices + aPrices(i)
          sAddEcom
        Next
  
      End If
  
    End If
  
  End Sub
    
%>

