<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Cont.asp"-->

<% 
  '...ensure users and/or facilitators don't try to run this report by bypassing the menu page
  If svMembLevel < 4 Then Response.Redirect "Menu.asp"

  '...Excel variables
  Dim oWs, oCell, oStyleC, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 

  Dim vCustIdPrev, vPrograms, vStrDate, vEndDate, vNameCust, vNameUser, vFreebie, vWhere

  '...split values: owner %s come from prod table and cust % comes from the customer table
  '   for the admin report, aOwnr build summaries
  Dim vEcom_SplitVubz, vEcom_SplitCust, vEcom_SplitOwnr, aOwnr_CA(), aOwnr_US(), vOwnrCnt    
  Dim vSplitVubz, vSplitOwnr, vSplitCust, vAmount, vPrice,  vProgram

  Dim vMthStr, vMthEnd, vDate, vDateUrl, vOption1, vOption2, vDateMonth, vSelected, vExpires, vIssued, vCurrDate, vOwnerId, vSource
  Dim vOrganization, vAddress, vAddressInfo1, vAddressInfo2, vAddressInfo3, vAddressInfo4
  
  vPrograms = Trim(Request("vPrograms"))
  vAddress = Request("vAddress")
  vOwnerId = Request("vOwnerId")
  vStrDate = Request("vStrDate")
  vEndDate = Request("vEndDate")
  vFreebie = Request("vFreebie")
  vSource  = Request("vSource")


  Function fMedia
    Select Case vEcom_Media
      Case "CDs"       : fMedia = "CD "
      Case "Prods"     : fMedia = "PR "
      Case "Group"     : fMedia = "G1 "
      Case "Group2"    : fMedia = "G2 "
      Case "AddOn2"    : fMedia = "G2 "
      Case "Spec_01"   : fMedia = "S1 "
      Case Else        : fMedia = "IO "
    End Select
  End Function
  
  vCustIdPrev            = ""
  vOwnrCnt               = 0

  Redim Preserve aOwnr_CA(vOwnrCnt)
  Redim Preserve aOwnr_US(vOwnrCnt)

  '...access restriction rules (create the "WHERE" part of the sql statement)
  '   administrators and supermanagers can only see all accounts (ie need to be a manager to just see ur own stuff)
  '   owners can see all accounts with their owner id
  '   rest just see their account
  '   anyone can select a month or all months

  vWhere = " WHERE "

      vSql = " SELECT * FROM Ecom Ec WITH (nolock) " _
           & "   LEFT OUTER JOIN Cust Cu WITH (nolock) ON Ec.Ecom_CustId = Cu.Cust_Id " _
           & "   LEFT OUTER JOIN Memb Me WITH (nolock) ON Ec.Ecom_MembNo = Me.Memb_No " 

      If Len(vOwnerId) = 4 Then
        vSql = vSql & " LEFT OUTER JOIN V5_Base.dbo.Prog Pr WITH (nolock) ON (Ec.Ecom_Programs = Pr.Prog_Id) AND (Pr.Prog_Owner = '" & vOwnerId & "') "
      End If

      If Len(vOwnerId) = 4 Then
        If svMembLevel = 4 And Not svMembManager Then
          vWhere = vWhere & " ((Cu.Cust_AcctId = '" & svCustAcctId & "') OR (Pr.Prog_Owner = '" & vOwnerId & "')) AND "       
        ElseIf svMembLevel = 4 And svMembManager Then
          vWhere = vWhere & " ((Cu.Cust_Id LIKE '" & Left(svCustId, 4) & "%')  OR (Pr.Prog_Owner = '" & vOwnerId & "')) AND "
        End If
      Else
        If svMembLevel = 4 And Not svMembManager Then
          vWhere = vWhere & "(Cu.Cust_AcctId = '" & svCustAcctId & "') AND "       
        ElseIf svMembLevel = 4 And svMembManager Then
          vWhere = vWhere & " (Cu.Cust_Id LIKE '" & Left(svCustId, 4) & "%') AND "
        End If
      End If


      If Len(vStrDate) > 6  Then vWhere = vWhere & " (Ecom_Issued >= '" & vStrDate & "') AND "
      If Len(vEndDate) > 6  Then vWhere = vWhere & " (Ecom_Issued < DATEADD(d, 1, '" & vEndDate & "')) AND " 

      If vFreebie <> "Y"    Then vWhere = vWhere & " (Ecom_Amount <> 0) AND "
			If Len(vPrograms) > 0 Then vWhere = vWhere & " (CHARINDEX(Ecom_Programs, '" & vPrograms & "') > 0) AND "
      vWhere = vWhere & " (CHARINDEX(Ecom_Source, '" & vSource & "') > 0) "  '...note make this the LAST WHERE else will have training AND

'     vSql = vSql & vWhere & " ORDER BY Ecom_CustId, Ecom_Issued, Ecom_Prices DESC"
      vSql = vSql & vWhere & " ORDER BY Ecom_CustId, Ecom_Issued, Ecom_CardName"

' sDebug

  '...initialize 
  sExcelInit

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof

    sReadEcom     
    sReadCust
    sReadMemb

    '...sometimes accounts have been inactivated or are from a older platform
    If fNoValue(vCust_Title) Then vCust_Title = "(Inactive Account)"      
    If fNoValue(vCust_EcomSplit) Then vCust_EcomSplit = 0

    vNameUser     = ""
    vNameCust     = ""
    vAddressInfo1 = ""
    vAddressInfo2 = ""
    vAddressInfo3 = ""
    vAddressInfo4 = ""

    If svMembLevel = 5 Or svMembManager Or Left(svCustId, 4) = Left(vCust_Id, 4) Then
      vNameUser   = fLeft(vEcom_FirstName & " " & vEcom_LastName, 16)
      vNameCust   = vEcom_CardName

      vOrganization = vMemb_Organization
      vOrganization = fIf(Lcase(vOrganization) = "none", "", vOrganization)
      vOrganization = fIf(Lcase(vOrganization) = "x",    "", vOrganization)
      vOrganization = fIf(Lcase(vOrganization) = "xxx",  "", vOrganization)

      '...sometimes this is added on the ecom-edit an not on the member record
      If Len(vOrganization) = 0 Then vOrganization = vEcom_Organization

      '...get address info
      If vAddress = "Y" Then
        vAddressInfo1 = vEcom_Address
        vAddressInfo2 = vEcom_City & ", " & vEcom_Province & ", " & vEcom_Country
        vAddressInfo3 = vEcom_Phone
        vAddressInfo4 = vEcom_Email
      End If

    End If

    vIssued     = ""
    vExpires    = ""
    vPrice      = ""
    vAmount     = ""
    vSplitVubz  = ""
    vSplitCust  = ""
    vSplitOwnr  = ""
    
    '...get ecom splits
    sGetProg vEcom_Programs '...see if any own splits

    If Len(vProg_Owner) > 0 Then

      '...owners split (if sales within their channel)
      If Left(vProg_Owner, 4) = Left(vCust_Id, 4) Then
        vEcom_SplitOwnr = vEcom_Prices * vProg_EcomSplitOwner1 / 100
      '...owners split if sales in other channels
      Else
        vEcom_SplitOwnr = vEcom_Prices * vProg_EcomSplitOwner2 / 100
      End If
      
			'...capture owner totals in table for level 5
			If svMembLevel = 5 Then
        If Ucase(vEcom_Currency) = "CA" Then
          Redim Preserve aOwnr_CA(vOwnrCnt)
          aOwnr_CA(vOwnrCnt) = aOwnr_CA(vOwnrCnt) + vEcom_SplitOwnr				
        Else
          Redim Preserve aOwnr_US(vOwnrCnt)
          aOwnr_US(vOwnrCnt) = aOwnr_US(vOwnrCnt) + vEcom_SplitOwnr				
        End If
      End If

    Else  
      vEcom_SplitOwnr = 0
    End If            
    
    '...compute the channel split from what's left over (unless sold by owner)
    If Left(vProg_Owner, 4) = Left(vCust_Id, 4) Then
      vEcom_SplitCust = 0
    Else
      vEcom_SplitCust = (vEcom_Prices - vEcom_SplitOwnr) * vCust_EcomSplit / 100
    End If

    '...vubiz get whats left
    vEcom_SplitVubz = vEcom_Prices - vEcom_SplitCust - vEcom_SplitOwnr


    '...display the same issue date beside each program
    vIssued = fFormatDate(vEcom_Issued)
        
    '...if vExpires is invalid then get the duration from the customer program string
    On Error Resume Next '...if no customer record, fall thru
    If Not IsDate(vEcom_Expires) Then 
      vExpires = vExpires & fFormatDate(DateAdd("d", fCustProgDuration (vEcom_CustId, vEcom_Programs), vEcom_Issued))
    Else  
      vExpires = vExpires & fFormatDate(vEcom_Expires)
    End If
    On Error GoTo 0

    vPrice     = FormatNumber(vEcom_Prices, 2)
    vAmount    = FormatNumber(vEcom_Amount, 2)
    vSplitVubz = FormatNumber(vEcom_SplitVubz, 2)
    vSplitCust = FormatNumber(vEcom_SplitCust, 2)
    vSplitOwnr = FormatNumber(vEcom_SplitOwnr, 2)

    '...write out worksheet line
    sExcelRow 

    oRs.MoveNext	        
  Loop
  sCloseDB

  '...close the worksheet 
  sExcelClose


  '...call this one time when ready to setup the worksheet
  Sub sExcelInit

    Set oWs = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell = oWs.Worksheets(1).Cells

    Set oStyleD      	 	  	  	 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
    Set oStyleC      	 		  		 = oWs.CreateStyle

    oStyleD.Number      				 = 14    '...format date m/d/yy
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234
    oStyleC.Number      				 = 2     '...currency

    vRow = 1
    oCell.RowHeight(vRow) = 30

    oCell(vRow, 08).Style = oStyleR
    oCell(vRow, 11).Style = oStyleR
    oCell(vRow, 12).Style = oStyleR
    oCell(vRow, 13).Style = oStyleR
    oCell(vRow, 14).Style = oStyleR
    oCell(vRow, 15).Style = oStyleR

    oCell(vRow, 01) = "Account"						: oCell(vRow, 01).Format.Font.Bold = True   : oCell.ColumnWidth(01) = 15
    oCell(vRow, 02) = "Learner Id"				: oCell(vRow, 02).Format.Font.Bold = True   : oCell.ColumnWidth(02) = 20
    oCell(vRow, 03) = "Card Holder"				: oCell(vRow, 03).Format.Font.Bold = True   : oCell.ColumnWidth(03) = 20
    oCell(vRow, 04) = "Organization"			: oCell(vRow, 04).Format.Font.Bold = True   : oCell.ColumnWidth(04) = 20
    oCell(vRow, 05) = "Source"	 					: oCell(vRow, 05).Format.Font.Bold = True   : oCell.ColumnWidth(05) = 05
    oCell(vRow, 06) = "Type"	  					: oCell(vRow, 06).Format.Font.Bold = True   : oCell.ColumnWidth(06) = 05
    oCell(vRow, 07) = "Program"						: oCell(vRow, 07).Format.Font.Bold = True   : oCell.ColumnWidth(07) = 45
    oCell(vRow, 08) = "Qty"       				: oCell(vRow, 08).Format.Font.Bold = True   : oCell.ColumnWidth(08) = 05
    oCell(vRow, 09) = "Issued"						: oCell(vRow, 09).Format.Font.Bold = True   : oCell.ColumnWidth(09) = 15
    oCell(vRow, 10) = "Expires"						: oCell(vRow, 10).Format.Font.Bold = True   : oCell.ColumnWidth(10) = 15
    oCell(vRow, 11) = "Vubiz"					    : oCell(vRow, 11).Format.Font.Bold = True   : oCell.ColumnWidth(11) = 10
    oCell(vRow, 12) = "$Cust"				      : oCell(vRow, 12).Format.Font.Bold = True   : oCell.ColumnWidth(12) = 10
    oCell(vRow, 13) = "$Owner"	          : oCell(vRow, 13).Format.Font.Bold = True   : oCell.ColumnWidth(13) = 10
    oCell(vRow, 14) = "$Amount"						: oCell(vRow, 14).Format.Font.Bold = True   : oCell.ColumnWidth(14) = 10
    oCell(vRow, 15) = "$Total+Tax"				: oCell(vRow, 15).Format.Font.Bold = True   : oCell.ColumnWidth(15) = 10
    oCell(vRow, 16) = "Curr"       				: oCell(vRow, 16).Format.Font.Bold = True   : oCell.ColumnWidth(16) = 05   
    If vAddress = "Y" Then
    oCell(vRow, 17) = "Address 1"   			: oCell(vRow, 17).Format.Font.Bold = True   : oCell.ColumnWidth(17) = 32
    oCell(vRow, 18) = "Address 2"	   			: oCell(vRow, 18).Format.Font.Bold = True   : oCell.ColumnWidth(18) = 32
    oCell(vRow, 19) = "Phone"	      			: oCell(vRow, 19).Format.Font.Bold = True   : oCell.ColumnWidth(19) = 32
    oCell(vRow, 20) = "Email"  	     			: oCell(vRow, 20).Format.Font.Bold = True   : oCell.ColumnWidth(20) = 32
    End If    
    oCell(vRow, 21) = "Memo"  	     			: oCell(vRow, 21).Format.Font.Bold = True   : oCell.ColumnWidth(21) = 64
    oCell(vRow, 22) = "Order Id"     			: oCell(vRow, 22).Format.Font.Bold = True   : oCell.ColumnWidth(22) = 16
    
  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 02).Style = oStyleI
    oCell(vRow, 10).Style = oStyleI

    oCell(vRow, 11).Style = oStyleC
    oCell(vRow, 12).Style = oStyleC
    oCell(vRow, 13).Style = oStyleC
    oCell(vRow, 14).Style = oStyleC
    oCell(vRow, 15).Style = oStyleC
    oCell(vRow, 22).Style = oStyleL

    oCell(vRow, 01) = vEcom_CustId
    oCell(vRow, 02) = vEcom_Id
    oCell(vRow, 03) = vNameCust
    oCell(vRow, 04) = vOrganization
    oCell(vRow, 05) = vEcom_Source
    oCell(vRow, 06) = Trim(fMedia) 
    oCell(vRow, 07) = vEcom_Programs & " - " & fLeft(fProgTitle(vEcom_Programs), 40)
    oCell(vRow, 08) = vEcom_Quantity
    oCell(vRow, 09) = vIssued
    oCell(vRow, 10) = vExpires
    oCell(vRow, 11) = vSplitVubz
    oCell(vRow, 12) = vSplitCust
    oCell(vRow, 13) = vSplitOwnr
    oCell(vRow, 14) = vPrice
    oCell(vRow, 15) = vAmount
    oCell(vRow, 16) = vEcom_Currency
    If vAddress = "Y" Then
    oCell(vRow, 17) = vAddressInfo1
    oCell(vRow, 18) = vAddressInfo2
    oCell(vRow, 19) = vAddressInfo3
    oCell(vRow, 20) = vAddressInfo4
    End If    
    oCell(vRow, 21) = vEcom_Memo
    oCell(vRow, 22) = vEcom_OrderId

  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
'   Response.BinaryWrite(oWs.Save)
    oWs.Save "Vubiz Ecommerce Report as " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
  
%>

