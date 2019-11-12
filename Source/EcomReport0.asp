<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Cont.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->

<% 

  '...General Parms
  Dim aChannel, vChannelNo, vChannel, vChannelCnt, vCustIdPrev, vCustIdCnt, vCustomer, vNameInfo, vAddressInfo, vAgent, vPrograms, vMembEcom, vGroup, vGroupCnt
  Dim vSource, vStrDate, vEndDate, vStrDateErr, vEndDateErr, vFreebie, vReportType, vPrintType, vNewAcctIds, vIDs, vProgramsPrev, vProgramsCnt, vAddress, vType, vFinished
  Dim vAmount, vSplitOwnrRate, vSplitCustRate, vSplitPointer, vLastName, vEmail, vMemo
  
  '...Detail Totals
  Dim vDetailSplitVubz, vSplitVubz 
  Dim vDetailSplitOwnr, vSplitOwnr
  Dim vDetailSplitCust, vSplitCust
  Dim vDetailSplitTotl, vSplitTotl

  Dim vDetailSplitVubz_US, vDetailSplitVubz_CA
  Dim vDetailSplitOwnr_US, vDetailSplitOwnr_CA
  Dim vDetailSplitCust_US, vDetailSplitCust_CA
  Dim vDetailSplitTotl_US, vDetailSplitTotl_CA
  Dim vDetailPrice_US,     vDetailPrice_CA
  Dim vDetailTotal_US,     vDetailTotal_CA
  Dim vDetailAmount_US,    vDetailAmount_CA

  '...Channel Totals
  Dim vChannelSplitVubz_US, vChannelSplitVubz_CA 
  Dim vChannelSplitOwnr_US, vChannelSplitOwnr_CA
  Dim vChannelSplitCust_US, vChannelSplitCust_CA
  Dim vChannelSplitTotl_US, vChannelSplitTotl_CA  
  Dim vChannelPrice_US,  vChannelPrice_CA, vChannelTotal_US, vChannelTotal_CA, vChannelAmount_US, vChannelAmount_CA
  Dim vChannelOwnrEcom_US, vChannelOwnrEcom_CA, vChannelOwnrManV_US, vChannelOwnrManV_CA, vChannelOwnrManC_US, vChannelOwnrManC_CA
  Dim vChannelVubzEcom_US, vChannelVubzEcom_CA, vChannelVubzManV_US, vChannelVubzManV_CA, vChannelVubzManC_US, vChannelVubzManC_CA
  Dim vChannelCustEcom_US, vChannelCustEcom_CA, vChannelCustManV_US, vChannelCustManV_CA, vChannelCustManC_US, vChannelCustManC_CA

  Dim vChannelEcomAmount_US, vChannelManCAmount_US, vChannelManVAmount_US
  Dim vChannelEcomAmount_CA, vChannelManCAmount_CA, vChannelManVAmount_CA

  Dim vChannelShipping_CA, vChannelShipping_US, vChannelPST, vChannelGST, vChannelHST, vChannelTAX, vChannelCDs
  Dim vChannelProgsSold, vChannelProgsRefs

  Dim vChannelSplitKey, aChannelSplitTot(), aChannelSplit
  Redim aChannelSplitTot(15, 0)


  '...Grand Totals
  Dim vGrandSplitVubz_US, vGrandSplitVubz_CA 
  Dim vGrandSplitOwnr_US, vGrandSplitOwnr_CA
  Dim vGrandSplitCust_US, vGrandSplitCust_CA
  Dim vGrandSplitTotl_US, vGrandSplitTotl_CA  
  Dim vGrandPrice_US,  vGrandPrice_CA, vGrandTotal_US,  vGrandTotal_CA, vGrandAmount_US, vGrandAmount_CA
  Dim vGrandOwnrEcom_US, vGrandOwnrEcom_CA, vGrandOwnrManV_US, vGrandOwnrManV_CA, vGrandOwnrManC_US, vGrandOwnrManC_CA
  Dim vGrandVubzEcom_US, vGrandVubzEcom_CA, vGrandVubzManV_US, vGrandVubzManV_CA, vGrandVubzManC_US, vGrandVubzManC_CA
  Dim vGrandCustEcom_US, vGrandCustEcom_CA, vGrandCustManV_US, vGrandCustManV_CA, vGrandCustManC_US, vGrandCustManC_CA


  Dim vGrandEcomAmount_US, vGrandManCAmount_US, vGrandManVAmount_US
  Dim vGrandEcomAmount_CA, vGrandManCAmount_CA, vGrandManVAmount_CA

  Dim vGrandShipping_CA, vGrandShipping_US, vGrandPST, vGrandGST, vGrandHST, vGrandTAX, vGrandCDs
  Dim vGrandProgs, vGrandRefs

  '...ensure users and/or facilitators don't try to run this report by bypassing the menu page
  If svMembLevel < 4 Then Response.Redirect "Menu.asp"


  '...determine if this user has special ecom rights (true/false) allowing them to access the ecom editor and adjustments
  sGetMemb (svMembNo)
  vMembEcom = vMemb_Ecom 


  vChannel    = Replace(fDefault(Request("vChannel"), Left(svCustId, 4)), ",", "")
  aChannel    = Split(vChannel)
  vReportType = fDefault(Request("vReportType"), "D")
  vPrintType  = fDefault(Request("vPrintType"), "P")
  vCustomer   = Request("vCustomer")
  vAddress    = Request("vAddress")
  vFreebie    = Ucase(fDefault(Request("vFreebie"), "N"))
  vAgent      = Request("vAgent")

  '...added Aug 15, 2012
  vLastName   = fUnQuote(Trim(Request("vLastName")))
  vEmail      = Trim(Request("vEmail"))
  '...added Feb 27, 2018
  vMemo       = Trim(Request("vMemo"))


  '...get source of posting
  vSource = ""
  If fDefault(Request("vSource_E"), "Y") = "Y" Then vSource = "E"
  If fDefault(Request("vSource_V"), "N") = "Y" Then vSource = vSource & "V"
  If fDefault(Request("vSource_C"), "N") = "Y" Then vSource = vSource & "C"
  
  '...get group (if any)
  vGroup = Request("vGroup")
  
  '...defaults to current month
  If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then

    vStrDate  = Request("vStrDate")      : If Len(vStrDate) = 0 Then vStrDate = fFormatSqlDate(MonthName(Month(Now)) & " 1, " & Year(Now))
    vEndDate  = Request("vEndDate")      : If Len(vEndDate) = 0 Then vEndDate = fFormatSqlDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))

  Else
    vStrDate  = fFormatSqlDate(Request("vStrDate")) 
    If Request("vStrDate") = "" Then 
      vStrDate = ""
    ElseIf vStrDate = " " Then
      vStrDate  = Request("vStrDate") 
      vStrDateErr = "Error"
    End If
    vEndDate  = fFormatSqlDate(Request("vEndDate"))
    If Request("vEndDate") = "" Then 
      vEndDate = ""
    ElseIf vEndDate = " " Then
      vEndDate  = Request("vEndDate") 
      vEndDateErr = "Error"
    End If
    If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
      If DateDiff("d", vStrDate, vEndDate) < 0 Then
        vEndDateErr = "Error"
      End If
    End If
  End If

' sStart  '...generate header info _______________________________________________________________________________

  '...Get report criteria
  If Request.Form("vHidden").Count = 0 Or vStrDateErr <> "" Or vEndDateErr <> "" Then   '...If first pass then display prompts 

    sPrompts '...display prompts table  _______________________________________________________________________________

  '...Produce report
  Else 

    vCustIdPrev   = ""
    vCustIdCnt    = 0
    vChannelCnt   = 0
    vProgramsCnt  = 0
    vFinished     = False

    '...create "report" for each channel group selected
    For vChannelNo = 0 To Ubound(aChannel)
      vChannel = aChannel(vChannelNo)

      vSql =        " "
      vSql = vSql & " SELECT CASE SUBSTRING(Ecom.Ecom_CustId, 1, 4) WHEN '" & vChannel & "' THEN 'D' ELSE 'I' END AS [Type], Ecom.*, Memb.*, Cust.*, V5_Base.dbo.Prog.*"
      vSql = vSql & " FROM Ecom WITH (nolock) "
      vSql = vSql & " LEFT OUTER JOIN V5_Base.dbo.Prog ON Ecom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id "
      vSql = vSql & " LEFT OUTER JOIN Cust ON Ecom.Ecom_CustId = Cust.Cust_Id "
      vSql = vSql & " LEFT OUTER JOIN Memb WITH (nolock) ON Ecom.Ecom_MembNo = Memb.Memb_No "

      '...do not select Indirect sales via Vubz
      If vChannel <> "VUBZ" Then
        vSql = vSql & " WHERE (V5_Base.dbo.Prog.Prog_Owner = '" & vChannel & "' OR LEFT(Ecom.Ecom_CustId, 4) = '" & vChannel & "') "  
      Else
        vSql = vSql & " WHERE (LEFT(Ecom.Ecom_CustId, 4) = '" & vChannel & "') "  
      End If

      '...select Agents
      If Len(vAgent) > 0 Then
        vSql = vSql & " AND CHARINDEX(Cust.Cust_Agent, '" & vAgent & "') > 0 "  
      End If

      '...start and end dates 
      If Len(vStrDate) > 6 Then vSql = vSql & " AND Ecom_Issued >= '" & vStrDate & "'"
      If Len(vEndDate) > 6 Then vSql = vSql & " AND Ecom_Issued < DATEADD(d, 1, '" & vEndDate & "')"  

      '...Ignore any records without an amount unless Freebie
      If vFreebie <> "Y" Then
        vSql = vSql & " AND Ecom_Amount <> 0 "
      End If

      '...select only the selected source      
      vSql = vSql & " AND (CHARINDEX(Ecom_Source, '" & vSource & "') > 0) "


      '...select groups
      If Len(vGroup) > 0 And vGroup <> "None" Then
        If Instr("All", vGroup) > 0 Then
          vSql = vSql & " AND (Ecom.Ecom_Media LIKE 'Group%')"
        Else
          vSql = vSql & " AND (CHARINDEX(Ecom.Ecom_NewAcctId, '" & vGroup & "') > 0) "
        End If
      End If


      '...check for lastname - added Aug 15, 2012
      If Len(vLastName) > 0 Then
        vSql = vSql & " AND ((Ecom_Cardname LIKE '%" & vLastName & "%') OR  (Ecom_LastName LIKE '%" & vLastName & "%')) " 
      End If
      '...check for email address - added Aug 15, 2012
      If Len(vEmail) > 0 Then
        vSql = vSql & " AND (Ecom_Email LIKE '%" & vEmail & "%') " 
      End If
      '...check for a memo value - added Feb 27, 2018
      If Len(vMemo) > 0 Then
        vSql = vSql & " AND (Ecom_Memo LIKE '%" & vMemo & "%') " 
      End If
      

      '...sort them
'     vSql = vSql & " ORDER BY Ecom.Ecom_CustId, Ecom.Ecom_Issued, Ecom.Ecom_Programs, [Type] " 
      vSql = vSql & " ORDER BY Ecom.Ecom_CustId, Ecom.Ecom_Issued, Ecom.Ecom_Programs, Ecom.Ecom_No " 

'     sDebug
      sOpenDb
      Set oRs = oDb.Execute(vSql)

      '...print subtitles before starting and initialize fields
      If Not oRs.Eof Then
        vChannelCnt = vChannelCnt + 1 
        sInitializeChannelTotals
        sTitles  
      End If
      
      Do While Not oRs.Eof

        '...determine if program is Direct or Indirect
        vType = oRs("Type")

        sReadEcom     


        '...older ecom records did not contain taxes and shipping tax is not included in taxes
        If vEcom_Taxes = 0 Or vEcom_Shipping <> 0 Then
          If vEcom_Amount > 0 Then                    '...sales
            If vEcom_Amount > vEcom_Prices Then
              vEcom_Taxes = vEcom_Amount - vEcom_Prices - fDefault(vEcom_Shipping, 0)
            End iF
          Else                                        '...credits
            If vEcom_Amount < vEcom_Prices Then
              vEcom_Taxes = vEcom_Amount + vEcom_Prices + fDefault(vEcom_Shipping, 0)
            End iF
          End If
        End If
        

        '...ensure4 we have valid Province/Country values
        vEcom_Province = fDefault(vEcom_Province, "ON")
        vEcom_Country  = fDefault(vEcom_Country, "CA")
  
        If Len(vEcom_Country) <> 2 And vEcom_Taxes > 0 Then vEcom_Country = "CA"
        If Instr("AB BC MB NB NF NT NS NU ON PE QC SK YT", vEcom_Province) = 0 Then
          Select Case Ucase(vEcom_Province)
            Case Ucase("Alberta")                  : vEcom_Province = "AB"
            Case Ucase("British Columbia")         : vEcom_Province = "BC" 
            Case Ucase("Manitoba")                 : vEcom_Province = "MB" 
            Case Ucase("New Brunswick")            : vEcom_Province = "NB" 
            Case Ucase("Newfoundland")             : vEcom_Province = "NF" 
            Case Ucase("Northwest Territories")    : vEcom_Province = "NT" 
            Case Ucase("Nova Scotia")              : vEcom_Province = "NS"
            Case Ucase("Nunavut")                  : vEcom_Province = "NU"
            Case Ucase("Ontario")                  : vEcom_Province = "ON"
            Case Ucase("Prince Edward Island")     : vEcom_Province = "PE" 
            Case Ucase("Quebec")                   : vEcom_Province = "QC" 
            Case Ucase("Saskatchewan")             : vEcom_Province = "SK" 
            Case Ucase("Yukon")                    : vEcom_Province = "YT"
            Case ELSE                              : vEcom_Province = "ON"
          End Select
        End If

        If vProgramsCnt = 0 Then vProgramsPrev = vEcom_Programs
        vProgramsCnt = vProgramsCnt + fIf(vEcom_Amount > 0, 1, -1)

        '...put into this field so we can add a hyperlink for admins 
        vPrograms = vEcom_Programs 


        '...if an adjustment make green 
        If vEcom_Adjustment Then  vPrograms = "<font color='#808000'>" & vPrograms & "</font>"

        '...provide a link to add seats to a G1 Site (and eventually a G2 site)
        If svMembLevel = 5 Or vMembEcom Or svMembManager Then '...eventually allow level 4's to do this
          If vEcom_Media = "Group" And Len(Trim(vEcom_NewAcctId)) = 4 Then
            vPrograms = "<a " & fStatX & "href='EcomAdjustG1.asp?vEcom_NewAcctId=" & vEcom_NewAcctId & "'>" & vPrograms & "</a>"
          End If
        End If

        '...admins and ecom rights can link to the ecom editor
        If svMembLevel = 5 Or vMembEcom Or svMembManager Then
          If fNoValue (vEcom_Archived) Then
            vPrograms = vPrograms & "&nbsp;<a " & fStatX & "href='EcomEdit.asp?vEcom_No=" & vEcom_No & "'>X</a>&nbsp;"
          Else
            vPrograms = vPrograms & "&nbsp;<a onclick='alert(""This account has been archived\nand cannot be edited."");' style='color:red; font-weight:bold' " & fStatX & "href='#'>X</a>&nbsp;"
          End If
        End If


        sReadCust

        '...sometimes accounts have been inactivated
        If fNoValue(vCust_Title) Then vCust_Title = "(Inactive Account)"      
        If fNoValue(vCust_EcomSplit) Then vCust_EcomSplit = 0
        '...if new channel see if any previous totals need to be displayed
        If vEcom_CustId <> vCustIdPrev  Then sSubTotals

        '...get address info
        vAddressInfo = ""
        If vAddress = "Y" And vType = "D" Then
          '...only admins can see address cross account

          '...bug forever, should be svMembLevel
'         If svCustLevel = 5 Or Left(svCustId, 4) = Left(vCust_Id, 4) Then

          If svMembLevel = 5 Or Left(svCustId, 4) = Left(vCust_Id, 4) Then
            vAddressInfo = vAddressInfo & "<br />" & vEcom_Address
            vAddressInfo = vAddressInfo & "<br />" & vEcom_City & ", " & vEcom_Province & ", " & vEcom_Country
            vAddressInfo = vAddressInfo & "<br />" & vEcom_Phone
            vAddressInfo = vAddressInfo & "<br />" & vEcom_Email
          End If
        End If

        sReadMemb


        '...if owner (until Apr 2004 we did not carry cardname, so use first/last
        If Len(fOkValue(vEcom_CardName)) = 0 Then vEcom_CardName = vEcom_FirstName & " " & vEcom_LastName

        '...link to member file if there's a valid member no 
'       If Left(svCustId, 4) = Left(vCust_Id, 4) And vMemb_No > 0 Then
        If Left(svCustId, 4) = Left(vCust_Id, 4) Or svMembLevel = 5 Then          
          If vMemb_No > 0 Then 
            vNameInfo = "<a " & fStatX & "href='User" & fGroup & ".asp?vMembNo=" & vMemb_No & "'>" & fLeft(vEcom_CardName, 16) & "</a>" & vAddressInfo
          Else
            vNameInfo = fLeft(vEcom_CardName, 16) & vAddressInfo
          End If
        Else
          '...else do not display name info
          vNameInfo = ""
        End If


        '...get prog owner and split info (for Prods Vubiz gets whatever the channel does not)
        sReadProgEcom 
    
        '...rates to use on report
        vSplitOwnrRate = 0
        vSplitCustRate = 0
    
        If vEcom_Media = "Prods" Then
          vProg_Owner = "VUBZ"
          vProg_EcomSplitOwner1 = 0
          vProg_EcomSplitOwner2 = 0
        End If

        '...get ecom splits if sold directly via customer
        If vType = "D" Then

          '...get owner split then determine customer piece
          If vProg_Owner = vChannel Then
            vSplitOwnrRate   = vProg_EcomSplitOwner1

            vDetailSplitOwnr = vEcom_Prices * vProg_EcomSplitOwner1 / 100
            vDetailSplitCust = 0
            vDetailSplitVubz = vEcom_Prices - vDetailSplitOwnr             '...vubiz get whats left
          '...remove owner split then determine customer piece (will show up under channel section)
          Else  
            vSplitCustRate   = vCust_EcomSplit

            vDetailSplitOwnr = vEcom_Prices * vProg_EcomSplitOwner2 / 100
            vDetailSplitCust = (vEcom_Prices - vDetailSplitOwnr) * vCust_EcomSplit / 100
            vDetailSplitVubz = vEcom_Prices - vDetailSplitCust - vDetailSplitOwnr            '...vubiz get whats left
            vDetailSplitOwnr = 0
          End If

        '...get owners split if sold indirectly via customer
        ElseIf vType = "I" Then
          vSplitOwnrRate   = vProg_EcomSplitOwner2

          vDetailSplitOwnr = vEcom_Prices * vProg_EcomSplitOwner2 / 100
          vDetailSplitCust = 0
'         vDetailSplitVubz = 0
          vDetailSplitVubz = vEcom_Prices - vDetailSplitOwnr             '...vubiz get whats left
        End If

        vDetailSplitTotl = vDetailSplitOwnr + vDetailSplitCust + vDetailSplitVubz
        
        '...compute split of amount plus tax and shipping
        If vEcom_Prices = 0 Then
          vAmount = 0
        Else
          vAmount = vEcom_Amount * vDetailSplitTotl/vEcom_Prices
        End If
        

        '...compute taxes for bottom line of report
        vChannelTAX = vChannelTAX + vEcom_Taxes

        vPstRate = fPST(vEcom_Issued, vEcom_Country, vEcom_Province)
        vGstRate = fGST(vEcom_Issued, vEcom_Country, vEcom_Province)
        vHstRate = fHST(vEcom_Issued, vEcom_Country, vEcom_Province)

        If vEcom_Media = "CDs" Or vEcom_Media = "Prods" Then
          If vHstRate > 0 Then 
            vChannelHST = vChannelHST + vEcom_Taxes
          ElseIf vGstRate > 0 And vPstRate > 0 Then  
            vChannelPST = vChannelPST + vEcom_Taxes * vPstRate / (vPstRate + vGstRate)
            vChannelGST = vChannelGST + vEcom_Taxes * vGstRate / (vPstRate + vGstRate)
          End If
        Else  
          If vHstRate > 0 Then 
            vChannelHST = vChannelHST + vEcom_Taxes
          ElseIf vGstRate > 0 Then
            vChannelGST = vChannelGST + vEcom_Taxes
          End If
        End If  

'...turn on to confirm we are capturing all taxes
'if vChannelHST + vChannelGST <> vChannelTAX Then stop

        If Ucase(vEcom_Currency) = "CA" Then
          vDetailSplitVubz_CA    = vDetailSplitVubz_CA    + vDetailSplitVubz
          vChannelSplitVubz_CA   = vChannelSplitVubz_CA   + vDetailSplitVubz
          vDetailSplitCust_CA    = vDetailSplitCust_CA    + vDetailSplitCust
          vChannelSplitCust_CA   = vChannelSplitCust_CA   + vDetailSplitCust
          vDetailSplitOwnr_CA    = vDetailSplitOwnr_CA    + vDetailSplitOwnr
          vChannelSplitOwnr_CA   = vChannelSplitOwnr_CA   + vDetailSplitOwnr
          vDetailSplitTotl_CA    = vDetailSplitTotl_CA    + vDetailSplitTotl
          vChannelSplitTotl_CA   = vChannelSplitTotl_CA   + vDetailSplitTotl
          vDetailPrice_CA        = vDetailPrice_CA        + vEcom_Prices
          vChannelPrice_CA       = vChannelPrice_CA       + vEcom_Prices
          vDetailAmount_CA       = vDetailAmount_CA       + vAmount
          vChannelAmount_CA      = vChannelAmount_CA      + vAmount
          vChannelShipping_CA    = vChannelShipping_CA    + vEcom_Shipping

          Select Case vEcom_Source
            Case "E"
              vChannelVubzEcom_CA   = vChannelVubzEcom_CA   + vDetailSplitVubz
              vChannelCustEcom_CA   = vChannelCustEcom_CA   + vDetailSplitCust
              vChannelOwnrEcom_CA   = vChannelOwnrEcom_CA   + vDetailSplitOwnr
              vChannelEcomAmount_CA = vChannelEcomAmount_CA + vAmount
            Case "C"
              vChannelVubzManC_CA   = vChannelVubzManC_CA   + vDetailSplitVubz
              vChannelCustManC_CA   = vChannelCustManC_CA   + vDetailSplitCust
              vChannelOwnrManC_CA   = vChannelOwnrManC_CA   + vDetailSplitOwnr
              vChannelManCAmount_CA = vChannelManCAmount_CA + vAmount
            Case "V"
              vChannelVubzManV_CA   = vChannelVubzManV_CA   + vDetailSplitVubz
              vChannelCustManV_CA   = vChannelCustManV_CA   + vDetailSplitCust
              vChannelOwnrManV_CA   = vChannelOwnrManV_CA   + vDetailSplitOwnr
              vChannelManVAmount_CA = vChannelManVAmount_CA + vAmount
          End Select

        Else
          vDetailSplitVubz_US    = vDetailSplitVubz_US    + vDetailSplitVubz
          vChannelSplitVubz_US   = vChannelSplitVubz_US   + vDetailSplitVubz
          vDetailSplitCust_US    = vDetailSplitCust_US    + vDetailSplitCust
          vChannelSplitCust_US   = vChannelSplitCust_US   + vDetailSplitCust
          vDetailSplitOwnr_US    = vDetailSplitOwnr_US    + vDetailSplitOwnr
          vChannelSplitOwnr_US   = vChannelSplitOwnr_US   + vDetailSplitOwnr
          vDetailSplitTotl_US    = vDetailSplitTotl_US    + vDetailSplitTotl
          vChannelSplitTotl_US   = vChannelSplitTotl_US   + vDetailSplitTotl
          vDetailPrice_US        = vDetailPrice_US        + vEcom_Prices
          vChannelPrice_US       = vChannelPrice_US       + vEcom_Prices
          vDetailAmount_US       = vDetailAmount_US       + vAmount
          vChannelAmount_US      = vChannelAmount_US      + vAmount
          vChannelShipping_US    = vChannelShipping_US    + vEcom_Shipping

          Select Case vEcom_Source
            Case "E"
              vChannelVubzEcom_US   = vChannelVubzEcom_US   + vDetailSplitVubz
              vChannelCustEcom_US   = vChannelCustEcom_US   + vDetailSplitCust
              vChannelOwnrEcom_US   = vChannelOwnrEcom_US   + vDetailSplitOwnr
              vChannelEcomAmount_US = vChannelEcomAmount_US + vAmount
            Case "C"
              vChannelVubzManC_US   = vChannelVubzManC_US   + vDetailSplitVubz
              vChannelCustManC_US   = vChannelCustManC_US   + vDetailSplitCust
              vChannelOwnrManC_US   = vChannelOwnrManC_US   + vDetailSplitOwnr
              vChannelManCAmount_US = vChannelManCAmount_US + vAmount
            Case "V"
              vChannelVubzManV_US   = vChannelVubzManV_US   + vDetailSplitVubz
              vChannelCustManV_US   = vChannelCustManV_US   + vDetailSplitCust
              vChannelOwnrManV_US   = vChannelOwnrManV_US   + vDetailSplitOwnr
              vChannelManVAmount_US = vChannelManVAmount_US + vAmount
          End Select
        End If

        If vEcom_Amount < 0 Then 
          If vType = "D" Then 
            If vEcom_Media = "CDs" Or vEcom_Media = "Prods" Then vChannelCDs = vChannelCDs - 1  
          End If
        Else     
          If vType = "D" Then 
            If vEcom_Media = "CDs" Or vEcom_Media = "Prods" Then vChannelCDs = vChannelCDs + 1  
          End If
        End If  


        '...create total arrays by the various splits     ___________________________________________________________

        '...this is the key to the channel split array
        vChannelSplitKey = vSplitOwnrRate & "|" & vSplitCustRate
        
        '...if first pass then put key in 0,0
        If Len(aChannelSplitTot(0, 0)) = 0 Then aChannelSplitTot(0, 0) = vChannelSplitKey

        '...find which split to enter totals 
        vSplitPointer = -1
        For i = 0 To Ubound(aChannelSplitTot, 2) 
          If aChannelSplitTot(0, i) = vChannelSplitKey Then
            vSplitPointer = i
            Exit For
          End If
        Next

        '...if not in array then add 
        If vSplitPointer = -1 Then
          vSplitPointer = Ubound(aChannelSplitTot, 2) + 1
          Redim Preserve aChannelSplitTot(15, vSplitPointer)
          aChannelSplitTot(0, vSplitPointer) = vChannelSplitKey
        End If         

        If Ucase(vEcom_Currency) = "CA" Then
          aChannelSplitTot(01, vSplitPointer) = aChannelSplitTot(01, vSplitPointer) + vEcom_Prices
          aChannelSplitTot(02, vSplitPointer) = aChannelSplitTot(02, vSplitPointer) + vEcom_Amount
          aChannelSplitTot(03, vSplitPointer) = aChannelSplitTot(03, vSplitPointer) + vDetailSplitVubz
          aChannelSplitTot(04, vSplitPointer) = aChannelSplitTot(04, vSplitPointer) + vDetailSplitOwnr
          aChannelSplitTot(05, vSplitPointer) = aChannelSplitTot(05, vSplitPointer) + vDetailSplitCust
          aChannelSplitTot(06, vSplitPointer) = aChannelSplitTot(06, vSplitPointer) + vDetailSplitTotl
          aChannelSplitTot(07, vSplitPointer) = aChannelSplitTot(07, vSplitPointer) + vEcom_Amount
        Else  
          aChannelSplitTot(08, vSplitPointer) = aChannelSplitTot(08, vSplitPointer) + vEcom_Prices
          aChannelSplitTot(09, vSplitPointer) = aChannelSplitTot(09, vSplitPointer) + vEcom_Amount
          aChannelSplitTot(10, vSplitPointer) = aChannelSplitTot(10, vSplitPointer) + vDetailSplitVubz
          aChannelSplitTot(11, vSplitPointer) = aChannelSplitTot(11, vSplitPointer) + vDetailSplitOwnr
          aChannelSplitTot(12, vSplitPointer) = aChannelSplitTot(12, vSplitPointer) + vDetailSplitCust
          aChannelSplitTot(13, vSplitPointer) = aChannelSplitTot(13, vSplitPointer) + vDetailSplitTotl
          aChannelSplitTot(14, vSplitPointer) = aChannelSplitTot(14, vSplitPointer) + vEcom_Amount
        End If
        
        sDetails                         '____________________________________

        oRs.MoveNext	        

      Loop
      sCloseDb

      '...display totals
      sSubTotals                         '____________________________________

      vChannelProgsSold       = fProgs ("Sold")
      vChannelProgsRefs       = fProgs ("Refs")

      sChannelTotals                     '____________________________________

      sAggregrateTotals
      sInitializeChannelTotals

    ' This ends this channel group
    Next

    vFinished = True

    '...display grand totals if more than one channel (unless type = G)
    If vReportType = "G" Or (vReportType <> "G" And vChannelCnt > 1) Then 
      sTitles
      sGrandTotals
    End If
    
  End If  

  '...wrap up HTML                        _____________________________________
  sFinish
 



  '                                       *******************************************************************************
  '...Functions and SubRoutines 

  Function fCustOptions(vChannel)
    Dim vAll, vCust
    vAll = ""
    fCustOptions = ""
    vSql ="SELECT DISTINCT LEFT(Ecom_CustId, 4) AS Cust FROM Ecom WITH (nolock) ORDER BY Cust"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vCust = oRs("Cust")
      vAll  = vAll & " " & vCust
      fCustOptions  = fCustOptions & "<option " & fIf(vChannel=vCust, "selected ", "") & "value='" & vCust & "'>&nbsp;" & vCust & "&nbsp;</option>" & vbCrLf
      oRs.MoveNext	        
    Loop
    sCloseDb
    fCustOptions  = vbCrLf & "<option " & fIf(vChannel= "All", "selected ", "") & " value='" & vAll & "'>&nbsp;ALL&nbsp;&nbsp;</option>" & fCustOptions & vbCrLf
  End Function


  Function fGroupOptions(vChannel)
    Dim vStyle
    vGroupCnt = 1
    fGroupOptions = ""
'   vSql ="SELECT DISTINCT Ecom_NewAcctId AS [Group], Ecom_Archived FROM Ecom WHERE (LEFT(Ecom_CustId, 4) = '" & vChannel & "') AND (LEN(Ecom_NewAcctId) = 4) AND (ISNUMERIC(Ecom_NewAcctId) = 1) ORDER BY Ecom_NewAcctId, Ecom_Archived"
'   ... removed isnumeric(...) Jun 29, 2016
    vSql ="SELECT DISTINCT Ecom_NewAcctId AS [Group], Ecom_Archived FROM Ecom WHERE (LEFT(Ecom_CustId, 4) = '" & vChannel & "') AND (LEN(Ecom_NewAcctId) = 4) ORDER BY Ecom_NewAcctId, Ecom_Archived"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vGroupCnt = vGroupCnt + 1
      If fNoValue (oRs("Ecom_Archived")) Then
        vStyle = ""
      Else
        vStyle = " style='color:red; font-weight:bold;' "
      End If

      fGroupOptions  = fGroupOptions & "<option " & vStyle & " value='" & oRs("Group") & "'>&nbsp;" & oRs("Group") & "&nbsp;</option>" & vbCrLf 
      oRs.MoveNext	        
    Loop
    sCloseDb
    fGroupOptions  = vbCrLf & "<option value='All'>&nbsp;ALL&nbsp;&nbsp;</option>" & fGroupOptions & vbCrLf
  End Function


  Function fMedia
    Select Case vEcom_Media
      Case "CDs"       : fMedia = "CD "
      Case "Prods"     : fMedia = "PR "
      Case "Group"     : fMedia = "G1 "
      Case "Group2"    : fMedia = "G2 "
      Case "AddOn2"    : fMedia = "G2 "
      Case "Spec_01"   : fMedia = "S1 "
      Case "Corporate" : fMedia = "CP "
      Case Else        : fMedia = "IO "
    End Select
  End Function


  Function fProgs (vType)
    Dim vQuantity, vTotProgs, vLastProg
    vSql =        " SELECT Ecom.Ecom_Id, Ecom_Programs, Ecom.Ecom_Quantity, Ecom.Ecom_Prices, Ecom_Source "
    vSql = vSql & " FROM Ecom WITH (nolock) WHERE (Ecom_Media = 'Online' OR Ecom_Media = 'Group' OR Ecom_Media = 'Group2' OR Ecom_Media = 'Spec_01' OR Ecom_Media = 'Corporate') AND (LEFT(Ecom_CustId, 4) = '" & vChannel & "')"
    If vType = "Sold" Then 
      vSql = vSql & " AND Ecom_Prices >= 0 "
    ElseIf vType = "Refs" Then 
      vSql = vSql & " AND Ecom_Prices < 0 "
    End If    
    If Len(vStrDate) > 6 Then vSql = vSql & " AND Ecom_Issued >= '" & vStrDate & "'"
    If Len(vEndDate) > 6 Then vSql = vSql & " AND Ecom_Issued <= '" & vEndDate & "'"  
    vSql = vSql & " ORDER BY Ecom.Ecom_CustId, Ecom.Ecom_Issued, Ecom.Ecom_Media, Ecom.Ecom_Id, Ecom.Ecom_Programs"
'   sDebug "<br'>fProgs_" & vType , vSql
    fProgs = 0
    vLastProg = ""
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    fProgs = 0
    Do While Not oRs.Eof
      If oRs("Ecom_Prices") < 0 Or oRs("Ecom_Id") = "0" Or oRs("Ecom_Source") <> "E" Or oRs("Ecom_Id") & "|" & oRs("Ecom_Programs") <> vLastProg Then
        fProgs = fProgs + oRs("Ecom_Quantity")
        If oRs("Ecom_Prices") > 0 And oRs("Ecom_Id") <> "0" And oRs("Ecom_Source") = "E" Then
          vLastProg = oRs("Ecom_Id") & "|" & oRs("Ecom_Programs")
        End If
      End If
      oRs.MoveNext
    Loop
    sCloseDb
  End Function


  Function fFormatCurrency(vAmt, vCur)
    vAmt = fDefault(vAmt, 0)
    '... make red for refunds <font color="#FF0000">
    If vAmt < 0 Then 
      fFormatCurrency = "<font color='#FF0000'>" & FormatNumber(vAmt, 2) & " " & vCur & "</font>"
    ElseIf vAmt = 0 Then 
      fFormatCurrency = ""
    Else
      fFormatCurrency = FormatNumber(vAmt, 2) & " " & vCur
    End If
  End Function


  Sub sInitializeChannelTotals
    vChannelPrice_US      = 0
    vChannelPrice_CA      = 0
    vChannelAmount_US     = 0
    vChannelAmount_CA     = 0
    vChannelSplitVubz_US  = 0
    vChannelSplitVubz_CA  = 0
    vChannelSplitOwnr_US  = 0 
    vChannelSplitOwnr_CA  = 0
    vChannelSplitCust_US  = 0 
    vChannelSplitCust_CA  = 0
    vChannelSplitTotl_US  = 0 
    vChannelSplitTotl_CA  = 0
    vChannelOwnrEcom_US   = 0
    vChannelOwnrEcom_CA   = 0  
    vChannelOwnrManV_US   = 0
    vChannelOwnrManV_CA   = 0  
    vChannelOwnrManC_US   = 0
    vChannelOwnrManC_CA   = 0  
    vChannelVubzEcom_US   = 0
    vChannelVubzEcom_CA   = 0  
    vChannelVubzManV_US   = 0
    vChannelVubzManV_CA   = 0  
    vChannelVubzManC_US   = 0
    vChannelVubzManC_CA   = 0  
    vChannelCustEcom_US   = 0
    vChannelCustEcom_CA   = 0  
    vChannelCustManV_US   = 0
    vChannelCustManV_CA   = 0  
    vChannelCustManC_US   = 0
    vChannelCustManC_CA   = 0  

    vChannelEcomAmount_CA = 0
    vChannelManCAmount_CA = 0
    vChannelManVAmount_CA = 0

    vChannelShipping_CA   = 0
    vChannelShipping_US   = 0   

    vChannelPST           = 0
    vChannelGST           = 0
    vChannelHST           = 0
    vChannelTAX           = 0

    vChannelProgsSold     = 0
    vChannelProgsRefs     = 0
    vChannelCDs           = 0

    Redim aChannelSplitTot(15, 0)


  End Sub



  Sub sAggregrateTotals
    vGrandPrice_US	      = 	vGrandPrice_US	      + vChannelPrice_US
    vGrandPrice_CA        = 	vGrandPrice_CA        + vChannelPrice_CA
    vGrandAmount_US       = 	vGrandAmount_US       + vChannelAmount_US
    vGrandAmount_CA       = 	vGrandAmount_CA       + vChannelAmount_CA
    vGrandSplitVubz_US    = 	vGrandSplitVubz_US	  + vChannelSplitVubz_US
    vGrandSplitVubz_CA    = 	vGrandSplitVubz_CA	  + vChannelSplitVubz_CA
    vGrandSplitOwnr_US    = 	vGrandSplitOwnr_US  	+ vChannelSplitOwnr_US
    vGrandSplitOwnr_CA    = 	vGrandSplitOwnr_CA  	+ vChannelSplitOwnr_CA
    vGrandSplitCust_US    = 	vGrandSplitCust_US	  + vChannelSplitCust_US
    vGrandSplitCust_CA    = 	vGrandSplitCust_CA  	+ vChannelSplitCust_CA
    vGrandSplitTotl_US    = 	vGrandSplitTotl_US	  + vChannelSplitTotl_US
    vGrandSplitTotl_CA    = 	vGrandSplitTotl_CA  	+ vChannelSplitTotl_CA
    vGrandOwnrEcom_US     = 	vGrandOwnrEcom_US     + vChannelOwnrEcom_US
    vGrandOwnrEcom_CA     = 	vGrandOwnrEcom_CA     +	vChannelOwnrEcom_CA
    vGrandOwnrManV_US     = 	vGrandOwnrManV_US     +	vChannelOwnrManV_US
    vGrandOwnrManV_CA     = 	vGrandOwnrManV_CA     +	vChannelOwnrManV_CA
    vGrandOwnrManC_US     = 	vGrandOwnrManC_US     +	vChannelOwnrManC_US
    vGrandOwnrManC_CA     = 	vGrandOwnrManC_CA     +	vChannelOwnrManC_CA
    vGrandVubzEcom_US     = 	vGrandVubzEcom_US     +	vChannelVubzEcom_US
    vGrandVubzEcom_CA     = 	vGrandVubzEcom_CA     +	vChannelVubzEcom_CA
    vGrandVubzManV_US     = 	vGrandVubzManV_US     +	vChannelVubzManV_US
    vGrandVubzManV_CA     = 	vGrandVubzManV_CA     +	vChannelVubzManV_CA
    vGrandVubzManC_US     = 	vGrandVubzManC_US     +	vChannelVubzManC_US
    vGrandVubzManC_CA     = 	vGrandVubzManC_CA     +	vChannelVubzManC_CA
    vGrandCustEcom_US     = 	vGrandCustEcom_US     +	vChannelCustEcom_US
    vGrandCustEcom_CA     = 	vGrandCustEcom_CA     +	vChannelCustEcom_CA
    vGrandCustManV_US     = 	vGrandCustManV_US     +	vChannelCustManV_US
    vGrandCustManV_CA     = 	vGrandCustManV_CA     +	vChannelCustManV_CA
    vGrandCustManC_US     = 	vGrandCustManC_US     +	vChannelCustManC_US
    vGrandCustManC_CA     = 	vGrandCustManC_CA     +	vChannelCustManC_CA

    vGrandEcomAmount_CA   =   vGrandEcomAmount_CA   + vChannelEcomAmount_CA
    vGrandManCAmount_CA   =   vGrandManCAmount_CA   + vChannelManCAmount_CA
    vGrandManVAmount_CA   =   vGrandManVAmount_CA   + vChannelManVAmount_CA

    vGrandShipping_CA     =   vGrandShipping_CA     + vChannelShipping_CA
    vGrandShipping_US     =   vGrandShipping_US     + vChannelShipping_US

    vGrandPST             = 	vGrandPST             + vChannelPST
    vGrandGST             = 	vGrandGST             + vChannelGST
    vGrandHST             = 	vGrandHST             + vChannelHST
    vGrandTAX             = 	vGrandTAX             + vChannelTAX

    vGrandProgs           = 	vGrandProgs           + vChannelProgsSold
    vGrandRefs            = 	vGrandRefs            + vChannelProgsRefs
    vGrandCDs             = 	vGrandCDs             + vChannelCDs
  End Sub
%>

<html>

<head>
  <title>EcomReport0</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <style>
    .table .h1 td:nth-child(01) { text-align: left; }
    .table .h1 td:nth-child(02) { text-align: center; }
    .table .h1 td:nth-child(03) { text-align: center; }
    .table .h1 td:nth-child(04) { text-align: center; }
    .table .h1 td:nth-child(05) { text-align: center; }
    .table .h1 td:nth-child(06) { text-align: center; }
    .table .h1 td:nth-child(07) { text-align: center; }
    .table .h1 td:nth-child(08) { text-align: center; }
    .table .h1 td:nth-child(09) { text-align: center; }
    .table .h1 td:nth-child(10) { text-align: center; }

    .table .h2 td:nth-child(01) { text-align: left; }
    .table .h2 td:nth-child(02) { text-align: center; }
    .table .h2 td:nth-child(03) { text-align: center; }
    .table .h2 td:nth-child(04) { text-align: center; }
    .table .h2 td:nth-child(05) { text-align: center; }
    .table .h2 td:nth-child(06) { text-align: center; }
    .table .h2 td:nth-child(07) { text-align: center; }
    .table .h2 td:nth-child(08) { text-align: center; }
    .table .h2 td:nth-child(09) { text-align: center; }

    .table .d2 td:nth-child(01) { text-align: left; }
    .table .d2 td:nth-child(02) { text-align: center; }
    .table .d2 td:nth-child(03) { text-align: center; }
    .table .d2 td:nth-child(04) { text-align: right; }
    .table .d2 td:nth-child(05) { text-align: right; }
    .table .d2 td:nth-child(06) { text-align: right; }
    .table .d2 td:nth-child(07) { text-align: right; }
    .table .d2 td:nth-child(08) { text-align: right; }

    .table tr td:nth-child(01) { text-align: left; }
    .table tr td:nth-child(02) { text-align: left; }
    .table tr td:nth-child(03) { text-align: center; }
    .table tr td:nth-child(04) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(05) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(06) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(07) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(08) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(09) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(10) { text-align: right; white-space: nowrap; }
    .table tr td:nth-child(11) { text-align: right; white-space: nowrap; }

    .heading { color: white; background-color: #4c8be8; white-space: nowrap; font-weight: bold; border: 1px solid white; padding: 5px; margin-bottom: 20px; text-align: right; }

    .table tr:hover { background-color: yellow; }
  </style>
  <%
   '...this option is only for administrators 
   If svMembLevel = 5 Then 
  %>
  <script>
  function Validate(theForm)
  {  
    if (theForm.vChannel.selectedIndex < 0)
    {
      alert("Please select a Channel.");
      theForm.vChannel.focus();
      return (false);
    }
    return (true);
  }
  </script>
  <%
    End If
  %>
</head>

<body style="width: 95%; margin: auto;">

  <% 
  Server.Execute vShellHi 
  %>


  <%
  '______________________________________________________________________________________________________________________________
  '...this is the inital screen where criteria are entered
  Sub sPrompts()
  %>

  <form method="POST" action="EcomReport0.asp" <% If svMembLevel=5 Then %>onsubmit="return Validate(this)" <% End If %>>
    <input type="Hidden" name="vHidden" value="Hidden">
    <table style="width: 80%; margin: auto;">
      <tr>
        <td colspan="3">
          <h1>Advanced Ecommerce Sales Report</h1>
          <h2>Select report criteria then click Generate Report</h2>
        </td>
      </tr>
      <tr>
        <% If svMembLevel = 5 Then %>
        <td style="text-align: center; padding: 15px;">
          <h3>Channel(s)</h3>
          <select size="40" name="vChannel" multiple><%=fCustOptions(vChannel)%></select>
        </td>
        <% End If %>
        <td style="padding: 15px;">
          <p class="c3">Display Options</p>

          <% If svMembLevel = 5 Then %>

          <input type="radio" value="G" name="vReportType">Grand totals of all channels<br />
          <% End If %>
          <input type="radio" value="S" name="vReportType" <%=fcheck("s", vreporttype)%>>Channel totals<br />
          <input type="radio" value="D" name="vReportType" <%=fcheck("d", vreporttype)%>>Channel details
          <br />
          &nbsp;
          <input type="checkbox" name="vCustomer" value="Y" <%=fcheck("y", vcustomer)%>>Include Customer&#39;s name<br />
          &nbsp;
          <input type="checkbox" name="vAddress" value="Y" <%=fcheck("y", vaddress)%>>Include Address (for direct sales)<br />
          &nbsp;
          <input type="checkbox" name="vFreebie" value="Y" <%=fcheck("y", vfreebie)%>>Include &quot;No Charge&quot; programs (not recommended)<p>

          <p class="c3">Include Transactions</p>
          &nbsp;
            E:
          <input type="radio" value="Y" name="vSource_E" <%=fcheck("y", fdefault(request("vsource_e"), "y"))%>>Yes
            <input type="radio" value="N" name="vSource_E" <%=fcheck("n", fdefault(request("vsource_e"), "y"))%>>No (Normal <strong>E</strong>commerce)<br />
          &nbsp;
            V:
          <input type="radio" value="Y" name="vSource_V" <%=fcheck("y", fdefault(request("vsource_v"), "y"))%>>Yes
            <input type="radio" value="N" name="vSource_V" <%=fcheck("n", fdefault(request("vsource_v"), "y"))%>>No (Manual Payment to <strong>V</strong>ubiz)<br />
          &nbsp;
            C:
          <input type="radio" value="Y" name="vSource_C" <%=fcheck("y", fdefault(request("vsource_c"), "y"))%>>Yes
            <input type="radio" value="N" name="vSource_C" <%=fcheck("n", fdefault(request("vsource_c"), "y"))%>>No (Manual Payment to <strong>C</strong>hannel)

              <p class="c3">Start Date</p>
          <input type="text" name="vStrDate" size="10" value="<%=vStrDate%>">
          <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br />
          &nbsp; ie Jan 1, 2012 (Mmm D, YYYY). Leave empty to start at first sale.


              <p class="c3">End Date</p>
          <input type="text" name="vEndDate" size="10" value="<%=vEndDate%>">
          <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br />
          &nbsp; ie Mar 31, 2012 (Mmm D, YYYY). Leave empty to finish with last sale.


              <p class="c3">Optional search filters...</p>

          <div style="margin: 0 20px 20px;">
            <span class="c3">Last Name</span><br />
            <input type="text" name="vLastName" size="15" value="<%=vLastName%>"><br />
            Enter the last name of either a Cardholder or Learner.&nbsp; Note: Learner last name can only be identified in an ecommerce record if it was specified as different from the name of the &nbsp; Cardholder during the purchase transaction. Leave empty to include all names.
          </div>

          <div style="margin: 0 20px 20px;">
            <span class="c3">Email Address</span><br />
            <input type="text" name="vEmail" size="24" value="<%=vEmail%>"><br />
            Enter the Cardholder email address that was specified during the purchase transaction. Leave empty to include all Cardholder email addresses.<br />
          </div>

          <div style="margin: 0 20px 20px;">
            <span class="c3">Ecommerce Memo</span><br />
            <input type="text" name="vMemo" size="60" value="<%=vMemo%>"><br />
            Enter a memo value that was passed in via the Ecommerce Service. It could be the Customer's Order Id.<br />
          </div>


        </td>

        <td style="padding: 15px;">
          <% 
            i = fGroupOptions(vChannel) 
            If vGroupCnt > 1 Then 
          %>
          <h3>Only include <%=vChannel%> Group Sales to...<br />
            <br />
            <select size="<%=fMin(vGroupCnt + 1, 18)%>" name="vGroup" multiple>
              <option style="color: blue" selected value="None">None</option>
              <%=i%>
            </select>
          </h3>
          If you select ALL or any of the above, then only Group sales to the above Customers will be reflected if this report. 
          <%
            End If
          %>
          <br />
          <p class='c6'>Accounts in bold red have been archived.</p>
        </td>
      </tr>
    </table>

    <h6>
      <input type="submit" value="Generate Report" name="bPrint" id="bPrint" class="button" onclick="$(this).hide()"><br />
      <br />
      Note: this report can take several minutes. Please be patient.
    </h6>

  </form>
  <%
  End Sub
  %>





  <%
    '______________________________________________________________________________________________________________________________
    '...these are only displayed at top of the first channel (full table)
    Sub sTitles

      If vReportType <> "G" Or (vReportType = "G" And vFinished) Then 
  %>
  <p></p>
  <p class="page"></p>

  <!--- This will create a new page -->
  <table class="table">
    <tr>
      <td colspan="11" style="text-align: center">
        <h1>Advanced Ecommerce | <%=fIf(vReportType = "D", "Detailed", fif(vReportType = "G", "Overall", "Channel Summary"))%> Sales Report | <%=fFormatSqlDate(Now())%></h1>
        <div class="c3" style="width: 80%; margin: 0 auto 20px; text-align: left;">
          <% If IsDate(vStrDate) And IsDate(vEndDate) Then %> 
           Includes all programs sold directly and indirectly between <%=vStrDate%> and <%=vEndDate%>. 
          <% ElseIf IsDate(vStrDate) Then %> 
             Includes all programs sold directly and indirectly after <%=vStrDate%>. 
          <% ElseIf IsDate(vEndDate) Then %> 
             Includes all programs sold directly and indirectly before <%=vEndDate%>. 
          <% Else%> 
             Includes all programs sold directly and indirectly. 
          <% End If %>
          <% If vReportType = "D" And Not vFinished Then %>
             The Program Code (in green if manually created/adjusted) is followed by an &quot;E&quot; (normal ecommerce), &quot;C&quot; (manual payments to customer) or &quot;V&quot; (manual payments to Vubiz) followed by an &quot;IO&quot; (individual online), &quot;G1&quot; or &quot;G2&quot; (group online), &quot;CP&quot; or &quot;S1&quot;, &quot;S2&quot;... (corporate) or &quot;CD&quot; (CD Rom).&nbsp; Column &quot;%Own&quot; is the percentage revenue to the Content&#39;s Owner.&nbsp; Column &quot;%Chn&quot; is the percentage revenue to the Channel Reseller (calculated after the Owner&#39;s percentage has been deducted).&nbsp; Column &quot;Price&quot; is the dollar amount of the sale.&nbsp; Column &quot;$Total&quot; sums all the payments distributed and &quot;$Total+&quot; contains GST for all Canadian Sales, except for the Canadian Maritimes which includes HST, and shipping.&nbsp; Note that detailed &quot;split&quot; values have been rounded but that Totals are accurate.
          <% End If %>
        </div>
      </td>
    </tr>
    <% 
      End If

      If Not vFinished Then
        If vReportType = "D" Then 
    %>
    <tr class="h1">
      <td class="heading" colspan="2"><%=vChannel%> Details</td>
      <td class="heading">Date Sold<br />Expires</td>
      <td class="heading">%Own</td>
      <td class="heading">%Chn</td>
      <td class="heading">$Price<br />$Price+</td>
      <td class="heading">$Vubiz</td>
      <td class="heading">$Owner</td>
      <td class="heading">$Channel</td>
      <td class="heading">$Total</td>
      <td class="heading">$Total+</td>
    </tr>
    <tr>
      <td colspan="11">&nbsp;</td>
    </tr>
    <% 
        ElseIf vReportType = "S" Then
    %>
    <tr class="h2">
      <td class="heading" colspan="3"><%=vChannel%> Totals</td>
      <td class="heading">%Own</td>
      <td class="heading">%Chn</td>
      <td class="heading">$Price</td>
      <td class="heading">$Vubiz</td>
      <td class="heading">$Owner</td>
      <td class="heading">$Channel</td>
      <td class="heading">$Total</td>
      <td class="heading">$Total+</td>
    </tr>
    <tr>
      <td colspan="11">&nbsp;</td>
    </tr>
    <% 
        End If
      End If

    End Sub 


    '______________________________________________________________________________________________________________________________
    '...print each program details (no table) 
    Sub sDetails()
  
      vCustIdCnt  = vCustIdCnt + 1
      If vCustIdCnt = 1  Then 
       If vReportType = "D" Then 

    %>
    <tr>
      <td colspan="11" class="c3"><%=vChannel%> <% If vType = "D" Then %> Directly via <%=vEcom_CustId%> - <%=vCust_Title%> <% Else %> Indirectly via <%=vEcom_CustId%> - <%=vCust_Title%> <% End If %></td>
    </tr>
    <%
        End If
      End If
  
      '...display if details requested
      If vReportType = "D" Then 
        If vEcom_Prices <> 0 Or vFreebie = "Y" Then
    %>
    <tr class="details">
      <td><%=fIf (vCustomer="Y", "&nbsp;" & vNameInfo, "&ensp;")%></td>
      <td><%=vPrograms%> <%=vEcom_Source & " " & fMedia & " " & Right("00" & vProg_EcomSplitOwner1, 2) & "% " & Right("00" & vProg_EcomSplitOwner2, 2) & "% " & Right("00" & vCust_EcomSplit, 2) & "% " & fLeft(Trim(vProg_Title), 16)%></td>
      <td><%=fFormatSqlDate(vEcom_Issued) & "<br />" & fFormatSqlDate(vEcom_Expires) %></td>
      <td><%=fIf(vChannel = Left(vEcom_CustId, 4), vProg_EcomSplitOwner1, vProg_EcomSplitOwner2)%>% </td>
      <td><%=vCust_EcomSplit%>%</td>
      <td><%=fFormatCurrency(vEcom_Prices, vEcom_Currency) & "<br />" & fFormatCurrency(vEcom_Amount, vEcom_Currency)%></td>
      <td><%=fFormatCurrency(vDetailSplitVubz, vEcom_Currency)%></td>
      <td><%=fFormatCurrency(vDetailSplitOwnr, vEcom_Currency)%></td>
      <td><%=fFormatCurrency(vDetailSplitCust, vEcom_Currency)%></td>
      <td><%=fFormatCurrency(vDetailSplitTotl, vEcom_Currency)%></td>
      <td><%=fFormatCurrency(vAmount, vEcom_Currency)%></td>
    </tr>
    <%
           End If
        End If
        
      End Sub


  '______________________________________________________________________________________________________________________________
  '...these are displayed when channel changes (close table)
  Sub sSubTotals ()
  
    If vReportType <> "G" Then
  
      If vDetailSplitTotl_CA <> 0 Then 
    %>
    <tr>
      <td colspan="3" class="c3">Total via <%=vCustIdPrev %> :</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vDetailPrice_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vDetailSplitVubz_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vDetailSplitOwnr_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vDetailSplitCust_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vDetailSplitTotl_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vDetailAmount_CA, "CA")%></td>
    </tr>
    <%

      End If
  
      If vDetailSplitTotl_US <> 0 Then
    %>
    <tr>
      <td colspan="3" class="c3">Total via <%=vCustIdPrev%> : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vDetailPrice_US, "US")%> </td>
      <td><%=fFormatCurrency(vDetailSplitVubz_US, "US")%> </td>
      <td><%=fFormatCurrency(vDetailSplitOwnr_US, "US")%> </td>
      <td><%=fFormatCurrency(vDetailSplitCust_US, "US")%> </td>
      <td><%=fFormatCurrency(vDetailSplitTotl_US, "US")%> </td>
      <td><%=fFormatCurrency(vDetailAmount_US, "US")%></td>
    </tr>
    <%
      End If
  

  
    End If

    vDetailSplitVubz_US  = 0
    vDetailSplitVubz_CA  = 0
    vDetailSplitCust_US  = 0
    vDetailSplitCust_CA  = 0
    vDetailSplitOwnr_US  = 0
    vDetailSplitOwnr_CA  = 0
    vDetailSplitTotl_US  = 0
    vDetailSplitTotl_CA  = 0
    vDetailPrice_US      = 0
    vDetailPrice_CA      = 0
    vDetailAmount_US     = 0
    vDetailAmount_CA     = 0

    vCustIdCnt           = 0
    vCustIdPrev          = vEcom_CustId
 
  End Sub 
    %> <%
  '______________________________________________________________________________________________________________________________

  Sub sChannelTotals()
    %> <% If vReportType <> "G" Then %> <%  If vChannelSplitTotl_CA <> 0 Or vChannelSplitTotl_US <> 0 Then %>


    <tr>
      <td colspan="11">&nbsp;</td>
    </tr>
    <%    If vReportType = "D" Then '...don't print title for summary as it was printed earlier  %>
    <tr class="h2">
      <td class="heading" colspan="3"><%=vChannel%> Totals</td>
      <td class="heading">%Own</td>
      <td class="heading">%Chn</td>
      <td class="heading">$Price</td>
      <td class="heading">$Vubiz</td>
      <td class="heading">$Owner</td>
      <td class="heading">$Channel</td>
      <td class="heading">$Total</td>
      <td class="heading">$Total+</td>
    </tr>
    <%    End If %>

    <% For i = 0 To Ubound(aChannelSplitTot, 2) %>
    <%   If aChannelSplitTot(07, i) <> 0 Then %>
    <%     aChannelSplit = Split(aChannelSplitTot(0, i), "|") %>
    <tr class="d2">
      <td class="c3" colspan="3">Total by Split - CA : </td>
      <td><%=aChannelSplit(0)%>%</td>
      <td><%=aChannelSplit(1)%>%</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(aChannelSplitTot(03, i), "CA")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(04, i), "CA")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(05, i), "CA")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(06, i), "CA")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(07, i), "CA")%></td>
    </tr>
    <%   End If %>
    <%     If aChannelSplitTot(14, i) <> 0 Then %>
    <%       aChannelSplit = Split(aChannelSplitTot(0, i), "|") %>
    <tr class="d2">
      <td class="c3" colspan="3">Total by Split - US : </td>
      <td><%=aChannelSplit(0)%>%</td>
      <td><%=aChannelSplit(1)%>%</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(aChannelSplitTot(10, i), "US")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(11, i), "US")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(12, i), "US")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(13, i), "US")%></td>
      <td><%=fFormatCurrency(aChannelSplitTot(14, i), "US")%></td>
    </tr>
    <%   End If %> <% Next %> <%    If vChannelVubzEcom_CA + vChannelOwnrEcom_CA + vChannelCustEcom_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Ecommerce Sales - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelVubzEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelOwnrEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelCustEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelVubzEcom_CA + vChannelOwnrEcom_CA + vChannelCustEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelAmount_CA, "CA")%></td>
    </tr>
    <%    End If %> <%    If vChannelVubzEcom_US + vChannelOwnrEcom_US + vChannelCustEcom_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Ecommerce Sales - US : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelVubzEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelOwnrEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelCustEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelVubzEcom_US + vChannelOwnrEcom_US + vChannelCustEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelAmount_US, "US")%></td>
    </tr>
    <%    End If %> <%    If vChannelVubzManC_CA + vChannelOwnrManC_CA + vChannelCustManC_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Manual Sales via Channel - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelVubzManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelOwnrManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelCustManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelVubzManC_CA + vChannelOwnrManC_CA + vChannelCustManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelManCAmount_CA, "CA")%></td>
    </tr>
    <%    End If %> <%    If vChannelVubzManC_US + vChannelOwnrManC_US + vChannelCustManC_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Manual Sales via Channel - US : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelVubzManC_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelOwnrManC_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelCustManC_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelVubzManC_US + vChannelOwnrManC_US + vChannelCustManC_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelManCAmount_US, "US")%></td>
    </tr>
    <%    End If %> <%    If vChannelVubzManV_CA + vChannelOwnrManV_CA + vChannelCustManV_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Manual Sales via Vubiz - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelVubzManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelOwnrManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelCustManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelVubzManV_CA + vChannelOwnrManV_CA + vChannelCustManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelManVAmount_CA, "CA")%></td>
    </tr>
    <%    End If %> <%    If vChannelVubzManV_US + vChannelOwnrManV_US + vChannelCustManV_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Manual Sales via Vubiz - US : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelVubzManV_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelOwnrManV_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelCustManV_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelVubzManV_US + vChannelOwnrManV_US + vChannelCustManV_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelManVAmount_US, "US")%></td>
    </tr>
    <%    End If %> <%   If vChannelSplitTotl_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Total - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelPrice_CA, "CA")%></td>
      <td><%=fFormatCurrency(vChannelSplitVubz_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vChannelSplitOwnr_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vChannelSplitCust_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vChannelSplitTotl_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vChannelAmount_CA, "CA")%></td>
    </tr>
    <%   End If %> <%   If vChannelSplitTotl_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Total - US : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelPrice_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelSplitVubz_US, "US")%> </td>
      <td><%=fFormatCurrency(vChannelSplitOwnr_US, "US")%> </td>
      <td><%=fFormatCurrency(vChannelSplitCust_US, "US")%> </td>
      <td><%=fFormatCurrency(vChannelSplitTotl_US, "US")%> </td>
      <td><%=fFormatCurrency(vChannelAmount_US, "US")%></td>
    </tr>
    <%   End If %> <%   If vChannelShipping_CA <> 0 Or vChannelShipping_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelShipping_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Shipping - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelShipping_CA, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelShipping_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Shipping - US : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelShipping_US, "US")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelPST <> 0 Or vChannelGST <> 0 Or vChannelHST <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelPST <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">PST - CA : </td>
      <td colspan="8">&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelGST <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">GST - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelPST, "CA")%><%=fFormatCurrency(vChannelGST, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelHST <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">HST - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelHST, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vChannelTAX <> 0 Then %>
    <tr>
      <td class="c3" colspan="3">Total Tax - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vChannelTAX, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %>

    <% End If %>
    <%End If %>
    <%
    End Sub
    %>


    <%
    '______________________________________________________________________________________________________________________________
    Sub sGrandTotals()
  
      If vGrandAmount_CA <> 0 Or vGrandAmount_US <> 0 Then
    %>
    <tr>
      <td colspan="11">&nbsp;</td>
    </tr>
    <tr>
      <td class="c3" colspan="6">Grand Totals</td>
      <td class="c3">$Vubiz</td>
      <td class="c3">$Owner</td>
      <td class="c3">$Channel</td>
      <td class="c3">$Total</td>
      <td class="c3">$Total+</td>
    </tr>
    <% If vGrandVubzEcom_CA + vGrandOwnrEcom_CA + vGrandCustEcom_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Ecommerce Sales - CA : </td>
      <td><%=fFormatCurrency(vGrandVubzEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandOwnrEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandCustEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandVubzEcom_CA + vGrandOwnrEcom_CA + vGrandCustEcom_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandEcomAmount_CA, "CA")%></td>
    </tr>
    <% End If %> <% If vGrandVubzEcom_US + vGrandOwnrEcom_US + vGrandCustEcom_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Ecommerce Sales - US : </td>
      <td><%=fFormatCurrency(vChannelVubzEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelOwnrEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vChannelCustEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandVubzEcom_US + vGrandOwnrEcom_US + vGrandCustEcom_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandEcomAmount_US, "US")%></td>
    </tr>
    <% End If %> <% If vGrandVubzManC_CA + vGrandOwnrManC_CA + vGrandCustManC_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Manual Sales via Channel - CA : </td>
      <td><%=fFormatCurrency(vGrandVubzManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandOwnrManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandCustManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandVubzManC_CA + vGrandOwnrManC_CA + vGrandCustManC_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandManCAmount_CA, "CA")%></td>
    </tr>
    <% End If %> <% If vGrandVubzManC_US + vGrandOwnrManC_US + vGrandCustManC_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Manual Sales via Channel - US : </td>
      <td><%=fFormatCurrency(vGrandVubzManC_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandOwnrManC_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandCustManC_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandVubzManC_US + vGrandOwnrManC_US + vGrandCustManC_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandManCAmount_US, "US")%></td>
    </tr>
    <% End If %> <% If vGrandVubzManV_CA + vGrandOwnrManV_CA + vGrandCustManV_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Manual Sales via Vubiz - CA : </td>
      <td><%=fFormatCurrency(vGrandVubzManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandOwnrManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandCustManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandVubzManV_CA + vGrandOwnrManV_CA + vGrandCustManV_CA, "CA")%></td>
      <td><%=fFormatCurrency(vGrandManVAmount_CA, "CA")%></td>
    </tr>
    <% End If %> <% If vGrandVubzManV_US + vGrandOwnrManV_US + vGrandCustManV_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Manual Sales via Vubiz - US : </td>
      <td><%=fFormatCurrency(vGrandVubzManV_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandOwnrManV_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandCustManV_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandVubzManV_US + vGrandOwnrManV_US + vGrandCustManV_US, "US")%></td>
      <td><%=fFormatCurrency(vGrandManVAmount_US, "US")%></td>
    </tr>
    <% End If %>
    <tr>
      <td class="c3" colspan="6">Total - CA : </td>
      <td><%=fFormatCurrency(vGrandSplitVubz_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vGrandSplitOwnr_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vGrandSplitCust_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vGrandSplitTotl_CA, "CA")%> </td>
      <td><%=fFormatCurrency(vGrandAmount_CA, "CA")%></td>
    </tr>
    <% If vGrandSplitTotl_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Total - US : </td>
      <td><%=fFormatCurrency(vGrandSplitVubz_US, "US")%> </td>
      <td><%=fFormatCurrency(vGrandSplitOwnr_US, "US")%> </td>
      <td><%=fFormatCurrency(vGrandSplitCust_US, "US")%> </td>
      <td><%=fFormatCurrency(vGrandSplitTotl_US, "US")%> </td>
      <td><%=fFormatCurrency(vGrandAmount_US, "US")%></td>
    </tr>
    <% End If %>
    <tr>
      <td class="c3" colspan="6">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <%   If vGrandShipping_CA <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Shipping - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vGrandShipping_CA, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vGrandShipping_US <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Shipping - US : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vGrandShipping_US, "US")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vGrandPST <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">PST - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vGrandPST, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vGrandGST <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">GST - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vGrandGST, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vGrandHST <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">HST - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vGrandHST, "CA")%></td>
      <td>&nbsp;</td>
    </tr>
    <%   End If %> <%   If vGrandTAX <> 0 Then %>
    <tr>
      <td class="c3" colspan="6">Total Tax - CA : </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=fFormatCurrency(vGrandTAX, "CA")%></td>
      <td></td>
    </tr>
    <%   End If %> <%   End If %>
    <%

  End Sub


 




  '______________________________________________________________________________________________________________________________ 
  Sub sFinish

    If Request("vHidden").Count > 0 Then 
    %>
  </table>

  <% If vGrandAmount_CA + vGrandAmount_US = 0 Then %>
  <h6 style="margin: 25px">No ecommerce transactions were recorded during the period selected.</h6>
  <% End If %>
  <h6>
    <input onclick="location.href = 'javascript:history.back(1)'" type="button" value="Return" name="bReturn" id="bReturn" class="button"></h6>

  <% If vGrandAmount_CA + vGrandAmount_US <> 0 Then %>
  <h6><a <%=fstatx%> title="Click here to print this report locally." href="javascript:window.print()">
    <img border="0" src="../Images/Icons/Printer.gif" width="18" height="18"></a></h6>
  <% End If %>

  <% 
    End If
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
<%
  End Sub
%>
