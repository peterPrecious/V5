<%
  Dim vCatl_No, vCatl_ParentNo, vCatl_CustId, vCatl_Title, vCatl_Promo, vCatl_Order, vCatl_Active, vCatl_Programs, vCatl_TileColor, vCatl_TileIcon, vCatl_JITNo
  Dim vCatl_Eof


  '...Get Active Catl RecordSet for current customer
  Sub sGetCatl_Rs (vCustId)
    vSql = "SELECT * FROM Catl WHERE (Catl_CustId = '" & vCustId & "' AND Catl_Active = 1) ORDER BY Catl_Order, Catl_No"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
  End Sub
  
  
  '...Get All Catl RecordSet for current customer
  Sub sGetCatls_Rs (vCustId)
    vSql = "SELECT * FROM Catl WHERE (Catl_CustId = '" & vCustId & "') ORDER BY Catl_Order, Catl_No"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
  End Sub  
  
  
  Sub sGetCatlByTitle_Rs (vCustId)
    vSql = "SELECT * FROM Catl WHERE (Catl_CustId = '" & vCustId & "' AND Catl_Active = 1) ORDER BY Catl_Title"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
  End Sub


  '...Get Catl RecordSet for All
  Sub sGetCatlAll_Rs (vCustId, vActive)
    vActive = fSqlBoolean (vActive)
    If Len(vCustId) <> 8 Then
      vSql = "SELECT * FROM Catl WHERE (Catl_Active = " & vActive & ") ORDER BY Catl_Order, Catl_No"
    Else  
      vSql = "SELECT * FROM Catl WHERE (Catl_CustId = '" & vCustId & "' AND Catl_Active = " & vActive & ") ORDER BY Catl_Order"
    End If
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
  End Sub

  
  '...create a new catalogue from group ecom sale (works for Group and Group2) 
  '...note that we use sRecreateCatl below AddOnWebService Group2 sales

  Sub sCreateCatl (vCustId, vNewCustId, vCatlNos, vPrograms)
    '...extract matching programs from originating catalogue
    Dim aCatlNos, aPrograms, aProgs, aProg, i, vNewPrograms, bRead, vCatlNo_Prev, vProg_Id
    sOpenDb2    
    '...merge catlnos + programs then sort 
    aCatlNos  = Split(vCatlNos, "|")
    aPrograms = Split(vPrograms, "|")
    '...if Group1 remember that there is a license plus a seat charge, so just check every other one
    For i = 0 To Ubound(aCatlNos) Step fIf(vEcom_Media = "Group", 2, 1)
      vCatl_No = Clng(aCatlNos(i))
      vProg_Id = aPrograms(i)
      If i = 0 Then 
        vCatlNo_Prev = vCatl_No
        bRead = True
        vNewPrograms = ""        
      End If
      If vCatl_No <> vCatlNo_Prev Then
        vSql  = "INSERT INTO Catl (Catl_CustId, Catl_Title, Catl_Programs) VALUES ('" & vNewCustId & "', '" & fUnquote(vCatl_Title) & "', '" & Trim(vNewPrograms) & "')"
        oDb2.Execute(vSql)
        bRead = True
        vNewPrograms = ""
      End If
      If bRead Then 
        vSql = "SELECT Catl_No, Catl_Title, Catl_Programs FROM Catl WHERE Catl_No = " & vCatl_No      
        Set oRs2 = oDb2.Execute(vSql)
        vCatl_No        = oRs2("Catl_No")
        vCatl_Title     = oRs2("Catl_Title")
        vCatl_Programs  = oRs2("Catl_Programs")
        aProgs  = Split(Trim(vCatl_Programs), " ")
        Set oRs2 = Nothing
        bRead = False
      End If          
      For j = 0 to Ubound(aProgs)
        If vProg_Id = Left(aProgs(j), 7) Then
          aProg  = Split(aProgs(j), "~")
          aProg(1)  = 0   '...no ecommerce allowed
          aProg(2)  = 0   '...no ecommerce allowed
          aProg(4)  = 0   '...no expiry date
          vNewPrograms  = vNewPrograms & Join(aProg, "~") & " "
          Exit For
        End If
      Next
      vCatlNo_Prev = vCatl_No
    Next
    vSql  = "INSERT INTO Catl (Catl_ParentNo, Catl_CustId, Catl_Title, Catl_Programs) VALUES (" & vCatl_No & ", '" & vNewCustId & "', '" & fUnquote(vCatl_Title) & "', '" & Trim(vNewPrograms) & "')"
    oDb2.Execute(vSql)
    sCloseDb2    
  End Sub






  '...This creates the Group2 catalogue for Ecom WS and AddOns (revised - again)
  '   this version was replaced Apr 19, 2016 to better handle creating catalogues, orginal (long version) is below but not used
  Sub sRecreateCatl (vCustId, vNewCustId)

    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp6catalogueReCreate"
      .Parameters.Append .CreateParameter("@newCustId", adChar, adParamInput, 8, vNewCustId)
    End With
	  oCmdApp.Execute()
    Set oCmdApp = Nothing
    sCloseDbApp

  End Sub




 '...this is the previous version that didn't handle archived accounts properly - now ignored
  Sub sRecreateCatl_previous (vCustId, vNewCustId)

    Dim i, vOldNos, aOldNos, vOldOrder, vOldProgs, aProgs, bOk
    
    '...create a string of previously purchased programs to determine what is NOT on the catl
    '   as well as the parent no - so we can know which category to add in a new purchase
    '   and the order - so we know next one in case we need to add on an item  
    vSql = "SELECT * FROM Catl WHERE Catl_CustId = '" & vNewCustId & "' ORDER BY Catl_Order"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    Do While Not oRs2.Eof
      sReadCatl
      vOldNos = vOldNos & vCatl_ParentNo & " " 
      vOldOrder = vCatl_Order
      aProgs = Split(Trim(vCatl_Programs))
      For i = 0 to Ubound(aProgs)
        vOldProgs = vOldProgs & Left(aProgs(i), 7) & " "
      Next
      oRs2.MoveNext
    Loop   
    sCloseDb2
    aOldNos = Split(Trim(vOldNos))
    

    '...get all the programs from the Ecom file (Programs do not expire for G2 Accounts)
    vSql = "SELECT Ecom_Programs, Ecom_CatlNo FROM Ecom WHERE Ecom_Archived IS NULL AND (Ecom_NewAcctId = '" & Right(vNewCustId, 4) & "')"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    Do While Not oRs2.Eof
    
      vEcom_Programs  = oRs2("Ecom_Programs")
      vEcom_CatlNo    = oRs2("Ecom_CatlNo")
      
      '...if this program is not on the Catl then add it else leave Catl as is
      If Instr(Trim(vOldProgs), vEcom_Programs) = 0 Then

        '...add it to the string of programs on the new catl
        vOldProgs = vOldProgs & vEcom_Programs & " "

        '...we need to either add this program to an existing catl item on the new catl or create a new catl item
        bOk = False '...if false then keep trying to rebuild the catl
        
               
        '...see if the sold program is on the new catl based on the parent catl no, if so update the programs on that item
        If Not bOk Then
          For i = 0 To Ubound(aOldNos)
            '...if we have an existing item with the parent catl no then add this program to it
            If cLng(aOldNos(i)) = vEcom_CatlNo Then
              vSql = "SELECT * FROM Catl WHERE Catl_CustId = '" & vNewCustId & "' AND Catl_ParentNo = " & vEcom_CatlNo
              sOpenDb3
              Set oRs3 = oDb3.Execute(vSql)          
              vSql = "UPDATE Catl SET Catl_Programs = Catl_Programs + ' " & vEcom_Programs & "~0~0~1~0' WHERE Catl_No = " & oRs3("Catl_No")
              oDb3.Execute(vSql)          
              sCloseDb3
              bOk = True
              Exit For
            End If
          Next
        End If

        '   else see if we have the catl no on the parent catl and if so then insert this record with the purchased programs
        '...this checks if the newly purchased program is on the current catalogue
        If Not bOk Then
          vSql = "SELECT * FROM Catl WHERE Catl_No = " & vEcom_CatlNo 
          sOpenDb3
          Set oRs3 = oDb3.Execute(vSql)
          If Not oRs3.Eof Then 
            vOldOrder = vOldOrder+ 1
            vSql  = "INSERT INTO Catl (Catl_ParentNo, Catl_CustId, Catl_Title, Catl_Order, Catl_Programs) VALUES (" & vEcom_CatlNo & ", '" & vNewCustId & "', '" & fUnQuote(oRs3("Catl_Title")) & "', " & vOldOrder & ", '" & vEcom_Programs & "~0~0~1~0')"
            oDb3.Execute(vSql)
            bOk = True
          End If
          sCloseDb3
        End If

        '   if this catl no does not exist on the new or part catl then create a new orphan else see if we have the catl no on the parent catl and if so then insert this record with the purchased programs
        If Not bOk Then
          vOldOrder = vOldOrder+ 1
          vSql  = "INSERT INTO Catl (Catl_ParentNo, Catl_CustId, Catl_Title, Catl_Order, Catl_Programs) VALUES (" & vEcom_CatlNo & ", '" & vNewCustId & "', '', " & vOldOrder & ", '" & vEcom_Programs & "~0~0~1~0')"
          sOpenDb3
          oDb3.Execute(vSql)
          sCloseDb3
          bOk = True
        End If

      End If

      oRs2.MoveNext
    Loop

    sCloseDb2
    Set oRs2 = Nothing

  End Sub









  '...clone the catalogue for catl editor
  Sub sCloneCatl (vCustId)
    Dim vCatlOrder
    '...get next order no
    vSql = "SELECT TOP 1 Catl_Order FROM Catl WHERE Catl_CustId = '" & svCustId & "' ORDER BY Catl_Order DESC"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then 
      vCatlOrder = 0
    Else
      vCatlOrder = oRs2("Catl_Order")
    End If         
    Set oRs = Nothing
    '...get the other customers catalogue of active groups
    vSql = "SELECT * FROM Catl WHERE Catl_CustId = '" & vCustId & "' AND Catl_Active = 1 ORDER BY Catl_Order"
    Set oRs2 = oDb2.Execute(vSql)
    '...add into current catalogue
    Do While Not oRs2.Eof
      sReadCatl
      vCatlOrder = vCatlOrder + 1
      vSql  = "INSERT INTO Catl (Catl_CustId, Catl_Title, Catl_Order, Catl_Programs) VALUES ('" & svCustId & "', '" & fUnQuote(vCatl_Title) & "', " & vCatlOrder & ", '" & vCatl_Programs & "')"
'     sDebug
      oDb2.Execute(vSql)
      oRs2.MoveNext
    Loop
    Set oRs2 = Nothing
    sCloseDb2    
  End Sub


  '...Reorder the groups after a delete
  Sub sOrderCatl  
    Dim vCatlOrder
    vCatlOrder = 0
    
    '...get the current catalogue
    sOpenDb2
    vSql = "SELECT * FROM Catl WHERE Catl_CustId = '" & svCustId & "' ORDER BY Catl_Order"
    Set oRs2 = oDb2.Execute(vSql)
    '...rewrite the order
    Do While Not oRs2.Eof
      sReadCatl
      vCatlOrder = vCatlOrder + 1
      vSql  = "UPDATE Catl Set Catl_Order = " & vCatlOrder & " WHERE Catl_No = " & vCatl_No
'     sDebug
      oDb2.Execute(vSql)
      oRs2.MoveNext
    Loop
    Set oRs2 = Nothing
    sCloseDb2    
  End Sub


  '...Add a Catl Record
  Sub sAddCatl
    vSql = "SELECT TOP 1 Catl_Order FROM Catl WHERE Catl_CustId = '" & svCustId & "' ORDER BY Catl_Order DESC"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then 
      vCatl_Order = 1
    Else
      vCatl_Order = oRs2("Catl_Order") + 1
    End If         
    Set oRs2 = Nothing      
    sCloseDb2
    vCatl_Active   = 1
    vCatl_Title    = "[new catalogue item]"
    vCatl_Programs = ""
    sInsertCatl
  End Sub
  
  
  '...Get Catl Record
  Sub sGetCatl (vCatlNo)
    vCatl_Eof = False
'   If Not IsNumeric(vCatlNo) Then Exit Sub
    If Not IsNumeric(fOkValue(vCatlNo)) Then Exit Sub '...added May 11, 2016 since event logs show bombing on SQL ... suspect no value for vCatlNo
    vSql = "SELECT * FROM Catl WHERE Catl_No = " & vCatlNo
'   sDebug
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      sReadCatl
      vCatl_Eof = True
    End If
    Set oRs2 = Nothing
    sCloseDb2    
  End Sub

  Sub sReadCatl
    vCatl_No             = oRs2("Catl_No")
    vCatl_ParentNo       = oRs2("Catl_ParentNo")
    vCatl_CustId         = oRs2("Catl_CustId")
    vCatl_Order          = oRs2("Catl_Order")
    vCatl_Active         = oRs2("Catl_Active")
    vCatl_Title          = oRs2("Catl_Title")
    vCatl_Promo          = oRs2("Catl_Promo")
    vCatl_Programs       = oRs2("Catl_Programs")
    vCatl_TileColor      = oRs2("Catl_TileColor")
    vCatl_TileIcon       = oRs2("Catl_TileIcon")
    vCatl_JITNo          = oRs2("Catl_JITNo")
  End Sub

  Sub sExtractCatl
    vCatl_No             = Request.Form("vCatl_No")
    vCatl_Order          = fDefault(Request.Form("vCatl_Order"), 0)
    vCatl_Active         = Request.Form("vCatl_Active")
    vCatl_Title          = fUnQuote(Request.Form("vCatl_Title"))
    vCatl_Promo          = fUnQuote(Request.Form("vCatl_Promo"))
    vCatl_Programs       = fNoQuote(Request.Form("vCatl_Programs"))
    vCatl_TileColor      = Lcase(Request.Form("vCatl_TileColor"))
    vCatl_TileIcon       = Lcase(Request.Form("vCatl_TileIcon"))
    vCatl_JITNo          = fDefault(Request.Form("vCatl_JITNo"), 0)
  End Sub
  
  Sub sInsertCatl
    vSql = "INSERT INTO Catl "
    vSQL = vSQL & "(Catl_CustId, Catl_Order, Catl_Active, Catl_Title, Catl_Promo, Catl_Programs, Catl_TileColor, Catl_TileIcon, Catl_JITNo)"
    vSQL = vSQL & " VALUES ('" & svCustId & "',  " & vCatl_Order & ", " & vCatl_Active & ", '" & fUnquote(vCatl_Title) & "', '" & fUnquote(vCatl_Promo) & "', '" & vCatl_Programs & "', '" & vCatl_TileColor & "', '" & vCatl_TileIcon & "', " & fDefault(vCatl_JITNo, 0) & ")"
'   sDebug
    sOpenDb2
    oDb2.Execute(vSQL)
    sCloseDb2
  End Sub

  Sub sUpdateCatl
    vSQL = "UPDATE Catl SET"
    vSQL = vSQL & " Catl_Order     =  " & vCatl_Order              & " , " 
    vSQL = vSQL & " Catl_Active    =  " & vCatl_Active             & " , " 
    vSQL = vSQL & " Catl_Title     = '" & vCatl_Title              & "', " 
    vSQL = vSQL & " Catl_Promo     = '" & vCatl_Promo              & "', " 
    vSQL = vSQL & " Catl_Programs  = '" & vCatl_Programs           & "', " 
    vSQL = vSQL & " Catl_TileColor = '" & vCatl_TileColor          & "', " 
    vSQL = vSQL & " Catl_TileIcon  = '" & vCatl_TileIcon           & "', " 
    vSQL = vSQL & " Catl_JITNo     =  " & vCatl_JITNo              & "   " 

    vSQL = vSQL & " WHERE Catl_No = " & vCatl_No
'   sDebug
    sOpenDb2
    oDb2.Execute(vSQL)
    sCloseDb2
  End Sub

  
  Sub sDeleteCatl
    vSQL = "DELETE FROM Catl WHERE Catl_No = " & vCatl_No
    sOpenDb2
    oDb2.Execute(vSQL)
    sCloseDb2
  End Sub


  '...return Catl subsets with program strings for channel setup
  Function fCatlDropdown (vIds, vSubSet)
    Dim vCurrentId, vSelected, aProg, vFree, vEcom
    '...save the current Catlp id
    fCatlDropDown = vbCrLf
    vSql = "SELECT * FROM Catl WHERE (Catl_Active =1) AND Len(Catl_Programs) > 0 AND Right(Catl_Order, 2) = '" & fIf(Ucase(svLang) = "ROOT", "EN", svLang)  & "' AND Len(Catl_Title) > 0" 
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    Do While Not oRs2.Eof
      sReadCatl
      '...If Catl includes items for sell, add $ after the Catl name
      aProg = Split(vCatl_Programs)
      vEcom = "" : vFree = ""
      For i = 0 To Ubound(aProg)
        If Mid(aProg(i), 8, 3) = "~0~" Then
          vFree = "F"
        Else
          vEcom = "$"
        End If
      Next 
      If Instr(vIds, vCatl_Order) > 0 Then
        vSelected = " selected" 
      Else
        vSelected = ""
      End If
      i = "          <option value=" & Chr(34) & vCatl_Order & Chr(34) & vSelected & ">" & fLeft(vCatl_Title, 48) & " [" & vFree & vEcom & "]</option>" & vbCrLf
      fCatlDropdown = fCatlDropdown & i
      oRs2.MoveNext	        
    Loop
    Set oRs2 = Nothing      
    sCloseDb2
    '...save the current Catlaign id
    vCatl_Order = vIds
  End Function 


  Sub sShiftOrder (vCatlNo, vAction)
  
    Dim aRs(), vCurrNo, vSeekNo, vCatlOrder
    
    '...if shifting to top or bottom, shift, sort then exit
    If vAction = "TP" or vAction = "BT" Then
      If vAction = "TP" Then vCatlOrder = 0 Else vCatlOrder = 9999
      sOpenDb2
      vSql = "UPDATE Catl SET Catl_Order = " & vCatlOrder & " WHERE Catl_No = " & vCatlNo
      oDb2.Execute(vSql)
      sCloseDb2
      '...now reorder
      sOrderCatl
      Exit Sub
    End If
    
    
    i = 0
    vCurrNo = 0
    vSeekNo = vCatlNo
    
    '...put the recordset into an array and flag the current vCatlNo so we can find the value before or after
    
    vSql = "SELECT * FROM Catl WHERE Catl_CustId = '" & svCustId & "' ORDER BY Catl_Order, Catl_No"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)

    Do While Not oRs2.Eof 
      sReadCatl
      i = i + 1
      ReDim Preserve aRs(2, i)
      aRs(1, i) = vCatl_No      
      aRs(2, i) = vCatl_Order
      If vCatl_No = vSeekNo Then
        vCurrNo = i
      End If
      oRs2.MoveNext
    Loop
    Set oRs2 = Nothing
    sCloseDb2
    
    '...if "up" and top record or "down" and bottom record then return without shifting
    If (vAction = "UP" and vCurrNo = 1) Or (vAction = "DN" and vCurrNo = ubound(aRs,2)) Then
      Exit Sub
    End If

    '...now that we have the two nos, swap their order value
    sOpenDb2
    '...store the current order
    vSql = "UPDATE Catl Set Catl_Order = " & -999999         & " WHERE Catl_No = " & aRs(1, vCurrNo)
    Set oRs2 = oDb2.Execute(vSql)
    If vAction = "DN" Then
      vSql = "UPDATE Catl Set Catl_Order = " & aRs(2, vCurrNo) & " WHERE Catl_No = " & aRs(1, vCurrNo + 1)
      Set oRs2 = oDb2.Execute(vSql)
      '...set the next order to the current order
      vSql = "UPDATE Catl Set Catl_Order = " & aRs(2, vCurrNo + 1) & " WHERE Catl_No = " & aRs(1, vCurrNo)
      Set oRs2 = oDb2.Execute(vSql)
    ElseIf vAction = "UP" Then
      vSql = "UPDATE Catl Set Catl_Order = " & aRs(2, vCurrNo) & " WHERE Catl_No = " & aRs(1, vCurrNo - 1)
      Set oRs2 = oDb2.Execute(vSql)
      '...set the next order to the current order
      vSql = "UPDATE Catl Set Catl_Order = " & aRs(2, vCurrNo - 1) & " WHERE Catl_No = " & aRs(1, vCurrNo)
      Set oRs2 = oDb2.Execute(vSql)
    End If
    sCloseDb2
  End Sub


  '...Get Catl No for specific program
  Function fCatlNo (vCustId, vProgId)
    fCatlNo = 0
    vSql = "SELECT TOP 1 Catl_No FROM Catl WHERE (Catl_CustId = '" & vCustId & "') AND (CHARINDEX('" & vProgId & "', Catl_Programs) > 0) ORDER BY Catl_CustId"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fCatlNo = oRs2("Catl_No")
    Set oRs2 = Nothing
    sCloseDb2
  End Function 


  '...Get Catl expiry for specific program
  Function fCatlExpires (vCustId, vProgId)
    Dim aProgs, aProg
    fCatlExpires = 90 '...defaults to 90 days
    vSql = "SELECT TOP 1 Catl_Programs FROM Catl WHERE (Catl_CustId = '" & vCustId & "') AND (CHARINDEX('" & vProgId & "', Catl_Programs) > 0) ORDER BY Catl_CustId"
'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then       
      aProgs = Split(oRs2("Catl_Programs"))   '...find program group
      For i = 0 To Ubound(aProgs)
        If Left(aProgs(i), 7) = vProgId Then
          aProg = Split(aProgs(i), "~")       '...find Expiry date
          fCatlExpires = aProg(4)
        End If
      Next    
    End If  
    Set oRs2 = Nothing
    sCloseDb2
  End Function 


  Sub sClearCatl (vCustId)
    vSql  = "DELETE Catl WHERE Catl_CustId = '" & vCustId & "'"
    sOpenDb2 
    oDb2.Execute(vSql)
    sCloseDb2    
  End Sub

  '...Get the Catalogue Nos for third parties
  Sub spCatlByCustId (vCustId)
    sOpenCmd
    With oCmd
      .CommandText = "spCatlByCustId"
      .Parameters.Append .CreateParameter("@CustId",    		adChar,  adParamInput,    8, vCustId)
    End With
	  Set oRs = oCmd.Execute()
  End Sub
  
  '...Get Siblings
  Function spCatlSiblings (vCust)
    spCatlSiblings = ""
    sOpenCmd
    With oCmd
      .CommandText = "spCatlSiblings"
      .Parameters.Append .CreateParameter("@Cust", adChar, adParamInput, 4, vCust)
    End With
	  Set oRs = oCmd.Execute()
    Do While Not oRs.Eof       
      spCatlSiblings = spCatlSiblings & " " & oRs("Cust_Id")   
      oRs.MoveNext
    Loop
    spCatlSiblings = Trim(spCatlSiblings)
  End Function

  '...Get Master (while technically there could be more than one master we only we only retrieve the first since more than one is an error)
  Function spCatlMaster (vCust)
    spCatlMaster = ""
    sOpenCmd
    With oCmd
      .CommandText = "spCatlMaster"
      .Parameters.Append .CreateParameter("@Cust", adChar, adParamInput, 4, vCust)
    End With
	  Set oRs = oCmd.Execute()
    spCatlMaster = oRs("Cust_Id")   
  End Function

  '...Copy Master Catalogue to Siblings Catalogues
  Function spCatlCopy (vCustId)
    Dim ReturnValue
    sOpenCmd
    With oCmd
      .CommandText = "spCatlCopy"
      .Parameters.Append .CreateParameter("RetValue",   adInteger,  adParamReturnValue)
      .Parameters.Append .CreateParameter("@CustId",    adChar,     adParamInput,       8, vCustId)
    End With
	  oCmd.Execute()
    spCatlCopy = oCmd.Parameters("RetValue").Value '...this is the number of Sibling Accounts that were copies
  End Function
 
  

  '...Get the Catalogue Nos for third parties
  Sub spCatlByCustId (vCustId)
    sOpenCmd
    With oCmd
      .CommandText = "spCatlByCustId"
      .Parameters.Append .CreateParameter("@CustId",    		adChar,  adParamInput,    8, vCustId)
    End With
	  Set oRs = oCmd.Execute()
  End Sub
    
%>