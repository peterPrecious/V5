<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<% 
  Server.ScriptTimeout = 60 * 10
  Response.Buffer = False

  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleC, oStyleR, oStyleL, oStyleI, vRow, vCol 

  '...Determine the number of programs purchased
  Dim aProgs(), vCnt, aAssigned, vAssigned, custId, custAcctId

  custId = Request("custId")
  custAcctId = Right(custId, 4)

  vSql = " SELECT"_
       & "   Ecom.Ecom_Programs AS ProgId,"_
       & "   SUM(CASE WHEN Ecom_Amount < 0 THEN Ecom_Quantity * -1 ELSE Ecom_Quantity END) AS Purchased"_
       & " FROM Cust INNER JOIN Ecom ON Cust.Cust_AcctId = Ecom.Ecom_NewAcctId "_
       & " WHERE (Cust.Cust_AcctId = '" & custAcctId & "')  AND Ecom_Archived IS NULL "_
       & " GROUP BY Cust.Cust_Id, Ecom.Ecom_Programs "_
       & " ORDER BY Ecom.Ecom_Programs "

  vCnt = 0


  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRS.eof
    vCnt = vCnt + 1
    ReDim Preserve aProgs (2, vCnt)
    aProgs (1, vCnt) = oRs("ProgId")
    aProgs (2, vCnt) = oRs("Purchased")
    oRs.MoveNext	  
  Loop
  Set oRs = Nothing
  sCloseDb
  

  '...Determine the total number of programs assigned
  vSql = " SELECT Memb_Programs"_
       & " FROM Memb "_
       & " WHERE (Memb_AcctId = '" & custAcctId & "') "_
       & "   AND (Len(Memb_Programs) > 0) "_
       & "   AND (Memb_Internal = 0) "
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRS.eof
    aAssigned = Split(oRs("Memb_Programs"))
    For i = 1 to vCnt
      For j = 0 To Ubound(aAssigned)
        If aProgs(1, i) = aAssigned(j) Then
          aProgs(0, i) = aProgs(0, i) + 1
          Exit For
        End If
      Next
    Next
    oRs.MoveNext	  
  Loop
  Set oRs = Nothing
  sCloseDb


  sExcelInitHeader

  vSql = " SELECT Memb_Id, Memb_FirstName, Memb_LastName, Memb_Programs"_
       & " FROM Memb "_
       & " WHERE (Memb_AcctId = '" & custAcctId & "') "_
       & "   AND (Len(Memb_Programs) > 0) "_
       & "   AND (Memb_Internal = 0) "_
       & " ORDER BY Memb_LastName, Memb_FirstName "
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.eof

    sExcelDetailHeader
  
    aAssigned = Split(oRs("Memb_Programs"))
    For i = 1 to vCnt
      For j = 0 To Ubound(aAssigned)
        If aProgs(1, i) = aAssigned(j) Then
          sExcelDetails (i)
          Exit For
        End If
      Next
    Next
    oRs.MoveNext	  

  Loop
  Set oRs = Nothing
  sCloseDb


  '...close the worksheet 
  sExcelClose


  '...call this to setup the header info
  Sub sExcelInitHeader

    Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell                    = oWs.Worksheets(1).Cells

    Set oStyleD      	 	  	  	 = oWs.CreateStyle
    Set oStyleC      	 	  	  	 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
  
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleC.HorizontalAlignment  = 2     '...center justify

    vRow = 1
    vCol = 1

    oCell(vRow + 1, vCol) = "<!--{{-->Programs<!--}}-->"             : oCell(vRow + 1, 01).Format.Font.Bold = True
    oCell(vRow + 2, vCol) = "<!--{{-->Purchased Total<!--}}-->"      : oCell(vRow + 2, 01).Format.Font.Bold = True
    oCell(vRow + 3, vCol) = "<!--{{-->Assigned Total<!--}}-->"       : oCell(vRow + 3, 01).Format.Font.Bold = True
    oCell(vRow + 4, vCol) = "<!--{{-->Balance Remaining<!--}}-->"    : oCell(vRow + 4, 01).Format.Font.Bold = True
    oCell(vRow + 5, vCol) = "<!--{{-->Assigned...<!--}}-->"          : oCell(vRow + 5, 01).Format.Font.Bold = True

    oCell.ColumnWidth(1) = 32

    For i = 1 To vCnt
      oCell(vRow + 1, vCol + i) = aProgs(1, i) & " - " & fLeft(fProgTitleClean(aProgs(1, i)), 29)
      oCell(vRow + 2, vCol + i) = aProgs(2, i)
      oCell(vRow + 3, vCol + i) = aProgs(0, i)
      oCell(vRow + 4, vCol + i) = aProgs(2, i) - aProgs(0, i)

      oCell(vRow + 1, vCol + i).Style = oStyleC : oCell(vRow + 1, vCol + i).Format.Font.Bold = True
      oCell(vRow + 2, vCol + i).Style = oStyleC : oCell(vRow + 2, vCol + i).Format.Font.Bold = True
      oCell(vRow + 3, vCol + i).Style = oStyleC : oCell(vRow + 3, vCol + i).Format.Font.Bold = True
      oCell(vRow + 4, vCol + i).Style = oStyleC : oCell(vRow + 4, vCol + i).Format.Font.Bold = True

      oCell.ColumnWidth(vCol + i) = 40
    Next
   
    vRow = 8

  End Sub



  '...write out details
  Sub sExcelDetailHeader
    vRow = vRow + 1 
    oCell(vRow, 01) = oRs("Memb_Id").Value & " (" & oRs("Memb_FirstName").Value & " " & oRs("Memb_LastName").Value & ")"
    vCol = 1
  End Sub


  Sub sExcelDetails (vHit)
    oCell(vRow, vCol + vHit) = "X"
    oCell(vRow, vCol + vHit).Style = oStyleC
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Dim vTitle
    vTitle = "<!--{{-->Programs Purchased and Assigned<!--}}-->"
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTitle & " - " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

%>
