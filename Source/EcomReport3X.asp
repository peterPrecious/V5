<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyle, oStyleL, vRow, vCol, vStrDate, vEndDate, vChannels
  Dim vQuantity, vProgsSold, vProgsRefs, vLastProg, vOk

  vStrDate  = Request("vStrDate")
  vEndDate  = Request("vEndDate") 
  vChannels = fDefault(Request("vChannels"), "All")

  sExcelInit '...initialize 

  vSql = "SELECT " _
       & "  Ecom.Ecom_CustId, Ecom.Ecom_Id, Ecom.Ecom_Organization, Ecom.Ecom_Programs, Ecom.Ecom_Quantity, Ecom.Ecom_NewAcctId, Ecom.Ecom_Media, Ecom.Ecom_Source, Ecom.Ecom_Issued, V5_Base.dbo.Prog.Prog_Title1, Ecom.Ecom_Prices, Ecom_MembNo, Ecom.Ecom_FirstName, Ecom.Ecom_LastName, Ecom.Ecom_CardName, Ecom.Ecom_Adjustment, Ecom.Ecom_OrderId " _
       & "FROM "_
       & "  Ecom LEFT OUTER JOIN " _ 
       & "  Cust ON Ecom.Ecom_CustId = Cust.Cust_Id LEFT OUTER JOIN "_
       & "  Memb ON Ecom.Ecom_MembNo = Memb.Memb_No LEFT OUTER JOIN " _ 
       & "  V5_Base.dbo.Prog ON Ecom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id " _
       & "WHERE " _  
       & "  (Ecom_Media <> 'CDs') AND (Ecom_Media <> 'Prods') " _
       &    fIf(vChannels = "All",  " AND ((LEFT(Ecom.Ecom_CustId, 4) = '" & Left(svCustId, 4) & "') OR (Cust.Cust_Agent = '" & Left(svCustId, 4) & "')) ", fIf(vChannels <> "Global", " AND (CHARINDEX(Ecom.Ecom_CustId, '" & vChannels & "') > 0) ", " ")) _
       &    fIf(Len(vStrDate) > 6, " AND (Ecom_Issued >= '" & vStrDate & "') ", " ") _
       &    fIf(Len(vEndDate) > 6, " AND (Ecom_Issued <= '" & vEndDate & "') ", " ") _
       & "ORDER BY "_
       & "  Ecom.Ecom_CustId, Ecom.Ecom_Issued, Ecom.Ecom_Id, Ecom.Ecom_Programs, Ecom.Ecom_Media"

 'sDebug

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
    sExcelRow  '...write out worksheet line
    oRs.MoveNext	        
  Loop
  sCloseDB
  
  sExcelClose   '...close the worksheet 
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell                    = oWs.Worksheets(1).Cells
    Set oStyle                   = oWs.CreateStyle
    Set oStyleL                  = oWs.CreateStyle
    oStyle.Number      				   = 14    '...format date m/d/yy
    oStyle.HorizontalAlignment   = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify

    vRow = 1
    oCell.RowHeight(vRow) = 50

'    oCell(vRow, 09).Style = oStyle

    oCell(vRow, 01) = "Channel Id"    : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 12
    oCell(vRow, 02) = "Learner"       : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 24
    oCell(vRow, 03) = "CardHolder"    : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 24
    oCell(vRow, 04) = "Organization"  : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 24
    oCell(vRow, 05) = "Program "      : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 12
    oCell(vRow, 06) = "Quantity"      : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 08
    oCell(vRow, 07) = "New Id"        : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 12
    oCell(vRow, 08) = "Type"          : oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 12
    oCell(vRow, 09) = "Source"        : oCell(vRow, 09).Format.Font.Bold = True : oCell.ColumnWidth(09) = 08
    oCell(vRow, 10) = "Issued "       : oCell(vRow, 10).Format.Font.Bold = True : oCell.ColumnWidth(10) = 12 
    oCell(vRow, 11) = "Title"         : oCell(vRow, 11).Format.Font.Bold = True : oCell.ColumnWidth(11) = 48
    oCell(vRow, 12) = "Order Id"      : oCell(vRow, 12).Format.Font.Bold = True : oCell.ColumnWidth(12) = 16
  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    Dim vDate
    '...ignore records with same ID and Program (if purchase via Ecom "E" or bypass ecom "C")
    If oRs("Ecom_Prices") < 0 Or oRs("Ecom_Id") = "0" Or oRs("Ecom_Adjustment") = True Or vLastProg = "" Then
     vOk = True
    ElseIf oRs("Ecom_Id") & "|" & oRs("Ecom_Programs") <> vLastProg Then
      vOk = True
    Else 
      vOk = False
    End If

vOk = True

    If vOk Then


      vQuantity = Abs(oRs("Ecom_Quantity"))
      If oRs("Ecom_Prices") >= 0 Then
        vProgsSold = vProgsSold + vQuantity
      Else
        vProgsRefs = vProgsRefs + vQuantity
        vQuantity = vQuantity * -1
      End If

      vLastProg = oRs("Ecom_Id") & "|" & oRs("Ecom_Programs")
      vDate =  oRs("Ecom_Issued")
      vRow = vRow + 1
      oCell(vRow, 01) = oRs("Ecom_CustId").Value
      oCell(vRow, 02) = oRs("Ecom_FirstName").Value & " " & oRs("Ecom_LastName").Value
      oCell(vRow, 03) = oRs("Ecom_CardName").Value
      oCell(vRow, 04) = oRs("Ecom_Organization").Value
      oCell(vRow, 05) = oRs("Ecom_Programs").Value
      oCell(vRow, 06) = vQuantity 
      oCell(vRow, 07) = oRs("Ecom_NewAcctId").Value
      oCell(vRow, 08) = oRs("Ecom_Media").Value
      oCell(vRow, 09) = oRs("Ecom_Source").Value
      oCell(vRow, 10) = vDate : oCell(vRow, 10).Style = oStyle
      oCell(vRow, 11) = fClean(oRs("Prog_Title1"))
      oCell(vRow, 12) = oRs("Ecom_OrderId").Value : oCell(vRow, 12).Style = oStyleL
    End If
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Program Sales Report as " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
  
  
  Function fClean(i)
    If Instr(i, "<") > 0 Then
      fClean = Left(i, Instr(i, "<")-1)
    Else
      fClean = i
    End If
    fClean = fLeft(fClean, 48)
  End Function
  
  
  
%>