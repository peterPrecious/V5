<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<%
  Dim oWs, oCell, vRow, vCol, oStyle_1, oStyle_2, oStyle_3
  Dim vStrDate, vEndDate, vTitle
  
  Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
  Set oCell                    = oWs.worksheets(1).Cells  

  Set oStyle_1                 = oWs.CreateStyle
  Set oStyle_2                 = oWs.CreateStyle
  Set oStyle_3                 = oWs.CreateStyle

  oStyle_1.HorizontalAlignment = 3
  oStyle_2.Number              = 14
  oStyle_3.Number              = 44

  oCell("O3").Style            = oStyle_1
  oCell("P3").Style            = oStyle_1

  vRow = 1
  oCell(vRow, 1) = "Ecommerce Summary Report" : oCell(vRow, 01).Format.Font.Bold = True

  vRow = 3
  oCell(vRow, 01) = "Parent Account"					: oCell(vRow, 01).Format.Font.Bold = True   : oCell.ColumnWidth(01) = 12
  oCell(vRow, 02) = "New Account"						  : oCell(vRow, 02).Format.Font.Bold = True   : oCell.ColumnWidth(02) = 12
  oCell(vRow, 03) = "Order Date"							: oCell(vRow, 03).Format.Font.Bold = True   : oCell.ColumnWidth(03) = 12
  oCell(vRow, 04) = "Expires"							    : oCell(vRow, 04).Format.Font.Bold = True   : oCell.ColumnWidth(04) = 12
  oCell(vRow, 05) = "Agent"     							: oCell(vRow, 05).Format.Font.Bold = True   : oCell.ColumnWidth(05) = 08
  oCell(vRow, 06) = "Source"     							: oCell(vRow, 06).Format.Font.Bold = True   : oCell.ColumnWidth(06) = 18
  oCell(vRow, 07) = "First Name"							: oCell(vRow, 07).Format.Font.Bold = True   : oCell.ColumnWidth(07) = 15
  oCell(vRow, 08) = "Last Name"								: oCell(vRow, 08).Format.Font.Bold = True   : oCell.ColumnWidth(08) = 15
  oCell(vRow, 09) = "Organization"  					: oCell(vRow, 09).Format.Font.Bold = True   : oCell.ColumnWidth(09) = 15
  oCell(vRow, 10) = "Email Address"						: oCell(vRow, 10).Format.Font.Bold = True   : oCell.ColumnWidth(10) = 30
  oCell(vRow, 11) = "Mailing Address"					: oCell(vRow, 11).Format.Font.Bold = True   : oCell.ColumnWidth(11) = 30
  oCell(vRow, 12) = "City"							  		: oCell(vRow, 12).Format.Font.Bold = True   : oCell.ColumnWidth(12) = 20
  oCell(vRow, 13) = "Prov/State"							: oCell(vRow, 13).Format.Font.Bold = True   : oCell.ColumnWidth(13) = 10
  oCell(vRow, 14) = "PC/Zip"									: oCell(vRow, 14).Format.Font.Bold = True   : oCell.ColumnWidth(14) = 10
  oCell(vRow, 15) = "Country"									: oCell(vRow, 15).Format.Font.Bold = True   : oCell.ColumnWidth(15) = 10
  oCell(vRow, 16) = "Programs"								: oCell(vRow, 16).Format.Font.Bold = True   : oCell.ColumnWidth(16) = 10
  oCell(vRow, 17) = "Quantity"								: oCell(vRow, 17).Format.Font.Bold = True   : oCell.ColumnWidth(17) = 10
  oCell(vRow, 18) = "Price"							  		: oCell(vRow, 18).Format.Font.Bold = True   : oCell.ColumnWidth(18) = 15

  vRow = 4
   
  vStrDate = fFormatDate(fDefault(Request("vStrDate"), "Jan 1, 2000"))
  vEndDate = fFormatDate(fDefault(Request("vEndDate"), Now))

  vSql = " SELECT *"_
        & " FROM Ecom" _
        & " WHERE Ecom_Issued BETWEEN '" & vStrDate & "' AND '" & vEndDate & "'" _
        & "   AND (LEFT(Ecom_CustId, 4) = '" & Left(svCustId, 4) & "')" _
        & "   AND Ecom_Amount <> 0" _
        & " ORDER BY Ecom_CustId, Ecom_Issued, Ecom_LastName, Ecom_FirstName "

' sDebug "", vSql

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRS.eof
    sReadEcom
  
    '...write out a line/row
    vRow = vRow + 1

    oCell(vRow, 03).Style = oStyle_2
    oCell(vRow, 04).Style = oStyle_2
'   oCell(vRow, 16).Style = oStyle_3
    oCell(vRow, 18).Style = oStyle_3

    oCell(vRow, 01) = vEcom_CustId
    oCell(vRow, 02) = fIf(Len(Trim(vEcom_NewAcctid)) > 0 , Left(vEcom_CustId, 4), "") & vEcom_NewAcctId
    oCell(vRow, 03) = vEcom_Issued
    oCell(vRow, 04) = vEcom_Expires
    oCell(vRow, 05) = vEcom_Agent
    oCell(vRow, 06) = vEcom_Source
    oCell(vRow, 07) = vEcom_FirstName
    oCell(vRow, 08) = vEcom_LastName
    oCell(vRow, 09) = vEcom_Organization
    oCell(vRow, 10) = vEcom_Email
    oCell(vRow, 11) = vEcom_Address
    oCell(vRow, 12) = vEcom_City
    oCell(vRow, 13) = vEcom_Province
    oCell(vRow, 14) = vEcom_Postal
    oCell(vRow, 15) = vEcom_Country
    oCell(vRow, 16) = vEcom_Programs
    oCell(vRow, 17) = vEcom_Quantity
    oCell(vRow, 18) = vEcom_Prices 
     
    oRs.MoveNext	  
  Loop
  sCloseDb
  
 '...output spreadsheet if there are any rows
  vTitle = "Ecommerce Sales Report dated"
  Response.ContentType = "application/vnd.ms-excel"
  oWs.Save vTitle & " " & fFormatDate(Now) & ".xls", 1
  Response.End
 
%>