<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->

<%
  '........................................................................................
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 

  Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
  Set oCell                    = oWs.Worksheets(1).Cells

  Set oStyleD      	 	  	  	 = oWs.CreateStyle
  Set oStyleR      	 		  		 = oWs.CreateStyle
  Set oStyleL      	 		  		 = oWs.CreateStyle
  Set oStyleI      	 		  		 = oWs.CreateStyle

  oStyleD.Number      				 = 14    '...format date m/d/yy
  oStyleR.HorizontalAlignment  = 3     '...right justify
  oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
  oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234
  '........................................................................................


	vRow = 1
	oCell(vRow, 1) = "Customer ID"
	oCell(vRow, 2) = "Title"
	oCell(vRow, 3) = "Catalogue No"

	oCell(vRow, 1).Format.Font.Bold = True   : oCell.ColumnWidth(1) = 14
	oCell(vRow, 2).Format.Font.Bold = True   : oCell.ColumnWidth(2) = 64
	oCell(vRow, 3).Format.Font.Bold = True   : oCell.ColumnWidth(3) = 14

  spCatlByCustId svCustId

  Do While Not oRs.Eof
    vRow = vRow + 1    
    oCell(vRow, 1) = oRs("Catl_CustId").Value
    oCell(vRow, 2) = oRs("Catl_Title").Value
    oCell(vRow, 3) = oRs("Catl_No").Value
    oRs.MoveNext	  
  Loop
  
  Set oCmd = Nothing
  sCloseDb  

  '...output spreadsheet if there are any rows
  Response.ContentType = "application/vnd.ms-excel"
  oWs.Save "Module Access Summary Report dated " & fFormatDate(Now) & ".xls", 1
  Response.End

%>


