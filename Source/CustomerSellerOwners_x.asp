<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<% 
  '........................................................................................
  Dim oWs, oCell, oStyleD, oStyleC, oStyleR, oStyleL, oStyleI, vRow, vCol 

  Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
  Set oCell                    = oWs.Worksheets(1).Cells

  Set oStyleD      	 	  	  	 = oWs.CreateStyle
  Set oStyleR      	 		  		 = oWs.CreateStyle
  Set oStyleC      	 		  		 = oWs.CreateStyle
  Set oStyleL      	 		  		 = oWs.CreateStyle
  Set oStyleI      	 		  		 = oWs.CreateStyle

  oStyleD.Number      				 = 14    '...format date m/d/yy
  oStyleR.HorizontalAlignment  = 3     '...right justify
  oStyleC.HorizontalAlignment  = 2     '...center align
  oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
  oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234
  '........................................................................................


  Dim vCustId, vCustUrl, vCustTitle, vCustSeller, vCustOwner


  '...initialize 
  sExcelInit

  vSql =        " "
  vSql = vSql & " SELECT DISTINCT Left(Cust_Id, 4) AS [Cust], Cust_EcomSeller, Cust_EcomOwner"
  vSql = vSql & " FROM Cust"
  vSql = vSql & " WHERE Cust_Active = 1 AND (Cust_EcomSeller = 1 OR Cust_EcomOwner = 1) "
  vSql = vSql & " ORDER BY Cust " 
' sDebug "", vSql

  sOpenDb2
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRS.eof

    '...get customer title      
    vSql =        " "
    vSql = vSql & " SELECT TOP 1 Cust_Title, Cust_Url"
    vSql = vSql & " FROM Cust WITH (NOLOCK)"
    vSql = vSql & " WHERE Cust_Active = 1 AND (Cust_EcomSeller = 1 OR Cust_EcomOwner = 1) AND Left(Cust_Id, 4) = '" & oRs("Cust") & "'"
'   sDebug "", vSql
    Set oRs2 = oDb2.Execute(vSql)

    vCustId     = oRs("Cust")
    vCustUrl    = oRs2("Cust_Url")
    vCustTitle  = oRs2("Cust_Title")
    vCustSeller = fIf(oRs("Cust_EcomSeller") = True, "Y", "")
    vCustOwner  = fIf(oRs("Cust_EcomOwner") = True, "Y", "")

    '...write out worksheet line
    sExcelRow 
     
    Set oRs2 = Nothing
    oRs.MoveNext	        
  Loop
  sCloseDB


  '...close the worksheet 
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    vRow = 1
    vCol = 0

    vCol = vCol + 1 : oCell(vRow, vCol) = "Cust Id" 			: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 08
    vCol = vCol + 1 : oCell(vRow, vCol) = "Web Site"      : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "Name"    		  : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 48
    vCol = vCol + 1 : oCell(vRow, vCol) = "Seller?"       :                                             oCell.ColumnWidth(vCol) = 08 : oCell(vRow, vCol).Style = oStyleC : oCell(vRow, vCol).Format.Font.Bold = True 
    vCol = vCol + 1 : oCell(vRow, vCol) = "Owner?"  			:                                             oCell.ColumnWidth(vCol) = 08 : oCell(vRow, vCol).Style = oStyleC : oCell(vRow, vCol).Format.Font.Bold = True 
  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0

    oCell(vRow, 01).Style = oStyleL
    oCell(vRow, 01).Style = oStyleI

    vCol = vCol + 1 : oCell(vRow, vCol) = vCustId 
    vCol = vCol + 1 : oCell(vRow, vCol) = vCustUrl
    vCol = vCol + 1 : oCell(vRow, vCol) = vCustTitle
    vCol = vCol + 1 : oCell(vRow, vCol) = vCustSeller     :                                             oCell.ColumnWidth(vCol) = 08 : oCell(vRow, vCol).Style = oStyleC
    vCol = vCol + 1 : oCell(vRow, vCol) = vCustOwner      :                                             oCell.ColumnWidth(vCol) = 08 : oCell(vRow, vCol).Style = oStyleC
  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Customer Sellers|Owners Report as of " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

  
%>

