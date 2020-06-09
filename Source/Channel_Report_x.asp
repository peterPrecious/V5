<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Chan.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 

  '...initialize 
  sExcelInit

  sGetChan_Rs
  Do While Not oRs.Eof  
    sReadChan  
    sExcelRow
    oRs.MoveNext
  Loop
  Set oRs = Nothing
 
  '...close the worksheet 
  sExcelClose

  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs         	 				   = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell       	   				 = oWs.Worksheets(1).Cells

    Set oStyleD      	 	  	  	 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle

    oStyleD.Number      				 = 14    '...format date m/d/yy
'   oStyleD.Number      				 = "mmm dd, yyyy" '...this does not seem to work 
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234

    vRow = 1
    oCell.RowHeight(vRow) = 50

    For vCol = 3 To 29
      oCell(vRow, vCol).Style = oStyleR
    Next 

    vCol = 0

    vCol = vCol + 1 : oCell(vRow, vCol) = "Channel"			: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Title"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 60

    vCol = vCol + 1 : oCell(vRow, vCol) = "2004e"     	: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2005e"		    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2006e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2007e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2008e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2009e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2010e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2011e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2012e"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12

    vCol = vCol + 1 : oCell(vRow, vCol) = "2004m"     	: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2005m"		    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2006m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2007m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2008m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2009m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2010m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2011m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2012m"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12

    vCol = vCol + 1 : oCell(vRow, vCol) = "2004t"     	: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2005t"		    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2006t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2007t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2008t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2009t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2010t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2011t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "2012t"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12

    vCol = vCol + 1 : oCell(vRow, vCol) = "Contacts"    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 60
    vCol = vCol + 1 : oCell(vRow, vCol) = "Notes"       : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 60

  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 01).Style = oStyleL

    vCol = 0

    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_Id
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_Title

    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2004e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2005e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2006e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2007e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2008e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2009e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2010e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2011e
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2012e

    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2004m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2005m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2006m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2007m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2008m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2009m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2010m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2011m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2012m

    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2004e + vChan_2004m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2005e + vChan_2005m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2006e + vChan_2006m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2007e + vChan_2007m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2008e + vChan_2008m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2009e + vChan_2009m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2010e + vChan_2010m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2011e + vChan_2011m
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_2012e + vChan_2012m

    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_Contacts
    vCol = vCol + 1 : oCell(vRow, vCol) = vChan_Notes
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Annual Ecommerce Sales dated " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
 
%>

