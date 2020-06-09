<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleC, oStyleR, oStyleL, oStyleI, vRow, vCol 

  sExcelInit

  sOpenDb
  vSql = "SELECT * FROM Snap WHERE UserNo = " & svMembNo & " ORDER BY ParentId, CustId "
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
    sExcelDetails
    oRs.MoveNext	  
  Loop
  Set oRs = Nothing
  sCloseDb  

  sExcelClose

  Sub sExcelInit

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
    vCol = 0
 
    oCell(vRow, vCol + 1) = "Account"       : oCell(vRow, vCol + 1).Format.Font.Bold = True : oCell.ColumnWidth(vCol + 1) = 12 : oCell(vRow, vCol + 1).Style = oStyleL
    oCell(vRow, vCol + 2) = "Course"        : oCell(vRow, vCol + 2).Format.Font.Bold = True : oCell.ColumnWidth(vCol + 2) = 12 : oCell(vRow, vCol + 2).Style = oStyleL
    oCell(vRow, vCol + 3) = "Title"         : oCell(vRow, vCol + 3).Format.Font.Bold = True : oCell.ColumnWidth(vCol + 3) = 32 : oCell(vRow, vCol + 3).Style = oStyleL
    oCell(vRow, vCol + 4) = "# Completed"   : oCell(vRow, vCol + 4).Format.Font.Bold = True : oCell.ColumnWidth(vCol + 4) = 12 : oCell(vRow, vCol + 4).Style = oStyleR
  
    vRow = 2

  End Sub


  '...write out details
  Sub sExcelDetails
    vRow = vRow + 1 
    vCol = 0

    oCell(vRow, vCol + 1) = oRs("CustId").Value       : oCell(vRow, vCol + 1).Style = oStyleL
    oCell(vRow, vCol + 2) = oRs("ProgId").Value       : oCell(vRow, vCol + 2).Style = oStyleL
    oCell(vRow, vCol + 3) = oRs("Title").Value        : oCell(vRow, vCol + 3).Style = oStyleL
    oCell(vRow, vCol + 4) = oRs("Completed").Value    : oCell(vRow, vCol + 4).Style = oStyleR
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Dim vTitle
    vTitle = "Course Completion Snapshot"
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTitle & " - " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

%>


















</html>