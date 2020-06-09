<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, vRow, vCol 
  Dim vTit, vHdr, aHdr, aSql, vCols

  vTit = Request("vTit")
  vHdr = Request("vHdr")
  vSql = Request("vSql")
  
  aHdr = Split(vHdr, "|")
  vCols = Ubound(aHdr)

  '...initialize 
  sExcelInit

  '...generate body
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
    sExcelRow  
    oRs.MoveNext
  Loop
  Set oRs = Nothing
  sCloseDb 
 
  '...close the worksheet 
  sExcelClose


  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs         	 				   = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell       	   				 = oWs.Worksheets(1).Cells
    vRow = 1
    vCol = 0
    For i = 0 To vCols
      vCol = vCol + 1 : oCell(vRow, vCol) = aHdr(i)        : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 20
    Next
    vRow = 2
  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0
    For i = 0 To vCols
      vCol = vCol + 1 : oCell(vRow, vCol) = oRs(i).Value
    Next
  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTit & " - " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
%>



