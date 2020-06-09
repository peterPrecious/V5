<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, vRow

  sExcelInit
  vSql = "select distinct lower(memb_email) as email from memb inner join cust on memb_acctid = cust_acctid where cust_active = 1 and memb_active = 1 and memb_vunews = 1 and Len(memb_email) > 0"
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
    vMemb_Email = oRs("email")
    sExcelRow  
    oRs.MoveNext
  Loop
  Set oRs = Nothing
  sCloseDb 
  sExcelClose

  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs    = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell  = oWs.Worksheets(1).Cells
    oCell(1,1) = "Email Address"        
    oCell(1,1).Format.Font.Bold = True
    oCell.ColumnWidth(1) = 40
    vRow = 1
  End Sub

  '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    oCell(vRow, 1) = vMemb_Email
  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "VuNews Email Address" & " " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
%>

