<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim oWs, oCell, oStyle, vRow, vCol 
  Dim vAcct, vSort

  vAcct    = fDefault(Request("vAcct"), "c")
  vSort    = fDefault(Request("vSort"), "u")

  '...Excel Ouput____________________________________________________________________________________________________

  Set oWs         = Server.CreateObject("SoftArtisans.ExcelWriter")
  Set oCell       = oWs.Worksheets(1).Cells
  Set oStyle      = oWs.CreateStyle
  oStyle.WrapText = True
  vRow = 1
  oCell.RowHeight(vRow) = 50
  oCell(vRow, 01) = "Module Id"       : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 10
  oCell(vRow, 02) = "Title"  					: oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 60
  oCell(vRow, 03) = "Time Spent"      : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(05) = 10 

  vSql = " SELECT Mods_Id, Mods_Title, SUM(CAST(RIGHT(Logs_Item, 6) AS int)) AS [TimeSpent]"_
       & " FROM Logs INNER JOIN V5_Base.dbo.Mods ON SUBSTRING(Logs.Logs_Item, 9, 6) = V5_Base.dbo.Mods.Mods_ID"_
       & " WHERE (Logs_Type = 'P')"_
       &     fIf(vAcct = "c",       " AND (Logs_AcctId = '"  & svCustAcctId & "')", "")_
       & " GROUP BY Mods_Id, Mods_Title"
  If vSort = "u" Then
    vSql = vSql & " ORDER BY TimeSpent DESC, Mods_Title, Mods_Id"
  ElseIf vSort = "t" Then
    vSql = vSql & " ORDER BY Mods_Title, Mods_Id"
  Else
    vSql = vSql & " ORDER BY Mods_Id"
  End If

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRS.Eof
    If oRs("TimeSpent") > 0 Then
      vRow = vRow + 1
      oCell(vRow, 01) = oRs("Mods_Id").Value
      oCell(vRow, 02) = oRs("Mods_Title").Value
      oCell(vRow, 03) = oRs("TimeSpent").Value
      oRs.MoveNext	        
    End If 
  Loop
  sCloseDb


 '...output spreadsheet if there are any rows
  Response.ContentType = "application/vnd.ms-excel"
  oWs.Save "Module Usage Report as of " & fFormatDate(Now) & ".xls", 1
  Response.End
  
%>