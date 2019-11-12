<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, vRow, vCol 

' vSql = "SELECT TOP (5000) "_
  vSql = "SELECT "_
       & "  Memb.Memb_Id          AS [Learner ID], "_ 
       & "  Crit.Crit_Id          AS [Group], "_ 
       & "  Memb.Memb_FirstName   AS [First Name], "_ 
       & "  Memb.Memb_LastName    AS [Last Name], "_ 
       & "  Memb.Memb_Email       AS [Email Address], "_ 
       & "  Memb.Memb_Pwd         AS Password, "_ 
       & "  Memb.Memb_Programs    AS Programs, "_ 
       & "  Memb.Memb_Memo        AS Memo, "_ 
       & "  Memb.Memb_Jobs        AS Jobs "_
       & "FROM "_         
       & "  Memb LEFT OUTER JOIN "_
       & "  Crit ON Memb.Memb_AcctId = Crit.Crit_AcctId AND Memb.Memb_Criteria = Crit.Crit_No "_
       & "WHERE "_     
       & "      (isnumeric(Memb.Memb_Criteria) = 1) "_  
       & "  AND (Memb.Memb_AcctId = '" & svCustAcctId & "') "_ 
       & "  AND (Memb.Memb_Level = 2) "_  
       & "  AND (Memb.Memb_Internal = 0) "_
       & "  AND (Memb.Memb_Active = 1) "

' sDebug
' stop

  sExcelInit  '...initialize 

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
    sExcelRow
    oRs.MoveNext
  Loop
  Set oRs = Nothing
  sCloseDb
  
  sExcelClose  '...close the worksheet 

  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs         	 				   = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell       	   				 = oWs.Worksheets(1).Cells

    vRow = 1
    vCol = 0
    oCell.RowHeight(vRow) = 50

    vCol = vCol + 1 : oCell(vRow, vCol) = "Learner ID"		  : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "Group"		        : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "First Name"	    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Last Name"		    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Email Address"		: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "Password"        : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Programs"        : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 20
    vCol = vCol + 1 : oCell(vRow, vCol) = "Memo"            : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "Jobs"            : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 20

  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0

    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Learner ID").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Group").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("First Name").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Last Name").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Email Address").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Password").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Programs").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Memo").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Jobs").Value

  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save svCustId & "_LEARNERS.xls", 1
    Response.End
  End Sub
 
%>

