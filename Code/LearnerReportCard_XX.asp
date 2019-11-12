<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim oWs, oCell, vRow, vCol, oStyleD, oStyleR, oStyleL, oStyleI, aMemo, vActive
  
  If Request("vFindActive") = "a" Then
    vActive = 1
  ElseIf Request("vFindActive") = "i" Then
    vActive = 0
  Else
    vActive = ""
  End If


  sExcelInit

  vSql = "SELECT TOP 50000 * FROM vCSV WHERE "_
       & "[Acct Id] = '" & svCustAcctId & "' AND "_ 
       & "[Date] BETWEEN '" & Request("vStrDate") & "' AND '" & Request("vEndDate") & "'" _
       & fIf(vActive <> "", " AND [Active] = " & vActive, "")

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
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
    Set oStyleD      	 	  	  	 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
  
    oStyleD.Number      				 = 14    '...format date m/d/yy
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234

    vRow = 1
    vCol = 0

    oCell.RowHeight(vRow) = 50
    oCell(vRow, 09).Style = oStyleR
    oCell(vRow, 10).Style = oStyleR

    vCol = vCol + 1 : oCell(vRow, vCol) = "Group 1"						: oCell.ColumnWidth(vCol) = 20
    vCol = vCol + 1 : oCell(vRow, vCol) = "Group 2"						: oCell.ColumnWidth(vCol) = 08
    vCol = vCol + 1 : oCell(vRow, vCol) = "ID/Password"				: oCell.ColumnWidth(vCol) = 20
    vCol = vCol + 1 : oCell(vRow, vCol) = "Last Name"					: oCell.ColumnWidth(vCol) = 14
    vCol = vCol + 1 : oCell(vRow, vCol) = "First Name"				: oCell.ColumnWidth(vCol) = 14
    vCol = vCol + 1 : oCell(vRow, vCol) = "Active"				    : oCell.ColumnWidth(vCol) = 08
    vCol = vCol + 1 : oCell(vRow, vCol) = "Assessment ID"			: oCell.ColumnWidth(vCol) = 14
    vCol = vCol + 1 : oCell(vRow, vCol) = "Assessment Title"	: oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "Score"							: oCell.ColumnWidth(vCol) = 08
    vCol = vCol + 1 : oCell(vRow, vCol) = "Date"							: oCell.ColumnWidth(vCol) = 20
    vCol = vCol + 1 : oCell(vRow, vCol) = "Memo"							: oCell.ColumnWidth(vCol) = 20    

    For vCol = 1 to 11 : oCell(vRow, vCol).Format.Font.Bold = True : Next

  End Sub


  '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0

    oCell(vRow, 09).Style = oStyleR
    oCell(vRow, 10).Style = oStyleD

    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Group 1").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Group 2").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Password").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Last Name").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("First Name").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = fYN (oRs("Active").Value)
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Assessment ID").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Assessment Title").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Score").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Date").Value
		aMemo = Split(fOkValue(oRs("Memo").Value), "|")
		For i = 0 To Ubound(aMemo)
	    vCol = vCol + 1 : oCell(vRow, vCol) = aMemo(i)
	  Next    
  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save svCustID & " Learner Scores taken between " & fFormatDate(Request("vStrDate")) & " and " &  fFormatDate(Request("vEndDate")) & ".xls", 1
    Response.End
  End Sub
%>



