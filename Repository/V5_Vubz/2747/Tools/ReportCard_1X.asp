<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Mods.asp"-->
<!--#include virtual = "V5\Inc\Db_Crit.asp"-->

<% 
  Server.ScriptTimeout = 60 * 60

  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol, vCnt

  Dim vStrDate, vBold, vGrade, vTestExam, vTitle, vCertUrl, vCertType, vFind, vFindId, vFindFirstName, vFindLastName, vFindCriteria, vFindEmail, vUrl, aMemo, aCrit
  Dim vLogs_No, vLogs_AcctId, vLogs_Type, vLogs_Item, vLogs_Posted, vLogs_MembNo
  Dim vLogs_Module, vScore, vCurList, vMaxList, vSum, vCriteria
  
  vCurList        = Request("vCurList") 
  vStrDate        = Request("vStrDate")
  vFind           = fDefault(Request("vFind"), "S")
  vFindId         = fUnQuote(Request("vFindId"))
  vFindFirstName  = fUnQuote(Request("vFindFirstName"))
  vFindLastName   = fUnQuote(Request("vFindLastName"))
  vFindCriteria   = Request("vFindCriteria") 
  vFindEmail      = fNoQuote(Request("vFindEmail"))

  vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Memo, Memb.Memb_Criteria, " 
  vSql = vSql & " AVG(CAST(RIGHT(Logs.Logs_Item, 3) AS FLOAT)) AS [Score], MAX(Logs.Logs_Posted) AS Logs_Posted, SUM(1) AS [Sum]"
  vSql = vSql & " FROM Logs WITH (NOLOCK) INNER JOIN Memb WITH (NOLOCK) ON Logs_MembNo = Memb_No "
  vSql = vSql & " WHERE (Logs_AcctId= '" & svCustAcctId & "') AND (Logs_Type = 'T')"

  If vFindCriteria <> "0" Then
    vSql = vSql & " AND (CHARINDEX(Memb.Memb_Criteria, '" & vFindCriteria & "') > 0)"
  End If

  vSql = vSql & " AND (CHARINDEX(LEFT(Logs.Logs_Item, 4), '9427 9495 9497 9498') > 0) "
  vSql = vSql & " AND (Logs.Logs_Posted > '" & vStrDate & "')"
  vSql = vSql & " AND (Memb_Level = 2) "
  vSql = vSql & " AND (Memb.Memb_Criteria <> '0')"

  If vFind = "S" Then
    If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '" & vFindId         & "%')"
    If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '" & vFindFirstName  & "%')"
    If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '" & vFindLastName   & "%')"
    If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '" & vFindEmail      & "%')"
  Else
    If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '%" & vFindId        & "%')"
    If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '%" & vFindFirstName & "%')"
    If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '%" & vFindLastName  & "%')"
    If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '%" & vFindEmail     & "%')"
  End If

  vSql = vSql & " GROUP BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo "
  vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo "

' sDebug

  sOpenDb
  Set oRs = oDb.Execute(vSql)

  '...initialize 
  sExcelInit

  '...read until either eof or end of group
  Do While Not oRs.Eof

    vScore                      = oRs("Score")
    vLogs_Posted                = oRs("Logs_Posted")
    vSum                        = oRs("Sum")

    vMemb_No                    = oRs("Memb_No")
    vMemb_Id                    = oRs("Memb_Id")

    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")

    vMemb_Criteria   			      = oRs("Memb_Criteria")
	  vCriteria				            = fCriteria (vMemb_Criteria)
	  aCrit					              = Split(vCriteria , "|")
    If Ubound(aCrit) <> 1 Then 
      aCrit                     = Split("||", "|")
    End If

    vMemb_Memo					        = fOkValue(oRs("Memb_Memo"))
    aMemo                       = Split(vMemb_Memo, "|")
    If Ubound(aMemo) < 5 Then 
      vMemb_Memo                = vMemb_Memo & "||||"
      aMemo                     = Split(vMemb_Memo, "|")
    End If
    
    vSum                        = Cint(oRs("Sum"))

   '...write out worksheet line
   If vSum = 4 Then sExcelRow 
      
    oRs.MoveNext
  Loop 
  
  '...close the worksheet 
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs								 = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell							 = oWs.Worksheets(1).Cells

    Set oStyleD      	 	  	  		 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
  
    oStyleD.Number      				 = 14    '...format date m/d/yy
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234
    
    vRow = 1
    oCell.RowHeight(vRow) = 50

    oCell(vRow, 02).Style = oStyleR
    oCell(vRow, 07).Style = oStyleR

    oCell(vRow, 01) = "Name"           : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 20
    oCell(vRow, 02) = "Score"          : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 12
    oCell(vRow, 03) = "Student #"      : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 16
    oCell(vRow, 04) = "Institution"    : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 24
    oCell(vRow, 05) = "Faculty"        : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 24
    oCell(vRow, 06) = "Course"	  	   : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 24
    oCell(vRow, 07) = "Completed"      : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 16
    oCell(vRow, 08) = "Email"          : oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 36
  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 02).Style = oStyleR
    oCell(vRow, 03).Style = oStyleI
    oCell(vRow, 07).Style = oStyleD

    oCell(vRow, 01)       = vMemb_LastName & ", " & vMemb_FirstName
    oCell(vRow, 02)       = FormatPercent(vScore/100,0)
    oCell(vRow, 03)       = aMemo(0)
    oCell(vRow, 04)       = aCrit(0)
    oCell(vRow, 05)       = aCrit(1)
    oCell(vRow, 06)       = aMemo(4)
    oCell(vRow, 07)       = vLogs_Posted
    oCell(vRow, 08)       = vMemb_Id


  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Assessment|Survey Report dated " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

 
%>

