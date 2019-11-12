<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Mods.asp"-->

<% 
  Server.ScriptTimeout = 60 * 60

  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol, vCnt

'  Dim vId_Pwd, vStrDate, vBold, vGrade, vTestExam, vTitle, vCertUrl, vCertType, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vUrl

  Dim vStrDate, vGrade, vTestExam, vTitle, vCertUrl, vCertType, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vUrl, aMemo, vDetails, vResults, aResults
  Dim vLogs_No, vLogs_AcctId, vLogs_Type, vLogs_Item, vLogs_Posted, vLogs_MembNo, vLogs_Module, vLogs_Result

  vDetails       = Request("vDetails") 
  vStrDate       = Request("vStrDate")
  vFind          = fDefault(Request("vFind"), "S")
  vFindId        = fUnQuote(Request("vFindId"))
  vFindFirstName = fUnQuote(Request("vFindFirstName"))
  vFindLastName  = fUnQuote(Request("vFindLastName"))
  vFindEmail     = fNoQuote(Request("vFindEmail"))


  vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Memo, Memb.Memb_Level "
  If vDetails = "y" Then  '...details of assessments
    vSql = vSql & ", Left(Logs.Logs_Item, 6) AS Logs_Module, CAST(Right(Logs.Logs_Item, 3) AS FLOAT) AS Logs_Result, Logs.Logs_Posted "
  ElseIf vDetails = "n" Then '...summary
    vSql = vSql & ", Left(Logs.Logs_Item, 6) AS Logs_Module,  MAX(CAST(Right(Logs.Logs_Item, 3) AS FLOAT)) AS Logs_Result, MAX(Logs.Logs_Posted) AS Logs_Posted "
  ElseIf vDetails = "s" Then '...details of surveys
    vSql = vSql & ", SUBSTRING(Logs.Logs_Item, 9, 6) AS Logs_Module,  SUBSTRING(Logs.Logs_Item, 16, 999) AS Logs_Result, Logs.Logs_Posted "
  End If

  vSql = vSql & " FROM Logs INNER JOIN Memb ON Logs_MembNo = Memb_No "
  vSql = vSql & " WHERE Logs_AcctId= '" & svCustAcctId & "' AND Logs_Type = '" & fIf(vDetails = "s", "U", "T") & "'"
  vSql = vSql & " AND Logs.Logs_Posted > '" & vStrDate & "'"
  vSql = vSql & " AND Memb_Level < 4 "

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

  If vDetails = "y" or vDetails = "s" Then    
    vSql = vSql & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo, Logs.Logs_Posted "'
  Else
    vSql = vSql & " GROUP BY Memb.Memb_LastName, Memb.Memb_FirstName, Left(Logs.Logs_Item, 6), Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo, Memb.Memb_Level"
    vSql = vSql & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Left(Logs.Logs_Item, 6), Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo "
  End If

  '...only allow 10000 records (else system will blow)
  vCnt = 0

  sOpenDB
  Set oRs = oDB.Execute(vSql)

  '...initialize 
  sExcelInit

  '...read until either eof or end of group
  Do While Not oRs.Eof

    vLogs_Module                = oRs("Logs_Module")
    vLogs_Result                = oRs("Logs_Result")
    vLogs_Posted                = oRs("Logs_Posted")

    vMemb_Level                 = oRs("Memb_Level")
    vMemb_No                    = oRs("Memb_No")
    vMemb_Id                    = oRs("Memb_Id")
    vMemb_Id                    = fIf(vMemb_Level > 2, "******", fDefault(vMemb_Id, "N/A"))

    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")
    vMemb_Memo									= oRs("Memb_Memo")

    aMemo                       = Split(vMemb_Memo, "|")
    If Ubound(aMemo) < 5 Then 
      vMemb_Memo                = vMemb_Memo & "||||"
      aMemo                     = Split(vMemb_Memo, "|")
    End If
    
    vTitle = fModsTitle(vLogs_Module)

    '...if over 10k it will blow
    vCnt = vCnt + 1
'   If vCnt > 10000 Then Exit Do
    
    '...write out worksheet line
    sExcelRow 
      
    oRs.MoveNext
  Loop 
  
  '...close the worksheet 
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
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
    
    vRow = 1
    oCell.RowHeight(vRow) = 50

    oCell(vRow, 08).Style = oStyleR

    oCell(vRow, 01) = "Learner"        : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 20
    oCell(vRow, 02) = "Email"          : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 36

    oCell(vRow, 03) = "Student Id"     : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 16
    oCell(vRow, 04) = "Institution"    : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 24
    oCell(vRow, 05) = "Faculty"        : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 24
    oCell(vRow, 06) = "Course"         : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 24
    oCell(vRow, 07) = "Academic Year"  : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 16
    
    oCell(vRow, 08) = "Date"					 : oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 12
    oCell(vRow, 09) = "Title"          : oCell(vRow, 09).Format.Font.Bold = True : oCell.ColumnWidth(09) = 30
    oCell(vRow, 10) = "Score | Result" : oCell(vRow, 10).Format.Font.Bold = True
  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 03).Style = oStyleI
    oCell(vRow, 08).Style = oStyleD

    oCell(vRow, 01)       = vMemb_LastName & ", " & vMemb_FirstName
    oCell(vRow, 02)       = vMemb_Id

    oCell(vRow, 03)       = aMemo(0)
    oCell(vRow, 04)       = aMemo(1)
    oCell(vRow, 05)       = aMemo(2)
    oCell(vRow, 06)       = aMemo(4)
    oCell(vRow, 07)       = aMemo(3)

    oCell(vRow, 08)       = vLogs_Posted '...this allows sorting
    oCell(vRow, 09)       = vTitle
    If vDetails <> "s" Then
      oCell(vRow, 10)     = FormatPercent(vLogs_Result/100,0)
    Else
      '...post results - each in it's own cell (if test then only one cell)
      aResults = Split(vLogs_Result, "|")
      For i = 0 To Ubound(aResults)
        oCell(vRow, 10 + i) = aResults(i)
      Next
    End If 




  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Assessment|Survey Report dated " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

  Function fId (i)
    fId = fDefault(i, "N/A")
    If oRs("Memb_Level") > 2 Then 
      fId = "******"
    ElseIf IsNumeric(i) Then
      If Left(i, 1) = "0" Then
        fId = "'" & fId
      End If
    End If
  End Function
 
  
%>

