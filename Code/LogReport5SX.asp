<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include virtual = "V5/Inc/Db_Parm.asp"-->



<% 
  '...Excel variables
  Dim oWs, oCell, oStyle, vRow, vCol 
  Dim vCookie, vLevel, vId_Pwd, vStrDate, vTitle, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vUrl, vResults, aResults, vCriteria, vParmNo, aGroup1, aMemo
  vCookie   = svCustAcctId & "_LogReport5"
  vId_Pwd = fIf(svCustPwd, fPhraH(000411), fPhraH(000211))
  
  vDetails       = Request.Cookies(vCookie)("vDetails") 
  vLevel         = Request.Cookies(vCookie)("vLevel")
  vCurList       = Request.Cookies(vCookie)("vCurList") 
  vStrDate       = Request.Cookies(vCookie)("vStrDate")
  vFind          = Request.Cookies(vCookie)("vFind")
  vFindId        = Request.Cookies(vCookie)("vFindId")
  vFindFirstName = Request.Cookies(vCookie)("vFindFirstName")
  vFindLastName  = Request.Cookies(vCookie)("vFindLastName")
  vFindEmail     = Request.Cookies(vCookie)("vFindEmail")
  vFindMemo      = Request.Cookies(vCookie)("vFindMemo")
  vFindCriteria  = Request.Cookies(vCookie)("vFindCriteria")
  vParmNo        = Request.Cookies(vCookie)("vParmNo")
      


  vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Criteria, Memb.Memb_Level, Logs.Logs_Item, Logs_Posted, Memb.Memb_Memo "
  vSql = vSql & " FROM Logs INNER JOIN Memb WITH (nolock) ON Logs_MembNo = Memb_No "
  vSql = vSql & " WHERE Memb_AcctId= '" & svCustAcctId & "'"
  vSql = vSql & " AND Logs.Logs_AcctId = '" & svCustAcctId & "'"
  vSql = vSql & " AND Logs.Logs_Type = 'U'"
  vSql = vSql & " AND Logs.Logs_Posted > '" & vStrDate & "'"
  vSql = vSql & " AND Memb.Memb_Level IN (" & vLevel & ")"

  If vFind = "S" Then
    If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '"  & vFindId         & "%')"
    If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '"  & vFindFirstName  & "%')"
    If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '"  & vFindLastName   & "%')"
    If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '"  & vFindEmail      & "%')"
    If Len(vFindMemo)      > 0 Then vSql = vSql & " AND (Memb_Memo      LIKE '"  & vFindMemo       & "%')"
  Else
    If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '%" & vFindId         & "%')"
    If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '%" & vFindFirstName  & "%')"
    If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '%" & vFindLastName   & "%')"
    If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '%" & vFindEmail      & "%')"
    If Len(vFindMemo)      > 0 Then vSql = vSql & " AND (Memb_Memo      LIKE '%" & vFindMemo      & "%')"
  End If

  '...Group1
  j = 0
  If Len(vFindCriteria)    > 1 Then 
    aGroup1 = Split(vFindCriteria)
    For i = 0 To Ubound(aGroup1)
      If Cint(aGroup1(i)) > 0 Then
        j = j + 1
        If j = 1 Then 
          vSql = vSql & " AND ((Memb_Criteria LIKE '%" & aGroup1(i) & "%')"
        Else
          vSql = vSql & "  OR (Memb_Criteria LIKE '%" & aGroup1(i) & "%')"
        End If
      End If
    Next
    If j > 0 Then 
       vSql = vSql & " )"
    End If         
  End If

  '...allow a module filter to be extracted from the vParm table via the url [?vParm=2] so report only displays modules required by this user - syntax must be perfect, ie:
' vSql = vSql & " AND CHARINDEX(LEFT(Logs.Logs_Item, 4), '0350|0225|0227|0334|0226|0333|0336|0335|0337|0338') > 0 "
  vSql = vSql & " " & fParmValue (vParmNo)
  vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Logs.Logs_Posted "'


  sOpenDB
  Set oRs = oDB.Execute(vSql)

  '...initialize 
  sExcelInit

  '...read until either eof or end of group
  Do While Not oRs.Eof

    sReadLogsMembSurvey  
    '...contains a program id 
    If Left(vLogs_Item, 1) = "P" Then  
      vLogs_Module = Mid(vLogs_Item, 9, 6)
      vResults = Mid(vLogs_Item, 16) 
    '...contains an 'undefined' program id 
    Else
      vLogs_Module = Mid(vLogs_Item, 11, 6)
      vResults = Mid(vLogs_Item, 18) 
    End If

    vTitle   = fModsTitle(vLogs_Module)   

    '...ensure you can only see members with same criteria
'   If fCriteriaOk (svMembCriteria, vMemb_Criteria) Then

      vCriteria = fCriteria(vMemb_Criteria)

      '...get title
      If vLogs_Assess = "E" Then
        vTitle = vLogs_Module & " - " & fExamTitle(vLogs_Module)
      Else  
        vTitle = vLogs_Module & " - " & fModsTitle(vLogs_Module)
      End If
  
      '...write out worksheet line
      sExcelRow 
'   End If  
      
    oRs.MoveNext
  Loop 
  
  '...close the worksheet 
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell                    = oWs.Worksheets(1).Cells
    Set oStyle       	 	  	  	 = oWs.CreateStyle
    oStyle.Number      				   = 14    '...format date m/d/yy
    oStyle.HorizontalAlignment   = 3     '...right justify
    
    vRow = 1
    oCell.RowHeight(vRow) = 50

    oCell(vRow, 01) = fPhraH(000369)	        : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 20
    oCell(vRow, 02) = vId_Pwd                   	      : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 20
    oCell(vRow, 03) = fPhraH(000165)				: oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 36
    oCell(vRow, 04) = fPhraH(000112)					: oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 12
    oCell(vRow, 05) = fPhraH(000019)	      	: oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 36
    oCell(vRow, 06) = fPhraH(000173)          : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 16
    oCell(vRow, 07) = fPhraH(000475)    : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 12
  
  
  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 04).Style = oStyle

    oCell(vRow, 01)       = vCriteria
    oCell(vRow, 02)       = fId(vMemb_Id)
    oCell(vRow, 03)       = vMemb_LastName & ", " & vMemb_FirstName
    oCell(vRow, 04)       = vLogs_Posted '...this allows sorting
    oCell(vRow, 05)       = vTitle
    oCell(vRow, 06)       = vMemb_Memo

    '...post results - each in it's own cell (if test then only one cell)
    aResults = Split(vResults, "|")
    For i = 0 To Ubound(aResults)
      oCell(vRow, 07 + i) = aResults(i)
    Next

    If vRow > 10000 Then 
      vRow = vRow + 5
      oCell(vRow, 01)       = "Report terminated.  Too many records selected for this format....."
      sExcelClose  
    End If

  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Assessment Report dated " & fFormatDate(Now) & ".xls", 1
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



