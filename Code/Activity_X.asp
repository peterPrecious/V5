<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 

  Dim vNext, vActive, vMods, vModsOnly, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vStrDate
  Dim vMembIdLast, vBg, vTitle, vScore, vTimeSpent, vPassword
  Dim vLogsType, vLogsModule, vLogsValue, vLogsTimespent, vLogsPosted, aMemo         

  vCurList       = Request("vCurList") 
  vMaxList       = fDefault(Request("vMaxList"), 50)
  vStrDate       = Request("vStrDate") 
  vActive        = fDefault(Request("vActive"), "y")
  vModsOnly      = Request("vModsOnly")
  If Len(vModsOnly) > 0 Then
    vMods        = vModsOnly '...format: Activity.asp?vModsOnly=1234EN,1235EN,1440FR
  Else
    vMods        = Request("vMods")
  End If
  vFind          = fDefault(Request("vFind"), "S")
  vFindId        = fUnQuote(Request("vFindId"))
  vFindFirstName = fUnQuote(Request("vFindFirstName"))
  vFindLastName  = fUnQuote(Request("vFindLastName"))
  vFindEmail     = fNoQuote(Request("vFindEmail"))
  vFindMemo      = fUNQuote(Request("vFindMemo"))
  vFindCriteria  = Request("vFindCriteria")

  '...initialize spreadsheet
  sExcelInit

  vSql = "SELECT Memb_Criteria, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Level, Memb_Active, Memb_Memo,  " _
       & "CASE Len(Logs.Logs_Item) WHEN 21 THEN SUBSTRING (Logs.Logs_Item,9, 6) ELSE LEFT(Logs.Logs_Item, 6) END AS [Logs_Module], " _
       & "CAST(RIGHT(Logs.Logs_Item, 3) AS FLOAT) AS Logs_Value, " _
       & "SUBSTRING(Logs.Logs_Item, 8, 1) AS Logs_Attempt, " _
       & "Logs.Logs_Posted AS Logs_Posted, " _
       & "CASE LEN(Logs_Item) WHEN 21 THEN 'P' WHEN 10 THEN 'M' WHEN 12 THEN 'E' END AS [Logs_Type], " _
       & "RIGHT(Logs.Logs_Item, 6) AS [Logs_Timespent] " _

       & "FROM Memb WITH (nolock) " & fIf(vActive="y", "INNER", "LEFT OUTER") & " JOIN Logs WITH (nolock) " _ 

       & "ON Logs_MembNo = Memb_No " _    
       & "AND (CHARINDEX(Logs.Logs_Type, 'PT') > 0) " _
       & fIf(Len(Trim(vStrDate)) > 0, "AND (Logs.Logs_Posted >= '" & vStrDate & "') ", "" ) _  

       & fIf(Len(Trim(vMods)) > 0, "AND ((LEN(Logs.Logs_Item) = 21) AND (CHARINDEX(SUBSTRING(Logs.Logs_Item, 9, 6), '" & vMods & "') > 0) OR (LEN(Logs.Logs_Item) = 10) AND (CHARINDEX(LEFT(Logs.Logs_Item, 6), '" & vMods & "') > 0)) ", "" )_  


       & "WHERE (Memb_AcctId = '" & svCustAcctId & "') " _
       & "AND (Logs.Logs_AcctId = '" & svCustAcctId & "') " _             
       & "AND (Memb_Level <= " & svMembLevel & ") " _

       & fIf(vFind = "S" And Len(vFindId) > 0,        "AND (Memb_Id        LIKE '" & vFindId         & "%') ", "" ) _
       & fIf(vFind = "S" And Len(vFindFirstName) > 0, "AND (Memb_FirstName LIKE '" & vFindFirstName  & "%') ", "" ) _
       & fIf(vFind = "S" And Len(vFindLastName) > 0,  "AND (Memb_LastName  LIKE '" & vFindLastName   & "%') ", "" ) _
       & fIf(vFind = "S" And Len(vFindEmail) > 0,     "AND (Memb_Email     LIKE '" & vFindEmail      & "%') ", "" ) _
       & fIf(vFind = "S" And Len(vFindMemo) > 0,      "AND (Memb_Memo      LIKE '" & vFindMemo       & "%') ", "" ) _

       & fIf(vFind = "C" And Len(vFindId) > 0,        "AND (Memb_Id        LIKE '%" & vFindId        & "%') ", "" ) _
       & fIf(vFind = "C" And Len(vFindFirstName) > 0, "AND (Memb_FirstName LIKE '%" & vFindFirstName & "%') ", "" ) _
       & fIf(vFind = "C" And Len(vFindLastName) > 0,  "AND (Memb_LastName  LIKE '%" & vFindLastName  & "%') ", "" ) _
       & fIf(vFind = "C" And Len(vFindEmail) > 0,     "AND (Memb_Email     LIKE '%" & vFindEmail     & "%') ", "" ) _
       & fIf(vFind = "C" And Len(vFindMemo) > 0,      "AND (Memb_Memo      LIKE '%" & vFindMemo      & "%') ", "" ) _

       & fIf(Len(vFindCriteria)> 2,                   "AND (Memb_Criteria = '"      & vFindCriteria  & "')  ", "" ) _

       & "ORDER BY Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, " _
       & "CASE Len(Logs.Logs_Item) WHEN 21 THEN Substring(Logs_Item, 9, 6) ELSE LEFT(Logs_Item, 6) END, " _ 
       & "CASE LEN(Logs_Item) WHEN 21 THEN 'P' WHEN 10 THEN 'M' ELSE 'E' END DESC" 
 
'     sDebug

  sOpenDb
  Set oRs = oDb.Execute(vSql)

  '...read until either eof or end of group
  Do While Not oRs.Eof

    vMemb_No        = oRs("Memb_No")
    vMemb_Id        = oRs("Memb_Id")
    vMemb_FirstName = oRs("Memb_FirstName")
    vMemb_LastName  = oRs("Memb_LastName")
    vMemb_Active    = oRs("Memb_Active")
    vMemb_Level     = oRs("Memb_Level")
    vMemb_Memo      = oRs("Memb_Memo")
    vMemb_Criteria  = oRs("Memb_Criteria")
    
    vLogsType       = fOkValue(oRs("Logs_Type"))
    vLogsModule     = fOkValue(oRs("Logs_Module"))
    vLogsValue      = fOkValue(oRs("Logs_Value"))
    vLogsTimespent  = fOkValue(oRs("Logs_Timespent"))
    vLogsPosted     = fOkValue(oRs("Logs_Posted"))
    
    If Len(vLogsType) > 0 Then
      If vLogsType = "E" Then
        vTitle = fExamTitle(vLogsModule)
        vScore = vLogsValue
      Else
        vTitle = fModsTitle(vLogsModule)
        '...store the Score and print with next record (timespent)
        If vLogsType = "M" Then
          vScore = vLogsValue
        Else
          vTimeSpent = Cdbl(vLogsTimespent)
        End If
      End If
       
      vTitle = Replace(vTitle, "<b>", "")        
      vTitle = Replace(vTitle, "<B>", "")        
      vTitle = Replace(vTitle, "</b>", "")        
      vTitle = Replace(vTitle, "</B>", "")        
      vTitle = vTitle
    Else
      vTitle = ""
    End If
            
    vPassword = "********"
    If (svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And InStr(vMemb_Id, vPasswordx) = 0 Then
      vPassword = vMemb_Id & fIf(vMemb_Level = 3, " *", "") & fIf(vMemb_Level = 4, " **", "")
    End If
  
    '...write out worksheet line
    sExcelRow 

    vScore = ""
    vTimeSpent = ""
    oRs.MoveNext
  Loop 

  Set oRs  = Nothing
  sCloseDb

 
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
'   oCell(vRow, 05).Style = oStyleR
    oCell(vRow, 09).Style = oStyleR
    oCell(vRow, 10).Style = oStyleR

    oCell(vRow, 01) = fPhraH(000369)             : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 16
    oCell(vRow, 02) = fPhraH(000156)        : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 16
    oCell(vRow, 03) = fPhraH(000163)         : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 16
    oCell(vRow, 04) = fIf(svCustPwd, fPhraH(000411), fPhraH(000211))                  
                                                              oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 16
    oCell(vRow, 05) = fPhraH(000551)           : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 16
    oCell(vRow, 06) = fPhraH(000272)            : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 16    
    oCell(vRow, 07) = fPhraH(000019)             : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 40
    oCell(vRow, 08) = fPhraH(000552)        : oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 16
    oCell(vRow, 09) = fPhraH(000232)             : oCell(vRow, 09).Format.Font.Bold = True : oCell.ColumnWidth(09) = 16
    oCell(vRow, 10) = fPhraH(000112)              : oCell(vRow, 10).Format.Font.Bold = True : oCell.ColumnWidth(10) = 16
    oCell(vRow, 11) = fPhraH(000173)              : oCell(vRow, 11).Format.Font.Bold = True : oCell.ColumnWidth(11) = 16
   
  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 10).Style = oStyleD

    oCell(vRow, 01) = fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", fCriteria(vMemb_Criteria))
    oCell(vRow, 02) = vMemb_FirstName
    oCell(vRow, 03) = vMemb_LastName
    oCell(vRow, 04) = vPassword
    oCell(vRow, 05) = fYN(vMemb_Active)
    oCell(vRow, 06) = fIf(vLogsType = "E", "Exam", vLogsModule)
    oCell(vRow, 07) = vTitle
    oCell(vRow, 08) = vTimeSpent
    oCell(vRow, 09) = vScore
    oCell(vRow, 10) = vLogsPosted


    '...post results - each in it's own cell (if test then only one cell)
    vMemb_Memo = fOkValue(vMemb_Memo) '...if null set to ""
    aMemo = Split(vMemb_Memo, "|")
    For i = 0 To Ubound(aMemo)
      oCell(vRow, 11 + i) = aMemo(i)
    Next
    
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Dim vTitle
    vTitle = fPhraH(000562)
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTitle & " " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

%>

