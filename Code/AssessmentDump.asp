<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include file = "LearnerReportCard_Functions.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol, vRowMax

  Dim vCustId, vFind, vFindId, vFindFailing, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFindActive, vFindCompleted, vFindBookmarks, vStrDate, vEndDate
  Dim vLearner, vPassword, vMods, vModNo, aMods, vTimeSpent, aBest, vBestScore, vBestDate, vNoAttempts, vExam_Id, aMemo, vTitle
  Dim vProgId, vProgTitle, vProgMods, vProgAssessment, vProgAssessmentScore, vProgExam, aTimeSpent, vCompleted
  
  Dim vStrTime
  vStrTime = Now()
 
  vCompleted = "Completed" '...use this as a constant
  
  vRowMax = 5000
 
  vStrDate       = Request("vStrDate") 
  vEndDate       = Request("vEndDate") 
  vCustId        = fDefault(Request("vCustId"), svCustId)
  vFind          = Request("vFind")
  vFindId        = fUnQuote(Request("vFindId"))
  vFindFailing   = Request("vFindFailing")
  vFindCompleted = Request("vFindCompleted")
  vFindBookmarks = Request("vFindBookmarks")
  vFindFirstName = fUnQuote(Request("vFindFirstName"))
  vFindLastName  = fUnQuote(Request("vFindLastName"))
  vFindEmail     = fNoQuote(Request("vFindEmail"))
  vFindMemo      = fUnQuote(Request("vFindMemo"))
  vFindCriteria  = Request("vFindCriteria")
  vFindActive    = Request("vFindActive")

  sGetCust vCustId

  '...initialize spreadsheet
  sExcelInit

  sOpenDb
  vSql = " SELECT [Group], [First Name], [Last Name], [Password], [Assessment Id], [Assessment Title], [Last Score], [Best Score], [No Attempts], [Memo]"_
       & " FROM V5_Vubz.dbo.vHistory"_
       & " WHERE [Acct Id] = '" & svCustAcctId & "'"_
       & " ORDER BY [Group], [Last Name], [First Name], [Assessment Title]"
  Set oRs = oDb.Execute(vSql)

  Do While Not oRs.Eof

    vMemb_No        = oRs("Memb_No")
    vMemb_Id        = oRs("Memb_Id")
    vMemb_FirstName = oRs("Memb_FirstName")
    vMemb_LastName  = oRs("Memb_LastName")
    vMemb_Criteria  = oRs("Memb_Criteria")
    vMemb_Level     = oRs("Memb_Level")
    vMemb_Memo      = oRs("Memb_Memo")

    sExcelDetails      

    oRs.MoveNext
  Loop 
  Set oRs = Nothing

  sCloseDb
  sCloseDb4

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

    oCell.RowHeight(vRow) = 20
    oCell(vRow, 09).Style = oStyleR
    oCell(vRow, 10).Style = oStyleR
    oCell(vRow, 11).Style = oStyleR

  vSql = " SELECT [Group], [First Name], [Last Name], [Password], [Assessment Id], [Assessment Title], [Last Score], [Best Score], [No Attempts], [Memo]"_



    oCell(vRow, 01) = "Group"             											: oCell(vRow, 01).Format.Font.Bold = True :	oCell.ColumnWidth(01) = 18
    oCell(vRow, 02) = "First Name"        											: oCell(vRow, 02).Format.Font.Bold = True :	oCell.ColumnWidth(02) = 12
    oCell(vRow, 03) = "Last Name"         											: oCell(vRow, 03).Format.Font.Bold = True : 	oCell.ColumnWidth(03) = 12
    oCell(vRow, 04) = fIf(svCustPwd, "Learner Id", "Password")  : oCell(vRow, 03).Format.Font.Bold = True : 	oCell.ColumnWidth(03) = 12                
    oCell(vRow, 06) = "Title"             											: oCell(vRow, 06).Format.Font.Bold = True :	oCell.ColumnWidth(06) = 48
    oCell(vRow, 07) = "Assessment Id" 													: oCell(vRow, 07).Format.Font.Bold = True : 	oCell.ColumnWidth(07) = 12

    oCell(vRow, 10) = "# Attempts"        											: oCell(vRow, 10).Format.Font.Bold = True : 	oCell.ColumnWidth(10) = 12   
    oCell(vRow, 11) = "Best Score"        											: oCell(vRow, 11).Format.Font.Bold = True : 	oCell.ColumnWidth(11) = 12
    oCell(vRow, 12) = "Date"         											: oCell(vRow, 12).Format.Font.Bold = True : 	oCell.ColumnWidth(12) = 12
    oCell(vRow, 13) = "Memo"              											: oCell(vRow, 13).Format.Font.Bold = True :	oCell.ColumnWidth(13) = 16
																																							oCell.ColumnWidth(14) = 16
                                                                     	                                	oCell.ColumnWidth(15) = 16
                                                                           		                         	oCell.ColumnWidth(16) = 16
                                                                                  		                  	oCell.ColumnWidth(17) = 16
  End Sub



  '...write out details
  Sub sExcelDetails
    vRow = vRow + 1 
    
    oCell(vRow, 11).Style = oStyleR
    
    oCell(vRow, 01) = fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", fCriteria(vMemb_Criteria))
    oCell(vRow, 02) = vMemb_FirstName
    oCell(vRow, 03) = vMemb_LastName
    oCell(vRow, 04) = vPassword
    oCell(vRow, 05) = vProgId
    oCell(vRow, 06) = Replace(vProgTitle, "<br>", " - ")
    oCell(vRow, 07) = fIf(Len(aMods(vModNo)) = 6, aMods(vModNo), "Exam")
    oCell(vRow, 08) = fIf(Len(aMods(vModNo)) = 6, fModsTitle(aMods(vModNo)), "Exam")
    oCell(vRow, 09) = fIf(vTimeSpent=0, "", vTimeSpent)
    oCell(vRow, 10) = fIf(vNoAttempts=0, "", vNoAttempts)


    If Cint(vBestScore) = -999 Then
      If vFindCompleted = "n" Then
        oCell(vRow, 11) = 100
      Else
        oCell(vRow, 11) = vCompleted
      End If
      oCell(vRow, 12) = vBestDate
    Else
      oCell(vRow, 11) = fIf(vBestScore <=0, "", vBestScore)
      oCell(vRow, 12) = fIf(vBestScore <=0, "", vBestDate)
    End If

    '...post results - each in it's own cell (if test then only one cell)
    vMemb_Memo = fOkValue(vMemb_Memo) '...if null set to ""
    aMemo = Split(vMemb_Memo, "|")
    For i = 0 To Ubound(aMemo)
      oCell(vRow, 13 + i) = aMemo(i)
    Next

    If vRow > vRowMax Then sExcelClose '...dont' blow memory

  End Sub


  '...write out orphans
  Sub sExcelOrphans
    vRow = vRow + 1 
    oCell(vRow, 01) = fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", fCriteria(vMemb_Criteria))
    oCell(vRow, 02) = vMemb_FirstName
    oCell(vRow, 03) = vMemb_LastName
    oCell(vRow, 04) = vPassword
    oCell(vRow, 05) = ""
    oCell(vRow, 06) = ""
    oCell(vRow, 07) = oRs3("Exam Id").Value
    oCell(vRow, 08) = oRs3("Exam Title").Value
    oCell(vRow, 09) = ""
    oCell(vRow, 10) = oRs3("Attempts").Value
    oCell(vRow, 11) = vBestScore
    oCell(vRow, 12) = fFormatDate(oRs3("Best Date").Value)

    If vRow > vRowMax Then sExcelClose '...dont' blow memory

  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Dim vTitle
    vTitle = "Learner Report Card dated" & " " & fFormatDate(Now) 
    If vRow > vRowMax Then vTitle = vTitle  & " (Aborted after " & vRow - 1 & " lines!)"
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTitle & ".xls", 1
    Response.End
  End Sub

%>


