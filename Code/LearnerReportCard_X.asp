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
  Dim oWs, oCell, oStyleC, oStyleR, oStyleL, oStyleI, vRow, vCol, vRowMax

  Dim vCustId, vFind, vFindId, vFindFailing, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFindActive, vFindCompleted, vFindBookmarks, vStrDate, vEndDate
  Dim vLearner, vPassword, vMods, vModNo, aMods, vTimeSpent, aBest, vBestScore, vBestDate, vNoAttempts, vExam_Id, aMemo, vTitle
  Dim vProgId, vProgTitle, vProgMods, vProgAssessment, vProgAssessmentScore, vProgExam, aTimeSpent, vCompleted
  
  Dim vStrTime
  vStrTime = Now()
 
  vCompleted =fPhraH(000953) '...use this as a constant
  
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


  '...read the learners
  sOpenDb
  sOpenDb4
  vSql = "SELECT DISTINCT TOP 5000 Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Criteria, Memb.Memb_Level, Memb.Memb_Memo" _
       & "  FROM Memb WITH (NOLOCK) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo " _
       & "  WHERE "_
       & "    (CHARINDEX(Logs.Logs_Type, 'TP') > 0) "_
       & "    AND (Logs.Logs_AcctId = '" & vCust_AcctId & "') "_
       & "    AND (Logs.Logs_Posted BETWEEN '" & vStrDate & "' AND '" & vEndDate & "') "_

       & fIf(vFind = "S" And Len(vFindId) > 0,        " AND (Memb_Id        LIKE '"      & vFindId        & "%')   ", "" ) _
       & fIf(vFind = "S" And Len(vFindFirstName) > 0, " AND (Memb_FirstName LIKE '"      & vFindFirstName & "%')   ", "" ) _
       & fIf(vFind = "S" And Len(vFindLastName) > 0,  " AND (Memb_LastName  LIKE '"      & vFindLastName  & "%')   ", "" ) _
       & fIf(vFind = "S" And Len(vFindEmail) > 0,     " AND (Memb_Email     LIKE '"      & vFindEmail     & "%')   ", "" ) _
       & fIf(vFind = "S" And Len(vFindMemo) > 0,      " AND (Memb_Memo      LIKE '"      & vFindMemo      & "%')   ", "" ) _

       & fIf(vFind = "C" And Len(vFindId) > 0,        " AND (Memb_Id        LIKE '%"     & vFindId        & "%')   ", "" ) _
       & fIf(vFind = "C" And Len(vFindFirstName) > 0, " AND (Memb_FirstName LIKE '%"     & vFindFirstName & "%')   ", "" ) _
       & fIf(vFind = "C" And Len(vFindLastName) > 0,  " AND (Memb_LastName  LIKE '%"     & vFindLastName  & "%')   ", "" ) _
       & fIf(vFind = "C" And Len(vFindEmail) > 0,     " AND (Memb_Email     LIKE '%"     & vFindEmail     & "%')   ", "" ) _
       & fIf(vFind = "C" And Len(vFindMemo) > 0,      " AND (Memb_Memo      LIKE '%"     & vFindMemo      & "%')   ", "" ) _

       & fIf(vFindCriteria <> "0",                     "AND (CHARINDEX(Memb_Criteria, '" & vFindCriteria  & "') > 0)", "") _
       & fIf(vFindCriteria <> "0",                     "AND (Memb_Criteria <> '0')                                  ", "") _
       & fIf(vFindActive = "a",                        "AND (Memb_Active = 1)                                       ", "") _
       & fIf(vFindActive = "i",                        "AND (Memb_Active = 0)                                       ", "") _
       & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName"
'sDebug

  Set oRs = oDb.Execute(vSql)
  '...read learners (oRs)
  Do While Not oRs.Eof

    vMemb_No        = oRs("Memb_No")
    vMemb_Id        = oRs("Memb_Id")
    vMemb_FirstName = oRs("Memb_FirstName")
    vMemb_LastName  = oRs("Memb_LastName")
    vMemb_Criteria  = oRs("Memb_Criteria")
    vMemb_Level     = oRs("Memb_Level")
    vMemb_Memo      = oRs("Memb_Memo")

    vPassword = "********"
    If (svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And InStr(vMemb_Id, vPasswordx) = 0 Then
      vPassword = vMemb_Id & fIf(vMemb_Level = 3, " *", "") & fIf(vMemb_Level = 4, " **", "")
    End If

    '...read programs (oRs4) that each learner has accessed
    vSql = "SELECT DISTINCT "_
         & "  LEFT(Logs.Logs_Item, 7) AS [Prog_Id], "_ 
         & "  V5_Base.dbo.Prog.Prog_Title1 AS [Prog_Title], "_
         & "  V5_Base.dbo.Prog.Prog_Mods, "_
         & "  V5_Base.dbo.Prog.Prog_Assessment, "_
         & "  V5_Base.dbo.Prog.Prog_AssessmentScore, "_
         & "  V5_Base.dbo.Prog.Prog_Exam "_
         & "FROM Logs WITH (nolock)  "_
         & "  INNER JOIN V5_Base.dbo.Prog WITH (NOLOCK) ON LEFT(Logs.Logs_Item, 7) = V5_Base.dbo.Prog.Prog_Id "_
         & "WHERE (Logs.Logs_MembNo = " & vMemb_No & ") " _
         & "  AND (Logs.Logs_Type = 'P') "_
         & "  AND (LEFT(Logs.Logs_Item, 7) <> 'P0000XX') "_
         & "  AND (Logs.Logs_Posted BETWEEN '" & vStrDate & "' AND '" & vEndDate & "') "_
         & "ORDER BY " _
         & "  V5_Base.dbo.Prog.Prog_Title1 "         
'   sDebug
    Set oRs4 = oDb4.Execute(vSql)

'   vSql = vMemb_No & " " & vRow : sDebug

    Do While Not oRs4.Eof    
      vProgId              = oRs4("Prog_Id")
      vProgTitle           = oRs4("Prog_Title")
      vProgMods            = oRs4("Prog_Mods")
      vProgAssessment      = oRs4("Prog_Assessment")
      vProgAssessmentScore = oRs4("Prog_AssessmentScore")
      vProgExam            = oRs4("Prog_Exam")
      
      '...see if there are any platform exams or program assessments
      vMods = vProgMods
      If Len(vProgExam) > 2 Then 
        vMods = vMods & " " & Mid(vProgExam, 22, 6) & "_E" '...add on _E so we know not to display the modid - actually an exam id
      End If
      If Len(vProgAssessment) > 2 Then 
        vMods = vMods & " " & vProgAssessment 
      End If

      aMods = Split(vMods)
      For vModNo = 0 To Ubound(aMods)
        If Len(aMods(vModNo)) = 8 Then 
          vTimeSpent  = 0
        Else  
          aTimeSpent  = Split(spLogsTimeSpent(vMemb_No, vProgId, Left(aMods(vModNo), 6), vStrDate, vEndDate), "|")
          vTimeSpent  = aTimeSpent(1)
        End If
'       vSql = "T/S" : sDebug

        aBest = Split(spLogsBestValues(vMemb_No, Left(aMods(vModNo), 6), vStrDate, vEndDate), "|")
        If Ubound(aBest) = 0 Then
          vBestScore = 0
          vBestDate = ""
        Else
          vBestScore = Cint(aBest(0))
          vBestDate  = aBest(1)
        End If

        '...see if we want to display failing scores (unless it's "Completed")
        If vBestScore <> -999 Then
          If vFindFailing = "n" And (vBestScore / 100 < fIf(vProgAssessmentScore = 0, .8, vProgAssessmentScore)) Then
            vBestScore = -1
          End If
        End If
'       vSql = "B/S" : sDebug

        vNoAttempts = spLogsAttempts(vMemb_No, Left(aMods(vModNo), 6), vStrDate, vEndDate)             
'       vSql = "N/A" : sDebug

        '...write out worksheet line for modules
        sExcelDetails
        
      Next
  
      oRs4.MoveNext
    Loop 
    Set oRs4 = Nothing

    sOrphans

    oRs.MoveNext
  Loop 
  Set oRs = Nothing

  sCloseDb
  sCloseDb4

' vSql = "Seconds expired: " & DateDiff("s", vStrTime, Now) : sDebug


  Sub sOrphans
    vSql = " SELECT DISTINCT"_ 
         & "   Logs.Logs_MembNo, "_
         & "   MAX(Logs.Logs_Posted) AS [Best Date], "_
         & "   LEFT(Logs.Logs_Item, 6) AS [Exam Id], "_
         & "   MAX(RIGHT(Logs.Logs_Item, 3)) AS [Best Score], "_
         & "   V5_Base.dbo.TstH.TstH_Title AS [Exam Title], "_
         & "   MAX(SUBSTRING(Logs.Logs_Item, 8, 1)) AS [Attempts], "_
         & "   Memb.Memb_FirstName, "_
         & "   Memb.Memb_LastName, "_
         & "   Memb.Memb_No "_
  
         & " FROM"_  
         & "   Logs               WITH (NOLOCK) LEFT OUTER JOIN "_ 
         & "   Catl_Prog          WITH (NOLOCK) ON Logs.Logs_AcctId = Catl_Prog.Catl_Prog_AcctId AND LEFT(Logs.Logs_Item, 6) <> Catl_Prog.Catl_Prog_ExamId INNER JOIN "_ 
         & "   V5_Base.dbo.TstH   WITH (NOLOCK) ON LEFT(Logs.Logs_Item, 6) = V5_Base.dbo.TstH.TstH_Id INNER JOIN "_
         & "   Memb               WITH (NOLOCK) ON Logs.Logs_MembNo = Memb.Memb_No "_
  
         & " WHERE"_ 
         & "   (LEN(Logs.Logs_Item) = 12) AND (RIGHT(Logs.Logs_Item, 3) <> '000') AND (Logs.Logs_Type = 'T') AND (Logs.Logs_AcctId = '" & svCustAcctId & "') AND (Logs.Logs_MembNo = " & vMemb_No & ") "_
         & "  GROUP BY "_ 
         & "   Logs.Logs_MembNo, LEFT(Logs.Logs_Item, 6), V5_Base.dbo.TstH.TstH_Title, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_No "_
         & "  ORDER BY "_ 
         & "   V5_Base.dbo.TstH.TstH_Title "
  
  '       sDebug
    Set oRs3 = oDb.Execute(vSql)
    Do While Not oRs3.Eof
      vBestScore = Cint(oRs3("Best Score"))
      '...write out worksheet line for modules
      sExcelOrphans
      oRs3.MoveNext
    Loop 

  End Sub


  '...close the worksheet 
  sExcelClose

  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell                    = oWs.Worksheets(1).Cells

    Set oStyleC      	 		  		 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
 
    oStyleC.HorizontalAlignment  = 2     '...center align
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234

    vRow = 1
    oCell.RowHeight(vRow) = 20
    oCell(vRow, 09).Style = oStyleR
    oCell(vRow, 10).Style = oStyleR
    oCell(vRow, 11).Style = oStyleR

    oCell(vRow, 01) = fPhraH(000369)             : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 18
    oCell(vRow, 02) = fPhraH(000156)        : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 12
    oCell(vRow, 03) = fPhraH(000163)         : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 12
    oCell(vRow, 04) = fIf(svCustPwd, fPhraH(000411), fPhraH(000211))                  
                                                              oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 16
    oCell(vRow, 05) = fPhraH(000201)           : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 12
    oCell(vRow, 06) = fPhraH(000019)             : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 48
    oCell(vRow, 07) = fPhraH(000793) : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 12
    oCell(vRow, 08) = fPhraH(000019)             : oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 48
    oCell(vRow, 09) = fPhraH(000552)        : oCell(vRow, 09).Format.Font.Bold = True : oCell.ColumnWidth(09) = 12    
    oCell(vRow, 10) = fPhraH(000792)        : oCell(vRow, 10).Format.Font.Bold = True : oCell.ColumnWidth(10) = 12   
    oCell(vRow, 11) = fPhraH(000361)        : oCell(vRow, 11).Format.Font.Bold = True : oCell.ColumnWidth(11) = 12
    oCell(vRow, 12) = fPhraH(000832)         : oCell(vRow, 12).Format.Font.Bold = True : oCell.ColumnWidth(12) = 12
    oCell(vRow, 13) = fPhraH(000173)              : oCell(vRow, 13).Format.Font.Bold = True : oCell.ColumnWidth(13) = 16
                                                                                                        oCell.ColumnWidth(14) = 16
                                                                                                        oCell.ColumnWidth(15) = 16
                                                                                                        oCell.ColumnWidth(16) = 16
                                                                                                        oCell.ColumnWidth(17) = 16
  End Sub



  '...write out details
  Sub sExcelDetails
    vRow = vRow + 1 
    
    oCell(vRow, 11).Style = oStyleR
    oCell(vRow, 12).Style = oStyleC
    
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
        oCell(vRow, 11) = fExcelDate(vCompleted)
      End If
      oCell(vRow, 12) = fExcelDate(vBestDate)
    Else
      oCell(vRow, 11) = fIf(vBestScore <=0, "", vBestScore)
      oCell(vRow, 12) = fIf(vBestScore <=0, "", fExcelDate(vBestDate))
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
    oCell(vRow, 12) = fExcelDate(oRs3("Best Date").Value)

    If vRow > vRowMax Then sExcelClose '...dont' blow memory

  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Dim vTitle
    vTitle = fPhraH(000794) & " " & fFormatDate(Now) 
    If vRow > vRowMax Then vTitle = vTitle  & " (Aborted after " & vRow - 1 & " lines!)"
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTitle & ".xls", 1
    Response.End
  End Sub

%>


