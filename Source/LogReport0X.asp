<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->


<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 
  Dim vAccount, vLevel, vLearners, vStrDate, vEndDate

  vAccount  = Request("vAccount")
  vLevel    = Request("vLevel")
  vLearners = Request("vLearners")
  vStrDate  = Request("vStrDate")
  vEndDate  = Request("vEndDate")

  Dim vId, vModule, vTitle, vLearner, vTimeSpent, vLogItemLength, vLogItemTitle, vOk

  '...if summarizing to prog level just select the left 7 chars (ie P1001EN) else select all 14 chars (ie P1001EN|1234EN)
  vLogItemLength = 14: If vLevel = "prog" Then vLogItemLength = 7
  vLogItemTitle  = "<!--{{-->Modules<!--}}-->" : If vLevel = "Prog" Then vLogItemTitle  = "<!--{{-->Programs<!--}}-->" 

  vSql = "SELECT Memb.Memb_LastName + ',  ' + Memb.Memb_FirstName AS [Learner], Memb.Memb_Id AS [Id], Memb.Memb_Criteria AS [Criteria], Left(Logs.Logs_Item, " & vLogItemLength & ") AS MODULE, SUM(CONVERT(integer, RIGHT(Logs_Item, 6))) AS TIMESPENT, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Level "
  vSql = vSQL & " FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
  vSql = vSQL & " WHERE (Memb.Memb_AcctId = '" & svCustAcctId & "') "
  vSql = vSQL & " AND (Logs.Logs_Type = 'P') "
  If vLearners = "Y" Then
	  vSql = vSQL & " AND (Memb.Memb_Level < 3)"
  Else
	  vSql = vSQL & " AND (Memb.Memb_Level < 4)"
  End If
  If Len(vStrDate) > 0 Then    
    vSql = vSql & " AND (Logs_Posted >= '" & vStrDate & "')"
  End If
  If Len(vEndDate) > 0 Then    
    vSql = vSql & " AND (Logs_Posted <= '" & vEndDate & "')"
  End If
  vSql = vSQL & " GROUP BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_LastName + ', ' + Memb.Memb_FirstName, Memb.Memb_Id, Memb.Memb_Criteria, LEFT(Logs.Logs_Item, " & vLogItemLength & "), Memb.Memb_Level "
  vSql = vSQL & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_ID "
' sDebug

  sOpenDb
  Set oRs = oDb.Execute(vSql)

  sExcelInit    '...initialize 

  Do While Not oRS.eof
    '...ensure you can only see members with same criteria
    If svMembLevel > 2 And svMembCriteria <> "0" And oRs("Criteria") <> svMembCriteria Then 
      vOk = False
    Else
      vId         = Trim(oRs("Id"))
      vModule     = oRs("Module")
      vTimeSpent  = oRs("TimeSpent")
      If vLevel   = "prog" Then
        vTitle    = fProgTitle (Left(vModule, 7))
      Else
        vTitle    = fModsTitle (Right(vModule, 6))
      End If
      vLearner    = Trim(oRs("Learner"))
      If vLearner = "," Then vLearner = ""
      sExcelRow  '...write out worksheet line
    End If        
    oRs.MoveNext	        
  Loop
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
    oCell(vRow, 01) = "<!--{{-->Name<!--}}-->"															: oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 32
    oCell(vRow, 02) = fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")	: oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 20
    oCell(vRow, 03) = "<!--{{-->Time<!--}}-->"		 											              : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 10
    oCell(vRow, 04) = vLogItemTitle                                                                    : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 64
  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 02).Style = oStyleL
    oCell(vRow, 02).Style = oStyleI

    oCell(vRow, 01) = vLearner
    oCell(vRow, 02) = fId(vId)
    oCell(vRow, 03) = vTimeSpent
    oCell(vRow, 04) = vModule & " - " & vTitle
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Timespent Report dated " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub


  Function fId (vId)
    fId = fIf(oRs("Memb_Level") > 2, "******", fDefault(vId, "N/A"))
  End Function

%>