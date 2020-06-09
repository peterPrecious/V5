<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Parm.asp"-->

<% 
  '........................................................................................
  Dim oWs, oCell, oStyleD, oStyleC, oStyleR, oStyleL, oStyleI, vRow, vCol 

  Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
  Set oCell                    = oWs.Worksheets(1).Cells

  Set oStyleD      	 	  	  	 = oWs.CreateStyle
  Set oStyleR      	 		  		 = oWs.CreateStyle
  Set oStyleC      	 		  		 = oWs.CreateStyle
  Set oStyleL      	 		  		 = oWs.CreateStyle
  Set oStyleI      	 		  		 = oWs.CreateStyle

  oStyleD.Number      				 = 14    '...format date m/d/yy
  oStyleR.HorizontalAlignment  = 3     '...right justify
  oStyleC.HorizontalAlignment  = 2     '...center align
  oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
  oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234

  Server.ScriptTimeout = 60 * 10
  
  Dim vCriteria, vDays, vLearners, vDateFormat, vSelect, vPrevious, vCnt, vParmNo
  Dim vDate, vType, aCrit, bOk

  vCriteria   = Replace(fDefault(Request("vCriteria"), "0"), " ", "")
  vDays       = fDefault(Request("vDays"), "30")
  vLearners   = fDefault(Trim(Request("vLearners")), "2")
  vDateFormat = fDefault(Request("vDateFormat"), "X")
  vSelect     = Ucase(Trim(Request("vSelect")))
  vPrevious   = Request("vPrevious")
  vCnt        = Clng(fDefault(Request("vCnt"), 0))
  vParmNo     = Request("vParmNo")

  '...initialize 
  sExcelInit

  vSql = "SELECT "_
  	   & "  Memb.Memb_Id, "_
  	   & "  Memb.Memb_LastName, "_
  	   & "  Memb.Memb_FirstName, "_
  	   & "  Memb.Memb_Criteria, "_
  	   & "  Memb.Memb_Level, "_
  	   & "  Logs.Logs_Posted as Posted, "_ 
  	   & "  CASE Logs_Type WHEN 'S' THEN Logs.Logs_Item WHEN 'L' THEN LEFT(Logs.Logs_Item, 6) END AS Id "_ 

  	   & "FROM "_
  	   & "  Memb WITH (nolock) "_
  	   & "  INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo "_

  	   & "WHERE "_

  	   & "  (Logs.Logs_AcctId = '" & svCustAcctId & "') "_
  	   & "  AND (Memb.Memb_LastName + Memb.Memb_FirstName + Memb.Memb_Id >= '" & vPrevious & "') "_
  	   & "  AND (Memb.Memb_Active = 1) "_
  	   & "  AND (Memb.Memb_Level IN (" & vLearners & ")) "_
       &    fIf(vDays = 0, "", "AND (Logs.Logs_Posted >='" & fFormatSqlDate(DateAdd("d", Now(), -vDays)) & "') ") _
  	   & "  AND (Logs.Logs_Type = 'S' AND LEN(Logs.Logs_Item) = 6) "_
       &    fParmValue (vParmNo) _
  	   
  	   & "OR "_
  	   & "  (Logs.Logs_AcctId = '" & svCustAcctId & "') "_
  	   & "  AND (Memb.Memb_LastName + Memb.Memb_FirstName + Memb.Memb_Id >= '" & vPrevious & "') "_
  	   & "  AND (Memb.Memb_Active = 1) "_
  	   & "  AND (Memb.Memb_Level IN (" & vLearners & ")) "_
       &    fIf(vDays = 0, "", "AND (Logs.Logs_Posted >='" & fFormatSqlDate(DateAdd("d", Now(), -vDays)) & "') ") _
  	   & "  AND (Logs.Logs_Type = 'L' AND Logs.Logs_Item LIKE '%_completed') "_
       &    fParmValue (vParmNo) _

  	   & "ORDER BY "_
  	   & "  Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Id "

' sDebug 
' stop

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof

    bOk = False : If vSelect = "" Or Instr(vSelect, oRs("Id")) > 0 Then bOk = True

    '...criteria can be 129,330 or just 129
    If bOk Then
      If Len(vCriteria) > 3 Then          
        bOk = False
        aCrit = Split(vCriteria, ",")
        For i = 0 To Ubound(aCrit)
          If Instr(oRs("Memb_Criteria"), aCrit(i)) > 0 Then
            bOk = True
            Exit For
          End If
        Next    
      End If   
    End If   
  
    If bOk Then
      vDate = oRs("Posted")
      If vDateFormat = "S" Then 
        vDate = fFormatDate(vDate)
      Else
        vDate = Right("00" & Month(vDate), 2) & "/" & Right("00" & Day(vDate), 2) & "/" & Year(vDate)
      End If 
      sExcelRow 
    End If

    oRs.MoveNext	        
  Loop
  sCloseDb
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    vRow = 1
    vCol = 0

    vCol = vCol + 1 : oCell(vRow, vCol) = "Group" 	
                      oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 18
    vCol = vCol + 1 : oCell(vRow, vCol) = "First Name"     
                      oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 18
    vCol = vCol + 1 : oCell(vRow, vCol) = "Last Name"    
                      oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 18
    vCol = vCol + 1 : oCell(vRow, vCol) = fIf(svCustPwd, "Learner Id", "Password")
                      oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 24
    vCol = vCol + 1 : oCell(vRow, vCol) = "Date Completed"    
                      oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 18
    vCol = vCol + 1 : oCell(vRow, vCol) = "Module"    
                      oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 48
  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0

    oCell(vRow, 01).Style = oStyleL
    oCell(vRow, 01).Style = oStyleI

    vCol = vCol + 1 : oCell(vRow, vCol) = fIf(Len(Trim(oRs("Memb_Criteria"))) < 3 Or Trim(oRs("Memb_Criteria")) = "0" , "", Replace(fCriteria(oRs("Memb_Criteria")), " + ", "<br>")) 
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Memb_FirstName").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Memb_LastName").Value     
                      oCell(vRow, vCol).Style = oStyleL
    vCol = vCol + 1 : oCell(vRow, vCol) = fIf(oRs("Memb_Level")=5, "********", oRs("Memb_Id"))     
                      oCell(vRow, vCol).Style = oStyleL
    vCol = vCol + 1 : oCell(vRow, vCol) = vDate
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("Id") & " - " & fLeft(fModsTitle(oRs("Id")), 40)
  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Customer Sellers|Owners Report as of " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

  
%>

