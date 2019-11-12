<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<%
  Dim vId, vLevel, vIdLast, vCriteria, vName, vNamePrev, vModule, vModules, vModulePrev, vRows, vStrDate, vEndDate, vOk
  
  vStrDate  = Request("vStrDate")
  vEndDate  = Request("vEndDate")


  '........................................................................................
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 

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
  '........................................................................................



  vRow = 1
  oCell(vRow, 1) = Session("CustTitle") & " - " & "Modules accessed between " &  fIf(Len(vStrDate)>0, vStrDate, "Jan 1st of last year ") & " and " & fIf(Len(vEndDate)>0, vEndDate, fFormatDate(Now))
	oCell(vRow, 1).Format.Font.Bold = True   : oCell.ColumnWidth(1) = 30

  vRow = 3
  oCell(vRow, 1) = "Group"
  oCell(vRow, 2) = "Name"
  oCell(vRow, 3) = fIf(svCustPwd, "Learner Id", "Password")
  oCell(vRow, 4) = "Module accessed"

	oCell(vRow, 1).Format.Font.Bold = True   : oCell.ColumnWidth(1) = 24
	oCell(vRow, 2).Format.Font.Bold = True   : oCell.ColumnWidth(2) = 24
	oCell(vRow, 3).Format.Font.Bold = True   : oCell.ColumnWidth(3) = 24
	oCell(vRow, 4).Format.Font.Bold = True   : oCell.ColumnWidth(4) = 10

  vRow = 4
  
  vSql = "SELECT Memb.Memb_FirstName + ' ' + Memb.Memb_LastName AS [Name], Memb.Memb_Id AS [Id], Memb.Memb_Criteria AS [Criteria], SUBSTRING(Logs.Logs_Item, 9, 6) AS MODULE, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Level AS [Level]"
  vSql = vSQL & " FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
  vSql = vSQL & " WHERE (Logs.Logs_AcctId = '" & svCustAcctId & "') AND (Logs.Logs_Type = 'P') AND (Memb.Memb_Level < 5)"
  If Len(vStrDate) > 0 Then    
    vSql = vSql & " AND (Logs_Posted >= '" & vStrDate & "')"
  End If
  If Len(vEndDate) > 0 Then    
    vSql = vSql & " AND (Logs_Posted <= '" & vEndDate & "')"
  End If
  vSql = vSQL & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName "

' sDebug
  sOpenDb
  Set oRs = oDb.Execute(vSql)

  Do While Not oRs.Eof

    '...ensure you can only see members with same criteria
    If Not (svMembLevel > 2 And svMembCriteria <> "0" And oRs("Criteria") <> svMembCriteria) Then 

      vId      = Trim(oRs("Id"))
      vModule  = oRs("Module")
      vLevel   = oRs("Level")
      vName    = fDefault(Trim(oRs("Name")), "N/A")
      vCriteria = Trim(oRs("Criteria"))
      vCriteria = fIf(vCriteria="0", "", fCriteria(vCriteria))      

      vRow = vRow + 1
      oCell(vRow, 03).Style = oStyleL
      oCell(vRow, 03).Style = oStyleI

      oCell(vRow, 01) = vCriteria
      oCell(vRow, 02) = vName
      oCell(vRow, 03) = fId(vId, vLevel)
      oCell(vRow, 04) = vModule      

    End If

    oRs.MoveNext	  
  Loop
  sCloseDB


  '...output spreadsheet if there are any rows
  Response.ContentType = "application/vnd.ms-excel"
  oWs.Save "Module Access Summary Report dated " & fFormatDate(Now) & ".xls", 1
  Response.End


  Function fId (vId, vLevel)
    fId = fIf(vLevel > 2, "******", fDefault(vId, "N/A"))
  End Function

%>

