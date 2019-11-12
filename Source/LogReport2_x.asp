<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<%
  Dim vId, vIdPrev, vIdLast, vCriteria, vName, vNamePrev, vModule, vModules, vModulePrev, vRows, vStrDate, vEndDate, vTitle, vOk, vLevel

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
  oCell(vRow, 4) = "Modules accessed"

	oCell(vRow, 1).Format.Font.Bold = True   : oCell.ColumnWidth(1) = 24
	oCell(vRow, 2).Format.Font.Bold = True   : oCell.ColumnWidth(2) = 24
	oCell(vRow, 3).Format.Font.Bold = True   : oCell.ColumnWidth(3) = 24
	oCell(vRow, 4).Format.Font.Bold = True   : oCell.ColumnWidth(4) = 80


  vRow = 4
    
  vIdprev = "": vModules = "": vIdLast = ""
 
  vSql = "SELECT Memb.Memb_LastName + ',  ' + Memb.Memb_FirstName AS Name, Memb.Memb_Id AS [Id], Memb.Memb_Criteria AS [Criteria], SUBSTRING(Logs.Logs_Item, 9, 6) AS MODULE, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Level AS [Level]"
  vSql = vSQL & " FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
  vSql = vSQL & " WHERE (Logs.Logs_AcctID = '" & svCustAcctId & "') AND (Logs.Logs_Type = 'P') AND (Memb.Memb_Level < 4)"
  If Len(vStrDate) > 0 Then    
    vSql = vSql & " AND (Logs_Posted >= '" & vStrDate & "')"
  End If
  If Len(vEndDate) > 0 Then    
    vSql = vSql & " AND (Logs_Posted <= '" & vEndDate & "')"
  End If
  vSql = vSQL & " GROUP BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_LastName + ', ' + Memb.Memb_FirstName, Memb.Memb_Id, Memb.Memb_Level, SUBSTRING(Logs.Logs_Item, 9, 6) "
  vSql = vSQL & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_ID "

  sOpenDb
  Set oRs = oDb.Execute(vSql)

  Do While Not oRs.eof

    '...ensure you can only see members with same criteria
    If svMembLevel > 2 And svMembCriteria <> "0" And oRs("Criteria") <> svMembCriteria Then 
      vOk = False
    Else
      vOk = True
    End If

    If vOk Then
      vId       = Trim(oRs("Id"))
      vLevel    = oRs("Level")
      vName     = Trim(oRs("Name"))
      vCriteria = Trim(oRs("Criteria"))
      vCriteria = fif(vCriteria="0", "", fCriteria(vCriteria))
      If vName  = "," Then vName = ""    
      vModule   = oRs("Module")
      vTitle    = fModsTitle (vModule)
      vRow = vRow + 1
      
      'sDebug vName, vModule
      
      oCell(vRow, 1) = vCriteria
      oCell(vRow, 2) = vName
      oCell(vRow, 3) = fId(vId, vLevel)
      oCell(vRow, 4) = vModule & " - " & vTitle    
    End If
            
    oRs.MoveNext	        
  Loop
  sCloseDB

  '...output spreadsheet if there are any rows
  Response.ContentType = "application/vnd.ms-excel"
  oWs.Save "Module Access Details Report dated " & fFormatDate(Now) & ".xls", 1
  Response.End

  Function fId (vId, vLevel)
    fId = fIf(vLevel > 2, "******", fDefault(vId, "N/A"))
  End Function

  
%>