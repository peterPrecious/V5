<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleI, oStyleL, oStyleR, oStyleC, vRow, vCol
  
  Dim vActive, vPassword, vDate, vCompleted
  Dim vGroupId, vMembNo, vMembId, vMembActive, vMembLevel, vFirstName, vLastName, vMembMemo, vProgNo, vProgId, vProgTitle, vModsNo, vModsId, vModsTitle, vTimeSpent, vScore, vLastDate, vExpired, vRecurs, vSessionId, vObjectiveId
  
	Function fExpired(vExpired, vRecurs)
	  If vRecurs = 0 Then
	    fExpired = "na"
	  ElseIf IsDate(vExpired) Then  
	    fExpired= fYN (True)
	  Else
	    fExpired= fYN (False)
	  End If
	End Function

  '...initialize spreadsheet
  sExcelInit

	vSql = "SELECT DISTINCT * FROM LogsR WHERE UserNo = " & svMembNo
	sOpenDb
	Set oRs = oDb.Execute(vSql)

  Do While Not oRs.Eof
	
	  vGroupId       = oRs("GroupId")
	  vMembNo        = oRs("MembNo")
	  vMembId        = oRs("MembId")
	  vMembMemo      = oRs("MembMemo")
	  vMembActive		 = fYN(oRs("MembActive"))
	  vMembLevel 		 = oRs("MembLevel")
	  vFirstName     = oRs("FirstName")
	  vLastName      = oRs("LastName")
	
	  vProgNo        = oRs("ProgNo")
	  vProgId        = oRs("ProgId")
	  vProgTitle     = oRs("ProgTitle")
	  vModsNo        = oRs("ModsNo")
	  vModsId        = oRs("ModsId")
	  vModsTitle     = oRs("ModsTitle")
	
	  vTimeSpent     = oRs("TimeSpent")
	  vScore    		 = oRs("Score")
	  vLastDate      = oRs("LastDate")
	  vCompleted     = fYN(oRs("Completed"))
	  vExpired       = fExpired(oRs("Expired"), oRs("Recurs"))

    vPassword      = fIf(svMembLevel > vMembLevel, vMembId, "***")

    sExcelDetails      

    oRs.MoveNext
  Loop 
  Set oRs = Nothing
  sCloseDb


  '...close the worksheet 
  sExcelClose
	
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell                    = oWs.Worksheets(1).Cells

    Set oStyleD      	 	  	  	 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleC      	 	  	  	 = oWs.CreateStyle
  
    oStyleD.Number      				 = 14    '...format date m/d/yy
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234

    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleC.HorizontalAlignment  = 2     '...center justify

 
    vRow = 1

    oCell.RowHeight(vRow) = 20

    oCell(vRow, 05).Style = oStyleC
    oCell(vRow, 06).Style = oStyleC
    oCell(vRow, 08).Style = oStyleC

    oCell(vRow, 10).Style = oStyleC
    oCell(vRow, 11).Style = oStyleC
    oCell(vRow, 12).Style = oStyleC
    oCell(vRow, 13).Style = oStyleC
    oCell(vRow, 14).Style = oStyleC

    oCell(vRow, 01) = "Group"             											: oCell(vRow, 01).Format.Font.Bold = True :	oCell.ColumnWidth(01) = 18
    oCell(vRow, 02) = "First Name"        											: oCell(vRow, 02).Format.Font.Bold = True :	oCell.ColumnWidth(02) = 12
    oCell(vRow, 03) = "Last Name"         											: oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 12
    oCell(vRow, 04) = fIf(svCustPwd, "Learner Id", "Password")  : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 12                
    oCell(vRow, 05) = "Active"            											: oCell(vRow, 05).Format.Font.Bold = True :	oCell.ColumnWidth(05) = 08
    oCell(vRow, 06) = "Program"            											: oCell(vRow, 06).Format.Font.Bold = True :	oCell.ColumnWidth(06) = 08
    oCell(vRow, 07) = "Title"             											: oCell(vRow, 07).Format.Font.Bold = True :	oCell.ColumnWidth(07) = 48
    oCell(vRow, 08) = "Module" 						        							: oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 08
    oCell(vRow, 09) = "Title"             											: oCell(vRow, 09).Format.Font.Bold = True :	oCell.ColumnWidth(09) = 48
    oCell(vRow, 10) = "TimeSpent"        											  : oCell(vRow, 10).Format.Font.Bold = True : oCell.ColumnWidth(10) = 12   
    oCell(vRow, 11) = "Score"        											      : oCell(vRow, 11).Format.Font.Bold = True : oCell.ColumnWidth(11) = 12
    oCell(vRow, 12) = "Date"            											  : oCell(vRow, 12).Format.Font.Bold = True : oCell.ColumnWidth(12) = 12
    oCell(vRow, 13) = "Completed"              									: oCell(vRow, 13).Format.Font.Bold = True :	oCell.ColumnWidth(13) = 12
    oCell(vRow, 14) = "Closed"              										: oCell(vRow, 14).Format.Font.Bold = True :	oCell.ColumnWidth(14) = 12

  End Sub




  '...write out details
  Sub sExcelDetails
    vRow = vRow + 1 
    
    oCell(vRow, 05).Style = oStyleC
    oCell(vRow, 06).Style = oStyleC
    oCell(vRow, 08).Style = oStyleC

    oCell(vRow, 10).Style = oStyleC
    oCell(vRow, 11).Style = oStyleC
    oCell(vRow, 12).Style = oStyleC
    oCell(vRow, 13).Style = oStyleC
    oCell(vRow, 14).Style = oStyleC

    oCell(vRow, 12).Style = oStyleD

    
    oCell(vRow, 01) = vGroupId
    oCell(vRow, 02) = vFirstName
    oCell(vRow, 03) = vLastName
    oCell(vRow, 04) = vPassword
    oCell(vRow, 05) = vMembActive
    oCell(vRow, 06) = vProgId
    oCell(vRow, 07) = vProgTitle
    oCell(vRow, 08) = vModsId
    oCell(vRow, 09) = vModsTitle

    oCell(vRow, 10) = vTimeSpent
    oCell(vRow, 11) = vScore
    oCell(vRow, 12) = vLastDate
    oCell(vRow, 13) = vCompleted
    oCell(vRow, 14) = vExpired

  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Dim vTitle
    vTitle = "<!--{{-->Learner Report Card<!--}}-->"
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save vTitle & " - " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub

%>