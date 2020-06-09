<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, vRow, vCol, oStyleC, oStyleR, oStyleL, oStyleI, oStyleD
  Dim vStrDate, vEndDate, vPrograms, vPassword

  vStrDate  = Request("vStrDate")
  vEndDate  = Request("vEndDate") 
  vPrograms = Request("vPrograms")
  vPassword = Request("vPassword")

  sExcelInit '...initialize 

  Function fProgs(vPrograms)
    fProgs = ""
    If Len(Trim(vPrograms)) > 14 Then
      fProgs = "(Pr.Prog_Id IN ('" & Replace(vPrograms, " ", "', '") & "')) AND "
    Elseif Len(Trim(vPrograms)) = 7 Then
      fProgs = "(Pr.Prog_Id = '" & vPrograms & "') AND "
    End If        
  End Function

  vSql = "SELECT TOP 50000" _     
       & "  Me.Memb_FirstName + ' ' + Me.Memb_LastName AS Learner, " _ 
       & "  Me.Memb_Id AS Password, " _ 
       & "  Ec.Ecom_Programs AS Program, " _ 
       & "  Pr.Prog_Title1 AS Title,  " _
       & "  Ec.Ecom_Issued AS Purchased, " _ 
       & "  Ec.Ecom_Expires AS Expired, " _ 
       & "  Sc.pcnCompleted AS Completed " _
       & "FROM " _         
       & "  V5_Vubz.dbo.Memb                          AS Me INNER JOIN " _
       & "  V5_Vubz.dbo.Ecom                          AS Ec ON Me.Memb_No = Ec.Ecom_MembNo INNER JOIN " _
       & "  V5_Base.dbo.Prog                          AS Pr ON Ec.Ecom_Programs = Pr.Prog_Id LEFT OUTER JOIN " _
       & "  vuGoldSCORM.dbo.LearnerProgramCompleted   AS Sc ON Me.Memb_No = Sc.pcnMembID AND Pr.Prog_No = Sc.pcnProgramID " _
       & "WHERE "_
       & "  (Ec.Ecom_Media = 'Online') AND "_
       &    fProgs(vPrograms) _
       & "  (Ec.Ecom_Issued BETWEEN '" & vStrDate & "' AND '" & vEndDate & "') AND "_
       & "  (Me.Memb_AcctId = '" & svCustAcctId & "') " _
       &    fIf(vPassword = "", "", " AND Ec.Ecom_Id = '" & vPassword & "'") _
       & "ORDER BY "_
       & "  Me.Memb_LastName, Me.Memb_FirstName, Purchased "

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRs.Eof
    sExcelRow
    oRs.MoveNext	        
  Loop
  sCloseDb  
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs                      = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell                    = oWs.Worksheets(1).Cells

    Set oStyleC      	 		  		 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle
    Set oStyleD      	 	  	  	 = oWs.CreateStyle

    oStyleC.HorizontalAlignment  = 2     '...center align
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234
    oStyleD.Number      				 = 14    '...format date m/d/yy

    vRow = 1
    oCell.RowHeight(vRow) = 50

    oCell(vRow, 03).Style = oStyleC

    oCell(vRow, 05).Style = oStyleR
    oCell(vRow, 06).Style = oStyleR
    oCell(vRow, 07).Style = oStyleR





    oCell(vRow, 01) = "Learner"     : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 24
    oCell(vRow, 02) = fIf(svCustPwd, "Id", "Password")    : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 24
    oCell(vRow, 03) = "Program "    : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 12
    oCell(vRow, 04) = "Title"       : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 48
    oCell(vRow, 05) = "Purchased"   : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 12
    oCell(vRow, 06) = "Expired"     : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 12
    oCell(vRow, 07) = "Completed"   : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 12
  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 02).Style = oStyleL
    oCell(vRow, 03).Style = oStyleC
    oCell(vRow, 05).Style = oStyleR
    oCell(vRow, 06).Style = oStyleR
    oCell(vRow, 07).Style = oStyleR
    oCell(vRow, 05).Style = oStyleD
    oCell(vRow, 06).Style = oStyleD
    oCell(vRow, 07).Style = oStyleD

    oCell(vRow, 01) = oRs("Learner").Value
    oCell(vRow, 02) = oRs("Password").Value
    oCell(vRow, 03) = oRs("Program").Value
    oCell(vRow, 04) = oRs("Title").Value
    oCell(vRow, 05) = oRs("Purchased").Value
    oCell(vRow, 06) = oRs("Expired").Value
    oCell(vRow, 07) = oRs("Completed").Value
  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Ecommerce Completion Report (Basic) as " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
  
  

  
  
%>