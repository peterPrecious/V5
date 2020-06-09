<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->



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
  oStyleD.HorizontalAlignment  = 2     '...center align
  oStyleR.HorizontalAlignment  = 3     '...right justify
  oStyleC.HorizontalAlignment  = 2     '...center align
  oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
  oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234
  '........................................................................................


  '...initialize 
  sExcelInit

  vSql = " " _
        & "SELECT TOP 5000 "_
        & "	Cust_Id AS CustId, "_ 
        & "  Cust_Title AS CustTitle, "_ 
        & "	Count(Memb_Id) AS MembCount, "_ 
        & "  Cust_Expires AS CustExpires "_
        & "FROM "_ 
        & "  Cust INNER JOIN "_ 
        & "	Memb ON Cust_AcctId = Memb_AcctId "_ 
        & "WHERE "_ 
        & "	(Cust_ParentId = '" & RIGHT(svCustId, 4) & "') AND "_ 
        & "	(Cust_Active = 1) AND "_ 
        & "	(Memb_Level < 4) AND "_   
        & "	(Cust_Expires > getDate()) AND "_
        & "	(Cust_Expires < DATEADD(year, 1, getDate())) "_
        & "GROUP BY "_ 
        & "  Cust_Id, "_
        & "  Cust_Title, "_ 
        & "  Cust_Expires "_
        & "ORDER BY  "_
        & "  Cust_Expires "

  vSql = " " _
        & "SELECT TOP 5000 "_
	      & "  cu.Cust_Id AS custId, "_
	      & "  cu.Cust_Title AS custTitle, "_
	      & "  Count(m1.Memb_Id) AS membCount, "_
	      & "  Cust_Expires AS custExpires, "_
	      & "  m2.Memb_Id AS facId, "_
        & "	m2.Memb_Email AS facEmail, "_
        & "	m2.Memb_LastVisit AS facLast "_
        & "FROM "_  
        & "	V5_Vubz.dbo.Cust cu											                  INNER JOIN "_
        & "	V5_Vubz.dbo.Memb m1 ON cu.Cust_AcctId = m1.Memb_AcctId		INNER JOIN "_
        & "	V5_Vubz.dbo.Memb m2 ON cu.Cust_AcctId = m2.Memb_AcctId "_
        & "WHERE "_  
        & "	(Cust_ParentId = '" & RIGHT(svCustId, 4) & "') AND "_ 
        & "	(cu.Cust_Active = 1) AND "_  
        & "	(m1.Memb_Level < 4) AND "_    
        & "	(m1.Memb_Internal = 0) AND "_
        & "	(cu.Cust_Expires > getDate()) AND "_ 
        & "	(cu.Cust_Expires < DATEADD(year, 1, getDate())) AND "_
        & "	(m2.Memb_Level = 3) AND "_
        & "	(m2.Memb_Internal = 0) AND "_
        & "	(m2.Memb_LastVisit is not null) "_ 
        & "GROUP BY "_  
        & "	cu.Cust_Id, "_ 
        & "	cu.Cust_Title, "_  
        & "	cu.Cust_Expires, "_
        & "	m2.Memb_Id, "_
        & "	m2.Memb_Email, "_
        & "	m2.Memb_LastVisit "_
        & "ORDER BY "_  
        & "	cu.Cust_Expires " 


'     Response.Write vSql
  sOpenDb      
  Set oRs = oDb.Execute(vSql)      

  Do While Not oRs.Eof
    '...write out worksheet line
    sExcelRow 
     
    Set oRs2 = Nothing
    oRs.MoveNext	        
  Loop
  sCloseDb


  '...close the worksheet 
  sExcelClose
  
  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    vRow = 1
    vCol = 0

    vCol = vCol + 1 : oCell(vRow, vCol) = "Customer Id" 	: oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Title"         : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 48
    vCol = vCol + 1 : oCell(vRow, vCol) = "# Learners"    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12 : oCell(vRow, vCol).Style = oStyleC : oCell(vRow, vCol).Format.Font.Bold = True 
    vCol = vCol + 1 : oCell(vRow, vCol) = "Expiry Date"   :                                             oCell.ColumnWidth(vCol) = 12 : oCell(vRow, vCol).Style = oStyleC : oCell(vRow, vCol).Format.Font.Bold = True 

    vCol = vCol + 1 : oCell(vRow, vCol) = "Facilitator"   :                                             oCell.ColumnWidth(vCol) = 22 
    vCol = vCol + 1 : oCell(vRow, vCol) = "Email"         :                                             oCell.ColumnWidth(vCol) = 22 
    vCol = vCol + 1 : oCell(vRow, vCol) = "Last Visit"    :                                             oCell.ColumnWidth(vCol) = 22 : oCell(vRow, vCol).Style = oStyleC : oCell(vRow, vCol).Format.Font.Bold = True 

    vRow = 2

  End Sub

 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0

    oCell(vRow, 9).Style = oStyleD

    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("custId").Value                                                               : oCell(vRow, 01).Style = oStyleL : oCell(vRow, 01).Style = oStyleI
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("custTitle").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("membCount").Value                                                                                                 
    vCol = vCol + 1 : oCell(vRow, vCol) = fFormatDate(oRs("custExpires").Value)                                                                                                

    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("facId").Value
    vCol = vCol + 1 : oCell(vRow, vCol) = oRs("facEmail").Value                                                                                                
    vCol = vCol + 1 : oCell(vRow, vCol) = fFormatDate(oRs("facLast").Value)                                                                                                


  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Channel Expiry Report as of " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
  
%>

