<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 

  Dim vActive, vGlobal, vNext, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vFormat, vLearners
  Dim vLastValue, vDetails, vCurList, vRecnt, vWhere, aCrit, vEdit, vExtend, vRole

  vNext            = Request("vNext")
  vActive          = fDefault(Request("vActive"), "1")
  vGlobal          = fDefault(Request("vGlobal"), "0")
  vFind            = fDefault(Request("vFind"), "S")
  vFindId          = fUnQuote(Request("vFindId"))
  vFindFirstName   = fUnQuote(Request("vFindFirstName"))
  vFindLastName    = fUnQuote(Request("vFindLastName"))
  vFindEmail       = fNoQuote(Request("vFindEmail"))
  vFindCriteria    = Request("vFindCriteria")
  vFormat          = fDefault(Request("vFormat"), "o")
  vLearners        = fDefault(Request("vLearners"), "n")

  vDetails         = Request("vDetails") 
  vLastValue       = Request("vLastValue") 
  vCurList         = fDefault(Request("vCurList"), 0)

  '...initialize 
  sExcelInit

  vSql = " SELECT * FROM "_
        & "   Memb WITH (NOLOCK) INNER JOIN "_ 
        & "   Crit WITH (NOLOCK) ON TRY_CAST(Memb.Memb_Criteria AS INT) = Crit.Crit_No "_      
        & " WHERE "_
        & "   (Memb_AcctId = '" & svCustAcctId & "')"_ 
        & "   AND (Memb_Level <= " & Session("Completion_Level") & ")"_ 
        & "   AND (Memb_Internal = 0)"_ 
        & "   AND (ISNUMERIC(Memb_Criteria) = 1)"_ 
        &     fIf (Len(vLastValue)    > 0, " AND (ISNULL(Memb_LastName,'') + ISNULL(Memb_FirstName,'') + CAST(Memb_No AS VARCHAR(10)) >= '" & fUnquote(vLastValue) & "')","") _
        &     fIf (Len(vFindId)       > 0, " AND (Memb_Id       LIKE '" & vFindId         & "%')", "") _
        &     fIf (Len(vFindLastName) > 0, " AND (Memb_LastName LIKE '" & vFindLastName   & "%')", "") _   
        &     fIf (vFindCriteria <> "All", " AND (Crit_No IN (" & vFindCriteria & "))", "") _
        &     fIf (vActive <> "*", " AND (Memb_Active = " & vActive & ")", "") _
        & " ORDER BY "_
        & "   Crit_Id, ISNULL(Memb_LastName,'') + ISNULL(Memb_FirstName,'') + CAST(Memb_No AS varchar(10))"

  sCompletion_Debug 
  sOpenDb
  Set oRs = oDb.Execute(vSql)

  Do While Not oRs.Eof
    sReadMemb
    vCrit_Id = oRs("Crit_Id")
    vCurList = vCurList + 1 
    '...write out worksheet line
    sExcelRow  
    oRs.MoveNext
  Loop
  Set oRs = Nothing
  sCloseDb 
 
  '...close the worksheet 
  sExcelClose



  '...call this one time when ready to setup the worksheet
  Sub sExcelInit
    Set oWs         	 				   = Server.CreateObject("SoftArtisans.ExcelWriter")
    Set oCell       	   				 = oWs.Worksheets(1).Cells

    Set oStyleD      	 	  	  	 = oWs.CreateStyle
    Set oStyleR      	 		  		 = oWs.CreateStyle
    Set oStyleL      	 		  		 = oWs.CreateStyle
    Set oStyleI      	 		  		 = oWs.CreateStyle

    oStyleD.Number      				 = 14    '...format date m/d/yy
'   oStyleD.Number      				 = "mmm dd, yyyy" '...this does not seem to work 
    oStyleR.HorizontalAlignment  = 3     '...right justify
    oStyleL.HorizontalAlignment  = 1     '...left justify (numbers)
    oStyleI.Number      				 = 49    '...consider as text, ie leave as 01234

    vRow = 1
    vCol = 0
    oCell.RowHeight(vRow) = 50

'   oCell(vRow, 08).Style = oStyleR

    vCol = vCol + 1 : oCell(vRow, vCol) = "Region"        : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Location"      : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 12
    vCol = vCol + 1 : oCell(vRow, vCol) = "Role"          : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 06
    vCol = vCol + 1 : oCell(vRow, vCol) = "Password"      : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "First Name"    : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 32
    vCol = vCol + 1 : oCell(vRow, vCol) = "Last Name"     : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 22
    vCol = vCol + 1 : oCell(vRow, vCol) = "Active"        : oCell(vRow, vCol).Format.Font.Bold = True : oCell.ColumnWidth(vCol) = 06
  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1
    vCol = 0

    oCell(vRow, 01).Style = oStyleL
    oCell(vRow, 02).Style = oStyleL
    oCell(vRow, 02).Style = oStyleI

    vCol = vCol + 1 : oCell(vRow, vCol) = Left(vCrit_Id, Session("Completion_L1len"))
    vCol = vCol + 1 : oCell(vRow, vCol) = Mid(vCrit_Id, Session("Completion_L0len"), Session("Completion_L0str"))
    vCol = vCol + 1 : oCell(vRow, vCol) = Right(vCrit_Id, Session("Completion_RLlen"))
    vCol = vCol + 1 : oCell(vRow, vCol) = vMemb_Id
    vCol = vCol + 1 : oCell(vRow, vCol) = vMemb_FirstName
    vCol = vCol + 1 : oCell(vRow, vCol) = vMemb_LastName
    vCol = vCol + 1 : oCell(vRow, vCol) = fYN(vMemb_Active)
  End Sub

 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Learner Report dated" & " " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
%>

