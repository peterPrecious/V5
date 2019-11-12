<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  '...Excel variables
  Dim oWs, oCell, oStyleD, oStyleR, oStyleL, oStyleI, vRow, vCol 
  Dim vActive, vGlobal, vCustId, vNext, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vLearners, vLevel
  Dim vLastId, vDetails, vWhere, aCrit, aMemo

  vNext            = Request("vNext")
  vCustId          = fDefault(Request("vCustId"), svCustId)
  vActive          = fDefault(Request("vActive"), "1")
  vGlobal          = fDefault(Request("vGlobal"), "0")
  vFind            = fDefault(Request("vFind"), "S")
  vFindId          = fUnQuote(Request("vFindId"))
  vFindFirstName   = fUnQuote(Request("vFindFirstName"))
  vFindLastName    = fUnQuote(Request("vFindLastName"))
  vFindEmail       = fNoQuote(Request("vFindEmail"))
  vFindMemo        = fUnQuote(Request("vFindMemo"))
  vFindCriteria    = Request("vFindCriteria")
  vFormat          = fDefault(Request("vFormat"), "o")
  vLearners        = Request("vLearners")
  vDetails         = Request("vDetails") 
  
  vWhere = ""
  If svMembLevel < 4 Then vWhere = vWhere & " AND (Memb_Level < 4)"
  If vActive = "0" Then vWhere = vWhere & " AND (Memb_Active = 1)"

  If vFind = "S" Then
    If Len(vFindId)        > 0 Then vWhere = vWhere & " AND (Memb_Id        LIKE '" & vFindId         & "%')"
    If Len(vFindFirstName) > 0 Then vWhere = vWhere & " AND (Memb_FirstName LIKE '" & vFindFirstName  & "%')"
    If Len(vFindLastName)  > 0 Then vWhere = vWhere & " AND (Memb_LastName  LIKE '" & vFindLastName   & "%')"
    If Len(vFindEmail)     > 0 Then vWhere = vWhere & " AND (Memb_Email     LIKE '" & vFindEmail      & "%')"
    If Len(vFindMemo)      > 0 Then vWhere = vWhere & " AND (Memb_Memo      LIKE '" & vFindMemo       & "%')"
  Else
    If Len(vFindId)        > 0 Then vWhere = vWhere & " AND (Memb_Id        LIKE '%" & vFindId        & "%')"
    If Len(vFindFirstName) > 0 Then vWhere = vWhere & " AND (Memb_FirstName LIKE '%" & vFindFirstName & "%')"
    If Len(vFindLastName)  > 0 Then vWhere = vWhere & " AND (Memb_LastName  LIKE '%" & vFindLastName  & "%')"
    If Len(vFindEmail)     > 0 Then vWhere = vWhere & " AND (Memb_Email     LIKE '%" & vFindEmail     & "%')"
    If Len(vFindMemo)      > 0 Then vWhere = vWhere & " AND (Memb_Memo      LIKE '%" & vFindMemo      & "%')"
  End If

  If Len(vFindCriteria)    > 2 Then '...criteria can be 129,330 or just 129
    vWhere = vWhere & " AND ("
    aCrit = Split(vFindCriteria, ",")
    For i = 0 To Ubound(aCrit)
      vWhere = vWhere & fIf(i = 0, "", " OR ") & "Memb_Criteria = '" & aCrit(i)  & "'"
    Next    
    vWhere = vWhere & ")"
  End If 

  vLevel = ""  
  If Instr(vLearners, "s") > 0 Then 
    vLevel = "2"
    vWhere = vWhere & " AND (Memb_Sponsor > 0)"
  End If

  If Instr(vLearners, "2") > 0 Then
    If vLevel = "" Then vLevel = "2" '...if sponsors (above) then we are already grabbing level 2 (learner)
  End If
  If Instr(vLearners, "3") > 0 Then
    If vLevel = "" Then 
      vLevel = "3"
    Else
      vLevel = vLevel & ",3"
    End If   
  End If
  If Instr(vLearners, "4") > 0 Then
    If vLevel = "" Then 
      vLevel = "4"
    Else
      vLevel = vLevel & ",4"
    End If   
  End If
  If Instr(vLearners, "5") > 0 Then
    If vLevel = "" Then 
      vLevel = "5"
    Else
      vLevel = vLevel & ",5"
    End If   
  End If

  vWhere = vWhere & " AND (Memb_Level IN (" & vLevel & "))"

  vWhere = vWhere & " AND (Memb_Id NOT LIKE '" & vPasswordx & "%')"

  sGetCust vCustId

  '...initialize 
  sExcelInit

  sGetMemb_Rs vCust_AcctId, vWhere, vGlobal
  Do While Not oRs.Eof
    sReadMemb
    sExcelRow
    oRs.MoveNext
  Loop
  Set oRs = Nothing
  sCloseDb
  
  '...close the worksheet 
  sExcelClose

  
  '...this prefixes the password with the Customer AcctId 
  Function fGlobal
    If vGlobal = 1 Then
      fGlobal = "(" & vMemb_AcctId & ") "
    Else
      fGlobal = ""
    End If 
  End Function




  Function fId (vId)
    fId = fDefault(vId, "N/A")
    '...ensure you can only see users below your level
    j = ""
    If     vMemb_Level = 3 Then
      j = " * "
    ElseIf vMemb_Level = 4 Then
      j = " ** "
    ElseIf vMemb_Level = 5 Then
      j = " *** "
    End If 
    
    If (svMembLevel > vMemb_Level Or vMemb_No= svMembNo) And InStr(vMemb_Id, vPasswordx) = 0 Then
      fId = fGlobal & vId & j
    Else
      fId = "********"
    End If
  End Function


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
    oCell.RowHeight(vRow) = 50

    oCell(vRow, 08).Style = oStyleR
    oCell(vRow, 09).Style = oStyleR
    oCell(vRow, 10).Style = oStyleR
    oCell(vRow, 11).Style = oStyleR
    oCell(vRow, 12).Style = oStyleR

    oCell(vRow, 01) = "Group"				   : oCell(vRow, 01).Format.Font.Bold = True : oCell.ColumnWidth(01) = 12
    oCell(vRow, 02) = "First Name"	   : oCell(vRow, 02).Format.Font.Bold = True : oCell.ColumnWidth(02) = 16
    oCell(vRow, 03) = "Last Name"		   : oCell(vRow, 03).Format.Font.Bold = True : oCell.ColumnWidth(03) = 16
    oCell(vRow, 04) = "Organization"   : oCell(vRow, 04).Format.Font.Bold = True : oCell.ColumnWidth(04) = 16
    oCell(vRow, 05) = "Email"		       : oCell(vRow, 05).Format.Font.Bold = True : oCell.ColumnWidth(05) = 24
    oCell(vRow, 06) = "Active"         : oCell(vRow, 06).Format.Font.Bold = True : oCell.ColumnWidth(06) = 12
    oCell(vRow, 07) = "Password"       : oCell(vRow, 07).Format.Font.Bold = True : oCell.ColumnWidth(07) = 20
    oCell(vRow, 08) = "First Visit"    : oCell(vRow, 08).Format.Font.Bold = True : oCell.ColumnWidth(08) = 12
    oCell(vRow, 09) = "Last Visit"     : oCell(vRow, 09).Format.Font.Bold = True : oCell.ColumnWidth(09) = 12
    oCell(vRow, 10) = "Expires"        : oCell(vRow, 10).Format.Font.Bold = True : oCell.ColumnWidth(10) = 12
    oCell(vRow, 11) = "No Site Visits" : oCell(vRow, 11).Format.Font.Bold = True : oCell.ColumnWidth(11) = 12
    oCell(vRow, 12) = "Hours Oniste"   : oCell(vRow, 12).Format.Font.Bold = True : oCell.ColumnWidth(12) = 12
    oCell(vRow, 13) = "Memo"           : oCell(vRow, 13).Format.Font.Bold = True : oCell.ColumnWidth(13) = 16
                                                                                   oCell.ColumnWidth(14) = 16
                                                                                   oCell.ColumnWidth(15) = 16
                                                                                   oCell.ColumnWidth(16) = 16
                                                                                   oCell.ColumnWidth(17) = 16

  End Sub


 '...write out a detail line/row
  Sub sExcelRow
    vRow = vRow + 1

    oCell(vRow, 04).Style = oStyleL
    oCell(vRow, 07).Style = oStyleI
    oCell(vRow, 08).Style = oStyleD
    oCell(vRow, 09).Style = oStyleD
    oCell(vRow, 10).Style = oStyleD
    oCell(vRow, 13).Style = oStyleL
    oCell(vRow, 14).Style = oStyleL
    oCell(vRow, 15).Style = oStyleL
    oCell(vRow, 16).Style = oStyleL

    oCell(vRow, 01) = fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", fCriteria(vMemb_Criteria))
    oCell(vRow, 02) = vMemb_FirstName
    oCell(vRow, 03) = vMemb_LastName
    oCell(vRow, 04) = vMemb_Organization
    oCell(vRow, 05) = vMemb_Email
    oCell(vRow, 06) = fYN(vMemb_Active)
    oCell(vRow, 07) = fId(vMemb_Id)
    oCell(vRow, 08) = fIf(fFormatDate(vMemb_FirstVisit) <> " ", vMemb_FirstVisit, "")
    oCell(vRow, 09) = fIf(fFormatDate(vMemb_LastVisit) <> " ", vMemb_LastVisit, "")  
    oCell(vRow, 10) = fIf(fFormatDate(vMemb_Expires) <> " ", vMemb_Expires, "")      
    oCell(vRow, 11) = vMemb_NoVisits
    oCell(vRow, 12) = FormatNumber(vMemb_NoHours/60,1)
    
    
    
    '...post results - each in it's own cell (if test then only one cell)
    aMemo = Split(vMemb_Memo, "|")
    For i = 0 To Ubound(aMemo)
      oCell(vRow, 13 + i) = aMemo(i)
    Next

  End Sub


 '...output spreadsheet if there are any rows
  Sub sExcelClose
    Response.ContentType = "application/vnd.ms-excel"
    oWs.Save "Learner Report dated " & fFormatDate(Now) & ".xls", 1
    Response.End
  End Sub
 
%>



