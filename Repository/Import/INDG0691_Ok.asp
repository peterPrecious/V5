<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim oFs, oFile, vFile, vRecord, aRecord, vCnt, bOk, vError, vUrl
  Dim vMemb_Id, vMemb_FirstName, vMemb_LastName, vMemb_Criteria, vCrit_Id
  Dim v1, v2, v3, v4, v5, vb1, vb2, vb3, vb4, vb5, va1, va2, va3, va4, va5

  Const ForReading = 1
  Set oFs = CreateObject("Scripting.FileSystemObject")   
  vFile = Server.MapPath("INDG0691.TXT")

  '...ensure all criteria values are valid  
  Set oFile = oFs.OpenTextFile(vFile, ForReading)
  i = 0 : bOk = True
  Do While oFile.AtEndOfStream <> True
    i = i + 1
    vRecord = oFile.ReadLine
    If i > 1 Then '...ignore header
      aRecord = Split(vRecord, vbTab)
      vCrit_Id = Right("0000" & Ucase(aRecord(0)), 4) & " " & Right("0000" & Ucase(aRecord(1)), 4) & " " & Ucase(aRecord(2))
      If fMembCriteria("0691", vCrit_Id) = 0 Then
        vError = vError & "<br>" & vCrit_Id & " (Row: " & i & ")"
        bOk = False        
      End If
    End If  
  Loop
  oFile.Close          

  If Not bOk Then
    vUrl = "Message.asp?vNext=Default.asp&vMsg=" & Replace("The following Group(s) must be setup before you can Import Learners :<br>" & fLeft(vError, 1000) & "<br>", " ", "+")
    Response.Redirect vUrl

  Else
    '...count users before update and store in "before variables"
    sCountUsers
    vb1 = v1:   vb2 = v2:   vb3 = v3:   vb4 = v4:   vb5 = v1+v2+v3+v4
  

    '...then inactivate all members
    sInactivateMembers
  
   '...imported names into a table
    Set oFile = oFs.OpenTextFile(vFile, ForReading)
    i = 0
    Do While oFile.AtEndOfStream <> True
      i = i + 1
      vRecord = oFile.ReadLine
      If i > 1 Then '...ignore header
        aRecord = Split(vRecord, vbTab)
    
        vMemb_Id          = Right("000000" & Ucase(aRecord(3)), 6)
        vMemb_LastName    = fUnquote(aRecord(4))  
        vMemb_FirstName   = fUnquote(aRecord(5))  
        vCrit_Id          = Right("0000" & Ucase(aRecord(0)), 4) & " " & Right("0000" & Ucase(aRecord(1)), 4) & " " & Ucase(aRecord(2))
        vMemb_Criteria    = fMembCriteria("0691", vCrit_Id)
        sUpdateImportedMembers
      End If  
    Loop
    oFile.Close          
  
    '...count users after update and store in "after variables"
    sCountUsers
    va1 = v1:   va2 = v2:   va3 = v3:   va4 = v4:   va5 = v1+v2+v3+v4

  End If
  
  '...inactive all members (imported members will be active)
  Sub sInactivateMembers
    vSql = "UPDATE Memb SET Memb_Active = 0 WHERE Memb_AcctId = '0691' AND Memb_Level = 2"
    sOpenDb 
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  '...inactive all members (imported members will be active)
  Sub sCountUsers
    sOpenDb 
    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '0691' AND Memb_Level = 2 AND Memb_Active = 1"
    Set oRs = oDb.Execute(vSql)
    v1 = oRs("Cnt")

    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '0691' AND Memb_Level = 2 AND Memb_Active = 0"
    Set oRs = oDb.Execute(vSql)
    v2 = oRs("Cnt")

    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '0691' AND Memb_Level = 3 AND Memb_Id NOT LIKE 'VUV5_%'"
    Set oRs = oDb.Execute(vSql)
    v3 = oRs("Cnt")

    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '0691' AND Memb_Level = 4 AND Memb_Id NOT LIKE 'VUV5_%'"
    Set oRs = oDb.Execute(vSql)
    v4 = oRs("Cnt")

    sCloseDb
  End Sub


  '...insert a new record if no Memb_No
  Sub sUpdateImportedMembers

    '...try to insert
    vSql = "INSERT INTO Memb"
    vSql = vSql & " (Memb_AcctId, Memb_Id, Memb_FirstName, Memb_LastName, Memb_Criteria)"
    vSql = vSql & " VALUES ('0691', '" & vMemb_Id & "', '" & vMemb_FirstName & "', '" & vMemb_LastName & "', " & vMemb_Criteria & ")"                               
    On Error Resume Next
    sOpenDb 
    oDb.Execute(vSql)
    If Err.Number = 0 Or Err.Number = "" Then 
      sCloseDb
      Exit Sub
    End If

    '...if on file then update
    On Error GoTo 0
    vSql = "UPDATE Memb SET"
    vSql = vSql & " Memb_FirstName  = '" & vMemb_FirstName & "', " 
    vSql = vSql & " Memb_LastName   = '" & vMemb_LastName  & "', " 
    vSql = vSql & " Memb_Criteria   =  " & vMemb_Criteria  & " , " 
    vSql = vSql & " Memb_Active     =  " & 1               & "  " 
    vSql = vSql & " WHERE Memb_Id   = '" & vMemb_Id        & "'  "
    vSql = vSql & " AND Memb_AcctId = '0691' "
    sOpenDb 
    oDb.Execute(vSql)
    sCloseDb
  End Sub


    '...if on table as a single integer return criteria no (0 is not acceptable)
  Function fMembCriteria(vCritAcctId, vCriteria)
    fMembCriteria = 0
    vSql = "SELECT Crit_No FROM Crit WHERE Crit_Id = '" & vCriteria & "' AND Crit_AcctId = '" & vCritAcctId & "'"
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      On Error Resume Next
      fMembCriteria = Clng(oRs2("Crit_No"))
      On Error Goto 0
      Set oRs2 = Nothing
      sCloseDb2    
      Exit Function
    End If
  End Function


%>
<html>

<head>
  <meta charset="UTF-8">
  <link href="<%=svDomain%>/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
</head>


<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
    Server.Execute vShellHi 
  %>

  <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
    <tr>
      <td align="center">
      <h1>Chapter Indigo | Import Learners into VUBIZ LMS</h1>
      <h1>Results</h1>
      <h2>Thank you.&nbsp; All records have been imported.</h2>
      <table border="1" id="table2" cellspacing="3" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9">
        <tr>
          <td colspan="3" align="center">
          <h2>Totals</h2>
          </td>
        </tr>
        <tr>
          <td>
          <h2>&nbsp;</h2>
          </td>
          <td align="right">
          <h2>Before</h2>
          </td>
          <td align="right">
          <h2>After</h2>
          </td>
        </tr>
        <tr>
          <th align="right">No active users : </th>
          <td align="right"><%=vb1%></td>
          <td align="right"><%=va1%></td>
        </tr>
        <tr>
          <th align="right">No inactive users :</th>
          <td align="right"><%=vb2%></td>
          <td align="right"><%=va2%></td>
        </tr>
        <tr>
          <th align="right">No facilitators :</th>
          <td align="right"><%=vb3%></td>
          <td align="right"><%=va3%></td>
        </tr>
        <tr>
          <th align="right">No managers :</th>
          <td align="right"><%=vb4%></td>
          <td align="right"><%=va4%></td>
        </tr>
        <tr>
          <th align="right">No users total: </th>
          <td align="right"><%=vb5%></td>
          <td align="right"><%=va5%></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_LoLite.asp"-->

</body>

</html>
