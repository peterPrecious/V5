<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<%
  Dim sRecords, aRecord, vCnt, vError
  Dim vMemb_Id, vMemb_Pwd, vMemb_FirstName, vMemb_LastName, vMemb_Email, vMemb_Criteria

 '...update imported names
  If Request("vAction") = "update" Then
    sImportNames
    For vCnt = 0 To Ubound(sRecords)
      aRecord = Split(sRecords(vCnt) & ",", ",")
      vMemb_Id          = fNoQuote(Trim(Ucase(aRecord(0))))
      vMemb_Pwd         = fUnQuote(Trim(Ucase(aRecord(1))))
      vMemb_FirstName   = fUnQuote(Trim(aRecord(2)))  
      vMemb_LastName    = fUnQuote(Trim(aRecord(3)))  
      vMemb_Email       = fNoQuote(Trim(aRecord(4)))
      vMemb_Criteria    = fUnQuote(fDefault(Trim(aRecord(5)), 0))
      If Len(vMemb_Id) > 0 Then
        sUpdateImportedMembers
      End If
    Next
    Response.Redirect "SignOff.asp"
  End If

  Function fIsDupInList (vMembId, vCurrentNo)
    Dim i, aI
    fIsDupInList = False
    If vCurrentNo > 0 Then 
      '...check all uploaded sorted records, up to the current record for dup ids
      For i = 0 To vCurrentNo - 1
        aI = Split(sRecords(i), ",")
        If aI(0) = vMembId Then
          fIsDupInList = True
          Exit Function
        End If
      Next
    End If
  End Function
  
  Function fIsDupOnFile (vMembId)
    fIsDupOnFile = False
    If Len(vMembId) > 0 Then
      vSql  = "        SELECT Memb.Memb_No FROM Cust INNER JOIN Memb ON Cust.Cust_AcctId = Memb.Memb_AcctId"
      vSql  = vSql & " WHERE (Memb.Memb_Id = '" & vMembId & "') AND (Cust.Cust_Id = '" & svCustId & "')"
      sOpenDb    
      Set oRs  = oDB.Execute(vSql)
      If Not oRs.Eof Then 
        fIsDupOnFile = True
      End If
      Set oRs  = Nothing
      sCloseDB 
    End If
  End Function

  Sub sImportNames
    Dim oFs, oFile, vFile, vRecord, aRecords()
    Const ForReading = 1
    Set oFs = CreateObject("Scripting.FileSystemObject")   
    vFile = Server.MapPath("\" & svSite & "\Import\" & svCustId & ".csv")
    Set oFile = oFs.OpenTextFile(vFile, ForReading)
    i = -2
    Do While oFile.AtEndOfStream <> True
      i = i + 1
      '...ignore first title record
      If i = -1 Then 
        vRecord = oFile.ReadLine
      '...put into array
      Else
        ReDim Preserve aRecords (i)
        aRecords(i) = oFile.ReadLine
      End If
    Loop
    oFile.Close          
    '...sort the array
    sRecords = fSortArray(aRecords) 
  End Sub

  '...insert a new record if no Memb_No
  Sub sUpdateImportedMembers
    '...try to insert
    vSql = "INSERT INTO Memb"
    vSql = vSql & " (Memb_AcctId, Memb_Id, Memb_No, Memb_FirstName, Memb_LastName, Memb_Email, Memb_Pwd, Memb_Criteria)"
    vSql = vSql & " VALUES (" & svCustAcctId & ", '" & vMemb_Id & "', " & fNextMembNo & ", '" & vMemb_FirstName & "', '" & vMemb_LastName & "', '" & vMemb_Email & "', '" & vMemb_Pwd & "', " & vMemb_Criteria & ")"                               
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
    vSql = vSql & " Memb_Pwd        = '" & vMemb_Pwd       & "', " 
    vSql = vSql & " Memb_Email      = '" & vMemb_Email     & "', " 
    vSql = vSql & " Memb_Criteria   =  " & vMemb_Criteria  & "  " 
    vSql = vSql & " WHERE Memb_Id   = '" & vMemb_Id        & "'  "
    vSql = vSql & " AND Memb_AcctId = '" & svCustAcctId    & "'  "
    sOpenDb 
    oDb.Execute(vSql)
    sCloseDb
  End Sub

  Function fNextMembNo
    fNextMembNo = 0
    vSql = "SELECT TOP 1 Memb_No FROM Memb ORDER BY Memb_No DESC"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fNextMembNo = oRs2("Memb_No")
    Set oRs2 = Nothing      
    sCloseDb2
    fNextMembNo = fNextMembNo + 1
  End Function
%>


<html>

<head>
  <meta charset="UTF-8">
  <link href="http://vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <title>Upload</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
    Server.Execute vShellHi 

   
    <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
      <tr>
        <td><h1 align="center">Imported Learners</h1>
        <h2>This shows you the upload CSV file, sorted by Id, and will note any possible duplications, etc.&nbsp; If there are any problems, you can abort this job, fix the CSV file and start again or you can click &quot;Return&quot; if you do NOT wish to Continue with the upload.&nbsp; Error Messages can one of: <br>&nbsp; &quot;List Dup&quot; - there is a duplicate Id of this record on this list - all will be updated<br>&nbsp; &quot;File Dup&quot; - there is already a record on file with this User Id - ok to update<br>&nbsp; &quot;No Pwd&quot;&nbsp; - the Password is missing - ok to update<br>&nbsp; &quot;No Id&quot;&nbsp;&nbsp;&nbsp;&nbsp; - missing User Id - cannot update this record<br>If you are ok to update the database then click &quot;<b>Continue</b>&quot; remembering that any database duplicates will be overwritten and missing User Ids will be ignored. Otherwise click &quot;<b>Return</b>&quot; to NOT update the database.</h2>
        </td>
      </tr>
      <tr>
        <td align="center">&nbsp;</td>
      </tr>
      <tr>
        <th align="center">
        <table border="1" id="table1" cellspacing="1" cellpadding="3" bordercolor="#DCEDF8" style="border-collapse: collapse">
          <tr>
            <th nowrap bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="left">User Id </th>
            <th nowrap bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="left">Password </th>
            <th nowrap bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="left">First <br>Name </th>
            <th nowrap bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="left">Last <br>Name </th>
            <th nowrap bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="left">Email <br>Address</th>
            <th nowrap bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="left">Department <br>No</th>
            <th nowrap align="left">&nbsp;</th>
            <td bgcolor="#DCEDF8" bordercolor="#FFFFFF" align="center">
            <p class="c5">Errors</td>
          </tr>
          <%
            '...import and sort names, then split and check for dups
            sImportNames
            For vCnt   = 0 To Ubound(sRecords)
              aRecord  = Split(sRecords(vCnt) & ",", ",")  '...tack on "," in case there is no criteria field
              vMemb_Id = Trim(aRecord(0))
              vError   = ""
              vError   = vError & fIf(vMemb_Id = "", " | No Id", "")
              vError   = vError & fIf(fIsDupInList (vMemb_Id, vCnt), " | List Dup", "")
              vError   = vError & fIf(fIsDupOnFile (vMemb_Id), " | File Dup", "")
              vError   = vError & fIf(Trim(aRecord(1)) = "", " | No Pwd", "")
              If Len(vError) > 0 Then vError = Mid(vError, 4)
          %> 
          <tr>
            <td><%=Trim(Ucase(aRecord(0)))%></td>
            <td><%=Trim(Ucase(aRecord(1)))%></td>
            <td><%=Trim(aRecord(2))%></td>
            <td><%=Trim(aRecord(3))%></td>
            <td><%=Trim(aRecord(4))%></td>
            <td><%=Trim(aRecord(5))%></td>
            <td>&nbsp;</td>
            <td align="center" class="c6"><%=vError%></td>
          </tr>
          <%
            Next
          %>         
        </table>
        <p><a href="javascript:history.back(1)"><img border="0" src="Images/Buttons/Return_EN.gif" width="81" height="20"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="ImportOk.asp?vAction=update"><img border="0" src="Images/Buttons/Continue_EN.gif" width="110" height="20"></a></th>
      </tr>
      <tr>
        <td align="center">
        &nbsp;</td>
      </tr>
    </table>
    
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
