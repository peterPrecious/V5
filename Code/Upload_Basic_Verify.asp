<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vList, vFname, vLname, vEmail, vPassw, vProgs, vMemos, vCrit, vReturnInfo
  Dim vRecordSep, vFieldSep
  Dim aRecords, aFieldNames, aFieldNamesDb, vEmailPos, vNumRecords, vFieldCount, aTemp, vProgsPos
  Dim aAllRecords, vRecWidth

  Dim vUserIds, vEmails, vWordPos, vUserIdsDb, vEmailsDb
  Dim vErrMess

  vList      = Trim(Request.Form("tList"))
  vFieldSep  = Request.Form("oField")
  vRecordSep = Request.Form("oRecord")

  vFname     = Request.Form("cFName") : If fNoValue(vFname) Then vFname = "False"
  vLname     = Request.Form("cLName") : If fNoValue(vLname) Then vLname = "False"
  vEmail     = Request.Form("cEmail") : If fNoValue(vEmail) Then vEmail = "False" : vEmailPos = 0
  vPassw     = Request.Form("cPassw") : If fNoValue(vPassw) Then vPassw = "False"
  vProgs     = Request.Form("cProgs") : If fNoValue(vProgs) Then vProgs = "False"
  vMemos     = Request.Form("cMemos") : If fNoValue(vMemos) Then vMemos = "False"

  vCrit      = Replace(Request.Form("dCrit"), ",", "") '...remove the commas

  '...Determine amount of Fields
  Redim aFieldNames(0)   : aFieldNames(0)   = "Learner Id"
  Redim aFieldNamesDb(0) : aFieldNamesDb(0) = "Memb_Id"

  vFieldCount = 1

  If vFname = "True" Then
    vFieldCount = vFieldCount + 1
    Redim Preserve aFieldNames(Ubound(aFieldNames)+1)
    aFieldNames(Ubound(aFieldNames)) = "First Name"
    Redim Preserve aFieldNamesDb(Ubound(aFieldNamesDb)+1)
    aFieldNamesDb(Ubound(aFieldNamesDb)) = "Memb_FirstName"
  End If

  If vLname = "True" Then
    vFieldCount = vFieldCount + 1
    Redim Preserve aFieldNames(Ubound(aFieldNames)+1)
    aFieldNames(Ubound(aFieldNames)) = "Last Name"
    Redim Preserve aFieldNamesDb(Ubound(aFieldNamesDb)+1)
    aFieldNamesDb(Ubound(aFieldNamesDb)) = "Memb_LastName"
  End If

  If vEmail = "True" Then
    vFieldCount = vFieldCount + 1
    vEmailPos = vFieldCount
    Redim Preserve aFieldNames(Ubound(aFieldNames)+1)
    aFieldNames(Ubound(aFieldNames)) = "Email"
    Redim Preserve aFieldNamesDb(Ubound(aFieldNamesDb)+1)
    aFieldNamesDb(Ubound(aFieldNamesDb)) = "Memb_Email"
  End If

  If vPassw = "True" Then
    vFieldCount = vFieldCount + 1
    Redim Preserve aFieldNames(Ubound(aFieldNames)+1)
    aFieldNames(Ubound(aFieldNames)) = "Password"
    Redim Preserve aFieldNamesDb(Ubound(aFieldNamesDb)+1)
    aFieldNamesDb(Ubound(aFieldNamesDb)) = "Memb_Pwd"
  End If

  If vProgs = "True" Then
    vFieldCount = vFieldCount + 1
    vProgsPos = vFieldCount
    Redim Preserve aFieldNames(Ubound(aFieldNames)+1)
    aFieldNames(Ubound(aFieldNames)) = "Programs"
    Redim Preserve aFieldNamesDb(Ubound(aFieldNamesDb)+1)
    aFieldNamesDb(Ubound(aFieldNamesDb)) = "Memb_Programs"
  End If

  If vMemos = "True" Then
    vFieldCount = vFieldCount + 1
    Redim Preserve aFieldNames(Ubound(aFieldNames)+1)
    aFieldNames(Ubound(aFieldNames)) = "Memo"
    Redim Preserve aFieldNamesDb(Ubound(aFieldNamesDb)+1)
    aFieldNamesDb(Ubound(aFieldNamesDb)) = "Memb_Memo"
  End If



  '...Set the Original info to display on main page
  vReturnInfo = "&vCrit=" & vCrit & "&vFieldSep=" & vFieldSep & "&vRecordSep=" & vRecordSep & "&vFname=" & vFname & "&vLname=" & vLname & "&vEmail=" & vEmail & "&vPassw=" & vPassw & "&vProgs=" & vProgs & "&vMemos=" & vMemos
  Session("BulkImportList") = vList

  '...Check if Text is empty
  If Len(vList) = 0 Then 
    Response.Redirect "Upload_Basic.asp?vErrMess=" & Server.URLEncode("Please enter data to Import into the Text Area.") & vReturnInfo
  '...Check if delimiters are the same for Field and Record
  ElseIf vFieldSep = vRecordSep Then
    Response.Redirect "Upload_Basic.asp?vErrMess=" & Server.URLEncode("Please ensure that the Field Delimeter and Record Delimeter are <i>not</i> the same.") & vReturnInfo
  End If

  '...Define delimiters
  Select Case vFieldSep
    Case "comma" : vFieldSep = ","
    Case "semi"  : vFieldSep = ";"
    Case "pipe"  : vFieldSep = "|"
    Case "tilde" : vFieldSep = "~"
    Case "tab"   : vFieldSep = vbTab
  End Select

  Select Case vRecordSep
    Case "comma" : vRecordSep = ","
    Case "semi"  : vRecordSep = ";"
    Case "pipe"  : vRecordSep = "|"
    Case "tilde" : vRecordSep = "~"
    Case "tab"   : vRecordSep = vbTab
    Case "enter" : vRecordSep = vbCrLf
  End Select

  '...Make sure the last character is not a trailing Record delimiter
  If vRecordSep = vbCrLf And Right(vList,2) = vRecordSep Then
    vList = Left(vList, Len(vList)-2)
  ElseIf Right(vList,1) = vRecordSep Then
    vList = Left(vList, Len(vList)-1)
  End If

  '...Make sure we have correct amount of Fields
  aRecords = Split(vList,vRecordSep)
  vNumRecords = Ubound(aRecords)
  For i = 0 to vNumRecords
    aTemp = Split(aRecords(i), vFieldSep)
    If Ubound(aTemp) <> (vFieldCount - 1) Then
      Response.Redirect "Upload_Basic.asp?vErrMess=" & Server.URLEncode("Invalid number of Fields in Record " & i+1 & ".") & vReturnInfo
    '...Check if Password is blank
    ElseIf Len(aTemp(0)) = 0 Then
      Response.Redirect "Upload_Basic.asp?vErrMess=" & Server.URLEncode("Password is blank in Record " & i+1 & ".") & vReturnInfo
    End If
    '...Build Master Password and Email list...to check for duplicates
    vUserIds = vUserIds & "@@@" & aTemp(0) & "***"
    If vEmailPos > 0 Then
      vEmails = vEmails & "@@@" & aTemp(vEmailPos-1) & "***"
    End If
  Next

  '...Convert to upper case for comparisons
  vUserIds   = Ucase(vUserIds)
  vEmails    = Ucase(vEmails)

  '...Get all UserIds and Emails from the Memb table
  sOpenDb
  vSql = "SELECT * FROM Memb WITH (nolock) WHERE Memb_AcctId=" & svCustAcctId
  Set oRs = oDb.Execute(vSql)
  While Not oRs.Eof
    vUserIdsDb = vUserIdsDb & "@@@" & oRs("Memb_Id") & "***"
    If vEmailPos > 0 Then
      vEmailsDb = vEmailsDb & "@@@" & oRs("Memb_Email") & "***"
    End If
    oRs.MoveNext
  Wend
  Set oRs = Nothing
  sCloseDb

  '...Convert to upper case for comparisons
  vUserIdsDb = Ucase(vUserIdsDb)
  vEmailsDb  = Ucase(vEmailsDb)

  vErrMess = ""

  '...Make sure no Password or Email duplicates
  For i = 0 to vNumRecords
    aTemp = Split(aRecords(i), vFieldSep)

    '...Check UserIds from User List
    vWordPos = InStr(1, vUserIds,"@@@" & Ucase(aTemp(0)) & "***")
    vWordPos = InStr(vWordPos+1, vUserIds,"@@@" & Ucase(aTemp(0)) & "***")
    If vWordPos > 0 Then
      vErrMess = vErrMess & "Learner Id [" & aTemp(0) & "]" & " in record " & i+1 & " is duplicated within List.<br>"
    End If

    '...Check UserIds from Database
    vWordPos = InStr(1, vUserIdsDb,"@@@" & Ucase(aTemp(0)) & "***")
    If vWordPos > 0 Then
      vErrMess = vErrMess & "Learner Id [" & aTemp(0) & "]" & " in record " & i+1 & " is already in the Database.<br>"
    End If

    '...Check Emails from User List (ignore this code)
    If vEmailPos > 0 Then
      '...ignore blank emails
      If Len(aTemp(vEmailPos-1)) > 0 Then
        vWordPos = InStr(1, vEmails, "@@@" & Ucase(aTemp(vEmailPos-1)) & "***")
        vWordPos = InStr(vWordPos+1, vEmails, "@@@" & Ucase(aTemp(vEmailPos-1)) & "***")
        If vWordPos > 0 Then
'         vErrMess = vErrMess & "Email [" & aTemp(vEmailPos-1) & "]" & " in record " & i+1 & " is duplicated within List.<br>"
        End If  
        '...Check Emails from Database
        vWordPos = InStr(1, vEmailsDb, "@@@" & Ucase(aTemp(vEmailPos-1)) & "***")
        If vWordPos > 0 Then
'         vErrMess = vErrMess & "Email [" & aTemp(vEmailPos-1) & "]" & " in record " & i+1 & " is already in the Database.<br>"
        End If
      End If       
    End If

  Next

  '...Bail out if any ErroRs
  If Len(vErrMess) > 0 Then
    Response.Redirect "Upload_Basic.asp?vErrMess=" & Server.URLEncode(fLeft(vErrMess, 1000)) & vReturnInfo
  End If

  ReDim aAllRecords(vFieldCount-1,vNumRecords)
  For i = 0 to vNumRecords
    aTemp = Split(aRecords(i), vFieldSep)

    '...capitalize Progs
    If vProgsPos > 0 Then aTemp(vProgsPos-1) = Ucase(aTemp(vProgsPos-1))

    For j = 0 to Ubound(aTemp)
      aAllRecords(j,i) = aTemp(j)
    Next
  Next

  '...Store data needed on Save page into Session vars
  Session("ImportFieldNamesDb") = aFieldNamesDb
  Session("ImportAllRecords")   = aAllRecords

  vRecWidth = 100/(vFieldCount+1)

  '...Response.Write "Number of Records is " & vNumRecords + 1 & ".<br>"
  '...Response.Write "Number of Fields is " & vFieldCount & "."
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" style="border-collapse: collapse" width="100%" bordercolor="#DDEEF9" cellspacing="0" cellpadding="3">
    <tr>
      <td align="center" colspan="<%=vFieldCount %>"><h1>Upload Learner Profiles (Basic) - Verify</h1><h2>Please verify the following information is ok to Upload...</h2></td>
    </tr>
    <%
      For i = 0 to Ubound(aAllRecords, 2)
        If i = 0 Then
          Response.Write "<tr>"
          For k = 0 to Ubound(aAllRecords, 1)
            Response.Write "<th width='" & vRecWidth & "%' align='left' bgcolor='#DDEEF9'>" & aFieldNames(k) & "</th>"
          Next
          Response.Write "</tr>"
        End If
  
        Response.Write "<tr>"
        For j = 0 to Ubound(aAllRecords,1)
          If Len(aAllRecords(j, i)) = 0 Then
            Response.Write "<td width='" & vRecWidth & "%'>&nbsp;</td>"
          Else
            Response.Write "<td width='" & vRecWidth & "%'>" & aAllRecords(j, i) & "</td>"
          End If
        Next
        Response.Write "</tr>"
      Next
    %>
    <tr>
      <td align="center" colspan="<%=vFieldCount %>">&nbsp;<p>
        <input onclick="location.href='Upload_Basic.asp?abc=<%=vReturnInfo%>'" type="button" value="Edit Values" name="bEdit" class="button"><%=f10%> 
        <input onclick="location.href='Upload_Basic_Ok.asp?abc=<%=vReturnInfo%>'" type="button" value="Save Values" name="bSave" class="button"> </p><h2><a href="Users.asp">Learner List</a></h2>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


