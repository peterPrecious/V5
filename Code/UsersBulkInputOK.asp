<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim aFieldNamesDb, aAllRecords, vReturnInfo, vNextMembNo
  Dim vRecNo, vFldNo

  '...Retrieve data needed to Import into Db
  aFieldNamesDb   = Session("ImportFieldNamesDb")
  aAllRecords     = Session("ImportAllRecords")
  vReturnInfo     = Request.QueryString

  vMemb_Criteria = Request.QueryString("vCrit")

  For vRecNo = 0 To UBound(aAllRecords,2)

    For vFldNo = 0 To Ubound(aFieldNamesDb)

      vFld    = aFieldNamesDb(vFldNo)
      vValue  = fUnquote(aAllRecords(vFldNo, vRecNo))

'     Response.Write vFld & " = " & vValue

      Select Case vFld
        Case "Memb_Id"        : vMemb_Id            = Trim(Ucase(Ucase(vValue)))
        Case "Memb_FirstName" : vMemb_FirstName     = vValue      
        Case "Memb_LastName"  : vMemb_LastName      = vValue      
        Case "Memb_Pwd"       : vMemb_Pwd           = Ucase(vValue)
        Case "Memb_Email"     : vMemb_Email         = Lcase(vValue)
        Case "Memb_Programs"  : vMemb_Programs      = Ucase(vValue)
        Case "Memb_Memo"      : vMemb_Memo          = vValue      
      End Select

    Next

    vMemb_No = 0
    vMemb_Criteria = fIf(Len(vMemb_Criteria) = 0, "0", vMemb_Criteria)
    sAddMemb svCustAcctId

  Next
%>

<html>

<head>
  <title>UsersBulkInputOL</title>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Learners were Uploaded Successfully.</h1>
  <h2><a href="UsersBulkInput.asp">Restart Upload Program</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="Users.asp">Learner List</a></h2>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


