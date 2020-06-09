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
        Case "Memb_Id"        : vMemb_Id            = Ucase(vValue)
        Case "Memb_FirstName" : vMemb_FirstName     = vValue      
        Case "Memb_LastName"  : vMemb_LastName      = vValue      
        Case "Memb_Pwd"       : vMemb_Pwd           = Ucase(vValue)
        Case "Memb_Email"     : vMemb_Email         = Lcase(vValue)
        Case "Memb_Programs"  : vMemb_Programs      = Ucase(vValue)
        Case "Memb_Memo"      : vMemb_Memo          = vValue      
      End Select

    Next

    vMemb_No = 0
'   sAddMemb vMemb_AcctId
    sAddMemb svCustAcctId

  Next
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" style="border-collapse: collapse" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9">
    <tr>
      <td align="center">&nbsp;<h1>Upload Learner Profiles (Basic) - Confirmation</h1>
      <h2 align="center">All learner profiles have been uploaded successfully!</h2>
      <p align="center">&nbsp;</p><h2 align="center">&nbsp;<a href="Upload_Basic.asp">Restart Upload</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="Users.asp">Learner List</a></h2></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
