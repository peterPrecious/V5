<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <base target="_self">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %> <center>
  <table border="1" width="90%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellspacing="0" cellpadding="3">
    <tr>
      <th width="25%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" align="left"><%=fIf(svCustPwd, "Learner Id", "Password")%></th>
      <th width="25%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" align="left">First Name</th>
      <th width="25%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" align="left">Last Name</th>
      <th width="25%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" align="left">Email Address</th>
    </tr>
    <%
      Dim vFld2

      '...go through the collection and find the vMemb field 
      For Each vFld In Request.Form
        If Left(vFld, 5) = "vMemb" Then
          vMemb_Id = Request.Form(vFld)
          vMemb_No = Mid(vFld, 7)

          '...then get all related fields, update memb file then display
          For Each vFld2 In Request.Form
            If Mid(vFld2, 7) = vMemb_No Then
              Select Case Left(vFld2, 5)
                Case "vFrst" : vMemb_FirstName = Trim(Request.Form(vFld2))
                Case "vLast" : vMemb_LastName  = Trim(Request.Form(vFld2))
                Case "vEmai" : vMemb_Email     = Trim(Request.Form(vFld2))
              End Select
            End If
          Next
           
         '...update member file 
         If Len(vMemb_FirstName) > 0 Or Len(vMemb_LastName) > 0 Or Len(vMemb_Email) > 0 Then 
           sUpdateMemb_Profile
         End If
           
         '...display new members
      %>
    <tr>
      <td width="25%" align="left"><a href="User<%=fGroup%>.asp?vMembNo=<%=vMemb_No%>"><%=vMemb_Id%></a> </td>
      <td width="25%"><%=vMemb_FirstName%>&nbsp; </td>
      <td width="25%"><%=vMemb_LastName%>&nbsp; </td>
      <td width="25%"><%=vMemb_Email%>&nbsp; </td>
    </tr>
    <%
        End If
      Next
    %>
  </table>
  </center>
  <h2 align="center"><a href="IssueIds.asp">Return to Issue Passwords</a></h2>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


