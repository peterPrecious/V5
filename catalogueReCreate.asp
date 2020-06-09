<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  '...Recreate Child Catalogue Utility (for Lori) - developed Apr 1, 2016 after archiving challenges


  Dim newCustId

  newCustId = Request.Form("newCustId")
  If Len(newCustId) = 8 Then

    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp6catalogueReCreate"
      .Parameters.Append .CreateParameter("@newCustId", adChar, adParamInput, 8, newCustId)
    End With
	  oCmdApp.Execute()
    Set oCmdApp = Nothing
    sCloseDbApp

    Response.Write newCustId & "recreated."

  End If

%>

<html>

<head>
  <title>sp6catalogueReCreate</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <%
  Server.Execute vShellHi
  %>

  <h1>Recreate A Child Account</h1>
  <h2>Enter the 8 char NEW Child Account</h2>
  <form method="POST" action="catalogueReCreate.asp">
    <h3>New Client Cust Id :
      <input name="newCustId" type="text" value="LAKE1124" />
      <input type="submit" value="Submit" name="bSubmit" class="button070"></h3>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
