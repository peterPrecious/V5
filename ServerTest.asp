<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
    Set oDb = Server.CreateObject("ADODB.Connection")
    svSQL = "(local)"
    svSQL = "SQL01,1400"

    oDb.ConnectionString = "Provider=SQLOLEDB.1;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=V5_Vubz;Data Source=" & svSQL
    response.write oDb.ConnectionString
    oDb.Open
    
    Set oRs = oDb.Execute("Select Count(*) AS [Count] FROM Cust")


    oDb.Close
    Set oDb = Nothing
%>


<html>
  <head>
    <meta http-equiv="Content-Language" content="en-us">
  </head>
  <body>
    <p>bite me</p>
  </body>
</html>
