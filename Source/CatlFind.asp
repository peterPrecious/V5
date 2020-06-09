<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  Dim vMsg, vId, vVb, vProgram
  vId = Ucase(Request("vId"))
  vVb = fDefault(Request("vVb"), "n")

  '...form submitted?
  If Len(vId) >= 6 Then

    '...looking for programs
    If Left(vId, 1) = "P" Then

      vSql        = " SELECT Catl.Catl_CustId As [Cust_Id], Cust.Cust_Title As Title" _
                  & " FROM Catl" _
                  & " INNER JOIN Cust ON Catl.Catl_CustId = Cust.Cust_Id"
      If vVb = "n" Then
      vSql = vSql & " WHERE (Cust.Cust_AcctId NOT LIKE '7%') AND (Catl_Programs LIKE '%" & vId & "%')"
      Else
      vSql = vSql & " WHERE (Catl_Programs LIKE '%" & vId & "%')"
      End If
      vSql = vSql & " ORDER BY Cust.Cust_Id"

    '...looking for Modules
    Else
   
      vSql        = " SELECT     Catl.Catl_CustId AS Cust_Id, Cust.Cust_Title AS Title, V5_Base.dbo.Prog.Prog_Id AS Program" _
                  & " FROM Catl" _ 
                  & " INNER JOIN Cust ON Catl.Catl_CustId = Cust.Cust_Id" _ 
                  & " INNER JOIN V5_Base.dbo.Prog ON CHARINDEX(V5_Base.dbo.Prog.Prog_Id, Catl.Catl_Programs) > 0"
      If vVb = "n" Then
      vSql = vSql & " WHERE (Cust.Cust_AcctId NOT LIKE '7%') AND (CHARINDEX('" & vId & "', V5_Base.dbo.Prog.Prog_Mods) > 0)"
      Else
      vSql = vSql & " WHERE (CHARINDEX('" & vId & "', V5_Base.dbo.Prog.Prog_Mods) > 0)"
      End If
      vSql = vSql & " ORDER BY Cust.Cust_Id"

    End If
   
'   sDebug

    sOpenDb
    Set oRs = oDb.Execute(vSql)    
    
    If Not oRs.Eof Then
      vMsg = "<h1>" & vId & " can be found in Account(s):</h1>"
      vMsg = vMsg &  "<table border='0'>"
      Do While Not oRs.Eof

        If Left(vId, 1) <> "P" Then
          vProgram = oRs("Program")
        Else
          vProgram = ""
        End If

        vMsg = vMsg & "<tr><td>" & oRs("Cust_Id") & "</td><td>" & fLeft(oRs("Title"), 32) & "</td><td>" & vProgram & "</td></tr>"
        oRs.MoveNext
      Loop
      vMsg = vMsg &  "</table>"

    Else
      vMsg = "That " & fIf(Left(vId, 1) = "P", "Program", "Module") & " Id is not in any Account."
    End If

    sCloseDb  

  End If
%>

<html>

<head>
  <title>CatFind</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <table>
    <tr>
      <td>
        <h1>Catalogue Search</h1>
        <h2>This lists Accounts whose Catalogue contains the following Module ID or Program ID.</h2>
        <br />
        <br />
        <form method="POST" action="CatlFind.asp">
          <h3>
            <input type="radio" value="y" name="vVB" <%=fCheck(vVb, "y")%>>Include 
            <input type="radio" value="n" name="vVB" <%=fCheck(vVb, "n")%>>Exclude VuBuild Accounts
            <br />
            <br />
            Enter Module or Program Id :
            <input type="text" name="vId" size="8" maxlength="7" value="<%=vId%>">
            <input type="submit" value="Go" name="bGo" class="button">
          </h3>
        </form>
      </td>
    </tr>

    <% If Len(vMsg) > 0 Then %>
    <tr>
      <td>
        <br>
        <%=vMsg%></td>
    </tr>
    <% End If %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
