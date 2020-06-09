<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Sort_Routine.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <title>:: Session Variables</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->
  <table style="text-align: center">
    <tr>
      <th colspan="2" class="rowshade">&nbsp; Session Variables&nbsp; </th>
    </tr>
    <%
        Dim aSessions(), sSessions, aSession, aProd, vBg
        i = -1
        For Each j In Session.Contents
           If j <> "HostDbPwd" Then
             i = i + 1
             ReDim Preserve aSessions(i)
             If VarType(Session(j)) < 16 AND VarType(Session(j)) <> 9 Then 
               aSessions(i) = j & "||" & Session(j)
             Else
               If j = "Prod" Then
                 aProd        = Session("Prod")
                 aSessions(i) = j & "||" & "<i>Prod</i>"
               Else
                 aSessions(i) = j & "||" & "<i>Not printable</i>"
               End If
             End If
           End If
        Next

        sSessions = fSortArray(aSessions)
        
        For i = 0 To Ubound(sSessions)
          aSession = Split(sSessions(i), "||")
          If aSession(0) = "QueryString" Then
            aSession(1) = Replace(aSession(1), "&", " &")
          ElseIf aSession(0) = "Browser" Then
            aSession(1) = Replace(aSession(1), "|", " |")
          ElseIf aSession(0) = "Prod" Then
            aSession(1) = "caca pooh"
          End If
    %>
    <tr>
      <th><%=aSession(0)%> :</th>
      <td>
        <% 
           If aSession(1) = "caca pooh" Then
        %>
        <table class="table">
          <%
           For k = 1 To Ubound(aProd, 2)
             For j = 0 to 6
             vBg = "" : If k Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9'"
          %>
          <tr>
            <td><b><%=k%></b></td>
            <td>&nbsp;- <%=j%> - </td>
            <td><%=fLeft(aProd(j, k), 24)%></td>
          </tr>
          <%
             Next
           Next
          %>
        </table>
        <%
           ElseIf aSession(0) = "SessionStarted" Then
             Response.Write aSession(1) & "<br>(" & Session.Timeout - DateDiff("n", Session("SessionStarted"), Now())  & " mins remaining)"
           Else
             Response.Write aSession(1)
           End If
        %>
      </td>
    </tr>
    <%  
        Next
    %>
    <tr>
      <td colspan="2" style="text-align:center"><br>
        <input onclick="location.href = location.href" type="button" value="Refresh" name="bRefresh" class="button100"><%=f10() %>
        <input onclick="window.close()" type="button" value="Close" name="bClose" class="button100"><br>&nbsp;</td>
    </tr>
  </table>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
</body>

</html>
