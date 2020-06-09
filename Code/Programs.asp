<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/CustomCertRoutines.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->


<% 
  Dim vCustId, vAddProgId, vEditProgId, vFunction, vMods, vRange, vTitle, vLingo, vNextId

  vNextId = sp7nextProgramId()

  vFunction = ""
  vLingo = fDefault(Request("vLingo"), "EN, FR, ES, PT")
  vRange = fDefault(Request("vRange"), "")
  vTitle = fDefault(Request("vTitle"), "")

  '...get next Program Id
  Function sp7nextProgramId()
    sp7nextProgramId=""
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp7nextProgramId"     
    End With
    Set oRsApp = oCmdApp.Execute()
    If Not oRsApp.Eof Then
      sp7nextProgramId = oRsApp("nextProgramId")
    End If
    Set oRsApp = Nothing
    Set oCmdApp = Nothing
    sCloseDbApp
  End Function 
    
%>

<html>

<head>
  <title>Programs</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
    function validate(ele) {
      if ($(ele)[0].value.length != 7) {
        alert("Please enter a 7 character Program Id.");
        $(ele)[0].focus();
        return (false);
      } 
       return (true) 
    }
  </script>
</head>

<body>

  <% 
    Server.Execute vShellHi 
  %>

  <h1>Program Table</h1>
  <table class="table">
    <tr>

      <td style="width: 65%; padding: 10px;">
        <p class="c2">
          Select the Programs to list, max 50 then click <span class="code">Next.</span></p>
        <form method="POST" action="Programs.asp">
          <table class="table">
            <tr>
              <th>Ids that start with :</th>
              <td class="debug">
                <input type="text" name="vRange" id="vRange" size="8" value="<%=vRange%>">
              </td>
            </tr>
            <tr>
              <th>of language :</th>
              <td>
                <input type="checkbox" name="vLingo" value="EN" <%=fchecks(vlingo, "en")%>>EN 
                <input type="checkbox" name="vLingo" value="FR" <%=fchecks(vlingo, "fr")%>>FR 
                <input type="checkbox" name="vLingo" value="ES" <%=fchecks(vlingo, "es")%>>ES
                <input type="checkbox" name="vLingo" value="PT" <%=fchecks(vlingo, "pt")%>>PT 
              </td>
            </tr>
            <tr>
              <th>whose Titles containing :</th>
              <td>
                <input type="text" name="vTitle" size="30" value="<%=vTitle%>">
              </td>
            </tr>
            <tr>
              <td colspan="2" style="text-align: center;">
                <input type="submit" value="Next" name="bGo" class="button070">
              </td>
            </tr>
          </table>
        </form>
      </td>

      <td style="text-align: center; width: 35%; padding: 10px;">
        <form method="POST" id="programAdd" action="Program.asp" name="fAdd" onsubmit="return validate('#vAddProgId')">
          <p class="c2">
            To add a new Program, use the next available Program Id below then click <span class="code">Add</span>.&nbsp;
          </p>
          <br /><br />
          <input type="hidden" name="vRange" value="<%=vRange%>">
          <input type="hidden" name="vLingo" value="<%=vLingo%>">
          <input type="text" name="vAddProgId" id="vAddProgId" size="8" value="<%=vNextId %>">
          <input type="submit" value="Add" name="bAdd" class="button070"></form>
      </td>

    </tr>
  </table>

  <table class="table">
    <tr>
      <td class="rowshade">Id</td>
      <td class="rowshade">Title</td>
      <td class="rowshade">Owner<br>Id</td>
      <td class="rowshade">Retired?</td>
      <td class="rowshade">Clone into<br /><span style="background-color: #FFFF00">New</span> Program Id</td>
      <td class="rowshade">Certificate using<br>Cust Id | Acct Id</td>
    </tr>
    <%
      i = 0 '...used to determine if we have listed all 50 programs
      '...read Prog
      sOpenDbBase
      vSql = "SELECT TOP 50 * FROM Prog "_
           & " WHERE Prog_Id > '" & Left(vRange, 5) & "'" _
           & "   AND CHARINDEX(RIGHT(Prog_Id, 2), '" & vLingo & "') > 0 "_
           & fIf (vTitle <> "",   " AND Prog_Title1 LIKE '%" & vTitle & "%'", "") _ 
           & " ORDER BY Prog_Id"  

   '   stop
      Set oRsBase = oDbBase.Execute(vSQL)    
      Do While Not oRsBase.Eof 
        sReadProg  
        i = i + 1  
    %>
    <tr>
      <td><a href="Program.asp?vEditProgId=<%=vProg_Id%>&vHidden=n&vLingo=<%=vLingo%>&vRange=<%=vRange%>"><%=vProg_Id%></a></td>
      <td><%=fLeft(vProg_Title1, 32)%></td>
      <td style="text-align: center"><%=vProg_Owner%></td>
      <td style="text-align: center"><%=fIf(vProg_Retired,"Y","N")%></td>
      <td style="text-align: center; white-space: nowrap;">
        <form method="POST" action="Program.asp" name="fClone" onsubmit="return validate('.clone_<%=i%>')">
          <input type="text" name="vCloneId" id="vCloneId" class="clone_<%=i%>" size="8">
          <input type="submit" value="Clone" name="bClone" class="button070">
          <a title="Clone using this Program Id ..." class="debug" onclick="$('.clone_<%=i%>')[0].value='<%=vNextId%>';" href="#">&#937;</a>
          <input type="hidden" name="vRange" value="<%=vRange%>">
          <input type="hidden" name="vLingo" value="<%=vLingo%>">
          <input type="hidden" name="vFunction" value="clone">
          <input type="hidden" name="vProgId" value="<%=vProg_Id%>">
        </form>
      </td>
      <td style="text-align: center">
        <form method="POST" id="programCert" action="Program.asp" name="fCert" target="_blank">
          <input type="text" name="vCust" size="4" value="<%=Left(svCustId, 4)%>">
          <input type="text" name="vAcct" size="4" value="<%=svCustAcctId%>">
          <input type="submit" value="Cert" name="bCert" class="button070">
          <input type="hidden" name="vRange" value="<%=vRange%>">
          <input type="hidden" name="vLingo" value="<%=vLingo%>">
          <input type="hidden" name="vFunction" value="cert">
          <input type="hidden" name="vProgId" value="<%=vProg_Id%>">
        </form>
      </td>
    </tr>
    <%  
        oRsBase.MoveNext
      Loop
      Set oRsBase = Nothing
      sCloseDbBase    
    %>
  </table>

  <div style="text-align: center">
    <br /><br />
    <% If i = 0 Then %>
    <h5>No Programs match your selection criteria.</h5>
    <% End If %>
    <% If i = 50 Then %>
    <input type="button" onclick="location.href = 'Programs.asp?vRange=<%=Left(vProg_Id, Len(vProg_Id) - 2)%>'" value="Next" name="bNext" class="button070">
    <% End If %>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


