<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <%
  Dim vModId, vStr, aQue, vQue, aAns, vAns, vCheck, vChecked, vFormOK
  Const cAlpha = "abcdefg" 

  vModId = Request("vModId")
  vStr = GetStr (vModId)
' Response.write "<P><P>Str: " & vStr
  aQue = split(vStr,"~~"): vQue = ubound(aQue)
' response.write "<P><P>No Que: " & vQue

  Server.Execute vShellHi 
  
  %>

  <form method="POST" action="TestEdit.asp">
    <h1 align="left">Self Assessment - <%=vModId%></h1>
    <table border="1" width="100%" cellspacing="1" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <%  
      vFormOK = False
      For i = 0 To vQue - 1  
        aAns = split(aQue(i),"||"): vAns = Ubound(aAns) 
        If Len(aAns(0)) > 0 Then
          vFormOK = true
      %>
      <tr>
        <td bgcolor="#DDEEF9" valign="top" width="30">&nbsp;<%=i+1%>.</td>
        <td bgcolor="#DDEEF9" valign="top"><b><%=aAns(0)%></b> </td>
      </tr>
      <%
          For j = 2 To vAns
            vCheck = 1
            On Error Resume Next
            vCheck =cInt(aAns(1))
            If Cint(j-1) = vCheck Then vChecked = " Checked" Else vChecked = ""
            If Len(aAns(j)) > 0 then  
      %>
      <tr>
        <td valign="top" width="30" align="right"><input name="Q<%=right("00" & i+1,2)%>" type="radio" value="<%=j+1%>" <%=vchecked%>></td>
        <td valign="top"><%=aAns(j)%> </td>
      </tr>
      <%  
            End If
          Next
        End If
      Next
    %>
    </table>
    <% If Not vFormOK Then %> <p align="center"><br>There are no questions on file for this self assessment!</p><% End If %> <center><p><input border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif" name="I1" type="image"><br><br><a href="TestMenu.asp">Self Assessment Menu</a></p><input type="hidden" name="vModId" value="<%=vModId%>"></center>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
