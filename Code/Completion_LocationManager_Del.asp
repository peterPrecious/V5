<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
  vUnit_No = Request("vUnit_No")
  sGetUnit vUnit_No
  
  '...get the region values from the selected list before edditing
  '   <option value='0002|Distribution Centre|0'">0002 (Distribution Centre)</option>
  If Request.QueryString("vDel") = "y" Then
    sDeleteUnit vUnit_NO
    vMsg = Session("Completion_L0tit") & " : " & vUnit_L1 & "|" & vUnit_L0  & " was deleted successfully."
  End If
 
%>

<html>

<head>
  <title>Completion_LocationManager_Add</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--#include file = "Completion_LocationManager_Top.asp"-->

  <div style="margin-bottom: 30px;">
    <h1>Delete <%=Session("Completion_L0Tit") & ":&ensp;&ensp;" & vUnit_L1Title & " | " & vUnit_L0Title & " (" & vUnit_L1 & "|" & vUnit_L0 & ")"%></h1>
    <h2>Click Confirm to delete this <%=Session("Completion_L0Tit")%></h2>
  </div>

  <%
    If(Len(vMsg) > 0) Then 
      Response.Write "<h5>" & vMsg & "</h5>"
    Else
  %>
  <input type="button" onclick="location.href = 'Completion_LocationManager_Del.asp?vDel=y&vUnit_No=<%=vUnit_No%>'" value="Confirm" name="bDelete" class="button070">
  <%
    End If     
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>


