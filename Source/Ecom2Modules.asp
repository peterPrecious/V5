<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
    Session("Ecom_Prog") = Request("vProgId")
    Session("Ecom_Mods") = ""

    sGetProg Session("Ecom_Prog")
%>

<html>

<head>
  <title>Ecom2Modules</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <base target="_self">
</head>

<body>

  <% Server.Execute vShellHi %>

  <table class="table">
    <tr>
      <td>
        <h1><!--[[-->Learning Modules<!--]]--></h1>
        <p><!--[[-->These are our basic elements of learning.&nbsp; This is the module set contained in the program you selected.&nbsp; For details of any Module, click on the module title.<!--]]--></p>
        <h2 style="text-align: center; margin-bottom:20px;"><a <%=fStatX%> href="javascript:history.back(1)"><!--[[-->Return to Programs<!--]]--></a></h2>
      </td>
    </tr>
  </table>


  <table class="table">
    <tr>
      <td>
        <h2><%=vProg_Title%></h2>
        <p><%=vProg_Desc%></p>
        <h3><!--[[-->Estimated program length<!--]]-->: <%=vProg_Length%>&nbsp;<!--[[-->Hour(s)<!--]]-->.</h3>
        <p>
      </td>
    </tr>
    <%
      Dim aMods, vBg
      aMods = Split(Trim(vProg_Mods), " ")
      For i = 0 To Ubound(aMods)
        sGetMods aMods(i)
				If vMods_Active Then 
					vBg = "" : If i Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color ever other line        
    %>
    <tr>
      <td <%=vBg%>>
          <a <%=fStatX%> href="Ecom2Module.asp?vModsId=<%=vMods_Id%>"><%=vMods_Title%></a><%=vMods_Features%>
      </td>
    </tr>
    <%  
				End If
      Next

      '...exam included?
      If vMods_Active Then 
				If Len(vProg_Assessment) > 1 Then  
					vBg = "" : If i Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"
    %>
    <tr>
      <td <%=vBg%>>
        <p class="c1"><!--[[-->An Examination is available with this program.<!--]]--></p>
      </td>
    </tr>
    <% 
				End If
      End If

      If vProg_Cert <> 0 Then  
        vBg = "" : If i Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"
    %>
    <tr>
      <td <%=vBg%>>
        <p class="c1"><!--[[-->A Certificate of Completion is available with this program.<!--]]--></p>
      </td>
    </tr>
    <%
      End If
    %>
  </table>

  <p align="center">
    <a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>
  </p>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>
</html>
