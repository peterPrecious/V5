<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prod.asp"-->

<% 
  Session("Ecom_Media") = "Prods"
  Dim vCnt, vBg, vUrl
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>


  <title>Vubiz Product Catalogue</title>
  <base target="_self">
</head>

<body>

  <% Server.Execute vShellHi %>

  <table border="0" style="border-collapse: collapse" width="100%" id="table3" cellpadding="3">
    <tr>
      <td><img border="0" src="../Images/Ecom/Categories.gif"></td>
      <td width="0" align="center"><h1>Categories</h1><h2>Click on any of the titles below and a list of products available for purchase will appear on the right.</h2></td>
    </tr>
  </table>

  <table cellspacing="0" cellpadding="3" border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" id="table2">
    <%
      '...get selected product groups
      sGetProdLeft_Rs fIf(Request("vProdSpecials")="n","00000000", svCustId)
      vCnt = 0
      Do While Not oRs.Eof
        vCnt = vCnt + 1
        vBg = "" : If vCnt Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color ever other line        
        vUrl = "Ecom2ProdsRight.asp?vProd_Id=" & oRs("ProdId")
        '...send first catalogue item to right frame
        If vCnt = 1 Then Response.Write "<script>{parent.frames.Right.location.href='" & vUrl & "';}</script>"
    %> 
    <tr>
      <td align="left" valign="top" width="90%" class="c1" <%=vbg%>><p class="c2"><a <%=fStatX%> target="Right" href="<%=vUrl%>"><%=oRs("Prod_CatTitle")%></a></p>
      </td>
    </tr>
    <%
        oRs.MoveNext
      Loop
      sCloseDb
    %>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>