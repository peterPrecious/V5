<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prod.asp"-->

<% 
  If Len(Request("vProd_Id")) > 0 Then Session("Id")  = Request("vProd_Id")
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script Language="JavaScript">
    function Validate(theForm)
    {    
      if (theForm.vQty.value == "")
      {
        alert("Please enter a value for the \"Quantity\" field.");
        theForm.vQty.focus();
        return (false);
      }
    
      if (theForm.vQty.value.length < 1)
      {
        alert("Please enter at least 1 characters in the \"Quantity\" field.");
        theForm.vQty.focus();
        return (false);
      }
    
      if (theForm.vQty.value.length > 3)
      {
        alert("Please enter at most 3 characters in the \"Quantity\" field.");
        theForm.vQty.focus();
        return (false);
      }
    
      var checkOK = "0123456789-";
      var checkStr = theForm.vQty.value;
      var allValid = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0;  i < checkStr.length;  i++)
      {
        ch = checkStr.charAt(i);
        for (j = 0;  j < checkOK.length;  j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length)
        {
          allValid = false;
          break;
        }
        allNum += ch;
      }
      if (!allValid)
      {
        alert("Please enter only digit characters in the \"Quantity\" field.");
        theForm.vQty.focus();
        return (false);
      }
    
      var chkVal = allNum;
      var prsVal = parseInt(allNum);
      if (chkVal != "" && !(prsVal >= "1" && prsVal <= "999"))
      {
        alert("Please enter a value greater than or equal to \"1\" and less than or equal to \"999\" in the \"Quantity\" field.");
        theForm.vQty.focus();
        return (false);
      }
      return (true);
    }
  </script>
  
    
    
    
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>

  <table border="0" style="border-collapse: collapse" width="100%" id="table1" cellpadding="3">
    <tr>
      <td nowrap valign="top">&nbsp;&nbsp; <img border="0" src="../Images/Ecom/Book.jpg" width="80" height="73"></td>
      <td align="center"><h1>Product List</h1><h2 align="left">For details on any product, click on the Title. To purchase an item, enter Quantity then click &quot;<b>Add</b>&quot; and it will added to your basket. </h2></td>
    </tr>
  </table>

  <table border="1" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%" cellspacing="0" cellpadding="3">
    <tr>
      <th align="left" height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" rowspan="2">Title</th>
      <th align="left" height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" rowspan="2">Quantity</th>
      <th align="middle" bgcolor="#DDEEF9" height="10" bordercolor="#FFFFFF" colspan="2">Cost per Item</th>
      <th align="middle" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF" rowspan="2">Add to basket</th>
    </tr>
    <tr>
      <th align="right" bgcolor="#DDEEF9" height="10" bordercolor="#FFFFFF">$US&nbsp; </th>
      <th align="right" bgcolor="#DDEEF9" height="10" bordercolor="#FFFFFF">$CA&nbsp; </th>
    </tr>
    <%
      Dim vCnt, vBg
      '...get selected product groups
      sGetProdRight_Rs Session("Id")
      vCnt = 0
      Do While Not oRs.Eof
        sReadProd
        vCnt = vCnt + 1
        vBg = "" : If vCnt Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color every other line       
   %>
    <form action="Ecom2ProdsBasket.asp" method="POST" target="Right" onsubmit="return Validate(this)" name="Form<%=vCnt%>">
      <tr>
        <td valign="top" <%=vbg%>><p class="c2"><a <%=fStatX%> href="Ecom2ProdsDesc.asp?vProdId=<%=vProd_Id%>"><%=vProd_Title%></a></p>
        </td>
        <td valign="top" <%=vbg%>><input type="text" name="vQty" size="2" value="1" maxlength="3"></td>
        <td valign="top" align="right" <%=vbg%>><%=FormatNumber(vProd_Price * vCurrency, 2)%></td>
        <td valign="top" align="right" <%=vbg%>><%=FormatNumber(vProd_Price, 2)%></td>
        <td valign="top" align="middle" <%=vbg%>><input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="bAdd" type="image"></td>
      </tr>
      <input type="hidden" name="vOrder" value="<%=vProd_Id & "~" & vProd_Price * vCurrency & "~" & vProd_Price & "~" & fHtmlUnquote(vProd_Title)%>">
    </form>
    <%
        oRs.MoveNext
      Loop
      sCloseDb
    %> 
    <tr>
      <td valign="top" colspan="5" align="center"><p><br>
      <!--[[-->Applicable taxes extra for Canadian orders.<!--]]-->&nbsp; <br><a <%=fStatX%> href="javascript:SiteWindow('Ecom2PoliciesProds.asp?vClose=Y')">
      <!--[[-->Shipping and handling charges extra.<!--]]--></a><br>&nbsp;</p>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

