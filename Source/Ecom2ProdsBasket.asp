<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Discounts.asp"-->


<%
  Dim aValues, aValue, vOnFile, vCnt, vBg, vQty, vOk, vStr, vEnd

  '...retrieve/create the basket info
  svProdNo       = Session("ProdNo")
  svProdMax      = Session("ProdMax")
  If svProdNo > 0 Then  
    saProd       = Session("Prod")
  Else
    Dim saProd()
    ReDim Preserve saProd (6, svProdNo)
  End If

  '...see if clearing basket
  If Request("vAction") = "Clear" Then
    svProdNo  = 0
    svProdMax = 0
    ReDim Preserve saProd (6, svProdNo)
  End If

  '...determine if any discounts apply from the customer file
  sGetCust svCustId

  '...add the selected item into the array
  If Request("vOrder").Count > 0 Then
    '...grab the quantity which must accompany vOrder
    vQty            = Request("vQty")
    vValue          = Replace(Request("vOrder"),"'"," ")
    aValues         = Split (vValue,"||")
    For i = 0 To Ubound(aValues)
      aValue  = Split (aValues(i),"~")

      '...see if each product is in basket
      vOnFile = False

      If svProdMax > 0 Then
        For j = 1 to svProdMax
          If saProd(1, j) = aValue(0) Then 
            saProd(2, j) = vQty                               '...quantity
            vOnFile = True
            Exit For
          End If
        Next 
      End If
  
      '...if not on file, then add item to bottom of array
      If Not vOnFile Then
        svProdNo            = svProdNo + 1
        svProdMax           = svProdMax + 1
        ReDim Preserve saProd (6, svProdNo)
        saProd(0, svProdNo) = 0                              '...percentage discount
        saProd(1, svProdNo) = aValue(0)                      '...product Id
        saProd(2, svProdNo) = vQty                           '...quantity
        saProd(3, svProdNo) = aValue(1)                      '...US price each
        saProd(4, svProdNo) = aValue(2)                      '...CA price each
        saProd(5, svProdNo) = 0                              '...no days duration
        saProd(6, svProdNo) = vQty & " of <b>" & aValue(3) & "</b>"  '...title

        '...save basket values
        Session("ProdNo")   = svProdNo
        Session("ProdMax")  = svProdMax
        Session("Prod")     = saProd       
      End If
   Next

  End If


  '...put any discounts into cell 0  
  sCheckDiscounts

 
  '...save basket values
  Session("ProdNo")  = svProdNo
  Session("ProdMax") = svProdMax
  Session("Prod")    = saProd       
 
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <title>Ecommerce Product Basket</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" style="border-collapse: collapse" width="100%" cellpadding="3">
    <tr>
      <td valign="top"><img border="0" src="../Images/Ecom/Basket.gif">&nbsp;&nbsp; </td>
      <td align="center"><h1>My Basket </h1><h2 align="left">When you have completed your selection(s) click &quot;Continue&quot; to proceed with your e-commerce purchase. Click &quot;Return&quot; to return to the catalogue and make further selections to add to your basket. To remove all selections from your basket, click &quot;clear&quot;.</h2>
       <h1 align="center"><a <%=fStatX%> href="Ecom2ProdsRight.asp?vProdId=<%=Session("Id")%>">Return to Product List</a></h1>
      </td>
    </tr>
  </table>
  <table border="1" width="100%" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3">
    <tr>
      <th align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF" rowspan="2" nowrap>Product Id</th>
      <th bgcolor="#DDEEF9" align="left" bordercolor="#FFFFFF" rowspan="2">Title</th>
      <th bgcolor="#DDEEF9" colspan="2" bordercolor="#FFFFFF" nowrap>Unit Cost</th>
    </tr>
    <tr>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="right" nowrap>$US&nbsp; </th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="right" nowrap>$CA&nbsp; </th>
    </tr>
    <%
      vCnt = 0
      For i = 1 to svProdMax
        If saProd(2, i) > 0 Then
          vCnt = vCnt + 1
          vBg = "" : If vCnt Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color every other line       

    %>
      <tr>
        <td class="c2" valign="top"><font color="#008000"><%=saProd(1, i)%></font></td>
<!--    <td class="c2" valign="top"><font color="#008000"><%=saProd(6, i)%></font></td>  -->
        <td class="c2" valign="top"><font color="#008000"><%=saProd(6, i) & fIf(saProd(0, i) > 0, "<br>[Note: " & saProd(0, i) & "% discount applied]", "") %></font></td>


        <td nowrap class="c2" align="right"><font color="#008000"><%=FormatNumber(saProd(3, i) * (1 - saProd(0, i) / 100) ,2)%></font></td>
        <td nowrap class="c2" align="right"><font color="#008000"><%=FormatNumber(saProd(4, i) * (1 - saProd(0, i) / 100) ,2)%></font></td>
      </tr>
    <%
          End If
        Next
  
        '...nothing in the basket?  empty basket in case lines were removed, thus still there with qty = 0
        If vCnt = 0 Then
          svProdNo  = 0
          svProdMax = 0
          ReDim Preserve saProd (6, svProdNo)
          '...save basket values
          Session("ProdNo")    = svProdNo
          Session("ProdMax")   = svProdMax
          Session("Prod")      = saProd       
      %> 


    <tr>
      <td valign="top" colspan="4" align="center">
      <h6>No items have been selected, please click &quot;Return&quot;.</h6>
      <p><a <%=fStatX%> href="Ecom2ProdsRight.asp"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif" alt="Click to return to the Program list"></a></p>
      </td>
    </tr>
    <% 
        Else
      %> <tr>
      <td valign="top" colspan="4" align="center"><br> <a <%=fStatX%> href="Ecom2ProdsRight.asp"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif" alt="Click to return to the Program list"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="Ecom2ProdsBasket.asp?vAction=Clear"><img border="0" src="../Images/Buttons/Clear_<%=svLang%>.gif" alt="Click to remove ALL items"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="Ecom2Customer.asp" target="main"><img border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" alt="Click to continue to the next screen"></a><br>&nbsp;</td>
    </tr>
    <% 
        End If
      %>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>