<!--#include virtual = "V5/Inc/Setup.asp"-->

<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<!--#include virtual = "V5/Inc/Elavon.asp"-->

<%
  Dim orderNo, vMsg, vSource, orderCnt

  If Len(Request("orderNo")) > 0 Then
 
    orderNo = Request("orderNo")
    If fPureInt(orderNo) = 0 Then 
      vMsg = "'" & orderNo & "' is not a valid Order No!"
    Else
      sGetSqlForm(orderNo)
      If vEcom_Eof Then
        vMsg = "That order was not processed by Elavon."
      Else
        orderCnt = sp5elavonToEcom(orderNo)
        If orderCnt > 0 Then
          vMsg = "That order is already in the Ecom system."
        Else
          vMsg = "That " & vEcom_Media & " order below is NOT in the Ecom system."
        End If   
      End If
    End If 
  Else
    vMsg = "Enter the Order No, ie 12345"
  End If 

%>

<html>

<head>
  <title>EcomEdit</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <style>
    form { margin: auto; width: 400px; padding: 20px; text-align: center; }
    div { vertical-align: baseline; display: inline-block; }
  </style>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>(Re)Post an Elavon Transaction</h1>
  <h3>If a Elavon transaction does not get into our Ecom system,<br />typically because the Customer did not return properly, <br />this app will complete the transaction, <br />ie post the data approved by Elavon into the Vubiz Ecom Database.</h3>

  <form method="POST" action="ElavonRePost.asp">
    <div>Order No: </div>
    <div><input class="c3" type="text" name="orderNo" id="orderNo" value="<%=orderNo%>"></div>
    <div><input type="submit" value="Lookup" name="postOrder" class="button070"></div>
  </form>

  <h5><%=vMsg %></h5>

  <table>
    <tr><th>vEcom_CustId : </th><td><%=vEcom_CustId %></td></tr>
    <tr><th>vEcom_AcctId : </th><td><%=vEcom_AcctId %></td></tr>
    <tr><th>vEcom_Agent : </th><td><%=vEcom_Agent %></td></tr>
    <tr><th>vEcom_Id : </th><td><%=vEcom_Id %></td></tr>
    <tr><th>vEcom_CatlNo : </th><td><%=vEcom_CatlNo %></td></tr>
    <tr><th>vEcom_Programs : </th><td><%=vEcom_Programs %></td></tr>
    <tr><th>vEcom_Prices : </th><td><%=vEcom_Prices %></td></tr>
    <tr><th>vEcom_Taxes : </th><td><%=vEcom_Taxes %></td></tr>
    <tr><th>vEcom_Expires : </th><td><%=vEcom_Expires %></td></tr>
    <tr><th>vEcom_Amount : </th><td><%=vEcom_Amount %></td></tr>
    <tr><th>vEcom_Currency : </th><td><%=vEcom_Currency %></td></tr>
    <tr><th>vEcom_Lang : </th><td><%=vEcom_Lang %></td></tr>
    <tr><th>vEcom_Quantity : </th><td><%=vEcom_Quantity %></td></tr>
    <tr><th>vEcom_Media : </th><td><%=vEcom_Media %></td></tr>
    <tr><th>vEcom_Source : </th><td><%=vEcom_Source %></td></tr>
    <tr><th>vEcom_CardName : </th><td><%=vEcom_CardName %></td></tr>
    <tr><th>vEcom_Address : </th><td><%=vEcom_Address %></td></tr>
    <tr><th>vEcom_City : </th><td><%=vEcom_City %></td></tr>
    <tr><th>vEcom_Postal : </th><td><%=vEcom_Postal %></td></tr>
    <tr><th>vEcom_Province : </th><td><%=vEcom_Province %></td></tr>
    <tr><th>vEcom_Country : </th><td><%=vEcom_Country %></td></tr>
    <tr><th>vEcom_Phone : </th><td><%=vEcom_Phone %></td></tr>
    <tr><th>vMemb_Organization : </th><td><%=vMemb_Organization %></td></tr>
    <tr><th>vMemb_FirstName : </th><td><%=vMemb_FirstName %></td></tr>
    <tr><th>vMemb_LastName : </th><td><%=vMemb_LastName %></td></tr>
    <tr><th>vMemb_Email : </th><td><%=vMemb_Email %></td></tr>
    <tr><th>vEcom_FirstName : </th><td><%=vEcom_FirstName %></td></tr>
    <tr><th>vEcom_LastName : </th><td><%=vEcom_LastName %></td></tr>
    <tr><th>vEcom_Email : </th><td><%=vEcom_Email %></td></tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
