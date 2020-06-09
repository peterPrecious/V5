<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  Dim bTest 
  bTest = false

  If bTest Then
    If Request.Form.Count > 0 Then 
      For Each vFld in Request.Form
        Response.Write vFld & " - " & Request(vFld) & "<br>"
      Next  
    End If
  End If
%>

<html>

  <head>
    <meta charset="UTF-8">
    <script src="/V5/Inc/jQuery.js"></script>
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>
    function submitForm(theForm) {
      var txtAmount, innerHtml, bTest = <%=Lcase(bTest)%>;
      txtAmount = document.getElementById('txtAmount').value;
      if (txtAmount.length == 0) {
        alert('Please enter an amount, ie 5500');
        return (false);
      }
      innerHtml = '<input type="hidden" name="Products" value="' + txtAmount + '::1::Misc::Miscellaneous Payment (<%=svCustId%>)::{US}">';
      if (bTest) {
        alert(txtAmount);
        alert(innerHtml);
      }
      document.getElementById('divAmount').innerHTML = innerHtml;
      document.getElementById(theForm).submit();
    }
  </script>
  <title>Ecommerce Memo Posting</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>

  <form action="https://secure.internetsecure.com/process.cgi" method="POST" target="_top" id="theForm">
    <input type="hidden" name="language" 				value="English">
    <input type="hidden" name="MerchantNumber"  value="4089">
    <input type="hidden" name="xxxCompany"      value="(Customer ID: <%=svCustId%> - please do not remove)">
    <input type="hidden" name="ReturnCGI"       value="/V5/Code/EcomMemo.asp">
    <div id="divAmount"></div>
    <div align="center">
    <table border="0" width="600" bordercolor="#111111" cellpadding="10" cellspacing="0">
      <tr>
        <td align="center">
          <h1>Credit Card Memo Payment</h1>
          <p align="left">Please enter the Amount you wish to pay Vubiz (ie 8000 for eight thousand dollars) then click Submit.&nbsp;You will be transferred to &quot;InternetSecure&quot; where you can enter the Credit Card details.<br>Note: you will not be returned to this site after the transaction is complete. An email will be sent by InternetSecure showing the transaction.</p>
          <input type="text" name="txtAmount" id="txtAmount" size="11"> 
          <input type="button" onclick="submitForm('theForm')" value="Submit" name="bContinue" class="button"> <br><br>
          Please email <a href="mailto:heggleston@vubiz.com">Helen Eggleston</a> at Vubiz when the transaction is complete to confirm payment. 
        </td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->


</body>

</html>