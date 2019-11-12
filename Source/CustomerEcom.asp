<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <form>
    <table border="1" width="100%" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3">
      <tr>
        <td colspan="2" align="center" valign="Top" width="100%"><h1>&nbsp;</h1><h1>Ecommerce Revenue Splits</h1><p>&nbsp;</p>
        </td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="Top">Cust Id : </th>
        <td width="70%" valign="Top"><h1><b>IBAO2314</b></h1></td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="50%">% Split if via this channel : </th>
        <td valign="Top" width="50%"><input type="text" name="T8" size="6" value="40">%&nbsp; <select size="1" name="D8">
        <option>IBAO2314</option>
        </select><br><input type="text" name="T9" size="6" value="10">%&nbsp; <select size="1" name="D9">
        <option>Larry Hopperton</option>
        </select><br><input type="text" name="T10" size="6" value="20">%&nbsp; <select size="1" name="D10">
        <option>Rosemary Rapino</option>
        </select><br><input type="text" name="T11" size="6" value="30">%&nbsp; <select size="1" name="D11">
        <option>IBAO2314</option>
        <option selected>Vubix</option>
        </select><br><input type="text" name="T12" size="6" value="0">%&nbsp; <select size="1" name="D12">
        <option>Select</option>
        </select><br><input type="text" name="T13" size="6" value="0">%&nbsp; <select size="1" name="D13">
        <option>Select</option>
        </select><br><input type="text" name="T14" size="6" value="0">%&nbsp; <select size="1" name="D14">
        <option>Select</option>
        </select><br>Note % must add up to 100<br>&nbsp;</td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="50%">% Split if via other channels : </th>
        <td valign="Top" width="50%"><input type="text" name="T1" size="6" value="20">%&nbsp; <select size="1" name="D1">
        <option>IBAO2314</option>
        </select><br><input type="text" name="T2" size="6" value="10">%&nbsp; <select size="1" name="D2">
        <option>Larry Hopperton</option>
        </select><br><input type="text" name="T3" size="6" value="20">%&nbsp; <select size="1" name="D3">
        <option>Rosemary Rapino</option>
        </select><br><input type="text" name="T4" size="6" value="50">%&nbsp; <select size="1" name="D4">
        <option>IBAO2314</option>
        <option selected>Vubix</option>
        </select><br><input type="text" name="T5" size="6" value="0">%&nbsp; <select size="1" name="D5">
        <option>Select</option>
        </select><br><input type="text" name="T6" size="6" value="0">%&nbsp; <select size="1" name="D6">
        <option>Select</option>
        </select><br><input type="text" name="T7" size="6" value="0">%&nbsp; <select size="1" name="D7">
        <option>Select</option>
        </select><br>Note % must add up to 100<br>&nbsp;</td>
      </tr>
      <tr>

<!--        <td colspan="2" align="center" valign="Top" width="100%">&nbsp;<p><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="javascript:jconfirm('CustomerEdit.asp?vDelCustId=<%=vCust_Id%>&vDelCustAcctId=<%=vCust_AcctId%>&vFunction=del', 'Ok to delete this customer and all supporting files?')"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a>&nbsp; </p>-->

        <td colspan="2" align="center" valign="Top" width="100%">&nbsp;<p><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="javascript:jconfirm('Customer.asp?vDelCustId=<%=vCust_Id%>&vDelCustAcctId=<%=vCust_AcctId%>&vFunction=del', 'Ok to delete this customer and all supporting files?')"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a>&nbsp; </p>


        <p>&nbsp;</p>
        </td>
      </tr>
    </table>
  </form>

</body>

</html>