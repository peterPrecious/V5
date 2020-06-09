<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
  Dim vMsg

  If Request("vCust_Expires").Count = 1 Then
    vCust_Expires = Trim(Request("vCust_Expires"))
    If Len(Trim(fFormatSqlDate(vCust_Expires))) > 0 Then
      sUpdateCustExpires svCustAcctId, vCust_Expires
      vMsg = "Updated Successfully"
    Else
      vMsg = vCust_Expires & " is an invalid expiry date."
      vCust_Expires = fCustExpires
    End If
  Else  
    vMsg = ""
    vCust_Expires = fCustExpires
  End If
%>

<html>

<head>
  <title>CustomerExpires</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
    function validate(theForm)
    {

      if (theForm.vCust_Expires.value == "")
      {
        alert("Please enter a value for the \"Customer Expiry\" field.");
        theForm.vCust_Expires.focus();
        return (false);
      }

      if (theForm.vCust_Expires.value.length < 11)
      {
        alert("Please enter at least 11 characters in the \"Customer Expiry\" field.");
        theForm.vCust_Expires.focus();
        return (false);
      }

      if (theForm.vCust_Expires.value.length > 12)
      {
        alert("Please enter at most 12 characters in the \"Customer Expiry\" field.");
        theForm.vCust_Expires.focus();
        return (false);
      }
      return (true);
    }
  </script>
</head>

<body>

  <% 
    Server.Execute vShellHi
  %>

  <h1>Customer Expiry Date</h1>
  <h2>Edit/Modify the Expiry Date then click Update.</h2>
  <p>&nbsp;</p>
  <% If Len(vMsg) > 0 Then %><h5><%=vMsg%></h5>
  <% End If %>


  <form method="POST" action="CustomerExpires.asp" onsubmit="return validate(this)">
    <table class="table">
      <tr>
        <th style="width: 50%"><%=svCustId%> Expiry Date :</th>
        <td style="width: 50%">
          <input type="text" size="14" name="vCust_Expires" value="<%=fFormatDate(vCust_Expires)%>" maxlength="12">
        </td>
      </tr>
      <tr>
        <td colspan="2" style="text-align:center;"><br />Use English date format only, ie Jan 15, 2008. <br />If no valid date appears then there is no expiry date for this Account.</td>
      </tr>
    </table>
    <div style="margin: 30px; text-align:center;">
      <input type="submit" value="Update" name="bUpdate" class="button100">
    </div>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->


</body>

</html>
