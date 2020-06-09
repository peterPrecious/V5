<%@  codepage="65001" %>

<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<html>

<head>
  <title>RTE_Ecom</title>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <%
    Dim vMembNo, vProgId, vProgTitle, vPurchaser, vPurchased, vExpires

		vMembNo = (Request("vMembNo")) 
		vProgId = (Request("vProgId")) 
    vProgTitle = fProgTitle (vProgId)

    vSql = "SELECT "_
    		 & "  Ec.Ecom_Issued, "_
    		 & "  Ec.Ecom_Expires, "_
    		 & "  Ec.Ecom_CardName, "_
    		 & "  Ec.Ecom_FirstName, "_
    		 & "  Ec.Ecom_LastName "_ 
    		 & "FROM "_
    		 & "  V5_Vubz.dbo.Ecom AS Ec WITH (NOLOCK) "_
    		 & " WHERE "_
         & "  (Ec.Ecom_MembNo = " & vMembNo & ") AND "_
         & "  (Ec.Ecom_Programs = '" & vProgId & "') "

'   sDebug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql) 

  %>

  <fieldset>
    <legend><!--[[-->Purchase Details<!--]]--></legend>

    <h1><%=vProgTitle %></h1>
    <h2><!--[[-->Program Id<!--]]-->: <%=vProgId %></h2>

    <table class="table">
      <tr>
        <th class="rowshade" style="width: 50%; text-align: left;">Purchaser</th>
        <th class="rowshade" style="width: 25%; text-align: center;">Purchased</th>
        <th class="rowshade" style="width: 25%; text-align: center;">Expires</th>
      </tr>
      <%
        Do While Not oRs2.Eof
    	    vPurchaser = fIf(Len(oRs2("Ecom_CardName")) = 0, oRs2("Ecom_FirstName") & " " & oRs2("Ecom_LastName"), oRs2("Ecom_CardName"))
    	    vPurchased = oRs2("Ecom_Issued")  
    	    vExpires   = oRs2("Ecom_Expires")
      %>
      <tr><td colspan="3">&nbsp;</td></tr>
      <tr>
        <td style="text-align: left; vertical-align:bottom;"><%=vPurchaser%></td>
        <td style="text-align: center"><%=fFormatDate(vPurchased)%></td>
        <td style="text-align: center"><%=fFormatDate(vExpires)%></td>
      </tr>
      <% 
          oRs2.MoveNext
        Loop
      %>
    </table>

    <br><br>
  </fieldset>

</body>
</html>
