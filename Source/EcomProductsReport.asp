<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<html>

  <head>
    <meta charset="UTF-8">
    <script src="/V5/Inc/jQuery.js"></script>
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi 
    
    If Len(Request("vEcom_No")) = 0 Then
  %>
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td colspan="2" align="center"><h1>Ecommerce Products Sales</h1><h2>To access a Products Sale, click on the Order No below.&nbsp; Note: this list the latest 50 Product orders.</h2></td>
    </tr>
    <tr>
      <th height="30" nowrap align="left">Order No</th>
      <th height="30" nowrap align="left">Customer</th>
    </tr>
    <%
      '...read Ecom  
      Dim vOrderNo, vNoCds, vNoRecs
      vOrderNo = ""  
      vNoRecs = 0
      vSql = "SELECT Ecom_OrderNo, Ecom_No, Ecom_FirstName, Ecom_LastName FROM Ecom WITH (nolock) WHERE (Ecom_Media = 'CDs') OR (Ecom_Media = 'Prods') ORDER BY Ecom_OrderNo DESC"
      sOpenDb
      Set oRs = oDb.Execute(vSQL)    
      Do While Not oRs.Eof     
        If oRs("Ecom_OrderNo") <> vOrderNo Then '...only display first line of multi program order
    %> <tr>
      <td><a <%=fStatX%> href="EcomProductsReport.asp?vEcom_No=<%=oRs("Ecom_No")%>"><%=oRs("Ecom_OrderNo")%></a>&nbsp; </td>
      <td><%=oRs("Ecom_FirstName") & " " & oRs("Ecom_LastName")%>&nbsp; </td>
    </tr>
    <%  
          vNoRecs = vNoRecs + 1
          If vNoRecs = 50 Then Exit Do
          vOrderNo = oRs("Ecom_OrderNo")
        End If
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb    
    %> <tr>
      <td colspan="2" align="center"><br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><p>&nbsp;</p>
      </td>
    </tr>
  </table>
  <%
  Else

    '...get the values (even if trying to add)
    Dim vAddress, vLabel
    vAddress = ""
    vLabel = ""
    
    vSql = "SELECT * FROM Ecom WITH (nolock) WHERE Ecom_No = " & Request("vEcom_No")
    sOpenDb
    Set oRs = oDb.Execute(vSQL)    
    sReadEcom

    vAddress = vEcom_FirstName & " " & vEcom_LastName & " (" & vEcom_OrderNo & ")<br>"
    vAddress = vAddress & vEcom_Address & "<br>"
    vAddress = vAddress & vEcom_City & ", " & vEcom_Postal & ", " & vEcom_Province & ", " & vEcom_Country & "<br>"
    vAddress = vAddress & "(" & vEcom_Phone & ")<br>"
    
    If Len(vEcom_Label) > 0 Then
      vLabel   = Replace(vEcom_Label, vbCrLf, "<br>" )
      If Right(vLabel, 4) = "<br>" Then vLabel = Left(vLabel, Len(vLabel)-4)
      If Right(vLabel, 4) = "<br>" Then vLabel = Left(vLabel, Len(vLabel)-4)
      If Right(vLabel, 4) = "<br>" Then vLabel = Left(vLabel, Len(vLabel)-4)
      vLabel   = vLabel & "<br>(" & vEcom_Phone & ")<br>"
    End If

%>
  <table border="1" width="100%" cellspacing="3" cellpadding="2" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <th width="30%">Order Address : </th>
      <td width="70%"><%=vAddress%></td>
    </tr>
    <tr>
      <th width="30%">Shipping Label : </th>
      <td width="70%"><%=vLabel%> </td>
    </tr>
    <tr>
      <td align="center" width="100%" valign="top" colspan="2" height="38"><br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><p>&nbsp;</p>
      </td>
    </tr>
  </table>

<%
  End If  

  Server.Execute vShellLo 
  
%>

</body>

</html>