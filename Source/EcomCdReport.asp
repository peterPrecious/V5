<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %> <%   
  If Len(Request("vEcom_No")) = 0 Then
%>
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td colspan="2" align="center"><h1>Ecommerce CD Sales</h1><h2 align="left">To Access a CD Sale, click on the Order No below.&nbsp; Note: this list the latest 50 CD orders.</h2></td>
    </tr>
    <tr>
      <td><font face="Verdana" size="1"><b>Order No</b></font></td>
      <td><font face="Verdana" size="1"><b>Customer</b></font></td>
    </tr>
    <%
    '...read Ecom  
    Dim vOrderNo, vNoCds, vNoRecs
    vOrderNo = ""  
    vNoRecs = 0
    vSql = "SELECT Ecom_OrderNo, Ecom_No, Ecom_FirstName, Ecom_LastName FROM Ecom WITH (nolock) WHERE (Ecom_Media = 'CDs') ORDER BY Ecom_OrderNo DESC"
    sOpenDb
    Set oRs = oDb.Execute(vSQL)    
    Do While Not oRs.Eof     
      If oRs("Ecom_OrderNo") <> vOrderNo Then '...only display first line of multi program order
    %> <tr>
      <td><font face="Verdana" size="1"><a <%=fstatx%> href="EcomCdReport.asp?vEcom_No=<%=oRs("Ecom_No")%>"><%=oRs("Ecom_OrderNo")%></a></font>&nbsp; </td>
      <td><font face="Verdana" size="1"><%=oRs("Ecom_FirstName") & " " & oRs("Ecom_LastName")%></font>&nbsp; </td>
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
      <td colspan="2" align="center"><br><a <%=fstatx%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><p align="center"><a <%=fstatx%> href="Menu.asp"><img border="0" src="../Images/Icons/Administration.gif" alt="Click here for the Menu"></a><br>&nbsp;</p>
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
      <td align="right" width="30%" valign="top"><b><font face="Verdana" size="1">Order Address : </font></b></td>
      <td width="70%" valign="top"><font face="Courier New" size="2"><%=vAddress%></font><p>&nbsp;</p>
      </td>
    </tr>
    <tr>
      <td align="right" width="30%" valign="top"><b><font face="Verdana" size="1">Shipping Label : </font></b></td>
      <td width="70%" valign="top"><font face="Courier New" size="2"><%=vLabel%></font> </td>
    </tr>
    <tr>
      <td align="center" width="100%" valign="top" colspan="2" height="38"><br><a <%=fstatx%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><p align="center"><font face="Verdana" size="1"><a <%=fstatx%> href="EcomCdReport.asp">Ecom CD Sales List</a></font></p>
      <p align="center"><a <%=fstatx%> href="Menu.asp"><img border="0" src="../Images/Icons/Administration.gif" alt="Click here for the Menu"></a><br>&nbsp;</p>
      </td>
    </tr>
  </table>
  <%

  End If  
%>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>