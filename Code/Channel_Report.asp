<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Chan.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
  	Server.Execute vShellHi
  %>
  <table width="100%" border="1" cellspacing="0" bordercolor="#DDEEF9" cellpadding="2" style="border-collapse: collapse">
    <tr>
      <th colspan="13">
      <form method="POST" action="CustomerExpiryReport.asp">
        <h1><br>Channel Report</h1><h2>This lists the channels with their <b>approximate</b> annual sales.&nbsp; Note that US and CA sales are lumped together.&nbsp; <br>Chick on the Channel Id for details.&nbsp; Source is E:Ecommerce, M:Manual (V or C) and T:Total (sum of E+M).</h2><h2><a href="Channel_Report_x.asp">Click here for Excel Version of this report.</a></h2>
      </form>
      </th>
    </tr>
    <tr>
      <th bgcolor="#DDEEF9" height="32" bordercolor="#FFFFFF">Channel</th>
      <th align="left" bgcolor="#DDEEF9" height="32" bordercolor="#FFFFFF">Title</th>
      <th bgcolor="#DDEEF9" height="32" bordercolor="#FFFFFF">Source</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2004</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2005</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2006</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2007</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2008</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2009</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2010</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2011</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">2012</th>
      <th bgcolor="#DDEEF9" height="32" align="right" bordercolor="#FFFFFF" width="65">Total</th>
    </tr>
    <%
      Dim vTotale, vT2004e, vT2005e, vT2006e, vT2007e, vT2008e, vT2009e, vT2010e, vT2011e, vT2012e, vGrande 
      Dim vTotalm, vT2004m, vT2005m, vT2006m, vT2007m, vT2008m, vT2009m, vT2010m, vT2011m, vT2012m, vGrandm 
      Dim vTotalt, vT2004t, vT2005t, vT2006t, vT2007t, vT2008t, vT2009t, vT2010t, vT2011t, vT2012t, vGrandt 
    
      sGetChan_Rs
      Do While Not oRs.Eof  
        sReadChan  
      
        vT2004e = vT2004e + vChan_2004e
        vT2005e = vT2005e + vChan_2005e
        vT2006e = vT2006e + vChan_2006e
        vT2007e = vT2007e + vChan_2007e
        vT2008e = vT2008e + vChan_2008e
        vT2009e = vT2009e + vChan_2009e
        vT2010e = vT2010e + vChan_2010e
        vT2011e = vT2011e + vChan_2011e
        vT2012e = vT2012e + vChan_2012e

        vTotale = vChan_2004e + vChan_2005e + vChan_2006e + vChan_2007e + vChan_2008e + vChan_2009e + vChan_2010e + vChan_2011e + vChan_2012e
        vGrande = vGrande + vTotale

        vT2004m = vT2004m + vChan_2004m
        vT2005m = vT2005m + vChan_2005m
        vT2006m = vT2006m + vChan_2006m
        vT2007m = vT2007m + vChan_2007m
        vT2008m = vT2008m + vChan_2008m
        vT2009m = vT2009m + vChan_2009m
        vT2010m = vT2010m + vChan_2010m
        vT2011m = vT2011m + vChan_2011m
        vT2012m = vT2012m + vChan_2012m

        vTotalm = vChan_2004m + vChan_2005m + vChan_2006m + vChan_2007m + vChan_2008m + vChan_2009m + vChan_2010m + vChan_2011m + vChan_2012m
        vGrandm = vGrandm + vTotalm

        vT2004t = vT2004e + vT2004m
        vT2005t = vT2005e + vT2005m
        vT2006t = vT2006e + vT2006m
        vT2007t = vT2007e + vT2007m
        vT2008t = vT2008e + vT2008m
        vT2009t = vT2009e + vT2009m
        vT2010t = vT2010e + vT2010m
        vT2011t = vT2011e + vT2011m
        vT2012t = vT2012e + vT2012m

        vTotalt = vChan_2004e + vChan_2005e + vChan_2006e + vChan_2007e + vChan_2008e + vChan_2009e + vChan_2010e + vChan_2011e + vChan_2012e + vChan_2004m + vChan_2005m + vChan_2006m + vChan_2007m + vChan_2008m + vChan_2009m + vChan_2010m + vChan_2011m + vChan_2012m

        vGrandt = vGrandt + vTotalt



    %> 
    <tr>
      <td colspan="13">&nbsp;</td>
    </tr>
    <tr>
      <td valign="Top" align="center" height="20" rowspan="3"><a href="Channel_Details.asp?vChan_Id=<%=vChan_Id%>"><%=vChan_Id%></a></td>
      <td valign="Top" height="20" rowspan="3"><%=fLeft(vChan_Title, 60)%></td>
      <td valign="Top" align="center" height="8">E</td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2004e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2005e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2006e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2007e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2008e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2009e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2010e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2011e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8"><%=FormatNumber(vChan_2012e, 0)%></td>
      <td valign="Top" align="right" width="65" height="8" bgcolor="#DDEEF9" bordercolor="#FFFFFF" ><%=FormatNumber(vTotale, 0)%></td>
    </tr>
    <tr>
      <td valign="Top" align="center" height="6">M</td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2004m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2005m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2006m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2007m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2008m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2009m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2010m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2011m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2012m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6" bgcolor="#DDEEF9" bordercolor="#FFFFFF" ><%=FormatNumber(vTotalm, 0)%></td>
    </tr>
    <tr>
      <td valign="Top" align="center" height="6">T</td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2004e + vChan_2004m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2005e + vChan_2005m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2006e + vChan_2006m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2007e + vChan_2007m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2008e + vChan_2008m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2009e + vChan_2009m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2010e + vChan_2010m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2011e + vChan_2011m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6"><%=FormatNumber(vChan_2012e + vChan_2012m, 0)%></td>
      <td valign="Top" align="right" width="65" height="6" bgcolor="#DDEEF9" bordercolor="#FFFFFF" ><%=FormatNumber(vTotalt, 0)%></td>
    </tr>
    <%  
        oRs.MoveNext
      Loop
      Set oRs = Nothing
    %> 
    <tr>
      <td colspan="13">&nbsp;</td>
    </tr>
    <tr>
      <th height="30" bgcolor="#DDEEF9" rowspan="3" valign="top" bordercolor="#FFFFFF" colspan="2" align="left">Total</th>
      <th height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9">E</th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2004e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2005e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2006e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2007e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2008e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2009e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2010e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2011e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2012e, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vGrande, 0)%></th>
    </tr>
    <tr>
      <th height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9">M</th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2004m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2005m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2006m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2007m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2008m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2009m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2010m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2011m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2012m, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vGrandm, 0)%></th>

    </tr>
    <tr>
      <th height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9">T</th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2004t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2005t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2006t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2007t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2008t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2009t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2010t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2011t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vT2012t, 0)%></th>
      <th align="right" width="65" height="10" bordercolor="#FFFFFF" bgcolor="#DDEEF9"><%=FormatNumber(vGrandt, 0)%></th>

    </tr>
    </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->


</body>

</html>

