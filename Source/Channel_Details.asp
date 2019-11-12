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

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

<% 
  Server.Execute vShellHi 

  '...update tables
  If Request.QueryString("vChan_Id").Count > 0 Then
    vChan_Id = Request.QueryString("vChan_Id")
    sGetChan
  ElseIf Request.Form("vChan_Id").Count > 0 Then
    sExtractChan
    sUpdateChan
    Response.Redirect "Channel_Report.asp"
  End If  
  
 
%>
  <form method="POST" action="Channel_Details.asp">

    <input type="hidden" name="vChan_ID" value="<%=vChan_ID%>">

    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <th align="right" width="30%" valign="top" nowrap>Channel Id :</th>
        <td width="70%"><%=vChan_Id%>&nbsp;&nbsp; </td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top" nowrap>Title : </th>
        <td width="70%"><input type="text" size="72" name="vChan_Title" value="<%=vChan_Title%>"></td>
      </tr>
      <tr>
        <th width="100%" valign="top" nowrap colspan="2">
        <div align="center"><br>
          <table border="1" cellpadding="7" cellspacing="0" bordercolor="#DDEEF9" id="table1" style="border-collapse: collapse">
            <tr>
              <th align="right">&nbsp;</th>
              <th align="right">2004</th>
              <th align="right">2005</th>
              <th align="right">2006</th>
              <th align="right">2007</th>
              <th align="right">2008</th>
              <th align="right">2009</th>
              <th align="right">2010</th>
              <th align="right">2011</th>
              <th align="right">2012</th>
            </tr>
            <tr>
              <th align="right">Ecommerce :</th>
              <th align="right"><%=vChan_2004e%></th>
              <th align="right"><%=vChan_2005e%></th>
              <th align="right"><%=vChan_2006e%></th>
              <th align="right"><%=vChan_2007e%></th>
              <th align="right"><%=vChan_2008e%></th>
              <th align="right"><%=vChan_2009e%></th>
              <th align="right"><%=vChan_2010e%></th>
              <th align="right"><%=vChan_2011e%></th>
              <th align="right"><%=vChan_2012e%></th>
            </tr>
            <tr>
              <th align="right">Manual :</th>
              <th align="right"><%=vChan_2004m%></th>
              <th align="right"><%=vChan_2005m%></th>
              <th align="right"><%=vChan_2006m%></th>
              <th align="right"><%=vChan_2007m%></th>
              <th align="right"><%=vChan_2008m%></th>
              <th align="right"><%=vChan_2009m%></th>
              <th align="right"><%=vChan_2010m%></th>
              <th align="right"><%=vChan_2011m%></th>
              <th align="right"><%=vChan_2012m%></th>
            </tr>
            <tr>
              <th align="right">Total :</th>
              <th align="right"><%=vChan_2004e + vChan_2004m%></th>
              <th align="right"><%=vChan_2005e + vChan_2005m%></th>
              <th align="right"><%=vChan_2006e + vChan_2006m%></th>
              <th align="right"><%=vChan_2007e + vChan_2007m%></th>
              <th align="right"><%=vChan_2008e + vChan_2008m%></th>
              <th align="right"><%=vChan_2009e + vChan_2009m%></th>
              <th align="right"><%=vChan_2010e + vChan_2010m%></th>
              <th align="right"><%=vChan_2011e + vChan_2011m%></th>
              <th align="right"><%=vChan_2012e + vChan_2012m%></th>
            </tr>
          </table>
        </div>

        </th>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top" nowrap>Owner : </th>
        <td width="70%"><input type="text" size="72" name="vChan_Owner" value="<%=vChan_Owner%>"></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top" nowrap>Contacts : </th>
        <td width="70%"><textarea rows="6" name="vChan_Contacts" cols="50"><%=vChan_Contacts%></textarea></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top" nowrap>Notes : </th>
        <td width="70%"><textarea rows="6" name="vChan_Notes" cols="50"><%=vChan_Notes%></textarea></td>
      </tr>
      <tr>
        <td align="center" width="100%" colspan="2" height="100">
          <input onclick="location.href='Channel_Report.asp'" type="button" value="Return" name="bReturn" id="bReturn"class="button"><%=f10%>
          <input type="submit" value="Update" name="bUpdate" class="button"></td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>