<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
  vCust_Id = Request("vCust_Id")

  If Request("vForm") = "y" Then  
    vCust_Programs = Replace(Request("vCust_Programs"), vbCrLf, " ")
    sUpdateCustPrograms
'   Response.Redirect "CustomerEdit.asp?vEditCustId=" & vCust_Id & "&vHidden=n#Programs"
    Response.Redirect "Customer.asp?vEditCustId=" & vCust_Id & "&vHidden=n#Programs"
  End If  
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body text="#000080" vlink="#000080" alink="#000080" link="#000080" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <% Server.Execute vShellHi %>
  <form action="ContentEditAll.asp" method="POST" target="_self">
    <table style="BORDER-COLLAPSE: collapse" bordercolor="#DDEEF9" border="0" width="100%" cellpadding="5">
      <tr>
        <td valign="top" align="center"><h1 align="left">Edit Program String</h1><h2 align="left">This allows you to edit the entire program string.&nbsp; Ensure each program begins on a new line.&nbsp; Click <b>Update</b> when finished and this string will replace the one on the customer file and you will be returned to the customer table. Note that you can preface one or more programs with a title from the Catalogue Table.&nbsp; Enter the Catalogue Id just before the program as [G1234EN].&nbsp; The format of each string is:&nbsp; <br><br>[CatalogueId]Program~Online$US~Online$CA~Len Hours~Len Days~CDs?~VuBooks?~Self Assessments?~Exam Id~Min Ques~Max Attempts~Time/Bank~Pass Grade<br><br>[G0001EN]P0004EN~95~149~80~90~n~n~y~1018EN~10~50~4~75</h2>
        Customer :&nbsp;<%=vCust_Id%> <p><textarea rows="15" name="vCust_Programs" cols="80"><%=Replace(Request("vCust_Programs"), " ", vbCrLf)%></textarea></p>&nbsp;<p>
        
        <input onclick="javascript:history.back(1)" type="button" value="Return" name="bReturn" id="bReturn"class="button"><%=f10%>
        <input type="submit" value="Update" name="bUpdate" class="button"></p>
        </td>
      </tr>
      </table>
    <input type="hidden" name="vForm" value="y"><input type="hidden" name="vCust_Id" value="<%=Request("vCust_Id")%>">
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>