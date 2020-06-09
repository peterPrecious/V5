<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  '...you can get here from SeatThreshold.asp as ...ProgramsAssigned.asp?custAcctId=1234
  Dim custId, custAcctId
  If Request("custId").Count = 0 Then
    custId = svCustId
    custAcctId = svCustAcctId   
  Else
    custId = Request("custId")
    custAcctId = Right(custId, 4)
  End If


'      & " WHERE (Cust.Cust_AcctId = '" & svCustAcctId & "') "_


  '...First see if any programs were purchased
  Dim bOk : bOk = False
  vSql = " SELECT"_
       & "   Ecom.Ecom_Programs AS ProgId,"_
       & "   SUM(CASE WHEN Ecom_Amount < 0 THEN Ecom_Quantity * -1 ELSE Ecom_Quantity END) AS Purchased"_
       & " FROM Cust INNER JOIN Ecom ON Cust.Cust_AcctId = Ecom.Ecom_NewAcctId "_
       & " WHERE (Cust.Cust_AcctId = '" & custAcctId & "') "_
       & " GROUP BY Cust.Cust_Id, Ecom.Ecom_Programs "
  sOpenDb
  Set oRs = oDb.Execute(vSql)
  If Not oRS.eof Then bOk = True
  Set oRs = Nothing
  sCloseDb
%>

<html>

<head>
  <meta charset="UTF-8">
  <title></title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi
  %>

  <table style="width: 50%; margin: auto;" class="tableBorder">
    <tr>
      <td style="text-align: center; padding: 20px;">
        <h1><!--webbot bot='PurpleText' PREVIEW='Programs Purchased and Assigned for'--><%=fPhra(000883)%>&ensp;<%=custId %></h1><br />
        <h2 class="c2"><!--webbot bot='PurpleText' PREVIEW='This report displays all learners, sorted by Last Name, who have had Programs assigned to them.'--><%=fPhra(000879)%></h2>
        <% If Not bOk Then %>
        <br />
        <h5 class="c2"><!--webbot bot='PurpleText' PREVIEW='This Account does not contain any Purchased Programs.'--><%=fPhra(001770)%></h5>
        <% Else  %>
        <h3 class="c3"><!--webbot bot='PurpleText' PREVIEW='Note: When either the amount of Learners or the amount of Programs assigned is large (ie more than 100 learners and/or more than 20 programs) only use the Excel version.'--><%=fPhra(001771)%></h3>
        <br /><br />
        <input onclick="location.href = 'ProgramsAssigned_O.asp?custId=<%=custId%>'" type="button" value="<%=bOnline%>" name="bOnline" class="button070">
        <%=f10() %>
        <input onclick="location.href = 'ProgramsAssigned_X.asp?custId=<%=custId%>'" type="button" value="Excel" name="bExcel" class="button070">
        <% End If %>
      </td>
    </tr>
  </table>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
</body>

</html>


