<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 

  Dim vChannel, vStrDate, vEndDate

  vChannel = fDefault(Request("vChannel"), Left(svCustId, 4))

  '...defaults to current month
  vStrDate  = Request("vStrDate")      : If Len(vStrDate) = 0 Then vStrDate = fFormatDate(MonthName(Month(Now) - 1) & " 1, " & Year(Now))
  vEndDate  = Request("vEndDate")      : If Len(vEndDate) = 0 Then vEndDate = fFormatDate(DateAdd("d", -1, MonthName(Month(Now)) & " 1, " & Year(DateAdd("m", +1, Now))))

  Function fCustOptions(vChannel)
    Dim vAll, vCust
    vAll = ""
    fCustOptions = ""
    vSql ="SELECT DISTINCT LEFT(Ecom_CustId, 4) AS Cust FROM Ecom ORDER BY Cust"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vCust = oRs("Cust")
      vAll  = vAll & " " & vCust
      fCustOptions  = fCustOptions & "<option " & fIf(vChannel=vCust, "selected ", "") & "value='" & vCust & "'>&nbsp;" & vCust & "&nbsp;</option>" & vbCrLf
      oRs.MoveNext	        
    Loop
    sCloseDb
    fCustOptions  = vbCrLf & "<option " & fIf(vChannel= "All", "selected ", "") & " value='" & vAll & "'>&nbsp;ALL&nbsp;&nbsp;</option>" & fCustOptions & vbCrLf
  End Function
%>

<html>

<head>
  <title>EcomExcel</title>
  <meta charset="UTF-8">

  <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
  <script src="//code.jquery.com/jquery-1.10.2.js"></script>
  <script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>

  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <style>
   .table tr td, .table tr th { width: 50%; }
  </style>
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>

  $(function() {
    $( "#strDate" ).datepicker({dateFormat: "M d, yy"});
    $( "#endDate" ).datepicker({dateFormat: "M d, yy"});
  });

  function validate(theForm) {

    if ($("#strDate")[0].value.length == 0) { 
      alert ("Please enter a Start Date");   
      $("#strDate")[0].focus();
      return (false);
    };
    if ($("#endDate")[0].value.length == 0) { 
      alert ("Please enter an End Date");   
      $("#endDate")[0].focus();
      return (false);
    };
    if (!isDate($("#strDate")[0].value, "en")) { 
      alert ("Please enter a valid Start Date");   
      $("#strDate")[0].focus();
      return (false);
    };
    if (!isDate($("#endDate")[0].value, "en")) { 
      alert ("Please enter a valid End Date");   
      $("#endDate")[0].focus();
      return (false);
    };

    return (true);
  }
  </script>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_Hi.asp"-->

  <h1>Ecommerce Sales - Excel [...coming]</h1>
  <h2><%=vChannel%> Channels / Authors</h2>
  <h3>Select a date range (maximum 12 months).</h3>
  <br />

  <form method="POST" action="EcomExcel.asp" onsubmit="return validate(this)">
    <table class="table">
      <tr>
        <td style="width: 90%;">
          <table class="table">
            <tr>
              <th>Channel/Author :</th>
              <td class="c2">
                <% If svMembLevel < 5 Then %>
                <%=vChannel%>
                <% Else  %>
                <select name="vChannel"><%=fCustOptions(vChannel)%></select>
                <% End If %>
              </td>
            </tr>
            <tr>
              <th>Select Start Date :</th>
              <td>
                <input type="text" id="strDate" name="vStrDate" size="12" value="<%=vStrDate%>">
              </td>
            </tr>
            <tr>
              <th>Select End Date :</th>
              <td>
                <input type="text" id="endDate" name="vEndDate" size="12" value="<%=vEndDate%>">
              </td>
            </tr>
          </table>
        </td>
      </tr>

    </table>

    <div style="text-align: center; margin-top: 40px;">
      <input type="submit" value="Generate Excel" name="bxcel" id="bExcel" class="button" onclick="$(this).hide(); alert('Coming...')";>
    </div>

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
