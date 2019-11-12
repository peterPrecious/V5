<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Elog.asp"-->

<%
	Dim vCustId, vNext

	If Request("vCustId").Count > 0 Then
		If Len(Request("vCustId")) = 8 Then
			vCustId = Request("vCustId")
		ElseIf Len(Request("vCustId")) = 0 Then
			vCustId = ""			
		End If
	Else
		vCustId = svCustId	
	End If

%>

<html>

<head>
  <title>EcomHistory</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Ecommerce History Report</h1>
  <p class="c2">This lists in descending date order, the latest ecommerce transactions (max 500) that were submitted for processing <span class="red">but did not appear on our ecommerce system.</span>&nbsp; This does not mean that the transaction failed, it is possible that the customer decided not to continue.&nbsp; If you wish to see more than this account, enter the Customer Id (or leave empty for all).<br /><br />Tap on any History item for details of the transactions.&nbsp; If that is the missing one then click &quot;Post Transaction&quot; to generate the IDs and put the data into the ecommerce system. [Note: this has not been tested for some time - see Peter if needed.]</p>

  <div style="white-space: nowrap; text-align: center; width: 500px; margin: auto; border: 1px solid red;">
    <form method="POST" action="EcomHistory.asp">
      Customer ID:
          <input type="text" name="vCustId" size="10" value="<%=vCustId%>" class="c2" style="text-align: center">
      <br />
      <select name="vElog_No" size="20"><%=fElogById(vCustId)%></select>
      <br /><br />
      <input type="submit" value="Get History" name="bSubmit" class="button">
    </form>
  </div>

  <div style="white-space: nowrap; text-align: left; width: 500px; margin: auto; border: 1px solid red;">
    <%
			Dim vData
		  If Request("vElog_No").Count > 0 Then
		    i = ""
		    vElog_No = fDefault(vElog_No, Request("vElog_No"))
		    If vElog_No > 0 Then
		      sOpenDb
		      vSql = "SELECT * FROM Elog WHERE Elog_No = " & vElog_No 
		      Set oRs = oDb.Execute(vSql)
		      If Not oRs.Eof Then 
		        vData = oRs("Elog_Data")             
		        i = Server.HtmlEncode(vData)
		        i = Replace(i, "&gt;&lt;", "&gt;<br>&lt;")
		      End If  
		      Set oRs = Nothing
		      sCloseDb        
		      Response.Write "<br>" & i & "<br><br>"
		    End If
		  End If
								
			If Instr(vData,"Ecom2GenerateId.asp") > 0 Then 
				vNext = "Ecom2GenerateId.asp"
			ElseIf Instr(vData,"Ecom3GenerateId.asp") > 0 Then 
				vNext = "Ecom3GenerateId.asp"
			Else 
				vNext = ""
			End If
					If Len(vNext) > 0 Then	
    %>
    <div style="text-align: center;">
      <form method="POST" action="<%=vNext%>" name="myForm">
        <input type="submit" value="Post Transaction" name="bPost" class="button">
        <%=vData%>
      </form>
    </div>
    <%
								End If
    %>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


