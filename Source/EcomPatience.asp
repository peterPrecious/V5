<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Elog.asp"-->

<% 

  '...values can be received by querystring or form but are forward only by querystring
  Dim vNext, vFilter, vParms, vElogs

  stop


  '...If info received via a querystring (or form get)
  vNext  = fDefault(Request("vNext"), "Timeout.asp")
  vParms = vbCrLf

  '...extract and forward by form post to vNext all filtered fields 
  vFilter = Request.Form("vFilter")
  For Each vFld in Request.Form

  	'...only pass key variables to eLogs
		If Left(vFld, 5) = "vNext" Or Left(vFld, 5) = "vMemb" Or Left(vFld, 5) = "vEcom" Then
	    vElogs = vElogs & "<input type=""hidden"" name=""" & vFld & """ value=""" & Request(vFld) & """>" & vbCrLf & "    "
	  End If

	  '...only pass filtered variables to I/S
    If Instr(vFilter, vFld) > 0 Then
      vParms = vParms & "<input type=""hidden"" name=""" & vFld & """ value=""" & Request(vFld) & """>" & vbCrLf & "    "
    End If

  Next

  '...log transaction if we get a vElog field containing an 8 character customer id plus optional identifier, ie ABCD1234 or ABCD1234eatme
  If Len(Request("vElog")) > 7 Then
    vElog_CustId = Ucase(Left(Request("vElog"), 8))
    vElog_Id = fIf(Len(Request("vElog"))> 8, Mid(Request("vElog"), 9), NULL)
    spElogInsert vElog_CustId, vElog_Id, vElogs
  End If

%>

<html>

<head>
  <title>EcomPatience</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
     $(function() { $("#myForm").submit() });
  </script>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->

  <div style="text-align: center; margin-top: 20px;">
    <h1><!--[[-->Please be patient.<!--]]--></h1>
    <h2><!--[[-->It can take several minutes for the next page to appear.<!--]]--></h2>
    <% If Len(Request("vMsg")) > 0 Then %><h3><%=Request("vMsg")%></h3><% End If %>
    <br />
    <p><img border="0" src="../Images/Common/ProgressBar.gif"></p>

    <form method="POST" action="<%=vNext%>" name="myForm" id="myForm">
      <!--<input type="submit" value="Submit" name="bSubmit" class="button">-->
      <%=vParms%>
    </form>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
