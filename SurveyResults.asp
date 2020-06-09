<% 
  Dim vParms 

  '...transfer tracking to new system 
  If Request.QueryString.Count > 0 Then
    vParms = Request.QueryString.Item
  Else
    vParms = Request.Form.Item
  End If

  '...if accessible modules then close
  If Lcase(Request("vAccess")) = "y" Then 
  	Response.Redirect "CloseObjects.asp?" & vParms
	Else
	  Response.Redirect "Logx.asp?" & vParms
	End If
%>








<!--- original survey tracking -->


<!-- include virtual = "V5/Inc/Setup.asp"-->
<!-- include virtual = "V5/Inc/Initialize.asp"-->
<!-- include virtual = "V5/Inc/Db_Cust.asp"-->
<!-- include virtual = "V5/Inc/Db_Logs.asp"-->

<%
'  vLogs_Item = fDefault(Request("vProgId"), "P0000XX") & "|" & Request("vModId") & "_" & fUnquote(Request("vResults"))
'  sLogSurveyResults

  '...if accessible modules then close
' If Lcase(Request("vAccess")) = "y" Then Response.Redirect "CloseObjects.asp?vProgId=" & Request("vProgId") & "&vModId=" & Request("vModId") & "&vTimeSpent=" & Request("vTimeSpent")  
%>