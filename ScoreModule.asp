<% 
  Dim vParms 

  '...transfer tracking to new system 
  If Request.QueryString.Count > 0 Then
    vParms = Request.QueryString.Item
  Else
    vParms = Request.Form.Item
  End If

  Response.Redirect "Logx.asp?" & vParms
%>
