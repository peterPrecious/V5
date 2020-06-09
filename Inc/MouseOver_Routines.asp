
<%
  '...Mouseover functions for links displays in status line (note: can also use fStatusOff in Initialize.com)
  Function fMouseOver (vStatus)
    fMouseOver = "onmouseover=" & Chr(34) & "jMouseOver('" & vStatus & "');return true" & Chr(34) & " onmouseout=" & Chr(34) & "jMouseOut();return true" & Chr(34) & ""
  End Function
%>


<script>
  function jMouseOver(vStatus)
  {
    window.status = vStatus
  }
    
  function jMouseOut()
  {
    window.status = ""
  }
</script>