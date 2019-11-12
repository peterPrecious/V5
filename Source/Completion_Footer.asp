  <% If Lcase(svHost) = "localhost/v5" Or svMembLevel = 5 Or Session("Completion_Level") > 4 Then %>

    <div style="text-align:center">
      Debug: 
      <a href="<%=svPage%>?vCompletion_Debug=y">On</a> | <a href="<%=svPage%>?vCompletion_Debug=n">Off</a><br><br>
      <input type="button" value="Sessions" name="B1" onclick="window.open('Sessions.asp?vClose=y','Session','toolbar=no,width=450,height=800,left=10,top=10,status=yes,scrollbars=yes,resizable=yes')"></div>
  
    <% If Session("Completion_Debug") Then %>
      <br>
      <div style="text-align:center">
        <table style="border:1px solid orange">
          <tr>
            <td><%=vCompletion_Debug & "<br>"%></td>
          </tr>
        </table>
      </div>
    <% End If %>

    <br><br>

  <% End If %>