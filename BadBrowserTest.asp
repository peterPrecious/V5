<%

  '...This page tries to read a Session variable set previously.  This is a popup window...and if the
  '   variable is NOT read, that means we have launched via "My Computer", Outlook, etc.
  '   Once determined, we re-launch the opener (Cookies.asp) adding an additional parameter with the 
  '   results of this page (GoodBrowser=Yes/No)

  Dim vValue
  '...check if we can read a Session variable...
  '...if not, we have accessed this web via a Bad Browser (My Computer, Outlook, etc)
  If Session("Cookies") = "Yes" Then
    vValue = "Yes"
  Else
    vValue = "No"
  End If

%>

<html>
  <script>
    opener.location.href = opener.location.href + '&GoodBrowser=<% =vValue %>'
    close()
  </script>
  <body></body>
</html>