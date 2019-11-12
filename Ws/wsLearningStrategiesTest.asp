
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <title>New Page 1</title>
</head>

<body>

  <form method="POST" action="wsLearningStrategies.asp">
    <p><input type="text" name="vMemb_FirstName" size="20" value="Peter"></p>
    <p><input type="text" name="vMemb_LastName" size="20" value="Bulloch"></p>
    <p>&nbsp;</p>
    <p><input type="submit" value="Submit" name="B1"></p>
    <input type="hidden" name="vPassword" value="1010101">
    <input type="hidden" name="vCust_Id" value="LNST2598">
    <input type="hidden" name="vMemb_Email" value>
    <input type="hidden" name="vMemb_Programs" value="P1001EN">
    <input type="hidden" name="vMemb_Expires" value="90">
    <%
      vPassword           = Request.Form("vPassword")
      vCust_Id            = Request.Form("vCust_Id")
      vMemb_Id            = "" '...always setup new user
      vMemb_FirstName     = Request.Form("vMemb_FirstName") 
      vMemb_LastName      = Request.Form("vMemb_LastName")
      vMemb_Email         = Request.Form("vMemb_Email")
    
      '...extract learning program(s), ie P1001EN|P1202EN then change the pipes to spaces
      vMemb_Programs      = Request.Form("vMemb_Programs") 
      vMemb_Programs      = Replace(vMemb_Programs, "|", " ")
    
      '...expires can be empty, or no days, or date
      i = Request.Form("vMemb_Expires")
  %>
  </form>

</body>

</html>