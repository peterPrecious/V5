<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<html>

<head>
  <title>Completion_LocationManager_Loc</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
  
    // field tests
    var reAlphaNumeric = new RegExp(/^[0-9A-Za-z]+$/);
    var reAlpha        = new RegExp(/^[A-Za-z]+$/);
    var reNumeric      = new RegExp(/^[0-9]+$/);

    var L0len          = <%=Session("Completion_L0len")%>;
    var L1len          = <%=Session("Completion_L1len")%>;
    
    var L0tit          = "<%=Session("Completion_L0tit")%>";
    var L1tit          = "<%=Session("Completion_L1tit")%>";

    function validateReg (theForm) {
      if (theForm.vUnit_L1.value.length != L1len) {
        alert("Please enter a " + L1len + " character " + L1tit + " Id.");
        theForm.vUnit_L1.focus();
        return (false);
      }
      if (theForm.vUnit_L1Title.value == "") {
        alert("Please enter the " + L1tit + " Name.");
        theForm.vUnit_L1Title.focus();
        return (false);
      }    
      if (theForm.vUnit_L1Title.value.length < 4 || theForm.vUnit_L1Title.value.length > 128) {
        alert("The " + L1tit + " Name must be between 4 and 128 characters.");
        theForm.vUnit_L1Title.focus();
        return (false);  
      }
      if (theForm.vUnit_L0.value.length != L0len) {
        alert("Please enter a " + L0len + " character " + L0tit + " Id.");
        theForm.vUnit_L0.focus();
        return (false);
      }
      if (theForm.vUnit_L0Title.value == "") {
        alert("Please enter the " + L0tit + " Name.");
        theForm.vUnit_L0Title.focus();
        return (false);
      }    
      if (theForm.vUnit_L0Title.value.length < 4 || theForm.vUnit_L0Title.value.length > 128) {
        alert("The " + L0tit + " Name must be between 4 and 128 characters.");
        theForm.vUnit_L0Title.focus();
        return (false);  
      }
      return (true);
    }   
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--#include file = "Completion_LocationManager_Top.asp"-->

  <div style="margin-bottom: 30px;">
    <h1>Add a new <%=Session("Completion_L0Tit")%></h1>
    <h2>First click on the <%=Session("Completion_L1Tit")%> where you wish to add the <%=Session("Completion_L0Tit")%>.</h2>
  </div>

  <p style="text-align: center">
    <% i = fL1s ("All") %>
    <select name="vUnit_L2" onclick="location.href='Completion_LocationManager_Add.asp?vReg=' + value" size="<%=fDefault(vSelectNo , 7)%>">
      <%=i%>
    </select>
  </p>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>
