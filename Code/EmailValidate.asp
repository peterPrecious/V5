
<html>

<head>
  <script>
    var validRegExp = "^[a-z0-9][a-z0-9_\.-]{0,}[a-z0-9]@[a-z0-9][a-z0-9_\.-]{0,}[a-z0-9][\.][a-z0-9]{2,4}$";

    function Validate(theForm) 
    {
      if (theForm.xxxEmail.value.toLowerCase().search(validRegExp) == -1)
      {
        alert("Invalid Email");
        theForm.xxxEmail.focus();
        return (false);
      }
      return (true);
    }
  </script>
</head>

<body>
  <form name="fForm" onsubmit="return Validate(this)" action="EmailValidate.asp" method="POST">
    Email Address : &nbsp;<input type="text" name="xxxEmail" size="32"><input type="submit" value="Continue" name="bContinue" class="button">
  </form>
</body>

</html>


