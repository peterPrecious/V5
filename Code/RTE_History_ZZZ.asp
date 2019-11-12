<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="//code.jquery.com/jquery-1.9.1.js"></script>
  <script src="/V5/Inc/Functions.js"></script>
  <script>
    // expand or contract test
    var less = true;
    $(function () {renderExpandable()})
    $("#expandable").click(function () {renderExpandable()})

    function renderExpandable() {
      if (less) {
        $(".expandable").addClass("less");
        $("#expandable").text("More Text");
      } else {
        $(".expandable").removeClass("less");
        $("#expandable").text("Less Text");

      }
      less = !less;
    }
    </script>

    <style>
      .expandable { width: 50px; font-style: italic; }
      .less { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; display:block; -o-text-overflow: ellipsis; }
    </style>
  </head>

<body>

  <a href="#" class="green" onclick="renderExpandable()" id="expandable"></a>
    <table width="100%">

      <tr>

        <td>
          <div class="expandable">Once upon a time when the pigs were swine</div>
        </td>
      </tr>
    </table>
 

</body>

</html>


