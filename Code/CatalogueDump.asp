<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
  Dim vPromo, vCatlT, vCatlN, vProgS, vProgN, vProgD, vProgL, vProgM, vProgE, vModsN, vModsI, vModsD, vFormt, vOrder

  vCatlT = fDefault(Request("vCatlT"), "0")
  vCatlN = fDefault(Request("vCatlN"), "y")
  vProgS = fDefault(Request("vProgS"), "n")
  vProgN = fDefault(Request("vProgN"), "n")
  vProgD = fDefault(Request("vProgD"), "y")
  vProgL = fDefault(Request("vProgL"), "y")
  vProgM = fDefault(Request("vProgM"), "y")
  vProgE = fDefault(Request("vProgE"), "n")
  vModsN = fDefault(Request("vModsN"), "n")
  vModsI = fDefault(Request("vModsI"), "n")
  vModsD = fDefault(Request("vModsD"), "n")
  vPromo = fDefault(Request("vPromo"), "y")
  vFormt = fDefault(Request("vFormt"), "o") '...eventually might add in "PDF" format
  vOrder = fDefault(Request("vOrder"), "a")

  If Request("vHidden").Count > 0 Then Response.Redirect "CatalogueDump_" & vFormt & ".asp?" & Request.ServerVariables("QUERY_STRING")

%>

<html>

<head>
  <title>CatalogueDump</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
    function validate(theForm)
    {

      if (theForm.vCatlT.selectedIndex < 0)
      {
        alert("Please select one of the \"Categories\" options.");
        theForm.vCatlT.focus();
        return (false);
      }

      var numSelected = 0;
      var i;
      for (i = 0;  i < theForm.vCatlT.length;  i++)
      {
        if (theForm.vCatlT.options[i].selected)
            numSelected++;
      }
      if (numSelected < 1)
      {
        alert("Please select at least 1 of the \"Categories\" options.");
        theForm.vCatlT.focus();
        return (false);
      }
      return (true);
    }
  </script>

</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Catalogue Dump</h1>
  <p class="c2">This displays/dumps the key items of your catalogue which can then be cut/pasted into another document.<br /><br />It will always display the Program Title but the rest of the fields are optional.&nbsp; Promo Info is notation about new productions, pricing, etc.&nbsp; It will sort in Catalogue Name order followed by Program Name order.<br /><br /></p>

  <form method="GET" action="CatalogueDump.asp" onsubmit="return validate(this)">

    <table style="width: 80%; margin: auto;">
      <tr>
        <th>Include Categories :&nbsp;</th>
        <td>
          <select size="5" name="vCatlT" multiple>
            <option value="0">All</option>
            <%
                Dim vTitle
                vSql = "SELECT Catl_Title, Catl_No FROM Catl WHERE (Catl_CustId = '" & svCustId & "') AND (Catl_Active = 1) ORDER BY Catl_Title"
                sOpenDb    
                Set oRs = oDb.Execute(vSql)
                Do While Not oRs.Eof 
                  vTitle = oRs("Catl_Title") 
                  i = Instr(vTitle, "<") 
                  If i > 0 Then vTitle = Left(vTitle, i - 1)                  
            %>
            <option value="<%=oRs("Catl_No")%>"><%=fLeft(vTitle, 48)%></option>
            <%              
	                oRs.MoveNext
                Loop
                Set oRs = Nothing
                sCloseDb   
            %>
          </select><br>&nbsp;Use Ctrl+Click for multiple categories
        </td>
      </tr>
      <tr>
        <th>Include Category Name :&nbsp;</th>
        <td>
          <input type="checkbox" name="vCatlN" value="y" <%=fCheck("y", vCatlN)%>>
        </td>
      </tr>
      <% If svMembLevel = 5 Then %>
      <tr>
        <th>Include Program Strings :</th>
        <td>
          <input type="checkbox" name="vProgS" value="y" <%=fCheck("y", vProgS)%>>
          (Vubiz Administrators only)
        </td>
      </tr>
      <% End If %>

      <tr>
        <th>Include Program Description : </th>
        <td>
          <input type="checkbox" name="vProgD" value="y" <%=fCheck("y", vProgD)%>>
        </td>
      </tr>

      <tr>
        <th class="red">Hide Program Length (Hrs) : </th>
        <td>
          <input type="checkbox" name="vProgL" value="y" <%=fCheck("y", vProgL)%>>
        </td>
      </tr>

      <tr>
        <th class="red">Show Count of Modules : </th>
        <td>
          <input type="checkbox" name="vProgM" value="y" <%=fCheck("y", vProgM)%>>
        </td>
      </tr>

      <tr>
        <th>Include Module Name (and Id) : </th>
        <td>
          <input type="checkbox" name="vModsN" value="y" <%=fCheck("y", vModsN)%>>
        </td>
      </tr>
      <tr>
        <th>Include Module Id Only : </th>
        <td>
          <input type="checkbox" name="vModsI" value="y" <%=fCheck("y", vModsI)%>>
          (Do not select above option with this option)
        </td>
      </tr>


      <tr>
        <th>Include Module Description : </th>
        <td>
          <input type="checkbox" name="vModsD" value="y" <%=fCheck("y", vModsD)%>>
        </td>
      </tr>
      <tr>
        <th class="red">Hide "this course has an examination" : </th>
        <td>
          <input type="checkbox" name="vPromo" value="y" <%=fCheck("y", vPromo)%>>
        </td>
      </tr>
      <tr>
        <th>Remove Promo Info : </th>
        <td>
          <input type="checkbox" name="vPromo" value="y" <%=fCheck("y", vPromo)%>>
        </td>
      </tr>
      <tr>
        <td style="text-align: center;" colspan="2">
          <input type="submit" value="Continue" name="bContinue" class="button">
          <h2>Be patient, this will take several minutes.</h2>
        </td>
      </tr>

    </table>
    <input type="hidden" name="vHidden" value="y">

  </form>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->


</body>

</html>


