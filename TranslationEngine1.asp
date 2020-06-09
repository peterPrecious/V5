<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim vClass, vSelected  
  Dim vPages, vSource, vFile, vLine, oFs, oFolder, oFiles, oInp, vCntIn, vCntNo, vCntTr
  Const ForReading = 1, ForWriting = 2

  vCntIn = 0
  vCntNo = 0
  vCntTr = 0
  
  vSource = "\V5\Source"

  Set oFs = CreateObject("Scripting.FileSystemObject")   
  Set oFolder = oFs.GetFolder(Server.MapPath(vSource))
  Set oFiles = oFolder.Files

  '...get all .asp translatable files from Source
  For Each vFile in oFiles
    vCntIn = vCntIn + 1
   	Set oInp = oFs.OpenTextFile(Server.MapPath(vSource) & "\" & vFile.Name, ForReading, True)
    vLine = oInp.ReadAll
    If Instr(vLine, "[[") > 0 Or Instr(vLine, "{{") > 0 Or Instr(vLine, "[{") > 0 Then
      vCntTr = vCntTr + 1
      If Cint(DateDiff("d", Now, vFile.DateLastModified)) > -1 Then
        vClass = "d2"
        vSelected = " selected"
      Else
        vClass = "d2"
        vSelected = ""
      End If
      vPages = vPages & "<option  Class=" & Chr(34) & vClass & Chr(34) & vSelected & " value=" & Chr(34) & vFile.Name & Chr(34) & ">" & vFile.Name & "</option>" & vbCrLf
    Else
      vCntNo = vCntNo + 1
      If Cint(DateDiff("d", Now, vFile.DateLastModified)) > -1 Then
        vClass = "d4"
        vSelected = " selected"
      Else
        vClass = "d4"
        vSelected = ""
      End If
      vPages = vPages & "<option  Class=" & Chr(34) & vClass & Chr(34) & vSelected & " value=" & Chr(34) & vFile.Name & Chr(34) & ">" & vFile.Name & "</option>" & vbCrLf
    End If
  Next      

%>

<html>

<head>
  <title>:: Translation Engine 1/2</title>
  <meta charset="UTF-8">
  <script src="Inc/jQuery.js"></script>
  <link href="https://vubiz.com/v5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <style>
    .d1 {
      COLOR: #000000;
    }

    .d2 {
      COLOR: #000080;
    }

    .d3 {
      COLOR: #3977B6;
    }

    .d4 {
      COLOR: ORANGE;
    }
  </style>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->

  <form method="POST" action="TranslationEngine2.asp">
    <table style="width: 600px; margin: auto;">
      <tr>
        <td><h1>Vubiz Translation Engine</h1></td>
      </tr>
      <tr>
        <td class="c2">This service will analyze all selected pages and convert tagged phrases into database function calls.&nbsp; It defaults to selecting translatable pages that were modified today.&nbsp; Click &quot;...all Pages&quot; for all or individually click one or more pages.&nbsp; Click <b>Go</b> when ready.<br /><br /></td>
      </tr>
      <tr>
        <td class="c3" style="text-align: center"><%=vCntTr%> pages using translation tags. | <span style="color:#FFA500"><%=vCntNo%> pages without translation tags.</span></td>
      </tr>
      <tr>
        <td style="text-align: center">
          <br />
          <select size="12" name="vSelectPages" multiple>
            <option selected value="today">...pages modified today</option>
            <option value="all">...all Pages</option>
            <%=vPages %>
          </select>
          <input id="bGo" type="submit" value="Go" name="vGo" class="button040">
          <br>Use CTRL+Enter for multiple selections.
        </td>
      </tr>
    </table>
  </form>

  <script>
    $(document).on("load", function () { $("#bGo").show() }); 
    $("#bGo").on("click", function () { $(this).hide() }); 
  </script>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
