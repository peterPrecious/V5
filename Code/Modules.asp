<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
  Dim vCookie, vRange, vLingo, vActiv, vFormt, vTitle, vMemos, vAccOk, vAccNo, vFeats, vTypes, vScript

  vCookie = svCustAcctId & "_Modules"

  '...if we did NOT arrive here from THIS form, check for a previously save cookie or a current cookie if we are returning/restarting from the report(s)
  If Request.Form.Count = 0 Then
    vRange = fDefault(Request.Cookies(vCookie)("vRange"), "")
    vLingo = fDefault(Request.Cookies(vCookie)("vLingo"), "EN, FR, ES, PT")
    vFormt = fDefault(Request.Cookies(vCookie)("vFormt"), "0")
    vActiv = fDefault(Request.Cookies(vCookie)("vActiv"), "*")
    vTitle = fDefault(Request.Cookies(vCookie)("vTitle"), "")
    vMemos = fDefault(Request.Cookies(vCookie)("vMemos"), "")
    vFeats = fDefault(Request.Cookies(vCookie)("vFeats"), "*")
    vTypes = fDefault(Request.Cookies(vCookie)("vTypes"), "*")
  '...else assume we arrived here from THIS form...
  Else
    vRange = Request("vRange")
    vLingo = fDefault(Request("vLingo"), "EN, FR, ES, PT")
    vFormt = fDefault(Request("vFormt"), "0")
    vActiv = fDefault(Request("vActiv"), "1")
    vTitle = Request("vTitle")
    vMemos = Request("vMemos")
    vFeats = fDefault(Request("vFeats"), "*")
    vTypes = fDefault(Request("vTypes"), "*")
  End If

  If Request.QueryString("vRange").Count = 1 Then vRange = Request.QueryString("vRange")

  '...save selection criteria in a cookie for session
  Response.Cookies(vCookie)("vRange") = vRange
  Response.Cookies(vCookie)("vLingo") = vLingo
  Response.Cookies(vCookie)("vFormt") = vFormt
  Response.Cookies(vCookie)("vActiv") = vActiv
  Response.Cookies(vCookie)("vTitle") = vTitle
  Response.Cookies(vCookie)("vMemos") = vMemos
  Response.Cookies(vCookie)("vFeats") = vFeats
  Response.Cookies(vCookie)("vTypes") = vTypes

  Select Case Request("vFunction")
    Case "ActiveY" : sUpdateModsActive Request.QueryString("vMods_Id"), 1
    Case "ActiveN" : sUpdateModsActive Request.QueryString("vMods_Id"), 0
  End Select
  
  If vRange = "0000" Then
    vScript = " onload=""divOn('divAdvanced')"""
  Else
    vScript = ""
  End If

  Function fLaunchUrl
    If Len(Trim(vMods_Script)) > 0 Then 
      fLaunchURL = vMods_Script & "('" & vMods_Id & "|N|N|N')"
    ElseIf vMods_FullScreen Then
      fLaunchURL = "fullScreen('P0000XX|" & vMods_Id & "|N|N|N')"
    Else
      fLaunchURL = "location.href='/V5/LaunchObjects.asp?vModId=" & vMods_Id & "|N|N|N&vNext=" & svPage & "'"
    End If  
  End Function
  
%>

<html>

<head>
  <title>Modules</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Launch.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
//    function zerofill(theElement, len) {
//      if (theElement.value.length < len) {
//        theElement.value = '100000'.substring(0,len-theElement.value.length) + theElement.value;
//      }   
//    };
    function empty(id){
      $("#" + id)[0].value="";
    };
  </script>
</head>

<body <%=vscript%>>

  <% 
    Server.Execute vShellHi 
  %>

  <div>
    <h1>Module Table</h1>
    <table class="table">
      <tr>
        <td style="width: 65%; padding: 10px;">
          <p class="c2">Select which Modules you would like to list (max 50, click <b>Next</b> if there are more.)</p>
          <form method="POST" action="Modules.asp">
            <table class="table">
              <tr>
                <th>whose ID starts with :<br />
                  (do not enter LANG values)&nbsp;&nbsp; </th>
                <td class="debug">
                  <!--              <input type="text" name="vRange" id="vRange" size="6" value="<%=vRange%>" maxlength="5" xonblur="zerofill(this, 5)">
                  <a title="Start at the beginning..." class="debug" onclick="fillField('vRange', '0'); divOn('divAdvanced');" href="#">&#937;</a>
modified May 7, 2015 to clear field onclick
  
-->
                  <input type="text" name="vRange" id="vRange" size="6" value="<%=vRange%>" maxlength="6">
                  <a title="Start at the beginning..." class="debug" onclick="empty('vRange'); divOn('divAdvanced');" href="#">&#937;</a>
                </td>
              </tr>
              <tr>
                <th>of type :</th>
                <td>
                  <select size="1" name="vTypes">
                    <option <%=fselect(vtypes,  "*")%> value="*">All</option>
                    <option <%=fselect(vtypes,  "z")%> value="Z">Z</option>
                    <option <%=fselect(vtypes,  "f")%> value="F">F</option>
                    <option <%=fselect(vtypes, "fx")%> value="FX">FX</option>
                    <option <%=fselect(vtypes,  "x")%> value="X">X</option>
                    <option <%=fselect(vtypes, "xx")%> value="XX">XX</option>
                    <option <%=fselect(vtypes,  "a")%> value="A">A</option>
                    <option <%=fselect(vtypes,  "u")%> value="U">U</option>
                    <option <%=fselect(vtypes,  "H")%> value="H">H</option>
                  </select>
                </td>
              </tr>

              <tr>
                <th>of feature :</th>
                <td>
                  <select size="1" name="vFeats">
                    <option <%=fselect(vfeats, "*")%> value="*">All</option>
                    <option <%=fselect(vfeats, "a")%> value="a">A : Audio</option>
                    <option <%=fselect(vfeats, "v")%> value="v">V : Video</option>
                    <option <%=fselect(vfeats, "c")%> value="c">C : Accessible</option>
                    <option <%=fselect(vfeats, "h")%> value="h">H : Hybrid</option>
                    <option <%=fselect(vfeats, "m")%> value="m">M : Mobile</option>
                  </select>
                </td>
              </tr>
              <tr>
                <th>and with status :</th>
                <td>
                  <input type="radio" value="*" name="vactiv" <%=fcheck(vactiv, "*")%>>All
                  <input type="radio" value="1" name="vactiv" <%=fcheck(vactiv, "1")%>>Active
                  <input type="radio" value="0" name="vactiv" <%=fcheck(vactiv, "0")%>>Inactive
                </td>
              </tr>
              <tr>
                <th>of language :</th>
                <td>
                  <input type="checkbox" name="vLingo" value="EN" <%=fchecks(vlingo, "en")%>>EN 
                  <input type="checkbox" name="vLingo" value="FR" <%=fchecks(vlingo, "fr")%>>FR 
                  <input type="checkbox" name="vLingo" value="ES" <%=fchecks(vlingo, "es")%>>ES
                  <input type="checkbox" name="vLingo" value="PT" <%=fchecks(vlingo, "pt")%>>PT 
                </td>
              </tr>
              <tr>
                <th>with Titles containing :</th>
                <td>
                  <input type="text" name="vTitle" size="30" value="<%=vTitle%>">
                </td>
              </tr>
              <tr>
                <th>with Memo containing :</th>
                <td>
                  <input type="text" name="vMemos" size="30" value="<%=vMemos%>">
                </td>
              </tr>
              <tr>
                <td colspan="2" style="text-align: center;">
                  <input type="submit" value="Go" name="bContinue" class="button070">
                </td>
              </tr>
            </table>
          </form>
        </td>
        <td style="text-align: center; width: 35%; padding: 10px;">
          <p class="c2"><b>Add</b> will generate the next available Module ID that is typically 7 characters long, ie 20023EN.</p>
          <br />
          <br />
          <input type="button" onclick="location.href='Module.asp?vMods_Id='" value="Add" name="bAdd" class="button070">
        </td>
      </tr>
    </table>
  </div>

  <!---Edit List-->
  <table style="width: 90%; margin: auto;" class="table">
    <tr>
      <td colspan="8" class="c2">Actions: <b>Edit</b> displays the Module details, <b>View</b> will launch the module without tracking, <b>Clone This</b> creates a new module with the same 4-5 character ID but allows you to add a new Language, ie 12345EN to 12345ES. <b>Clone New</b> creates a new Module Id.&nbsp; Note: you can click on the <b>Active</b> field to toggle status from Active (Y) to Inactive (N) and vice versa.<br />
        <br />
      </td>
    </tr>
    <tr>
      <td class="rowshade" style="text-align:center">Id</td>
      <td class="rowshade" style="text-align:center">Type </td>
      <td class="rowshade" style="text-align:center">Features</td>
      <td class="rowshade" style="text-align:center"><a title="Click value to change state." href="#">Active</a>? </td>
      <td class="rowshade" style="text-align: left">Title</td>
      <td class="rowshade" style="text-align:center" colspan="4">Action</td>
    </tr>
    <%
      Function fFeats()
        fFeats = ""
        If vMods_FeaAud Then fFeats = fFeats + "A"
        If vMods_FeaVid Then fFeats = fFeats + "V"
        If vMods_FeaAcc Then fFeats = fFeats + "C"
        If vMods_FeaHyb Then fFeats = fFeats + "H"
        If vMods_FeaMob Then fFeats = fFeats + "M"
      End Function

      i = 0
      sOpenDbBase


      vSql = "SELECT TOP 50 * FROM Mods " _ 
           & " WHERE " _
           & "   Mods_Id >= '" & vRange & "'" _
           & "   AND LEN(Mods_Id) >= " & Len(vRange) + 2 _
           & "   AND CHARINDEX(RIGHT(Mods_Id, 2), '" & vLingo & "') > 0 "_
           & "   AND CHARINDEX(CAST(Mods_Format AS CHAR(1)), '" & vFormt & "') > 0 "_
           & fIf (vActiv <> "*",  " AND Mods_Active = " & vActiv, "") _
           & fIf (vTypes <> "*",  " AND Mods_Type = '" & vTypes & "'", "") _ 
           & fIf (vTitle <> "",   " AND Mods_Title LIKE '%" & vTitle & "%'", "") _ 
           & fIf (vMemos <> "",   " AND Mods_Memo  LIKE '%" & vMemos & "%'", "") _
           & fIf (vFeats = "a",   " AND Mods_FeaAud = 1", "") _
           & fIf (vFeats = "v",   " AND Mods_FeaVid = 1", "") _
           & fIf (vFeats = "c",   " AND Mods_FeaAcc = 1", "") _
           & fIf (vFeats = "h",   " AND Mods_FeaHyb = 1", "") _
           & fIf (vFeats = "m",   " AND Mods_FeaMob = 1", "") _
           & " ORDER BY Mods_Id"  

'     Response.Write (vSql)
      Set oRsBase = oDbBase.Execute(vSQL)    
      Do While Not oRsBase.EOF 
        sReadMods
        i = i + 1
    %>
    <form method="POST" action="Modules.asp" name="fClone">
      <input type="hidden" name="vRange" value="<%=vMods_Id%>">
      <input type="hidden" name="vLingo" value="<%=vLingo%>">
      <tr>
        <td style="text-align: center"><%=vMods_Id%></td>
        <td style="text-align: center"><%=Ucase(vMods_Type)%></td>
        <td style="text-align: center"><%=fFeats()%></td>
        <td style="text-align: center"><% If vMods_Active Then%> <a class="d2" href="Modules.asp?vFunction=ActiveN&vMods_Id=<%=vMods_Id%>">Y</a> <% Else%> <a class="d2" href="Modules.asp?vFunction=ActiveY&vMods_Id=<%=vMods_Id%>">N</a> <% End If %> </td>
        <td><%=fLeft(vMods_Title, 48)%>&nbsp; </td>
        <td>
          <input type="button" onclick="location.href='Module.asp?vMods_Id=<%=vMods_Id%>'" value="Edit" name="bEdit" class="button070">
        </td>
        <td>
          <input type="button" onclick="<%=fLaunchUrl%>" value="View" name="bView" class="button070">
        </td>
        <td>
          <% 
          '  stop
            if (vMods_Id = "10007EN") Then Stop

            If Instr(vMods_Langs, "EN") = 0 Or Instr(vMods_Langs, "FR") = 0 Or Instr(vMods_Langs, "ES") = 0 Or Instr(vMods_Langs, "PT") = 0 Then 
          %>
              <input type="button" onclick="location.href='Module.asp?vCloneThis=<%=vMods_Id%>'" value="Clone This" name="bCloneThis" class="button070">
          <% 
            End If 
          %> 
        </td>
        <td>
          <input type="button" onclick="location.href='Module.asp?vCloneNew=<%=vMods_Id%>'" value="Clone New" name="bCloneNew" class="button070">
        </td>
      </tr>
    </form>
    <%  
        oRsBase.MoveNext
      Loop

      Set oRsBase = Nothing
      sCloseDbBase    
    %>
  </table>

  <script>
    // document.getElementById("vRange").value = "<%=Left(vMods_Id, 4)%>"
  </script>


  <div style="text-align: center">
    <br />
    <br />
    <% If i = 0 Then %>
    <h5>No Modules match your selection criteria.</h5>
    <% End If %>
    <% If i = 50 Then %>
    <input type="button" onclick="location.href='Modules.asp?vRange=<%=Left(vMods_Id, Len(vMods_Id) - 2)%>'" value="Next" name="bNext" class="button070">
    <% End If %>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


