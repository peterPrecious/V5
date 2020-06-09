<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Comp.asp"-->

<% 
  Dim vModsLang, bLive : bLive = False

  '...update tables
  If Request.Form("bUpdate").Count = 1 Then

    sExtractMods
    If Len(vMods_Id) >= 4 And IsNumeric(vMods_Id) Then 
      vModsLang = Ucase(Request("vModsLang"))
      If Len(vModsLang) = 2 Then
        vMods_Id = vMods_Id & Ucase(Request("vModsLang"))
      Else
        Response.Redirect "Error.asp?vErr=You must select a language for this Module."
      End If
    End If

    spModsAlterById
		spUpdateProgLengths vMods_Id
    Response.Redirect "Modules.asp?vMods_Id=" & vMods_Id

  ElseIf Len(Request("vDelModId")) > 3 Then 
    vMods_Id = Request("vDelModId")
    sDeleteMods
		spUpdateProgLengths vMods_Id
    Response.Redirect "Modules.asp"

' ElseIf Len(Request("vCloneNew")) = 6 Then
  ElseIf Len(Request("vCloneNew")) >= 6 Then
    sGetMods Request("vCloneNew")
    vMods_Id = fNextModsId

' ElseIf Len(Request("vCloneThis")) = 6 Then
  ElseIf Len(Request("vCloneThis")) >= 6 Then
    sGetMods Request("vCloneThis")
'   vMods_Id = Left(vMods_Id, 4)
    vMods_Id = Left(vMods_Id, Len(vMods_Id) - 2)

  Else
    vMods_Id = Request("vMods_Id")  
    If Len(vMods_Id) = 0 Then vMods_Id = fNextModsId

  End If  

' If Len(vMods_Id) = 6 And Len(Request("vCloneNew")) <> 6 And Len(Request("vCloneThis")) <> 6 Then
  If Len(vMods_Id) >= 6 And Len(Request("vCloneNew")) < 6 And Len(Request("vCloneThis")) < 6 Then
    sGetMods vMods_Id
    bLive = True
  End If

  
  '...this is legacy way of finding a free module id
  Function fNextModsId_prev
    Dim vCurr, vPrev
    vCurr = 0
    sOpenDbBase
    vSql = "SELECT LEFT(Mods_Id, 4) AS Mods_No FROM Mods ORDER BY Mods_Id"
    Set oRsBase = oDbBase.Execute(vSQL)    
    Do While Not oRsBase.Eof 
      vCurr = Clng(oRsBase("Mods_No"))
      If vCurr - vPrev > 1 Then 
        fNextModsId = Right("0000" & vPrev + 1, 4)
        sCloseDbBase
        Exit Function
      End If
      oRsBase.MoveNext
      vPrev = vCurr
    Loop
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


  '...this is new way to get the new module Ids , ie 20031EN
  Function fNextModsId
    sOpenDbBase
    vSql = "SELECT MAX(CAST(LEFT(Mods_Id, LEN(Mods_Id) - 2) AS int)) + 1 AS Next_No FROM Mods"
    Set oRsBase = oDbBase.Execute(vSQL)    
    fNextModsId = Clng(oRsBase("Next_No"))
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function


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
  <title>Module</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Launch.js"></script>
  <script>
    function fSize(x, y) {
      if (x > 0) {
        document.getElementById("vMods_Width").value = x;  
        document.getElementById("vMods_FullScreen").checked = false;
      } else {
        document.getElementById("vMods_Width").value = 0;  
      }      
      if (y > 0) {
        document.getElementById("vMods_Height").value = y;  
        document.getElementById("vMods_FullScreen").checked = false;
      } else {
        document.getElementById("vMods_Height").value = 0;  
      }            
    }

    function fScript(type, modId, player) {
      if (type.toLowerCase() == "z" || type.toLowerCase() == "xx" || type.toLowerCase() ==  "fx" || type.toLowerCase() ==  "h") {
        document.getElementById("vMods_Script").value = "";  
      } else {
        document.getElementById("vMods_Script").value = "fmodulewindow";  
      }
    }

    // this was changed on Nov 1, 2016 to NOT lose existing values when changing types
    function fUrl(type, modId, player) {
      if (document.getElementById("vMods_Url").value == "") {
        if (type.toLowerCase() ==  "z") {
          document.getElementById("vMods_Url").value = "/V5/ZMODULES/" + modId + "/IMSMANIFEST.XML";  
        } else  if (type.toLowerCase() == "fx") {
          document.getElementById("vMods_Url").value = "/V5/FMODULES/" + modId + "/IMSMANIFEST.XML";  
        } else  if (type.toLowerCase() == "xx") {
          document.getElementById("vMods_Url").value = "/V5/XMODULES/" + modId + "/IMSMANIFEST.XML";  
        } else  if (type.toLowerCase() == "h") {
          document.getElementById("vMods_Url").value = "/V5/HMODULES/...guid.../" + modId + "/IMSMANIFEST.XML";  
        } else {
          document.getElementById("vMods_Url").value = "";
        };    
      };
    };


    function fClearScript() {
      document.getElementById("vMods_Script").value = "";    
    };

    function fClearFullScreen() {
      $("#vMods_FullScreen")[0].checked = "";    
    };

    function fClearFluid() {
      $("#vMods_Fluid")[0].checked = "";
    };



    function fType (type, modId, player) {
      if (type.toLowerCase() !=  "fx") {      
        document.getElementsByName("vMods_Player")[0].checked = false;
        document.getElementsByName("vMods_Player")[1].checked = false;     
        player = 0;
      }  
      if (type.toLowerCase() ==  "fx" && player == 0) {      
        document.getElementsByName("vMods_Player")[0].checked = false;
        document.getElementsByName("vMods_Player")[1].checked = true;     
        player = 2;
      }  


      if (type.toLowerCase() == "xx") {
        fSize(785,600);
      } else if (type.toLowerCase() ==  "z") {
        fSize(785,600);
      } else if (type.toLowerCase() ==  "fx" && player == 0) {
        fSize(785,575);
      } else if (type.toLowerCase() ==  "fx" && player == 1) {
        fSize(750,475);
      } else if (type.toLowerCase() ==  "fx" && player == 2) {
        fSize(785,600);
      } else {
        fSize(0,0);
      }

      fUrl(type, modId, player);

      fScript(type, modId, player);
    }


    function check(theElement) {
      if (isNaN(theElement.value)) {
        alert("Enter a valid number for this field.");
        theElement.value = "";
        return false;
      }  
    }

    function Validate(theForm){  
      var vOk = false;
      for (i = 0;  i < theForm.vMods_VuCert.length;  i++) {
        if (theForm.vMods_VuCert[i].checked) vOk = true;
      }
      if (!vOk) {
        alert("Please select one of the 'VuAssess Certificate' options.");
        return (false);
      }
      if ((theForm.vMods_ParentId.value.length == 0) || (theForm.vMods_ParentId.value.length == 4 && !isNaN(theForm.vMods_ParentId.value))) {
      } else {
        alert("Please enter a valid 4 digit Parent Id or leave field empty.");
        theForm.vMods_ParentId.value= "";
        theForm.vMods_ParentId.focus();
        return (false);
      }
      return (true);
    }

  </script>
  <style>
    th { width: 30%; }
    .auto-style1 { width: 39%; }
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

  <form method="POST" action="Module.asp" target="_self" onsubmit="return Validate(this)" id="fForm">
    <table class="table">
      <tr>
        <th style="text-align: center" colspan="2">
          <h1>Module Details</h1>
          <h2>
            <% If Len(vMods_Id) = 4 OR Len(vMods_Id) Then %>
              Select the Language for this Module, enter the values then click <b>Update</b>.
            <% Else %>
              Modify any Module values then click <b>Update</b>.
            <% End If %>
            <br />
            <br />
          </h2>
        </th>
      </tr>
      <tr>
        <th class="auto-style1">Module Id :</th>
        <td>
          <% 
            If Len(vMods_Id) = 4 OR Len(vMods_Id) = 5  Then 
            vModsLang        = ""
            vMods_Type   = fDefault(vMods_Type, "FX")
            vMods_Player = fDefault(vMods_Player, 2)
            vMods_Width  = fDefault(vMods_Width, 785)
            vMods_Height = fDefault(vMods_Height, 600)
            vMods_Active = fDefault(vMods_Active, True)
            sModsLangs (vMods_Id)
          %>

          <span style="background-color: yellow; font-weight: bold">
            <%=vMods_Id%>
          </span>

          &ensp;&ensp;&ensp;

          <% If Instr(vMods_Langs, "EN") = 0 Then %>
          <input type="radio" name="vModsLang" value="EN" onclick="fUrl('<%=vMods_Type%>', '<%=vMods_Id & "EN"%>')" <%=fcheck(vModsLang, "en")%>>EN 
          <% End If %>
          <% If Instr(vMods_Langs, "FR") = 0 Then %>
          <input type="radio" name="vModsLang" value="FR" onclick="fUrl('<%=vMods_Type%>', '<%=vMods_Id & "FR"%>')" <%=fcheck(vModsLang, "fr")%>>FR 
          <% End If %>
          <% If Instr(vMods_Langs, "ES") = 0 Then %>
          <input type="radio" name="vModsLang" value="ES" onclick="fUrl('<%=vMods_Type%>', '<%=vMods_Id & "ES"%>')" <%=fcheck(vModsLang, "es")%>>ES 
          <% End If %>
          <% If Instr(vMods_Langs, "PT") = 0 Then %>
          <input type="radio" name="vModsLang" value="PT" onclick="fUrl('<%=vMods_Type%>', '<%=vMods_Id & "PT"%>')" <%=fcheck(vModsLang, "pt")%>>PT 
          <% End If %>

          <% Else %>
          <%=vMods_Id%>
          <% End If %>

          <% = fIf (svMembLevel = 5 And Len(vMods_No) > 0 And (Len(vMods_Id) = 6 OR Len(vMods_Id) = 7), f10 & f10 & f10 & "<span>[Vubiz Internal Module No : " & vMods_No & "]</span>", "") %>

          <input type="hidden" name="vMods_Id" value="<%=vMods_Id%>">
        </td>
      </tr>

      <!--
      <tr>
        <th>Format :</th>
        <td>
          <input type="radio" name="vMods_Format" value="0" <%=fcheck(vMods_Format, "0")%>>Online 
        <input type="radio" name="vMods_Format" value="1" <%=fcheck(vMods_Format, "1")%>>Classroom 
        <input type="radio" name="vMods_Format" value="9" <%=fcheck(vMods_Format, "9")%>>Other
        </td>
      </tr>
      -->


      <tr>
        <th class="auto-style1">Title :</th>
        <td>
          <input type="text" size="72" name="vMods_Title" value="<%=vMods_Title%>"></td>
      </tr>
      <tr>
        <th class="auto-style1">Active :</th>
        <td>
          <input type="radio" name="vMods_Active" value="1" <%=fcheck(fsqlboolean(vmods_active), 1)%>>Yes  
        <input type="radio" name="vMods_Active" value="0" <%=fcheck(fsqlboolean(vmods_active), 0)%>>No 
        </td>
      </tr>

      <!--
    <tr>
      <th>Allow Access to :<font color="#008000"><br>[coming]&nbsp;&nbsp; </font> </th>
      <td>
        <% If svMembLevel < 5 Then %>
        <%=vMods_AccessOk%>
        <input type="hidden" name="vMods_AccessOk" value="<%=vMods_AccessOk%>">
        <% Else %>
        <input type="text" size="35" name="vMods_AccessOk" value="<%=vMods_AccessOk%>" style="width: 500px"> <br>4 Character codes separated by spaces, ie &quot;CCHS 2818&quot;.&nbsp;&nbsp; If entered, access to this module is available by Managers from these Accounts - unless denied below.&nbsp; Leave empty if this module is only available to Administrators.
        <% End If %>
        [Future]
      </td>
    </tr>
    <tr>
      <th>Deny Access to :<font color="#008000"><br>[coming]&nbsp;&nbsp; </font></th>
      <td>
        <% If svMembLevel < 5 Then %>
        <%=vMods_AccessNo%>
        <input type="hidden" name="vMods_AccessNo" value="<%=vMods_AccessNo%>">
        <% Else %>
        <input type="text" size="35" name="vMods_AccessNo" value="<%=vMods_AccessNo%>" style="width: 500px"> <br>4 Character codes separated by spaces, ie &quot;CCHS 2818&quot;.&nbsp;&nbsp; If entered, access to this module is denied to Managers from these Accounts.&nbsp; Leave empty if this module is only available to Administrators.
        <% End If %>
        [Future]
      </td>
    </tr>
-->

      <tr>
        <th class="auto-style1">Parent Id :<br>
          &nbsp;</th>
        <td>
          <input type="text" size="4" name="vMods_ParentId" id="vMods_ParentId" value="<%=Trim(vMods_ParentId)%>" maxlength="4">
          If this module has been cloned, specify the Parent Module Id.</td>
      </tr>
      <tr>
        <th class="auto-style1">Reviewed ?</th>
        <td>
          <input type="radio" name="vMods_Reviewed" value="1" <%=fcheck(fsqlboolean(vmods_Reviewed), 1)%>>Yes 
        	<input type="radio" name="vMods_Reviewed" value="0" <%=fcheck(fsqlboolean(vmods_Reviewed), 0)%>>No<br>
          If Yes then this module has been cleared for migration/upgrade.</td>
      </tr>
      <tr>
        <th class="auto-style1">Description :</th>
        <td>
          <textarea name="vMods_Desc"><%=vMods_Desc%></textarea></td>
      </tr>
      <tr>
        <th class="auto-style1">Outline :</th>
        <td>
          <textarea name="vMods_Outline"><%=vMods_Outline%></textarea></td>
      </tr>
      <tr>
        <th class="auto-style1">Goals :</th>
        <td>
          <textarea name="vMods_Goals"><%=vMods_Goals%></textarea><br>
          Enter goals separated by a double colon (::).</td>
      </tr>
      <tr>
        <th class="auto-style1">Skill Set&nbsp; :</th>
        <td>
          <input type="text" name="vMods_SkillSet" size="72" value="<%=vMods_SkillSet%>" style="width: 500px"><br>
          Separated skills with a double colon (::).</td>
      </tr>

      <% Dim ii : ii = 0
If ii = 1 Then %>
      <tr>
        <th class="auto-style1">Competency :</th>
        <td>
          <select size="1" name="vMods_Competency" multiple>
            <% = spCompTitles ("2818", "EN", vMods_Competency) %>
          </select>
        </td>
      </tr>
      <% End If  %>

      <tr>
        <th class="auto-style1">Content Type :</th>

        <td>
          <strong><em>Note: &quot;Large&quot; Module Ids, ie 10123EN only support FX or XX module types!</em></strong><br />
          <table style="width: 600px;">
            <tr>
              <th colspan="2" style="width: 30px; text-align: center">Type</th>
              <th style="text-align: center"><a href="#" onclick="$('.legacy').toggle()">(more/less)</a></th>
              <th style="text-align: center">Tracking </th>
              <th style="text-align: center" colspan="2">Player</th>
              <th style="text-align: center">Default<br>
                Content Size</th>
              <th style="text-align: center">PopUp<br>
                Script?</th>
              <th style="text-align: center">Need<br>
                Location?</th>
            </tr>

            <tr class="legacy">
              <td style="text-align: center; width: 15px;">
                <input type="radio" name="vMods_Type" value="F" <%=fcheck(vmods_type, "f")%> onclick="fType('f', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center; width: 15px;"><b>F</b></td>
              <td style="white-space: nowrap;">VuBuild Legacy</td>
              <td style="text-align: center; width: 30px;">V5</td>
              <td style="text-align: center; width: 30px;">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">750 x 475</td>
              <td style="text-align: center">Y</td>
              <td style="text-align: center">N</td>
            </tr>

            <tr class="legacy">
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="FS" <%=fcheck(vmods_type, "fs")%> onclick="fType('fs', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center"><b>FS</b></td>
              <td>VuBuild Legacy SCORM</td>
              <td style="text-align: center">V5</td>
              <td style="text-align: center">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">750 x 475</td>
              <td style="text-align: center">Y</td>
              <td style="text-align: center">N</td>
            </tr>

            <tr class="legacy">
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="A" <%=fcheck(vmods_type, "a")%> onclick="fType('a', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center"><b>A</b></td>
              <td>VuBuild Accessible</td>
              <td style="text-align: center">V5</td>
              <td style="text-align: center">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">750 x 475</td>
              <td style="text-align: center">Y</td>
              <td style="text-align: center">N</td>
            </tr>

            <tr class="legacy">
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="X" <%=fcheck(vmods_type, "x")%> onclick="fType('x', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center"><b>X</b></td>
              <td>3rd Party Content</td>
              <td style="text-align: center">V5</td>
              <td style="text-align: center">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">800 x 600</td>
              <td style="text-align: center">Y</td>
              <td style="text-align: center">Y</td>
            </tr>

            <tr class="legacy">
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="U" <%=fcheck(vmods_type, "u")%> onclick="fType('u', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center"><b>U</b></td>
              <td>3rd party Server</td>
              <td style="text-align: center">None</td>
              <td style="text-align: center">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">800 x 600</td>
              <td style="text-align: center">Y</td>
              <td style="text-align: center">Y</td>
            </tr>

            <tr>
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="Z" <%=fcheck(vmods_type, "z")%> onclick="fType('z', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center"><b>Z</b></td>
              <td>Vubiz Legacy HTML</td>
              <td style="text-align: center">RTE</td>
              <td style="text-align: center">Flex</td>
              <td>&nbsp;</td>
              <td style="text-align: center">785 x 600</td>
              <td style="text-align: center">N</td>
              <td style="text-align: center">Y</td>
            </tr>


            <tr>
              <td style="text-align: center" rowspan="2">
                <input type="radio" name="vMods_Type" value="FX" <%=fcheck(vmods_type, "fx")%> onclick="fType('fx', '<%=vMods_Id%>', 0)"></td>
              <td style="text-align: center" rowspan="2"><b>FX</b></td>
              <td rowspan="2">VuBuild SCORM</td>
              <td style="text-align: center">RTE</td>
              <td style="text-align: center" class="d2">Old</td>
              <td style="text-align: center" class="d2">
                <input type="radio" name="vMods_Player" id="vMods_Player" value="1" <%=fcheck(vmods_Player, "1")%> onclick="fType('fx', '<%=vMods_Id%>', 1)"></td>
              <td style="text-align: center">750 X 475</td>
              <td style="text-align: center">N</td>
              <td style="text-align: center">Y</td>
            </tr>
            <tr>
              <td style="text-align: center">RTE</td>
              <td style="text-align: center" class="d2">Flex</td>
              <td style="text-align: center" class="d2">
                <input type="radio" name="vMods_Player" id="vMods_Player" value="2" <%=fcheck(vmods_player, "2")%> onclick="fType('fx', '<%=vMods_Id%>', 2)"></td>
              <td style="text-align: center">785 x 600</td>
              <td style="text-align: center">N</td>
              <td style="text-align: center">Y</td>
            </tr>

            <tr>
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="XX" <%=fcheck(vmods_type, "xx")%> onclick="fType('xx', '<%=vMods_Id%>', 2)"></td>
              <td style="text-align: center"><b>XX</b></td>
              <td>3rd party SCORM</td>
              <td style="text-align: center">RTE</td>
              <td style="text-align: center">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">785 x 600</td>
              <td style="text-align: center">N</td>
              <td style="text-align: center">Y</td>
            </tr>

            <tr>
              <td style="text-align: center">
                <input type="radio" name="vMods_Type" value="H" <%=fcheck(vmods_type, "H")%> onclick="fType('H', '<%=vMods_Id%>', 2)"></td>
              <td style="text-align: center"><b>H</b></td>
              <td>Vubiz 2.0 HTML</td>
              <td style="text-align: center">RTE</td>
              <td style="text-align: center">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="text-align: center">785 x 600</td>
              <td style="text-align: center">N</td>
              <td style="text-align: center">Y</td>
            </tr>

          </table>
        </td>
      </tr>
      <% vMods_VuCert = fDefault(vMods_VuCert, 1) %>
      <tr>
        <th class="auto-style1">Content Size : </th>
        <td>
          <input type="text" name="vMods_Width" id="vMods_Width" size="4" value="<%=vMods_Width%>" onchange="check(this)">Width&nbsp;&nbsp;&nbsp;
          <input type="text" name="vMods_Height" id="vMods_Height" size="4" value="<%=vMods_Height%>" onchange="check(this)">Height&nbsp;&nbsp;&nbsp; 
          <%=f10%>OR Full Screen: 

          <input type="checkbox" name="vMods_FullScreen" id="vMods_FullScreen" value="1" <%=fcheck(fsqlboolean(vmods_fullscreen), 1)%> onclick="fSize(0,0); fClearScript(); fClearFluid();">Popup
          <input type="checkbox" name="vMods_Fluid" id="vMods_Fluid" value="1" <%=fcheck(fsqlboolean(vmods_fluid), 1)%> onclick="fSize(0,0); fClearScript(); fClearFullScreen();">Fluid

          <br>
          Specify the unique size for in-screen content or for full screen select either PopUp or Fluid (best for HTML). 
          For either Full Screen option the PopUp Script below must be empty.</td>
      </tr>
      <tr>
        <th class="auto-style1">PopUp Script : </th>
        <td>
          <input type="text" name="vMods_Script" id="vMods_Script" size="59" value="<%=vMods_Script%>"><br>
          Leave empty for Z, FX, XX, H and Full Screen content. 
          Otherwise enter the javascript function that launches the appropriate popup window. 
          X typically uses a custom launch script while F uses &quot;fmodulewindow&quot;.
          NOTE THAT SCRIPTS ARE CASE SENSITIVE.
          The full Script Set are contained in <a target="_blank" href="/V5/Inc/Launch.js">Launch.js</a>.
        </td>
      </tr>
      <tr>
        <th><span style="background-color: yellow">Content Location</span> : </th>
        <td>
          <span style="background-color: yellow; font-weight: bold">If you modify an existing Module Type, this field will NO LONGER be automatically modified. Rules: </span>
          <br />
          <input type="text" name="vMods_Url" id="vMods_Url" value="<%=vMods_Url%>" style="width: 500px" size="40"><br>

          <table>
            <tr>
              <td style="text-align: center">Z</td>
              <td>Requires imsmanifest.xml file, ie: /V5/zmodules/5505ES/imsmanifest.xml</td>
            </tr>
            <tr>
              <td style="text-align: center">FX</td>
              <td>Requires imsmanifest.xml file, ie: /V5/fmodules/5505ES/imsmanifest.xml</td>
            </tr>
            <tr>
              <td style="text-align: center">XX</td>
              <td>Requires imsmanifest.xml file, ie: /V5/xmodules/&lt;company/sco&gt;/imsmanifest.xml</td>
            </tr>
            <tr>
              <td style="text-align: center">X</td>
              <td>Enter launch page, ie: /company/start.htm</td>
            </tr>
            <tr>
              <td style="text-align: center">U</td>
              <td>Enter full URL where content start page resides on other server, ie: //&lt;company.com/folder/start.htm&gt;</td>
            </tr>
            <tr>
              <td style="text-align: center">H</td>
              <td>Requires imsmanifest.xml file, ie: /V5/hmodules/...guid.../12345FR/imsmanifest.xml</td>
            </tr>
          </table>


        </td>
      </tr>
      <tr>
        <th class="auto-style1">Features : </th>
        <td>
          <input type="checkbox" name="vMods_FeaAud" value="1" <%=fcheck(fsqlboolean(vMods_FeaAud), 1)%>>Audio&nbsp;&nbsp;
          <input type="checkbox" name="vMods_FeaVid" value="1" <%=fcheck(fsqlboolean(vMods_FeaVid), 1)%>>Video&nbsp;&nbsp;
          <input type="checkbox" name="vMods_FeaAcc" value="1" <%=fcheck(fsqlboolean(vMods_FeaAcc), 1)%>>Accessible&nbsp;&nbsp;
          <input type="checkbox" name="vMods_FeaHyb" value="1" <%=fcheck(fsqlboolean(vMods_FeaHyb), 1)%>>Hybrid&nbsp;&nbsp;
          <input type="checkbox" name="vMods_FeaMob" value="1" <%=fcheck(fsqlboolean(vMods_FeaMob), 1)%>>Mobile&nbsp;&nbsp;
        </td>

      </tr>
      <tr>
        <th class="auto-style1">VuAssess Certificate ?</th>
        <td>
          <input type="radio" name="vMods_VuCert" value="1" <%=fcheck(fsqlboolean(vmods_vucert), 1)%>>Yes&nbsp;
            <input type="radio" name="vMods_VuCert" value="0" <%=fcheck(fsqlboolean(vmods_vucert), 0)%>>No&nbsp;
          <br>
          If this module contains an assessment, specify if a certificate is required.
        </td>
      </tr>
      <tr>
        <th class="auto-style1">Include in Completion System ?</th>
        <td>
          <input type="radio" name="vMods_Completion" value="1" <%=fcheck(fsqlboolean(fDefault(vmods_completion, "1")), 1)%>>Yes&nbsp; 
        <input type="radio" name="vMods_Completion" value="0" <%=fcheck(fsqlboolean(fDefault(vmods_completion, "1")), 0)%>>No&nbsp;
          <br>
          The Completion System is designed for customers having multiple locations (Indigo, Cineplex, etc).&nbsp; When configured as such in the Customer Table, this flag determines if this module is to be included in this subsystem.&nbsp; It should be initially set to &quot;Yes&quot; and then only set to &quot;No&quot; when if it has been replaced, eliminated or updated by another module.</td>
      </tr>
      <tr>
        <th class="auto-style1">Length :</th>
        <td>
          <input type="text" name="vMods_Length" size="4" value="<%=fDefault(vMods_Length, 1)%>">
          Hours</td>
      </tr>
      <tr>
        <th class="auto-style1">Memo :</th>
        <td>
          <textarea name="vMods_Memo"><%=vMods_Memo%></textarea><br>
          This is a free form field which can be used for searching.</td>
      </tr>

      <% If bLive Then %>
      <tr>
        <th class="auto-style1">Used in Programs :</th>
        <td><%=spProgByMods (vMods_Id)%></td>
      </tr>
      <tr>
        <th class="auto-style1">Used in MultiScos :</th>
        <td><%=spProgByScos (vMods_Id)%></td>
      </tr>
      <% End If %>

      <!--
    <tr>
      <th>Used by Customers : <font color="#008000"><br>[coming... maybe]&nbsp;&nbsp; </font> </th>
      <td>&nbsp;</td>
    </tr>
-->
      <tr>
        <td style="text-align: center" colspan="2" height="100">
          <br>
          <h2>Ensure you Update before you View the module or its Description. </h2>
          <br>

          <input onclick="jconfirm('Module.asp?vDelModId=<%=vMods_Id%>&vFunction=del', 'Ok to delete?')" type="button" value="Delete" name="bDelete" class="button070"><%=f10%>
          <input type="submit" value="Update" name="bUpdate" class="button070"><%=f10%>
          <input type="button" onclick="location.href='Modules.asp'" value="Return" name="bReturn" id="bReturn" class="button070"><%=f10%>
          <input type="button" onclick="<%=fLaunchUrl%>" value="View" name="bView" class="button070"><%=f10%>
          <input type="button" onclick="SiteWindow('ModuleDescription.asp?vModId=<%=vMods_Id%>&vClose=Y')" value="Description" name="bDescription" class="button070">
        </td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  <script>
    $(".legacy").hide();
  </script>

</body>

</html>
