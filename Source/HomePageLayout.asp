<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Inc/Db_Clus.asp"-->

<%
  Dim vForm

  vForm = Request.Form("vForm")
  If vForm = "Clus" Then
    vClus_Id = Request.Form("vClus_Id")
    sGetClus  
  ElseIf vForm = "Feat" Then
    sGetQueryString
    sExtractClus
    sInsertClus 
    Response.Redirect "Default.asp"            
  Else
    vClus_ID = svCustCluster
    sGetClus  
  End If
%>

<html><head><meta charset="UTF-8"><% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <meta http-equiv="Cache-Control" content="no-cache">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="-1"><meta name="GENERATOR" content="Microsoft FrontPage 6.0"><meta name="ProgId" content="FrontPage.Editor.Document"></head>
<script src="/V5/inc/Functions.js"></script>
<body>

<% Server.Execute vShellHi %>

<div align="center">
  <center>
  <table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td colspan="2"><b><font face="Verdana" size="1">Home Page Layout</font></b><p><font face="Verdana" size="1">Select the Cluster Id, then click &quot;next&quot;, then Edit your values, then click &quot;update&quot;.<br><br>&nbsp;</font></p></td>
    </tr>
    <tr>
      <td align="right" valign="top" width="25%"><b><font face="Verdana" size="1">Cluster Id :&nbsp;&nbsp;&nbsp; </font></b></td>
      <td valign="top">

      <form method="POST" action="HomePageLayout.asp">
        <select size="1" name="vClus_Id"><%=fClusDropdown(vClus_Id)%></select> 
        <input border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" name="I1" type="image"> 
        <input type="hidden" name="vForm" value="Clus">
      </form>

      </td>
    </tr>

    <form method="POST" action="HomePageLayout.asp" target="_top">
      <tr>
        <td valign="top" align="center" width="25%"><font face="Verdana" size="1">&nbsp;</font></td>
        <td valign="top" align="center"><hr size="1" color="#000080"></td>
      </tr>
      <tr>
        <td align="right" valign="top" width="25%"><b><font face="Verdana" size="1">Cluster Id :&nbsp;&nbsp; </font></b></td>
        <td valign="top"><input type="text" name="vClus_Id" size="17" value="<%=vClus_Id%>"><font face="Verdana" size="1"><b><br>Note:</b> if you want to clone a Cluster (ie copy a current cluster info to create a new cluster), simply enter a new Cluster Id into the text box and Update.&nbsp; Be careful you don&#39;t enter an existing Cluster Id or you will wipe out all the old values and replace then with the current values.</font></td>
      </tr>
      <tr>
        <td valign="top" align="center" width="25%"><font face="Verdana" size="1">&nbsp;</font></td>
        <td valign="top" align="center"><hr size="1" color="#000080"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><b><font face="Verdana" size="1">Tab Function&nbsp;&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><b><font face="Verdana" size="1">Active?&nbsp;&nbsp; Tab Text<br>&nbsp;</font></b></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font size="1" face="Verdana">Info Page </font><b><font face="Verdana" size="1">:&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp; <input type="checkbox" name="vClus_Tab1" value="1" <%=fCheck(fSqlBoolean(vClus_Tab1), 1)%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab1_Name" size="36" value="<%=vClus_Tab1_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font size="1" face="Verdana" size="1">My Learning <b>:&nbsp;&nbsp;&nbsp; </b></font></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp; <input type="checkbox" name="vClus_Tab2" value="1" <%=fCheck(fSqlBoolean(vClus_Tab2), 1)%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab2_Name" size="36" value="<%=vClus_Tab2_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font size="1" face="Verdana">My Content </font><b><font face="Verdana" size="1">:&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp; <input type="checkbox" name="vClus_Tab3" value="1" <%=fCheck(fSqlBoolean(vClus_Tab3), 1)%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab3_Name1" size="36" value="<%=vClus_Tab3_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font size="1" face="Verdana">My Bookmarks </font><b><font face="Verdana" size="1">:&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp; <input type="checkbox" name="vClus_Tab4" value="1" <%=fCheck(fSqlBoolean(vClus_Tab4), 1)%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab4_Name1" size="36" value="<%=vClus_Tab4_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><b><font face="Verdana" size="1">&nbsp;</font></b><font size="1" face="Verdana">More Content </font><b><font face="Verdana" size="1">:&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp; <input type="checkbox" name="vClus_Tab5" value="1" <%=fCheck(fSqlBoolean(vClus_Tab5), 1)%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab5_Name1" size="36" value="<%=vClus_Tab5_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font size="1" face="Verdana">Benefits</font><b><font size="1" face="Verdana" size="1"> :&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp; <input type="checkbox" name="vClus_Tab6" value="1" <%=fCheck(fSqlBoolean(vClus_Tab6), 1)%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab6_Name" size="36" value="<%=vClus_Tab6_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font size="1" face="Verdana" size="1">Administration<b> :&nbsp;&nbsp;&nbsp; </b></font></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab7_Name" size="36" value="<%=vClus_Tab7_Name%>"></td>
      </tr>
      <tr>
        <td valign="top" align="right" width="25%"><font face="Verdana" size="1">Sign Off</font><b><font size="1" face="Verdana" size="1"> :&nbsp;&nbsp;&nbsp; </font></b></td>
        <td valign="top"><font size="1" face="Verdana">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="text" name="vClus_Tab8_Name" size="36" value="<%=vClus_Tab8_Name%>"></td>
      </tr>
      <input type="hidden" name="vForm" value="Feat">
      <tr>
        <td valign="top" align="center" colspan="2">
          <p align="center"><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I3" type="image"></p><p>&nbsp;</td>
      </tr>
    </form>

  </table>
  </center>
</div>

<p align="center"><a href="Menu.asp"><img border="0" src="../Images/Icons/Administration.gif" alt="Click here for the Menu"></a></p>

<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body></html>