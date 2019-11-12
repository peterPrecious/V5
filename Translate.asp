<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  If svServer <> "localhost" And svMembLevel <> 5 Then Response.Redirect "/V5/Code/" & svCustCluster & ".asp"

  '...if called from normal page
  If Request.Form.Count = 0 Then
    If Request("vNext").Count = 0 Then Response.Redirect "Message.asp?vMsg=" & Server.UrlEncode("You cannot access this service with providing the calling page name.")
    vPhra_No = Request("vPhraNo")
    sGetPhra (vPhra_No)

  '...if responding to the editor
  Else
    sExtractPhra
    sUpdatePhra
    Response.Redirect "Code/" & Request("vNext")
  End If 
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="<%=svDomain%>/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>:: Translate</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vPhra_EN.value == "")
  {
    alert("Please enter a value for the \"English Phrase\" field.");
    theForm.vPhra_EN.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Translate.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <table border="1" style="border-collapse: collapse" width="100%" id="table1" bordercolor="#DDEEF9" cellpadding="2" cellspacing="0">
      <tr>
        <td class="navTableHeader" colspan="2">
        <h1>Translation Editor.&nbsp; </h1>
        <h2>Update the required phrases then click Update, or simply Return without making any changes.&nbsp; You must always have an English phrase on file - if you leave the other phrases empty the English phrase will appear.</h2>
        </td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Phrase No :</th>
        <td width="70%"><%=vPhra_No%></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">EN :</th>
        <td width="70%"><!--webbot bot="Validation" s-display-name="English Phrase" b-value-required="TRUE" --><textarea rows="9" name="vPhra_EN" cols="57"><%=Server.HtmlEncode(vPhra_EN)%></textarea></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">FR :</th>
        <td width="70%"><textarea rows="10" name="vPhra_FR" cols="57"><%=Server.HtmlEncode(vPhra_FR)%></textarea></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">ES :</th>
        <td width="70%"><textarea rows="10" name="vPhra_ES" cols="57"><%=Server.HtmlEncode(vPhra_ES)%></textarea></td>
      </tr>
      <tr>
        <td nowrap valign="top" colspan="2" align="center"><br><input onclick="javascript:history.back(1)" type="button" value="Return" name="B3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="submit" value="Update" name="B4"><br>&nbsp;</td>
      </tr>
    </table>

    <input type="hidden" name="vPhra_No" value="<%=vPhra_No%>">
    <input type="hidden" name="vNext" value="<%=Request("vNext")%>">

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>








