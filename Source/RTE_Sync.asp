<%@  codepage="65001" %>
<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim vCnt, vSessionId, vMembId, vProgId, vModsId

  vSessionId = fDefault(Request("vSessionId"), 0)

  vMembId = Request("vMembId")
  vProgId = Request("vProgId")
  vModsId = Request("vModsId")

  sOpenCmd
  With oCmd
    .CommandText = "spRTE_Sync"
    .Parameters.Append .CreateParameter("@MembNo", adInteger, adParamInput, , vSessionId)
  End With
  Set oRs = oCmd.Execute()
  vCnt = oRs("Updated")
  Set oCmd = Nothing
  sCloseDb

%>

<html>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <script type="text/javascript" src="/V5/Inc/jQuery.js"></script>
    <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
  </head>

  <body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

      <fieldset style="padding:40px;">
        <legend>Sync LMS with RTE FX session items</legend>

          <br/>
          <h1 style="text-align:center"> <br />The LMS log items have been deleted and<br />replaced <%=vCnt%> RTE session items for...<br /><br /><%=vMembId%> &nbsp;| <%=vProgId %> | <%=vModsId %><br /><br /></h1>
          <br/>


      </fieldset>

  </body>

</html>
