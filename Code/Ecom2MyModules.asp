<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
  If Len(Request("vProgId")) = 0 Then Response.Redirect "Ecom2NoModules.asp"
  Session("Ecom_Prog") = Request("vProgId")
  Session("Ecom_Mods") = ""
  sGetProg Session("Ecom_Prog")

  Function fPassed (vModId)
    Dim vSql, oRs2
    fPassed = False
    sOpenDb2
    vSql = "SELECT * FROM Logs WITH (nolock) WHERE (Logs_MembNo = " & svMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "') AND (Right(Logs_Item, 3) >= '070')"
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then fPassed = True
    sCloseDb2
    Set oRs2 = Nothing
  End Function
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

      <base target="_self">
  </head>

  <body>

  <% Server.Execute vShellHi %>

  <table border="0" width="100%" cellpadding="3" style="border-collapse: collapse">
    <tr>
      <td nowrap>
      <img border="0" src="../Images/Ecom/Modules.gif" width="75" height="67"></td>
      <td align="center">
      <h2 align="left"><!--webbot bot='PurpleText' PREVIEW='Click a module title below to launch the module and start e-learning.&nbsp; If an examination is included in this learning program it will appear below.&nbsp; Click the examination link to access the exam. '--><%=fPhra(000096)%></h2>
      </td>
    </tr>
  </table>


  <table cellspacing="0" border="1" id="table3" width="100%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td colspan="2">
      <h1><%=vProg_Title%></h1>
      <h2><%=vProg_Desc%></h2>
      </td>
    </tr>
    <%
      Dim aMods, vBg, vLine
      aMods = Split(Trim(vProg_Mods), " ")
      For vLine = 0 To Ubound(aMods)
        sGetMods aMods(vLine)
        vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color ever other line        
    %>
    <tr>

      <td width="95%" valign="top" <%=vBg%>>
        <p class="c2">
        <a <%=fStatX%> href="javascript:<%=vMods_Script%>('<%=vProg_Id%>|<%=vMods_Id%>|<%=vProg_Test%>|<%=vProg_Bookmark%>|<%=vProg_CompletedButton%>')"><%=vMods_Title%></a> 
      </td>

      <td align="right" nowrap valign="top" <%=vBg%>><p class="c2">
      
        <% If Len(Trim(vMods_Desc)) > 0 Then %>
        &nbsp;&nbsp;[<a <%=fStatX%> href="Ecom2Module.asp?vModsId=<%=vMods_Id%>"><!--webbot bot='PurpleText' PREVIEW='Desc'--><%=fPhra(000117)%></a>]
        <% End If %>
      
        <%
          '...if assessment included then show link or score?
          If Len(Trim(vMods_AssessmentUrl)) > 0 Then %>
          <% If fPassed(vMods_Id) Then %>
          [Passed]
          <% Else %>
          &nbsp;&nbsp;[<a <%=fStatX%> href="javascript:<%=vMods_AssessmentScript%>('<%=vMods_AssessmentUrl%>','<%=vMods_Id%>')"><!--webbot bot='PurpleText' PREVIEW='Assessment'--><%=fPhra(000073)%></a>]
          <% End If %>
        <% End If %>


      </td>
    </tr>
    <%  
      Next



      If Len(vProg_Assessment) > 0 Then  
        vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"
    %>
    <tr>
      <td width="90%" <%=vBg%> colspan="2"><p class="c1">
        <a <%=fStatX%> href="javascript:assessmentwindow('<%=vProg_Assessment%>')"><!--webbot bot='PurpleText' PREVIEW='Examination'--><%=fPhra(000132)%></a>
      </td>  
    </tr>
    <%

      '...exam included?
      ElseIf Lcase(vProg_Exam) <> "n" Then  
        vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"
        Session("CertProg") = vProg_Id '...use this for custom cert
    %>
    <tr>
      <td width="90%" <%=vBg%> colspan="2"><p class="c1">
        <a <%=fStatX%> href="javascript:examwindow('<%=vProg_Exam%>')"><!--webbot bot='PurpleText' PREVIEW='Examination'--><%=fPhra(000132)%></a>
      </td>  
    </tr>
    <%
      End If

      If vProg_Cert <> 0 Then  
        vBg = "" : If vLine Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"
    %>
    <tr>
      <td width="90%" <%=vBg%> colspan="2"><p class="c1"><a <%=fStatX%> href="javascript:vuwindow('CertCompletion.asp?vProgId=<%=vProg_Id%>&vClose=Y',650,425,100,100,'no','no','no')"><!--webbot bot='PurpleText' PREVIEW='Certificate of Completion'--><%=fPhra(000095)%></a></td>
    </tr>
    <%
      End If
    %>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  </body>
</html>



