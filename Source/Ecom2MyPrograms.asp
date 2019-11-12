<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<% 
  Dim aCatl, vCatl, aProgs, aProg, vLinCnt, vGrpCnt, vBg, vProg_Value, vProg_Name, vMode, vOk, vOnloadScript, vExpires, vProgsOk

  '...assume listing of just "My" programs, but admin can see all programs
  vMode = "My"
  If Request("vMode") = "All" And svMembLevel > 2 Then vMode = "All"

  '...get customer product string
  sGetCust svCustId

  '...get any user and ecom program 
  sGetMemb svMembno
  vEcom_Programs = fEcomPrograms (svCustId, svMembId)
 
  '...Use to see if any lines are printed, if not go to another page
  vLinCnt = 0

  '...keep a list of ok progs so dups don't appear
  vProgsOk = ""

  '...get the Catalogue info
  sGetCatl_Rs svCustId
  Do While Not oRs2.Eof
    vGrpCnt = 0

    '...get catalogue info from catalogue table
    sReadCatl

    '...extract the program strings from the catalogue content string
    aProgs = Split(vCatl_Programs)

    '...process each program
    For j = 0 To Ubound(aProgs) '...aProgs(j): "P1001EN~50~79~23.5~90"
      aProg = Split(aProgs(j), "~") 

      '...get program info from the prog table
      sGetProg aProg(0)

      '...get pricing unless price is 9999  
      vProg_US       = aProg(1)
      vProg_CA       = aProg(2)
      vProg_Duration = aProg(4)

      vOk = False

      '...if facilitator/manager/administrator show all programs that are chargeable except the inactive ones  
      If vMode = "All" And vProg_US > 0 And vProg_US <> 9999 Then
	      vOk = True

      '...else for users prog must be free, purchased (but not via group2) or put onto the member table
'     ElseIf vProg_US = 0 Or Instr(vEcom_Programs, vProg_Id) > 0 Or Instr(vMemb_Programs, vProg_Id) > 0 Then
      ElseIf (vProg_US = 0 And vCust_MaxUsers >=0) Or (Instr(vEcom_Programs, vProg_Id) > 0 And vCust_MaxUsers >=0) Or Instr(vMemb_Programs, vProg_Id) > 0 Then

         '...if free ensure there is no duration or not expired
         If vProg_US = 0 Then
           If vProg_Duration = 0 Then
             vOk = True
           ElseIf DateAdd("d", vProg_Duration, svMembFirstVisit) > Now Then
             vOk = True
           End If
         Else
           vOk = True
         End If

         '...ensure prog only displayed once 
         If vOk Then
           If Instr(vProgsOk, vProg_Id) > 0 Then
             vOk = False
           Else
             vProgsOk = vProgsOk & " " & vProg_Id
           End If
         End If

      End If

     If vOk Then
        vGrpCnt = vGrpCnt + 1
        vLinCnt = vLinCnt + 1
        
        '...determine expiry date...
        
        '...if from the member record
        If Instr(vMemb_Programs, vProg_Id) > 0 Then 
                     
          '...if entered an expirey date
          If fFormatDate(vMemb_Expires) <> " " Then
            vExpires = vMemb_Expires
          '...else a duration  
          Else
            vExpires = DateAdd("d", vMemb_Duration, svMembFirstVisit)
          End If

        '...if from the ecom record
        ElseIf Instr(vEcom_Programs, vProg_Id) > 0 Then 
          k = Instr(vEcom_Programs, vProg_Id) '...is this program in the ecom string?
          l = Instr(k, vEcom_Programs, "|") - 1 '...find the end of the pair
          If l = -1 Then l = Len(vEcom_Programs) '...else get the end of string
          vExpires = Mid(vEcom_Programs, k+8, l-k-7)            
  
        '...if free (then no expirey if duration=0 else expires after firstvisit plus duration)
        ElseIf vProg_US = 0 And vProg_Duration = 0 Then
          vExpires = ""
        '...else create the expirey date
        Else
          vExpires = DateAdd("d", vProg_Duration, svMembFirstVisit)
        End If
        If fFormatDate(vMemb_Expires) = " " Then vMemb_Expires = Now
        
        '...if there is an expirey on the customer file this is the latest available date
        If fFormatDate(vCust_Expires) <> " " Then
          If DateDiff("d", vCust_Expires, vExpires) > 0 Then
            vExpires = vCust_Expires
          End If
        End If
        
        '...On first program, create onload script, generate top HTML
        If vLinCnt = 1 Then 
          vOnLoadScript = "onLoad=""parent.frames[1].location.href='Ecom2MyModules.asp?vProgId=" & vProg_Id & "';" & Chr(34)
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

    <title>My Content</title>
    <base target="Details">
  </head>

  <body <%=vonloadscript%> topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">
  <% Server.Execute vShellHi %>

  <table cellpadding="3" border="0" id="table2" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse">

    <tr>
      <td nowrap valign="top"><img border="0" src="../Images/Ecom/User1.gif"> </td>
      <td align="center"><h1><!--[[-->My Programs<!--]]--></h1><h2 align="left"> <!--[[-->Learning programs available to you are listed by category below. Each program can contain one or more learning modules. Clicking on a program title will display the included modules in the right frame.&nbsp; Access any module by clicking on the module title.<!--]]--></h2>
      
      <% 
        '... let non learners in unless group2 
        If svMembLevel > 2 Then 
          If svMembLevel < 5 And vCust_MaxUsers >= 0 Then         
      %> 
      <h6 align="center">As a <%=fIf(svMembLevel=3, "facilitator", fIf(svMembLevel=4, "manager", "administrator"))%> you can click on &quot;All Programs&quot; to see a complete listing. </h6>
      <h2 align="center"><a <%=fStatX%> target="_self" href="Ecom2MyPrograms.asp"><!--[[-->My Programs<!--]]--></a> | <a <%=fStatX%> target="_self" href="Ecom2MyPrograms.asp?vMode=All"><!--[[-->All Programs<!--]]--></a></h2>
      <% 
          End If      
        End If 
      %> 

      </td>
    </tr>

  </table>

  <table cellspacing="0" cellpadding="3" border="1" id="table2" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse">

<%
        End If

        '...display the catalogue title
        If vGrpCnt = 1 Then
%>

    <tr>
      <td bgcolor="#DDEEF9" valign="top" bordercolor="#FFFFFF" colspan="2"><p class="c1" align="left"><%=vCatl_Title%></p></td>
      <% If vMode <> "All" Then %>
      <td bgcolor="#DDEEF9" valign="top" bordercolor="#FFFFFF" align="center" nowrap><p class="c2"><!--[[-->Expires<br>beginning of...<!--]]--></p></td>
      <% End If %>
    </tr>

<%
        End If
%>

    <tr>
      <td valign="top" align="left" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
      <td valign="top" align="left" width="99%"><p class="c2"><a <%=fStatX%> href="Ecom2MyModules.asp?vProgId=<%=vProg_Id%>"><%=vProg_Title%></a></p></td>
      <% If vMode <> "All" Then %>
      <td valign="top" align="center" class="c2" nowrap><%=fFormatDate(vExpires)%></td>
      <% End If %>
    </tr>
<%  

        End If

      Next
      
      oRs2.MoveNext
    Loop

    '...close off page if details
    If vLinCnt = 0 Then Response.Redirect "Ecom2NoPrograms.asp"

%>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>

