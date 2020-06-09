<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->
<!--#include file = "MyWorld2CodeRoutines.asp"-->

<%
  Dim vFunction, vWs, aMods, vProgId, vExamId, vUrl, vUrlTitle1, vAlt, vAttempts, vMembId

  If Request("vFunction") = "modules" Then
  
    '...function returns one of:
    '   "error"          meaning that this service is not configured properly
    '   "eof"            meaning that program is not on file
    '   otherwise return the table of modules, etc

    '...pass this value in so the session variable Session("MembId") does not screw up live run
    vMembId     = Request("vMembId")

    '...set these to ensure you can access the DB without signing in 
    Session("CustId")     = svCustId
'   Session("MembId")     = "WebService"
    Session("MembId")     = vMembId
    Session("CustAcctId") = svCustAcctId
    Session("HostDb")     = "V5_Vubz"

    sGetCust svCustId

    vProgId     = Request("vProgId")
    sGetProg vProgId
    aMods       = Split(vProg_Mods)
    vUrl        = ""
    vWs         = "<span class='c4'><a href='javascript:toggle(""div_" & vProgId & """)'><b>" & Server.HtmlEncode(vProg_Title) & "</b></a></span>" & vbCrLf
    vUrlTitle1  = Server.HtmlEncode("<!--{{-->Click here to launch the Module<!--}}-->")

    vWs       = "  <table style='BORDER-COLLAPSE: collapse' bordercolor='#ddeef9' cellpadding='2' border='0'>" & vbCrLf
    For k = 0 to Ubound(aMods)

      sGetMods aMods(k)

      If vMods_Active Then

        vUrl = vProgId & "|" & vMods_Id  & "|" & vProg_Test & "|" & vProg_Bookmark & "|" & vProg_CompletedButton
  
        If (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z") Or Ucase(vMods_Type) = "H") And Not vMods_FullScreen Then
          vUrl   = "/V5/LaunchObjects.asp?vModId=" & vUrl & "&vNext=MyWorld2.asp"  
        ElseIf (Ucase(vMods_Type) = "FX" Or Ucase(vMods_Type) = "XX" Or Ucase(vMods_Type) = "Z") Or Ucase(vMods_Type) = "H") And vMods_FullScreen Then
          vUrl   = "javascript:fullScreen('" & vUrl & "')"
        Else
          vUrl   = "javascript:" & vMods_Script & "('" & vUrl & "')"
        End If
  
        vWs = vWs & "    <tr>" & vbCrLf
        vWs = vWs & "      <td width='550'>&ensp;&ensp;&ensp;<a " & fStatX & " href=""" & vUrl & """ title='" & vUrlTitle1 & "'>" & Server.HtmlEncode(vMods_Title) & "</a>"
        If fModsDesc (vMods_Id) Then
          vWs = vWs & " [<a " & fStatX & " href=""javascript:SiteWindow('ModuleDescription.asp?vClose=Y&vModId=" & vMods_Id & "')"" title='" & vUrlTitle1 & "'>" & Server.HtmlEncode("<!--{{-->Description<!--}}-->") & "</a>]"
        End If
        vWs = vWs & "&nbsp;<span class='green'>[" & fModStatusLink (svMembNo, vProg_Id, vMods_Id) & "]</span></td>" & vbCrLf
        vWs = vWs & "    </tr>" & vbCrLf

      End If

    Next

    '...assessment included?
    If Len(Trim(vProg_Assessment)) > 0 Then  

      sGetMods (vProg_Assessment)
      vWs = vWs & "    <tr>" & vbCrLf
      vWs = vWs & "      <td width='550'>&ensp;&ensp;&ensp;"
      vWs = vWs &          fVuAssessLink (vProg_Assessment, "<!--{{-->Examination<!--}}-->", Server.HtmlEncode(vMods_Title))
      vWs = vWs & "      </td>"
      vWs = vWs & "    </tr>" & vbCrLf

    '...platform exam included? ...these should no longer be offered
'    ElseIf Lcase(vProg_Exam) <> "n" Then  
     ElseIf Len(vProg_Exam) > 1 Then  
      Session("CertProg") = vProg_Id
      vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch examination<!--}}-->")
      vWs = vWs & "    <tr>" & vbCrLf
      vWs = vWs & "      <td width='550'>&ensp;&ensp;&ensp;<span class='c2'><a " & fStatX & " href=""javascript:examwindow('" & vProg_Exam & "')"" title='" & vUrlTitle1 & "'>" & "<!--{{-->Examination<!--}}-->" & "</a></span>"
      vExamId = Mid(vProg_Exam, 22, 6)
      If fExamOk(vExamId) Then
      vWs = vWs & "        <span class='green'>[" & fAssessmentStatus (svMembNo, vExamId) & "]</span>"
      End If

      vWs = vWs & "    </tr>" & vbCrLf
    End If

    vWs = vWs & "  </table>" & vbCrLf

    '...reset this as this will be used for the icon
    vUrl   = "'javascript:toggle(""div_" & vProgId & """)'"

    Response.Write vWs

  Else
    Response.Write "error"

  End If  
  
%>


