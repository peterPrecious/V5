<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Keys.asp"-->
<!--#include virtual = "V5\Inc\Db_Prog.asp"-->
<!--#include virtual = "V5\Code\ModuleStatusRoutines.asp"-->

<%
  '...is section 2 locked
  If fIsLocked (8747, svMembNo) Then
    '...see if assessment was passed
    If fBestScore (svMembNo, "9549EN") = 100 Then
      sUnlock 8747, svMembNo
    End If
  End If      

  '...is section 3 locked
  If fIsLocked (8750, svMembNo) Then
    '...see if assessment was passed
    If fBestScore (svMembNo, "9586EN") = 100 Then
      sUnlock 8750, svMembNo
    End If
  End If      
%>

