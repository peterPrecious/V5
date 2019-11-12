<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_LogX.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->
<!--#include virtual = "V5/Inc/ProgramStatusRoutines.asp"-->
<!--#include file = "Code\ModuleStatusRoutines.asp"-->

<%
  Dim vModId, aModId, vProgId, vPageNo, vTest, vBookmark, vCompletedButton, vUrl, vPreviewMax, vPreviewFX, vUpdatePage, vFolder, vNext, vScorm, aGuid, vExpires, vShowTree, vBuild, vMembNo
  Dim vSessionUrl

  Session.Timeout = 60 * 6 '...(this should now be set in SignIn.asp - no longer needed)

  '...get module and page info info
  vModId           = Request.QueryString("vModId")
  vPageNo          = Request.QueryString("vPageNo")
  vBookmark        = fDefault(Request.QueryString("vBookmark"), "Y")
  vCompletedButton = fDefault(Request.QueryString("vCompletedButton"), "N")
  vNext            = Request("vNext") 
  vBuild           = fIf(Request.QueryString("vBuild") = "Y", "Y", "N")  
  vMembNo          = svMembNo   '...normally we use the caller but if no bookmarking/scorm=0 we create a dummy user

  If Len(vNext) > 0 Then 
    vNext = Server.UrlEncode("/V5/Code/" & vNext)
    '...if there is no vNext then don't include this for the RTE
    vSessionUrl = "&SessionReturnURL=" & vNext
  Else
    vSessionUrl = ""
  End If

  '...for Scorm SSCOs you can turn on the tree view by adding "showtree=1" - currently only used in qmodid (vQmodId=P1234EN|12345EN^)
  vShowTree = fIf(Request.QueryString("showtree") = "1", "&showtree=1", "")  

  '...determine if we use the /modules or /review folders to access modules
  If Ucase(Request("vReview")) = "Y" Then
    vFolder = "Review"
  Else
    vFolder = "Modules"
  End If
 
  vProgId     = ""
  vTest       = ""
  vPreviewMax = 0 

  '...break down the vModid (P1234EN|x9876EN|N|Y|N)
  '   where P1234EN is program (optional)
  '   where  9876EN is module (mandatory) - NOTE: as of May 2014 modules are 7 chars, ie 100012EN
  '   where       x is optional no of preview pages: a-z (optional)
  '   where       N is the test/sa flag:          Y/N (optional but must precede bookmark flag)
  '   where       Y is the bookmark flag:         Y/N (optional but must follow test/sa flag)
  '   where       N is the completed button flag: Y/N (optional but must follow bookmark flag)
  '   all must be separated by pipes           

  If Ucase(Left(vModId, 1)) = "P" Then
    vProgId = Left(vModId, 7) 
    vModId  = Mid(vModId, 9) '...strip off program id and the "|"
  End If

  '...see if pipes to extract test/bookmark/completed button values - must be either 2 or none
  aModId = Split(vModId, "|")
  If Ubound(aModId) >= 2 Then '...ie at least 2 bars
    vModId    = aModId(0)
    vTest     = aModId(1)
    vBookmark = aModId(2)
  End If
  If Ubound(aModId) = 3 Then '...ie includes 3 bars, get completed button
    vCompletedButton = aModId(3)
  End If

  '...for Scorm turn off tracking using scorm=0 - currently only turn off if we don't need bookmarking (samplers, qmodid etc)
  '  vScorm = fIf(vBookmark = "N", "&scorm=0", "") ...modified Feb 23, 2015 to allow 3rd party courses to use dummy learners since scorm=0 has no effect
  If vBookmark = "N" Then
    vScorm = "&scorm=0"
    Randomize 
    vMembNo = Int((1000000) * Rnd + 1) * -1
  Else
    vScorm = ""
  End If


  '...the basic module id format is: PPPPPP|RMMMMMM_B where:
  '   PPPPPP is an option program code (7 chars plus "|") - used in Content.htm but not ModuleEdit.asp
  '   R is an optional preview code (1 char if present (a-z) where a = 1 page, b = 2 pages, etc)
  '   MMMMMM is the module code
  '   _B is a sampler code for using the Vubuild Server (not sure this is still valid - code comes from qmodid/signin.asp here with &vBuild=Y


  '...see if preview code present, and if so, strip off
  If Len(vModId) = 7 Or Len(vModId) = 8 Then
    vPreviewMax  = Asc(Ucase(Left(vModId, 1))) - 65 + 1       
    If vPreviewMax > 0 and vPreviewMax < 27 Then              '...is first char A-Z?
      vPreviewFX   = Mid(vModId, 2) & Left(vModId, 1)         '...for FX player pass C1234EN as 1234ENC
      vModId       = Mid(vModId, 2)
      Session("PreviewMax_" & vModId) = vPreviewMax           '...put in session variable for F Module player     
    Else
     vPreviewFX  = vModId     
    End If
  End If

  '...get Mods/Prog data
  sGetMods (vModId)
  vMods_Type = Ucase(vMods_Type)

  If Len(vProgId) <> 7 Then vProgId = "P0000XX"
  sGetProg (vProgId)

	'...if recurring, pass the expiry date when an FX modules will expire the session upon completion
	vExpires = ""
	If vProg_ResetStatus > 0 Then	
		vExpires = fFormatSqlDate(DateAdd("d", vProg_ResetStatus, Now()))
	End If


  '...if completed button required on send if module is NOT completed
  If vCompletedButton = "Y" Then
    '...use this routine to return date module was last completed - it handles reset values (Code/ModuleStatusRoutines.asp)
    If Len(fCompleted (svMembNo, vModId)) > 0 Then
      vCompletedButton = "N"
    End If
  End If


  '...if we are using bookmarks, then see if any bookmarks on file and get starting pageno
  If vBookmark = "Y" Then
    If fNoValue(vPageNo) And Not fNoValue(svMembNo) Then 
      vPageNo = 0
      sOpenDb
      vSql = "SELECT Logs_No, Logs_Item FROM Logs WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_Type = 'B' AND Logs_MembNo = " & svMembNo & " AND Left(Logs_Item, 6) = '" & vModId & "'"
  '   sDebug "<br>vSql", vSql  
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then 
        '...for vubiz mods the page no (bookmark) will be numeric (from 1-999)
        '   for scorm, it can be anything, ie "page42"
        vPageNo = Mid(oRs("Logs_Item"), 8)
        If IsNumeric(vPageNo) Then vPageNo = Clng(vPageNo)
      End If
      sCloseDb
    End If
  End If


  If vBookmark = "Y" Then
	  fLogTimespent vProgId, vModId, 1 
	End If


  '...Initiate an RTE session for NON Scorm content (unless no bookmarking)
  If vBookmark = "Y" And (vMods_Type <> "FX" And vMods_Type <> "XX" And vMods_Type <> "Z" And vMods_Type <> "H") Then
    i = fRTE (svMembNo, vProg_No, vMods_No, "Initialize", Null, Null, Null, Null, Null)  
  End If


  If Len(Trim(vMods_Title)) = 0 Then vMods_Title = "Learning Module"
  vMods_Title= Replace(vMods_Title, " ", "+")     


  '...If a sampler wanting access the vubuild fModules (Mike setup a special web to access this content)
  If vBuild = "Y" Then
    vUrl = "/vuBuildfModules/" & vModId & "/LMSStart.html?vModId=" & vPreviewFX & "&jumpto=" & Request("jumpto") & "&psi=" & Request("psi") & "&pei=" & Request("pei") & "&deployment=vubuild&scorm=0"
 
  '...If "f" get modules in fModules using V5 communication
  ElseIf Lcase(vMods_Url) = "f" Or vMods_Type = "F" Then
    vUrl = "f" & vFolder & "/" & vModId & "/default.asp?vModType=f&vModId=" & vModId & "&vProgId=" & vProgId & "&vPageNo=" & vPageNo & "&vTest=" & vTest & "&vBookmark=" & vBookmark & "&vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&mode=" & vPreviewMax & "&vEmail=" & svMembEmail & "&vTitle=" & vMods_Title & "&vCompletedButton=" & vCompletedButton

  '...If "fs" get modules in fModules using Scorm communication (vURL is generated within Wrapper to call indexScorm.asp)
  ElseIf Lcase(vMods_Url) = "fs" Or vMods_Type = "FS" Then
    vUrl = "/V5/Wrapper/Wrapper.asp?vModType=f&vModId=" & vModId & "&vProgId=" & vProgId & "&vPageNo=" & vPageNo & "&vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vEmail=" & svMembEmail & "&vTitle=" & vMods_Title & "&vLessonStatus=" & fLessonStatus(svMembNo, vModId) & "&deployment=scorm"

  '...If "a" get local HTML vuBuild accessible modules in aModules
  ElseIf Lcase(vMods_Url) = "a" Or vMods_Type = "A"  Then
    vUrl = "a" & vFolder & "/" & vModId & "/default.htm?vModType=a&vModId=" & vModId & "&vProgId=" & vProgId & "&vPageNo=" & vPageNo & "&vTest=" & vTest & "&vBookmark=" & vBookmark & "&vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&mode=" & vPreviewMax & "&vEmail=" & svMembEmail & "&vTitle=" & vMods_Title & "&vCompletedButton=" & vCompletedButton

  '...If "u" launch from another get local Flash vuBuild modules in fModules
  ElseIf Instr(Lcase(vMods_Url), "http") > 0 Or vMods_Type = "U"  Then
    vUrl = "/V5/Wrapper/Wrapper.asp?vModType=u&vModId=" & vModId & "&vModUrl=" & vMods_Url & "&vProgId=" & vProgId & "&vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vEmail=" & svMembEmail & "&vTitle=" & vMods_Title

  '...If "fx" get modules in fModules, if "xx" get third party modules from xModules using new Scorm communication and new Player  **** Note as of Feb 25, 2015 we are using vMembNo rather than svMembNo (see bookmark above)
  ElseIf vMods_Type = "FX" Or vMods_Type = "XX" Or vMods_Type = "Z" Or vMods_Type = "H" Then
    vUrl = "/Gold/vuSCORM/SCOContainer.aspx?vCustId=" & svCustId & "&vMods_No=" & vMods_No & "&vProg_No=" & vProg_No & "&vMemb_No=" & vMembNo & "&vLang=" & svLang & "&vModId=" & vPreviewFX & vShowTree & vScorm & vSessionUrl & "&jumpto=" & Request("jumpto") & "&psi=" & Request("psi") & "&pei=" & Request("pei") 

  '...if "x" launch third party modules on our server
  Else
    vUrl = "/V5/Wrapper/Wrapper.asp?vModType=x&vModId=" & vModId & "&vModUrl=" & "x" & vFolder & "/" & vMods_Url & "&vProgId=" & vProgId & "&vPageNo=" & vPageNo & "&vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vEmail=" & svMembEmail & "&vTitle=" & vMods_Title & "&vLessonStatus=" & fLessonStatus(svMembNo, vModId)

  End If

'  Response.Write "<P>" & vUrl
  Response.Redirect vUrl
%>