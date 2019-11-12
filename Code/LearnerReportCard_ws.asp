<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->

<%
  Dim vMsg, vLogsNo, vLogsItem, vScoreProg, vScoreDate, vScoreMods, vScoreValue, vMastery, vExpired, vCompleted, vGuid, vTicks

  vMsg = ""

  '...this updates the timespent value
  If Request("vFunction") = "updateTS" Then

    '...post to LMS	
    vLogsNo   = Clng(Request("vLogsNo"))
    vLogsItem = Right("000000" & Request("vLogsItem"), 6)    '...normal timespent format: P1009EN|0197EN_000001

    sOpenDb

    vSql = "UPDATE Logs SET Logs_Item = LEFT(Logs_Item, 15) + '" & vLogsItem & "' WHERE Logs_No = " & vLogsNo
    oDb.Execute (vSql)

    '...get data to post to RTE
    vSql = "SELECT "_
         & "	Lo.Logs_MembNo AS vMemb_No, Pr.Prog_No AS vProg_No, Mo.Mods_No AS vMods_No, CAST(SUBSTRING(Lo.Logs_Item, 16, 6) AS int) AS vTimeSpentMins "_
         & "FROM "_         
         & "	V5_Vubz.dbo.Logs WITH (NOLOCK)		AS Lo																	INNER JOIN "_
         & "	V5_Base.dbo.Prog WITH (NOLOCK)		AS Pr ON LEFT(Lo.Logs_Item, 7) = Pr.Prog_Id		INNER JOIN "_
         & "	V5_Base.dbo.Mods WITH (NOLOCK)		AS Mo ON SUBSTRING(Lo.Logs_Item, 9, 6) = Mo.Mods_Id "_
         & "WHERE "_     
         & "	(Lo.Logs_No = " & vLogsNo & ")"
    Set oRs = oDb.Execute (vSql)

    vTicks = Cdbl(oRs("vTimeSpentMins") * 60000000)
    vGuid = fRTE (oRs("vMemb_No"), oRs("vProg_No"), oRs("vMods_No"), "GetSessionGuid", Null, Null, Null, Null, Null)
    vSql = "UPDATE vuGoldSCORM.dbo.Session SET sesTotalTime = " & vTicks & " WHERE sesGUID = '" & vGuid & "'"
    oDb.Execute (vSql)

    sCloseDb


  '...this updates the bookmark values in the logs table as long as they are numeric between 1 and 999
  ElseIf Request("vFunction") = "updateBM" Then
    vLogsNo   = Clng(Request("vLogsNo"))
    vLogsItem = Right("000" & Request("vLogsItem"), 3)    '...normal bookmark format: 0706EN_023
    vSql = "UPDATE Logs SET Logs_Item = LEFT(Logs_Item, 7) + '" & vLogsItem & "' WHERE Logs_No = " & vLogsNo
    sOpenDb
    oDb.Execute (vSql)
    sCloseDb
    vMsg = "ok"



  '...this adds a score value
  ElseIf Request("vFunction") = "addSC" Then

    vScoreProg   = Request("vScoreProg")
    vScoreMods   = Request("vScoreId")
    vScoreValue  = Request("vScoreValue")
    vScoreDate   = Request("vScoreDate")

    '...check assessment date
    If fFormatSqlDate(Request("vScoreDate")) = " " Then
      vMsg = "inv date"
    End If
  
    '...check Cust Id
    If vMsg = "" Then
	    sGetCust Request("vCust_Id")
	    If vCust_Eof Then vMsg = "inv cust"
    End If

    '...check Prog Id
    If vMsg = "" Then
	    sGetProg vScoreProg
	    If vProg_Eof Then vMsg = "inv prog"
    End If

    '...check assessment Id
    If vMsg = "" Then
	    sGetMods vScoreMods
	    If vMods_Eof Then 
	    	vMsg = "inv mods"
	   	Else
	   		If Instr(vProg_Mods, vScoreMods) = 0 And vProg_Assessment <> vScoreMods Then 
		    	vMsg = "inv mods"
				End If
	   	End If
    End If

    '...check assessment score
    If vMsg = "" And Not IsNumeric(vScoreValue) Then 
      vMsg = "inv val"
	    If vMsg = "" Then
	      If vScoreValue < 0 or vScoreValue > 100 Then 
	        vMsg = "inv val"
	      End if
	    End If
    End If

    If vMsg = "" Then

			'...post to LMS	
      vLogsItem = vScoreMods & "_" & Right("000" & vScoreValue, 3)
      vSql = " INSERT INTO Logs " _
           & " (Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo, Logs_Posted) " _
           & " VALUES " _
           & " ('" & vCust_AcctId & "', 'T', '" & vLogsItem & "', " & Request("vMemb_No") & ", '" & vScoreDate & "')"
  '   sDebug
      sOpenDb
      oDb.Execute(vSql)
      sCloseDb


			'...post to RTE	
			vMastery = 80      
			If vCust_AssessmentScore > 0 Then vMastery = vCust_AssessmentScore
		  If vProg_AssessmentScore > 0 Then vMastery = vProg_AssessmentScore	
	   	vExpired 		= Null
			vCompleted 	= "incomplete"			
			'...completed/passed?					
			If vScoreValue >= vMastery Then
				vCompleted = "completed"
				If vProg_ResetStatus > 0 Then
					vExpired = fFormatSqlDate(vScoreDate)
				End If
			End If	
	
  		'...pass the score and completion status, fRTE will generate the appropriate objective entries
    	fRTE Request("vMemb_No"), vProg_No, vMods_No, "SetValue", "cmi.core.score.raw", vScoreValue, vExpired, vScoreDate , vCompleted		
    	fRTE Request("vMemb_No"), vProg_No, vMods_No, "Terminate", Null, Null, Null, Null, Null

      vMsg = "ok"
    End If

  Else

    vMsg = "err"
 
  End If
  
  
  Response.Write vMsg
%>


