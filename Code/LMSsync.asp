<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->

<% 
  Dim vAcctId, vUrl, vScore, vMastery, vExpired, vCompleted, vPosted

  vCust_Id = fDefault(Request("vCust_Id"), svCustId)
  vMemb_Id = fDefault(Request("vMemb_Id"), "pbulloch@vubiz.com")
  vMods_Id = fDefault(Request("vMods_Id"), "1024XN")

  vAcctId  = Right(vCust_Id, 4)

	If Request.Form.Count > 0 Or Request.QueryString.Count > 0 Then

		'...get everything except the score and date
	  vSql = " SELECT DISTINCT *"_
				 & " FROM"_
				 & " 	 Cust AS Cu 																																		INNER JOIN"_
				 & " 	 Memb AS Me ON Cu.Cust_AcctId = Me.Memb_AcctId																	INNER JOIN"_
				 & "   Logs AS L1 ON Me.Memb_No = L1.Logs_MembNo 																			INNER JOIN"_
				 & "   Logs AS L2 ON L1.Logs_AcctId = L2.Logs_AcctId AND L1.Logs_MembNo = L2.Logs_MembNo AND LEFT(L1.Logs_Item, 6) = SUBSTRING(L2.Logs_Item, 9, 6) 		INNER JOIN"_
				 & " 	 V5_Base.dbo.Prog AS Pr ON Pr.Prog_Id = LEFT(L2.Logs_Item, 7) 													INNER JOIN"_
				 & "   V5_Base.dbo.Mods AS Mo ON Mo.Mods_Id = LEFT(L1.Logs_Item, 6)															   "_
				 & " WHERE"_     
				 & " 	 (Me.Memb_AcctId = '" & vAcctId & "') 			AND"_ 
				 & " 	 (Me.Memb_Id = '" & vMemb_Id & "') 					AND"_ 
				 & " 	 (L2.Logs_Type = 'P') 											AND"_ 
				 & " 	 (L1.Logs_Type = 'T') 											AND"_ 
				 & " 	 (LEFT(L1.Logs_Item, 6) = '" & vMods_Id & "')"
	
	' sDebug     
		sOpenDb
	  Set oRs = oDb.Execute(vSql)
		vProg_Id = oRs("Prog_Id")
		sReadCust
		sReadMemb
		sCloseDb

	 	sGetProg vProg_Id
	 	sGetMods vMods_Id

	End If

  If Request.QueryString.Count > 0 Then

		If Request("vScore").Count > 0 Then 
			vScore	 = Request("vScore")
		  vPosted  = Request("vPosted")	
			vMastery = 80      
			If vCust_AssessmentScore > 0 Then vMastery = vCust_AssessmentScore
		  If vProg_AssessmentScore > 0 Then vMastery = vProg_AssessmentScore	
	   	vExpired 		= Null
			vCompleted 	= "incomplete"			
			'...completed/passed?					
			If vScore >= vMastery Then
				vCompleted = "completed"
				If vProg_ResetStatus > 0 Then
					vExpired = fFormatSqlDate(vPosted)
				End If
			End If	
			'...pass the score and completion status, fRTE will generate the appropriate objective entries
			fRTE vMemb_No, vProg_No, vMods_No, "SetValue", "cmi.core.score.raw", vScore, vExpired, vPosted, vCompleted		
		End If
		
		If Request("sesObjId").Count > 0 Then 
			sRTEdelObj Request("sesObjId")
		End If    

  End If

%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <style type="text/css">
  .auto-style1 { text-align: left; }
  </style>
</head>

<body>

  <% Server.Execute vShellHi %> 
  
  <center>
    <form method="POST" action="LMSsync.asp">
      <table border="1" width="90%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
        <tr>
          <th align="right" height="32" colspan="3">&nbsp;<h1 align="center">LMS/RTE Synch</h1>
          <h2 class="auto-style1">Copy a legacy platform exam score from the LMS to the RTE when no RTE session exists for this learner/P0000XX/exam. The Module Id must be the &quot;converted&quot; value of the Exam Id, ie if the Exam Id is 1024EN then enter the Module Id equivalent as 1024XN!</h2>
          <p align="center">&nbsp;</p>
          </th>
        </tr>
        <tr>
          <th align="right">Customer Id :</th>
          <td colspan="2"><input type="text" name="vCust_Id" size="10" value="<%=vCust_Id%>" maxlength="8" class="c2"> ie ABCD1234 (can be any account)</td>
        </tr>
        <tr>
          <th align="right">Learner Id / Password :</th>
          <td colspan="2"><input type="text" name="vMemb_Id" size="20" value="<%=vMemb_Id%>" class="c2"> (Learner Id / Password from above account)</td>
        </tr>
        <tr>
          <th align="right"><i>Module Id</i> :</th>
          <td><input type="text" name="vMods_Id" size="7" value="<%=vMods_Id%>" maxlength="6" class="c2"> ie 1024XN (modifyed Assessment Id of 1024EN)</td>
          <td align="right"><input type="submit" value="Go" name="bGo" class="button"></td>
        </tr>
      </table>
    </form>
    <% If  Request.QueryString.Count > 0 Or Request.Form.Count > 0 Then %>
    <table border="1" cellpadding="10" width="600" style="border-collapse: collapse" bordercolor="#00FFFF" cellspacing="0">
      <tr>
        <td valign="top" align="center" class="c1"><b><% =vMemb_FirstName & " " & vMemb_LastName & "<br>" %> <% =vProg_Id & " (" & vProg_No & ") | " & vMods_Id & " (" & vMods_No & ") | " & vMods_Type & "<br>" & vMods_Title & "<br><br>"%> </b></td>
      </tr>
      <tr>
        <td valign="top" width="50%">
        <div align="right">
          <table border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse" bordercolor="#00FFFF">
            <tr>
              <th align="center" colspan="2" class="c1">LMS</th>
            </tr>
            <tr>
              <th align="center">Posted</th>
              <th align="center">Score</th>
            </tr>
            <% 	
  		        vSql = " SELECT RIGHT(Logs_Item, 3) AS Score, Logs_Posted AS Posted"_
  								 & " FROM Memb INNER JOIN Logs AS Logs ON Memb_No = Logs_MembNo  "_
  								 & " WHERE Memb_AcctId = '" & vAcctId & "' AND Logs_Type = 'T' AND Memb_Id = '" & vMemb_Id & "' AND LEFT(Logs_Item, 6) = '" & vMods_Id & "'"_    
  		             & " ORDER BY Logs_Posted"
  '		        sDebug     
  						sOpenDb
  			      Set oRs = oDb.Execute(vSql)
  			      Do While Not oRs.Eof
  			      	vUrl = "LMSsync.asp"_
  			      	     & "?vCust_Id=" & vCust_Id _
  			      	     & "&vMemb_Id=" & vMemb_Id _
  			      	     & "&vMods_Id=" & vMods_Id _
  			      	     & "&vScore="   & oRs("Score") _
  			      	     & "&vPosted="  & fFormatSqlDate(oRs("Posted"))
  		      %>
            <tr>
              <td align="center"><%=fFormatSqlDate(oRs("Posted"))%></td>
              <td align="center"><a class="c2" href="javascript:jconfirm('<%=vUrl%>', 'Ok to Synch?')"><%=oRs("Score")%></a> </td>
            </tr>
            <%
  							oRs.MoveNext
  						Loop
  						sCloseDb
  					%>
          </table>
        </div>
        </td>
      </tr>
      <tr>
        <td valign="top" align="center">
        <p align="left">To copy a missing score from the LMS to the RTE, click on the LMS score. </p>
        </td>
      </tr>
    </table>
    <% End If %> 
  </center>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


