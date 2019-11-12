<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/RTE.asp"-->

<% 
  Dim vCustId, vMembId, vModsId, vProgId, vAcctId, vMembNo, vProgNo, vModsNo, vMessage, bOk
  Dim vGuidId, vInitId

  vCustId = fDefault(Ucase(Request("vCustId")), "NQCD2399")
  vMembId = fDefault(Ucase(Request("vMembId")), "")
  vProgId = fDefault(Request("vProgId"), "P0000XX")
  vModsId = fDefault(Request("vModsId"), "4189EN")

  vMessage = ""
  bOk = false

  If Request.Form.Count > 0 Then


    vAcctId = fCustAcctId (vCustId)
    vMembNo = fMembNoById (vAcctId, vMembId)

    If vAcctId = "" Then vMessage = vMessage & "<br>Missing or Invalid Customer Id."
    If vMembNo = 0  Then vMessage = vMessage & "<br>Missing or Invalid Learner Id."


		'...get ProgNo/ModsNo if ProgId  = P0000XX
    If vProgId = "P0000XX" Then

      vProgNo = fProgNoById (vProgId)
      vModsNo = fModsNoById (vModsId)
      If vProgNo = 0 Then vMessage = vMessage & "<br>Invalid Program Id."
      If vModsNo = 0 Then vMessage = vMessage & "<br>Invalid Module Id."

    Else

  		'...get ProgNo/ModsNo if ProgId <> P0000XX
  	  vSql = " SELECT"_
  				  & "   ("_
  				  & "     SELECT"_ 
  				  & "     [Prog_Mods_ProgNo]"_
  				  & "     FROM [V5_Base].[dbo].[Prog_Mods]"_
  				  & "     WHERE Prog_Mods_ProgId='P1668EN'  and Prog_Mods_ModsId= '1209EN'"_
  				  & "   ) AS ProgNo,"_
  				  & "   ("_ 
  				  & "     SELECT"_ 
  				  & "     [Prog_Mods_ModsNo]"_
  				  & "     FROM [V5_Base].[dbo].[Prog_Mods]"_
  				  & "     WHERE Prog_Mods_ProgId='P1668EN'  and Prog_Mods_ModsId= '1209EN'"_
  				  & "   ) AS ModsNo"
  		sOpenDb
  	  Set oRs = oDb.Execute(vSql)
  		vMembNo = oRs("MembNo")
  		vProgNo = oRs("ProgNo")
      If vProgNo = 0 Or vModsNo = 0 Then vMessage = vMessage & "<br>Invalid Program/Module Combination."
  		sCloseDb

    End If

    If vMessage = "" Then
			'...send in a dummy 1 minute TS value into the session so it will render in the new LRC
      fRTE vMembNo, vProgNo, vModsNo, "Initialize", Null, Null, Null, Null, Null          
      fRTE vMembNo, vProgNo, vModsNo, "SetValue", "cmi.core.session_time", fRTEts (1), Null, Null, Null 
      fRTE vMembNo, vProgNo, vModsNo, "Terminate", Null, Null, Null, Null, Null

      vGuidId			= "guid" & "_" & vProgNo & "_" & vModsNo
      vInitId			= "init" & "_" & vProgNo & "_" & vModsNo
      Session(vGuidId) = ""
      Session(vInitId) = ""

      vMessage ="<br>Session has been created successfully."
      bOk = True
    End If      


 End If

%>
<html>

  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
    <script>
      function peek(type) {
        if (type == "ass") {
          var peekWindow = window.open('QuickExamList.asp','peek','toolbar=no,width=668,height=590,left=10,top=10,status=no,scrollbars=yes,resizable=yes').focus()
        } else {
//          var peekWindow = window.open('/V5/Default.asp?vSource=close&vCust=<%=vCustId%>&vId=VUV5_ADM&vGoto=Default.asp~3vPage~2RTE_History_F.asp~3vPass~2<%=vMembId%>~1vFrom~2RTE_CreateSession.asp','peek','toolbar=no,width=668,height=590,left=10,top=10,status=no,scrollbars=yes,resizable=yes').focus()
            var peekWindow = window.open('/V5/Default.asp?vSource=close&vCust=<%=vCustId%>&vId=<%=vPassword5%>&vGoto=Default.asp~3vPage~2RTE_History_F.asp~3vPass~2<%=vMembId%>~1vFrom~2RTE_CreateSession.asp','peek','toolbar=no,width=668,height=590,left=10,top=10,status=no,scrollbars=yes,resizable=yes').focus()
        }
      }
    </script>
  </head>

  <body>

    <% Server.Execute vShellHi %>
    <center>
    <form method="POST" action="RTE_CreateSession.asp">
      <table border="1" width="90%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
        <tr>
          <th colspan="2" >&nbsp;<h1 align="center">Create an RTE Session</h1>
          <h2 align="left">Enter the 4 fields to create an empty RTE session with a TS of 1 minute 
            (allowing it to render on the LRC).&nbsp; 
            This will allow you to use the LRC editor to then modify the session accordingly.&nbsp; NO LOG DATA WILL BE TRANSFERRED. 
            Remember if you wish to create a session for a legacy exam score, the Module Id must be a modified Exam Id, 
            ie an Exam Id of 1024EN would be entered as Module Id 1024XN! 
            <a href="javascript:peek('ass')">Click here to render a list of legacy Exam Ids.</a>
          </h2>
          </th>
        </tr>
        <tr>
          <th align="right">Customer Id :</th>
          <td><input type="text" name="vCustId" size="10" value="<%=vCustId%>" style="width:70px"> ie ABCD1234 (can be any account)</td>
        </tr>
        <tr>
          <th align="right">Learner Id :</th>
          <td><input type="text" name="vMembId" size="20" value="<%=vMembId%>"  style="width:230px"> any Learner Id from above account</td>
        </tr>
        <tr>
          <th align="right">Program Id :</th>
          <td><input type="text" name="vProgId" size="7" value="<%=vProgId%>"  style="width:60px"> can be P0000XX (for legacy scores)</td>
        </tr>
        <tr>
          <th align="right"><font color="#FF0000">Module Id :</font></th>
          <td><input type="text" name="vModsId" size="7" value="<%=vModsId%>"  style="width:50px"> ie 1024XN (converted 1024EN Exam Id)</td>
        </tr>
        <tr>
          <td align="center" colspan="2" style="padding:30px;">

            <p style="background-color:<%=fIf(bOk, "aliceblue","yellow")%>; color:<%=fIf(bOk, "blue", "red")%>; padding-bottom:20px; font-weight:bold;"><%=vMessage%></p>
            <input type="submit" value="Create" name="bCreate" class="button070">
            <p><a href="javascript:peek('lrc')">Click here to view the LRC if your Create was successful.</a></p>
          </td>
        </tr>
      </table>
    </form>
    </center>

    <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>

</html>
