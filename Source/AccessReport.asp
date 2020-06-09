<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<html>

<head>
  <meta charset="UTF-8">

  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>


  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi
    
    '...If first pass then display the drop down form
    If Request.Form("vHidden").Count = 0 Then
  
  %>
  <form method="POST" action="AccessReport.asp">
    <input type="Hidden" name="vHidden" value="Hidden">
    <table border="1" width="100%" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td align="center">
        &nbsp;<h1>Programs Assigned Report</h1>
        <h2>This report displays either a full list of learners and the programs assigned to them or a summary count.</h2>
        <p class="c6">Please be patient, this report can take several minutes.</p>
        <p class="c2">Include Learner Details? <input type="radio" value="y" name="vAll">Yes <input type="radio" value="n" name="vAll" checked>No&nbsp;&nbsp; <input type="submit" value="Go" name="bContinue"></p><p>&nbsp;</p></td>
      </tr>
    </table>
  </form>
  <%
    Else
      Dim vId, vAll
      vAll = Request("vAll")    
  %>
  <div align="center">
    <table border="1" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="2" cellspacing="0">
      <tr>
        <td colspan="3" align="center">
          &nbsp;<h1>Program Assignment Report</h1>
          <% If vAll = "n" Then %>
          <h2>This displays a count of assigned programs.</h2>
          <% Else %>
          <h2>This report displays a full list of learners who have been assigned programs of learning.</h2>
          <% End If %>
        </td>
      </tr>

      <% If vAll = "y" Then %>
      <tr>
        <th align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" width="200">Group</th>
        <th align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" width="200">&nbsp;Learner</th>
        <th align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" width="200">&nbsp;Programs</th>
      </tr>
      <tr>
        <td valign="top" class="c3" width="200">&nbsp;</td>
        <td valign="top" class="c3" width="200">&nbsp;</td>
        <td valign="top" class="c3" width="200">&nbsp;</td>
      </tr>
      <% End If %>

      <%
        Dim vJobs, vProg, vProgs, vProgsCnt, vTotalCnt, aProgs, aTotal, vProgOk, vCritPrev, vCritTitle
  
        vSql = "SELECT * FROM Memb WITH (nolock) WHERE Memb_AcctId = '" & svCustAcctId & "' ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName "
        vCritPrev = ""          
        '...use this to capture all Progs assigned - used to figure totals  
        vProgsCnt = 0
        ReDim aProgs(1, 0)

        vTotalCnt = 0
        ReDim aTotal(1, 0)
        
'       sDebug
        sOpenDb
        Set oRs = oDb.Execute(vSql)
  
        Do While Not oRS.eof
          sReadMemb

          If vMemb_Criteria <> vCritPrev And vCritPrev <> "" Then 
            sSubTotals
            vProgsCnt = 0
            ReDim aProgs(1, 0)
          End If

          vCritPrev  = vMemb_Criteria
          vCritTitle = fCriteria (vMemb_Criteria)

          If Instr(vMemb_Jobs, "P") > 0 Then
            vJobs = vMemb_Jobs
            vProgs = ""          
            Do While Instr(vJobs, "P") > 0
  
              i = Instr(vJobs, "P")
  
              '...isolate the program
              vProg = Mid(vJobs, i, 7)
              '...add to programs used by this user
              vProgs = vProgs & vProg & " "
              '...remove programs extracted
              vJobs = Mid(vJobs, i + 6)
              

              '...first do the sub totals

              '...if prog in array then add 1
              vProgOk = False
              If vProgsCnt > 0 Then
                For j = 0 To vProgsCnt
                  If vProg = aProgs (0, j) Then
                    aProgs(1, j) = aProgs(1, j) + 1
                    vProgOk = True
                    Exit For
                  End If
                Next
              End If
              
              '...else add prog to array (note do not put in slot 0)
              If Not vProgOk Then
                vProgsCnt = vProgsCnt + 1
                ReDim Preserve aProgs (1, vProgsCnt)
                aProgs(0, vProgsCnt) = vProg
                aProgs(1, vProgsCnt) = aProgs(1, vProgsCnt) + 1
              End If


              '...now the overall totals

              '...if prog in array then add 1
              vProgOk = False
              If vTotalCnt > 0 Then
                For j = 0 To vTotalCnt
                  If vProg = aTotal (0, j) Then
                    aTotal(1, j) = aTotal(1, j) + 1
                    vProgOk = True
                    Exit For
                  End If
                Next
              End If
              
              '...else add prog to array (note do not put in slot 0)
              If Not vProgOk Then
                vTotalCnt = vTotalCnt + 1
                ReDim Preserve aTotal (1, vTotalCnt)
                aTotal(0, vTotalCnt) = vProg
                aTotal(1, vTotalCnt) = aTotal(1, vTotalCnt) + 1
              End If



            Loop

            If vAll = "y" Then 
      %>
      <tr>
        <td valign="top" width="200"><%=vCritTitle%></td>
        <td valign="top" width="200">&nbsp;<%=vMemb_LastName & ", " & vMemb_FirstName%></td>
        <td valign="top" width="200">&nbsp;<%=vProgs%> </td>
      </tr>
      <%
  	        End If

          End If
          oRs.MoveNext	        
        Loop
        sCloseDB

        sSubTotals
        vProgsCnt = 0
        ReDim aProgs(1, 0)

        sTotals
      %>

      <tr>
        <td colspan="3" align="center">&nbsp;<p><a <%=fstatx%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp; </p></td>
      </tr>
    </table>
  </div>
  <%
    End If

   Server.Execute vShellLo 

   Sub sSubTotals
  %>
      <tr>
        <td valign="top" class="c3" align="center" colspan="3">
          <table border="1" id="table3" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9">
            <tr>
              <td>
                <div align="right">
                  <table border="0" id="table4" cellspacing="0" cellpadding="0" width="300">
                    <tr>
                      <td align="left" colspan="3"><p class="c1"><%=vCritTitle%></p></td>
                    </tr>
                    <%
                      Dim vCnt
                      vCnt = 0
                      For i = 1 To Ubound(aProgs, 2)
                        vCnt = vCnt + aProgs(1, i)
                    %>
                    <tr>
                      <td align="right" width="200">&nbsp;</td>
                      <td align="right" width="200"><%=aProgs(0, i)%></td>
                      <td align="right" width="200"><%=aProgs(1, i)%></td>
                    </tr>
                    <%
                      Next
                    %>
                    <tr>
                      <td align="right" width="200">&nbsp;</td>
                      <td align="right" width="200">&nbsp;</td>
                      <td align="right" width="200">&nbsp;</td>
                    </tr>
                    <tr>
                      <th align="right" width="200">&nbsp;</th>
                      <th align="right" width="200">Total</th>
                      <th align="right" width="200"><%=vCnt%></th>
                    </tr>
                  </table>
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
  <%
   End Sub

   Sub sTotals
  %>
      <tr>
        <td valign="top" class="c3" align="center" colspan="3">
          <p>&nbsp;</p>
          <table border="1" id="table5" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9">
            <tr>
             <td>
               <div align="right">
                 <table border="0" id="table6" cellspacing="0" cellpadding="0" width="300">
                  <tr>
                    <td align="left" colspan="3"><p class="c1">All Programs Assigned</p></td>
                  </tr>
                  <%
                    Dim vCnt
                    vCnt = 0
                    For i = 1 To Ubound(aTotal, 2)
                      vCnt = vCnt + aTotal(1, i)
                  %>
                  <tr>
                    <td align="right" width="200">&nbsp;</td>
                    <td align="right" width="200"><%=aTotal(0, i)%></td>
                    <td align="right" width="200"><%=aTotal(1, i)%></td>
                  </tr>
                  <%
                    Next
                  %>
                  <tr>
                    <td align="right" width="200">&nbsp;</td>
                    <td align="right" width="200">&nbsp;</td>
                    <td align="right" width="200">&nbsp;</td>
                  </tr>
                  <tr>
                    <th align="right" width="200">&nbsp;</th>
                    <th align="right" width="200">Total</th>
                    <th align="right" width="200"><%=vCnt%></th>
                  </tr>
                  </table>
               </div>
             </td>
            </tr>
          </table>
        </td>
      </tr>
  <%
   End Sub
  %>

</body>

</html>