<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  Dim vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vStrDate, vEndDate, vCredit, vCurList, vMaxList
  Dim vMemb_No, vCardLast, vDateLast, vProgLast 
  Dim vTotTS, vTotBS, vTotNA, vNumTS, vNumBS, vNumNA 
  Dim vTimeSpent, vBestScore, vNoAttempts, vExam_Id


  For Each i in Request.QueryString
'   Response.Write i & " - " & Request(i) & "<br>"
  Next

  vMemb_No = Request("vMemb_No")

  vCurList       = Request("vCurList")
  vMaxList       = Request("vMaxList")
  vStrDate       = Request("vStrDate") 
  vEndDate       = Request("vEndDate") 
  vCredit        = Request("vCredit")
  vFind          = Request("vFind")
  vFindId        = Request("vFindId")
  vFindFirstName = Request("vFindFirstName")
  vFindLastName  = Request("vFindLastName")
  vFindEmail     = Request("vFindEmail")
%>
<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td valign="top" align="center">
      <h1 align="center">Ecommerce Report Card for <br> <%=fMembName (vMemb_No)%></h1>
      <h2>This displays this learner's complete learning activities. <br>Note that Time Spent is summarized while the Best Scores and No of Attempts are averaged for each Program.</h2>
      <table border="1" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
        <tr>
          <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Purchaser</th>
          <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Purchased</th>
          <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Program</th>
          <th align="left" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Module</th>
          <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Time <br>Spent</th>
          <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF">Best <br>Score %</th>
          <th bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF"># <br>Attempts</th>
        </tr>

        <%
          vTotTS = 0 : vTotBS = 0 : vTotNA = 0
          vNumTS = 0 : vNumBS = 0 : vNumNA = 0

          vSql = " SELECT "_  
               & "   Ecom.Ecom_Programs, Ecom.Ecom_Issued, Ecom.Ecom_CardName, "_ 
               & "   V5_Base.dbo.Prog.Prog_Title1, "_
               & "   Memb.Memb_FirstName, Memb.Memb_LastName, " _
               & "   V5_Base.dbo.Mods.Mods_ID, V5_Base.dbo.Mods.Mods_Title, "_
               & "   SUBSTRING(V5_Base.dbo.Prog.Prog_Exam, 22, 6) AS Exam_Id "_
               & " FROM Ecom WITH (nolock)  " _
               & "   INNER JOIN V5_Base.dbo.Prog ON Ecom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id  " _
               & "   INNER JOIN Memb WITH (nolock) ON Ecom.Ecom_MembNo = Memb.Memb_No " _
               & "   INNER JOIN V5_Base.dbo.Mods ON CHARINDEX(V5_Base.dbo.Mods.Mods_ID, V5_Base.dbo.Prog.Prog_Mods) > 0 " _
               & " WHERE (Memb.Memb_No = " & vMemb_No & ") " _
               & " ORDER BY Ecom.Ecom_Programs "

'         sDebug
          sOpenDB
          Set oRs = oDB.Execute(vSql)
  
          '...read until either eof or end of group
          Do While Not oRs.Eof

            If vProgLast <> oRs("Ecom_Programs") AND vProgLast <> "" Then
            
               If vNumBS > 0 Then vTotBS = vTotBS/vNumBS Else vTotBS = 0
               If vNumNA > 0 Then vTotNA = vTotNA/vNumNA Else vTotNA = 0
               
               If fExamOk(vExam_Id) Then 
                 vBestScore  = fBestScore(vMemb_No, vExam_Id)
                 vNoAttempts = fNoAttempts(vMemb_No, vExam_Id)

                 vTotBS = vTotBS + vBestScore  : vNumBS = vNumBS + 1
                 vTotNA = vTotNA + vNoAttempts : vNumNA = vNumNA + 1

        %>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td><%=vExam_ID & " - Examination"%></td>
          <td align="center">&nbsp;</td>
          <td align="center"><%=vBestScore%></td>
          <td align="center"><%=vNoAttempts%></td>
        </tr>
        <%
               End If
        %>
        <tr>
          <td height="30" colspan="4">&nbsp;</td>
          <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" valign="top"><%=fIf(vTotTS > 0, vTotTS, "")%></th>
          <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" valign="top"><%=fIf(vTotBS > 0, vTotBS, "")%></th>
          <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" valign="top"><%=fIf(vTotNA > 0, vTotNA, "")%></th>
        </tr>

        <%
              vTotTS = 0 : vTotBS = 0 : vTotNA = 0
              vNumTS = 0 : vNumBS = 0 : vNumNA = 0
            End If

            vTimeSpent  = fTimeSpent(vMemb_No, oRs("Ecom_Programs"), oRs("Mods_ID"))
            vBestScore  = fBestScore(vMemb_No, oRs("Mods_ID"))
            vNoAttempts = fNoAttempts(vMemb_No, oRs("Mods_ID"))

            If vTimeSpent  > 0 Then vTotTS = vTotTS + vTimeSpent  : vNumTS = vNumTS + 1
            If vBestScore >= 0 Then vTotBS = vTotBS + vBestScore  : vNumBS = vNumBS + 1
            If vNoAttempts > 0 Then vTotNA = vTotNA + vNoAttempts : vNumNA = vNumNA + 1
        %>

        <tr>
          <td><%=fIf(vCardLast <> oRs("Ecom_CardName"), oRs("Ecom_CardName"), "")%></td>
          <td><%=fIf(vDateLast <> fFormatDate(oRs("Ecom_Issued")), fFormatDate(oRs("Ecom_Issued")), "")%></td>
          <td><%=fIf(vProgLast <> oRs("Ecom_Programs"), oRs("Ecom_Programs") & " - " & fLeft(fClean(oRs("Prog_Title1")), 32), "")%></td>
          <td><%=oRs("Mods_ID") & " - " & fLeft(fClean(oRs("Mods_Title")), 32)%></td>
          <td align="center"><%=fIf(vTimeSpent  > 0,vTimeSpent, "") %></td>
          <td align="center"><%=fIf(vBestScore  > 0,vBestScore, "") %></td>
          <td align="center"><%=fIf(vNoAttempts > 0,vNoAttempts, "")%></td>
        </tr>

        <%
            vExam_Id  = oRs("Exam_Id")
            vCardLast = oRs("Ecom_CardName")
            vDateLast = fFormatDate(oRs("Ecom_Issued"))
            vProgLast = oRs("Ecom_Programs")
            oRs.MoveNext
          Loop 

        
          If fExamOk(vExam_Id) Then 

            vBestScore  = fBestScore(vMemb_No, vExam_Id)
            vNoAttempts = fNoAttempts(vMemb_No, vExam_Id)

            vTotBS = vTotBS + vBestScore  : vNumBS = vNumBS + 1
            vTotNA = vTotNA + vNoAttempts : vNumNA = vNumNA + 1

        %>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td><%=vExam_ID & " - Examination"%></td>
          <td align="center">&nbsp;</td>
          <td align="center"><%=vBestScore%></td>
          <td align="center"><%=vNoAttempts%></td>
        </tr>
        <%
           End If
           
           If vNumBS > 0 Then vTotBS = vTotBS/vNumBS Else vTotBS = 0
           If vNumNA > 0 Then vTotNA = vTotNA/vNumNA Else vTotNA = 0

        %>
        <tr>
          <td height="30" colspan="4">&nbsp;</td>
          <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" valign="top"><%=fIf(vTotTS > 0, vTotTS, "")%></th>
          <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" valign="top"><%=fIf(vTotBS > 0, vTotBS, "")%></th>
          <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" valign="top"><%=fIf(vTotNA > 0, vTotNA, "")%></th>
        </tr>

        </table>
      <p><input type="button" onclick="location.href='EcomReportCard1.asp?vStrDate=<%=vStrDate%>&amp;vEndDate=<%=vEndDate%>&amp;vCurList=<%=vCurList%>&amp;vFind=<%=vFind%>&amp;vFindId=<%=vFindId%>&amp;vFindFirstName=<%=vFindFirstName%>&amp;vFindLastName=<%=vFindLastName%>&amp;vFindEmail=<%=vFindEmail%>'" value="Return" name="bReturn" id="bReturn"class="button085"></p>
      </td>
    </tr>


    <tr>
      <td valign="top" align="center">&nbsp;</td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <%

    Function fMembName (vMembNo)
      fMembName = ""
      vSql = "SELECT Memb_FirstName, Memb_LastName FROM Memb WITH (nolock) WHERE Memb_No = " & vMembNo
      sOpenDb2
      Set oRs2 = oDb2.Execute(vSql)
      fMembName = oRs2("Memb_FirstName") & " " & oRs2("Memb_LastName")
      Set oRs2 = Nothing      
      sCloseDb2
    End Function 
  
     Function fClean(i) '...strip off html tags and notes in brackets
       j = Instr(i, "<")
       If j = 0 Then
         fClean = i
       Else
         fClean = Left(i, j-1)
       End If
       j = Instr(i, "(")
       If j > 1Then
         fClean = Left(fClean, j-1)
       End If
     End Function

    Function fTimeSpent(vMembNo, vProgId, vModId)
      sOpenDb3
      vSql = "SELECT RIGHT(Logs_Item, 5) AS TimeSpent FROM Logs WITH (nolock) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (LEFT(Logs_Item, 14) = '" & vProgId & "|" &  vModId & "')"
  '   sDebug
      fTimeSpent = 0
      Set oRs3 = oDb3.Execute(vSql)
      If Not oRs3.Eof Then
        If Len(oRs3("TimeSpent")) > 0 Then
          fTimeSpent = Cint(oRs3("TimeSpent"))
        End If
      End If
      sCloseDb3
      Set oRs3 = Nothing
    End Function


    Function fBestScore (vMembNo, vModId)
      vSql = "SELECT MAX(CAST(Right(Logs.Logs_Item, 3) AS FLOAT)) AS Logs_Grade FROM Logs WITH (nolock)"
      vSql = vSql & " WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (Left(Logs_Item, 6) = '" & vModId & "')"
      sOpenDb3
      Set oRs3 = oDb3.Execute(vSql)
      If oRs3.Eof Then 
        fBestScore = -1
      ElseIf IsNull(oRs3("Logs_Grade")) Then
        fBestScore = -1
      Else
        fBestScore = Cint(oRs3("Logs_Grade"))
      End If
      sCloseDb3
      Set oRs3 = Nothing
    End Function


    Function fNoAttempts(vMembNo, vModId)
      Dim vSql, oRs3
      fNoAttempts= 0
      sOpenDb3
      vSql = "SELECT COUNT(*) AS NoAttempts FROM Logs WITH (nolock) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'T') AND (LEFT(Logs_Item, 6) = '" & vModId & "')"
      Set oRs3 = oDb3.Execute(vSql)
      If Not oRs3.Eof Then fNoAttempts= Cint(oRs3("NoAttempts"))
      sCloseDb3
      Set oRs3 = Nothing
    End Function

    Function fExamOk (vModId)
      vSql = "SELECT * FROM Mods WHERE Mods_Id= '" & vModId & "'"
      sOpenDbBase    
      Set oRsBase = oDbBase.Execute(vSql)
      If oRsBase.Eof Then 
        fExamOk = False
      Else
        fExamOk = True
      End If      
      Set oRsBase = Nothing
      sCloseDbBase    
    End Function


  %>

</body>

</html>