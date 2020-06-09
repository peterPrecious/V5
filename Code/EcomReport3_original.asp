<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<% 
  Dim vButton, vStrDate, vEndDate, vStrDateErr, vEndDateErr, vChannels, vDetails
  
  '...ensure users and/or facilitators don't try to run this report by bypassing the menu page
  If svMembLevel < 4 Then Response.Redirect "Menu.asp"

  sGetCust svCustId

  '...if Excel then go to Excel version
  vButton = "Online Report"
  If Request.Form("bExcel").Count = 1 Then 
    Response.Redirect "EcomReport3X.asp?vChannels=" & fDefault(Request("vChannels"), "All") & "&vStrDate=" & Server.UrlEncode(Request("vStrDate")) & "&vEndDate=" & Server.UrlEncode(Request("vEndDate"))
  End If
  
  '...defaults to current month
  If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then
    vStrDate  = Request("vStrDate")      : If Len(vStrDate) = 0 Then vStrDate = fFormatSqlDate(MonthName(Month(Now)) & " 1, " & Year(Now))
    vEndDate  = Request("vEndDate")      : If Len(vEndDate) = 0 Then vEndDate = fFormatSqlDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))

  Else
    vStrDate  = fFormatSqlDate(Request("vStrDate")) 
    If Request("vStrDate") = "" Then 
      vStrDate = ""
    ElseIf vStrDate = " " Then
      vStrDate  = Request("vStrDate") 
      vStrDateErr = "Error"
    End If
    vEndDate  = fFormatSqlDate(Request("vEndDate"))
    If Request("vEndDate") = "" Then 
      vEndDate = ""
    ElseIf vEndDate = " " Then
      vEndDate  = Request("vEndDate") 
      vEndDateErr = "Error"
    End If
    If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
      If DateDiff("d", vStrDate, vEndDate) < 0 Then
        vEndDateErr = "Error"
      End If
    End If
  End If

  vChannels = fDefault(Request("vChannels"), "All")
  vDetails  = fDefault(Request("vDetails"), "n")
  

  Function fClean(i)
    If Instr(i, "<") > 0 Then
      fClean = Left(i, Instr(i, "<")-1)
    Else
      fClean = i
    End If
    fClean = fLeft(fClean, 32)
  End Function


  Function fChannels(vChannel)
    Dim vAll, vCust, vCnt
    vCnt = 0
    vAll = ""
    fChannels = ""
    vSql ="SELECT DISTINCT Ecom.Ecom_CustId AS Cust FROM Ecom LEFT OUTER JOIN Cust ON Ecom.Ecom_CustId = Cust.Cust_Id WHERE (LEFT(Cust.Cust_Id, 4) = '" & Left(svCustId, 4) & "') OR (Cust.Cust_Agent = '" & Left(svCustId, 4) & "') ORDER BY Cust"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      vCnt = vCnt + 1
      vCust = oRs("Cust")
      vAll  = vAll & " " & vCust
      fChannels  = fChannels & "<option " & fIf(Instr(vChannel, vCust) > 0, "selected ", "") & "value='" & vCust & "'>" & vCust & "</option>" & vbCrLf
      oRs.MoveNext	        
    Loop
    sCloseDb
  End Function
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Launch.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Ecommerce Report</title>
</head>

<body>

  <% Server.Execute vShellHi %>

  <div align="center">

  <table border="1" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#DDEEF9" width="0">
    <tr>
      <td colspan="11" valign="top" align="center">
      <h1 align="center">Program Sales Report</h1>
      <h2>This report displays the online programs sold during the selected dates<br>for <%=Left(svCustId, 4)%> and any related accounts.</h2>
      <table border="0" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#111111" width="523">
        <form method="POST" action="EcomReport3_original.asp">
          <input type="Hidden" name="vHidden" value="Hidden">
          <tr>
            <th align="right" valign="top" width="30%" nowrap>Select Start Date :</th>
            <td width="68%" nowrap><input type="text" name="vStrDate" size="15" value="<%=vStrDate%>" class="c2"> <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br>ie Jan 1, 2010 (MMM DD, YYYY). <br>Leave empty to start at first sale.</td>
          </tr>
          <tr>
            <th align="right" valign="top" width="30%" nowrap>Select End Date :</th>
            <td width="68%" nowrap><input type="text" name="vEndDate" size="15" value="<%=vEndDate%>" class="c2"> <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br>ie Mar 31, 2010 (MMM DD, YYYY). <br>Leave empty to finish with last sale.</td>
          </tr>
          <%
            i = fChannels(vChannels)
            If Len(i) > 0 Then
          %>
          <tr>
            <th align="right" valign="top" width="30%" nowrap>Select Accounts :</th>
            <td align="left" width="68%" nowrap>
              <select size="6" name="vChannels" multiple class="c2">
                <% If svMembLevel = 5 Then %>
                <option value='Global' <%=fSelect(vChannels, "Global")%>>All Accounts</option>
                <% End If %>
                <option value='All' <%=fSelect(vChannels, "All")%>>All <%=Left(svCustId, 4) %> +</option>
                <%=i%>
              </select> 
              <br>Use Ctrl+Enter to for multiple selections.</td>
          </tr>
          <%
             End If
          %>
          <tr>
            <th align="right" valign="top" width="30%" nowrap>Show Online Details ?</th>
            <td align="left" width="68%" nowrap>
              <input type="radio" value="y" name="vDetails" checked>Yes (should only be used on a month to month basis.)<br><input type="radio" value="n" name="vDetails">No, summary only.<br><br>&nbsp;</td>
          </tr>
          <tr>
            <th align="right" valign="top" width="30%" nowrap>Then click either :</th>
            <td align="left" width="68%" nowrap><input type="submit" value="<%=vButton%>" name="bPrint" id="bPrint" class="button">&nbsp;&nbsp; or ...&nbsp; <input type="submit" value="MS Excel File" name="bExcel" class="button"><p>&nbsp;Note: MS Excel always shows details.</td>
          </tr>
        </form>
      </table>
      </td>
    </tr>

    <% 
      If Request.Form("vHidden").Count = 0 Or vStrDateErr <> "" Or vEndDateErr <> "" Then
    %>

    <% 
      Else
    
        If vDetails = "y" Then 
    %>
    <tr>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Issued</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Channel</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Learner</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Cardholder</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Organization</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">New<br>Channel</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Type</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Source</th>
      <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Program</th>
      <th align="left" height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Title</th>
      <th align="right" height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Quantity </th>
    </tr>
    <% 
        End If 

        Dim vQuantity, vProgsSold, vProgsRefs, vLastProg, vOk
  

'             & "  AND (Ecom.Ecom_NewAcctId = '8401') AND (Ecom_Programs = 'P1258EN') "_


        vSql = "SELECT " _
             & "  Ecom.Ecom_CustId, Ecom.Ecom_Id, Ecom.Ecom_Organization, Ecom.Ecom_Programs, Ecom.Ecom_Quantity, Ecom.Ecom_NewAcctId, Ecom.Ecom_Media, Ecom.Ecom_Source, Ecom.Ecom_Issued, V5_Base.dbo.Prog.Prog_Title1, Ecom.Ecom_Prices, Ecom_MembNo, Ecom.Ecom_FirstName, Ecom.Ecom_LastName, Ecom.Ecom_CardName, Ecom.Ecom_Adjustment " _
             & "FROM "_
             & "  Ecom LEFT OUTER JOIN " _ 
             & "  Cust ON Ecom.Ecom_CustId = Cust.Cust_Id LEFT OUTER JOIN "_
             & "  Memb ON Ecom.Ecom_MembNo = Memb.Memb_No LEFT OUTER JOIN " _ 
             & "  V5_Base.dbo.Prog ON Ecom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id " _
             & "WHERE " _  
             & "  (Ecom_Media <> 'CDs') AND (Ecom_Media <> 'Prods') " _
             &    fIf(vChannels = "All",  " AND ((LEFT(Ecom.Ecom_CustId, 4) = '" & Left(svCustId, 4) & "') OR (Cust.Cust_Agent = '" & Left(svCustId, 4) & "')) ", fIf(vChannels <> "Global", " AND (CHARINDEX(Ecom.Ecom_CustId, '" & vChannels & "') > 0) ", " ")) _
             &    fIf(Len(vStrDate) > 6, " AND (Ecom_Issued >= '" & vStrDate & "') ", " ") _
             &    fIf(Len(vEndDate) > 6, " AND (Ecom_Issued <= '" & vEndDate & "') ", " ") _
             & "ORDER BY "_
             & "  Ecom.Ecom_CustId, Ecom.Ecom_Issued, Ecom.Ecom_Id, Ecom.Ecom_Programs, Ecom.Ecom_Media"

'       sDebug

        vProgsSold = 0
        vProgsRefs = 0

        vLastProg  = ""
  
        sOpenDb
        Set oRs = oDb.Execute(vSql)
        Do While Not oRs.Eof
  
          '...ignore records with same ID and Program (if purchase via Ecom "E" Or bypass ecom "C")
          If oRs("Ecom_Prices") < 0 Or oRs("Ecom_Id") = "0" Or oRs("Ecom_Adjustment") = True Or vLastProg = "" Then
            vOk = True
          ElseIf oRs("Ecom_Id") & "|" & oRs("Ecom_Programs") <> vLastProg Then
            vOk = True
          Else 
            vOk = False
          End If


'...ignore above checks for some reason
vOk = True

          If vOk Then

            vQuantity = Abs(oRs("Ecom_Quantity"))
            If oRs("Ecom_Prices") >= 0 Then
              vProgsSold = vProgsSold + vQuantity
            Else
              vProgsRefs = vProgsRefs + vQuantity
              vQuantity = vQuantity * -1
            End If

            vLastProg = oRs("Ecom_Id") & "|" & oRs("Ecom_Programs")
            If vDetails = "y" Then 
    %>
    <tr>
      <td valign="top" nowrap align="left"><%=fFormatDate(oRs("Ecom_Issued"))%></td>
      <td valign="top" nowrap align="left"><%=oRs("Ecom_CustId")%></td>
      <td valign="top" align="left"><%=fLeft(oRs("Ecom_FirstName") & " " & oRs("Ecom_LastName"), 32)%></td>
      <td valign="top" align="left"><%=fLeft(oRs("Ecom_CardName"), 32)%></td>
      <td valign="top" align="left"><%=fLeft(oRs("Ecom_Organization"), 32)%></td>
      <td valign="top" align="center"><%=oRs("Ecom_NewAcctId")%></td>
      <td valign="top" align="center"><%=oRs("Ecom_Media")%></td>
      <td valign="top" align="center"><%=oRs("Ecom_Source")%></td>
      <td valign="top" align="center"><%=oRs("Ecom_Programs")%></td>
      <td valign="top"><%=fClean(oRs("Prog_Title1"))%></td>
      <td valign="top" align="right"><%=vQuantity%> </td>
    </tr>
    <%
            End If 
            

          End If
          oRs.MoveNext	        
        Loop
        sCloseDB
    %>
    <tr>
      <td valign="top" colspan="11" bgcolor="#DDEEF9">&nbsp;</td>
    </tr>
    <tr>
      <th align="left" colspan="10" height="30"># Programs Sold</th>
      <th align="right" height="30"><%=vProgsSold%> </th>
    </tr>
    <tr>
      <th align="left" colspan="10" height="30"># Programs Refunded</th>
      <th align="right" height="30"><%=vProgsRefs%> </th>
    </tr>
    <tr>
      <th align="left" colspan="10" height="30"># Programs Total</th>
      <th align="right" height="30"><%=vProgsSold - vProgsRefs%> </th>
    </tr>
    <tr>
      <th colspan="11" height="50"><input type="button" onclick="jPrint();" value="<%=bPrint%>" name="bPrint0" id="bPrint0" class="button100"></th>
    </tr>
   

    <% 
      End If 
    %>


    </table>

  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>

</html>

