<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<%
  Dim vAccounts, vInactive, vPrograms, vStrDate, vEndDate, vCurList, vMaxList, vFormat, vBg, vProgram, vAcctId
  vCurList       = fDefault(Request("vCurList"),   0)
  vMaxList       = fDefault(Request("vMaxList"), 100)
  vAccounts      = Request("vAccounts")
  vInactive      = Request("vInactive")
  vStrDate       = Request("vStrDate") 
  vEndDate       = Request("vEndDate") 
  vPrograms      = Request("vPrograms")
  vFormat        = Request("vFormat")
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %>

  <table border="1" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" cellspacing="1" width="1415">
    <tr>
      <td colspan="17" align="left">
      <h1><br>Client Activity Report</h1>
      <h2>The Report shows all selected Programs within the selected Self Service, Group, and/or Custom Sites<% If vStrDate <> "" And vEndDate <> "" Then %> that were setup between <%=vStrDate%> and <%=vEndDate%> <% ElseIf vStrDate <> "" And vEndDate = "" Then %> that were setup after <%=vStrDate%> <% ElseIf vStrDate = "" And vEndDate <> "" Then %> that were setup before <%=vEndDate%> <% End If %>.</h2>
      </td>
    </tr>

    <%
'       --------------------------------- Group Sites ----------------------------------------------

        '...include group sites?
        If Instr(vAccounts, "g") > 0 Then         
    %>

    <tr>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Original <br>Cust Id</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>New <br>Cust Id</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Account <br>Type</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Programs<br>Purchased</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Quantity</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Course</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Max #<br>Learners</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Facilitator Name</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Facilitator Email</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Setup Date</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Expiry Date</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Name</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Organization</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Email</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Phone</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Address1</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Address2</th>
    </tr>

    <%
          vSql = "SELECT  " _   
               & "  Cust.Cust_Id                        AS [Original Cust Id], " _
               & "  Cust_1.Cust_MaxUsers                AS [Max Learners], "_
               & "  vEcom.Ecom_NewAcctId                AS [New Cust Id], " _
               & "  vEcom.Ecom_Media                    AS [Account Type], " _
               & "  vEcom.Ecom_Programs                 AS Programs, " _
               & "  V5_Base.dbo.Prog.Prog_Title1        AS Course, " _
               & "  SUM(vEcom.Ecom_Quantity)            AS Quantity, " _
               & "  vMemb.Memb_FirstName + ' ' + vMemb.Memb_LastName AS [Facilitator Name], " _ 
               & "  vMemb.Memb_Email                    AS [Facilitator Email], " _
               & "  Memb.Memb_FirstVisit                AS [Setup Date], " _ 
               & "  CASE WHEN Cust_1.Cust_Expires < CONVERT(DATETIME, 'Jan 1, 2000') THEN NULL ELSE Cust_1.Cust_Expires END AS [Expiry Date], " _ 
               & "  vEcom.Ecom_CardName                 AS [Ecom CardName], " _ 
               & "  vEcom.Ecom_Organization             AS [Ecom Organization], " _ 
               & "  vEcom.Ecom_Email                    AS [Ecom Email], " _ 
               & "  vEcom.Ecom_Phone                    AS [Ecom Phone], " _
               & "  vEcom.Ecom_Address                  AS [Ecom Address1], " _
               & "  vEcom.Ecom_City + ', ' + vEcom.Ecom_Postal + ', ' + vEcom.Ecom_Province + ', ' + vEcom.Ecom_Country AS [Ecom Address2] " _
    
               & "FROM " _        
               & "  vEcom WITH (nolock) " _ 
               & "    LEFT  OUTER JOIN Memb WITH (nolock) ON vEcom.Ecom_NewAcctId = Memb.Memb_AcctId AND Memb.Memb_Level = 5 " _
               & "    RIGHT OUTER JOIN Cust WITH (nolock) ON vEcom.Ecom_CustId = Cust.Cust_Id AND vEcom.Ecom_Media LIKE 'Group%' " _
               & "    LEFT  OUTER JOIN Cust Cust_1 ON '" & Left(svCustId, 4) & "' + vEcom.Ecom_NewAcctId = Cust_1.Cust_Id " _
               & "    INNER JOIN V5_Base.dbo.Prog WITH (nolock) ON vEcom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id " _
               & "    INNER JOIN vMemb ON vEcom.Ecom_NewAcctId = vMemb.Memb_AcctId " _
               & "    INNER JOIN vMem2 ON vMemb.Memb_AcctId = vMem2.Memb_AcctId AND vMemb.Memb_No = vMem2.Memb_No " _
    
               & "WHERE  " _   
               & "  (Cust.Cust_Id LIKE '" & Left(svCustId, 4) & "[^78]%') " _
               &    fIf(vInactive = "n",    "AND (Cust.Cust_Active = 1) ", "") _  
               &    fIf(vPrograms <> "ALL", "AND (CHARINDEX(vEcom.Ecom_Programs, '" & vPrograms & "') > 0) ", "") _
               & "  AND (Memb.Memb_FirstVisit BETWEEN CAST('" & vStrDate & "' AS DATETIME) AND CAST('" & vEndDate & "' AS DATETIME))" _   
   
               & "  OR "_
    
               & "  (Cust.Cust_Agent = '" & Left(svCustId, 4) & "') "_
               &    fIf(vInactive = "n",    "AND (Cust.Cust_Active = 1) ", "") _  
               &    fIf(vPrograms <> "ALL", "AND (CHARINDEX(vEcom.Ecom_Programs, '" & vPrograms & "') > 0) ", "") _
               & "  AND (Memb.Memb_FirstVisit BETWEEN CAST('" & vStrDate & "' AS DATETIME) AND CAST('" & vEndDate & "' AS DATETIME))" _   
    
               & "GROUP BY " _
               & "  Cust.Cust_Id, " _
               & "  Cust_1.Cust_MaxUsers, " _
               & "  vEcom.Ecom_NewAcctId, " _
               & "  vEcom.Ecom_Media, " _
               & "  vEcom.Ecom_Programs, " _
               & "  V5_Base.dbo.Prog.Prog_Title1, " _
               & "  vMemb.Memb_FirstName + ' ' + vMemb.Memb_LastName, " _
               & "  vMemb.Memb_Email," _
               & "  Memb.Memb_FirstVisit, " _
               & "  Cust_1.Cust_Expires," _  
               & "  vEcom.Ecom_CardName, " _
               & "  vEcom.Ecom_Organization, " _
               & "  vEcom.Ecom_Email," _ 
               & "  vEcom.Ecom_Phone, " _
               & "  vEcom.Ecom_Address, " _
               & "  vEcom.Ecom_City + ', ' + vEcom.Ecom_Postal + ', ' + vEcom.Ecom_Province + ', ' + vEcom.Ecom_Country " _
    
               & "ORDER BY " _
               & "  Cust.Cust_Id, " _ 
               & "  vEcom.Ecom_NewAcctId, " _ 
               & "  vEcom.Ecom_Programs" 
  
'         sDebug
          sOpenDB
          Set oRs = oDB.Execute(vSql)
    
          '...read until either eof or end of group
          Do While Not oRs.Eof
        
            If vCurList Mod 2 = 0 Then vBg = "bgcolor='#F2F9FD'" Else vBg = ""
      %>
      <tr>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Original Cust ID")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=fIf(Len(Trim(oRs("New Cust ID"))) = 4, "CCHS" & oRs("New Cust ID"), "") %></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Account Type")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Programs")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Quantity")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Course")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Max Learners")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Facilitator Name")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Facilitator Email")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=fFormatDate(oRs("Setup Date"))%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=fFormatDate(oRs("Expiry Date"))%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom CardName")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Organization")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Email")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Phone")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Address1")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Address2")%></td>
      </tr>
  
      <%    
            vCurList = vCurList + 1
            oRs.MoveNext
          Loop 
          Set oRs = Nothing

        End If




'       --------------------------------- Custom Sites ----------------------------------------------

        '...include "c" corporate accounts?
        If Instr(vAccounts, "c") > 0 Then  
      %>
  
      <tr>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Cust Id</th>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Master <br>Account</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Account <br>Type</th>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Programs<br>Offered</th>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Quantity</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Course</th>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap># <br>Learners</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Setup Date</th>
        <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Expiry Date</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Organization</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
        <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>&nbsp;</th>
      </tr>
  
      <%
        
        vSql = "  SELECT DISTINCT " _ 
             & "    Cust.Cust_Id                AS [Original Cust Id], " _ 
             & "    Cust.Cust_AcctId            AS [Account Id], " _ 
             & "    Cust.Cust_Title             AS [Cust Title], " _ 
             & "    Memb.Memb_FirstVisit        AS [Setup Date], " _ 
             & "    CASE WHEN Cust_Expires < CONVERT(DATETIME, 'Jan 1, 2000') THEN NULL ELSE Cust_Expires END AS [Expiry Date], " _ 
             & "    vMem1.Memb_Count            AS Quantity, " _ 
             & "    CASE WHEN Prog1.Prog_Id     IS NOT NULL THEN Prog1.Prog_Id     ELSE 'n/a' END AS Programs, " _ 
             & "    CASE WHEN Prog1.Prog_TItle2 IS NOT NULL THEN Prog1.Prog_Title2 ELSE Prog2.Prog_Title1 END AS Titles " _
  
             & "  FROM " _ 
             & "    TskD " _ 
             & "    LEFT OUTER JOIN V5_Base.dbo.Prog Prog2 ON Left(TskD.TskD_Id, 7) = Prog2.Prog_Id " _ 
             & "    INNER JOIN TskH ON TskD.TskD_No = TskH.TskH_No AND LEFT(TskD.TskD_Id, 1) = 'P' " _ 
             & "    RIGHT OUTER JOIN Jobs " _ 
             & "    LEFT OUTER JOIN V5_Base.dbo.Prog Prog1 ON CHARINDEX(Prog1.Prog_Id, Jobs.Jobs_Mods) > 0 " _ 
             & "    RIGHT OUTER JOIN Cust WITH (nolock) " _ 
             & "    LEFT OUTER JOIN Memb WITH (nolock) ON Cust.Cust_AcctId = Memb.Memb_AcctId AND Memb.Memb_Level = 5 " _ 
             & "    INNER JOIN vMem1 WITH (nolock) ON Cust.Cust_AcctId = vMem1.Memb_AcctId ON Jobs.Jobs_AcctId = Cust.Cust_AcctId ON TskH.TskH_AcctId = Cust.Cust_AcctId " _
  
             & "  WHERE " _
             & "    (Cust.Cust_Id LIKE '" & Left(svCustId, 4) & "%') " _ 
             & "    AND (Cust.Cust_Level IN (3, 4)) " _
             &      fIf(vInactive = "n", "AND (Cust.Cust_Active = 1) ", "") _  
             & "    AND (Cust.Cust_Level <> 2) " _ 
             & "    AND (Memb.Memb_FirstVisit BETWEEN CAST('" & vStrDate & "' AS DATETIME) AND CAST('" & vEndDate & "' AS DATETIME))" _   
             &      fIf(vPrograms <> "ALL", "AND (CHARINDEX(Prog1.Prog_Id, '" & vPrograms & "') > 0) ", "") _

             & "    OR " _
             & "    (Cust.Cust_Agent = '" & Left(svCustId, 4) & "') " _  
             & "    AND (Cust.Cust_Level IN (3, 4)) " _
             & "    AND (Memb.Memb_FirstVisit BETWEEN CAST('" & vStrDate & "' AS DATETIME) AND CAST('" & vEndDate & "' AS DATETIME))" _   
             &      fIf(vPrograms <> "ALL", "AND (CHARINDEX(Prog1.Prog_Id, '" & vPrograms & "') > 0) ", "") _
  
             & "  ORDER BY " _ 
             & "    Cust.Cust_Id, Programs " 
  
  
        vSql = "  SELECT" _     
             & "    Cust.Cust_Id            AS [Original Cust Id]," _
             & "    Cust.Cust_AcctId        AS [Account Id]," _ 
             & "    vCustProg_Corp.Title    AS [Cust Title]," _
             & "    vMem3.Memb_FirstVisit   AS [Setup Date]," _  
             & "    Cust.Cust_Expires       AS [Expiry Date]," _
             & "    vMem1.Memb_Count        AS Quantity," _ 
             & "    vCustProg_Corp.Program  AS Programs," _ 
             & "    vCustProg_Corp.Title    AS Titles " _
             & "  FROM" _         
             & "    Cust" _
             & "    INNER JOIN vMem1            ON Cust.Cust_AcctId       = vMem1.Memb_AcctId" _
             & "    INNER JOIN vCustProg_Corp   ON Left(Cust.Cust_Id, 4)  = vCustProg_Corp.Cust" _
             & "    INNER JOIN vMem3            ON Cust.Cust_AcctId       = vMem3.Memb_AcctId" _
             & "  WHERE" _
             & "    (Cust.Cust_Id LIKE 'CCHS%') AND (Cust.Cust_Level IN (3, 4))" _
             & "    OR" _
             & "    (Cust.Cust_Agent  = 'CCHS') AND (Cust.Cust_Level IN (3, 4))" _
             & "  ORDER BY" _ 
             & "    Cust.Cust_Id, Programs " 
  
'       sDebug
        sOpenDB
        Set oRs = oDB.Execute(vSql)
        vAcctId = ""
  
        '...read until either eof or end of group
        Do While Not oRs.Eof    
      
          If vCurList Mod 2 = 0 Then vBg = "bgcolor='#F2F9FD'" Else vBg = ""
    %>
    <tr>
      <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Original Cust Id")%></td>
      <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Account Id")%></td>
      <td align="left"   valign="top" nowrap <%=vbg%>>Custom</td>
      <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Programs") %></td>
      <td align="center"   valign="top" nowrap>N/A</td>      
      <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Titles")%></td>      
      <% If oRs("Account Id") <> vAcctId Then %>
      <td align="center" valign="top" nowrap <%=vbg%>><%=fIf(oRs("Account Id") <> vAcctId, oRs("Quantity"), "")%></td>
      <td align="left"   valign="top" nowrap <%=vbg%>></td>
      <td align="left"   valign="top" nowrap <%=vbg%>>&nbsp;</td>
      <td align="center" valign="top" nowrap <%=vbg%>><%=fFormatDate(oRs("Setup Date"))%></td>
      <td align="center" valign="top" nowrap <%=vbg%>><%=fFormatDate(oRs("Expiry Date"))%></td>
      <td align="left"   valign="top" nowrap <%=vbg%>></td>
      <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Cust Title")%></td>
      <td align="left"   valign="top" nowrap <%=vbg%>></td>
      <td align="left"   valign="top" nowrap <%=vbg%>></td>
      <td align="left"   valign="top" nowrap <%=vbg%>></td>
      <td align="left"   valign="top" nowrap <%=vbg%>></td>
      <% Else %>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <td <%=vbg%> nowrap></td>
      <% End If %>
    </tr>
    <%    
            vAcctId = oRs("Account ID")
            vCurList = vCurList + 1
            oRs.MoveNext
          Loop 
          Set oRs = Nothing

        End If




'       --------------------------------- Self Service Sites ----------------------------------------------

        '...include other channel sites?
        If Instr(vAccounts, "h") > 0 Then 
    %>

   <tr>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Original <br>Cust Id</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>New <br>Cust Id</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Account <br>Type</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Programs<br>Purchased</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Quantity</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Course</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Max #<br>Learners</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Facilitator Name</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Facilitator Email</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Setup Date</th>
      <th bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Expiry Date</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Name</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Organization</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Email</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Phone</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Address1</th>
      <th align="left" bgcolor="#00FFFF" bordercolor="#FFFFFF" nowrap>Ecom Address2</th>
    </tr>
          
  <%
          vSql = "SELECT  " _   
               & "  Cust.Cust_Id AS [Original Cust Id], " _
               & "  Cust_1.Cust_MaxUsers AS [Max Learners], "_
               & "  vEcom.Ecom_NewAcctId AS [New Cust Id], " _
               & "  vEcom.Ecom_Media AS [Account Type], " _
               & "  vEcom.Ecom_Programs AS Programs, " _
               & "  V5_Base.dbo.Prog.Prog_Title1 AS Course, " _
               & "  SUM(vEcom.Ecom_Quantity) AS Quantity, " _
               & "  vMemb.Memb_FirstName + ' ' + vMemb.Memb_LastName AS [Facilitator Name], " _ 
               & "  vMemb.Memb_Email AS [Facilitator Email], " _
               & "  Memb.Memb_FirstVisit AS [Setup Date], " _ 
               & "  CASE WHEN Cust_1.Cust_Expires < CONVERT(DATETIME, 'Jan 1, 2000') THEN NULL ELSE Cust_1.Cust_Expires END AS [Expiry Date], " _ 
               & "  vEcom.Ecom_CardName AS [Ecom CardName], " _ 
               & "  vEcom.Ecom_Organization AS [Ecom Organization], " _ 
               & "  vEcom.Ecom_Email AS [Ecom Email], " _ 
               & "  vEcom.Ecom_Phone AS [Ecom Phone], " _
               & "  vEcom.Ecom_Address AS [Ecom Address1], " _
               & "  vEcom.Ecom_City + ', ' + vEcom.Ecom_Postal + ', ' + vEcom.Ecom_Province + ', ' + vEcom.Ecom_Country AS [Ecom Address2] " _
    
               & "FROM " _        
               & "  vEcom WITH (nolock) " _ 
               & "    LEFT  OUTER JOIN Memb WITH (nolock) ON vEcom.Ecom_NewAcctId = Memb.Memb_AcctId AND Memb.Memb_Level = 5 " _
               & "    RIGHT OUTER JOIN Cust WITH (nolock) ON vEcom.Ecom_CustId = Cust.Cust_Id " _
               & "    LEFT  OUTER JOIN Cust Cust_1 ON 'CCHS' + vEcom.Ecom_NewAcctId = Cust_1.Cust_Id " _
               & "    INNER JOIN V5_Base.dbo.Prog WITH (nolock) ON vEcom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id " _
               & "    INNER JOIN vMemb ON vEcom.Ecom_NewAcctId = vMemb.Memb_AcctId " _
               & "    INNER JOIN vMem2 ON vMemb.Memb_AcctId = vMem2.Memb_AcctId AND vMemb.Memb_No = vMem2.Memb_No " _
    
               & "WHERE  " _   
               & "  (Cust.Cust_Id LIKE 'CCHS[^78]%') " _
               &    fIf(vInactive = "n",    "AND (Cust.Cust_Active = 1) ", "") _  
               &    fIf(vPrograms <> "ALL", "AND (CHARINDEX(vEcom.Ecom_Programs, '" & vPrograms & "') > 0) ", "") _
               & "  AND (Memb.Memb_FirstVisit BETWEEN CAST('" & vStrDate & "' AS DATETIME) AND CAST('" & vEndDate & "' AS DATETIME))" _   

               & "  OR "_
    
               & "  (Cust.Cust_Agent = 'CCHS') "_
               &    fIf(vInactive = "n",    "AND (Cust.Cust_Active = 1) ", "") _  
               &    fIf(vPrograms <> "ALL", "AND (CHARINDEX(vEcom.Ecom_Programs, '" & vPrograms & "') > 0) ", "") _
               & "  AND (Memb.Memb_FirstVisit BETWEEN CAST('" & vStrDate & "' AS DATETIME) AND CAST('" & vEndDate & "' AS DATETIME))" _   
    
               & "GROUP BY " _
               & "  Cust.Cust_Id, " _
               & "  Cust_1.Cust_MaxUsers, " _
               & "  vEcom.Ecom_NewAcctId, " _
               & "  vEcom.Ecom_Media, " _
               & "  vEcom.Ecom_Programs, " _
               & "  V5_Base.dbo.Prog.Prog_Title1, " _
               & "  vMemb.Memb_FirstName + ' ' + vMemb.Memb_LastName, " _
               & "  vMemb.Memb_Email," _
               & "  Memb.Memb_FirstVisit, " _
               & "  Cust_1.Cust_Expires," _  
               & "  vEcom.Ecom_CardName, " _
               & "  vEcom.Ecom_Organization, " _
               & "  vEcom.Ecom_Email," _ 
               & "  vEcom.Ecom_Phone, " _
               & "  vEcom.Ecom_Address, " _
               & "  vEcom.Ecom_City + ', ' + vEcom.Ecom_Postal + ', ' + vEcom.Ecom_Province + ', ' + vEcom.Ecom_Country " _
    
               & "ORDER BY " _
               & "  Cust.Cust_Id, " _ 
               & "  vEcom.Ecom_NewAcctId, " _ 
               & "  vEcom.Ecom_Programs" 

'       sDebug
        sOpenDB
        Set oRs = oDB.Execute(vSql)
        vAcctId = ""
  
        '...read until either eof or end of group
        Do While Not oRs.Eof    
      
          If vCurList Mod 2 = 0 Then vBg = "bgcolor='#F2F9FD'" Else vBg = ""
    %>
      <tr>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Original Cust ID")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=fIf(Len(Trim(oRs("New Cust ID"))) = 4, "CCHS" & oRs("New Cust ID"), "") %></td>
        <td align="left"   valign="top" nowrap <%=vbg%>>Self Service</td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Programs")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Quantity")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Course")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=oRs("Max Learners")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Facilitator Name")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Facilitator Email")%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=fFormatDate(oRs("Setup Date"))%></td>
        <td align="center" valign="top" nowrap <%=vbg%>><%=fFormatDate(oRs("Expiry Date"))%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom CardName")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Organization")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Email")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Phone")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Address1")%></td>
        <td align="left"   valign="top" nowrap <%=vbg%>><%=oRs("Ecom Address2")%></td>
      </tr>
    </tr>
    <%    
            vCurList = vCurList + 1
            oRs.MoveNext
          Loop 
          Set oRs = Nothing

        End If

    %>


    <tr>
      <td valign="top" colspan="17" align="center">
        &nbsp;<form method="POST" action="ClientActivityReport_1.asp">
        <p align="left">&nbsp; 
        <input type="button" onclick="location.href='ClientActivityReport.asp?vStrDate=<%=vStrDate%>&vEndDate=<%=vEndDate%>&vInactive=<%=vInactive%>&vAccounts=<%=vAccounts%>&vCurList=<%=vCurList%>&vPrograms=<%=vPrograms%>&vFormat=<%=vFormat%>'" value="Restart" name="bRestart" class="button"> 
        <% If Cint(vCurList) > 0 And Cint(vCurList) Mod vMaxList = 0 Then '...If next group, get next starting value %> 
          <%=f10%>  
          <input type="hidden" name="vStrDate" value="<%=vStrDate%>">
          <input type="hidden" name="vCurList" value="<%=vCurList%>">
          <input type="hidden" name="vInactive" value="<%=vInactive%>">
          <input type="hidden" name="vAccounts" value="<%=vAccounts%>">
          <input type="hidden" name="vPrograms" value="<%=vPrograms%>">
          <input type="hidden" name="vFormat" value="<%=vFormat%>">
          <input type="submit" name="bNext" value="Next" class="button085"> 
        <% End If %>
        </p>
        </form>
      </td>
    </tr>
  </table>

</body>

</html>