<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_QCust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<!-- load the tabs to refresh any change in tabs - select tab=9 which means admin tab -->
<body onload="parent.frames.tabs.location.href='tabslive.asp?vTab=9'" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
	Server.Execute vShellHi
  Dim vAddCustId, vEditCustId, vFunction, vLeft, vRight, vLevel, vOk
  vFunction = Request("vFunction")

  If vFunction = "add" Then
    sExtractCust

    sInsertCorporate
    Session("CustLevel") = 1
  ElseIf vFunction = "edit" Then
    sExtractCust
    sUpdateCust
    Session("CustLevel") = 1
  ElseIf Len(Request("vDelCustId")) = 8 Then 
    vCust_Id = Request("vDelCustId")
    vCust_AcctId = Request("vDelCustAcctId")
    sDeleteCust
  End If  

  If Len(Request("vHidden")) = 0 Or Request("vFunction") = "del" Then
  
%>
  <!-- Add Customer -->
  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vAddCustId.value == "")
  {
    alert("Please enter a value for the \"New Customer Id\" field.");
    theForm.vAddCustId.focus();
    return (false);
  }

  if (theForm.vAddCustId.value.length < 4)
  {
    alert("Please enter at least 4 characters in the \"New Customer Id\" field.");
    theForm.vAddCustId.focus();
    return (false);
  }

  if (theForm.vAddCustId.value.length > 4)
  {
    alert("Please enter at most 4 characters in the \"New Customer Id\" field.");
    theForm.vAddCustId.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
  var checkStr = theForm.vAddCustId.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter characters in the \"New Customer Id\" field.");
    theForm.vAddCustId.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="CorporateEdit.asp" target="_self" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <input type="hidden" name="vHidden" value="Y">
    <table border="1" width="100%" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="3" id="table4">
      <tr>
        <td align="center" width="90%" rowspan="3"><h1>Quick Corporate Site Setup</h1><h2>To Add a new Corporate site, enter a 4 character Customer Id at right.&nbsp; <br>Note: use an existing Id for existing Customers otherwise create a new Id avoiding the letters &quot;IOL&quot;.&nbsp; <br>The next available Customer No (4xxx) will be generated and display below.</h2><h2>Display corporate sites by <a <%=fStatX%> href="CorporateEdit.asp?vSort=C">first 4 characters</a>&nbsp; |&nbsp; <a <%=fStatX%> href="CorporateEdit.asp?vSort=N">last 4 numbers</a><br>If the catalogue is selected you can click on the Facilitator or Manger link to <br>access the site as they would see it.</h2></td>
        <th nowrap colspan="2" valign="bottom"><p class="c1">Add a new Corporate Site</p>
        </th>
      </tr>
      <tr>
        <th align="right" nowrap>Customer Id:</th>
        <td align="left" nowrap>&nbsp;<!--webbot bot="Validation" s-display-name="New Customer Id" s-data-type="String" b-allow-letters="TRUE" b-value-required="TRUE" i-minimum-length="4" i-maximum-length="4" --><input type="text" name="vAddCustId" size="3" maxlength="4"></td>
      </tr>
      <tr>
        <th align="right" nowrap colspan="2" valign="top"><input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="I6" type="image"></th>
      </tr>
    </table>
  </form>
  <!-- Select Criteria for List -->
  <table width="100%" border="1" cellspacing="0" bordercolor="#DDEEF9" cellpadding="0" style="border-collapse: collapse">
    <tr>
      <th bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF" align="left">Cust Id</th>
      <th align="left" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF">Title</th>
      <th align="left" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF">Lang</th>
      <th align="left" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF">Url</th>
      <th align="left" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF">Catalogue</th>
      <th align="left" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF">Facilitator</th>
      <th align="left" bgcolor="#DDEEF9" height="20" bordercolor="#FFFFFF">Manager</th>
    </tr>
    <%
      '...read Cust
      If Request("vSort") = "N" Then
        sGetCust_Rs_AcctId
      Else
        sGetCust_Rs
      End If
      vOk = False
      i = 0
      Do While Not oRs.Eof
        sReadCust

        If Left(vCust_AcctId, 1) = "4" Then
    %> <tr>
      <td><a <%=fStatX%> href="ChannelEdit.asp?vEditCustId=<%=vCust_Id%>&vHidden=n"><%=vCust_Id%></a></td>
      <td><%=vCust_Title%></td>
      <td><%=vCust_Lang%></td>
      <td><%=vCust_Url%></td>
      <td><% If Len(vCust_Catalogue) > 0  Then %> <%=vCust_Catalogue%> <% Else %> [No Catalogue] <% End If %> </td>
      <td><a <%=fStatX%> target="_top" href="../Default.asp?vCust=<%=vCust_Id%>&vId=<%=vCust_FacilitatorId%>"><%=vCust_FacilitatorId%></a> </td>
      <td><a <%=fStatX%> target="_top" href="../Default.asp?vCust=<%=vCust_Id%>&vId=<%=vCust_ManagerId%>"><%=vCust_ManagerId%></a> </td>
    </tr>
    <%  
        End If
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb    
    %> <tr>
      <td valign="Top" align="center" colspan="7"><br><br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><p><a <%=fStatX%> href="ChannelEdit.asp?vEditCustId=<%=svCustId%>&vHidden=n">Customer Profile</a><br>&nbsp;</p>
      </td>
    </tr>
  </table>
  <!-- Customer Details --><%
  Else  

    Dim aProgs, aProg, vProgram
  
    If Request("vAddCustId").Count > 0 Then 
      vCust_Id = Ucase(Request("vAddCustId")) & fNextCustNo (4)
      vFunction = "add"
    ElseIf Len(Request.QueryString("vEditCustId")) = 8 Then 
      vCust_Id = Request.QueryString("vEditCustId")
      vFunction = "edit"
    Else
       Response.Redirect "ChannelEdit.asp"          
    End If

    '...get the values (even if trying to add)
    sGetCust (vCust_Id)
    sGetMemb (svMembNo)
%>
  <form method="POST" action="CorporateEdit.asp" target="_self" language="JavaScript">
    <input type="hidden" name="vFunction" value="<%=vFunction%>"><input type="hidden" name="vCust_Id" value="<%=vCust_Id%>"><input type="hidden" name="vCust_AcctId" value="<%=Right(vCust_Id, 4)%>">
    <table border="1" width="100%" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3">
      <tr>
        <td align="center" width="100%" valign="Top" colspan="2"><h1>Quick Corporate Site Setup</h1><h2>This allows you to add or edit your corporate customers.</h2></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="Top">Channel Customer Id :</th>
        <td width="35%" valign="Top"><h1><%=vCust_Id%></h1></td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Title :</th>
        <td valign="Top" width="60%"><input type="text" size="46" name="vCust_Title" value="<%=vCust_Title%>"><br>Appears as the browser title.</td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Language :</th>
        <td valign="Top" width="60%"><input type="radio" name="vCust_Lang" value="EN" <%=fcheck(vcust_lang, "en")%>>EN&nbsp;&nbsp;&nbsp; <input type="radio" name="vCust_Lang" value="FR" <%=fcheck(vcust_lang, "fr")%>>FR&nbsp;&nbsp;&nbsp; <input type="radio" name="vCust_Lang" value="ES" <%=fcheck(vcust_lang, "es")%>>ES<br>If empty, defaults to EN.</td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Logo :</th>
        <td valign="Top" width="60%"><input type="text" size="46" name="vCust_Banner" value="<%=vCust_Banner%>"><br>Channel logo that appears on top left of every page.<br>If empty, uses &quot;vubz.jpg&quot;</td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Logo Url :</th>
        <td valign="Top" width="60%"><input type="text" size="46" name="vCust_URL" value="<%=vCust_URL%>"><br>Channel URL accessed when above logo is clicked.&nbsp; Do not preface with &quot;//&quot;.<br>If empty, defaults to &quot;Vubix.com&quot; </td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Catalogue :</th>
        <td valign="Top" width="60%"><% i = fCatlDropdown (vCust_Catalogue, "NORM") %><select size="7" name="vCust_Catalogue" multiple><% = i %></select><br>Catalogue groups can be free &quot;F&quot;, sold via ecommerce &quot;$&quot; or both.<br>Use CTL+Enter to select multiple groups.</td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Facilitator Password:</th>
        <td valign="Top" width="60%"><input type="text" name="vCust_FacilitatorId" size="46" value="<%=fIf(vCust_FacilitatorId = "", vCust_Id & "_FAC", vCust_FacilitatorId)%>"><br>Give this Password to the channel to review their usage reports.<br>Note: once setup, this should NOT be changed.</td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="30%">Manager Password :</th>
        <td valign="Top" width="60%"><input type="text" name="vCust_ManagerId" size="46" value="<%=fIf(vCust_ManagerId = "", vCust_Id & "_MGR", vCust_ManagerId)%>"><br>Give to client if they use ecommerce.<br>Note: once setup, this should NOT be changed.</td>
      </tr>
      <tr>
        <td colspan="2" align="center" valign="Top" width="100%"><br><br><a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image"> <% 
            If svMembLevel = 5 Or (svMembLevel = 4 AND vMemb_Channel) Then 
              i = "javascript:jconfirm('ChannelEdit.asp?vDelCustId=" & vCust_Id & "&vDelCustAcctId=" & vCust_AcctId & "&vFunction=del', 'Ok to delete this customer and all supporting files?')"
          %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="<%=i%>"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a> <% End If %>
        <h6>Warning, if you delete an existing corporate account, history logs will also be deleted.<br>This is an irreversible, non-recoverable action.</h6>
        <p align="center"><a <%=fStatX%> href="CorporateEdit.asp">Customer List</a><br>&nbsp;</p>
        </td>
      </tr>
    </table>
  </form>
  <% 
    End If
  
  Server.Execute vShellLo
%>

</body>

</html>