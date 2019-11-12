<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<%
  Dim vUrl, vActive, vGlobal, vCustId, vNext, vEdit, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vLearners

  Select Case svMembLevel
    Case 3 : vLearners = "123"
    Case 4 : vLearners = "1234"
    Case 5 : vLearners = "12345"
  End Select

  vNext            = Request("vNext")
  vEdit            = fDefault(Request("vEdit"), "User" & fGroup & ".asp")
  vCustId          = fDefault(Request("vCustId"), svCustId)
  vActive          = fDefault(Request("vActive"), "0")
  vGlobal          = fDefault(Request("vGlobal"), "0")
  vLearners        = Replace(Replace(fDefault(Request("vLearners"), vLearners), " ", ""), ",", "")
  vFind            = fDefault(Request("vFind"), "S")
  vFindId          = fUnQuote(Request("vFindId"))
  vFindFirstName   = fUnQuote(Request("vFindFirstName"))
  vFindLastName    = fUnQuote(Request("vFindLastName"))
  vFindEmail       = fNoQuote(Request("vFindEmail"))
  vFindMemo        = fUnQuote(Request("vFindMemo"))
  vFindCriteria    = Replace(fDefault(Request("vFindCriteria"), "0"), " ", "")
  vFormat          = fDefault(Request("vFormat"), "o")

  If Request.Form.Count > 0 Then
    vUrl = ""
    vUrl = vUrl & "?vNext="          & vNext
    vUrl = vUrl & "&vEdit="          & vEdit
    vUrl = vUrl & "&vCustId="        & vCustId
    vUrl = vUrl & "&vActive="        & vActive
    vUrl = vUrl & "&vGlobal="        & vGlobal
    vUrl = vUrl & "&vFind="          & vFind
    vUrl = vUrl & "&vFindId="        & vFindId
    vUrl = vUrl & "&vFindFirstName=" & vFindFirstName
    vUrl = vUrl & "&vFindLastName="  & vFindLastName
    vUrl = vUrl & "&vFindEmail="     & vFindEmail
    vUrl = vUrl & "&vFindMemo="      & vFindMemo
    vUrl = vUrl & "&vFindCriteria="  & vFindCriteria
    vUrl = vUrl & "&vFormat="        & vFormat
    vUrl = vUrl & "&vLearners="      & vLearners
'   Response.Write "Users_" & vFormat & ".asp" & vUrl
    Response.Redirect "Users_" & vFormat & ".asp" & vUrl
  End If
  
  sGetCust vCustId

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
  <title>Users</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
  <% If vRightClickOff Then %><script type="text/javascript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet" />
  <script type="text/javascript" src="/V5/Inc/jQuery.js"></script>
  <script type="text/javascript" src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1><!--[[-->Learner Report<!--]]--></h1>
  <p class="c3"><!--[[-->This displays the details of Learners registered within this account by Last Name.&nbsp; Using any <b>Find</b> values below will produce a partial report displaying learners that match the selections.&nbsp; Using <b>Finding values that start with</b> values like <b>Hen</b> will return <b>Henry</b>, <b>Hendrickson, </b>etc. while using <b>...that contain </b>values like <b>Hen</b> will return <b>Henry</b>, <b>O&#39;Henry</b>, etc.<!--]]--><br /><br /></p>

  <form  method="post" action="Users.asp">
    <table style="width:500px; margin:auto;" class="table">
      <% If svMembLevel = 5 Or svMembManager Then %>
      <tr>
        <th>Include ALL Customers ?</th>
        <td>
          <input type="radio" name="vGlobal" value="1" <%=fcheck("1", vglobal)%> />Yes&nbsp; (Useful for a NARROW global search) <br /> 
          <input type="radio" name="vGlobal" value="0" <%=fcheck("0", vglobal)%> />No
        </td>
      </tr>
      <% Else %> 
      <input type="hidden" name="vGlobal" value="0" />
      <% End If %>

      <tr>
        <th style="width:50%"><!--[[-->Include<!--]]--> :</th>
        <td style="width:50%">

          <% If vCust_ChannelGuests Then %>
          <input type="checkbox" name="vLearners" value="1" <%=fchecks(vLearners, "1")%> /><!--[[-->Guests<!--]]--><br />
          <% End If %>
          <input type="checkbox" name="vLearners" value="2" <%=fchecks(vLearners, "2")%> /><!--[[-->Learners<!--]]--><br />
          <input type="checkbox" name="vLearners" value="3" <%=fchecks(vLearners, "3")%> /><!--[[-->Facilitators<!--]]-->
          <% If svMembLevel > 3 Then %><br />
          <input type="checkbox" name="vLearners" value="4" <%=fchecks(vLearners, "4")%> />Managers<% End If %>
          <% If svMembLevel = 5 Then %><br />
          <input type="checkbox" name="vLearners" value="5" <%=fchecks(vLearners, "5")%> />Administrators<% End If %>
          <% If vCust_MaxSponsor > 0 Then %><br />
          <input type="checkbox" name="vLearners" value="s" <%=fchecks(vLearners, "s")%> /><!--[[-->Sponsored Learners<!--]]-->
          <% End If %>

        </td>
      </tr>

      <tr>
        <th><!--[[-->and Inactive Learners<!--]]--> ?</th>
        <td>
          <input type="radio" value="1" <%=fcheck("1", vactive)%> name="vActive" /><%=fYN (1)%><br /> 
          <input type="radio" value="0" <%=fcheck("0", vactive)%> name="vActive" /><%=fYN (0)%>
        </td>
      </tr>
      <tr>
        <th><!--[[-->Finding values that<!--]]--> : </th>
        <td><input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%> /><!--[[-->start with<!--]]--><br /> 
          <input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%> /><!--[[-->contain<!--]]-->
        </td>
      </tr>
      <tr>
        <th><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> : </th>
        <td><input type="text" name="vFindId" size="29" value="<%=vFindId%>" /></td>
      </tr>
      <tr>
        <th><!--[[-->First Name<!--]]--> : </th>
        <td><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>" />&nbsp; </td>
      </tr>
      <tr>
        <th><!--[[-->Last Name<!--]]--> :</th>
        <td><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>" /></td>
      </tr>
      <tr>
        <th><!--[[-->Email Address<!--]]--> :</th>
        <td><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>" /></td>
      </tr>
      <tr>
        <th>Memo :</th>
        <td><input type="text" name="vFindMemo" size="29" value="<%=vFindMemo%>" /></td>
      </tr>      
      <% 
        If svMembLevel > 3 Then svMembCriteria = 0
        i = fCriteriaList (vCust_AcctId, "REPT:" & svMembCriteria)
        If vCriteriaListCnt > 1 Then
      %>

      <tr>
        <th><!--[[-->from Group<!--]]--> :</th>
        <td><select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select></td>
      </tr>
      <%  
          Else 
      %>
      <input type="hidden" name="vFindCriteria" value="<%=svMembCriteria%>" />
      <tr>
        <th><!--[[-->from Group<!--]]--> :</th>
        <td><%=fCriteria (svMembCriteria)%></td>
      </tr>
      <% 
        End If 
      %>
      <tr>
        <th><!--[[-->Format<!--]]--> : </th>
        <td>
          <input type="radio" name="vFormat" value="o" <%=fCheck("o", vFormat)%> /><!--[[-->Online<!--]]--><br />
          <input type="radio" name="vFormat" value="x" <%=fCheck("x", vFormat)%> />MS Excel</td>
      </tr>
      <tr>
        <td colspan="2" style="text-align:center">
          <table style="margin:30px; padding:30px;">
            <% If Len(vNext) > 0 Then %>
            <tr>
              <td style="text-align:center"><input type="button" onclick="location.href='<%=vNext%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button085" /></td>
              <td style="text-align:center"><input type="submit" value="<%=bContinue%>" name="bContinue" class="button085" /></td>
            </tr>
            <% Else %>
            <tr>
              <td style="text-align:center" colspan="2"><input type="submit" value="<%=bContinue%>" name="bContinue" class="button085" /></td>
            </tr>
            <% End If %>
            <% If vCust_Id = svCustId Then %>
            <tr>
              <td style="text-align:left"><% If (vCust_InsertLearners) Then %><a <%=fstatx%> href="<%=vEdit%>?vMembNo=0"><!--[[-->Add a Learner<!--]]--></a><% End If %></td>
              <td style="text-align:right"><a <%=fstatx%> href="<%=vEdit%>"><!--[[-->My Profile<!--]]--></a></td>
            </tr>
            <% End If %>

          </table>
          <h3><%=vCust_Id & "  (" & vCust_Title & ")"%></h3>
        </td>
      </tr>
    </table>

    <input type="hidden" name="vNext"   value="<%=vNext%>" />
    <input type="hidden" name="vEdit"   value="<%=vEdit%>" />
    <input type="hidden" name="vCustId" value="<%=vCustId%>" />


  </form>
  

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>