<%
   sGetCust svCustId
   If vCust_InfoEditProfile Then
%>

<p align="left"><font color="#3977B6"><a name="MyProfile"></a><font face="Arial Black" size="2">::&nbsp;
<!--webbot bot='PurpleText' PREVIEW='My profile'--><%=fPhra(000185)%></font></font></p>
<p align="left"><font face="Verdana" size="1" color="#3977B6">
<!--webbot bot='PurpleText' PREVIEW='Enter/edit your name and email address below. Your name will then appear on any certificates issued for successful completion of assessments or exams.'--><%=fPhra(000129)%></font></p>
<form method="POST" action="<%=svCustCluster%>.asp">
  <input type="hidden" name="fProfile" value="Y">
  <div align="center">
    <center>
    <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0">
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='First Visit'--><%=fPhra(000157)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><%=fFormatDate(svMembFirstVisit)%></font></td>
      </tr>
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='Last Visit'--><%=fPhra(000164)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><%=fFormatDate(svMembLastVisit)%></font></td>
      </tr>
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='Number of Visits'--><%=fPhra(000198)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><%=svMembNoVisits%></font></td>
      </tr>
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='Hours Online'--><%=fPhra(000144)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><%=fFormatDecimals(FormatNumber(svMembNoHours,1))%></font></td>
      </tr>
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><% If Request.QueryString("vAction") = "edit" Then %> <input type="text" name="vMemb_FirstName" size="19" value="<%=svMembFirstName%>"> <% Else %> <%=svMembFirstName%> <% End If %> </font></td>
      </tr>
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><% If Request.QueryString("vAction") = "edit" Then %> <input type="text" name="vMemb_LastName" size="19" value="<%=svMembLastName%>"> <% Else %> <%=svMembLastName%> <% End If %> </font></td>
      </tr>
      <tr>
        <td align="right"><b><font face="Verdana" size="1" color="#3977B6">
        <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> :&nbsp; </font></b></td>
        <td><font face="Verdana" size="1"><% If Request.QueryString("vAction") = "edit" Then %> <input type="text" name="vMemb_Email" size="19" value="<%=svMembEmail%>"> <% Else %> <%=svMembEmail%> <% End If %> </font></td>
      </tr>
      <tr>
        <td colspan="2" align="right"><% If Request.QueryString("vAction") = "edit" Then %> <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="iUpdate" type="image"> <% Else %> <a href="<%=svCustCluster%>.asp?vAction=edit#MyProfile"><img border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif"></a> <% End If %> </td>
      </tr>
    </table>
    </center>
  </div>
</form>

<% End If %>



