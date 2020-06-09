







<head>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <style>
    <!--
    div.Section1
    	{page:Section1;}
    -->
  </style>
</head>

<% 
  Dim vIntro, vParagraph
  vIntro = "<!--{{-->Welcome<!--}}-->" 
  If svSecure and Len(svMembFirstName) > 1 Then 
    vIntro = vIntro & " " & svMembFirstName
  End If 

  '...determine what tabs are available and sponsored learners
  sGetCust svCustId

  '...determine if sponsored learner
  sGetMemb svMembNo


  '...get to see if pc or mac (for bookmarking)
  sGetQueryString
%> 


<p align="left"><br><font face="Arial Black" size="2" color="#3977B6"><font color="#FF0000">::</font>&nbsp; <%=Trim(vIntro)%></font></p><p> 

<div class="c2" align="left">

  <% If vCust_Tab2 Then %><!--[[-->Click on the <b>My Learning</b> tab above to access your programs.<!--]]-->&nbsp;<% End If %>
  <% If vCust_Tab3 Then %><!--[[-->Click on the <b>My Content</b> tab above to access your free or purchased programs.<!--]]-->&nbsp;<% End If %>
  <% If vCust_Tab5 Then %><!--[[-->To purchase e-learning programs, click <b>More Content</b> to complete a secure e-commerce process.&nbsp;&nbsp; Any programs purchased will then appear under the <b>My Content</b> tab above.<!--]]-->&nbsp;<% End If %>
  
  <% 
  	If svLang = "EN" Then
      '...intro paragraph 
      Select Case svCustCluster
        Case "C0001" : vParagraph = ""
        Case "C0002" : vParagraph = "The organizations represented here share a common goal of helping their constituents embrace e-business technologies. They are further committed to maximizing your opportunity to share in their respective benefits. Click the <a " & fStatX & " href='javascript:ebizwindow()'><font color='#3977B6'>eLearning For Business</font></a> logo below and discover how to move your business into the online world step by step!"
        Case "C0003" : vParagraph = ""
        Case "C0004" : vParagraph = "The Halifax Inner City Initiative is an initiative of the North End Council of Churches. The mission of the Halifax Inner City Initiative is to support the community in building a healthy, safe environment in which the citizens can become fully employed, using practical and intelligent practices."
        Case  Else   : vParagraph = ""
      End Select
      If Len(vParagraph) > 0 Then
        Response.Write "<br><font face='Verdana' size='1' color='#3977B6'>" & vParagraph & "</font>"
      End If
    End If 
  %>&nbsp;<!--[[-->If you have any questions or comments please contact us using the link at the bottom of the page.<!--]]-->
  
  <% If vCust_MaxSponsor > 0 And vMemb_Sponsor = 0 Then '...if accounts allow sponsors then ensure this link is not for a sponsored learner %>
  <br><br><font face="Arial Black" size="2" color="#FF0000">::</font><font face="Arial Black" size="2" color="#3977b6">&nbsp; My Sponsored Learners</font><br><br>If you would like to offer members of your organization access to your content, click ... <a class="c2" href="Sponsors.asp"><u>Sponsored Learners.</u></a>
  <% End If %> 

</div>

<% If vCust_InfoEditProfile Then %> 
  <p align="left"><font color="#3977B6"><a <%=fStatX%> name="MyProfile"></a></font><font face="Arial Black" size="2"><font color="#FF0000">::</font><font color="#3977B6">&nbsp;
  <!--[[-->My profile<!--]]--></font></font></p><h2 align="left">
  <!--[[-->Enter/edit your name and email address below. Your name will then appear on any certificates issued for successful completion of assessments or exams.<!--]]--></h2>
  


  <script>
    function Validate(theForm) 
    {

      //  only check password if used by this memeber, else ignore
      if (theForm.vPassword.value == "check") 
      {

        if (theForm.vMemb_Pwd.value == "")
        {
          alert("Please enter a value for the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
      
        if (theForm.vMemb_Pwd.value.length < 4)
        {
          alert("Please enter at least 4 characters in the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
      
        if (theForm.vMemb_Pwd.value.length > 64)
        {
          alert("Please enter at most 64 characters in the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
      
        var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-_-@.";
        var checkStr = theForm.vMemb_Pwd.value;
        var allValid = true;
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
          alert("Please enter only letter, digit and \"_-@.\" characters in the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
        
      }
      return (true);
    }
  </script>


  <form method="POST" action="<%=svCustCluster%>.asp" onsubmit="return Validate(this)" name="fHome">
    <input type="hidden" name="fProfile" value="Y">
    <div align="center">
    <center>&nbsp;
    <table border="1" id="table1" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="10" bgcolor="#F2F9FD">
      <tr>
        <td>
          <table border="0" style="border-collapse: collapse" id="table2" cellpadding="2">
            <tr>
              <th align="right" nowrap><!--[[-->First Name<!--]]--> :</th>
              <td>
                <% If Request.QueryString("vAction") = "edit" Then %>
                  <input type="text" name="vMemb_FirstName" size="19" value="<%=svMembFirstName%>" maxlength="32"> 
                <% Else %> 
                  <%=svMembFirstName%> 
                <% End If %> 
              </td>
            </tr>
            <tr>
              <th align="right" nowrap><!--[[-->Last Name<!--]]--> :</th>
              <td>
                <% If Request.QueryString("vAction") = "edit" Then %>
                  <input type="text" name="vMemb_LastName" size="19" value="<%=svMembLastName%>" maxlength="64"> 
                <% Else %> 
                  <%=svMembLastName%> 
                <% End If %> 
              </td>
            </tr>
      
            <% If vCust_Pwd And svMembLevel = 2 Then %>
            <tr>
              <th align="right" nowrap><!--[[-->Password<!--]]--> :</th>
              <td>
                <% If Request.QueryString("vAction") = "edit" Then %>
                  <input type="password" name="vMemb_Pwd" size="19" value="<%=svMembPwd%>" maxlength="64"> 
                <% Else %> 
                  <%="****************"%> 
                <% End If %> 
              </td>
            </tr>
            <input type="hidden" name="vPassword" value="check">
            <% Else %>
            <input type="hidden" name="vPassword" value="ignore">
            <% End If %>
      
            <tr>
              <th align="right" nowrap><!--[[-->Email Address<!--]]--> :</th>
              <td>
                <% If Request.QueryString("vAction") = "edit" Then %>
      	          <input type="text" name="vMemb_Email" size="19" value="<%=svMembEmail%>"> 
                <% Else %> 
      	          <%=svMembEmail%> 
                <% End If %> 
              </td>
            </tr>
            <tr>
              <th colspan="2" align="right" nowrap>
              <br>
              <% If Request.QueryString("vAction") = "edit" Then %>
                <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I1" type="image"> 
              <% Else %> 
                <a <%=fStatX%> href="AnchorFix.asp?vNext=<%=svCustCluster%>.asp&vAction=edit&vAnchor=MyProfile"><img border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif"></a> 
              <% End If %> 
            </th>
            </tr>
            <tr>
              <th colspan="2" align="right" nowrap>&nbsp; </th>
            </tr>
            <tr>
              <th align="right" nowrap><!--[[-->First Visit<!--]]--> :</th>
              <td><%=fFormatDate(svMembFirstVisit)%></td>
            </tr>
      
            <tr>
              <th align="right" nowrap><!--[[-->Last Visit<!--]]--> :</th>
              <td><%=fFormatDate(svMembLastVisit)%></td>
            </tr>
            <%  
                If IsDate(svMembExpires) Then 
                  If svMembExpires > Now Then 
            %>
            <tr>
              <th align="right" nowrap><!--[[-->Access Expires<!--]]--> :</th>
              <td><%=fFormatDate(svMembExpires)%></td>
            </tr>
            <%
                  End If
                End If
            %>
      
      
          </table>
        </td>
      </tr>
    </table>
    </center></div>
  </form>

  <p>&nbsp;</p>

<% End If %> 


<!--- temporarily killed since the anchor issue seems to not allow this -->
<% If vBrowser = "msie" And svLang = "EN" And (svMembLevel > 2 Or (Len(vSource) = 0 And Not vCust_Auto)) Then %>

  <!--Bookmark to My Favorites-->
  <p align="left"><font face="Arial Black" size="2" color="#FF0000">::</font><font face="Arial Black" size="2" color="#3977b6">&nbsp; Bookmark this service</font></p><h2 align="left">If you are currently on your own computer you can bookmark this service for speedy access.&nbsp; However, please note that your Customer Id and Password will be stored in your Favorites List which may be a security breach</h2>
  
  <p class="c2">
  <script>
    var vUrl="//<%=svHost%>/default.asp?vCust=<%=svCustId%>&vId=<%=svMembId%>";
    var vTitle="<%=Trim(fNoQuote(svCustTitle))%>";
    document.write('To add to Favorites, click... <a HREF="javascript:window.external.AddFavorite(vUrl,vTitle);" ');
    document.write('onMouseOver=" window.status=');
    document.write("'Add ' + vTitle + ' to your favorites!'; return true ");
    document.write('"onMouseOut=" window.status=');
    document.write("' '; return true ");
    document.write('">' + vTitle + '</a>');
  </script>
  </p>
<% End If %>

<p>

<% If svLang = "EN" And svCustCluster <> "C2668" Then %>
  <p align="left"><font face="Arial Black" size="2" color="#FF0000">::</font><font face="Arial Black" size="2" color="#3977b6">&nbsp; Help using this service&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font> </p><h2 align="left"><a href="../Public/21_FAQ.asp?vReturn=y">Click here</a> if you have any questions about how this service works.</h2><font face="Verdana" size="1" color="#3977B6">
<% End If %>