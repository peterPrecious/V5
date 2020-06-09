
<script>

//  this page is no longer used, see info.asp

  <% 
    Dim vAlert
'   Session("Breach") = True
    '...if this user is flagged as online
    If svMembLevel < 5 And Session("Breach") Then 
      Session("Breach") = False
      vAlert = "<!--{{-->Warning! Your status shows that you are already online.\nThis can happen if you did not 'Sign Off' after your last session,\nare signed in with more than one browser window or\nsomeone else is accessing your account!\nThis can, at a minimum, cause loss of data integrity.\n\n[Breach Status ID:<!--}}-->" & svCustAcctId & "-" & Year(Now) & Month(Now) & Day(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & "]."  
  %>
      alert("<%=vAlert%>");
  <% 
    End If

    '...if used on staging
    If (Not svMembInternal) And svServer = "stagingweb.vubiz.com" And Len(Session("Staging")) = 0 Then             
      Session("Staging") = True
      vAlert = "<!--{{-->Warning! You are on the Vubiz STAGING server which is \nused for functional review NOT real time learning.\nNo records are maintained!\n\n If you intented to use our real time service\n please visit //vubiz.com.<!--}}-->"  
  %>
      alert("<%=vAlert%>");
  <% 
    End If
  %>
    
  // render popup status (twig to show "Yes" for enabled rather than for Blocker On)
  $(function() { 
    $("#popupStatus").html(parent.popupBlockerOn ? jYN("n", "<%=svLang%>") : jYN("y", "<%=svLang%>")); 
  });
</script>



<% 

  Dim vIntro, vParagraph
  vIntro = "<!--{{-->Welcome<!--}}-->" 
  If svSecure and Len(svMembFirstName) > 1 Then 
    vIntro = vIntro & " " & svMembFirstName
  End If 

  '...determine what tabs are available and sponsored learners
  sGetCust svCustId

  '...insert these into tab names
  p1 = fTabName(vCust_Tab1Name, "<!--{{-->Info Page<!--}}-->")
  p2 = fTabName(vCust_Tab2Name, "<!--{{-->My Learning<!--}}-->")
  p3 = fTabName(vCust_Tab3Name, "<!--{{-->My Content<!--}}-->")
  p4 = fTabName(vCust_Tab4Name, "<!--{{-->My Development<!--}}-->")
  p5 = fTabName(vCust_Tab5Name, "<!--{{-->More Content<!--}}-->")
  p6 = fTabName(vCust_Tab6Name, "<!--{{-->Administration<!--}}-->")
  p7 = fTabName(vCust_Tab7Name, "<!--{{-->Sign Off<!--}}-->")

  p8 = "<!--{{-->Id<!--}}-->"
  p9 = "<!--{{-->Password<!--}}-->"
  

  Function fTabName(vTabName, vOriginal)
    Dim aTabName
    If IsNull(vTabName) Or Len(vTabName) = 0 Then  
      fTabName = vOriginal
    Else
      aTabName = Split(vTabName, "|")
      If svLang = "EN" And Ubound(aTabName) >= 0 Then 
        If Len(aTabName(0)) > 0 Then fTabName = Trim(Left(aTabName(0) & Space(20), 20))
      End If
      If svLang = "FR" And Ubound(aTabName) >= 1 Then
        If Len(aTabName(1)) > 0 Then fTabName = Trim(Left(aTabName(1) & Space(20), 20))
      End If
      If svLang = "ES" And Ubound(aTabName) >= 2 Then
        If Len(aTabName(2)) > 0 Then fTabName = Trim(Left(aTabName(2) & Space(20), 20))
      End If
    End If
  End Function

    '...determine if sponsored learner
  sGetMemb svMembNo

  '...get to see if pc or mac (for bookmarking)
  sGetQueryString
%>

<div>
  <%
    '...Show|Hide the alert at admin's request via url below: vAlert=y | vAlert=n
    If Application("Alert") = "y" Then
  %>
  <h2 class="red">::&ensp;<!--[[-->NOTICEA!<!--]]--></h2>

  <% If svLang = "FR" Then %>
       Ce service fera l'objet de l'entretien courant et les améliorations des applications et ne sera pas disponible samedi 13 Sep de 08 jusqu'à 10 h HNE. Nous nous excusons pour tout inconvénient.
    <% ElseIf svLang = "ES" Then %>        
       Este servicio va a ser sometido a mantenimiento de rutina y mejoras a las aplicaciones y no estará disponible el sábado 13 de Sep de 8 hasta 10 am EST. Nos disculpamos por cualquier inconveniente.
    <% Else %>
       This service will be undergoing routine maintenance and application enhancements and will not be available on Saturday Sep 13th from 8am until 10am EST. We apologize for any inconvenience.
    <% End If %>
  <%
     End If  
  %>
  
  <h2 class="green">::&nbsp;&nbsp;<%=Trim(vIntro)%></h2>

  <% If vCust_Tab2 Then %>
  <!--[[-->Click on the <b>^2</b> tab above to access your programs.<!--]]-->&nbsp;
  <% End If %>

  <% If vCust_Tab3 Then %>
  <!--[[-->Click on the <b>^3</b> tab above to access your free or purchased programs.<!--]]-->&nbsp;
  <% End If %>

  <% If vCust_Tab5 Then %>
  <!--[[-->To purchase e-learning programs, click <b>^5</b> to complete a secure e-commerce process.&nbsp;&nbsp; Any programs purchased will then appear under the <b>^3</b> tab above.<!--]]-->&nbsp;
  <% End If %>

  <% 
  	If svLang = "EN" Then
      '...intro paragraph 
      Select Case svCustCluster
        Case "C0002" : vParagraph = "<br /><br />The organizations represented here share a common goal of helping their constituents embrace e-business technologies. They are further committed to maximizing your opportunity to share in their respective benefits.<br /><br /> Click the <a " & fStatX & " href='javascript:ebizwindow()'>eLearning For Business</a> logo below and discover how to move your business into the online world step by step!"
        Case  Else   : vParagraph = ""
      End Select
      If Len(vParagraph) > 0 Then Response.Write vParagraph
    End If 
  %>

  <!--[[-->If you have any questions or comments please contact us using the link at the bottom of the page.<!--]]-->
  <h2 class="green">::&nbsp;&nbsp;<!--[[-->Important<!--]]--></h2>
  <!--[[-->Please do NOT <b>Bookmark</b> or <b>Add to Favorites</b> the address that appears in your web browser when you are logged in. You must login with your user credentials each time you enter this service. Please click the Sign Off tab at the end of each visit.<!--]]-->


  <% If vCust_MaxSponsor > 0 And vMemb_Sponsor = 0 Then '...if accounts allow sponsors then ensure this link is not for a sponsored learner %>
  <h2 class="green">::&nbsp;&nbsp;<!--[[-->My Sponsored Learners<!--]]--></h2>
  <h2><!--[[-->If you would like to offer members of your organization access to your content, click ...<!--]]--><a href="Sponsors.asp"><u><!--[[-->Sponsored Learners.<!--]]--></u></a></h2>
  <% End If %>

  <% If vCust_Scheduler Then %>
  <h2 class="green">::&nbsp;&nbsp;<!--[[-->My Scheduler<!--]]--></h2>
  <h2><a href="Scheduler.asp"><u><!--[[-->Click here<!--]]--></u></a><!--[[-->to view your calendar.<!--]]--></h2>
  <% End If %>


  <% If vCust_InfoEditProfile Then %>

    <h2 class="green"><a <%=fStatX%> name="MyProfile"></a>::&nbsp; <!--[[-->My Profile<!--]]--></h2>
    <!--[[-->Enter/edit your name and email address below. Your name will then appear on any certificates issued for successful completion of assessments or exams.<!--]]-->

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
        <table style="width: 50%; margin: 20px auto auto auto;">
          <tr>
            <th>
              <!--[[-->First Name<!--]]-->
              :</th>
            <td>
              <% If Request.QueryString("vAction") = "edit" Then %>
              <input type="text" name="vMemb_FirstName" size="19" value="<%=svMembFirstName%>" maxlength="32">
              <% Else %>
              <%=svMembFirstName%>
              <% End If %> 
            </td>
          </tr>
          <tr>
            <th>
              <!--[[-->Last Name<!--]]-->
              :</th>
            <td>
              <% If Request.QueryString("vAction") = "edit" Then %><input type="text" name="vMemb_LastName" size="19" value="<%=svMembLastName%>" maxlength="64">
              <% Else %>
              <%=svMembLastName%>
              <% End If %> 
            </td>
          </tr>

          <% If vCust_Pwd And (svMembLevel = 2 Or svCustId = "CFIB1660") Then %>
          <tr>
            <th>
              <!--[[-->Password<!--]]-->
              :</th>
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
            <th>
              <!--[[-->Email Address<!--]]-->
              :
            </th>
            <td>
              <% If Request.QueryString("vAction") = "edit" Then %>
              <input type="text" name="vMemb_Email" size="19" value="<%=svMembEmail%>">
              <% Else %>
              <%=fDefault(svMembEmail, "...<i><font color='#FF0000'>[none]")%>
              <% End If %> 
            </td>
          </tr>

          <% If svLang = "EN" And vCust_VuNews Then %>
          <tr>
            <th>Send vuNews <b><a href="javascript:toggle('Div_VuNews');">?</a></b></th>
            <td>
              <p>
                <% If Request.QueryString("vAction") = "edit" Then %>
                <input type="radio" value="1" name="vMemb_VuNews" <%=fcheck(fsqlboolean(vMemb_VuNews), 1)%>>Yes&nbsp; 
                      <input type="radio" value="0" name="vMemb_VuNews" <%=fcheck(fsqlboolean(vMemb_VuNews), 0)%>>No&nbsp;
                    <% Else %>
                <%=fIf(vMemb_VuNews, "Yes", "No")%>
                <% End If %>
                <%=f5%>
            </td>
          </tr>
          <tr>
            <th nowrap colspan="2">
              <div align="center" id="Div_VuNews" class="div">
                <table border="0" id="table3" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9" bgcolor="#FFFFFF">
                  <tr>
                    <td>vuNews is an online newsletter that we publish quarterly.&nbsp; If interested, click Edit, select Yes to Send vuNews and your email address will be added to our distribution list.&nbsp; You can discontinue the newsletter at any time.<h6>Be assured, your profile will NEVER be released to any third parties.</h6>
                      <p align="center">
                        Thank you!</td>
                  </tr>
                </table>
              </div>
            </th>
          </tr>
          <% End If %>

          <tr>
            <th colspan="2" nowrap height="30">
              <% If Request.QueryString("vAction") = "edit" Then %>
              <input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button">
              <% Else %>
              <a <%=fStatX%> href="AnchorFix.asp?vNext=<%=svCustCluster%>.asp&vAction=edit&vAnchor=MyProfile"></a>
              &nbsp;<input onclick="location.href = 'AnchorFix.asp?vNext=<%=svCustCluster%>.asp&vAction=edit&vAnchor=MyProfile'" type="button" value="<%=bEdit%>" name="bEdit" class="button">
              <% End If %>
            </th>
          </tr>

        </table>
    </form>

  <% End If %>

  <h2 class="green">::&nbsp;&nbsp;<!--[[-->My Status<!--]]--></h2>

  <% 
    If vCust_Level = 4 OR svMembLevel = 5 Then 
  %>
      <a href="LearnerReportCard2.asp?vMemb_No=<%=svMembNo%>&vInfoPage=y"><!--[[-->Click here<!--]]--></a>&nbsp;<!--[[-->for your Report Card<!--]]-->.
  <% 
    End If 
  %>


  <% 
    If vCust_Level = 2 OR svMembLevel = 5 Then 
  %>
  <br><a href="RTE_History_F.asp?vPass=<%=svMembId%>&vFrom=<%=svPage%>"><!--[[-->Click here<!--]]--></a>&nbsp;<!--[[-->for your Report Card<!--]]-->.
  <% 
    End If 
  %>


  <%
    '...data is passed in as svBrowser and determined in the initial default.asp
    Dim aTools, vTouch, vHTML5, vFlash, vCookies


    svBrowser = svBrowser & "     "
    aTools = Split(Ucase(svBrowser), "|")

    vTouch   = fYN (aTools(0))
    vBrowser = aTools(1)
    vHTML5   = fYN (aTools(2)) 
    vFlash   = fIf(aTools(3) = "0", fYN(0), aTools(3)) 
    vCookies = fYN (aTools(4)) 
  %>

  <p />
  <table style="width: 50%; margin: 20px auto auto auto;">

    <tr>
      <th><!--[[-->Customer Id<!--]]-->:</th>
      <td><%=svCustId %></td>
    </tr>
    <tr>
      <th><% If svCustPwd Then %><!--[[-->Id<!--]]--><% Else %><!--[[-->Password<!--]]--><% End If %> :</th>
      <td><%=fIf(svMembInternal, "**********", svMembId)%></td>
    </tr>

    <tr>
      <th align="right" nowrap width="50%">&nbsp;</th>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--[[-->Popups Enabled<!--]]-->
        :</th>
      <td id="popupStatus" width="50%"></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--[[-->Touch Screen<!--]]-->
        :</th>
      <td id="touchStatus" width="50%"><%=vTouch%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--[[-->Browser<!--]]-->
        :</th>
      <td><%=vBrowser%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">HTML5 :</th>
      <td><%=vHTML5%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">Flash :</th>
      <td><%=vFlash%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--[[-->Cookies Enabled<!--]]-->
        :</th>
      <td><%=vCookies%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">&nbsp;</th>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--[[-->First Visit<!--]]-->
        :</th>
      <td><%=fFormatDate(svMembFirstVisit)%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--[[-->Last Visit<!--]]-->
        :</th>
      <td><%=fFormatDate(svMembLastVisit)%></td>
    </tr>
    <% If fIsGroup2 Then %>
    <tr>
      <th align="right" nowrap width="50%"><!--[[-->Account Expires<!--]]-->:</th>
      <td><%=fFormatDate(vCust_Expires)%></td>
    </tr>
    <% 
							Else		
				        If IsDate(svMembExpires) Then 
				          If svMembExpires > Now Then 
    %>
    <tr>
      <th align="right" nowrap width="50%"><!--[[-->Programs Expires<!--]]-->:</th>
      <td><%=fFormatDate(svMembExpires)%></td>
    </tr>
    <%
				          End If
				        End If
						  End If 
    %>
  </table>

  <% If svLang = "EN" Then %>
  <h2 class="green">::&nbsp; Help using this service</h2>
  <a href="../Public/21_FAQ.asp?vReturn=y">Click here</a>&nbsp;if you have any questions about how this service works.
  <% End If %>

  <% If svLang = "FR" Then %>
  <h2 class="green">::&nbsp; Problèmes liés aux navigateurs?</h2>
  <p><a href="../Public/BrowserIssues_FR.htm?vReturn=Y">&nbsp;Cliquez ici</a> pour options de réglage de votre navigateur Web</p>
  <% End If %>

  <br><br>
</div>
