
<script>

//  this page is no longer used, see info.asp

  <% 
    Dim vAlert
'   Session("Breach") = True
    '...if this user is flagged as online
    If svMembLevel < 5 And Session("Breach") Then 
      Session("Breach") = False
      vAlert = fPhraH(001690) & svCustAcctId & "-" & Year(Now) & Month(Now) & Day(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & "]."  
  %>
      alert("<%=vAlert%>");
  <% 
    End If

    '...if used on staging
    If (Not svMembInternal) And svServer = "stagingweb.vubiz.com" And Len(Session("Staging")) = 0 Then             
      Session("Staging") = True
      vAlert = fPhraH(001779)  
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
  vIntro = fPhraH(000011) 
  If svSecure and Len(svMembFirstName) > 1 Then 
    vIntro = vIntro & " " & svMembFirstName
  End If 

  '...determine what tabs are available and sponsored learners
  sGetCust svCustId

  '...insert these into tab names
  p1 = fTabName(vCust_Tab1Name, fPhraH(000160))
  p2 = fTabName(vCust_Tab2Name, fPhraH(000183))
  p3 = fTabName(vCust_Tab3Name, fPhraH(000182))
  p4 = fTabName(vCust_Tab4Name, fPhraH(000950))
  p5 = fTabName(vCust_Tab5Name, fPhraH(000180))
  p6 = fTabName(vCust_Tab6Name, fPhraH(000065))
  p7 = fTabName(vCust_Tab7Name, fPhraH(000240))

  p8 = fPhraH(000374)
  p9 = fPhraH(000211)
  

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
  <h2 class="red">::&ensp;<!--webbot bot='PurpleText' PREVIEW='NOTICEA!'--><%=fPhra(001654)%></h2>

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
  <!--webbot bot='PurpleText' PREVIEW='Click on the <b>^2</b> tab above to access your programs.'--><%=fPhra(001508)%>&nbsp;
  <% End If %>

  <% If vCust_Tab3 Then %>
  <!--webbot bot='PurpleText' PREVIEW='Click on the <b>^3</b> tab above to access your free or purchased programs.'--><%=fPhra(001509)%>&nbsp;
  <% End If %>

  <% If vCust_Tab5 Then %>
  <!--webbot bot='PurpleText' PREVIEW='To purchase e-learning programs, click <b>^5</b> to complete a secure e-commerce process.&nbsp;&nbsp; Any programs purchased will then appear under the <b>^3</b> tab above.'--><%=fPhra(001510)%>&nbsp;
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

  <!--webbot bot='PurpleText' PREVIEW='If you have any questions or comments please contact us using the link at the bottom of the page.'--><%=fPhra(000150)%>
  <h2 class="green">::&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='Important'--><%=fPhra(000153)%></h2>
  <!--webbot bot='PurpleText' PREVIEW='Please do NOT <b>Bookmark</b> or <b>Add to Favorites</b> the address that appears in your web browser when you are logged in. You must login with your user credentials each time you enter this service. Please click the Sign Off tab at the end of each visit.'--><%=fPhra(001318)%>


  <% If vCust_MaxSponsor > 0 And vMemb_Sponsor = 0 Then '...if accounts allow sponsors then ensure this link is not for a sponsored learner %>
  <h2 class="green">::&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='My Sponsored Learners'--><%=fPhra(000514)%></h2>
  <h2><!--webbot bot='PurpleText' PREVIEW='If you would like to offer members of your organization access to your content, click ...'--><%=fPhra(000822)%><a href="Sponsors.asp"><u><!--webbot bot='PurpleText' PREVIEW='Sponsored Learners.'--><%=fPhra(000515)%></u></a></h2>
  <% End If %>

  <% If vCust_Scheduler Then %>
  <h2 class="green">::&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='My Scheduler'--><%=fPhra(001252)%></h2>
  <h2><a href="Scheduler.asp"><u><!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></u></a><!--webbot bot='PurpleText' PREVIEW='to view your calendar.'--><%=fPhra(001254)%></h2>
  <% End If %>


  <% If vCust_InfoEditProfile Then %>

    <h2 class="green"><a <%=fStatX%> name="MyProfile"></a>::&nbsp; <!--webbot bot='PurpleText' PREVIEW='My Profile'--><%=fPhra(000185)%></h2>
    <!--webbot bot='PurpleText' PREVIEW='Enter/edit your name and email address below. Your name will then appear on any certificates issued for successful completion of assessments or exams.'--><%=fPhra(000129)%>

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
              <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%>
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
              <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%>
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
              <!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%>
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
              <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%>
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

  <h2 class="green">::&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='My Status'--><%=fPhra(001362)%></h2>

  <% 
    If vCust_Level = 4 OR svMembLevel = 5 Then 
  %>
      <a href="LearnerReportCard2.asp?vMemb_No=<%=svMembNo%>&vInfoPage=y"><!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a>&nbsp;<!--webbot bot='PurpleText' PREVIEW='for your Report Card'--><%=fPhra(001511)%>.
  <% 
    End If 
  %>


  <% 
    If vCust_Level = 2 OR svMembLevel = 5 Then 
  %>
  <br><a href="RTE_History_F.asp?vPass=<%=svMembId%>&vFrom=<%=svPage%>"><!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a>&nbsp;<!--webbot bot='PurpleText' PREVIEW='for your Report Card'--><%=fPhra(001511)%>.
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
      <th><!--webbot bot='PurpleText' PREVIEW='Customer Id'--><%=fPhra(000111)%>:</th>
      <td><%=svCustId %></td>
    </tr>
    <tr>
      <th><% If svCustPwd Then %><!--webbot bot='PurpleText' PREVIEW='Id'--><%=fPhra(000374)%><% Else %><!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%><% End If %> :</th>
      <td><%=fIf(svMembInternal, "**********", svMembId)%></td>
    </tr>

    <tr>
      <th align="right" nowrap width="50%">&nbsp;</th>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--webbot bot='PurpleText' PREVIEW='Popups Enabled'--><%=fPhra(001556)%>
        :</th>
      <td id="popupStatus" width="50%"></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--webbot bot='PurpleText' PREVIEW='Touch Screen'--><%=fPhra(001436)%>
        :</th>
      <td id="touchStatus" width="50%"><%=vTouch%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--webbot bot='PurpleText' PREVIEW='Browser'--><%=fPhra(001363)%>
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
        <!--webbot bot='PurpleText' PREVIEW='Cookies Enabled'--><%=fPhra(001557)%>
        :</th>
      <td><%=vCookies%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">&nbsp;</th>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--webbot bot='PurpleText' PREVIEW='First Visit'--><%=fPhra(000157)%>
        :</th>
      <td><%=fFormatDate(svMembFirstVisit)%></td>
    </tr>
    <tr>
      <th align="right" nowrap width="50%">
        <!--webbot bot='PurpleText' PREVIEW='Last Visit'--><%=fPhra(000164)%>
        :</th>
      <td><%=fFormatDate(svMembLastVisit)%></td>
    </tr>
    <% If fIsGroup2 Then %>
    <tr>
      <th align="right" nowrap width="50%"><!--webbot bot='PurpleText' PREVIEW='Account Expires'--><%=fPhra(001361)%>:</th>
      <td><%=fFormatDate(vCust_Expires)%></td>
    </tr>
    <% 
							Else		
				        If IsDate(svMembExpires) Then 
				          If svMembExpires > Now Then 
    %>
    <tr>
      <th align="right" nowrap width="50%"><!--webbot bot='PurpleText' PREVIEW='Programs Expires'--><%=fPhra(001312)%>:</th>
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


