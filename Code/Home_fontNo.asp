<script>
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

      '...if use in on staging
      If (Not svMembInternal) And svServer = "staging.vubiz.com" And Len(Session("Staging")) = 0 Then             
        Session("Staging") = True
        vAlert = fPhraH(001779)  
    %>
        alert("<%=vAlert%>");
    <% 
      End If
    %>



    
    $(function() { 

      // render popup status
      $("#popupStatus").html(parent.popupBlockerOn ? jYN("y", "<%=svLang%>") : jYN("n", "<%=svLang%>")); 

      // see if there's a cookie for the font size, else create one        
      var fontNo;
      var cookieOptions = { path: '/', expires: 365 };	
      
      // get fontNo from cookie on load else set to 3 and save in cookie
      fontNo = $.cookie("Profile_fontNo")
      if (isNumber(fontNo)) {
        if (fontNo < 1 || fontNo > 5)  {
          fontNo = 1;
          $.cookie("Profile_fontNo", fontNo, cookieOptions);
        }
      }

  
      // set/save font size
      if ($("#fontSizer").length > 0) {
        setFontSizer (fontNo);
      }
      else {
        setFontSize (fontNo);
      }

  
      // bind fontSizer
      $(".fontNo1").click(function() { setFontSizer("1"); ; return false; });
      $(".fontNo2").click(function() { setFontSizer("2"); ; return false; });
      $(".fontNo3").click(function() { setFontSizer("3"); ; return false; });
      $(".fontNo4").click(function() { setFontSizer("4"); ; return false; });
      $(".fontNo5").click(function() { setFontSizer("5"); ; return false; });


      // turn off other values and set selected fontNo to yellow background then change and save selection
      function setFontSizer(fontNo) {
        var classes = ".fontNo1, .fontNo2, .fontNo3, .fontNo4, .fontNo5";
        $(classes).css("background-color", "white");

        switch (fontNo) {
          case "1": $(".fontNo1").css("background-color", "yellow"); break;
          case "2": $(".fontNo2").css("background-color", "yellow"); break;
          case "3": $(".fontNo3").css("background-color", "yellow"); break;
          case "4": $(".fontNo4").css("background-color", "yellow"); break;
          case "5": $(".fontNo5").css("background-color", "yellow"); break;
        }

        setFontSize(fontNo);
        $.cookie("Profile_fontNo", null);
        $.cookie("Profile_fontNo", fontNo, cookieOptions);
      }

      // change fontSize and save fontNo
      function setFontSize (fontNo) {
        switch (fontNo) {
          case "1": $('html').css('font-size', '070%'); break;
          case "2": $('html').css('font-size', '080%'); break;
          case "3": $('html').css('font-size', '090%'); break;
          case "4": $('html').css('font-size', '100%'); break;
          case "5": $('html').css('font-size', '110%'); break;
          default : $('html').css('font-size', '070%'); break;
        }
      }

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

  '...determine if sponsored learner
  sGetMemb svMembNo

  '...get to see if pc or mac (for bookmarking)
  sGetQueryString
%>

<div align="left">
  <br>

  <%
    '...Show|Hide the alert at admin's request via url below: vAlert=y | vAlert=n
    If Application("Alert") = "y" Then
  %>
  <br><br>

  <h1><font color="#FF0000">:: <!--webbot bot='PurpleText' PREVIEW='NOTICE!'--><%=fPhra(001384)%></font></h1>
  <p>
    <% If svLang = "FR" Then %>
        Ce service sera en cours de maintenance de routine (y compris des améliorations pour nos clients) ne sera donc pas disponible le samedi 5 novembre à partir de 4 heures jusqu'à midi, HNE. Nous nous excusons pour tout inconvénient.
      <% ElseIf svLang = "ES" Then %>        
        Este servicio estará en mantenimiento de rutina (incluyendo mejoras para nuestros clientes), lo que no estará disponible el sábado 05 de noviembre desde las 4 hasta el mediodía hora del este. Nos disculpamos por cualquier inconveniente.
      <% Else %>
        This service will be undergoing routine maintenance (including enhancements for our customers) and will not be available on Saturday Nov 5th from 4am until noon EST. We apologize for any inconvenience.
      <% End If %>   
  </>
  <%
     End If  
  %>



    <% If vCust_Tab3 Then %>
    <% End If %>

    <h1><font color="#FF0000">::&nbsp;&nbsp; </font><%=Trim(vIntro)%></h1>
    <p>

      <% If vCust_Tab2 Then %>
      <!--webbot bot='PurpleText' PREVIEW='Click on the <b>My Learning</b> tab above to access your programs.'--><%=fPhra(000486)%>
      <% End If %>

      <% If vCust_Tab3 Then %>
      <!--webbot bot='PurpleText' PREVIEW='Click on the <b>My Content</b> tab above to access your free or purchased programs.'--><%=fPhra(000101)%>
      <% End If %>

  <% If vCust_Tab5 Then %>
  <!--webbot bot='PurpleText' PREVIEW='To purchase e-learning programs, click <b>More Content</b> to complete a secure e-commerce process.&nbsp;&nbsp; Any programs purchased will then appear under the <b>My Content</b> tab above.'--><%=fPhra(000513)%>
  <% End If %>

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
      If Len(vParagraph) > 0 Then Response.Write vParagraph
    End If 
  %>

  <!--webbot bot='PurpleText' PREVIEW='If you have any questions or comments please contact us using the link at the bottom of the page.'--><%=fPhra(000150)%>


  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='Important'--><%=fPhra(000153)%></h1>
  <p>
    <!--webbot bot='PurpleText' PREVIEW='Please do NOT <b>Bookmark</b> or <b>Add to Favorites</b> the address that appears in your web browser when you are logged in. You must login with your user credentials each time you enter this service. Please click the Sign Off tab at the end of each visit.'--><%=fPhra(001318)%></p>


  <% If vCust_MaxSponsor > 0 And vMemb_Sponsor = 0 Then '...if accounts allow sponsors then ensure this link is not for a sponsored learner %>
  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='My Sponsored Learners'--><%=fPhra(000514)%></h1>
  <h2>
  <!--webbot bot='PurpleText' PREVIEW='If you would like to offer members of your organization access to your content, click ...'--><%=fPhra(000822)%>
  <a href="Sponsors.asp"><u>
    <!--webbot bot='PurpleText' PREVIEW='Sponsored Learners.'--><%=fPhra(000515)%></u></a>
  </h2>
  <% End If %>

  <% If vCust_Scheduler Then %>
  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='My Scheduler'--><%=fPhra(001252)%></h1>
  <h2 align="left"><a href="Scheduler.asp"><u>
    <!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></u></a>
  <!--webbot bot='PurpleText' PREVIEW='to view your calendar.'--><%=fPhra(001254)%></h2>
  <% End If %>


  <% If vCust_InfoEditProfile Then %>
  <h1><a <%=fStatX%> name="MyProfile"></a><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='My Profile'--><%=fPhra(000185)%></h1>
  <p>
    <!--webbot bot='PurpleText' PREVIEW='Enter/edit your name and email address below. Your name will then appear on any certificates issued for successful completion of assessments or exams.'--><%=fPhra(000129)%></p>

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
      <table border="1" id="table1" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="5" bgcolor="#F2F9FD" width="300" cellspacing="0">
        <tr>
          <td>
            <table border="0" style="border-collapse: collapse" id="table2" width="100%" cellpadding="2">
              <tr>
                <th align="right" nowrap width="50%" valign="bottom">
                <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> :</th>
                <td width="50%" nowrap valign="bottom">
                  <% If Request.QueryString("vAction") = "edit" Then %><input type="text" name="vMemb_FirstName" size="19" value="<%=svMembFirstName%>" maxlength="32">
                  <% Else %>
                  <%=svMembFirstName%>
                  <% End If %> 
                </td>
              </tr>
              <tr>
                <th align="right" nowrap width="50%" valign="bottom">
                <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> :</th>
                <td width="50%" nowrap valign="bottom">
                  <% If Request.QueryString("vAction") = "edit" Then %><input type="text" name="vMemb_LastName" size="19" value="<%=svMembLastName%>" maxlength="64">
                  <% Else %>
                  <%=svMembLastName%>
                  <% End If %> 
                </td>
              </tr>

              <% If vCust_Pwd And svMembLevel = 2 Then %>
              <%'If vCust_Pwd And svMembLevel < 5 Then %>
              <tr>
                <th align="right" nowrap width="50%" valign="bottom">
                <!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%> :</th>
                <td width="50%" nowrap valign="bottom">
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
                <th align="right" nowrap width="50%" valign="bottom">
                <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> :</th>
                <td width="50%" nowrap valign="bottom">
                  <% If Request.QueryString("vAction") = "edit" Then %>
                    <input type="text" name="vMemb_Email" size="19" value="<%=svMembEmail%>">
                  <% Else %>
                    <%=fDefault(svMembEmail, "...<i><font color='#FF0000'>[none]")%> 
                  <% End If %> 
                </td>
              </tr>

              <% If svLang = "EN" And vCust_VuNews Then %>
              <tr>
                <th align="right" nowrap width="50%" valign="bottom">Send vuNews <b><a href="javascript:toggle('Div_VuNews');"> ?</a></b></th>
                <td width="50%" nowrap valign="bottom"><p>
                  <% If Request.QueryString("vAction") = "edit" Then %>
                    <input type="radio" value="1" name="vMemb_VuNews" <%=fcheck(fsqlboolean(vMemb_VuNews), 1)%>>Yes&nbsp; 
                    <input type="radio" value="0" name="vMemb_VuNews" <%=fcheck(fsqlboolean(vMemb_VuNews), 0)%>>No&nbsp;
                  <% Else %>
                    <%=fIf(vMemb_VuNews, "Yes", "No")%>
                  <% End If %><%=f5%>
                </td>
              </tr>
              <tr>
                <th nowrap colspan="2" valign="bottom">
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
          </td>
        </tr>
      </table>
    </div>
  </form>

  <% End If %>


  <% If vCust_Tab4Type= "GA" Then 
       If svMembGap Or svMembLevel = 5 Or vMemb_GapPositionNo > 0 Then 
  %>
  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='My Development Plan'--><%=fPhra(000935)%></h1>
  <p align="left">
    <!--webbot bot='PurpleText' PREVIEW='Click on the <b>My Development</b> tab above for your Self Assessment, Performance Rating and Development Plan.'--><%=fPhra(000936)%></p>
  <%   End If 
     End If 
  %>


  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='My Status'--><%=fPhra(001362)%></h1>

  <p>
    <a href="LearnerReportCard2.asp?vMemb_No=<%=svMembNo%>&vInfoPage=y">
      <!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a>
    <!--webbot bot='PurpleText' PREVIEW='for your Report Card.'--><%=fPhra(001365)%>
    <% If svMembLevel > 4 Then %><br>
    <a href="RTE_History_F.asp?vPass=<%=svMembId%>&vInfoPage=y">
      <!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a>
    <!--webbot bot='PurpleText' PREVIEW='for your Learning History [coming].'--><%=fPhra(001433)%>
    <% End If %>
  </p>

  <%
    '...data is passed in as svBrowser and determined in the initial default.asp
    Dim aTools, vTouch, vHTML5
    aTools = Split(Ucase(svBrowser), "|")
    vTouch   = fYN (aTools(0))
    vBrowser = aTools(1)
    vHTML5   = fYN (aTools(2)) 
  %>


  <div align="center">
    <table border="1" id="table1" style="border-collapse: collapse" bordercolor="#DDEEF9" bgcolor="#F2F9FD" width="300" cellpadding="5">
      <tr>
        <td>
          <table border="0" style="border-collapse: collapse" id="table2" width="100%">

            <tr>
              <th align="right" nowrap width="50%" valign="bottom"><!--webbot bot='PurpleText' PREVIEW='Font Size'--><%=fPhra(001434)%> :</th>
              <td width="50%" nowrap valign="bottom">
                <div id="fontSizer">
                <div class="fontNo1">1</div>
                <div class="fontNo2">2</div>
                <div class="fontNo3">3</div>
                <div class="fontNo4">4</div>
                <div class="fontNo5">5</div>
                </div>
              </td>
            </tr>

            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='Popups Blocked'--><%=fPhra(001435)%> :</th>
              <td id="popupStatus" width="50%"></td>
            </tr>
            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='Touch Screen'--><%=fPhra(001436)%> :</th>
              <td id="touchStatus" width="50%"><%=vTouch%></td>
            </tr>
            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='Browser'--><%=fPhra(001363)%>
              :</th>
              <td width="50%"><%=vBrowser%></td>
            </tr>
            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='HTML5'--><%=fPhra(001437)%>
              :</th>
              <td width="50%"><%=vHTML5%></td>
            </tr>

            <tr>
              <th align="right" nowrap width="50%">
              &nbsp;</th>
              <td width="50%">&nbsp;</td>
            </tr>





            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='First Visit'--><%=fPhra(000157)%>
              :</th>
              <td width="50%"><%=fFormatDate(svMembFirstVisit)%></td>
            </tr>
            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='Last Visit'--><%=fPhra(000164)%>
              :</th>
              <td width="50%"><%=fFormatDate(svMembLastVisit)%></td>
            </tr>

            <% If fIsGroup2 Then %>

            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='Account Expires'--><%=fPhra(001361)%>
              :</th>
              <td width="50%"><%=fFormatDate(vCust_Expires)%></td>
            </tr>

            <% 
							Else
			
				        If IsDate(svMembExpires) Then 
				          If svMembExpires > Now Then 
            %>
            <tr>
              <th align="right" nowrap width="50%">
              <!--webbot bot='PurpleText' PREVIEW='Programs Expires'--><%=fPhra(001312)%>
              :</th>
              <td width="50%"><%=fFormatDate(svMembExpires)%></td>
            </tr>
            <%
				          End If
				        End If
						  End If 
            %>
          </table>
        </td>
      </tr>
    </table>
  </div>

  <% If svLang = "EN" Then %>
  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>
    <!--webbot bot='PurpleText' PREVIEW='Help using this service'--><%=fPhra(000875)%></h1>
  <p><a href="../Public/21_FAQ.asp?vReturn=y">
    <!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a>
    <!--webbot bot='PurpleText' PREVIEW='if you have any questions about how this service works.'--><%=fPhra(000878)%></p>
  <% End If %>

  <% If svLang = "FR" Then %>
  <h1><font color="#FF0000">::&nbsp;&nbsp; </font>Problèmes liés aux navigateurs?</h1>
  <p><a href="../Public/BrowserIssues_FR.htm?vReturn=Y">Cliquez ici</a> pour options de réglage de votre navigateur Web</p>
  <% End If %>

  <br><br>
</div>


