<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Basket.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<!--#include virtual = "V5/Inc/EcomCountry.asp"-->


<% 
  '...determine media (need to know for PST and shipping reasons)
  Dim bPrePop

  vEcom_Media = Session("Ecom_Media")
  If fNoValue(vEcom_Media) Then Response.Redirect "EcomError.asp"
  sGetCust svCustId
  Session("vMemb_FirstName") = ""
  Session("vMemb_LastName")  = ""
  Session("vMemb_Email")     = ""

  '...pre populate if testing (note if email is pbulloch@vubiz.com then it bypasses I/S
  bPrePop = fIf(svServer <> "learn.vubiz.com" And svServer <> "www.learn.vubiz.com" And svServer <> "cloudweb.vubiz.com" And svServer <> "vubiz.com" And svServer <> "www.vubiz.com", True, False)
' previous code: (Not svEcomBypass And Lcase(svHost) <> "learn.vubiz.com/v5" And Lcase(svHost) <> "216.26.108.91/v5" And Lcase(svHost) <> "learn.vubiz.com/v5") Or Lcase(Session("MembId")) = "pbulloch@vubiz.com"  Then

  If bPrePop Then
    Session("xxxfirstname") = "Peter"
    Session("xxxlastname")  = "Bulloch"
    Session("xxxcompany")   = "none"
    Session("xxxaddress")   = "2112 Mississauga Rd"
    Session("xxxcity")      = "Mississauga"
    Session("xxxpostal")    = "L5H 2K6"
    Session("xxxprovince")  = "ON"
    Session("xxxcountry")   = "CA"
    Session("xxxphone")     = "905-891-9138"
    Session("xxxemail")     = "pbulloch@vubiz.com"
    Session("xxxcompany")   = "Vubiz Ltd"
  Else
    Session("xxxfirstname") = ""
    Session("xxxlastname")  = ""
    Session("xxxcompany")   = ""
    Session("xxxaddress")   = ""
    Session("xxxcity")      = ""
    Session("xxxpostal")    = ""
    Session("xxxprovince")  = ""
    Session("xxxcountry")   = ""
    Session("xxxphone")     = ""
    Session("xxxemail")     = ""
    Session("xxxcompany")   = ""
  End If

  Dim aErr(10), vOKtoSend
  Dim xxxName, xxxFirstName, xxxLastName, xxxCompany, xxxAddress, xxxCity, xxxPostal, xxxProvince, xxxCountry, xxxPhone, xxxEmail
  
  '...asssume all fields are bad, except 9
  For i = 0 To 8 : aErr(i) = True : Next
  aErr(9)  = False
  aErr(10) = True
  
  '...get values from session (links at bottom or basket) or customer form
  If Request.Form.Count = 0 Then  
  
    '...get values from session variables
    For Each vFld In Session.Contents
      If Lcase(Left(vFld, 3)) = "xxx" Or Lcase(Left(vFld, 1)) = "v" Then 
        vValue = Session.Contents(vFld)
'       sDebug vFld, vValue
        Select Case Lcase(vFld)
          Case "xxxfirstname"    : If vValue <> "" Then xxxFirstName    = vValue       : aErr(0)  = False
          Case "xxxlastname"     : If vValue <> "" Then xxxLastName     = vValue       : aErr(1)  = False
          Case "xxxcompany"      : If vValue <> "" Then xxxCompany      = vValue       : aErr(10) = False
          Case "xxxaddress"      : If vValue <> "" Then xxxAddress      = vValue       : aErr(2)  = False
          Case "xxxcity"         : If vValue <> "" Then xxxCity         = vValue       : aErr(3)  = False
          Case "xxxpostal"       : If vValue <> "" Then xxxPostal       = vValue       : aErr(4)  = False
          Case "xxxprovince"     : If vValue <> "" Then xxxProvince     = vValue 
          Case "xxxcountry"      : If vValue <> "" Then xxxCountry      = vValue  
          Case "xxxphone"        : If vValue <> "" Then xxxPhone        = vValue       : aErr(7)  = False
          Case "xxxemail"        : If vValue <> "" Then xxxEmail        = Lcase(vValue)    
          Case "xxxcompany"      : If vValue <> "" Then xxxCompany      = vValue  
          Case "vmemb_firstname" : If vValue <> "" Then vMemb_FirstName = vValue
          Case "vmemb_lastname"  : If vValue <> "" Then vMemb_LastName  = vValue
          Case "vmemb_email"     : If vValue <> "" Then vMemb_Email     = Lcase(vValue)
        End Select    
      End If
    Next
  
  Else  

    '...get values from form
    For Each vFld In Request.Form    
      vValue = Trim(Request.Form(vFld))

      '...store in Session Variable (do not unquote - will be done when we update db at last step)
      If Lcase(Left(vFld, 3)) = "xxx" Or Lcase(Left(vFld, 1)) = "v" Then 
        Session(vFld) = vValue
      End If
      
      Select Case Lcase(vFld)
        Case "xxxfirstname"    : If vValue <> "" Then xxxFirstName    = fUnQuote(Left(vValue & Space(50),   50)) : aErr(0)  = False
        Case "xxxlastname"     : If vValue <> "" Then xxxLastName     = fUnQuote(Left(vValue & Space(50),   50)) : aErr(1)  = False
        Case "xxxcompany"      : If vValue <> "" Then xxxCompany      = fUnQuote(Left(vValue & Space(128), 128)) : aErr(10) = False
        Case "xxxaddress"      : If vValue <> "" Then xxxAddress      = fUnQuote(Left(vValue & Space(128), 128)) : aErr(2)  = False
        Case "xxxcity"         : If vValue <> "" Then xxxCity         = fUnQuote(Left(vValue & Space(128), 128)) : aErr(3)  = False
        Case "xxxpostal"       : If vValue <> "" Then xxxPostal       = fUnQuote(Left(vValue & Space(128), 128)) : aErr(4)  = False
        Case "xxxprovince"     : If vValue <> "" Then xxxProvince     = fUnQuote(Left(vValue & Space(128), 128)) 
        Case "xxxcountry"      : If vValue <> "" Then xxxCountry      = fUnQuote(Left(vValue & Space(128), 128)) 
        Case "xxxphone"        : If vValue <> "" Then xxxPhone        = vValue                                   : aErr(7)  = False
        Case "xxxemail"        : If vValue <> "" Then xxxEmail        = Lcase(fUnQuote(Left(vValue & Space(128), 128)))    
        Case "xxxcompany"      : If vValue <> "" Then xxxCompany      = fUnQuote(Left(vValue & Space(128), 128))
        Case "vmemb_firstname" : If vValue <> "" Then vMemb_FirstName = fUnQuote(Left(vValue & Space(50),   50))
        Case "vmemb_lastname"  : If vValue <> "" Then vMemb_LastName  = fUnQuote(Left(vValue & Space(50),   50))
        Case "vmemb_email"     : If vValue <> "" Then vMemb_Email     = Lcase(fUnQuote(Left(vValue & Space(128), 128)))
      End Select    
    Next

  End If
  
  If xxxProvince <> "None" Or (xxxProvince = "None" And (xxxCountry <> "CA" And xxxCountry <> "US")) Then aErr(5) = False
  If xxxCountry  <> "" Then aErr(6) = False
  If Len(xxxEmail) > 0 And Instr(xxxEmail,"@") > 1 And Instr(xxxEmail,".") > 1 Then aErr(8) = False  
  If Len(vMemb_Email) > 0 Then
    If Instr(vMemb_Email,"@") > 1 And Instr(vMemb_Email,".")   > 1 Then 
      aErr(9) = False  
    Else
      aErr(9) = True
    End If
  End If
  
  '...determined if OK to Send to Checkout (from form only)
  If Request.Form.Count > 0 Then
    vOktoSend = True
    For i = 1 to 10
      If aErr(i) <> False Then vOKtoSend = False
    Next 

    If vOKtoSend Then 
      If Len(Trim(Session("vMemb_FirstName"))) = 0 Then Session("vMemb_FirstName") = Session("xxxFirstName")
      If Len(Trim(Session("vMemb_LastName")))  = 0 Then Session("vMemb_LastName")  = Session("xxxLastName")
      If Len(Trim(Session("vMemb_Email")))     = 0 Then Session("vMemb_Email")     = Session("xxxEmail")

      '...strip brackets from xxxCompany as we add in cust/id at checkout
      Session("xxxCompany") = Replace(Session("xxxCompany"), "(", "")
      Session("xxxCompany") = Replace(Session("xxxCompany"), ")", "")

      '...if there's a memo field, pass this through
      If Len(Session("Ecom_Memo")) > 0 Then vMemb_Memo = Session("Ecom_Memo")
          
      If (vEcom_Media = "Group2" Or vEcom_Media = "AddOn2") Then
        Response.Redirect "Ecom3Checkout.asp"
      Else
        Response.Redirect "Ecom2Checkout.asp"
      End If

    End If
  End If

  Function fProv (vProv)
    fProv = "" : If Ucase(vProv) = xxxProvince Then fProv = "selected" 
  End Function

  Function fCoun (vCoun)
    fCoun = "" : If Ucase(vCoun) = xxxCountry Then fCoun = "selected" 
  End Function

%>

<html>

<head>
  <title>Ecom2Customer</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js"></script>
  <script>
    function jTitle (vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
    }
  </script>


  <style>
    .notice {
      background-position: right; 
      background-color: #FFD5D5; 
      background-image: url('/V5/Images/Common/Back.gif'); 
      background-repeat: no-repeat;
    }
   #customer input[type=text] { width:300px; }
   #customer th { width:32%; }
   #customer td { width:65% }
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

<!--  <script>jTitle("<%=fPhraH(000085)%>", 'Carhholder.jpg')</script>-->
  <img src="../Images/Ecom/CardholderInformation_<%=svLang %>.png" />

  <h1><!--webbot bot='PurpleText' PREVIEW='Cardholder Information'--><%=fPhra(000085)%></h1>
  <h2><!--webbot bot='PurpleText' PREVIEW='Specific cardholder information is required for all e-commerce transactions.&nbsp; The e-commerce fields must correspond to the billing information of the cardholder.&nbsp; This also determines what currency and taxes apply.&nbsp; Programs sold outside of Canada are in $US.&nbsp; Arrows at right show fields requiring input.'--><%=fPhra(000412)%></h2>
  <h3><!--webbot bot='PurpleText' PREVIEW='Please click <b>Next</b> below to continue.'--><%=fPhra(000341)%></h3><br />

  <form method="POST" action="Ecom2Customer.asp" target="_self">
    <table class="table" id="customer">
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Cardholder First Name'--><%=fPhra(000084)%> : </th>
        <td <%=fIf(aErr(1), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxFirstName" value="<%=xxxFirstName%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> : </th>
        <td <%=fIf(aErr(1), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxLastName" value="<%=xxxLastName%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Organization'--><%=fPhra(000470)%> : </th>
        <td <%=fIf(aErr(10), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxCompany" value="<%=xxxCompany%>"> <!--webbot bot='PurpleText' PREVIEW='Enter None if not applicable'--><%=fPhra(001391)%></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Address'--><%=fPhra(000064)%> : </th>
        <td <%=fIf(aErr(2), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxAddress" maxlength="50" value="<%=xxxAddress%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='City'--><%=fPhra(000090)%> : </th>
        <td <%=fIf(aErr(3), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxCity" value="<%=xxxCity%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Postal Code'--><%=fPhra(000217)%> : </th>
        <td <%=fIf(aErr(4), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxPostal" value="<%=xxxPostal%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Province/State'--><%=fPhra(000277)%> : </th>
        <td <%=fIf(aErr(5), "class='notice'", "")%>>&nbsp;<select name="xxxProvince" size="1">
        <!--#include virtual = "V5/Inc/EcomProvince.asp"--></select> <!--webbot bot='PurpleText' PREVIEW='Canada &amp; US only'--><%=fPhra(000083)%></td>
      </tr>
      <tr><th><!--webbot bot='PurpleText' PREVIEW='Country'--><%=fPhra(000110)%> : </th>
        <td <%=fIf(aErr(6), "class='notice'", "")%>>&nbsp;<select name="xxxCountry" size="1"><%=sp5countryCodesDD(xxxCountry)%></select></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Phone Number'--><%=fPhra(000213)%> : </th>
        <td <%=fIf(aErr(7), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxPhone" value="<%=xxxPhone%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> : </th>
        <td <%=fIf(aErr(8), "class='notice'", "")%>>&nbsp;<input type="text" name="xxxEmail" value="<%=xxxEmail%>"></td>
      </tr>


      <tr>
        <td colspan="2" class="red" style="text-align:center; height:60px; vertical-align:middle;">
          <% If vEcom_Media = "Online" Then %> 
          <%   If svSecure Then %> 
          <!--webbot bot='PurpleText' PREVIEW='You are signed into the profile of the Learner below and the courses you are currently purchasing are for this Learner’s profile.<br><br>To purchase a course for a Learner OTHER than the one named below, you must log in first as that Learner and then proceed to purchase.'--><%=fPhra(001863)%>
          <%   Else %> 
          <!--webbot bot='PurpleText' PREVIEW='You MUST enter a Learner name below if it is different than the cardholder name above:'--><%=fPhra(001241)%>
          <%   End If %> 
          <% Else %> 
          <!--webbot bot='PurpleText' PREVIEW='You MUST enter a Facilitator name below if it is different than the cardholder name above:'--><%=fPhra(001242)%>
          <% End If %>
        </td>
      </tr>

      <tr>
        <th>
          <% =fIf(vEcom_Media = "Online", fPhraH(000165), fPhraH(000139)) %>&nbsp;<!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> : 
        </th>
        <td>&nbsp;
          <% If svSecure Then %>
            <% If svMembLevel = 5 Or svMembManager Then %>
              <input type="text" name="vMemb_FirstName" value="<%=fDefault(svMembFirstName, "None")%>">
            <% Else %>
              <% =fDefault(svMembFirstName, "None")%>
              <input type="hidden" name="vMemb_FirstName" value="<%=fDefault(svMembFirstName, "None")%>">
            <% End If %>
          <% Else %>
          <input type="text" name="vMemb_FirstName" value="<%=vMemb_FirstName%>">
          <% End If %>
        </td>
      </tr>
      <tr>
        <th>
          <% =fIf(vEcom_Media = "Online", fPhraH(000165), fPhraH(000139)) %>&nbsp;<!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> : 
        </th>
        <td>&nbsp;
          <% If svSecure Then %>
            <% If svMembLevel = 5 Or svMembManager Then %>
              <input type="text" name="vMemb_LastName" value="<%=fDefault(svMembLastName, "None")%>">
            <% Else %>
              <% =fDefault(svMembLastName, "None")%>
              <input type="hidden" name="vMemb_LastName" value="<%=fDefault(svMembLastName, "None")%>">
            <% End If %>
          <% Else %>
          <input type="text" name="vMemb_LastName" value="<%=vMemb_LastName%>">
          <% End If %>
        </td>
      </tr>
      <tr>
        <th>
          <% =fIf(vEcom_Media = "Online", fPhraH(000165), fPhraH(000139)) %>&nbsp;<!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> : 
        </th>
        <td>&nbsp;
          <% If svSecure Then %>
            <% If svMembLevel = 5 Or svMembManager Then %>
              <input type="text" name="vMemb_Email" value="<%=svMembEmail%>">
            <% Else %>
              <% =svMembEmail%>
              <input type="hidden" name="vMemb_Email" value="<%=svMembEmail%>">
            <% End If %>
          <% Else %>        
          <input type="text" name="vMemb_Email" value="<%=vMemb_Email%>">
          <% End If %>
        </td>
      </tr>

      <tr>
        <td colspan="2" style="text-align:center; vertical-align:middle; height:60px"><input type="submit" value="<%=bNext%>" name="bNext" class="button"></td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

