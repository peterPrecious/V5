<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
  Dim vOk, vBrowserOk, aRogue, vError

  vHostDb     = fNoQuote(Ucase(Trim(Request.QueryString("vHostDb")))) : If fNoValue(vHostDb)     Then vHostDb     = "V5_Vubz"
  vBrowserOk  = Trim(Lcase(Request("vBrowserOk")))                    : If vBrowserOk <> "y"     Then vBrowserOk  = "n"

  vCust       = fNoQuote(Ucase(Trim(Request.QueryString("vCust"))))   : If fNoValue(vCust)       Then vCust       = ""
  vId         = fNoQuote(Ucase(Trim(Request.QueryString("vId"))))     : If fNoValue(vId)         Then vId         = ""
  vPwd        = fNoQuote(Ucase(Trim(Request.QueryString("vPwd"))))    : If fNoValue(vPwd)        Then vPwd        = ""
  vFirstName  = Trim(Request.QueryString("vFirstName"))               : If fNoValue(vFirstName)  Then vFirstName  = ""
  vLastName   = Trim(Request.QueryString("vLastName"))                : If fNoValue(vLastName)   Then vLastName   = ""
  vEmail      = fNoQuote(Trim(Request.QueryString("vEmail")))         : If fNoValue(vEmail)      Then vEmail      = ""
  vMemo       = fNoQuote(Trim(Request.QueryString("vMemo")))          : If fNoValue(vMemo)       Then vMemo       = ""
  vTraining   = fNoQuote(Trim(Request.QueryString("vTraining")))      : If fNoValue(vTraining)   Then vTraining   = ""
  vCriteria   = Trim(Request.QueryString("vCriteria"))                : If fNoValue(vCriteria)   Then vCriteria   = ""
  vAction     = fNoQuote(Ucase(Trim(Request.QueryString("vAction")))) : If fNoValue(vAction)     Then vAction     = ""
  vLang       = fNoQuote(Ucase(Trim(Request.QueryString("vLang"))))   : If fNoValue(vLang)       Then vLang       = ""
  vErrUrl     = fNoQuote(Trim(Request.QueryString("vErrUrl")))        : If fNoValue(vErrUrl)     Then vErrUrl     = ""
  vQModId     = Trim(Request.QueryString("vQModId"))                  : If fNoValue(vQModId)     Then vQModId     = ""
  vQTestId    = Trim(Request.QueryString("vQTestId"))                 : If fNoValue(vQTestId)    Then vQTestId    = ""
  vTskH_Id    = Trim(Request.QueryString("vTskH_Id"))                 : If fNoValue(vTskH_Id)    Then vTskH_Id    = ""
  vGoTo       = Trim(Request.QueryString("vGoTo"))                    : If fNoValue(vGoTo)       Then vGoTo       = ""
  v8          = Trim(Request.QueryString("v8"))                       : If fNoValue(v8)          Then v8          = "n"
  vBrowser    = Request.QueryString("vBrowser")

	'... values should arrive urlEncoded then we need to twig to get past our querystring routine
  vSource = fNoQuote(Trim(Request.QueryString("vSource"))) 
  vSource = Replace (vSource, "&", "~1")
  vSource = Replace (vSource, "=", "~2")
  vSource = Replace (vSource, "?", "~3")

  Session("Secure")         = False
  Session("CustReturnUrl")  = vSource
  Session("Translate")      = False
  Session("TabActive")      = False '...use to see if tabs used for shell purposes
  
  '...if there's a request to purchase content from a landing page with an associated "agent" this will be passed through the system and stored on the ecommerce file
  If Len(Request("vEcomAgent")) > 0 Then Session("EcomAgent") = Request("vEcomAgent")

  '...if customer and no language then get language
  If vCust <> "" And vLang = "" Then
    Session("HostDb") = vHostDb                                            
    svHostDb = Session("HostDb")
    sGetCust (vCust)
    If vFileOk Then 
      vLang = vCust_Lang
    End If
  End If

  If Len(vLang) = 0 Or Instr(" EN FR ES ", vLang) = 0 Then
    vLang = "EN"
  End If

  '...protect against rogue sql injection - look for any field containing a space
  aRogue = Split("ALTER DROP DELETE INSERT UPDATE # ' = --")
  vOk = True
  For Each vFld In Request.QueryString
    vValue = Ucase(Trim(Request.QueryString(vFld)))
    If Instr(vValue, " ") > 0 Then
      For i = 0 To Ubound(aRogue)
        If Instr(" " & vValue & " ", " " & aRogue(i) & " ") > 0 Then vOk = False : Exit For
        If Left(vValue, Len(aRogue(i))) & " " = aRogue(i) & " "  Then vOk = False : Exit For
      Next
      If Not vOk Then
        vError = "<!--{{-->the values entered contain illegal characters.  Please contact Vubiz.<!--}}-->"
        vError = vError & "<br>(...ALTER DROP DELETE INSERT UPDATE...)"
        Response.Redirect "Code/SignInErr.asp?vClose=Y&vError=" & Server.UrlEncode(vError) & "&vLang=" & vLang
      End If  
    End If
  Next

  sPutQueryString
  sGetQueryString

  Response.Redirect "Code/SignIn.asp?vLang=" & vLang
' Response.Redirect "Source/SignIn.asp?vLang=" & vLang


  '...is value null, empty or ""
  Function fNoQuote (vTemp)
    fNoQuote = Replace(vTemp, "'", "")
  End Function

  '...remove sinqle quotes
  Function fNoValue (vTemp)
    fNoValue = False
    If VarType (vTemp) = vbEmpty Or VarType (vTemp) = vbNull Or vTemp = "" Then fNoValue = True  
  End Function

%>