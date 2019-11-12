<% 
//Response.Write "vDate=" & fFormatDate(Now())
  Response.Write "vDate=" & Server.UrlEncode(fFormatDate(Now()))
  
  Function fFormatDate (i)
    Dim aMonth, vLang
    fFormatDate = " "
    vLang = Request("vLang")
    If Len(vLang) = 0 Then vLang = "EN"    
    Select Case vLang
      Case "EN" : aMonth = Split ("January February March April May June July August Septempber October November December", " ")  : fFormatDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
      Case "FR" : aMonth = Split ("janvier février mars avril mai juin juillet août septembre octobre novembre décembre", " ")    : fFormatDate = Right("00" & Day(i), 2) & " " & aMonth(Month(i) -1) & " " & Year(i)
      Case "ES" : aMonth = Split ("Ene Feb Mar Abr May Jun Jul Ago Sep Oct Nov Dic", " ")                                         : fFormatDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
      Case "PT" : aMonth = Split ("Jan Fev Mar Abr Mai Jun Jul Ago Set Out Nov Dez", " ")                                         : fFormatDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
      Case "PL" : aMonth = Split ("sty. luty mar. kwi. maj cze. lip. sie. wrz. pa&#378;. lis. gru.", " ")                         : fFormatDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
    End Select
  End Function  
  
%>