<%
  '...similiar to Document.asp
  '...this creates a complete URL for a certificate by encoding the parameters sent then sending it to the certificate web service
  Function fCertificateUrl (vFirstName, vLastName, vScore, vDate, vModsId, vTitle, vLang, vCust, vAcctId, vProgId, vLogo, vMemo, vEmail)
    Dim vUrl, vParms, vFormat
    vParms = "" _
           & "&vFirstName=" & fDefault(vFirstName, svMembFirstName) _
           & "&vLastName="  & fDefault(vLastName, svMembLastName) _
           & "&vScore="     & vScore _
           & "&vDate="      & fFormatDate(fIf(IsDate(vDate), vDate, Now())) _
           & "&vModsId="    & vModsId _
           & "&vTitle="     & vTitle  _
           & "&vLang="      & fDefault(vLang, svLang) _
           & "&vCust="      & Left(fDefault(vCust, svCustId), 4) _
           & "&vAcctId="    & fDefault(vAcctId, svCustAcctId) _
           & "&vProgId="    & Left(vProgId, 5) _
           & "&vLogo="      & fDefault(vLogo, svCustBanner) _
           & "&vMemo="      & svMembCriteria & "|" & svMembEmail & "|" & fProgHours(vProgId) & "|" & vMemo & "|" & fProgNasbaCpe (vProgId) _
           & "&vEmailTo="   & fIf(vEmail = "", svMembEmail, vEmail) _    
           & "&vEmailFrom=" & fIf(svCustEmail = "", "info@vubiz.com", fIf(svCustEmail="none", "", svCustEmail))


    '...if there's any URLencoding, remove it
    vParms = URLDecode(vParms)
    '...if not UTF8 then convert it for the cert service to handle accents, etc
    '   but if it is already UTF8, leave - this is possible as a custom cert sometimes "calls itself" (don't ask)
    If Not IsValidUTF8(vParms) Then vParms = EncodeUTF8(vParms)

    '...encode for security (added in Server.URLEncode to handle accents)
    vParms = base64Encode(vParms) 

    vFormat = "jpeg"
    vUrl   = "/CertService/Default.aspx?format=" & vFormat & "&vParms="
    fCertificateUrl = vUrl & vParms

  End Function    

  '... Next two functions were added Jan 20, 2020 to improve base64encoding/decoding
  '... call a web service in INC to actually do the work
  '... key to accent handling was the Server.URLEncode(plainText)
  Function base64Encode(plainText)
    Dim xmlhttp, dataToSend, postUrl
'   dataToSend = "plainText=" & plainText
    dataToSend = "plainText=" & Server.URLEncode(plainText)
    postUrl = "http://" & Request.ServerVariables("HTTP_HOST") & "/V5/Inc/base64.asmx/base64Encode"
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", postUrl, false
    xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    xmlhttp.send dataToSend
    base64Encode = xmlhttp.responseXML.text
  End Function

  Function base64Decode(encodedText)
    Dim xmlhttp, dataToSend, postUrl
    dataToSend = "base64EncodedData=" & encodedText
    postUrl = "http://" & Request.ServerVariables("HTTP_HOST") & "/V5/Inc/base64.asmx/base64Decode"
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", postUrl, false
    xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    xmlhttp.send dataToSend
    base64Decode = xmlhttp.responseXML.text
  End Function




















 
  '...get number of hours in this program for those that want CE credits shown on their cert (vMemo)
  Function fProgHours(vProgId)
    Dim vMods, aMods
    fProgHours = 0
    sOpenDbBase    
    vSql = "SELECT Prog_Mods FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      vMods = Trim(oRsBase("Prog_Mods"))
      If Len(vMods) > 0 Then 
        aMods = Split(Trim(oRsBase("Prog_Mods")))
        For i = 0 to uBound(aMods)
          vSql = "SELECT Mods_Length FROM Mods WHERE Mods_Id= '" & aMods(i) & "'"
          Set oRsBase = oDbBase.Execute(vSql)
          If Not oRsBase.Eof Then fProgHours = fProgHours + oRsBase("Mods_Length")
        Next
      End If
    End If
    Set oRsBase = Nothing
    sCloseDbBase 
  End Function


    '...Get Prog Nasba_Cpe  (added Sep 11, 2018 to add to certificates - used in /inc/certificate.asp)
  Function fProgNasbaCpe (vProgId)
    Dim oRs
    fProgNasbaCpe = ""
    vSql = "SELECT Prog_Nasba_Cpe FROM Prog WHERE Prog_Id= '" & vProgId & "'"
    sOpenDbBase    
    Set oRs = oDbBase.Execute(vSql)
    If Not oRs.Eof Then 
      fProgNasbaCpe = oRs("Prog_Nasba_Cpe")
    End If
    Set oRs = Nothing
    sCloseDbBase    
  End Function


  Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If	
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")	
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")	
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If	
    URLDecode = sOutput
  End Function


  '...Determine if the string is valid UTF-8 encoded, Returns: true (valid UTF-8), false (invalid UTF-8 or not UTF-8 encoded string)
  function IsValidUTF8(s)
    dim i, c, n
    IsValidUTF8 = false
    i = 1
    do while i <= len(s)
      c = asc(mid(s,i,1))
      if c and &H80 then
        n = 1
        do while i + n < len(s)
          if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
            exit do
          end if
          n = n + 1
        loop
        select case n
        case 1
          exit function
        case 2
          if (c and &HE0) <> &HC0 then
            exit function
          end if
        case 3
          if (c and &HF0) <> &HE0 then
            exit function
          end if
        case 4
          if (c and &HF8) <> &HF0 then
            exit function
          end if
        case else
          exit function
        end select
        i = i + n
      else
        i = i + 1
      end if
    loop
    IsValidUTF8 = true 
  end function


  '...Encodes a Windows string in UTF-8, Returns: A UTF-8 encoded string
  function EncodeUTF8(s)
    dim i, c
    i = 1
    do while i <= len(s)
      c = asc(mid(s,i,1))
      if c >= &H80 then
        s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
        i = i + 1
      end if
      i = i + 1
    loop
    EncodeUTF8 = s 
  end function

  '...Decodes a UTF-8 string to the Windows character set Non-convertable characters are replace by an upside down question mark.  
  '  Returns a Windows string
  function DecodeUTF8(s)
    dim i, c, n
    i = 1
    do while i <= len(s)
      c = asc(mid(s,i,1))
      if c and &H80 then
        n = 1
        do while i + n < len(s)
          if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
            exit do
          end if
          n = n + 1
        loop
        if n = 2 and ((c and &HE0) = &HC0) then
          c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
        else
          c = 191 
        end if
        s = left(s,i-1) + chr(c) + mid(s,i+n)
      end if
      i = i + 1
    loop
    DecodeUTF8 = s 
  end function



  const BASE_64_MAP_INIT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  dim nl
  ' zero based arrays
  dim Base64EncMap(63)
  dim Base64DecMap(127)
  
  ' must be called before using anything else
  PUBLIC SUB initCodecs()
      ' init vars
      nl = "<P>" & chr(13) & chr(10)
      ' setup base 64
      dim max, idx
         max = len(BASE_64_MAP_INIT)
      for idx = 0 to max - 1
           ' one based string
           Base64EncMap(idx) = mid(BASE_64_MAP_INIT, idx + 1, 1)
      next
      for idx = 0 to max - 1
           Base64DecMap(ASC(Base64EncMap(idx))) = idx
      next
  END SUB
  
  '...these are the older functions, replaced by above

  '...encode base 64 encoded string
  PUBLIC FUNCTION xBase64Encode(plain)
  
      if len(plain) = 0 then
           base64Encode = ""
           exit function
      end if
  
      dim ret, ndx, by3, first, second, third
      by3 = (len(plain) \ 3) * 3
      ndx = 1
      do while ndx <= by3
           first  = asc(mid(plain, ndx+0, 1))
           second = asc(mid(plain, ndx+1, 1))
           third  = asc(mid(plain, ndx+2, 1))
           ret = ret & Base64EncMap(  (first \ 4) AND 63 )
           ret = ret & Base64EncMap( ((first * 16) AND 48) + ((second \ 16)AND 15 ) )
           ret = ret & Base64EncMap( ((second * 4) AND 60) + ((third \ 64)AND 3 ) )
           ret = ret & Base64EncMap( third AND 63)
           ndx = ndx + 3
      loop
      ' check for stragglers
      if by3 < len(plain) then
           first  = asc(mid(plain, ndx+0, 1))
           ret = ret & Base64EncMap(  (first \ 4) AND 63 )
           if (len(plain) MOD 3 ) = 2 then
                second = asc(mid(plain, ndx+1, 1))
                ret = ret & Base64EncMap( ((first * 16) AND 48) +((second \16) AND 15 ) )
                ret = ret & Base64EncMap( ((second * 4) AND 60) )
           else
                ret = ret & Base64EncMap( (first * 16) AND 48)
                ret = ret & "="
           end if
           ret = ret & "="
      end if
  
      Base64Encode = ret
  END FUNCTION



  '...decode base 64 encoded string
  PUBLIC FUNCTION xBase64Decode(scrambled)
  
      if len(scrambled) = 0 then
           base64Decode = ""
           exit function
      end if
  
      ' ignore padding
      dim realLen
      realLen = len(scrambled)
      do while mid(scrambled, realLen, 1) = "="
           realLen = realLen - 1
      loop
      dim ret, ndx, by4, first, second, third, fourth
      ret = ""
      by4 = (realLen \ 4) * 4
      ndx = 1
      do while ndx <= by4
           first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
           second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
           third  = Base64DecMap(asc(mid(scrambled, ndx+2, 1)))
           fourth = Base64DecMap(asc(mid(scrambled, ndx+3, 1)))
           ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3) )
           ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
           ret = ret & chr( ((third * 64) AND 255) +  ((fourth AND 63)) )
           ndx = ndx + 4
      loop
      ' check for stragglers, will be 2 or 3 characters
      if ndx < realLen then
           first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
           second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
           ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
           if realLen MOD 4 = 3 then
                third = Base64DecMap(asc(mid(scrambled,ndx+2,1)))
                ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
           end if
      end if
  
      Base64Decode = ret
  END FUNCTION

' initialize
  call initCodecs

  
%>