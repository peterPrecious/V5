<%
  '...similiar to Certificate.asp/Document.asp - variant for KMX

  '...this creates a complete URL for a document by encoding the parameters sent then sending it to the document web service
  Function fDocumentUrl (vFileName, vModsId, vLang, vCust, vAcctId, vProgId, vMemo)
    Dim vUrl, vParms 
    vParms = "" _
           & "&vFileName="  & vFileName _
           & "&vModsId="    & "" _
           & "&vLang="      & vLang _
           & "&vCust="      & vCust _
           & "&vAcctId="    & "" _
           & "&vProgId="    & "" _
           & "&vMemo="      & ""
 '...if there's any URLencoding, remove it
    vParms = URLDecode(vParms)
    '...typical post
    '   vCust=ERGP&vFileName=SexuallHarrasment.pdf&vAcctId=1234&vProgId=&vLang=EN&vModsId=&vMemo=&anticache=1953356480
    '...if not UTF8 then convert it for the cert service to handle accents, etc
    '   but if it is already UTF8, leave - this is possible as a custom cert sometimes "calls itself" (don't ask)
    If Not IsValidUTF8(vParms) Then vParms = EncodeUTF8(vParms)

    '...encode for security
    vParms = Base64Encode(vParms)
    vUrl   = "/DocService/Document.aspx?vParms="
    fDocumentUrl = vUrl & vParms

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
  
  
  '...encode base 64 encoded string
  PUBLIC FUNCTION Base64Encode(plain)
  
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
  PUBLIC FUNCTION Base64Decode(scrambled)
  
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