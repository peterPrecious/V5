<%        

    '...for CAAM only, get vMemo from Scorm

    Dim Account, MemberID, ProgramCode, vScore, vFirstName, vLastName, vModId, vTitle, vProgId, vLang, vAcctId, vDate
                      
    Account = Ucase(fDefault(Request("Account"), "3001"))
    MemberID = Ucase(fDefault(Request("MemberID"), "43015"))
    ProgramCode = Ucase(fDefault(Request("ProgramCode"), "P2313EN"))   

    //EXEC [dbo].[spGetProgramCertificate]                       
    //@Account = N'3001',
    //@MemberID = N'43015',
    //@ProgramCode = N'P2313EN'

    Set oDbVuGold = Server.CreateObject("ADODB.Connection")
    oDbVuGold.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=C8WDEzy9HPzjnDpWcFYm5UXk;Persist Security Info=True;User ID=apps;Initial Catalog=vuGold;Data Source=vmsql-01"
    oDbVuGold.Open
                         
    vSql = "EXEC [dbo].[spGetProgramCertificate] '" & Account & "', '"  & MemberID & "', '" & ProgramCode & "'"
    Set oRsBase = oDbVuGold.Execute(vSql)
    If Not oRsBase.Eof Then               
      vFirstName = oRsBase("memFirstName")
      vLastName = oRsBase("memLastName")  
      vTitle = oRsBase("prgTitle")       
      vDate = oRsBase("pcnCompleted")
      //Do While Not oRsBase.EOF 
        //fModsOptions = fModsOptions & "<option>" & oRsBase("Mods_Id") & "</option>" & vbCRLF
        //oRsBase.MoveNext
      //Loop      
    End If

    Set oRsBase = Nothing
    oDbVuGold.Close
    Set oDbVuGold = Nothing             
    
'   vScore = 80  ...caam does not use scores
    vScore = ""
    vModId = ""                   
    vProgId = ProgramCode       
    vLang = "EN"
    vCust = ""   
    vAcctId = Account
                                             
    const BASE_64_MAP_INIT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    dim nl
    ' zero based arrays
    dim Base64EncMap(63)
    dim Base64DecMap(127)

    ' initialize
    call initCodecs

    Response.Redirect  fCertificateUrl(vFirstName, vLastName, vScore, vDate, vModId, vTitle, vLang, vCust, vAcctId, vProgId, "", "")

//***************************************
    
  Function fDefault(i, j)
    If fNoValue(i) Then
      fDefault = j
    Else
      fDefault = i
    End If  
  End Function       
  
  '...is value null, empty or ""
  Function fNoValue (vTemp)
    fNoValue = False
    If VarType (vTemp) = vbEmpty Or VarType (vTemp) = vbNull Or vTemp = "" Then fNoValue = True  
  End Function
                          
  Function fCertificateUrl (vFirstName, vLastName, vScore, vDate, vModsId, vTitle, vLang, vCust, vAcctId, vProgId, vLogo, vMemo)
    Dim vUrl, vParms
    vParms = "" _
           & "&vFirstName=" & fDefault(vFirstName, "Unknown") _
           & "&vLastName="  & fDefault(vLastName, "Unknown") _
           & "&vScore="     & vScore _
           & "&vDate="      & fFormatDate(fIf(IsDate(vDate), vDate, Now())) _
           & "&vModsId="    & vModsId _
           & "&vTitle="     & vTitle  _
           & "&vLang="      & vLang _
           & "&vCust="      & vCust _
           & "&vAcctId="    & vAcctId _
           & "&vProgId="    & Left(vProgId, 5) _
           & "&vLogo="      & vLogo //_
           //& "&vMemo="      & fDefault(vMemo, svMembCriteria & "|" & svMembEmail & "|" & fProgHours(vProgId))   '...if no memo is passed in put through a few other key values
    '...if there's any URLencoding, remove it
    vParms = URLDecode(vParms)
    '...if not UTF8 then convert it for the cert service to handle accents, etc
    '   but if it is already UTF8, leave - this is possible as a custom cert sometimes "calls itself" (don't ask)
    If Not IsValidUTF8(vParms) Then vParms = EncodeUTF8(vParms)
    '...encode for security
    vParms = Base64Encode(vParms)
    vUrl   = "/CertService/Default.aspx?format=PDF&vParms="
    //vUrl   = "//learn.vubiz.com/CertService/Default.aspx?format=PDF&vParms="
    fCertificateUrl = vUrl & vParms
  End Function                     

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
  
  Function fFormatDate (i)
    Dim aMonth
'   If i = "" Then fFormatDate = "" : Exit Function '...if they clear out the date leave empty
    fFormatDate = " "
    If Not IsDate (i) Then Exit Function
    If Year(i) < 2000 Then Exit Function
    Select Case svLang
      Case "FR" : aMonth = Split ("janv. févr. mars avril mai juin juillet août sept. oct. nov. déc.", " ") : fFormatDate = Day(i) & " " & aMonth(Month(i) -1) & " " & Year(i)                 
      Case "ES" : aMonth = Split ("ene. feb. mar. abr. may. jun. jul. ago. sept. oct. nov. dic.", " ")      : fFormatDate = Day(i) & " " & aMonth(Month(i) -1) & " " & Year(i)
      Case Else : aMonth = Split ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", " ")                   : fFormatDate = aMonth(Month(i) -1) & " " & Right("00" & Day(i), 2) & ", " & Year(i)
    End Select
  End Function
  
  Function fIf(i, j, k)
    fIf = k
    If Vartype(i) = 11 Then     
      If i Then fIf = j
    End If
  End Function
  
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
%>