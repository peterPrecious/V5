<%
  '...this assumes that the code for creating base64 is loaded (documents/certificates, etc)
  Function fBase64(vParms)
    vParms = URLDecode(vParms)                                    '...if there's any URLencoding, remove it
    If Not IsValidUTF8(vParms) Then vParms = EncodeUTF8(vParms)   '...if not UTF8 then convert it to handle accents, etc
    vParms = Base64Encode(vParms)                                 '...encode for security
    fBase64 = vParms
  End Function
%>