<%
' response.write fEncode("governance") & "<br>" & fDecode(fEncode("governance"))

  '____ Passwords for MyWorld ______________________________________________

  Function fDecode (vPassword)
    Const vAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_"
    fDecode = ""
    For i = 1 To Len(vPassword) Step 4
      fDecode = fDecode & Mid(vAlpha, Cint(Mid(vPassword, i, 4)) - 4141, 1)
    Next
  End Function

  '...insert url and return url no
  Function fEncode (vPassword)
    Const vAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_"
    fEncode = ""
    For i = 1 to Len(vPassword)
      fEncode = fEncode & Right("0000" & Instr(vAlpha, Mid(Ucase(vPassword), i, 1)) + 4141, 4)
    Next
  End Function

%>