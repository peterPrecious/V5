<% 
  '...used after July 1, 2010 to handle complex taxes across Canada
  '   before we just used the setup.asp file
  
   
  Function fCurrency (vCountry)
    fCurrency = fIf(vCountry <> "CA", "US", "CA")
  End Function


  Function fGST (vDate, vCountry, vProvince)
    fGST = 0
    If vCountry <> "CA" Then Exit Function
    If cDate(vDate) < cDate("Jan 01, 2008") Then
      If Instr(" NS NB NF ", vProvince) = 0 Then
        fGST = .07
      End if
    ElseIf cDate(vDate) < cDate("Jul 01, 2010") Then
      If Instr(" NS NB NF ", vProvince) = 0 Then
        fGST = .05
      End if
    Else 
      If Instr(" AB BC MB NT NU PE PQ QC SK YT ", vProvince) > 0 Then
        fGST = .05
      End If
    End If
  End Function


  Function fHST (vDate, vCountry, vProvince)
    fHST = 0
    If vCountry <> "CA" Then Exit Function
    If cDate(vDate) < cDate("Jul 01, 2010") Then
      If Instr(" NS NB NF ", vProvince) > 0 Then
        fHST = .13
      End if
    Else
      Select Case vProvince
        Case "NB","NF","ON" : fHST = .13
        Case "NS"           : fHST = .15
      End Select
    End If
  End Function 

  
  Function fPST (vDate, vCountry, vProvince)
    fPST = 0
    If vCountry <> "CA" Or vProvince <> "ON" Then Exit Function
    If cDate(vDate) < cDate("Jul 01, 2010") Then
      fPST = .08  
    End If        
  End Function

%>