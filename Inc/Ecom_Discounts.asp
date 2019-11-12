<%
    '...Discount Routines

    Sub sCheckDiscounts
    
      '...If CDs, check for quantity discount
      If vEcom_Media = "CDs" Then
        i = 0
        For j = 1 to svProdMax
          '...ensure program is eligible for discounts
          i = i + saProd(2, j)
        Next        
        
        If i < 10 Then
          sQuantityDiscount (0)
        ElseIf i >= 10 And i < 25 Then
          sQuantityDiscount (10)
        ElseIf i >= 25 And i < 50 Then
          sQuantityDiscount (15)
        ElseIf i >= 50 And i < 250 Then 
          sQuantityDiscount (20)
        ElseIf i >= 250 Then   
          sQuantityDiscount (30)
        End If

        If i > 1 Then    
          Session("Prod") = saProd
        End If


      '...If Group2 aor AddOn2, check for quantity discount from the customer record
      
      
'     ElseIf vEcom_Media = "Group2" Then
      ElseIf vEcom_Media = "Group2" Or vEcom_Media = "AddOn2" Then

        Dim vGroup2Rates, aGroup2Rates, aGroup2Rate1, aGroup2Rate2, aGroup2Rate3, aGroup2Rate4, aGroup2Rate5 
        vGroup2Rates = fDefault(vCust_EcomGroup2Rates, "5|25~10|45~25|55~50|65~200|75")
        aGroup2Rates = Split(vGroup2Rates, "~")
        aGroup2Rate1 = Split (aGroup2Rates(0), "|")
        aGroup2Rate2 = Split (aGroup2Rates(1), "|")
        aGroup2Rate3 = Split (aGroup2Rates(2), "|")
        aGroup2Rate4 = Split (aGroup2Rates(3), "|")
        aGroup2Rate5 = Split (aGroup2Rates(4), "|")

      
        '...compute quantity overall, but not for "freebies" or programs that are not eligible for discounts
        i = 0
        For j = 1 to svProdMax
          '...ensure program is eligible for discounts
          If (saProd(3, j) > 0 Or saProd(4, j) > 0) And fProgDiscountsOk (saProd(1, j)) Then 
            i = i + saProd(2, j)
          End If
        Next        


        If i < Cint(aGroup2Rate1(0)) Then
          sQuantityDiscount (0)
        ElseIf i >= Cint(aGroup2Rate1(0)) And i < Cint(aGroup2Rate2(0)) Then
          sQuantityDiscount (Cint(aGroup2Rate1(1)))
        ElseIf i >= Cint(aGroup2Rate2(0)) And i < Cint(aGroup2Rate3(0)) Then
          sQuantityDiscount (Cint(aGroup2Rate2(1)))
        ElseIf i >= Cint(aGroup2Rate3(0)) And i < Cint(aGroup2Rate4(0)) Then
          sQuantityDiscount (Cint(aGroup2Rate3(1)))
        ElseIf i >= Cint(aGroup2Rate4(0)) And i < Cint(aGroup2Rate5(0)) Then
          sQuantityDiscount (Cint(aGroup2Rate4(1)))
        ElseIf i >= Cint(aGroup2Rate5(0)) Then
          sQuantityDiscount (Cint(aGroup2Rate5(1)))
        End If

        If i > 1 Then    
          Session("Prod") = saProd
        End If


      Else
  
        '...there are 4 ways to compute customer discounts
        If vCust_EcomDiscOptions > 0 Or Len(Session("Ecom_AdditionalDiscount")) > 0 Then 

          '... apply blanket discount only
          If vCust_EcomDiscOptions = "1" Or vCust_EcomDiscOptions = "4" Then
            fBlanketDiscount            
          End If
        
          '... apply blanket discount if no repurchase only
          If vCust_EcomDiscOptions = "3" Then
            If Not fRepurDiscount Then
              fBlanketDiscount            
            End If  
          End If
        
          '... apply repurchase discount only
          If vCust_EcomDiscOptions = "2" Or vCust_EcomDiscOptions = "4" Then
            fRepurDiscount
          End If

          '...apply "additional discount"
          sQuantityDiscount (0)
          
          Session("Prod") = saProd

        End If

      End If  

    End Sub


    Sub sQuantityDiscount (vDiscount)
      '...get additional discount (this offers x% off the discounted amount)
      If Len(Session("Ecom_AdditionalDiscount")) > 0 Then
        If IsNumeric(Session("Ecom_AdditionalDiscount")) Then
          vDiscount = vDiscount * (1 + Cint(Session("Ecom_AdditionalDiscount"))/100)
          If vDiscount < 0   Then vDiscount = 0
          If vDiscount > 100 Then vDiscount = 100
        End If
      End If    
      '...add onto compute basic discount
      For i = 1 to svProdMax
        '...ensure program is eligible for discounts
        If fDiscProgramsOk (saProd(1, i)) And fProgDiscountsOk (saProd(1, i)) Then
          If vEcom_Media = "Group2" Or vEcom_Media = "AddOn2" Then
            saProd(0, i) = vDiscount
          Else            
            saProd(0, i) = saProd(0, i) + vDiscount
          End If
        End If
      Next
    End Sub

    

    Function fBlanketDiscount
      fBlanketDiscount = False
      '...if a discount is defined
      If vCust_EcomDisc > 0 Then 
        vOk = True
        '...if a discount requires a minimum US purchase
        If vOk And vCust_EcomDiscMinUS > 0 Then
          If vCust_EcomDiscMinUS > vTotUS Then vOk = False
        End If
        '...if a discount requires a minimum CA purchase
        If vOk And vCust_EcomDiscMinCA > 0 Then
          If vCust_EcomDiscMinCA > vTotCA Then vOk = False
        End If
        '...if a discount requires a minimum qty purchase
        If vOk And vCust_EcomDiscMinQty > 0 Then
          If vCust_EcomDiscMinQty > vTotQty Then vOk = False
        End If
        '...if a discount doesn't apply to the original quantity
        If vOk And vCust_EcomDiscOriginal Then
          vStr = 1
        Else
          vStr = vCust_EcomDiscMinQty + 1
        End If
        '...if a discount is limited in quantity
        If vOk And vCust_EcomDiscLimit > 0 Then
          vEnd = vCust_EcomDiscMinQty + vCust_EcomDiscLimit
          If vEnd > svProdMax Then vEnd = svProdMax
        Else
          vEnd = svProdMax        
        End If

        '...store the discount percentage (ie 25%) in 0
        If vOk Then
          For i = vStr To vEnd
            fBlanketDiscount = False
            '...ensure program is eligible for discounts
            If fDiscProgramsOk (saProd(1, i)) Then
              saProd(0, i) = vCust_EcomDisc
            End If
          Next
        End If
      End If      
    End Function
    

    Function fRepurDiscount
      fRepurDiscount = False
      If vCust_EcomRepurDisc > 0 Then
        '...check if previously repurchase the program during the past x days
        For i = 1 to svProdMax
          '...ensure program is eligible for discounts
          If fDiscProgramsOk (saProd(1, i)) Then
            If fRepurchased(saProd(1, i)) Then 
              fRepurDiscount = True
              saProd(0, i) = vCust_EcomRepurDisc
            End If
          End If
        Next
      End If
    End Function
    

    '...are there restrictions on which programs are eligible for discounts?
    Function fDiscProgramsOk (vProg)
      fDiscProgramsOk = True
      '...any restrictions
      If Len(vCust_EcomDiscPrograms) > 0 Then
        '...is this program eligible
        If Instr(vCust_EcomDiscPrograms, vProg) = 0 Then  
          fDiscProgramsOk = False
        End If
      End If
    End Function
    
%>