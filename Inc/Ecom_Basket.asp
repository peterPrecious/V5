<%
  '...current basket pointer (Prod_no) And no of items in basket (Prod_max)
  Dim svProd_no, svProd_max

  '...setup order basket array info from Session
  svProd_no      = Session("Prod_no")
  If svProd_no   = "" Then svProd_no = 0  
  svProd_max     = Session("Prod_max")
  If svProd_max  = "" Then svProd_max = 0
  If svProd_no   > 0 Then  
    saProds      = Session("Prods")
  Else
    svProd_max   = 0
    Dim saProds()
  End If

  '...since there is only one pass at filling the basket, then set to zero
  svProd_No      = 0
  svProd_max     = 0
  

  '...put product id And qty into array (If it exists, Else add To array)
  Sub sStoreProds (ID, QTY, DESC, PriceUS, PriceCA)
    Dim i

'   sDebug "ID", ID
'   sDebug "Qty", Qty
'   sDebug "Dsc", Desc
'   sDebug "Prc", PriceCA
'   sDebug "Prod_max", svProd_max

    '...1st product in array?
    If svProd_max = 0 Then
      svProd_no = 1
      svProd_max = 1
      ReDim saProds (5, 1)
      saProds(1, 1) = ID
      saProds(2, 1) = Clng(QTY)
      saProds(3, 1) = DESC
      saProds(4, 1) = PriceUS
      saProds(5, 1) = PriceCA
      Session("Prod_no") = svProd_no
      Session("Prod_max") = svProd_max
      Session("Prods") = saProds
      Exit sub
    End If

    '...check If already there
    For i = 1 To svProd_max
      If ID = saProds (1, i) Then
        If saProds(2, i) > 0 And QTY = 0 Then svProd_no = svProd_no - 1
        If saProds(2, i) = 0 And QTY > 0 Then svProd_no = svProd_no + 1
        saProds(2, i) = clng(QTY)
        saProds(4, i) = PriceUS
        saProds(5, i) = PriceCA
        Session("Prod_no") = svProd_no
        Session("Prods") = saProds
        Exit sub      
      End If
    Next

    '...Else add To the End of the array
    svProd_no = svProd_no + 1
    svProd_max = svProd_max + 1

'   sDebug "svProd_no", svProd_no
'   sDebug "svProd_max", svProd_max

    ReDim Preserve saProds (4, svProd_no)
    saProds(1, svProd_no) = ID
    saProds(2, svProd_no) = clng(QTY)
    saProds(3, svProd_no) = DESC
    saProds(4, svProd_no) = PriceUS
    saProds(5, svProd_no) = PriceCA
    Session("Prod_no")    = svProd_no
    Session("Prod_max")   = svProd_max
    Session("Prods")      = saProds
  End sub

  '...find product in vproduct array
  Function fGetProdsQty (vProd)
    Dim i, j, k
    For i = 1 To svProd_max
      If saProds(1, i) = vProd Then 
        fGetProdsQty = saProds(2, i)
        Exit For
      End If
    Next
  End function

%>