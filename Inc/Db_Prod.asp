<%
  Dim vProd_Id, vProd_CatTitle, vProd_Title, vProd_Desc, vProd_Price
  Dim vProd_Eof, vProd_Ok


  '...get left side
  Sub sGetProdLeft_Rs (vProdCustId)
'   vSql = " SELECT DISTINCT Left(Prod_Id, 3) AS ProdId, Prod_CatTitle FROM Prod WHERE Left(Prod_Id, 8) = '" & vProdCustId & "'"
    vSql = " SELECT DISTINCT Left(Prod_Id, 12) AS ProdId, Prod_CatTitle FROM Prod WHERE Left(Prod_Id, 8) = '" & vProdCustId & "'"
'   vSql = " SELECT DISTINCT SUBSTRING(Prod_Id, 10, 3) AS ProdId, Prod_CatTitle FROM Prod WHERE Left(Prod_Id, 8) = '" & vProdCustId & "'"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
  End Sub


  '...get right side
  Sub sGetProdRight_Rs (vProdId)
    vProd_Eof = True
'   vSql =  " SELECT * FROM Prod WHERE (Left(Prod.Prod_Id, 3) = '" & vProdId & "')"
    vSql =  " SELECT * FROM Prod WHERE (Left(Prod.Prod_Id, 12) = '" & vProdId & "')"
'   vSql =  " SELECT * FROM Prod WHERE SUBSTRING(Prod_Id, 10, 3) = '" & vProdId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
  End Sub


  '...Get Prod Record 
  Sub sGetProd (vProdId)
    vProd_Eof = True
    vSql =  " SELECT * FROM Prod WHERE Prod.Prod_Id = '" & vProdId & "'"
    sOpenDb    
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      sReadProd
      vProd_Eof = False
    End If
    Set oRs = Nothing
    sCloseDb
  End Sub


  Sub sReadProd
    vProd_Id              = oRs("Prod_Id")
    vProd_CatTitle        = oRs("Prod_CatTitle")
    vProd_Title           = Trim(oRs("Prod_Title"))
    vProd_Desc            = Trim(oRs("Prod_Desc"))
    vProd_Price           = oRs("Prod_Price")
  End Sub


  '...Get Prod Title
  Function fProdTitle (vProdId)
    Dim oRs2
    fProdTitle = ""
    vSql = "SELECT Prod_Title FROM Prod WHERE Prod_Id= '" & vProdId & "'"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      fProdTitle = oRs2("Prod_Title")
    End If
    Set oRs2 = Nothing
    sCloseDb2    
  End Function


  '...Get Prod Title
  Function fProdSpecials ()
    Dim oRs2
    fProdSpecials = False
    vSql = "SELECT Top 1 Prod_Id FROM Prod WHERE Left(Prod_Id, 8) = '" & svCustId & "'"
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      fProdSpecials = True
    End If
    Set oRs2 = Nothing
    sCloseDb2    
  End Function


%>