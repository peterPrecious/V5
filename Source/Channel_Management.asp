<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Chan.asp"-->

<%
  '...get latest ecommerce sales and update the channel table

  sOpenDb
  sOpenDb2
  sOpenDb3


  vSql = "SELECT " _
       & "  LEFT(Ecom_CustId, 4) AS [Chan_Id], " _

       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2004' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2004e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2005' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2005e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2006' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2006e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2007' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2007e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2008' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2008e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2009' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2009e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2010' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2010e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2011' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2011e],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2012' AND Ecom_Source = 'E' THEN Ecom_Prices ELSE 0 END) AS [2012e],  " _ 

       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2004' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2004m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2005' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2005m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2006' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2006m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2007' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2007m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2008' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2008m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2009' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2009m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2010' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2010m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2011' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2011m],  " _ 
       & "    SUM(CASE WHEN YEAR(Ecom_Issued) = '2012' AND Ecom_Source <> 'E' THEN Ecom_Prices ELSE 0 END) AS [2012m]   " _ 

       & "  FROM  Ecom " _ 
       & "  WHERE Ecom_Prices <> 0 " _
       & "  GROUP BY LEFT(Ecom.Ecom_CustId, 4) " _
       & "  ORDER BY LEFT(Ecom_CustId, 4) " 

' sDebug
  Set oRs2 = oDb2.Execute(vSql)
  Do While Not oRs2.Eof    

    vChan_Id      =  oRs2("Chan_Id")

    '...see if this channel is already on file
    sGetChan

    '...get values to insert/update into table
    vChan_2004e   =  oRs2("2004e")
    vChan_2005e   =  oRs2("2005e")
    vChan_2006e   =  oRs2("2006e")
    vChan_2007e   =  oRs2("2007e")
    vChan_2008e   =  oRs2("2008e")
    vChan_2009e   =  oRs2("2009e")
    vChan_2010e   =  oRs2("2010e")
    vChan_2011e   =  oRs2("2011e")
    vChan_2012e   =  oRs2("2012e")

    vChan_2004m   =  oRs2("2004m")
    vChan_2005m   =  oRs2("2005m")
    vChan_2006m   =  oRs2("2006m")
    vChan_2007m   =  oRs2("2007m")
    vChan_2008m   =  oRs2("2008m")
    vChan_2009m   =  oRs2("2009m")
    vChan_2010m   =  oRs2("2010m")
    vChan_2011m   =  oRs2("2011m")
    vChan_2012m   =  oRs2("2012m")


    '...get channel title
    vSql = "SELECT Cust_Title FROM ChannelTitle('" & vChan_Id & "') ChannelTitle"
    Set oRs3 = oDb3.Execute(vSql)
    If Not oRs3.Eof Then 
      vChan_Title = oRs3("Cust_Title")
    Else
      vChan_Title = "Not available"
    End If


    '...add to channel table
    If bChan_Eof Then
      vSql = "INSERT INTO Chan "
      vSql = vSql & "(Chan_Id, Chan_Title, Chan_2004e, Chan_2005e, Chan_2006e, Chan_2007e, Chan_2008e, Chan_2009e, Chan_2010e, Chan_2011e, Chan_2012e, Chan_2004m, Chan_2005m, Chan_2006m, Chan_2007m, Chan_2008m, Chan_2009m, Chan_2010m, Chan_2011m, Chan_2012m)"
      vSql = vSql & " VALUES ('" & vChan_Id & "', '" & fUnQuote(vChan_Title) & "', " & vChan_2004e & ", " & vChan_2005e & ", " & vChan_2006e & ", " & vChan_2007e & ", " & vChan_2008e & ", " & vChan_2009e & ", " & vChan_2010e & ", " & vChan_2011e & ", " & vChan_2012e & ", " & vChan_2004m & ", " & vChan_2005m & ", " & vChan_2006m & ", " & vChan_2007m & ", " & vChan_2008m & ", " & vChan_2009m & ", " & vChan_2010m & ", " & vChan_2011m & ", " & vChan_2012m & ")"
  '   sDebug
      sOpenDb
      oDb.Execute(vSql)
    Else
      vSql = "UPDATE Chan SET"
      vSql = vSql & " Chan_2004e              =  " & vChan_2004e       & " , " 
      vSql = vSql & " Chan_2005e              =  " & vChan_2005e       & " , " 
      vSql = vSql & " Chan_2006e              =  " & vChan_2006e       & " , " 
      vSql = vSql & " Chan_2007e              =  " & vChan_2007e       & " , " 
      vSql = vSql & " Chan_2008e              =  " & vChan_2008e       & " , " 
      vSql = vSql & " Chan_2009e              =  " & vChan_2009e       & " , " 
      vSql = vSql & " Chan_2010e              =  " & vChan_2010e       & " , " 
      vSql = vSql & " Chan_2011e              =  " & vChan_2011e       & " , " 
      vSql = vSql & " Chan_2012e              =  " & vChan_2012e       & " , " 
      vSql = vSql & " Chan_2004m              =  " & vChan_2004m       & " , " 
      vSql = vSql & " Chan_2005m              =  " & vChan_2005m       & " , " 
      vSql = vSql & " Chan_2006m              =  " & vChan_2006m       & " , " 
      vSql = vSql & " Chan_2007m              =  " & vChan_2007m       & " , " 
      vSql = vSql & " Chan_2008m              =  " & vChan_2008m       & " , " 
      vSql = vSql & " Chan_2009m              =  " & vChan_2009m       & " , " 
      vSql = vSql & " Chan_2010m              =  " & vChan_2010m       & " , " 
      vSql = vSql & " Chan_2011m              =  " & vChan_2011m       & " , " 
      vSql = vSql & " Chan_2012m              =  " & vChan_2012m       & "   " 
      vSql = vSql & " WHERE Chan_Id          = '" & vChan_Id         & "' "
'     sDebug
      sOpenDb
      oDb.Execute(vSql)
    End If

    oRs2.MoveNext
  Loop
  Set oRs2 = Nothing
  sCloseDb2
  sCloseDb3
  
  Response.Redirect "Channel_Report.asp"

%> 