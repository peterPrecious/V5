<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 

  Server.ScriptTimeout = 60 * 60
  Dim vCustId, vAcctId 

  sOpenDb    

  vCustId = "CNPX2850" :  sDeleteCust
  vCustId = "BANK2332" :  sDeleteCust
  vCustId = "RSRT1543" :  sDeleteCust
  vCustId = "GVDA2465" :  sDeleteCust
  vCustId = "CGRH2750" :  sDeleteCust
  vCustId = "CMSS2852" :  sDeleteCust
  vCustId = "CFIB2985" :  sDeleteCust
  vCustId = "HOWE2071" :  sDeleteCust
  vCustId = "CSAN2978" :  sDeleteCust
  vCustId = "NAGP2292" :  sDeleteCust
  vCustId = "HRSS2328" :  sDeleteCust
  vCustId = "OHAA2407" :  sDeleteCust
  vCustId = "MBST2341" :  sDeleteCust
  vCustId = "OPPF2463" :  sDeleteCust
  vCustId = "SHAD2458" :  sDeleteCust
  vCustId = "VUBZ2567" :  sDeleteCust
  vCustId = "MNFN2554" :  sDeleteCust
  vCustId = "METR2416" :  sDeleteCust
  vCustId = "DSES2663" :  sDeleteCust
  vCustId = "VUBZ8117" :  sDeleteCust
  vCustId = "VUBZ2698" :  sDeleteCust
  vCustId = "MNFN2689" :  sDeleteCust
  vCustId = "MNFN2561" :  sDeleteCust
  vCustId = "MPEB2706" :  sDeleteCust
  vCustId = "CCPE2645" :  sDeleteCust
  vCustId = "VUBZ2614" :  sDeleteCust
  vCustId = "MTCU2626" :  sDeleteCust
  vCustId = "VUBZ8335" :  sDeleteCust
  vCustId = "OPGT2436" :  sDeleteCust
  vCustId = "VUBZ8135" :  sDeleteCust
  vCustId = "GVDA2333" :  sDeleteCust
  vCustId = "VUBZ8308" :  sDeleteCust
  vCustId = "VUBZ8322" :  sDeleteCust
  vCustId = "VUBZ8359" :  sDeleteCust
  vCustId = "HRFR2327" :  sDeleteCust
  vCustId = "HRMG2326" :  sDeleteCust
  vCustId = "VUBZ8282" :  sDeleteCust
  vCustId = "SALM2596" :  sDeleteCust
  vCustId = "NPSV2738" :  sDeleteCust
  vCustId = "LASN2477" :  sDeleteCust
  vCustId = "LTSN2516" :  sDeleteCust
  vCustId = "JAVA2789" :  sDeleteCust
  vCustId = "VUBZ2734" :  sDeleteCust
  vCustId = "RGMG2730" :  sDeleteCust
  vCustId = "MCMK2833" :  sDeleteCust
  vCustId = "MCMK2834" :  sDeleteCust
  vCustId = "LABA2810" :  sDeleteCust
  vCustId = "VUBZ8086" :  sDeleteCust
  vCustId = "NWFD2744" :  sDeleteCust
  vCustId = "THOM2964" :  sDeleteCust
  vCustId = "ERGP2914" :  sDeleteCust
  vCustId = "MNCH2386" :  sDeleteCust
  vCustId = "ERGP2897" :  sDeleteCust
  vCustId = "ERGP0035" :  sDeleteCust
  vCustId = "ERGP2901" :  sDeleteCust
  vCustId = "ERGP2925" :  sDeleteCust
  vCustId = "ERGP0012" :  sDeleteCust
  vCustId = "ERGP0037" :  sDeleteCust
  vCustId = "ERGP0050" :  sDeleteCust
  vCustId = "ERGP0099" :  sDeleteCust
  vCustId = "ERGP0060" :  sDeleteCust
  vCustId = "ERGP2929" :  sDeleteCust
  vCustId = "ERGP0061" :  sDeleteCust
  vCustId = "ERGP2926" :  sDeleteCust
  vCustId = "ERGP0097" :  sDeleteCust
  vCustId = "ERGP0142" :  sDeleteCust
  vCustId = "ERGP0136" :  sDeleteCust
  vCustId = "ERGP0027" :  sDeleteCust
  vCustId = "ERGP0112" :  sDeleteCust
  vCustId = "ERGP0069" :  sDeleteCust
  vCustId = "MCYS2960" :  sDeleteCust
  vCustId = "ERGP2895" :  sDeleteCust
  vCustId = "ERGP0113" :  sDeleteCust
  vCustId = "ERGP0250" :  sDeleteCust
  vCustId = "PRFC2975" :  sDeleteCust
  vCustId = "ERGP2904" :  sDeleteCust
  vCustId = "VUBZ2984" :  sDeleteCust
  vCustId = "ERGP2909" :  sDeleteCust
  vCustId = "ERGP2912" :  sDeleteCust
  vCustId = "ERGP2923" :  sDeleteCust
  vCustId = "ERGP0289" :  sDeleteCust
  vCustId = "ERGP0124" :  sDeleteCust
  vCustId = "ERGP2905" :  sDeleteCust
  vCustId = "ERGP0075" :  sDeleteCust
  vCustId = "ERGP0290" :  sDeleteCust
  vCustId = "ERGP0297" :  sDeleteCust
  vCustId = "ERGP0026" :  sDeleteCust
  vCustId = "ERGP0288" :  sDeleteCust
  vCustId = "ERGP0317" :  sDeleteCust
  vCustId = "ERGP0276" :  sDeleteCust
  vCustId = "ERGP0219" :  sDeleteCust
  vCustId = "ERGP2921" :  sDeleteCust
  vCustId = "ERGP2933" :  sDeleteCust
  vCustId = "ERGP0307" :  sDeleteCust
  vCustId = "ERGP0004" :  sDeleteCust
  vCustId = "ERGP0359" :  sDeleteCust
  vCustId = "ERGP0349" :  sDeleteCust
  vCustId = "ERGP0089" :  sDeleteCust
  vCustId = "ERGP0086" :  sDeleteCust
  vCustId = "ERGP0260" :  sDeleteCust
  vCustId = "ERGP0401" :  sDeleteCust
  vCustId = "ERGP2916" :  sDeleteCust
  vCustId = "ERGP0122" :  sDeleteCust
  vCustId = "ERGP2931" :  sDeleteCust
  vCustId = "ERGP0296" :  sDeleteCust
  vCustId = "ERGP0132" :  sDeleteCust
  vCustId = "ERGP0319" :  sDeleteCust
  vCustId = "ERGP0404" :  sDeleteCust
  vCustId = "NAYL0663" :  sDeleteCust
  vCustId = "ERGP0316" :  sDeleteCust
  vCustId = "ERGP0423" :  sDeleteCust
  vCustId = "ERGP0005" :  sDeleteCust
  vCustId = "ERGP0028" :  sDeleteCust
  vCustId = "ERGP0364" :  sDeleteCust
  vCustId = "ERGP0477" :  sDeleteCust
  vCustId = "ERGP2937" :  sDeleteCust
  vCustId = "ERGP0476" :  sDeleteCust
  vCustId = "ERGP0277" :  sDeleteCust
  vCustId = "TELE2289" :  sDeleteCust
  vCustId = "ERGP2922" :  sDeleteCust
  vCustId = "ERGP0098" :  sDeleteCust
  vCustId = "MYTR2829" :  sDeleteCust
  vCustId = "ERGP0034" :  sDeleteCust
  vCustId = "ERGP0365" :  sDeleteCust
  vCustId = "ERGP0199" :  sDeleteCust
  vCustId = "ERGP0110" :  sDeleteCust
  vCustId = "INDG2886" :  sDeleteCust
  vCustId = "OCGW2882" :  sDeleteCust
  vCustId = "ERGP0185" :  sDeleteCust
  vCustId = "COOK2770" :  sDeleteCust
  vCustId = "ERGP0300" :  sDeleteCust
  vCustId = "ERGP2907" :  sDeleteCust
  vCustId = "ERGP2913" :  sDeleteCust
  vCustId = "ERGP0371" :  sDeleteCust
  vCustId = "ERGP0078" :  sDeleteCust
  vCustId = "ERGP0254" :  sDeleteCust
  vCustId = "ERGP0278" :  sDeleteCust
  vCustId = "ERGP2932" :  sDeleteCust
  vCustId = "BANK2331" :  sDeleteCust
  vCustId = "ERGP0339" :  sDeleteCust
  vCustId = "ERGP0596" :  sDeleteCust
  vCustId = "ERGP0007" :  sDeleteCust
  vCustId = "ERGP0624" :  sDeleteCust
  vCustId = "ERGP2919" :  sDeleteCust
  vCustId = "ERGP0074" :  sDeleteCust
  vCustId = "ERGP0013" :  sDeleteCust
  vCustId = "ERGP2927" :  sDeleteCust
  vCustId = "ERGP0157" :  sDeleteCust
  vCustId = "ERGP0251" :  sDeleteCust
  vCustId = "ERGP0189" :  sDeleteCust
  vCustId = "ERGP2911" :  sDeleteCust
  vCustId = "ERGP0140" :  sDeleteCust
  vCustId = "ERGP0048" :  sDeleteCust
  vCustId = "FTBG2855" :  sDeleteCust
  vCustId = "ERGP2906" :  sDeleteCust
  vCustId = "ERGP2918" :  sDeleteCust
  vCustId = "ERGP0361" :  sDeleteCust
  vCustId = "FARM0954" :  sDeleteCust
  vCustId = "ERGP2903" :  sDeleteCust
  vCustId = "ERGP2915" :  sDeleteCust
  vCustId = "ERGP0489" :  sDeleteCust
  vCustId = "EDCC2725" :  sDeleteCust
  vCustId = "ERGP2898" :  sDeleteCust
  vCustId = "ERGP2920" :  sDeleteCust
  vCustId = "MACB2642" :  sDeleteCust
  vCustId = "ERGP2893" :  sDeleteCust
  vCustId = "ERGP0350" :  sDeleteCust
  vCustId = "STEW2680" :  sDeleteCust
  vCustId = "ERGP0200" :  sDeleteCust
  vCustId = "ERGP0076" :  sDeleteCust
  vCustId = "ERGP0077" :  sDeleteCust
  vCustId = "ERGP0745" :  sDeleteCust
  vCustId = "ERGP0071" :  sDeleteCust
  vCustId = "ERGP0152" :  sDeleteCust
  vCustId = "ERGP0188" :  sDeleteCust
  vCustId = "ERGP0133" :  sDeleteCust
  vCustId = "ERGP0218" :  sDeleteCust
  vCustId = "ERGP2894" :  sDeleteCust
  vCustId = "ERGP2934" :  sDeleteCust
  vCustId = "ERGP0448" :  sDeleteCust
  vCustId = "ERGP0249" :  sDeleteCust
  vCustId = "ERGP0318" :  sDeleteCust
  vCustId = "ERGP0024" :  sDeleteCust
  vCustId = "ERGP2910" :  sDeleteCust
  vCustId = "ERGP0435" :  sDeleteCust
  vCustId = "STEW2625" :  sDeleteCust
  vCustId = "ERGP0266" :  sDeleteCust
  vCustId = "ERGP0360" :  sDeleteCust
  vCustId = "ERGP0023" :  sDeleteCust
  vCustId = "ERGP2924" :  sDeleteCust
  vCustId = "CORP2837" :  sDeleteCust
  vCustId = "VUBZ2733" :  sDeleteCust
  vCustId = "SCDC2979" :  sDeleteCust
  vCustId = "SCDC2981" :  sDeleteCust
  vCustId = "NFIB2813" :  sDeleteCust
  vCustId = "ERGP2917" :  sDeleteCust
  vCustId = "ERGP2928" :  sDeleteCust
  vCustId = "CNAP3006" :  sDeleteCust
  vCustId = "ERGP0703" :  sDeleteCust
  vCustId = "ERGP2908" :  sDeleteCust
  vCustId = "VUBZ2343" :  sDeleteCust
  vCustId = "ERGP0704" :  sDeleteCust
  vCustId = "ERGP0032" :  sDeleteCust
  vCustId = "ERGP0115" :  sDeleteCust
  vCustId = "ERGP0120" :  sDeleteCust
  vCustId = "ERGP2900" :  sDeleteCust

  sCloseDb





  Sub sDeleteCust 

    vAcctId = Right(vCustId, 4)

    oDb.Execute("DELETE FROM Cust WHERE Cust_AcctId     = '" & vAcctId & "'")
    oDb.Execute("DELETE FROM Logs WHERE Logs_AcctId     = '" & vAcctId & "'")
    oDb.Execute("DELETE FROM Memb WHERE Memb_AcctId     = '" & vAcctId & "'")
    oDb.Execute("DELETE FROM Catl WHERE Catl_CustId     = '" & vCustId & "'")

    '...delete Task assets
    vSql  = "SELECT TskH_No FROM TskH WHERE TskH_AcctId  = '" & vAcctId & "'"
    Set oRs  = oDb.Execute(vSql)
    Do While Not oRs.Eof 
      vSql  = "DELETE FROM TskD WHERE TskD_No  = " & oRs("TskH_No")
      oDb.Execute(vSql)
      oRs.MoveNext
    Loop
    Set oRs  = Nothing 
    oDb.Execute("DELETE FROM TskH WHERE TskH_AcctId = '" & vAcctId & "'")

    Response.Write vCustId & "<br>"

  End Sub  

%>

