<%
  Option Explicit

  Dim i, j, k, l, vBypassSecurity 
  Dim vCurrency, vClose, vGstRate, vHstRate, vPstRate

  '...chance site wide access ids
  Dim vPassword2, vPassword3, vPassword4, vPassword5

  vPassword2 = "VUV5_LRN"
  vPassword3 = "VUV5_FAC"
  vPassword4 = "VUV5_MGR"
  vPassword5 = "VUV5_ADM"

  '...replace these next 3 lines but only after cleansed from all source
  Dim vPasswordx, vPassworda
  vPassworda = "VUV5_AUT" '...big authoring, like vPassword5 - assigned to sites when authoring accounts are not setup
  vPasswordx = "VUV5_"    '...this should no longer be needed as we now have "memb_internal" to flag users that shouldn't appear on reports or web services

  vBypassSecurity = False

' Response.Buffer = False 
  Server.ScriptTimeout = 60 * 10
 
  vCurrency       = 1.00 '...CA value relative to US (used on Customer Program Strings for Channels), ie if $CA is worth 90c US then enter 0.90
  '...as taxes have grown in complexity, use "Ecom_Routines.asp" for calculations
  vPstRate        =  .08 '...for old product sales (CDs, books, etc) - no longer used
  vPstRate        =  .00

  If Now() < cDate("Jan 01, 2008") Then
    vGstRate      =  .06
    vHstRate      =  .14
  ElseIf Now() < cDate("Jul 01, 2010") Then
    vGstRate      =  .05 
    vHstRate      =  .13
  Else
    vGstRate      =  .00 
    vHstRate      =  .13
  End If

  vClose          = ""      '...if set to "Y" after this page then the top frame of a "tab-less" window will appear 
  
%>