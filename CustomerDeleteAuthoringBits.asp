<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 

  Server.ScriptTimeout = 60 * 60
  Dim vCustId, vAcctId 
   

  sOpenDb    

  vCustId = "ACCC7026" :  sDeleteCust
  vCustId = "ACCI7384" :  sDeleteCust
  vCustId = "ADAY7053" :  sDeleteCust
  vCustId = "ADAY7081" :  sDeleteCust
  vCustId = "ADAY7173" :  sDeleteCust
  vCustId = "ALEI7243" :  sDeleteCust
  vCustId = "AMER7443" :  sDeleteCust
  vCustId = "AMER9815" :  sDeleteCust
  vCustId = "AODA7414" :  sDeleteCust
  vCustId = "ARCH7064" :  sDeleteCust
  vCustId = "ARCT7258" :  sDeleteCust
  vCustId = "ARCT7260" :  sDeleteCust
  vCustId = "AVEN7075" :  sDeleteCust
  vCustId = "BANK7049" :  sDeleteCust
  vCustId = "BANK7054" :  sDeleteCust
  vCustId = "BANK7370" :  sDeleteCust
  vCustId = "BANK7380" :  sDeleteCust
  vCustId = "BANK7397" :  sDeleteCust
  vCustId = "BANK8932" :  sDeleteCust
  vCustId = "BKSS7108" :  sDeleteCust
  vCustId = "BKSS7109" :  sDeleteCust
  vCustId = "BKSS7110" :  sDeleteCust
  vCustId = "BKSS7111" :  sDeleteCust
  vCustId = "BKSS7112" :  sDeleteCust
  vCustId = "BKSS7113" :  sDeleteCust
  vCustId = "BKSS7114" :  sDeleteCust
  vCustId = "BKSS7115" :  sDeleteCust
  vCustId = "BKSS7116" :  sDeleteCust
  vCustId = "BKSS7117" :  sDeleteCust
  vCustId = "BKSS7118" :  sDeleteCust
  vCustId = "BKSS7119" :  sDeleteCust
  vCustId = "BKSS7120" :  sDeleteCust
  vCustId = "BKSS7176" :  sDeleteCust
  vCustId = "BLUE7174" :  sDeleteCust
  vCustId = "BMBD7139" :  sDeleteCust
  vCustId = "BNAI7324" :  sDeleteCust
  vCustId = "BNAI7355" :  sDeleteCust
  vCustId = "BNAI7364" :  sDeleteCust
  vCustId = "BZLN7030" :  sDeleteCust
  vCustId = "CAAM7356" :  sDeleteCust
  vCustId = "CAAM7357" :  sDeleteCust
  vCustId = "CAAM7358" :  sDeleteCust
  vCustId = "CAAM7376" :  sDeleteCust
  vCustId = "CAAM7388" :  sDeleteCust
  vCustId = "CAAM7413" :  sDeleteCust
  vCustId = "CAAM7419" :  sDeleteCust
  vCustId = "CAAM7420" :  sDeleteCust
  vCustId = "CAAM7468" :  sDeleteCust
  vCustId = "CADI9855" :  sDeleteCust
  vCustId = "CAIS7320" :  sDeleteCust
  vCustId = "CANO7394" :  sDeleteCust
  vCustId = "CAST7387" :  sDeleteCust
  vCustId = "CBKS7041" :  sDeleteCust
  vCustId = "CBRE9383" :  sDeleteCust
  vCustId = "CBVC7378" :  sDeleteCust
  vCustId = "CBVC7417" :  sDeleteCust
  vCustId = "CCHS7034" :  sDeleteCust
  vCustId = "CCHS7043" :  sDeleteCust
  vCustId = "CCHS7046" :  sDeleteCust
  vCustId = "CCHS7063" :  sDeleteCust
  vCustId = "CCHS7095" :  sDeleteCust
  vCustId = "CCHS7096" :  sDeleteCust
  vCustId = "CCHS7097" :  sDeleteCust
  vCustId = "CCHS7163" :  sDeleteCust
  vCustId = "CCHS7164" :  sDeleteCust
  vCustId = "CCHS7167" :  sDeleteCust
  vCustId = "CCHS7226" :  sDeleteCust
  vCustId = "CCHS7236" :  sDeleteCust
  vCustId = "CCHS7237" :  sDeleteCust
  vCustId = "CCHS7238" :  sDeleteCust
  vCustId = "CCHS7252" :  sDeleteCust
  vCustId = "CCHS7298" :  sDeleteCust
  vCustId = "CCHS7347" :  sDeleteCust
  vCustId = "CCHS7390" :  sDeleteCust
  vCustId = "CCHS7395" :  sDeleteCust
  vCustId = "CCPE7073" :  sDeleteCust
  vCustId = "CFIB7346" :  sDeleteCust
  vCustId = "CFIB7422" :  sDeleteCust
  vCustId = "CHIN7259" :  sDeleteCust
  vCustId = "CIAC7418" :  sDeleteCust
  vCustId = "CICA7017" :  sDeleteCust
  vCustId = "CICA7144" :  sDeleteCust
  vCustId = "CICA7160" :  sDeleteCust
  vCustId = "CINE7214" :  sDeleteCust
  vCustId = "CINE7215" :  sDeleteCust
  vCustId = "CINE7216" :  sDeleteCust
  vCustId = "CINE7217" :  sDeleteCust
  vCustId = "CINE7218" :  sDeleteCust
  vCustId = "CINE7247" :  sDeleteCust
  vCustId = "CINE7336" :  sDeleteCust
  vCustId = "CINE7337" :  sDeleteCust
  vCustId = "CINE7352" :  sDeleteCust
  vCustId = "CINE7353" :  sDeleteCust
  vCustId = "CLSV7020" :  sDeleteCust
  vCustId = "CMSS7047" :  sDeleteCust
  vCustId = "CNAP7350" :  sDeleteCust
  vCustId = "COMP7444" :  sDeleteCust
  vCustId = "CPPI7151" :  sDeleteCust
  vCustId = "CPST7162" :  sDeleteCust
  vCustId = "CSHM7399" :  sDeleteCust
  vCustId = "CSSE7281" :  sDeleteCust
  vCustId = "DACO7465" :  sDeleteCust
  vCustId = "DBRM7061" :  sDeleteCust
  vCustId = "DNBI7141" :  sDeleteCust
  vCustId = "DRMC7449" :  sDeleteCust
  vCustId = "ELAP7003" :  sDeleteCust
  vCustId = "ELAP7191" :  sDeleteCust
  vCustId = "EPKG7377" :  sDeleteCust
  vCustId = "ERGP7277" :  sDeleteCust
  vCustId = "ERGP7279" :  sDeleteCust
  vCustId = "ERGP7286" :  sDeleteCust
  vCustId = "ERGP7287" :  sDeleteCust
  vCustId = "ERGP7294" :  sDeleteCust
  vCustId = "ERGP7295" :  sDeleteCust
  vCustId = "ERGP7296" :  sDeleteCust
  vCustId = "ERGP7297" :  sDeleteCust
  vCustId = "ERGP7316" :  sDeleteCust
  vCustId = "ERGP7317" :  sDeleteCust
  vCustId = "ERGP7327" :  sDeleteCust
  vCustId = "ERGP7400" :  sDeleteCust
  vCustId = "FLUP7082" :  sDeleteCust
  vCustId = "FTTC7058" :  sDeleteCust
  vCustId = "FTTC7212" :  sDeleteCust
  vCustId = "FTTC7254" :  sDeleteCust
  vCustId = "FTTC7262" :  sDeleteCust
  vCustId = "FTTC7263" :  sDeleteCust
  vCustId = "GCPW7169" :  sDeleteCust
  vCustId = "GLBM7018" :  sDeleteCust
  vCustId = "GLOB7365" :  sDeleteCust
  vCustId = "GONL7127" :  sDeleteCust
  vCustId = "HCCM7040" :  sDeleteCust
  vCustId = "HECS7048" :  sDeleteCust
  vCustId = "HECS7080" :  sDeleteCust
  vCustId = "HECS7084" :  sDeleteCust
  vCustId = "HIGH7360" :  sDeleteCust
  vCustId = "HMJF9831" :  sDeleteCust
  vCustId = "HMVG7272" :  sDeleteCust
  vCustId = "HMVG7273" :  sDeleteCust
  vCustId = "IBAO7068" :  sDeleteCust
  vCustId = "IBAO7154" :  sDeleteCust
  vCustId = "IBAO7362" :  sDeleteCust
  vCustId = "IBAO7453" :  sDeleteCust
  vCustId = "IBBC7088" :  sDeleteCust
  vCustId = "IBEW7251" :  sDeleteCust
  vCustId = "ICBA7426" :  sDeleteCust
  vCustId = "ICBA7432" :  sDeleteCust
  vCustId = "ICBA9211" :  sDeleteCust
  vCustId = "ICSA7145" :  sDeleteCust
  vCustId = "INDG7292" :  sDeleteCust
  vCustId = "INDG7303" :  sDeleteCust
  vCustId = "INDG7304" :  sDeleteCust
  vCustId = "INDG7305" :  sDeleteCust
  vCustId = "INDG7311" :  sDeleteCust
  vCustId = "INDG7321" :  sDeleteCust
  vCustId = "INDG7323" :  sDeleteCust
  vCustId = "INDG7456" :  sDeleteCust
  vCustId = "INDG7457" :  sDeleteCust
  vCustId = "INDG7458" :  sDeleteCust
  vCustId = "INDG7459" :  sDeleteCust
  vCustId = "INDG7472" :  sDeleteCust
  vCustId = "INFO7065" :  sDeleteCust
  vCustId = "JATF7374" :  sDeleteCust
  vCustId = "JUAL7285" :  sDeleteCust
  vCustId = "KMSI3636" :  sDeleteCust
  vCustId = "LEAP7392" :  sDeleteCust
  vCustId = "LEON7319" :  sDeleteCust
  vCustId = "LGEN7280" :  sDeleteCust
  vCustId = "LLRN7085" :  sDeleteCust
  vCustId = "LNST7010" :  sDeleteCust
  vCustId = "MACA7379" :  sDeleteCust
  vCustId = "MAST7405" :  sDeleteCust
  vCustId = "MCSS7161" :  sDeleteCust
  vCustId = "MCYS7274" :  sDeleteCust
  vCustId = "MCYS7275" :  sDeleteCust
  vCustId = "MCYS7276" :  sDeleteCust
  vCustId = "MCZN7016" :  sDeleteCust
  vCustId = "MCZN7050" :  sDeleteCust
  vCustId = "MIND7396" :  sDeleteCust
  vCustId = "MMAH7158" :  sDeleteCust
  vCustId = "MMAH7159" :  sDeleteCust
  vCustId = "MMAH7166" :  sDeleteCust
  vCustId = "MMAH7221" :  sDeleteCust
  vCustId = "MMAH7222" :  sDeleteCust
  vCustId = "MMAH7225" :  sDeleteCust
  vCustId = "MMAH7256" :  sDeleteCust
  vCustId = "MMAH7315" :  sDeleteCust
  vCustId = "MMCI7290" :  sDeleteCust
  vCustId = "MMCI7291" :  sDeleteCust
  vCustId = "MNDC7019" :  sDeleteCust
  vCustId = "MNFN7028" :  sDeleteCust
  vCustId = "MNFN7052" :  sDeleteCust
  vCustId = "MOEN7172" :  sDeleteCust
  vCustId = "MOEN7306" :  sDeleteCust
  vCustId = "MOGS7124" :  sDeleteCust
  vCustId = "MOLB7123" :  sDeleteCust
  vCustId = "MOLB7219" :  sDeleteCust
  vCustId = "MPAC7439" :  sDeleteCust
  vCustId = "MPEB7134" :  sDeleteCust
  vCustId = "MPMS7152" :  sDeleteCust
  vCustId = "MTHR7155" :  sDeleteCust
  vCustId = "MTHR7329" :  sDeleteCust
  vCustId = "MTHR7330" :  sDeleteCust
  vCustId = "MTHR7331" :  sDeleteCust
  vCustId = "MTHR7334" :  sDeleteCust
  vCustId = "MTHR7335" :  sDeleteCust
  vCustId = "MTHR7401" :  sDeleteCust
  vCustId = "MTHR7412" :  sDeleteCust
  vCustId = "MTHR7433" :  sDeleteCust
  vCustId = "MTHR7445" :  sDeleteCust
  vCustId = "MTHR7450" :  sDeleteCust
  vCustId = "MTHR7452" :  sDeleteCust
  vCustId = "MTLB7051" :  sDeleteCust
  vCustId = "MTRX7066" :  sDeleteCust
  vCustId = "MTRX7090" :  sDeleteCust
  vCustId = "MTRX7094" :  sDeleteCust
  vCustId = "MTRX7099" :  sDeleteCust
  vCustId = "MUIR7156" :  sDeleteCust
  vCustId = "MWKS7248" :  sDeleteCust
  vCustId = "MYTR7148" :  sDeleteCust
  vCustId = "MYTR7178" :  sDeleteCust
  vCustId = "MYTR7179" :  sDeleteCust
  vCustId = "MYTR7180" :  sDeleteCust
  vCustId = "MYTR7181" :  sDeleteCust
  vCustId = "MYTR7182" :  sDeleteCust
  vCustId = "MYTR7183" :  sDeleteCust
  vCustId = "MYTR7184" :  sDeleteCust
  vCustId = "MYTR7185" :  sDeleteCust
  vCustId = "MYTR7186" :  sDeleteCust
  vCustId = "MYTR7187" :  sDeleteCust
  vCustId = "MYTR7269" :  sDeleteCust
  vCustId = "MYTR7270" :  sDeleteCust
  vCustId = "MYTR7312" :  sDeleteCust
  vCustId = "NAGC7359" :  sDeleteCust
  vCustId = "NFIB7189" :  sDeleteCust
  vCustId = "NFIB7266" :  sDeleteCust
  vCustId = "NFIB7309" :  sDeleteCust
  vCustId = "NFIB7343" :  sDeleteCust
  vCustId = "NFIB7363" :  sDeleteCust
  vCustId = "NFLD7133" :  sDeleteCust
  vCustId = "NIEL7190" :  sDeleteCust
  vCustId = "NQCD7093" :  sDeleteCust
  vCustId = "NQCD7138" :  sDeleteCust
  vCustId = "NRTK7409" :  sDeleteCust
  vCustId = "NSMS7021" :  sDeleteCust
  vCustId = "NSRC7005" :  sDeleteCust
  vCustId = "NSRC7038" :  sDeleteCust
  vCustId = "NSRC7188" :  sDeleteCust
  vCustId = "NTLN7011" :  sDeleteCust
  vCustId = "NTLN7171" :  sDeleteCust
  vCustId = "NTLN7223" :  sDeleteCust
  vCustId = "NTLN7224" :  sDeleteCust
  vCustId = "NTLN7250" :  sDeleteCust
  vCustId = "NTLN7366" :  sDeleteCust
  vCustId = "NTLN7434" :  sDeleteCust
  vCustId = "NTLN9039" :  sDeleteCust
  vCustId = "NTSL7012" :  sDeleteCust
  vCustId = "NVBZ7282" :  sDeleteCust
  vCustId = "OCSC7240" :  sDeleteCust
  vCustId = "OGCA7206" :  sDeleteCust
  vCustId = "OPCC7255" :  sDeleteCust
  vCustId = "OPCI8947" :  sDeleteCust
  vCustId = "OSLC7100" :  sDeleteCust
  vCustId = "PACC7385" :  sDeleteCust
  vCustId = "PCWK7170" :  sDeleteCust
  vCustId = "PHAB7471" :  sDeleteCust
  vCustId = "PHAB7674" :  sDeleteCust
  vCustId = "PLUS7383" :  sDeleteCust
  vCustId = "PRFC7014" :  sDeleteCust
  vCustId = "PRFC7031" :  sDeleteCust
  vCustId = "PRFC7125" :  sDeleteCust
  vCustId = "PRFC7195" :  sDeleteCust
  vCustId = "PRFC7299" :  sDeleteCust
  vCustId = "PRFC7351" :  sDeleteCust
  vCustId = "PRFC7372" :  sDeleteCust
  vCustId = "PRON7344" :  sDeleteCust
  vCustId = "PSAC7229" :  sDeleteCust
  vCustId = "PSAC7261" :  sDeleteCust
  vCustId = "PSAC7302" :  sDeleteCust
  vCustId = "PSAC8645" :  sDeleteCust
  vCustId = "PTMC7024" :  sDeleteCust
  vCustId = "QTCM7389" :  sDeleteCust
  vCustId = "QUMU7239" :  sDeleteCust
  vCustId = "RBCD7131" :  sDeleteCust
  vCustId = "RESM7391" :  sDeleteCust
  vCustId = "REXA7257" :  sDeleteCust
  vCustId = "REXA7265" :  sDeleteCust
  vCustId = "REXA7268" :  sDeleteCust
  vCustId = "REXA7289" :  sDeleteCust
  vCustId = "REXA7301" :  sDeleteCust
  vCustId = "RISK7429" :  sDeleteCust
  vCustId = "ROCH7149" :  sDeleteCust
  vCustId = "ROHS7083" :  sDeleteCust
  vCustId = "SAPU7446" :  sDeleteCust
  vCustId = "SBHS7246" :  sDeleteCust
  vCustId = "SBMC7135" :  sDeleteCust
  vCustId = "SCAS7437" :  sDeleteCust
  vCustId = "SCOT7087" :  sDeleteCust
  vCustId = "SCOT7091" :  sDeleteCust
  vCustId = "SCOT7092" :  sDeleteCust
  vCustId = "SCOT7308" :  sDeleteCust
  vCustId = "SCOT7325" :  sDeleteCust
  vCustId = "SCPP7022" :  sDeleteCust
  vCustId = "SECC7421" :  sDeleteCust
  vCustId = "SEDT7025" :  sDeleteCust
  vCustId = "SEMC7153" :  sDeleteCust
  vCustId = "SGIC7157" :  sDeleteCust
  vCustId = "SHAK7467" :  sDeleteCust
  vCustId = "SIEU7470" :  sDeleteCust
  vCustId = "SNCL7230" :  sDeleteCust
  vCustId = "SNCL7333" :  sDeleteCust
  vCustId = "SNCL7345" :  sDeleteCust
  vCustId = "SNCL7368" :  sDeleteCust
  vCustId = "SNCL7375" :  sDeleteCust
  vCustId = "SNCL7402" :  sDeleteCust
  vCustId = "SNCL7416" :  sDeleteCust
  vCustId = "SNCL7430" :  sDeleteCust
  vCustId = "SNCL7431" :  sDeleteCust
  vCustId = "SNCL7442" :  sDeleteCust
  vCustId = "SNCL7451" :  sDeleteCust
  vCustId = "STDA7313" :  sDeleteCust
  vCustId = "STWD7071" :  sDeleteCust
  vCustId = "STWD7278" :  sDeleteCust
  vCustId = "TELE7042" :  sDeleteCust
  vCustId = "TELE7045" :  sDeleteCust
  vCustId = "TETK7009" :  sDeleteCust
  vCustId = "TEVA7367" :  sDeleteCust
  vCustId = "THNC7023" :  sDeleteCust
  vCustId = "TICN7060" :  sDeleteCust
  vCustId = "TTBX7002" :  sDeleteCust
  vCustId = "TTBX7004" :  sDeleteCust
  vCustId = "TUNK7340" :  sDeleteCust
  vCustId = "TXRD7101" :  sDeleteCust
  vCustId = "UFCW7032" :  sDeleteCust
  vCustId = "UWAY7175" :  sDeleteCust
  vCustId = "UWAY7196" :  sDeleteCust
  vCustId = "UWAY7197" :  sDeleteCust
  vCustId = "UWAY7198" :  sDeleteCust
  vCustId = "UWAY7199" :  sDeleteCust
  vCustId = "UWAY7200" :  sDeleteCust
  vCustId = "UWAY7201" :  sDeleteCust
  vCustId = "UWAY7202" :  sDeleteCust
  vCustId = "VERG7029" :  sDeleteCust
  vCustId = "VERS7177" :  sDeleteCust
  vCustId = "VERS7220" :  sDeleteCust
  vCustId = "VERS7228" :  sDeleteCust
  vCustId = "VERS7241" :  sDeleteCust
  vCustId = "VERS7242" :  sDeleteCust
  vCustId = "VERS7253" :  sDeleteCust
  vCustId = "VERS7288" :  sDeleteCust
  vCustId = "VERS7322" :  sDeleteCust
  vCustId = "VERS7361" :  sDeleteCust
  vCustId = "VUBZ7001" :  sDeleteCust
  vCustId = "VUBZ7006" :  sDeleteCust
  vCustId = "VUBZ7008" :  sDeleteCust
  vCustId = "VUBZ7015" :  sDeleteCust
  vCustId = "VUBZ7027" :  sDeleteCust
  vCustId = "VUBZ7033" :  sDeleteCust
  vCustId = "VUBZ7037" :  sDeleteCust
  vCustId = "VUBZ7055" :  sDeleteCust
  vCustId = "VUBZ7056" :  sDeleteCust
  vCustId = "VUBZ7057" :  sDeleteCust
  vCustId = "VUBZ7059" :  sDeleteCust
  vCustId = "VUBZ7070" :  sDeleteCust
  vCustId = "VUBZ7072" :  sDeleteCust
  vCustId = "VUBZ7074" :  sDeleteCust
  vCustId = "VUBZ7076" :  sDeleteCust
  vCustId = "VUBZ7077" :  sDeleteCust
  vCustId = "VUBZ7078" :  sDeleteCust
  vCustId = "VUBZ7079" :  sDeleteCust
  vCustId = "VUBZ7086" :  sDeleteCust
  vCustId = "VUBZ7089" :  sDeleteCust
  vCustId = "VUBZ7098" :  sDeleteCust
  vCustId = "VUBZ7102" :  sDeleteCust
  vCustId = "VUBZ7103" :  sDeleteCust
  vCustId = "VUBZ7104" :  sDeleteCust
  vCustId = "VUBZ7105" :  sDeleteCust
  vCustId = "VUBZ7106" :  sDeleteCust
  vCustId = "VUBZ7107" :  sDeleteCust
  vCustId = "VUBZ7121" :  sDeleteCust
  vCustId = "VUBZ7122" :  sDeleteCust
  vCustId = "VUBZ7126" :  sDeleteCust
  vCustId = "VUBZ7128" :  sDeleteCust
  vCustId = "VUBZ7129" :  sDeleteCust
  vCustId = "VUBZ7132" :  sDeleteCust
  vCustId = "VUBZ7136" :  sDeleteCust
  vCustId = "VUBZ7137" :  sDeleteCust
  vCustId = "VUBZ7140" :  sDeleteCust
  vCustId = "VUBZ7142" :  sDeleteCust
  vCustId = "VUBZ7143" :  sDeleteCust
  vCustId = "VUBZ7146" :  sDeleteCust
  vCustId = "VUBZ7147" :  sDeleteCust
  vCustId = "VUBZ7150" :  sDeleteCust
  vCustId = "VUBZ7168" :  sDeleteCust
  vCustId = "VUBZ7192" :  sDeleteCust
  vCustId = "VUBZ7193" :  sDeleteCust
  vCustId = "VUBZ7194" :  sDeleteCust
  vCustId = "VUBZ7203" :  sDeleteCust
  vCustId = "VUBZ7204" :  sDeleteCust
  vCustId = "VUBZ7207" :  sDeleteCust
  vCustId = "VUBZ7208" :  sDeleteCust
  vCustId = "VUBZ7209" :  sDeleteCust
  vCustId = "VUBZ7210" :  sDeleteCust
  vCustId = "VUBZ7211" :  sDeleteCust
  vCustId = "VUBZ7213" :  sDeleteCust
  vCustId = "VUBZ7227" :  sDeleteCust
  vCustId = "VUBZ7231" :  sDeleteCust
  vCustId = "VUBZ7232" :  sDeleteCust
  vCustId = "VUBZ7233" :  sDeleteCust
  vCustId = "VUBZ7234" :  sDeleteCust
  vCustId = "VUBZ7235" :  sDeleteCust
  vCustId = "VUBZ7245" :  sDeleteCust
  vCustId = "VUBZ7249" :  sDeleteCust
  vCustId = "VUBZ7264" :  sDeleteCust
  vCustId = "VUBZ7271" :  sDeleteCust
  vCustId = "VUBZ7283" :  sDeleteCust
  vCustId = "VUBZ7284" :  sDeleteCust
  vCustId = "VUBZ7293" :  sDeleteCust
  vCustId = "VUBZ7300" :  sDeleteCust
  vCustId = "VUBZ7307" :  sDeleteCust
  vCustId = "VUBZ7310" :  sDeleteCust
  vCustId = "VUBZ7314" :  sDeleteCust
  vCustId = "VUBZ7318" :  sDeleteCust
  vCustId = "VUBZ7326" :  sDeleteCust
  vCustId = "VUBZ7328" :  sDeleteCust
  vCustId = "VUBZ7332" :  sDeleteCust
  vCustId = "VUBZ7338" :  sDeleteCust
  vCustId = "VUBZ7339" :  sDeleteCust
  vCustId = "VUBZ7341" :  sDeleteCust
  vCustId = "VUBZ7342" :  sDeleteCust
  vCustId = "VUBZ7348" :  sDeleteCust
  vCustId = "VUBZ7349" :  sDeleteCust
  vCustId = "VUBZ7354" :  sDeleteCust
  vCustId = "VUBZ7369" :  sDeleteCust
  vCustId = "VUBZ7371" :  sDeleteCust
  vCustId = "VUBZ7373" :  sDeleteCust
  vCustId = "VUBZ7381" :  sDeleteCust
  vCustId = "VUBZ7382" :  sDeleteCust
  vCustId = "VUBZ7386" :  sDeleteCust
  vCustId = "VUBZ7393" :  sDeleteCust
  vCustId = "VUBZ7398" :  sDeleteCust
  vCustId = "VUBZ7403" :  sDeleteCust
  vCustId = "VUBZ7404" :  sDeleteCust
  vCustId = "VUBZ7406" :  sDeleteCust
  vCustId = "VUBZ7408" :  sDeleteCust
  vCustId = "VUBZ7411" :  sDeleteCust
  vCustId = "VUBZ7415" :  sDeleteCust
  vCustId = "VUBZ7425" :  sDeleteCust
  vCustId = "VUBZ7435" :  sDeleteCust
  vCustId = "VUBZ7436" :  sDeleteCust
  vCustId = "VUBZ7438" :  sDeleteCust
  vCustId = "VUBZ7440" :  sDeleteCust
  vCustId = "VUBZ7441" :  sDeleteCust
  vCustId = "VUBZ7447" :  sDeleteCust
  vCustId = "VUBZ7448" :  sDeleteCust
  vCustId = "VUBZ7454" :  sDeleteCust
  vCustId = "VUBZ7455" :  sDeleteCust
  vCustId = "VUBZ7460" :  sDeleteCust
  vCustId = "VUBZ7461" :  sDeleteCust
  vCustId = "VUBZ7462" :  sDeleteCust
  vCustId = "VUBZ7463" :  sDeleteCust
  vCustId = "VUBZ7464" :  sDeleteCust
  vCustId = "VUBZ7466" :  sDeleteCust
  vCustId = "VUBZ7469" :  sDeleteCust
  vCustId = "VUBZ7473" :  sDeleteCust
  vCustId = "VUBZ7474" :  sDeleteCust
  vCustId = "VUBZ7475" :  sDeleteCust
  vCustId = "VUBZ7499" :  sDeleteCust
  vCustId = "VUBZ7500" :  sDeleteCust
  vCustId = "VUBZ7657" :  sDeleteCust
  vCustId = "VUBZ7686" :  sDeleteCust
  vCustId = "VUBZ8913" :  sDeleteCust
  vCustId = "WBCO7423" :  sDeleteCust
  vCustId = "WDTW7039" :  sDeleteCust
  vCustId = "WFRD7267" :  sDeleteCust
  vCustId = "WLGN7036" :  sDeleteCust
  vCustId = "WLGN7044" :  sDeleteCust
  vCustId = "WLGN7130" :  sDeleteCust
  vCustId = "WOOD7244" :  sDeleteCust
  vCustId = "WSCH7007" :  sDeleteCust
  vCustId = "WSCH7013" :  sDeleteCust
  vCustId = "WSPS7427" :  sDeleteCust
  vCustId = "WSPS7428" :  sDeleteCust
  vCustId = "XBRL7410" :  sDeleteCust
  vCustId = "XRDS7035" :  sDeleteCust
  vCustId = "XRUN7067" :  sDeleteCust
  vCustId = "YWCA7165" :  sDeleteCust
  vCustId = "ZIIV7407" :  sDeleteCust



  sCloseDb



  Sub sDeleteCust 

    vAcctId = Right(vCustId, 4)

    oDb.Execute("DELETE FROM Cust WHERE Cust_IId        = '" & vCustId & "'")
    oDb.Execute("DELETE FROM Logs WHERE Logs_AcctId     = '" & vAcctId & "'")
    oDb.Execute("DELETE FROM Memb WHERE Memb_AcctId     = '" & vAcctId & "'")
    oDb.Execute("DELETE FROM Catl WHERE Catl_CustId     = '" & vCustId & "'")

    Response.Write vCustId & "<br>"

  End Sub  

%>

