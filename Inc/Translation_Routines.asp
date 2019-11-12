<%
  Function fFr (vPhrase)
    fFr = vPhrase
    fFr = Replace(fFr, "à", "&#224;") 
    fFr = Replace(fFr, "ç", "&#231;") 
    fFr = Replace(fFr, "è", "&#232;") 
    fFr = Replace(fFr, "é", "&#233;") 
    fFr = Replace(fFr, "ê", "&#234;") 

    fFr = Replace(fFr, "À", "&#192;") 
    fFr = Replace(fFr, "Ç", "&#199;") 
    fFr = Replace(fFr, "È", "&#200;") 
    fFr = Replace(fFr, "É", "&#201;") 
    fFr = Replace(fFr, "Ê", "&#202;") 
  End Function
%>  