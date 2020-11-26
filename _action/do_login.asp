<!--#INCLUDE FILE="../__INC/includes.asp"-->

  <%
    Session.Value("NFORUM_LOGIN") = False

    sAnvNamn    = Trim(GetF("r","ABC",50))
    sLosenord   = GetF("g","ABC",50)
    bRemember   = GetF("s","CHK",0)
    sPB         = GetF("postback","ABC",500)
  
    ' ### KÖR INLOGGNINGSRUTINEN
    LoginUser sAnvNamn, sLosenord, False, bRemember, sPB
  %>

<!--#INCLUDE FILE="../__INC/includes_end.asp"-->