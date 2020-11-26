<!--#INCLUDE FILE="../../../__INC/includes.asp"-->

  <%

    Call start_Rec2Session("chpass")
    
    sAnvNamn      = Trim(GetF("aNamn","ABC",30))
    sNyckel       = Trim(GetF("nyckel","ABC",10))
    sPass1        = GetF("passwd1","ABC",0)
    sPass2        = GetF("passwd2","ABC",0)

    If MakeLegal(sAnvNamn) <> sAnvNamn  Then Response.Redirect("../nyttlosenord.asp?fail=1")  ' Oglitigt användarnamn
    If Len(sAnvNamn) < 1                Then Response.Redirect("../nyttlosenord.asp?fail=1")  ' Inget användarnamn
    If MakeLegal(sNyckel) <> sNyckel    Then Response.Redirect("../nyttlosenord.asp?fail=1")  ' Oglitigt nyckel
    If Len(sNyckel) < 10                Then Response.Redirect("../nyttlosenord.asp?fail=1")  ' Ingen nyckel
    
    If Len(Trim(sPass1)) < 1            Then Response.Redirect("../nyttlosenord.asp?fail=2")  ' Inget lösenord
    If Len(Trim(sPass1)) < 7            Then Response.Redirect("../nyttlosenord.asp?fail=3")  ' För kort
    If sPass1 <> sPass2                 Then Response.Redirect("../nyttlosenord.asp?fail=4")  ' De stämmer inte
    
    RS_Open 1, "SELECT * FROM fsBB_Anv WHERE aAnvNamn = '" & MakeLegal(sAnvNamn) & "' AND aPassKey = '" & MakeLegal(sNyckel) & "'", True
    
      If rsDB(1).EOF                  Then Response.Redirect("../nyttlosenord.asp?fail=1")  ' Värdena stämde inte
      If Not rsDB(1)("aNewPass")      Then Response.Redirect("../nyttlosenord.asp?fail=1")  ' Lösenordsbyte har aldrig valts
      
      ' Ok, kör på. Allt godkänt byt lösenordet
      
      rsDB(1)("aPassKey")       = "0"
      rsDB(1)("aNewPass")       = False
      
      sDBSalt1  = rsDB(1)("aSalt1")
      sDBSalt2  = rsDB(1)("aSalt2")
      sHash     = config_Hash_Salt_1 & "" & sDBSalt1 & "" & sPass1 & "" & config_Hash_Salt_2 & "" & sDBSalt2
      sHash     = MD5(sHash)

      rsDB(1)("aPassWd")        = sHash
      rsDB(1)("aNyttLosenord")  = True
      
      rsDB(1).Update
      
    RS_Close 1
    
    Call stop_Rec2Session("chpass")
    Session.Value("form_saved") = True
    Call SayMe("Sparad","Ditt <strong>lösenord</strong> har nu ändrats!", "/avdelning/medlem/nyttlosenord.asp")
  
  %>

<!--#INCLUDE FILE="../../../__INC/includes_end.asp"-->