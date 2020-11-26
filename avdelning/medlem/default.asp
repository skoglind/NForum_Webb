<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  lMedlem = GetQ("m","ABC",50)
  If Trim(lMedlem) = Empty Then lMedlem = CONST_USERNAME
  
  anvID = GetIDFromUsername(lMedlem)

  RS_Open 1, "SELECT * FROM fsBB_Anv LEFT JOIN fsBB_Titlar ON fsBB_Anv.aTitelID = fsBB_Titlar.ttID WHERE aID = " & CLng(anvID), False
  
     If rsDB(1).EOF Then Response.Redirect(config_NotLoggedIn)
  
     text_ID        = CLng(rsDB(1)("aID"))
     text_EgenTitel = sEncode(rsDB(1)("aEgenTitel"))
     text_AnvNamn   = sEncode(rsDB(1)("aAnvNamn"))
     text_Namn      = sEncode(rsDB(1)("aNamn"))
     text_Plats      = sEncode(rsDB(1)("aPlats"))
     text_RegDatum  = FormatDateTime(rsDB(1)("aMedlemSedan"), vbShortDate)
     text_LoginDatum= DatumReplace(rsDB(1)("aInloggadSenast"))
     text_Titel     = sEncode(rsDB(1)("ttText"))
     text_PM        = Trim(BBCode(rsDB(1)("aPM"), True))
     text_MSN       = sEncode(rsDB(1)("aMSN"))
     text_ICQ       = sEncode(rsDB(1)("aICQ"))
     text_AktivPM   = rsDB(1)("aAktiveraPM")
     
     text_bAvatar   = rsDB(1)("aAvatar")
     text_Avatar    = config_Avatar & "u" & Right("000000" & anvID, 6) & ".jpg"
     
     text_Hemsida   = Trim(rsDB(1)("aHemsida"))
     If Len(text_Hemsida) > 1 Then
       If Left(LCase(text_Hemsida), 7) <> "http://" And Left(LCase(text_Hemsida), 8) <> "https://" Then text_Hemsida = "http://" & text_Hemsida
     End If
     
     If rsDB(1)("aTimeStamp") > DateAdd("n", -5, Now) Then
       text_OnOff = "<span style='color: #0A0; font-weight: bold;'>Online</span>"
     Else
       text_OnOff = "<span style='color: #A00;'>Offline</span>"
     End If
     
  RS_Close 1
  
  RS_Open 1, "SELECT TOP 7 abID, abTitel, abUppladdadDatum FROM cms_AnvBilder WHERE abUppladdadAv = " & CLng(anvID), False
                 
    If rsDB(1).EOF Then
      any_Bild = False
    Else
      any_Bild = True
      list_Bild = rsDB(1).GetRows()
      antal_bilder = UBound(list_Bild, 2) + 1
    End If
  
  RS_Close 1
  
  text_AntalSpel      = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Spel WHERE biAnv = " & CLng(anvID))(0)
  text_AntalKonsoler  = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Konsol WHERE biAnv = " & CLng(anvID))(0)
  text_AntalTillbehor = Con.ExeCute("SELECT COUNT(biID) FROM cms_Bind_Anv_Tillbehor WHERE biAnv = " & CLng(anvID))(0)
  
  text_AntalRecensioner      = Con.ExeCute("SELECT COUNT(rID) FROM cms_Recensioner WHERE rDatumPublicerad <= '" & Now & "' And rStatus = 4 AND rSkapadAv = " & CLng(anvID))(0)
  text_AntalArtiklar         = Con.ExeCute("SELECT COUNT(aaID) FROM cms_Artiklar WHERE aaDatumPublicerad <= '" & Now & "' And aaStatus = 4 AND aaSkapadAv = " & CLng(anvID))(0)
  text_AntalTipsoTrix        = Con.ExeCute("SELECT COUNT(xID) FROM cms_SpelTrix WHERE xStatus = 4 AND xSkapadAv = " & CLng(anvID))(0)
  
  If Not config_LockDown_Forum Then
    text_AntalInlagg    = Con.ExeCute("SELECT COUNT(tID) " & _
                                      "FROM fsBB_Tradar AS tbTrad " & _
                                      "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
                                      "WHERE tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tStatus_Raderad = 0 " & sFilter & " AND tAnv_Skapad = " & CLng(anvID))(0)
  Else
    text_AntalInlagg    = 0
  End If
%>
  
<%

  ' ## Globala variabler ##
  page_Title    = text_AnvNamn & " - Medlem"
  page_Header   = text_AnvNamn & "s profil"
  page_WhereAmI = "&gt; <a href='default.asp?m=" & lMedlem & "' title='Gå till &quot;Hem&quot; ...'>Profil</a> "
  page_SelMenu  = "user"
  page_Slide    = "medlem"
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <% If CONST_LOGIN And CLng(text_ID) = CLng(CONST_USERID) Then %>
    <!--#INCLUDE FILE="__menu_u.asp"-->
  <% Else %>
    <!--#INCLUDE FILE="__menu_other.asp"-->
  <% End If %>
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
   
    <div class="nf_datablock nf_size_full">
      <h1><% = lMedlem %></h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <% If CLng(anvID) = CLng(CONST_USERID) Then %>
        <div class="nf_msg">
          <p><strong>Permalänk:</strong> http://<% = page_NForum %>/q/profil/?m=<% = sEncode(lMedlem) %></p>
        </div>
      <% End If %>
      
      <% If any_Bild Then %>
        <div class="nf_images">
          <p><strong>Titta på några av <% = lMedlem %>s bilder...</strong></p>
          <% For zx = 0 To UBound(list_Bild, 2) %>
            <a href="/userimage.asp?e=<% = list_Bild(0, zx) %>" target="_blank" rel="lightbox[minabilder]"><img src="/userimage.asp?e=<% = list_Bild(0, zx) %>" title="<% = sEncode(list_Bild(1, zx)) %> / Uppladdad: <% = DatumReplace(list_Bild(2, zx)) %>" alt="<% = sEncode(list_Bild(1, zx)) %>"></a>
          <% Next %>
          <p>... eller <a href="minabilder.asp?m=<% = lMedlem %>">visa alla</a> bilder.</p>
        </div>
      <% End If %>
      
      <div class="nf_text">
        <% If Len(text_PM) > 0 Then %>
          <p><% = text_PM %></p>
        <% Else %>
          <p><em>Användaren har inte skrivit någon personlig text.</em></p>
        <% End If %>
      </div>
      
      <% If CLng(anvID) <> CLng(CONST_USERID) AND CONST_LOGIN Then %>
        <div class="nf_msg">
          <p>» <a href="/avdelning/listor/sokmedlem.asp">Sök upp en annan medlem...</a></p>
        </div>
      <% End If %>
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_minibox nf_blue">
        <h4>Profil</h4>
        <div class="nf_inside">
          <% If text_bAvatar Then %><p><img src="<% = text_Avatar %>" alt="Avatar" style="border: solid 1px #CCC; width: 100px; height: 100px; padding: 1px;"></p><% End If %>
          <p>
            <% If Len(Trim(text_Namn)) > 0 Then %><strong>Namn: </strong> <% = text_Namn %><br><% End If %>
            <% If Len(Trim(text_Plats)) > 0 Then %><strong>Plats: </strong> <% = text_Plats %><br><% End If %>
            <strong>Medlem sedan: </strong> <% = text_RegDatum %>
          </p>
          <p><strong>Senast inloggad: </strong> <% = text_LoginDatum %></p>
          <p><strong>Status: </strong> <% = text_OnOff %></p>
          <p>
            <% If Len(Trim(text_Hemsida)) > 0 Then %><strong>Hemsida: </strong> <a href="<% = text_Hemsida %>" rel="nofollow" target="_blank">Gå till hemsida</a><br><% End If %>
            <% If Len(Trim(text_MSN)) > 0 Then %><strong>MSN: </strong> <a href="msnim:chat?contact=<% = text_MSN %>" rel="nofollow"><% = text_MSN %></a><br><% End If %>
            <% If Len(Trim(text_ICQ)) > 0 Then %><strong>ICQ: </strong> <% = text_ICQ %><% End If %>
          </p>
          
          <% If CONST_LOGIN And text_AktivPM And anvID <> CONST_USERID THEN %><p>» <a href="skrivpm.asp?m=<% = text_AnvNamn %>">Skicka PM till <% = text_AnvNamn %></a></p><% End If %>
        </div>
      </div>
      
      <div class="nf_minibox">
        <h4>Spelsamling</h4>
        <div class="nf_inside">
          <p><span style="float: right;"><strong><% = text_AntalSpel %></strong></span> <img src="<% = config_GFXLocation %>icons/spel.png"> <a href="minaspel.asp?list=spel&m=<% = lMedlem %>">Spel</a> </p>
          <p><span style="float: right;"><strong><% = text_AntalKonsoler %></strong></span> <img src="<% = config_GFXLocation %>icons/konsol.png"> <a href="minaspel.asp?list=konsol&m=<% = lMedlem %>">Konsoler</a> </p>
          <p><span style="float: right;"><strong><% = text_AntalTillbehor %></strong></span> <img src="<% = config_GFXLocation %>icons/tillbehor.png"> <a href="minaspel.asp?list=tillbehor&m=<% = lMedlem %>">Tillbehör</a> </p>
        </div>
      </div>
    
      <div class="nf_minibox">
        <h4>Texter</h4>
        <div class="nf_inside">
          <p><span style="float: right;"><strong><% = text_AntalRecensioner %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> <a href="minatexter.asp?list=recensioner&m=<% = lMedlem %>">Recensioner</a> </p>
          <p><span style="float: right;"><strong><% = text_AntalArtiklar %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> <a href="minatexter.asp?list=artiklar&m=<% = lMedlem %>">Artiklar</a> </p>
          <p><span style="float: right;"><strong><% = text_AntalTipsoTrix %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> <a href="minatexter.asp?list=tipsotrix&m=<% = lMedlem %>">Tips & Trix</a> </p>
        </div>
      </div>
      
      <div class="nf_minibox">
        <h4>Övrigt</h4>
        <div class="nf_inside">
          <p><span style="float: right;"><strong><% = text_AntalInlagg %></strong></span> <img src="<% = config_GFXLocation %>icons/trad.png"> <a href="minainlagg.asp?m=<% = lMedlem %>">Foruminlägg</a> </p>
        </div>
      </div>
    </div>  
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->