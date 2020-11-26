<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  ' ## Hämta all data ##
  lID = GetQ("e", "123", 0)
  
  filter_kat = GetQ("k","123",0)
  If filter_kat > 0 And filter_kat < lstKSKategori(-1) + 1 Then kat_SQL = "AND ksKategori1 = " & CLng(filter_kat)  
  
  add_URL    = "&amp;page=" & iPageData & "&amp;q=" & sQUrl & "&amp;k=" & filter_kat
  
  If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
    RS_Open 1, "SELECT * FROM cms_KopSalj LEFT JOIN fsBB_Anv ON cms_KopSalj.ksSkapadAv = fsBB_Anv.aID WHERE ksID = " & lID, False
  Else
    RS_Open 1, "SELECT * FROM cms_KopSalj LEFT JOIN fsBB_Anv ON cms_KopSalj.ksSkapadAv = fsBB_Anv.aID WHERE ((ksSkapadDatum + " & CLng(config_AdDays) & " > '" & Now & "' AND ksSynlig = 1) OR ksSkapadAv = " & CLng(CONST_USERID) & ") And ksID = " & lID, False
  End if
  
    If rsDB(1).EOF Then Response.Redirect("default.asp")
    
    text_Titel      = rsDB(1)("ksTitel")
    text_Text       = sEncode(rsDB(1)("ksTextM"))
    text_TextE      = sEncode(CutText(BBCode_Remove(rsDB(1)("ksTextM")),200))
    text_AvNamn     = sEncode(rsDB(1)("aNamn"))
    text_AvIn       = sEncode(rsDB(1)("aAnvNamn"))
    text_AvID       = rsDB(1)("ksSkapadAv")
    text_Kategori   = rsDB(1)("ksKategori1")
    text_Typ        = rsDB(1)("ksTyp")
    text_Publicerad = rsDB(1)("ksSkapadDatum")
    text_Status     = rsDB(1)("ksStatus")
    text_Synlig     = rsDB(1)("ksSynlig")
  
  RS_Close 1
  
  ' #### FIX TEXT STRÄNG ####
    q = LCase(Trim(text_Titel))
    
    q = MakeLegal(q)
    w = Split(q, " ")
    
    For Each ww In w
      ww = Trim(ww)
      ww = Replace(ww, ":", "")
      ww = Replace(ww, "'", "")
      If IsNumeric(ww) Then If ww > 1979 And ww < 2050 Then ww = ""
      If IsNumeric(ww) Then If ww > 1979 And ww < 2050 Then ww = ""
      
      If Len(ww) > 3 Then
        p = p & """" & ww & """ OR "
      End If
    Next

    p = Left(p, Len(p)-4)
    p = "'(" & p & ")'"
    
    'Response.Write p
  ' #### ^
  
  ' #### Kommentarer
  RS_Open 1, "SELECT kskID, kskTextM, kskAnv, kskDatum, kskRaderadAv, kskAnnons, " & _
             "fsBB_Anv.aAnvNamn, fsBB_Anv.aID, fsBB_Anv.aAvatar, fsBB_Anv.aPlats, fsBB_Anv.aTimeStamp, fsBB_Anv.aAktiveraPM " & _
             "FROM cms_Kommentar_KopSalj " & _
             "LEFT JOIN fsBB_Anv ON cms_Kommentar_KopSalj.kskAnv = aID " & _
             "WHERE kskAnnons = " & CLng(lID) & " " & _
             "ORDER BY kskDatum ASC", False
   
     If rsDB(1).EOF Then
      any_Comments = False
    Else
      any_Comments = True
      list_Comments = rsDB(1).GetRows
    End If
   
  RS_Close 1
  
  ' ### Annonser
  RS_Open 1, "SELECT TOP 5 ksID, ksSkapadDatum, ksTitel, ksKategori1, ksStatus, ksKategori2, ksSkapadAv, ksTyp, ksTextM, (SELECT aAnvNamn FROM fsBB_Anv WHERE aID = ksSkapadAv) AS ksGetUser " & _
             "FROM cms_KopSalj WHERE ksID > 0 " & _ 
             "AND ksSkapadDatum + " & CLng(config_AdDays) & " > '" & Now & "' " & _
             "AND ksSynlig = 1 " & _
             "ORDER BY ksSkapadDatum DESC", False
  
    If rsDB(1).EOF Then
      any_Ads = False
    Else
      any_Ads = True
      list_Ads = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  canEdit = False
  If CLng(CONST_USERID) = CLng(text_AvID) Then canEdit = True
%>

<%
  ' ## Globala variabler ##
  page_Title    = sEncode(text_Titel) & " - Annonser"
  page_Header   = sEncode(text_Titel)
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Marknad&quot; ...'>Marknad</a> " & _
                  "&gt; Annons"
  page_SelMenu  = "buy"
  page_Slide    = "annonser"
  
  page_description    = "Du visar just nu annonsen (" & sEncode(text_Titel) & ") i vår köp- och säljavdelning på N-Forum.se, Nintendo Forum. " & Replace(text_TextE, vbCrlf, " ") & "..."
  page_keywords       = "visa annons, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
   
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/annonser/">Annonser</a></span><% = sEncode(text_Titel) %> <% If text_Status = 1 Then %> <span style="color: #A00;">(Avslutad!)</span><% End If %> <% If Not text_Synlig Then %> <span style="color: #00A;">(Osynlig!)</span><% End If %></h1>
      <h4><% = lstKSTyp(text_Typ) %> / <% = lstKSKategori(text_Kategori) %> / Upplagd <% = DatumReplace(text_Publicerad) %> av <a href="/avdelning/medlem/?m=<% = text_AvIn %>"><% = text_AvIn %></a></h4>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      <div class="nf_text">
        <p><% = BBCode(text_Text, False) %></p>
      </div>
      
      <div class="nf_msg">
        <p style="color: #AAA;"><a href="/avdelning/medlem/?m=<% = text_AvIn %>">» Annonsörens profil</a> | <a href="/avdelning/medlem/skrivpm.asp?m=<% = text_AvIn %>">» PMa annonsören</a> | <a href="/avdelning/medlem/omdome.asp?m=<% = text_AvIn %>">» Annonsörens omdömen</a></p>
      </div>
      
      <ul class="nf_list">
        
        <li class="nf_listsplit"> Kommentarer </li>
        
        <% If any_Comments Then %>
        
          <% For zx = 0 To UBound(list_Comments, 2) %>
            <li> 
              <div class="nf_header">
                <span class="nf_big">#<% = zx + 1 %></span>
              </div>
              <div class="nf_data">
                <span class="nf_medium nf_gray nf_bold">
                  <a href="/avdelning/medlem/?m=<% = list_Comments(6, zx) %>"><% = list_Comments(6, zx) %></a> / <% = DatumReplace(list_Comments(3, zx)) %>
                  <% If (CLng(CONST_USERID) = CLng(list_Comments(2, zx)) Or HasAcc(CONST_CMS_RIGHTS,"CMS700")) And CLng(list_Comments(4, zx)) = 0 Then %><img src="<% = config_GFXLocation %>icons/del.png" onclick="doActionWithPrompt('_action/deletecomment.asp?e=<% = list_Comments(0, zx) %>','Vill du ta bort kommentaren?');" style="float: right; cursor: pointer;" title="Ta bort kommentaren" alt="Radera"><% End If %>
                </span>
                
                <% If CLng(list_Comments(4, zx)) = 0 Then %>
                  <p><% = BBCode(list_Comments(1, zx), True) %></p>
                <% Else %>
                  <p style="font-size: 10px !important;font-style: italic !important; color: #A00 !important;">Kommentaren är borttagen av <strong><% If CLng(list_Comments(4, zx)) = CLng(list_Comments(2, zx)) Then %>användaren<% Else %>administratören<% End If %></strong>!</p>
                  
                  <% If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then %>
                    <p style="font-size: 10px !important; color: #CCC !important;"><% = BBCode(list_Comments(1, zx), True) %></p>
                  <% End If %>
                <% End If %>
                
                <% If CLng(list_Comments(4, zx)) = 0 And list_Comments(11, zx) And CONST_LOGIN Then %><span class="nf_small nf_bold">» <a href="/avdelning/medlem/skrivpm.asp?m=<% = list_Comments(6, zx) %>">Skicka PM</a></span><% End If %>
              </div>
            </li>
          <% Next %>
        
        <% End If %>
        
      </ul>
      
      <% If Not any_Comments Then %>
        <div class="nf_msg">
          <p> Det finns inga kommentarer. </p>
        </div>
      <% End If %>
      
      <% If CONST_LOGIN Then %>
        <form method="POST" action="_action/postcomment.asp">
          <div class="nf_form">
  
            <div class="nf_falt"><textarea name="aMsg" style="height: 100px; width: 576px"></textarea></div>
            
            <div class="nf_falt nf_buttons">
              <input type="hidden" name="e" value="<% = lID %>">
              <input type="submit" style="font-weight: bold;" value="Posta">
            </div>
  
          </div>
        </form>
      <% Else %>
        <div class="nf_msg nf_green">
          <p style="text-align: center;"><em>Du måste <strong><a href="<% = config_NotLoggedIn %>">logga in</a></strong> för att kunna lämna kommentarer.</em></p>
          <p style="text-align: center;"><em>Om du inte redan har en användare kan du <strong><a href="/avdelning/medlem/registreradig.asp">registrera dig</a></strong>.</em></p>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <% If canEdit Or HasAcc(CONST_CMS_RIGHTS,"CMS700") Then %>
        <div class="nf_minibox">
          <h4>Hantera annons</h4>
          <div class="nf_inside">
            <p><img src="<% = config_GFXLocation %>icons/edit.png"> <a href="ny_annons.asp?e=<% = lID %>">Redigera annons</a></p>
            <p><img src="<% = config_GFXLocation %>icons/del.png"> <a onclick="doActionWithPrompt('_action/deleteannons.asp?e=<% = lID %>','Vill du ta bort annonsen?');" style="cursor: pointer;">Ta bort annons</a></p>
          </div>
        </div>
      <% End If %>
      
      <div class="nf_minibox">
        <h4>Dela med dig</h4>
        <div class="nf_inside">
          <!-- AddThis Button BEGIN -->
            <div class="addthis_toolbox addthis_default_style" addthis:title="<% = text_Titel %>" addthis:description="<% = text_TextE %>">
              <a class="addthis_button_email" title="E-posta"></a>
              <a class="addthis_button_print" title="Skriv ut"></a>
              <span class="addthis_separator">|</span>
              <a class="addthis_button_facebook" title="Facebook"></a>
              <a class="addthis_button_twitter" title="Twitter"></a>
              <a class="addthis_button_digg" title="Digg"></a>
              <a class="addthis_button_pusha" title="Pusha"></a>
              <a class="addthis_button_blogger" title="Blogger"></a>
              <a class="addthis_button_delicious" title="Del.icio.us"></a>
              <a class="addthis_button_google" title="Google"></a>
            </div>
            <script type="text/javascript" src="http://s7.addthis.com/js/250/addthis_widget.js#username=nforum"></script>
          <!-- AddThis Button END -->
        </div>
      </div>
      
      <% If any_Ads Then %>
        <div class="nf_minibox nf_green">
          <h4>Annonser</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_Ads, 2) %>
                <li onclick="location.href='/avdelning/annonser/annons_visa.asp?e=<% = list_Ads(0, zx) %>';"><a href="/avdelning/annonser/annons_visa.asp?e=<% = list_Ads(0, zx) %>" title="<% = sEncode(list_Ads(2, zx)) %>"><% = sEncode(CutText(list_Ads(2, zx), 32)) %></a><% = lstKSTyp(list_Ads(7, zx)) %> / <% = DatumReplace(list_Ads(1, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/annonser/">Visa alla annonser</a></p>
          </div>
        </div>
      <% End If %>

    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->