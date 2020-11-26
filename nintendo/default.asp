<!--#INCLUDE FILE="../__INC/includes.asp"-->

<%

  ' ### Recension
  RS_Open 1, "SELECT TOP 1 * FROM cms_Recensioner LEFT JOIN fsBB_Anv ON fsBB_Anv.aID = rSkapadAv WHERE rDatumPublicerad <= '" & Now & "' AND rStatus = 4 ORDER BY rDatumPublicerad DESC", False
  
    If Not rsDB(1).EOF Then
      any_Rec     = True
      
      rec_ID      = CLng(rsDB(1)("rID"))
      
      rec_BildID  = CLng(rsDB(1)("rFlash"))
      
      rec_Titel   = sEncode(CutText(rsDB(1)("rTitel"),35))
      rec_FTitel  = sEncode(rsDB(1)("rTitel"))
    Else
      any_Rec     = False
    End If
  
  RS_Close 1

  ' ### Artikel
  RS_Open 1, "SELECT TOP 1 * FROM cms_Artiklar LEFT JOIN fsBB_Anv ON fsBB_Anv.aID = aaSkapadAv WHERE aaDatumPublicerad <= '" & Now & "' And aaStatus = 4 ORDER BY aaDatumPublicerad DESC", False
  
    'aaID IN(SELECT baArtikel FROM cms_Bind_Artikel_Img WHERE baSaved = 1) AND
  
    If Not rsDB(1).EOF Then
      any_Art     = True
    
      art_ID      = CLng(rsDB(1)("aaID"))
      art_BildID  = CLng(rsDB(1)("aaFlash"))
      art_Titel   = sEncode(CutText(rsDB(1)("aaTitel"),35))
      art_FTitel  = sEncode(rsDB(1)("aaTitel"))
    Else
      any_Art     = False
    End If
  
  RS_Close 1
  
  ' ### Forumtrådar
  If Not config_LockDown_Forum Then
    RS_Open 1, "SELECT TOP 8 tID, tAmne, tStatus_Trad, tStatus_UnderTrad, tDatum_Skapad, fsBB_Forum.fName, fsBB_Anv.aAnvNamn " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "LEFT JOIN fsBB_Anv ON tbTrad.tAnv_Skapad = fsBB_Anv.aID " & _ 
               "WHERE tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND NOT tForum IN (27,32) AND tStatus_Raderad = 0 " & sFilter & " ORDER BY tDatum_Skapad DESC", False
    
      If rsDB(1).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  
  ' ### Kommentarer
  If Not config_LockDown_Kommentarer Then
    RS_Open 1, "SELECT TOP 6 cID, cTextM, fsBB_Anv.aAnvNamn, cDatum, cAvdelning, cBindID, cms_Nyheter.nTitel, cms_Recensioner.rTitel, cms_Artiklar.aaTitel FROM cms_Kommentarer " & _
               "LEFT JOIN fsBB_Anv ON cms_Kommentarer.cAnv = fsBB_Anv.aID " & _
               "LEFT JOIN cms_Nyheter ON cms_Kommentarer.cBindID = cms_Nyheter.nID " & _
               "LEFT JOIN cms_Recensioner ON cms_Kommentarer.cBindID = cms_Recensioner.rID " & _
               "LEFT JOIN cms_Artiklar ON cms_Kommentarer.cBindID = cms_Artiklar.aaID " & _
               "ORDER BY cDatum DESC", False
    
      If rsDB(1).EOF Then
        any_Kommentarer = False
      Else
        any_Kommentarer = True
        list_Kommentarer = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  
  ' ### TextData
  lDaysBack = -365
  RS_Open 1, "SELECT nID, nDatumPublicerad AS tdPubl, nTitel, nKategori, nStatus, nText, nIdent, nFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 0 AND cBindID = nID) AS noComments FROM cms_Nyheter " & _
             "WHERE nStatus = 4 AND (nDatumPublicerad >= '" & DateAdd("d", Now, lDaysBack) & "' AND nDatumPublicerad <= '" & Now & "') " & _
             "UNION ALL " & _
             "SELECT aaID, aaDatumPublicerad AS tdPubl, aaTitel, aaKategori, aaStatus, aaText, aaIdent, aaFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 2 AND cBindID = aaID) AS noComments FROM cms_Artiklar " & _
             "WHERE aaStatus = 4 AND (aaDatumPublicerad >= '" & DateAdd("d", Now, lDaysBack) & "' AND aaDatumPublicerad <= '" & Now & "') " & _
             "UNION ALL " & _
             "SELECT rID, rDatumPublicerad AS tdPubl, rTitel, rKategori, rStatus, rText, rIdent, rFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 1 AND cBindID = rID) AS noComments FROM cms_Recensioner " & _
             "WHERE rStatus = 4 AND (rDatumPublicerad >= '" & DateAdd("d", Now, lDaysBack) & "' AND rDatumPublicerad <= '" & Now & "') " & _
             "ORDER BY tdPubl DESC", False
             
    If rsDB(1).EOF Then
      any_TD = False
    Else
      any_TD = True
      list_TD = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  If any_TD Then
    CreatePaging 12, UBound(list_TD, 2)
    CreatePagingChooser
  End If
  
  ' ### Slumpade spel
  RS_Open 1, "SELECT TOP 9 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE sSynlig = 1 AND (tBoxart_BoxFram > 0 OR tBoxart_Manual > 0 OR tBoxart_Kassett > 0) " & _
             "ORDER BY NewId()", False
  
    If rsDB(1).EOF Then
      any_Rnd  = False
    Else
      list_Rnd   = rsDB(1).GetRows
      any_Rnd     = True
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
  
  ' ### Populära spel
  RS_Open 1, "SELECT TOP 7 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE sSynlig = 1 " & _
             konsol_SQL & _
             "ORDER BY tVisningar DESC", False
  
    If rsDB(1).EOF Then
      any_PopSpel     = False
    Else
      any_PopSpel     = True
      list_PopSpel    = rsDB(1).GetRows(7)
    End If
  
  RS_Close 1
  
  ' ### Bra spel
  RS_Open 1, "SELECT TOP 7 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett, " & _
             "((SELECT SUM(bBetyg) FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID) / (SELECT COUNT(bID) FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID)) AS clBetyg, " & _
             "(SELECT COUNT(*) FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID) AS clBetyg_Antal " & _
             "FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE tSpelID IN(SELECT bSpelID FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID) AND sSynlig = 1 AND tID = sStandard_Titel " & _
             konsol_SQL & _
             "ORDER BY clBetyg DESC, clBetyg_Antal DESC", False
  
    If rsDB(1).EOF Then
      any_BraSpel     = False
    Else
      any_BraSpel     = True
      list_BraSpel    = rsDB(1).GetRows(7)
    End If
  
  RS_Close 1

%>

<%
  ' ## Globala variabler ##
  If pagingOnPage > 1 Then page_Title    = "Sida " & pagingOnPage
  page_Header   = "Nintendo Forum - Första sidan"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Hem&quot; ...'>Första sidan</a> "
  page_SelMenu  = "home"
  page_Slide    = "forum"
  
  page_description  = "Välkommen in till N-Forum.se, Nintendo Forum, där vi har listor över Nintendos alla konsoler, spel och tillbehör med boxart. Du kan även prata med oss i forumet eller köpa och sälja dina tv-spel i vår annonsavdelning."
  page_keywords     = "nintendo forum, "
%>

<!--#INCLUDE FILE="../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full" style="height: 12px;"> </div>
  
    <div class="nf_datablock nf_size_twothird">
    
      <% ' #### RECENSION %>
      <% If rec_BildID > 0 Then rec_Bild = config_ImageLocation & "?e=" & rec_BildID & "&amp;w=301&amp;h=200" Else rec_Bild = config_GFXLocation & "no_rec.png" %>
      <div class="nf_msg nf_flash" style="background-image: url('<% = rec_Bild %>'); margin-right: 20px;" title="<% = rec_FTitel %>">
        <a href="/avdelning/recensioner/recension_visa.asp?e=<% = rec_ID %>" class="nf_afill"> </a>
        <div class="nf_flash_title"><span>Recension</span><a href="/avdelning/recensioner/recension_visa.asp?e=<% = rec_ID %>" title="<% = rec_FTitel %>"><% = rec_Titel %></a><span style="text-align: right;"><a href="/avdelning/recensioner/" style="font-size: 10px;">Fler recensioner...</a></span></div>
      </div>
    
      <% ' #### ARTIKEL %>
      <% If art_BildID > 0 Then art_Bild = config_ImageLocation & "?e=" & art_BildID & "&amp;w=301&amp;h=200" Else art_Bild = config_GFXLocation & "no_art.png" %>
      <div class="nf_msg nf_flash" style="background-image: url('<% = art_Bild %>');" title="<% = art_FTitel %>">
        <a href="/avdelning/artiklar/artikel_visa.asp?e=<% = art_ID %>" class="nf_afill"> </a>
        <div class="nf_flash_title"><span>Artikel</span><a href="/avdelning/artiklar/artikel_visa.asp?e=<% = art_ID %>" title="<% = art_FTitel %>"><% = art_Titel %></a><span style="text-align: right;"><a href="/avdelning/artiklar/" style="font-size: 10px;">Fler artiklar...</a></span></div>
      </div>
      
      <% If any_TD Then %>
        <!-- #### NYHETER / RECENSIONER / ARTIKLAR #### -->

        <ul class="nf_list">
          
          <%
            remPost = CDate("2050-01-01 00:00:00")
            cntNoTimes = 0
            For zx = pagingBOF To pagingEOF
              If zx > UBound(list_TD, 2) Then Exit For
            
              Select Case CLng(list_TD(6, zx))
                Case 0 : sTypText = "Nyhet"     : sPrefixT = "news" : sXText = "nyheten"      : sCatText = "nyheter/nyheter"          : sMinCat = lstKategori(list_TD(3, zx))
                Case 1 : sTypText = "Recension" : sPrefixT = "rec"  : sXText = "recensionen"  : sCatText = "recensioner/recension"    : sMinCat = lstKonsol(list_TD(3, zx))
                Case 2 : sTypText = "Artikel"   : sPrefixT = "art"  : sXText = "artikeln"     : sCatText = "artiklar/artikel"         : sMinCat = lstKonsol(list_TD(3, zx))
              End Select
              
              sWeek = DatePart("ww", list_TD(1, zx), 2, 2)
              
              'If DateDiff("ww", list_TD(1, zx), remPost) <> 0 OR DateDiff("yyyy", list_TD(1, zx), remPost) <> 0 Then 
              If sWeek <> remWeek Then
                remWeek = sWeek
                %>
                <!-- <li class="nf_listsplit"> Vecka <% = remWeek %> </li> -->
                <%
              End If
              %>
                <li>
                  <% If CLng(list_TD(7, zx)) > 0 Then %>
                    <div class="nf_front" style="height: 120px; background-image: url('<% = config_ImageLocation %>?e=<% = list_TD(7, zx) %>&amp;w=180&amp;h=120')"><p>&nbsp;</p></div>
                  <% Else %>
                    <div class="nf_front" style="height: 120px; background-image: url('<% = config_GFXLocation %>icons/no_text.png')"><p>&nbsp;</p></div>
                  <% End If %>
                  <div class="nf_data">
                    <h3><a href="/avdelning/<% = sCatText %>_visa.asp?e=<% = list_TD(0, zx) %>" title="<% = sEncode(list_TD(2, zx)) %>"><% = sEncode(list_TD(2, zx)) %></a></h3>
                    <span class="nf_medium nf_gray nf_bold"><% = sTypText %> / <% = sMinCat %> / <% = DatumReplace(list_TD(1, zx)) %></span>
                    <p style="line-height: 18px;"><% = sEncode(CutText(BBCode_Remove(list_TD(5, zx)),180)) %></p>
                    
                    <div class="nf_morebtn">
                      <a href="/avdelning/<% = sCatText %>_visa.asp?e=<% = list_TD(0, zx) %>">Läs mer ...</a>
                      <a href="/avdelning/<% = sCatText %>_visa.asp?e=<% = list_TD(0, zx) %>#kommentarer" <% If CLng(list_TD(8, zx)) > 0 Then %>class="nf_hint"<% End If %>><% If CLng(list_TD(8, zx)) = 0 Then %>Kommentera texten<% ElseIf CLng(list_TD(8, zx)) = 1 Then %>1 kommentar<% Else %><% = list_TD(8, zx) %> kommentarer<% End If %></a>
                    </div>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
        
        <div class="nf_paging">
          <a href="?page=<% = pagingOnPage - 1 %>">««</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="?page=<% = zx %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
          | <a href="?page=<% = pagingOnPage + 1 %>">»»</a>
        </div>
      <% End If %>

    </div>
      
    <div class="nf_datablock nf_size_onethird">
     
      <!--#INCLUDE FILE="../__INC/_signup.asp"-->
    
      <% ' #### FORUMTRÅDAR %>
      <div class="nf_msg nf_forumflash" style="height: auto;">
        <ul>
          <% firstPost = True %>
          <% For zx = 0 To UBound(list_Tradar, 2) %>
            <%
              If list_Tradar(2,zx) Then
                tradAdd = list_Tradar(0,zx)
              Else
                tradAdd = list_Tradar(3,zx) & "&amp;go2=" & list_Tradar(0,zx)
              End If
            %>
            <li <% If firstPost Then Response.Write(" class='first'") %>><span><% = sEncode(list_Tradar(5,zx)) %> / <% = DatumReplace(list_Tradar(4,zx)) %> / <% = sEncode(CutText(list_Tradar(6,zx),15)) %></span><a href="/avdelning/forum/trad.asp?e=<% = tradAdd %>" title="<% = sEncode(list_Tradar(1,zx)) %>"><% = sEncode(CutText(list_Tradar(1,zx), 35)) %></a></li>
            <% firstPost = False %>
          <% Next %>
          <li class="last"><a href="/avdelning/forum/nyainlagg.asp">Visa fler nya inlägg i forumet</a></li>
        </ul>
      </div>
      
      <% ' #### KOMMENTARER %>
      <% If any_Kommentarer Then %>
        <div class="nf_minibox nf_blue">
          <h4>Senaste kommentarerna</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_Kommentarer, 2) %>
                <%
                Select Case CLng(list_Kommentarer(4, zx))
                  Case 0 : sTypText = "Nyhet"     : sPrefixT = "news" : sXText = "nyheten"      : sCatText = "nyheter/nyheter"          : sMinCat = lstKategori(list_Kommentarer(3, zx))
                  Case 1 : sTypText = "Recension" : sPrefixT = "rec"  : sXText = "recensionen"  : sCatText = "recensioner/recension"    : sMinCat = lstKonsol(list_Kommentarer(3, zx))
                  Case 2 : sTypText = "Artikel"   : sPrefixT = "art"  : sXText = "artikeln"     : sCatText = "artiklar/artikel"         : sMinCat = lstKonsol(list_Kommentarer(3, zx))
                End Select
                %>
              
                <li onclick="location.href='/avdelning/<% = sCatText %>_visa.asp?e=<% = list_Kommentarer(5, zx) %>#kommentar_<% = list_Kommentarer(0, zx) %>';">
                  <a href="/avdelning/<% = sCatText %>_visa.asp?e=<% = list_Kommentarer(5, zx) %>#kommentar_<% = list_Kommentarer(0, zx) %>"><% = sTypText %> / <% = sEncode(CutText(list_Kommentarer(6+list_Kommentarer(4, zx), zx),28)) %></a>
                  <span><% = DatumReplace(list_Kommentarer(3, zx)) %> av <% = sEncode(CutText(list_Kommentarer(2,zx),20)) %></span>
                  <p style="width: 241px; padding: 0; margin: 0; font-weight: normal;"><% = sEncode(CutText(BBCode_Remove(list_Kommentarer(1, zx)),80)) %></p>
                </li>
              <% Next %>
            </ul>
          </div>
        </div>
      <% End If %>
      
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
      
      <div class="nf_msg nf_halfbutton">
        <a href="/avdelning/spel/" title="Visa alla nintendospel">Visa alla spel</a>
      </div>
      
      <% If any_PopSpel Then %>
        <div class="nf_minibox nf_blue">
          <h4>Populära spel</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_PopSpel, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_PopSpel(8, zx)) > 0 Then text_UseBox = list_PopSpel(8, zx)
                If CLng(list_PopSpel(7, zx)) > 0 Then text_UseBox = list_PopSpel(7, zx)
                If CLng(list_PopSpel(2, zx)) > 0 Then text_UseBox = list_PopSpel(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/spel_lrg.png<% End If %>');" onclick="location.href='/avdelning/spel/spel_visa_info.asp?e=<% = list_PopSpel(0, zx) %>';"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = list_PopSpel(0, zx) %>" title="<% = sEncode(list_PopSpel(1, zx)) %>"><% = sEncode(CutText(list_PopSpel(1, zx), 32)) %></a><% = lstKonsol(list_PopSpel(5, zx)) %></li>
              <% Next %>
            </ul>
          </div>
        </div>
      <% End If %>
      
      <% If any_BraSpel Then %>
        <div class="nf_minibox nf_blue">
          <h4>Våra medlemmar gillar</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_BraSpel, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_BraSpel(8, zx)) > 0 Then text_UseBox = list_BraSpel(8, zx)
                If CLng(list_BraSpel(7, zx)) > 0 Then text_UseBox = list_BraSpel(7, zx)
                If CLng(list_BraSpel(2, zx)) > 0 Then text_UseBox = list_BraSpel(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/spel_lrg.png<% End If %>');" onclick="location.href='/avdelning/spel/spel_visa_info.asp?e=<% = list_BraSpel(0, zx) %>';"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = list_BraSpel(0, zx) %>" title="<% = sEncode(list_BraSpel(1, zx)) %>"><% = sEncode(CutText(list_BraSpel(1, zx), 32)) %></a><% = lstKonsol(list_BraSpel(5, zx)) %></li>
              <% Next %>
            </ul>
          </div>
        </div>
      <% End If %>
      
      <% If CONST_CMS Then %>
        <div class="nf_minibox nf_red">
          <h4>Administration</h4>
           <div class="nf_inside">
             <p><strong>Notera följande:</strong></p>
           
             <% If GetAcc("CMS111") Then %>
               <p><span style="float: right;"><strong> <% = Con.ExeCute("SELECT COUNT(nID) FROM cms_Nyheter WHERE nStatus = 2")(0) %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> Nyheter </p>
               <p><span style="float: right;"><strong> <% = Con.ExeCute("SELECT COUNT(rID) FROM cms_Recensioner WHERE rStatus = 2")(0) %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> Recensioner </p>
               <p><span style="float: right;"><strong> <% = Con.ExeCute("SELECT COUNT(aaID) FROM cms_Artiklar WHERE aaStatus = 2")(0) %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> Artiklar </p>
             <% End If %>
            
             <% If GetAcc("CMS3") Then %>
               <p><span style="float: right;"><strong> <% = Con.ExeCute("SELECT COUNT(anID) FROM fsBB_Anmal WHERE anDatum > '" & DATEADD("m", -1, Now) & "' And anNoterad = 0")(0) %></strong></span> <img src="<% = config_GFXLocation %>icons/text.png"> Anmälningar </p>
             <% End If %>
             
             <p><a href="http://cms.n-forum.se" target="_blank">» Gå till administrationen</a></p>
          </div>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_full">
      <% If any_Rnd Then %>
        <!-- #### 10 SLUMPADE SPEL #### -->
        <div class="nf_images nf_images_full">
          <p><strong>Upptäck följande spel...</strong></p>
          <% For zx = 0 To UBound(list_Rnd, 2) %>
            <%
            text_UseBox = 0
            If CLng(list_Rnd(8, zx)) > 0 Then text_UseBox = list_Rnd(8, zx)
            If CLng(list_Rnd(7, zx)) > 0 Then text_UseBox = list_Rnd(7, zx)
            If CLng(list_Rnd(2, zx)) > 0 Then text_UseBox = list_Rnd(2, zx)
            %>
            <a href="/avdelning/spel/spel_visa_info.asp?e=<% = list_Rnd(0,zx) %>"><img src="<% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=80&amp;h=80&amp;err=no" title="<% = sEncode(list_Rnd(1, zx)) %> (<% = list_Rnd(4, zx) %>)" alt="Nintendo Spel"></a>
          <% Next %>
          <p>... eller visa <a href="/avdelning/spel/default.asp">hela spellistan</a> istället.</p>
        </div>
      <% End If %>
    </div>
    
    
    
  </div>

<!--#INCLUDE FILE="../_page_bottom.asp"-->
<!--#INCLUDE FILE="../__INC/includes_end.asp"-->