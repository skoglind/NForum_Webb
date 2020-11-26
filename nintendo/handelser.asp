<!--#INCLUDE FILE="../__INC/includes.asp"-->

<%

  ' ### Fler nyheter
  lDaysBack = -120
  RS_Open 1, "SELECT nID, nTitel, nDatumPublicerad AS tdPubl, nKategori, nStatus, nText, nIdent, nFlash FROM cms_Nyheter " & _
             "WHERE nStatus = 4 AND (nDatumPublicerad >= '" & DateAdd("d", Now, lDaysBack) & "' AND nDatumPublicerad <= '" & Now & "') " & _
             "UNION ALL " & _
             "SELECT aaID, aaTitel, aaDatumPublicerad AS tdPubl, aaKategori, aaStatus, aaText, aaIdent, aaFlash FROM cms_Artiklar " & _
             "WHERE aaAnvandarArt = 0 AND aaStatus = 4 AND (aaDatumPublicerad >= '" & DateAdd("d", Now, lDaysBack) & "' AND aaDatumPublicerad <= '" & Now & "') " & _
             "UNION ALL " & _
             "SELECT rID, rTitel, rDatumPublicerad AS tdPubl, rKategori, rStatus, rText, rIdent, rFlash FROM cms_Recensioner " & _
             "WHERE rAnvandarRec = 0 AND rStatus = 4 AND (rDatumPublicerad >= '" & DateAdd("d", Now, lDaysBack) & "' AND rDatumPublicerad <= '" & Now & "') " & _
             "ORDER BY tdPubl DESC", False
  
    If rsDB(1).EOF Then
      any_TextData     = False
    Else
      any_TextData     = True
      list_TextData    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler spel
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE sSynlig = 1 " & _
             "ORDER BY sDatumSparad DESC", False
  
    If rsDB(1).EOF Then
      any_updSpel     = False
    Else
      any_updSpel     = True
      list_updSpel    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler konsoler
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, kKonsol, kTextM, tBoxart_Manual, tBoxart_Konsol FROM cms_KonsolTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Konsol ON kID = tKonsolID " & _
             "WHERE kSynlig = 1 " & _
             "ORDER BY kDatumSparad DESC", False
  
    If rsDB(1).EOF Then
      any_updKonsol     = False
    Else
      any_updKonsol     = True
      list_updKonsol    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler tillbehör
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, iKonsol, iTextM, tBoxart_Manual, tBoxart_Tillbehor FROM cms_TillbehorTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Tillbehor ON iID = tTillbehorID " & _
             "WHERE iSynlig = 1 " & _
             "ORDER BY iDatumSparad DESC", False
  
    If rsDB(1).EOF Then
      any_updTillbehor     = False
    Else
      any_updTillbehor     = True
      list_updTillbehor    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler foruminlägg
  If Not config_LockDown_Forum Then
    RS_Open 1, "SELECT TOP 10 tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, fIcon, fName " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "WHERE tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tForum <> " & CLng(config_Trashbin) & " AND tStatus_Raderad = 0 ORDER BY tDatum_Skapad DESC", False
    
      If rsDB(1).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(1).GetRows(10)
      End If
    
    RS_Close 1
  End If
  
  ' ### Fler senast registrerade
  RS_Open 1, "SELECT TOP 10 aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aMedlemSedan, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
             "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
             "WHERE aBlockadTill < '" & Date & "' AND aAktiverad = 1 ORDER BY aMedlemSedan DESC", False
  
    If rsDB(1).EOF Then
      any_On = False
    Else
      any_On = True
      list_On = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1

%>

<%
  ' ## Globala variabler ##
  page_Title    = "Händelser på N-Forum.se"
  page_Header   = "Händelser på N-Forum.se"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Hem&quot; ...'>Första sidan</a> "
  page_SelMenu  = "home"
  page_Slide    = "forum"
  
  page_description  = "Senaste uppdateringarna och händelserna på N-Forum.se, Nintendo Forum. Uppdaterade spel, konsoler och tillbehör. Senaste foruminläggen, texterna och registrerade medlemmarna."
  page_keywords     = "händelser, "
%>

<!--#INCLUDE FILE="../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_onethird">
     
      <!-- ## SENAST UPPDATERADE SPELEN ## -->
      <% If any_updSpel Then %>
        <div class="nf_minibox nf_green">
          <h4>Senast uppdaterade spelen</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_updSpel, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_updSpel(8, zx)) > 0 Then text_UseBox = list_updSpel(8, zx)
                If CLng(list_updSpel(7, zx)) > 0 Then text_UseBox = list_updSpel(7, zx)
                If CLng(list_updSpel(2, zx)) > 0 Then text_UseBox = list_updSpel(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/spel_lrg.png<% End If %>');" onclick="location.href='/avdelning/spel/spel_visa_info.asp?e=<% = list_updSpel(0, zx) %>';"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = list_updSpel(0, zx) %>" title="<% = sEncode(list_updSpel(1, zx)) %>"><% = sEncode(CutText(list_updSpel(1, zx), 32)) %></a><% = lstKonsol(list_updSpel(5, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/spel/">Visa alla spel</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /SENAST UPPDATERADE SPELEN ## -->
    
      <!-- ## SENASTE TEXTERNA ## -->
      <% If any_TextData Then %>
        <div class="nf_minibox">
          <h4>Senaste texterna</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_TextData, 2) %>
                <%
                Select Case CLng(list_TextData(6, zx))
                  Case 0 : sCatText = "nyheter/nyheter"
                  Case 1 : sCatText = "recensioner/recension"
                  Case 2 : sCatText = "artiklar/artikel"
                End Select
                %>
              
                <li onclick="location.href='/avdelning/<% = sCatText %>_visa.asp?e=<% = list_TextData(0, zx) %>';"><a href="/avdelning/<% = sCatText %>_visa.asp?e=<% = list_TextData(0, zx) %>" title="<% = sEncode(list_TextData(1, zx)) %>"><% = sEncode(CutText(list_TextData(1, zx), 32)) %></a><% = lstKategori(list_TextData(3, zx)) %> / <% = DatumReplace(list_TextData(2, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/texter/">Visa alla texter</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## SENASTE TEXTERNA ## -->
          
    </div>
      
    <div class="nf_datablock nf_size_onethird">
      
      <!-- ## SENAST UPPDATERADE KONSOLERNA ## -->
      <% If any_updKonsol Then %>
        <div class="nf_minibox nf_green">
          <h4>Senast uppdaterade konsolerna</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_updKonsol, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_updKonsol(8, zx)) > 0 Then text_UseBox = list_updKonsol(8, zx)
                If CLng(list_updKonsol(7, zx)) > 0 Then text_UseBox = list_updKonsol(7, zx)
                If CLng(list_updKonsol(2, zx)) > 0 Then text_UseBox = list_updKonsol(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/konsol_lrg.png<% End If %>');" onclick="location.href='/avdelning/konsol/konsol_visa_info.asp?e=<% = list_updKonsol(0, zx) %>';"><a href="/avdelning/konsol/konsol_visa_info.asp?e=<% = list_updKonsol(0, zx) %>" title="<% = sEncode(list_updKonsol(1, zx)) %>"><% = sEncode(CutText(list_updKonsol(1, zx), 32)) %></a><% = lstKonsol(list_updKonsol(5, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/konsol/">Visa alla konsoler</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /SENAST UPPDATERADE KONSOLERNA ## -->
      
      <!-- ## SENASTE FORUMINLÄGGEN ## -->
      <% If any_Tradar Then %>
        <div class="nf_minibox nf_blue">
          <h4>Senaste foruminläggen</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_Tradar, 2) %>
                <%
                  isTheThread = False
                  If list_Tradar(4,zx) Then isTheThread = True
                %>
                <li onclick="location.href='/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>';"><a href="/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>" title="<% = sEncode(list_Tradar(1,zx)) %>"><% = sEncode(CutText(list_Tradar(1,zx), 32)) %></a><% = list_Tradar(8, zx) %> / <% = DatumReplace(list_Tradar(3,zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/forum/nyainlagg.asp">Visa alla foruminlägg</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /SENASTE FORUMINLÄGGEN ## -->
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
     
      <!-- ## SENAST UPPDATERADE TILLBEHÖREN ## -->
      <% If any_updTillbehor Then %>
        <div class="nf_minibox nf_green">
          <h4>Senast uppdaterade tillbehören</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_updTillbehor, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_updTillbehor(8, zx)) > 0 Then text_UseBox = list_updTillbehor(8, zx)
                If CLng(list_updTillbehor(7, zx)) > 0 Then text_UseBox = list_updTillbehor(7, zx)
                If CLng(list_updTillbehor(2, zx)) > 0 Then text_UseBox = list_updTillbehor(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/tillbehor_lrg.png<% End If %>');" onclick="location.href='/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_updTillbehor(0, zx) %>';"><a href="/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_updTillbehor(0, zx) %>" title="<% = sEncode(list_updTillbehor(1, zx)) %>"><% = sEncode(CutText(list_updTillbehor(1, zx), 32)) %></a><% = lstKonsol(list_updTillbehor(5, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/tillbehor/">Visa alla tillbehör</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /SENAST UPPDATERADE TILLBEHÖREN ## -->
      
      <!-- ## SENASTE REGISTRERADE ## -->
      <% If any_On Then %>
        <div class="nf_minibox nf_red">
          <h4>Senaste registrerade medlemmarna</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_On, 2) %>
                <li onclick="location.href='/avdelning/medlem/?m=<% = sEncode(list_On(0,zx)) %>';"><a href="/avdelning/medlem/?m=<% = sEncode(list_On(0,zx)) %>" title="<% = sEncode(list_On(0,zx)) %>"><% = sEncode(CutText(list_On(0,zx), 32)) %></a> Registrerad: <% = DatumReplace(list_On(3,zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/listor/sokmedlem.asp">Sök medlem</a></p>
          </div>
        </div>
      <% End If %>
      <!-- ## /SENASTE REGISTRERADE ## -->
          
    </div>
    
  </div>

<!--#INCLUDE FILE="../_page_bottom.asp"-->
<!--#INCLUDE FILE="../__INC/includes_end.asp"-->