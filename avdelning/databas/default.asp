<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  ' ### Fler spel
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE sSynlig = 1 " & _
             "ORDER BY tVisningar DESC", False
  
    If rsDB(1).EOF Then
      any_XSpel     = False
    Else
      any_XSpel     = True
      list_XSpel    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler konsoler
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, kKonsol, kTextM, tBoxart_Manual, tBoxart_Konsol FROM cms_KonsolTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Konsol ON kID = tKonsolID " & _
             "WHERE kSynlig = 1 " & _
             "ORDER BY tVisningar DESC", False
  
    If rsDB(1).EOF Then
      any_XKonsol     = False
    Else
      any_XKonsol     = True
      list_XKonsol    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler tillbehör
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, iKonsol, iTextM, tBoxart_Manual, tBoxart_Tillbehor FROM cms_TillbehorTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Tillbehor ON iID = tTillbehorID " & _
             "WHERE iSynlig = 1 " & _
             "ORDER BY tVisningar DESC", False
  
    If rsDB(1).EOF Then
      any_XTillbehor     = False
    Else
      any_XTillbehor     = True
      list_XTillbehor    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1

%>

<%
  ' ## Globala variabler ##
  page_Title    = "Databas"
  page_Header   = "Databas"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Recensioner&quot; ...'>Recensioner</a> "
  page_SelMenu  = "databas"
  page_Slide    = "spel"
  
  page_description    = "Ett urval av alla spel, konsoler och tillbehör på N-Forum.se, Nintendo Forum. Här kan du söka efter alla våra listade objekt i databasen."
  page_keywords       = "databas, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full" style="height: 12px;"> </div>
  
    <!-- ### SPEL ### -->
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_msg nf_bigbutton">
        <a href="/avdelning/spel/" title="Visa alla spel">Spel</a>
      </div>
      
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter spel</h4>
        <div class="nf_inside">
          <form action="/avdelning/spel/sokspel.asp" method="GET">
            <div class="nf_selfinput_full">
              <select name="konsol" style="width: 264px;">
                <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> Alla konsoler </option>
                <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                <% For zx = 1 To lstKonsol(0) %>
                  <option value=<% = zx %> style="padding: 1px 0 1px 10px;"> <% = lstKonsol(zx) %> </option>
                <% Next %>
              </select>
            </div>
            <div class="nf_selfinput_full"><input name="q" type="text" style="width: 260px;"></div>
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="Sök..."></div>
          </form>
        </div>
      </div>
      <!-- ### /SÖKRUTA ### -->
      
      <% If any_XSpel Then %>
        <div class="nf_minibox nf_blue">
          <h4>Spel</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XSpel, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_XSpel(8, zx)) > 0 Then text_UseBox = list_XSpel(8, zx)
                If CLng(list_XSpel(7, zx)) > 0 Then text_UseBox = list_XSpel(7, zx)
                If CLng(list_XSpel(2, zx)) > 0 Then text_UseBox = list_XSpel(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/spel_lrg.png<% End If %>');" onclick="location.href='/avdelning/spel/spel_visa_info.asp?e=<% = list_XSpel(0, zx) %>';"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = list_XSpel(0, zx) %>" title="<% = sEncode(list_XSpel(1, zx)) %>"><% = sEncode(CutText(list_XSpel(1, zx), 32)) %></a><% = lstKonsol(list_XSpel(5, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/spel/">Visa alla spel</a></p>
          </div>
        </div>
      <% End If %>
    </div>
    
    <!-- ### KONSOLER ### -->
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_msg nf_bigbutton">
        <a href="/avdelning/konsol/" title="Visa alla konsoler">Konsoler</a>
      </div>
      
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter konsol</h4>
        <div class="nf_inside">
          <form action="/avdelning/konsol/sokkonsol.asp" method="GET">
            <div class="nf_selfinput_full">
              <select name="konsol" style="width: 264px;">
                <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> Alla konsoler </option>
                <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                <% For zx = 1 To lstKonsol(0) %>
                  <option value=<% = zx %> style="padding: 1px 0 1px 10px;"> <% = lstKonsol(zx) %> </option>
                <% Next %>
              </select>
            </div>
            <div class="nf_selfinput_full"><input name="q" type="text" style="width: 260px;"></div>
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="Sök..."></div>
          </form>
        </div>
      </div>
      <!-- ### /SÖKRUTA ### -->
      
      <% If any_XKonsol Then %>
        <div class="nf_minibox nf_blue">
          <h4>Konsoler</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XKonsol, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_XKonsol(8, zx)) > 0 Then text_UseBox = list_XKonsol(8, zx)
                If CLng(list_XKonsol(7, zx)) > 0 Then text_UseBox = list_XKonsol(7, zx)
                If CLng(list_XKonsol(2, zx)) > 0 Then text_UseBox = list_XKonsol(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/konsol_lrg.png<% End If %>');" onclick="location.href='/avdelning/konsol/konsol_visa_info.asp?e=<% = list_XKonsol(0, zx) %>';"><a href="/avdelning/konsol/konsol_visa_info.asp?e=<% = list_XKonsol(0, zx) %>" title="<% = sEncode(list_XKonsol(1, zx)) %>"><% = sEncode(CutText(list_XKonsol(1, zx), 32)) %></a><% = lstKonsol(list_XKonsol(5, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/konsol/">Visa alla konsoler</a></p>
          </div>
        </div>
      <% End If %>
    </div>
    
    <!-- ### TILLBEHÖR ### -->
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_msg nf_bigbutton">
        <a href="/avdelning/tillbehor/" title="Visa alla tillbehör">Tillbehör</a>
      </div>
      
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter tillbehör</h4>
        <div class="nf_inside">
          <form action="/avdelning/tillbehor/soktillbehor.asp" method="GET">
            <div class="nf_selfinput_full">
              <select name="konsol" style="width: 264px;">
                <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> Alla konsoler </option>
                <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                <% For zx = 1 To lstKonsol(0) %>
                  <option value=<% = zx %> style="padding: 1px 0 1px 10px;"> <% = lstKonsol(zx) %> </option>
                <% Next %>
              </select>
            </div>
            <div class="nf_selfinput_full"><input name="q" type="text" style="width: 260px;"></div>
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="Sök..."></div>
          </form>
        </div>
      </div>
      <!-- ### /SÖKRUTA ### -->
      
      <% If any_XTillbehor Then %>
        <div class="nf_minibox nf_blue">
          <h4>Tillbehör</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XTillbehor, 2) %>
                <%
                text_UseBox = 0
                If CLng(list_XTillbehor(8, zx)) > 0 Then text_UseBox = list_XTillbehor(8, zx)
                If CLng(list_XTillbehor(7, zx)) > 0 Then text_UseBox = list_XTillbehor(7, zx)
                If CLng(list_XTillbehor(2, zx)) > 0 Then text_UseBox = list_XTillbehor(2, zx)
                %>
                <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/tillbehor_lrg.png<% End If %>');" onclick="location.href='/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_XTillbehor(0, zx) %>';"><a href="/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_XTillbehor(0, zx) %>" title="<% = sEncode(list_XTillbehor(1, zx)) %>"><% = sEncode(CutText(list_XTillbehor(1, zx), 32)) %></a><% = lstKonsol(list_XTillbehor(5, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/tillbehor/">Visa alla tillbehör</a></p>
          </div>
        </div>
      <% End If %>
    </div>
      
    <div class="nf_datablock nf_size_onethird">
      &nbsp;
    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->