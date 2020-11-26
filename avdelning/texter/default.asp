<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%

  ' ### Fler nyheter
  RS_Open 1, "SELECT TOP 10 nID, nTitel, nDatumPublicerad, nKategori FROM cms_Nyheter WHERE nDatumPublicerad <= '" & Now & "' AND nStatus = 4 ORDER BY nDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_XNews     = False
    Else
      any_XNews     = True
      list_XNews    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1

  ' ### Fler recensioner
  RS_Open 1, "SELECT TOP 10 rID, rTitel, rDatumPublicerad, rKategori FROM cms_Recensioner WHERE rDatumPublicerad <= '" & Now & "' AND rStatus = 4 ORDER BY rDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_XRec     = False
    Else
      any_XRec     = True
      list_XRec    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1
  
  ' ### Fler artiklar
  RS_Open 1, "SELECT TOP 10 aaID, aaTitel, aaDatumPublicerad, aaKategori FROM cms_Artiklar WHERE aaDatumPublicerad <= '" & Now & "' AND aaStatus = 4 ORDER BY aaDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_XArt     = False
    Else
      any_XArt     = True
      list_XArt    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1

%>

<%
  ' ## Globala variabler ##
  page_Title    = "Texter"
  page_Header   = "Texter"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Recensioner&quot; ...'>Recensioner</a> "
  page_SelMenu  = "texter"
  page_Slide    = "nyheter"
  
  page_description    = "Ett urval av alla nyheter, recensioner och artiklar på N-Forum.se, Nintendo Forum. Här kan du söka efter alla våra texter om nintendo."
  page_keywords       = "texter, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full" style="height: 12px;"> </div>
  
    <!-- ### NYHETER ### -->
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_msg nf_bigbutton">
        <a href="/avdelning/nyheter/" title="Visa alla nyheter">Nyheter</a>
      </div>
      
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter nyheter</h4>
        <div class="nf_inside">
          <form action="/avdelning/nyheter/default.asp" method="GET">
            <div class="nf_selfinput_full"><input name="q" type="text" style="width: 260px;"></div>
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="Sök..."></div>
          </form>
        </div>
      </div>
      <!-- ### /SÖKRUTA ### -->
      
      <% If any_XNews Then %>
        <div class="nf_minibox nf_blue">
          <h4>Nyheter</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XNews, 2) %>
                <li onclick="location.href='/avdelning/nyheter/nyheter_visa.asp?e=<% = list_XNews(0, zx) %>';"><a href="/avdelning/nyheter/nyheter_visa.asp?e=<% = list_XNews(0, zx) %>" title="<% = sEncode(list_XNews(1, zx)) %>"><% = sEncode(CutText(list_XNews(1, zx), 32)) %></a><% = lstKategori(list_XNews(3, zx)) %> / <% = DatumReplace(list_XNews(2, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/nyheter/">Visa alla nyheter</a></p>
          </div>
        </div>
      <% End If %>
    </div>
    
    <!-- ### RECENSIONER ### -->
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_msg nf_bigbutton">
        <a href="/avdelning/recensioner/" title="Visa alla recensioner">Recensioner</a>
      </div>
      
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter recensioner</h4>
        <div class="nf_inside">
          <form action="/avdelning/recensioner/default.asp" method="GET">
            <div class="nf_selfinput_full"><input name="q" type="text" style="width: 260px;"></div>
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="Sök..."></div>
          </form>
        </div>
      </div>
      <!-- ### /SÖKRUTA ### -->
      
      <% If any_XRec Then %>
        <div class="nf_minibox nf_blue">
          <h4>Recensioner</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XRec, 2) %>
                <li onclick="location.href='/avdelning/recensioner/recension_visa.asp?e=<% = list_XRec(0, zx) %>';"><a href="/avdelning/recensioner/recension_visa.asp?e=<% = list_XRec(0, zx) %>" title="<% = sEncode(list_XRec(1, zx)) %>"><% = sEncode(CutText(list_XRec(1, zx), 32)) %></a><% = lstKonsol(list_XRec(3, zx)) %> / <% = DatumReplace(list_XRec(2, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/recensioner/">Visa alla recensioner</a></p>
          </div>
        </div>
      <% End If %>
    </div>
    
    <!-- ### ARTIKLAR ### -->
    <div class="nf_datablock nf_size_onethird">
      <div class="nf_msg nf_bigbutton">
        <a href="/avdelning/artiklar/" title="Visa alla artiklar">Artiklar</a>
      </div>
      
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter artiklar</h4>
        <div class="nf_inside">
          <form action="/avdelning/artiklar/default.asp" method="GET">
            <div class="nf_selfinput_full"><input name="q" type="text" style="width: 260px;"></div>
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="Sök..."></div>
          </form>
        </div>
      </div>
      <!-- ### /SÖKRUTA ### -->
      
      <% If any_XArt Then %>
        <div class="nf_minibox nf_blue">
          <h4>Artiklar</h4>
          <div class="nf_inside nf_stylelist">
            <ul>
              <% For zx = 0 To UBound(list_XArt, 2) %>
                <li onclick="location.href='/avdelning/artiklar/artikel_visa.asp?e=<% = list_XArt(0, zx) %>';"><a href="/avdelning/artiklar/artikel_visa.asp?e=<% = list_XArt(0, zx) %>" title="<% = sEncode(list_XArt(1, zx)) %>"><% = sEncode(CutText(list_XArt(1, zx), 32)) %></a><% = lstKonsol(list_XArt(3, zx)) %> / <% = DatumReplace(list_XArt(2, zx)) %></li>
              <% Next %>
            </ul>
            <p><a href="/avdelning/artiklar/">Visa alla artiklar</a></p>
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