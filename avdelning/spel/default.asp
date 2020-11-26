<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  ' #### FILTER ####
    filter_alfa   = FixAlfaList(GetQ("alfa","ABC",0))
    filter_region = CLng(GetQ("region","123",0))
    filter_konsol = CLng(GetQ("k","123",0))
    
    alfa_SQL      = AlfaToSQL(filter_alfa, "tTitel")
    
    If filter_Region > 0 And filter_Region < 50 Then
      region_SQL = "AND tRegion = " & CLng(filter_Region)
    Else
      filter_Region = 0
    End If
    
    If filter_Konsol > 0 And filter_Konsol <= lstKonsol(0) Then
      konsol_SQL = "AND sKonsol = " & CLng(filter_Konsol)
      konsol_Add = lstKonsol(filter_konsol)
    Else
      filter_konsol = 0
      konsol_Add = "Alla"
    End If
    
    filter_All = ""        
  ' ################

  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0

  RS_Open 1, "SELECT sID, tTitel, sKonsol, tSortNo, tBoxart_BoxFram, tBoxart_Manual, tBoxart_Kassett, tRegion, tRelease, fUtgivare.fNamn, fUtvecklare.fNamn, fUtgivare.fID, fUtvecklare.fID, sSingleplayer, sMultiplayer, sOnline, sPEGI, sESRB, tID, " & _
             "(SELECT COUNT(biID) FROM cms_Bind_Anv_Spel WHERE biTitelID = cms_SpelTitlar.tID AND biAnv = " & CLng(lAnvID) & ") AS tListadAntal, sOlicensierad, tExtra " & _
             "FROM cms_SpelTitlar " & _ 
             "LEFT JOIN cms_Spel ON cms_SpelTitlar.tSpelID = cms_Spel.sID " & _ 
             "LEFT JOIN cms_Foretag AS fUtgivare ON cms_SpelTitlar.tUtgivare = fUtgivare.fID " & _ 
             "LEFT JOIN cms_Foretag AS fUtvecklare ON cms_Spel.sUtvecklare = fUtvecklare.fID " & _
             "WHERE sSynlig = 1 " & _
             alfa_SQL & _
             region_SQL & _
             konsol_SQL & _
             "ORDER BY tTitel ASC", False
  
    If rsDB(1).EOF Then
      any_Games = False
    Else
      any_Games = True
      list_Games = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  RS_Open 2, "SELECT rID, rNamn, rHighLight FROM cms_Region WHERE rHighLight = 1 ORDER BY rNamn ASC", False
  
    If rsDB(2).EOF Then
      any_Regions = False
    Else
      any_Regions = True
      list_Regions = rsDB(2).GetRows
    End If
  
  RS_Close 2
  
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
  
  ' ### Slumpade spel
  RS_Open 3, "SELECT TOP 9 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE sSynlig = 1 AND (tBoxart_BoxFram > 0 OR tBoxart_Manual > 0 OR tBoxart_Kassett > 0) " & _
             konsol_SQL & _
             "ORDER BY NewId()", False
  
    If rsDB(3).EOF Then
      any_Rnd  = False
    Else
      list_Rnd   = rsDB(3).GetRows
      any_Rnd     = True
    End If
  
  RS_Close 3
  
  If any_Games Then
    CreatePaging config_MaxAntalPosterPerSida, UBound(list_Games, 2)
    CreatePagingChooser
  End If
  
  If pagingOnPage < 1 Then pagingOnPage = 1
%>

<%
  ' ## Globala variabler ##
  If CLng(filter_region) > 0 Then text_Region = "utgivna i " & GetRegion(filter_region)
  If Len(filter_alfa) > 0 Then If UCase(filter_alfa) = "NUM" Then text_Alfa   = " - [ # ]" Else text_Alfa   = " - [ " & filter_alfa & " ]"
  
  page_Title    = konsol_Add & " spel " & text_Region & " " & text_Alfa & " - Sida " & pagingOnPage
  page_Header   = konsol_Add & " spel - Nintendo"
  page_WhereAmI = "&gt; Spel "
  page_SelMenu  = "databas"
  page_Slide    = "spel"
  
  page_description  = konsol_Add & " spel " & text_Region & " till Nintendo listade på N-Forum.se, Nintendo Forum. Sida " & pagingOnPage & ". " & text_Alfa
  page_keywords     = konsol_Add & " spel, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/spel/">Spel</a></span><% = konsol_Add %> Spel <% = text_Region %> <% = text_Alfa %></h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <div class="nf_alfa">
        <a href="?alfa=<% = filter_alfa %>&amp;k=<% = filter_konsol %>" <% If CLng(filter_region) = CLng(0) Then Response.Write(" class='c'") %>>Alla</a> |
        <% If any_Regions Then %>
          <% For zx = 0 To Ubound(list_Regions,2) %>
            <a href="?region=<% = list_Regions(0,zx) %><% = filter_All %>&amp;alfa=<% = filter_alfa %>&amp;k=<% = filter_konsol %>" <% If CLng(filter_region) = CLng(list_Regions(0,zx)) Then Response.Write(" class='c'") %>><% = list_Regions(1,zx) %></a> <% If zx <> Ubound(list_Regions,2) Then %> | <% End If %>
          <% Next %>
        <% End If %>

        <div style="width: 100%; height: 4px; overflow: hidden;"> </div>
        
        <a href="?alfa=<% = filter_alfa %>&amp;region=<% = filter_region %>" <% If CLng(filter_konsol) = CLng(0) Then Response.Write(" class='c'") %>>Alla</a> |
        <% For zx = 1 To lstKonsolSuperShort(0) %>
          <a href="?k=<% = zx %><% = filter_All %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %>" <% If CLng(filter_konsol) = CLng(zx) Then Response.Write(" class='c'") %> title="<% = lstKonsol(zx) %>"><% = lstKonsolSuperShort(zx) %></a> <% If zx <> lstKonsolSuperShort(0) Then %> | <% End If %>
        <% Next %>

        <div style="width: 100%; height: 4px; overflow: hidden;"> </div>
        
        <a href="?alfa=<% = filter_All %>&amp;k=<% = filter_konsol %>&amp;region=<% = filter_region %>" <% If filter_alfa = "" Then Response.Write(" class='c'") %>>Alla</a> |
        <a href="?alfa=num<% = filter_All %>&amp;k=<% = filter_konsol %>&amp;region=<% = filter_region %>" <% If filter_alfa = "NUM" Then Response.Write(" class='c'") %>>#</a> |
        <% For zx = 65 To 90 %>
          <a href="?alfa=<% = Chr(zx) %><% = filter_All %>&amp;k=<% = filter_konsol %>&amp;region=<% = filter_region %>" <% If filter_alfa = Chr(zx) Then Response.Write(" class='c'") %>><% = Chr(zx) %></a> <% If zx <> 90 Then %> | <% End If %>
        <% Next %>
      </div>
    
      <% If any_Games Then %> 
        
        <ul class="nf_list">
          <%
            For zx = pagingBOF To pagingEOF
              If zx > UBound(list_Games, 2) Then Exit For
              
              miniBox = 0
              If CLng(list_Games(5, zx)) > 0 Then miniBox = list_Games(5, zx)
              If CLng(list_Games(6, zx)) > 0 Then miniBox = list_Games(6, zx)
              If CLng(list_Games(4, zx)) > 0 Then miniBox = list_Games(4, zx)
              %>
          
                <li>
                  <div class="nf_tiny">
                    <% If CLng(miniBox) > 0 Then %>
                      <img src="<% = config_ImageLocation %>?e=<% = miniBox %>&amp;w=50&amp;h=50&amp;err=no">
                    <% Else %>
                      <img src="<% = config_GFXLocation %>img/noimg_24x24.gif">
                    <% End If %>
                  </div>
                  <div class="nf_data">
                    <h4>
                      <img src="<% = config_GFXLocation %>icons/flags/<% = CLng(list_Games(7, zx)) %>.png" alt="<% = lstRegion(CLng(list_Games(7, zx))) %>" title="Region: <% = lstRegion(CLng(list_Games(7, zx))) %>">
                      <a href="spel_visa_info.asp?e=<% = list_Games(18, zx) %>" title="<% = sEncode(list_Games(1, zx)) %>"><% = sEncode(list_Games(1, zx)) %></a>
                    </h4>
                    <span class="nf_medium nf_gray nf_bold"><% = lstKonsol(list_Games(2, zx)) %></span>
                  </div>
                  <div class="nf_extend">
                    <% If CONST_LOGIN Then %>
                      <img src="<% = config_GFXLocation %>icons/plus_lrg.png" style="float: right; cursor: pointer;" alt="+" title="Lägg till titeln i din samling." onclick="OpenCollection('game',<% = list_Games(18, zx) %>,0,'list')">
                    <% Else %>
                      <img src="<% = config_GFXLocation %>icons/plus_lrg_bw.png" style="float: right;" alt="+" title="Du måste vara inloggad för att kunna lista dina spel.">
                    <% End If %>
                    <% If list_Games(13, zx) Then %><img src="<% = config_GFXLocation %>icons/sp.gif" title="Spelet stödjer en spelare" alt="SP"><% End If %>
                    <% If list_Games(14, zx) Then %><img src="<% = config_GFXLocation %>icons/mp.gif" title="Spelet stödjer flera spelare" alt="MP"><% End If %>
                    <% If list_Games(15, zx) Then %><img src="<% = config_GFXLocation %>icons/wifi.gif" title="Spelet stödjer onlinespel" alt="WiFi"><% End If %>
                    <% If Not list_Games(20, zx) Then %><img src="<% = config_GFXLocation %>icons/license.gif" title="Spelet är licensierat av Nintendo" alt="Seal"><% End If %>
                    <% If CONST_LOGIN Then %><img src="<% = config_GFXLocation %>icons/listed.gif" style="display: <% If CLng(list_Games(19, zx)) = 0 Then Response.Write("none") Else Response.Write("block") %>;" id="listicon_<% = list_Games(18, zx) %>" alt="LIST" title="Titeln finns i din samling."><% End If %>
                  </div>
                </li>
              
              <%
            Next
          %>
        </ul>
        
        <div class="nf_paging">
          <a href="default.asp?page=<% = pagingOnPage - 1 %>&amp;k=<% = filter_konsol %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %><% = filter_all %>">««</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="default.asp?page=<% = zx %>&amp;k=<% = filter_konsol %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
          | <a href="default.asp?page=<% = pagingOnPage + 1 %>&amp;k=<% = filter_konsol %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %><% = filter_all %>">»»</a>
        </div>
      <% Else %>
        <div class="nf_msg nf_red">
          <p style="text-align: center;"><strong>Det finns inga spel att visa med aktuella val.</strong></p>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <!-- ### SÖKRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>Sök efter spel</h4>
        <div class="nf_inside">
          <form action="sokspel.asp" method="GET">
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
    
      <div class="nf_minibox nf_green">
        <h4>Sök spel på regionskod</h4>
        <div class="nf_inside">
          <input style="width: 254px; color: #AAA;" type="text" id="regkod" value="NES-XX-REG" onfocus="clearField(this,'NES-XX-REG');" onblur="retypeField(this,'NES-XX-REG');" onkeyup="searchRegCode('regkod');">
          
          <div id="soktraff" style="float: left; width: 261px;"></div>
        </div>
      </div>
      
      <div class="nf_minibox nf_blue">
        <h4>Ikonförklaring</h4>
        <div class="nf_inside">
          <p> <img src="<% = config_GFXLocation %>icons/sp.gif"> Enspelarläge tillgänligt </p>
          <p> <img src="<% = config_GFXLocation %>icons/mp.gif"> Flerspelarläge tillgängligt </p>
          <p> <img src="<% = config_GFXLocation %>icons/wifi.gif"> Möjlighet till onlinespel </p>
          <p> <img src="<% = config_GFXLocation %>icons/listed.gif"> Titeln finns i din samling </p>
          <p> <img src="<% = config_GFXLocation %>icons/license.gif"> Spelet är licensierat av Nintendo </p>
        </div>
      </div>
      
      <div class="nf_minibox nf_blue">
        <h4>Lägg till i din samling</h4>
        <div class="nf_inside">
          <% If CONST_LOGIN Then %>
            <p>För att lägga till ett spel i din samling klickar du bara på plusset till höger om titeln.</p>
            <p>Du kan lista samma titel flera gånger.</p>
          <% Else %>
            <p style="text-align: center;"><em>Du måste <strong><a href="/avdelning/medlem/loggain.asp">logga in</a></strong> för att kunna lista dina spel.</em></p>
            <p style="text-align: center;"><em>Om du inte redan har en användare kan du <strong><a href="/avdelning/medlem/registreradig.asp">registrera dig</a></strong>.</em></p>
          <% End If %>
        </div>
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
        
    </div>
    
    <div class="nf_datablock nf_size_full">
      
      <% If any_Rnd Then %>
        <div class="nf_images nf_images_full">
          <p><strong>Upptäck följande spel...</strong></p>
          <% For zx = 0 To UBound(list_Rnd, 2) %>
            <%
            text_UseBox = 0
            If CLng(list_Rnd(8, zx)) > 0 Then text_UseBox = list_Rnd(8, zx)
            If CLng(list_Rnd(7, zx)) > 0 Then text_UseBox = list_Rnd(7, zx)
            If CLng(list_Rnd(2, zx)) > 0 Then text_UseBox = list_Rnd(2, zx)
            %>
            <a href="spel_visa_info.asp?e=<% = list_Rnd(0,zx) %>"><img src="<% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=80&amp;h=80&amp;err=no" title="<% = sEncode(list_Rnd(1, zx)) %> (<% = list_Rnd(4, zx) %>)" alt="Spel"></a>
          <% Next %>
        </div>
      <% End If %>
      
    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->