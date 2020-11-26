<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  ' ## Hämta all data ##
  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0
  
  lID = GetQ("e", "123", 0)
  lAG = GetF("vAldersgrans", "123", 0)
  RS_Open 1, "SELECT *, utg.fNamn AS iUtgivare, utg.fID AS iUtgivareID, utv.fNamn AS iUtvecklare, utv.fID AS iUtvecklareID, " & _
             "(SELECT COUNT(biID) FROM cms_Bind_Anv_Spel WHERE biTitelID = cms_SpelTitlar.tID AND biAnv = " & CLng(lAnvID) & ") AS tListadAntal " & _
             "FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Spel_Alder ON agSpelID = tSpelID " & _ 
             "LEFT JOIN cms_Spel ON cms_SpelTitlar.tSpelID = cms_Spel.sID " & _
             "LEFT JOIN cms_Foretag AS utg ON cms_Speltitlar.tUtgivare = utg.fID " & _
             "LEFT JOIN cms_Foretag AS utv ON cms_Spel.sUtvecklare = utv.fID " & _
             "WHERE tID = " & CLng(lID), False
  
    If rsDB(1).EOF Then Response.Redirect("default.asp")
    
    text_ID         = CLng(rsDB(1)("tSpelID"))
    
    ' ### Sätt åldersgräns
    If lAG > 5 Then lAG = 0
    If CONST_LOGIN And lAG <> 0 Then
      RS_Open 2, "SELECT * FROM cms_Spel_Alder WHERE agSpelID = " & CLng(text_ID) & " AND agAnv = " & CLng(CONST_USERID), True
        If rsDB(2).EOF Then
          noDelete = True
        
          If lAG > 0 Then
            rsDB(2).AddNew
            rsDB(2)("agAnv")    = CLng(CONST_USERID)
            rsDB(2)("agSpelID") = CLng(text_ID)
          End If
        End If
        
        If lAG > 0 Then
          rsDB(2)("agDatum") = Now
          rsDB(2)("agAlder") = lAG
          
          rsDB(2).Update
        Else
          If Not noDelete Then rsDB(2).Delete  
        End If
      RS_Close 2
      
      Response.Redirect("spel_visa_info.asp?e=" & CLng(rsDB(1)("tID")))
    End If
    ' ### ^
    
    text_Titel      = sEncode(rsDB(1)("tTitel"))
    text_TitelRaw   = rsDB(1)("tTitel")
    text_Region     = FixNum(rsDB(1)("tRegion"))
    text_ExtraInfo  = sEncode(rsDB(1)("tExtra"))
    text_RegKod     = sEncode(rsDB(1)("tRegionsKod"))
    text_Release    = sEncode(rsDB(1)("tRelease"))
    text_Konsol     = lstKonsol(rsDB(1)("sKonsol"))
    text_KonsolID   = rsDB(1)("sKonsol")
    text_SpelID     = FixNum(rsDB(1)("tSpelID"))
    
    If Not IsNull(rsDB(1)("agAlder")) Then text_MinAlder   = CLng(rsDB(1)("agAlder"))
    
    text_Utgivare   = sEncode(rsDB(1)("iUtgivare"))
    text_UtgivareID = FixNum((rsDB(1)("iUtgivareID")))
    text_Utvecklare = sEncode(rsDB(1)("iUtvecklare"))
    text_UtvecklareID = FixNum(rsDB(1)("iUtvecklareID"))
    
    text_ESRB       = CLng(rsDB(1)("sESRB"))
    text_PEGI       = CLng(rsDB(1)("sPEGI"))
    
    text_ListadAntal  = CLng(rsDB(1)("tListadAntal"))
    text_SinglePlayer = rsDB(1)("sSinglePlayer")
    text_Multiplayer  = rsDB(1)("sMultiplayer")
    text_Spelare      = CLng(rsDB(1)("sAntalSpelare"))
    text_Online       = rsDB(1)("sOnline")
    text_License      = rsDB(1)("sOlicensierad")
    
    text_LargeText  = Trim(BBCode(sEncode(rsDB(1)("sTextM")), False))
    text_LargeTextE = Trim(sEncode(CutText(rsDB(1)("sTextM"), 150)))
      
    text_Img1       = CLng(rsDB(1)("tBoxart_BoxFram"))
    text_Img2       = CLng(rsDB(1)("tBoxart_BoxBak"))
    text_Img3       = CLng(rsDB(1)("tBoxart_Manual"))
    text_Img4       = CLng(rsDB(1)("tBoxart_Kassett"))
    
    If text_Img4 > 0 Then text_UseArt = text_Img4 : text_UseText = "Kassett/Media"
    If text_Img3 > 0 Then text_UseArt = text_Img3 : text_UseText = "Manual"
    If text_Img2 > 0 Then text_UseArt = text_Img2 : text_UseText = "Boxart - Baksida"
    If text_Img1 > 0 Then text_UseArt = text_Img1 : text_UseText = "Boxart - Framsida"
  
  RS_Close 1
  
  If CONST_LOGIN Then
    RS_Open 1, "SELECT * FROM cms_SpelBetyg WHERE bAnv = " & CLng(lAnvID) & " AND bSpelID = " & CLng(text_ID), False
      If rsDB(1).EOF Then
        text_Betyg = 0
      Else
        text_Betyg = rsDB(1)("bBetyg")
      End If
    RS_Close 1
  End If
  
  lAntalBetyg = Con.ExeCute("SELECT COUNT(bID) FROM cms_SpelBetyg WHERE bSpelID = " & CLng(text_ID))(0)
  If lAntalBetyg > 0 Then
    lSumBetyg = Con.ExeCute("SELECT SUM(bBetyg) FROM cms_SpelBetyg WHERE bSpelID = " & CLng(text_ID))(0)
    
    lBetyg = Round(lSumBetyg / lAntalBetyg)
    If lBetyg < 1 Then lBetyg = 1
    If lBetyg > 6 Then lBetyg = 6
    
    text_MBetyg = lBetyg
  Else
    text_MBetyg = 0
  End if
  
  RS_Open 1, "SELECT tID, tRegion, tTitel, tRelease, tExtra FROM cms_SpelTitlar LEFT JOIN cms_Spel ON cms_SpelTitlar.tSpelID = cms_Spel.sID WHERE tSpelID = " & CLng(text_ID) & " ORDER BY tRelease ASC", False
  
    If rsDB(1).EOF Then
      any_Titles = False
    Else
      any_Titles = True
      list_Titles = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  RS_Open 1, "SELECT gID, gNamn FROM cms_Spelgengre WHERE gID IN(SELECT bgGenre FROM cms_Bind_Spel_Genre WHERE bgSpel = " & CLng(text_ID) & ") ORDER BY gNamn ASC", False
  
    If rsDB(1).EOF Then
      any_Genres = False
    Else
      any_Genres = True
      list_Genres = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  RS_Open 1, "SELECT ssID, ssNamn FROM cms_Spelserier WHERE ssID IN(SELECT bsSpelSerie FROM cms_Bind_Spel_Spelserie WHERE bsSpel = " & CLng(text_ID) & ") ORDER BY ssNamn ASC", False
  
    If rsDB(1).EOF Then
      any_Categories = False
    Else
      any_Categories = True
      list_Categories = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  ' #### HÄMTA RECENSIONER ####
  RS_Open 1, "SELECT TOP 10 rID, rTitel, rDatumPublicerad, aAnvNamn, rBetyg FROM cms_Recensioner LEFT JOIN fsBB_Anv ON aID = rSkapadAv WHERE rStatus = 4 AND rSpelID = " & CLng(text_ID) & " ORDER BY rAnvandarRec ASC, rDatumPublicerad DESC", False
  
    If rsDB(1).EOF Then
      any_Rec = False
    Else
      any_Rec = True
      list_Rec = rsDB(1).GetRows
    End If
  
  RS_Close 1
  ' ########################
  
  ' #### HÄMTA LIKNANDE SPEL ####
    ' #### FIX TEXT STRÄNG ####
      q = LCase(Trim(text_TitelRaw))
      
      q = MakeLegal(q)
      w = Split(q, " ")
      
      For Each ww In w
        ww = Trim(ww)
        ww = Replace(ww, ":", "")
        ww = Replace(ww, "'", "")
        ww = Replace(ww, "the", "")
        If IsNumeric(ww) Then If ww > 1979 And ww < 2050 Then ww = ""
        If IsNumeric(ww) Then If ww > 1979 And ww < 2050 Then ww = ""
        
        If Len(ww) > 2 Then
          p = p & """" & ww & """ OR "
        End If
      Next
  
      If Len(p) > 3 Then
        p = Left(p, Len(p)-4)
        p = "'(" & p & ")'"
  
        sSQL =     "SELECT TOP 9 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, Rank, tBoxart_Manual, tBoxart_Kassett " & _
                   "FROM cms_SpelTitlar " & _
                   "LEFT JOIN CONTAINSTABLE(cms_SpelTitlar, *, " & p & ") AS ct ON tID = ct.[KEY] " & _
                   "LEFT JOIN cms_Spel ON sID = tSpelID " & _
                   "LEFT JOIN cms_Region ON tRegion = rID " & _
                   "WHERE Rank > 0 AND sSynlig = 1 AND sID <> " & CLng(text_SpelID) & " AND (tBoxart_BoxFram > 0 OR tBoxart_Manual > 0 OR tBoxart_Kassett > 0) " & _
                   "ORDER BY Rank DESC"
      
        RS_Open 1, sSQL, False
        
          If rsDB(1).EOF Then
            any_Same = False
          Else
            any_Same  = True
            list_Same = rsDB(1).GetRows(10)
            text_Reko = "Rekommenderade spel..."
          End If
        
        RS_Close 1
      End If
    ' ##########################
    
    If Not any_Same Then
      RS_Open 3, "SELECT TOP 9 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tSpelID, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
                 "LEFT JOIN cms_Region ON tRegion = rID " & _
                 "LEFT JOIN cms_Spel ON sID = tSpelID " & _
                 "WHERE sID <> " & CLng(text_SpelID) & " AND sSynlig = 1 AND (tBoxart_BoxFram > 0 OR tBoxart_Manual > 0 OR tBoxart_Kassett > 0) " & _
                 "ORDER BY NewId()", False
      
        If rsDB(3).EOF Then
          any_Same  = False
        Else
          list_Same   = rsDB(3).GetRows
          any_Same    = True
          text_Reko   = "Upptäck följande spel..."
        End If
      
      RS_Close 3
    End If
  ' #############################
  
  ' #### HÄMTA BILDER ####
    RS_Open 1, "SELECT bID, bsSpel, bsBildText, bsBild, bsID FROM cms_Bind_Spel_Img LEFT JOIN cms_Bild ON cms_Bind_Spel_Img.bsBild = cms_Bild.bID WHERE bsSpel = " & CLng(text_SpelID) & " ORDER BY bsID ASC", False
  
      If rsDB(1).EOF Then
        any_Images = False
      Else
        any_Images = True
        list_Images = rsDB(1).GetRows
      End If
    
    RS_Close 1
  ' ######################
  
  ' #### HÄMTA TITLAR I SAMLINGEN ####
    If CONST_LOGIN Then
      RS_Open 1, "SELECT biID, biTitelID, biBox, biManual, biMedia, biExtra, biInPris, tTitel, tRegion, tExtra FROM cms_Bind_Anv_Spel LEFT JOIN cms_Speltitlar ON cms_Bind_Anv_Spel.biTitelID = cms_Speltitlar.tID WHERE biAnv = " & lAnvID & " AND biSpelID = " & CLng(text_ID), False
      
        If rsDB(1).EOF Then
          any_Samling = False
        Else
          any_Samling = True
          list_Samling  = rsDB(1).GetRows
        End If
      
      RS_Close 1
    End If
  ' ######################
  
  ' #### HÄMTA ALLA SOM SAMLAR ####
    RS_Open 1, "SELECT aAnvNamn, tRegion, biBox, biManual, biMedia, biExtra, biOvrigt FROM fsBB_Anv " & _
               "LEFT JOIN cms_Bind_Anv_Spel ON cms_Bind_Anv_Spel.biAnv = aID " & _
               "LEFT JOIN cms_SpelTitlar ON cms_SpelTitlar.tID = cms_Bind_Anv_Spel.biTitelID " & _
               "WHERE biTitelID IN (SELECT tID FROM cms_SpelTitlar WHERE tSpelID = " & CLng(text_SpelID) & ") AND aBlockadTill < '" & Date & "' AND aAktiverad = 1 AND aID <> " & CLng(lAnvID) & " ORDER BY biBox DESC, biMedia DESC, biManual DESC, biExtra DESC, aAnvNamn ASC", False
    
      If rsDB(1).EOF Then
        any_SomHar = False
      Else
        any_SomHar = True
        list_SomHar = rsDB(1).GetRows
      End If
    
    RS_Close 1
  ' ######################
  
  ' ### Fler foruminlägg
  If Not config_LockDown_Forum Then
    ' #### FIX TEXT STRÄNG ####
      p = ""
      q = ""
      ww = ""
    
      q = LCase(Trim(text_TitelRaw))
      
      q = MakeLegal(q)
      w = Split(q, " ")
      
      For Each ww In w
        ww = Trim(ww)
        
        If Len(ww) > 2 Then
          Select Case ww
            Case Else : p = p & """*" & ww & "*"" AND "
          End Select
        End If
      Next
      
      p = Left(p, Len(p)-5)
    ' #### ^
  
    RS_Open 2, "SELECT TOP 8 tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
               "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, fIcon, fName, Rank " & _
               "FROM fsBB_Tradar AS tbTrad " & _
               "LEFT JOIN CONTAINSTABLE(fsBB_Tradar, tTextM, '" & p & "') AS ct ON tbTrad.tID = ct.[KEY] " & _
               "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
               "WHERE Rank > 0 AND tDatum_Skapad <= '" & Now & "' AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tStatus_Raderad = 0" & _
               "ORDER BY Rank DESC, tAmne ASC", False
    
      If rsDB(2).EOF Then
        any_Tradar = False
      Else
        any_Tradar = True
        list_Tradar = rsDB(2).GetRows(8)
      End If
    
    RS_Close 2
  End If
  
  If CLng(Session.Value("LastGameSeen")) <> CLng(text_SpelID) Then
    Con.ExeCute("UPDATE cms_Speltitlar SET tVisningar = tVisningar + 1 WHERE tID = " & CLng(lID))
    Session.Value("LastGameSeen") = CLng(text_SpelID)
  End If
  
  bNoAllInfo = False
  If Len(text_RegKod) < 1 And Len(text_Release) < 1 Then bNoAllInfo = True ' Regionskod, Releasedatum
  If Len(text_Utgivare) < 1 And Len(text_Utvecklare) < 1 Then bNoAllInfo = True ' Utgivare och Utvecklare
  If text_Img1 = 0 And text_Img2 = 0 And text_Img3 = 0 And text_Img4 = 0 Then bNoAllInfo = True ' Boxart
%>

<%
  ' ## Globala variabler ##
  If Len(text_RegKod) > 0 Then
    page_Title    = text_Titel & " - " & UCase(text_RegKod) & " - Information - Spel"
    page_description  = "Visar " & text_Titel & " till " & text_Konsol & " utgiven i " & GetRegion(text_Region) & " på N-Forum.se, Nintendo Forum. Spelet har regionskoden " & UCase(text_RegKod) & "."
  Else
    page_Title    = text_Titel & " - Information - Spel"
    page_description  = "Visar " & text_Titel & " till " & text_Konsol & " utgiven i " & GetRegion(text_Region) & " på N-Forum.se, Nintendo Forum."
  End If
  
  page_Header   = text_Titel
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Spel&quot; ...'>Spel</a> " & _
                  "&gt; <a href='spel_visa_info.asp?e=" & lID & "'>" & text_Titel & "</a> " & _
                  "&gt; Information"
  page_SelMenu  = "databas"
  page_Slide    = "spel"
  
  page_keywords     = text_Titel & ", "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/spel/">Spel</a></span><img src="<% = config_GFXLocation %>icons/flags/<% = text_Region %>.png" alt="" title=""> <% = text_Titel %> <% If Len(text_ExtraInfo) > 0 Then %> <span style="color: #AAA;">- <% = text_ExtraInfo %></span><% End If %></h1>
      <h4><a href="default.asp?k=<% = text_KonsolID %>"><% = text_Konsol %></a></h4>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
      
      <% If any_Titles Then %>
        <div class="nf_msg">
          <p><strong>Utgivet i följande regioner...</strong></p>
          <ul class="nf_rowlist">
            <% For zx = 0 To UBound(list_Titles, 2) %>
              <li onclick="location.href='spel_visa_info.asp?e=<% = list_Titles(0, zx) %>'" <% If CLng(list_Titles(0, zx)) = CLng(lID) Then Response.Write(" class='c'") %>>
                <img src="<% = config_GFXLocation %>icons/flags/<% = list_Titles(1, zx) %>.png" alt="" title="">
                <a href="spel_visa_info.asp?e=<% = list_Titles(0, zx) %>" title="<% = sEncode(list_Titles(2, zx)) %>"><% = sEncode(CutText(list_Titles(2, zx), 65)) %></a> <% If Len(list_Titles(4, zx)) > 0 Then Response.Write("<span> - " & sEncode(list_Titles(4, zx)) & "</span>") %>
                <span style="float: right;"><% = sEncode(list_Titles(3, zx)) %></span>
              </li>
            <% Next %>
          </ul>
        </div>
      <% End If %>
      
      <% If text_UseArt > 0 Then %>
        <div class="nf_images">
          <div class="boxartblocker"></div>
          
          <% If text_Img1 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img1 %>&amp;w=640&amp;h=480" title="Boxart - Framsida" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img1 %>&amp;w=100&amp;h=100" title="Boxart - Framsida" alt="Boxart - Framsida">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <% If text_Img2 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img2 %>&amp;w=640&amp;h=480" title="Boxart - Baksida" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img2 %>&amp;w=100&amp;h=100" title="Boxart - Baksida" alt="Boxart - Baksida">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <% If text_Img3 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img3 %>&amp;w=640&amp;h=480" title="Manual" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img3 %>&amp;w=100&amp;h=100" title="Manual" alt="Manual">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <% If text_Img4 > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_Img4 %>&amp;w=640&amp;h=480" title="Kassett/Media" target="_blank" rel="lightbox[boxart]"><% End If %>
              <img class="boxart" src="<% = config_ImageLocation %>?e=<% = text_Img4 %>&amp;w=100&amp;h=100" title="Kassett/Media" alt="Kassett/Media">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% End If %>
          
          <div class="boxartblocker"></div>
        </div>
      <% End If %>
      
      <div class="nf_text">
        <p><strong>Information om <% = text_Titel %> <% If Len(text_ExtraInfo) > 0 Then %> <span style="color: #AAA;">- <% = text_ExtraInfo %></span><% End If %></strong></p>
        <% If Len(Trim(text_LargeText)) > 0 Then %>
          <p><% = text_LargeText %></p>
        <% Else %>
          <p><% = text_Titel %> <% If any_Genres Then %> är ett <% For zx = 0 To UBound(list_Genres, 2) %><% If zx > 0 And zx < UBound(list_Genres, 2) Then %>, <% End If %><% If zx > 0 And zx = UBound(list_Genres, 2) Then %> och <% End If %><% = sEncode(list_Genres(1, zx)) %><% Next %>-spel<% End If %>. Det är utgivet <strong><% = text_Release %></strong><% If Len(text_Utgivare) > 0 Then %> av utgivaren <strong><% = text_Utgivare %></strong><% End if %> <% If Len(text_Utvecklare) > 0 Then %> och är utvecklat av <strong><% = text_Utvecklare %></strong><% End If %>.</p>
          <p>Denna version av spelet är till <strong><a href="default.asp?k=<% = text_KonsolID %>"><% = text_Konsol %></a></strong> och man kan vara <strong><% If text_Spelare > 4 Then %>Fler än 4 spelare<% Else %><% = text_Spelare %> spelare<% End If %></strong> i detta spel. Det går <% If text_Online Then %>även<% Else %>inte<% End If %> att spela online.</p>
          <p>Det är <% If text_License Then %>inte<% End If %> licensierat av Nintendo.</p>
          <p style="color: #A00;"><strong>Hjälp oss!</strong></p> 
          <p>Som du ser har vi inte all information om detta spel, men du kan hjälpa oss genom att gå till <a href="/avdelning/forum/">forumet</a> och delge oss den information som saknas om spelet eller helt enkelt skriva en <a href="ny_recension.asp?e=<% = lID %>">recension</a> för att andra ska kunna få en uppfattning om detta spel.</p>
          <p>Om du själv dessutom är ägare av spelet får du gärna scanna in boxarten och sedan skicka den till oss, antingen via <a href="/avdelning/forum/">forumet</a> eller per mail till <a href="mailto:info@n-forum.se">info@n-forum.se</a>.</p>
        <% End If %>
      </div>
      
      <% If any_Images Then %>
        <div class="nf_images">
          <p><strong>Bilder från spelet...</strong></p>
          <% For zx = 0 To UBound(list_Images, 2) %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = list_Images(0, zx) %>&amp;w=640&amp;h=480" rel="lightbox[bilder]" title="<% = sEncode(list_Images(2, zx)) %>" target="_blank"><% End If %>
              <img src="<% = config_ImageLocation %>?e=<% = list_Images(0, zx) %>&amp;w=80&amp;h=80" title="<% = sEncode(list_Images(2, zx)) %>" alt="<% = sEncode(list_Images(2, zx)) %>">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% Next %>
        </div>
      <% End If %>
      
      <% If CONST_LOGIN Then %>
        <% ' #### SAMLINGEN #### %>
        <div class="nf_msg nf_green">
          <p><strong>Du har listat följande kopior av spelet...</strong></p>
          <ul class="nf_rowlist" id="titleListed_List" style="<% If Not any_Samling Then Response.Write("display: none;") %>"></ul>   
          <p class="nf_pretend_rowlist" id="titleListed_Mess" style="<% If any_Samling Then Response.Write("display: none;") %>">Du har inte listat spelet.</p>
          <p style="text-align: center;"><input style="float: none;" type="button" onclick="OpenCollection('game',<% = lID %>,0,'new');" value="Lägg till i samlingen"></p>
          <% If any_SomHar Then %>
            <p><strong><img id="toggleBt" src="<% = config_GFXLocation %>icons/plus.gif" onclick="toggleBox('listadav','toggleBt');" style="float: left; cursor: pointer; margin: 0 5px 0 0;"> <span style="float: left; margin: 1px 0 0 0;">Listat av följande medlemmar...</span> </strong></p>
            <ul class="nf_rowlist" id="listadav" style="display: none;">
              <% For zx = 0 To UBound(list_SomHar, 2) %>
                <li>
                  <img src="<% = config_GFXLocation %>icons/flags/<% = list_SomHar(1,zx) %>.png" alt="" title="">
                  <a href="/avdelning/medlem/?m=<% = sEncode(list_SomHar(0,zx)) %>"><% = sEncode(list_SomHar(0,zx)) %></a>
                  <div class="nf_collectionbar" style="background-image: url('<% = config_GFXLocation %>icons/samling/samling_alla_spel.png');">
                    <img alt="" title="Box" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(2,zx) Then Response.Write("blank") %>.png">
                    <img alt="" title="Media" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(4,zx) Then Response.Write("blank") %>.png">
                    <img alt="" title="Manual" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(3,zx) Then Response.Write("blank") %>.png"> 
                    <img alt="" title="Extra" src="<% = config_GFXLocation %>icons/samling/no<% If list_SomHar(5,zx) Then Response.Write("blank") %>.png"> 
                  </div>
                  <% If Len(list_SomHar(6,zx)) > 1 Then %><span style="float: right; margin: 0 5px 0 5px;" title="<% = sEncode(list_SomHar(6,zx)) %>"><% = sEncode(CutText(list_SomHar(6,zx), 50)) %></span> <% End If %>
                </li>
              <% Next %>
            </ul>
          <% Else %>
            <p><strong>Inga andra medlemmar har listat titeln. </strong></p>
          <% End If %>
        </div>
        
        <div id="titleListed_Clone" style="display: none;">
          <img src="<% = config_GFXLocation %>icons/flags/XXXX_REGION.png" alt="" title="">
          <a href="spel_visa_info.asp?e=XXXX_GAMEID" title="XXXX_GAME">XXXX_CUTGAME</a>
          <span style="float: right;">
            <img src="<% = config_GFXLocation %>icons/redigera.gif" alt="R" title="Redigera" title="" onclick="OpenCollection('game',XXXX_GAMEID,XXXX_POSTID,'edit');">
            <img src="<% = config_GFXLocation %>icons/radera.gif" alt="X" title="Radera" onclick="DeleteCollection('game',XXXX_POSTID);">
          </span>
          <div class="nf_collectionbar" style="background-image: url('<% = config_GFXLocation %>icons/samling/samling_alla_spel.png');">
            <img alt="" title="Box" src="<% = config_GFXLocation %>icons/samling/noXXXX_CBOX.png">
            <img alt="" title="Media" src="<% = config_GFXLocation %>icons/samling/noXXXX_CMEDIA.png">
            <img alt="" title="Manual" src="<% = config_GFXLocation %>icons/samling/noXXXX_CMANUAL.png"> 
            <img alt="" title="Extra" src="<% = config_GFXLocation %>icons/samling/noXXXX_CEXTRA.png"> 
          </div>
        </div>
        
        <script type="text/javascript">
          <% If any_Samling Then %>
            <% For zx = 0 To UBound(list_Samling, 2) %>
              <% titleTT = sEncode(CutText(list_Samling(7, zx), 65)) & "</a>" %>
              <% If Len(list_Samling(9, zx)) > 0 Then titleTT = titleTT & "<span> - " & sEncode(list_Samling(9, zx)) & "</span>" %>
              <% If list_Samling(2, zx) = True Then cBox = "blank" Else cBox = "" %>
              <% If list_Samling(4, zx) = True Then cMedia = "blank" Else cMedia = "" %>
              <% If list_Samling(3, zx) = True Then cManual = "blank" Else cManual = "" %>
              <% If list_Samling(5, zx) = True Then cExtra = "blank" Else cExtra = "" %>
              rh_cloneRow("titleListed_Clone", "titleListed_List", "titleListed_Row_", <% = list_Samling(0, zx) %>, "LI","REGION==<% = list_Samling(8, zx) %>;;GAMEID==<% = list_Samling(1, zx) %>;;GAME==<% = sEncode(list_Samling(7, zx)) %>;;CUTGAME==<% = titleTT %>;;POSTID==<% = list_Samling(0, zx) %>;;CBOX==<% = cBox %>;;CMEDIA==<% = cMedia %>;;CMANUAL==<% = cManual %>;;CEXTRA==<% = cExtra %>");
            <% Next %>
          <% End If %>
        </script>
        <% ' #### /SAMLINGEN #### %>
      <% Else %>
        <% ' #### SAMLINGEN #### %>
        <div class="nf_msg nf_green">
          <p style="text-align: center;">Du måste <strong><a href="/avdelning/medlem/loggain.asp">logga in</a></strong> för att kunna lista dina spel.</p>
          <p style="text-align: center;">Om du inte redan har en användare kan du <strong><a href="/avdelning/medlem/registreradig.asp">bli medlem</a> GRATIS</strong>.</p>
        </div>
        <% ' #### /SAMLINGEN #### %>
      <% End If %>
      
      <% ' #### RECENSIONER #### %>
        <div class="nf_msg nf_blue">
          <p><strong>Recensioner...</strong></p>
          <% If any_Rec Then %>
            <ul class="nf_rowlist">
              <% For zx = 0 To UBound(list_Rec, 2) %>
                <li>
                  <a href="/avdelning/recensioner/recension_visa.asp?e=<% = list_Rec(0, zx) %>"><% = sEncode(CutText(list_Rec(1, zx),65)) %></a>
                  <span style="float: right; color: #000; margin-left: 9px;"><strong><% = list_Rec(4, zx) %> / 10</strong></span>
                  <span style="float: right; margin-left: 5px;"><% = DatumReplace(list_Rec(2, zx)) %></span>
                  <span style="float: right; margin-left: 5px;"><a href="/avdelning/medlem/?m=<% = sEncode(list_Rec(3, zx)) %>"><% = sEncode(list_Rec(3, zx)) %></a></span>
                </li>
              <% Next %>
            </ul>
          <% Else %>
            <p class="nf_pretend_rowlist">Det finns inga recensioner.</p>
          <% End If %>
          <p style="text-align: center;"><input style="float: none;" type="button" onclick="location.href='ny_recension.asp?e=<% = lID %>';" value="Skriv en recension"></p>
        </div>
      <% ' #### /RECENSIONER #### %>
    </div>
    
    <div class="nf_datablock nf_size_onethird">
    
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <div class="nf_minibox nf_blue">
        <div class="nf_inside nf_boxart">
          <% If CLng(text_UseArt) > 0 Then %>
            <% If CONST_LOGIN Then %><a href="<% = config_ImageLocation %>?e=<% = text_UseArt %>&amp;w=800&amp;h=600" rel="lightbox" target="_blank" title="<% = text_UseText %>"><% End If %>
              <img src="<% = config_ImageLocation %>?e=<% = text_UseArt %>&amp;w=300&amp;h=300">
            <% If CONST_LOGIN Then %></a><% End If %>
          <% Else %>
            <img src="<% = config_GFXLocation %>img/noimg_200x150.png">        
          <% End If %>
        </div>
      </div>
      
      <div class="nf_minibox nf_blue">
        <div class="nf_inside">
          <% If text_SinglePlayer Then %><img src="<% = config_GFXLocation %>icons/sp.gif" title="Spelet stödjer en spelare" alt="SP"><% End If %>
          <% If text_Multiplayer Then %><img src="<% = config_GFXLocation %>icons/mp.gif" title="Spelet stödjer flera spelare" alt="MP"><% End If %>
          <% If text_Online Then %><img src="<% = config_GFXLocation %>icons/wifi.gif" title="Spelet stödjer onlinespel" alt="WiFi"><% End If %>
          <% If Not text_License Then %><img src="<% = config_GFXLocation %>icons/license.gif" title="Spelet är licensierat av Nintendo" alt="Seal"><% End If %>
        </div>
      </div>
      
      <div class="nf_minibox">
        <h4>Dela med dig</h4>
        <div class="nf_inside">
          <!-- AddThis Button BEGIN -->
            <div class="addthis_toolbox addthis_default_style" addthis:title="<% = text_Titel %>" addthis:description="<% = text_LargeTextE %>">
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
      
      <% If Len(text_RegKod) > 0 Then %>
        <div class="nf_minibox nf_blue">
          <h4>Regionskod</h4>
          <div class="nf_inside">
            <p class="nf_huge nf_center"><strong><% = UCase(text_RegKod) %></strong></p>
          </div>
        </div>
      <% End If %>
      
      <div class="nf_minibox">
        <h4>Betyg</h4>
        <div class="nf_inside nf_grades">
          <% If CONST_LOGIN Then %>
            <p style="text-align: center; margin-bottom: 12px;"><em>Klicka på valfri stjärna för att ange ditt betyg.</em></p>
          
            <span>Ditt betyg</span>
            <img src="<% = config_GFXLocation %>icons/grade_no.gif" id="bg6" onmouseover="showGrade(6);" onmouseout="showGrade(0);" onclick="setGrade(<% = text_ID %>,6);">
            <img src="<% = config_GFXLocation %>icons/grade_no.gif" id="bg5" onmouseover="showGrade(5);" onmouseout="showGrade(0);" onclick="setGrade(<% = text_ID %>,5);">
            <img src="<% = config_GFXLocation %>icons/grade_no.gif" id="bg4" onmouseover="showGrade(4);" onmouseout="showGrade(0);" onclick="setGrade(<% = text_ID %>,4);">
            <img src="<% = config_GFXLocation %>icons/grade_no.gif" id="bg3" onmouseover="showGrade(3);" onmouseout="showGrade(0);" onclick="setGrade(<% = text_ID %>,3);">
            <img src="<% = config_GFXLocation %>icons/grade_no.gif" id="bg2" onmouseover="showGrade(2);" onmouseout="showGrade(0);" onclick="setGrade(<% = text_ID %>,2);">
            <img src="<% = config_GFXLocation %>icons/grade_no.gif" id="bg1" onmouseover="showGrade(1);" onmouseout="showGrade(0);" onclick="setGrade(<% = text_ID %>,1);">
            
            <input type="hidden" id="userbg" value="<% = text_Betyg %>">
            <script type="text/javascript" language="javascript">showGrade(0);</script>
            
            <div class="nf_separator"></div>
          <% Else %>
            <p style="text-align: center; margin-bottom: 12px;"><em>Du måste vara <strong><a href="/avdelning/medlem/loggain.asp">inloggad</a></strong> för att kunna ange betyg.</em></p>
          <% End If %>
          <span>Medlemsbetyg</span>
          <% If text_MBetyg = 0 Then %>
            <% For zx = 1 To 6 %>
              <img src="<% = config_GFXLocation %>icons/grade_no.gif">
            <% Next %>
          <% Else %>
            <% For zx = 6 To 1 Step -1 %>
              <img src="<% = config_GFXLocation %>icons/grade_o<% If zx <= text_MBetyg Then Response.Write("n") Else Response.Write("ff") %>.gif">
            <% Next %>
          <% End If %>
        </div>
      </div> 
      
      <% If any_Genres Then %>
        <div class="nf_minibox nf_blue">
          <h4>Genre</h4>
          <div class="nf_inside">
            <p class="nf_center"><strong>
              <% For zx = 0 To UBound(list_Genres, 2) %><% If zx > 0 And zx < UBound(list_Genres, 2) Then %>, <% End If %><% If zx > 0 And zx = UBound(list_Genres, 2) Then %> och <% End If %><% = sEncode(list_Genres(1, zx)) %><% Next %>
            </strong></p>
          </div>
        </div>
      <% End If %>
      
      <div class="nf_minibox nf_blue">
        <h4>Speldata</h4>
        <div class="nf_inside">
          <div class="nf_rowhead">Release</div>
          <div class="nf_row"><% = text_Release %></div>
          
          <% If Len(text_Utgivare) > 0 Then %>
            <div class="nf_rowhead">Utgivare</div>
            <div class="nf_row"><% = text_Utgivare %></div>
          <% End If %>
          
          <% If Len(text_Utvecklare) > 0 Then %>
            <div class="nf_rowhead">Utvecklare</div>
            <div class="nf_row"><% = text_Utvecklare %></div>
          <% End If %>
          
          <% If CLng(text_Spelare) > 0 Then %>
            <div class="nf_rowhead">Antal spelare</div>
            <div class="nf_row"><% If text_Spelare > 4 Then %>Fler än 4 spelare<% Else %><% = text_Spelare %> spelare<% End If %></div>
          <% End If %>
        </div>
      </div>
      
      <% If CLng(text_ESRB) > 0 OR CLng(text_PEGI) > 0 Then %>
        <div class="nf_minibox nf_blue">
          <h4>Åldersmärkning</h4>
          <div class="nf_inside nf_ages" style="text-align: center;">
            <img style="float: none; width: 60px; height: 60px;" src="<% = config_GFXLocation %>rating/ESRB_<% = text_ESRB %>.gif" alt="ESRB">
            <img style="float: none; width: 60px; height: 60px;" src="<% = config_GFXLocation %>rating/PEGI_<% = text_PEGI %>.gif" alt="PEGI">
          </div>
        </div>
      <% End If %>
      
      <!-- ## FORUMINLÄGG ## -->
      <% If any_Tradar Then %>
        <div class="nf_minibox nf_green">
          <h4>Forum</h4>
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
      <!-- ## /FORUMINLÄGG ## -->

    </div>
    
    <% If any_Same Then %>
      <div class="nf_datablock nf_size_full">
        <div class="nf_images nf_images_full">
          <p><strong><% = text_Reko %></strong></p>
          <% For zx = 0 To UBound(list_Same, 2) %>
            <%
            text_UseBox = 0
            If CLng(list_Same(9, zx)) > 0 Then text_UseBox = list_Same(9, zx)
            If CLng(list_Same(8, zx)) > 0 Then text_UseBox = list_Same(8, zx)
            If CLng(list_Same(2, zx)) > 0 Then text_UseBox = list_Same(2, zx)
            %>
            <a href="spel_visa_info.asp?e=<% = list_Same(0,zx) %>"><img src="<% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=80&amp;h=80&amp;err=no" title="<% = sEncode(list_Same(1, zx)) %> (<% = list_Same(4, zx) %>)" alt="Spel"></a>
          <% Next %>
          <p>... eller visa <a href="default.asp">alla spel</a> istället.</p>
        </div>
      </div>
    <% End If %>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->