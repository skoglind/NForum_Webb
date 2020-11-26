<!--#INCLUDE FILE="../__INC/includes.asp"-->

<%

  ' ### Valt datum
    rDate   = GetQ("d", "ABC", 0)
    If IsDate(rDate) Then
      lDay    = Day(rDate)
      lMonth  = Month(rDate)
      lYear   = Year(rDate)
      If lYear < 1975 Or lYear > 2025 Then
        lDay    = Day(Date)
        lMonth  = Month(Date)
        lYear   = Year(Date)
      End If
    Else
      lDay    = Day(Date)
      lMonth  = Month(Date)
      lYear   = Year(Date)
    End If

    fDate = CDate(lYear & "-" & lMonth & "-" & lDay)
    lOnDay  = 0

    ' ## Månadens data
    sMonthName        = MonthName(lMonth)
    lFirstDayOfMonth  = CLng(WeekDay(CDate(lYear & "-" & Right("00" & lMonth, 2) & "-01"), vbMonday))
    lLastDayOfMonth = GetLastDayOfMonth(lYear,lMonth)
    
    ' ## Förra månadens data
    lPreviousYear   = lYear
    lPreviousMonth  = lMonth - 1
    If lPreviousMonth = 0 Then lPreviousMonth = 12 : lPreviousYear = lPreviousYear - 1
    lFirstMonth      = GetLastDayOfMonth(lPreviousYear,lPreviousMonth)
    
    ' ## Nästa månads Data
    lNextYear  = lYear
    lNextMonth = lMonth + 1
    If lNextMonth = 13 Then lNextMonth = 1 : lNextYear = lNextYear + 1
    lLastMonth = 0
    
    lDateInText = lDay & " " & sMonthName & " " & lYear

    
    ' ### HÄMTA ALL DATA AKTUELLT DATUM
    
    ' ### Fler nyheter
      RS_Open 1, "SELECT nID, nTitel, nDatumPublicerad AS tdPubl, nKategori, nStatus, nText, nIdent, nFlash FROM cms_Nyheter " & _
                 "WHERE nStatus = 4 AND (DATEDIFF(d, nDatumPublicerad, '" & fDate & "') = 0 AND nDatumPublicerad <= '" & Now & "') " & _
                 "UNION ALL " & _
                 "SELECT aaID, aaTitel, aaDatumPublicerad AS tdPubl, aaKategori, aaStatus, aaText, aaIdent, aaFlash FROM cms_Artiklar " & _
                 "WHERE aaAnvandarArt = 0 AND aaStatus = 4 AND (DATEDIFF(d, aaDatumPublicerad, '" & fDate & "') = 0 AND aaDatumPublicerad <= '" & Now & "') " & _
                 "UNION ALL " & _
                 "SELECT rID, rTitel, rDatumPublicerad AS tdPubl, rKategori, rStatus, rText, rIdent, rFlash FROM cms_Recensioner " & _
                 "WHERE rAnvandarRec = 0 AND rStatus = 4 AND (DATEDIFF(d, rDatumPublicerad, '" & fDate & "') = 0 AND rDatumPublicerad <= '" & Now & "') " & _
                 "ORDER BY tdPubl DESC", False
      
        If rsDB(1).EOF Then
          any_TextData     = False
        Else
          any_TextData     = True
          list_TextData    = rsDB(1).GetRows
        End If
      
      RS_Close 1
      
    ' ### Fler spel
      RS_Open 1, "SELECT tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett FROM cms_SpelTitlar " & _
                 "LEFT JOIN cms_Region ON tRegion = rID " & _
                 "LEFT JOIN cms_Spel ON sID = tSpelID " & _
                 "WHERE sSynlig = 1 AND tRelease = '" & fDate & "' " & _
                 "ORDER BY sDatumSparad DESC", False
      
        If rsDB(1).EOF Then
          any_Spel     = False
        Else
          any_Spel     = True
          list_Spel    = rsDB(1).GetRows
        End If
      
      RS_Close 1
      
    ' ### Fler konsoler
      RS_Open 1, "SELECT tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, kKonsol, kTextM, tBoxart_Manual, tBoxart_Konsol FROM cms_KonsolTitlar " & _
                 "LEFT JOIN cms_Region ON tRegion = rID " & _
                 "LEFT JOIN cms_Konsol ON kID = tKonsolID " & _
                 "WHERE kSynlig = 1 AND tRelease = '" & fDate & "' " & _
                 "ORDER BY kDatumSparad DESC", False
      
        If rsDB(1).EOF Then
          any_Konsol     = False
        Else
          any_Konsol     = True
          list_Konsol    = rsDB(1).GetRows
        End If
      
      RS_Close 1
        
   ' ### Fler tillbehör
      RS_Open 1, "SELECT tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, iKonsol, iTextM, tBoxart_Manual, tBoxart_Tillbehor FROM cms_TillbehorTitlar " & _
                 "LEFT JOIN cms_Region ON tRegion = rID " & _
                 "LEFT JOIN cms_Tillbehor ON iID = tTillbehorID " & _
                 "WHERE iSynlig = 1 AND tRelease = '" & fDate & "' " & _
                 "ORDER BY iDatumSparad DESC", False
      
        If rsDB(1).EOF Then
          any_Tillbehor     = False
        Else
          any_Tillbehor     = True
          list_Tillbehor    = rsDB(1).GetRows
        End If
      
      RS_Close 1
      
    ' ### Fler foruminlägg
      If Not config_LockDown_Forum Then
        RS_Open 1, "SELECT tID, tAmne, tTextM, tDatum_Skapad, tStatus_Trad, tStatus_UnderTrad, " & _
                   "(SELECT COUNT(tID) FROM fsBB_Tradar WHERE tStatus_UnderTrad = tbTrad.tID AND tStatus_Trad = 0) AS iAntalSvar, fIcon, fName " & _
                   "FROM fsBB_Tradar AS tbTrad " & _
                   "LEFT JOIN fsBB_Forum ON tbTrad.tForum = fsBB_Forum.fID " & _
                   "WHERE DATEDIFF(d, tDatum_Skapad, '" & fDate & "') = 0 AND (fSec_View = '0' OR fSec_View LIKE '%;" & SEC_TITEL & ";%') AND tForum <> " & CLng(config_Trashbin) & " AND tStatus_Raderad = 0 ORDER BY tDatum_Skapad DESC", False
        
          If rsDB(1).EOF Then
            any_Tradar = False
          Else
            any_Tradar = True
            list_Tradar = rsDB(1).GetRows(50)
          End If
        
        RS_Close 1
      End If
      
    ' ### Fler senast registrerade
      RS_Open 1, "SELECT aAnvNamn, aTimeStamp, fsBB_Titlar.ttText AS aTitelText, aMedlemSedan, aEgenTitel, aAvatar, aID FROM fsBB_Anv " & _
                 "LEFT JOIN fsBB_Titlar ON aTitelID = fsBB_Titlar.ttID " & _
                 "WHERE aBlockadTill < '" & Date & "' AND aAktiverad = 1 AND DATEDIFF(d, aMedlemSedan, '" & fDate & "') = 0 ORDER BY aAnvNamn DESC", False
      
        If rsDB(1).EOF Then
          any_On = False
        Else
          any_On = True
          list_On = rsDB(1).GetRows
        End If
      
      RS_Close 1
    
%>

<%
  ' ## Globala variabler ##
  page_Title    = lDateInText & " - Kalender"
  page_Header   = "Kalender"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Hem&quot; ...'>Första sidan</a> "
  page_SelMenu  = "home"
  page_Slide    = "forum"
  
  page_description  = "Vad hände inom spelvärlden och N-Forum.se, Nintendo Forum, valt datum: " & lDateInText & ". Kalender över alla dessa händelser."
  page_keywords     = "kalender, "
%>

<!--#INCLUDE FILE="../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
     
      <h1>Nintendo kalender - <% = lDateInText %></h1>
          
    </div>
  
    <div class="nf_datablock nf_size_onethird">
     
      <div class="nf_msg nf_msg_third">
        <p><strong>Vad hände på sidan detta datum?</strong></p>
      </div>
    
      <!-- ## SENASTE TEXTERNA ## -->
      
        <div class="nf_minibox">
          <h4>Publicerade texter</h4>
          <div class="nf_inside nf_stylelist">
            <% If any_TextData Then %>
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
            <% Else %>
              <p>Inga texter är publicerade valt datum.</p>
            <% End If %>
            <p><a href="/avdelning/texter/">Visa alla texter</a></p>
          </div>
        </div>
      
      <!-- ## SENASTE TEXTERNA ## -->
      
      <!-- ## SENASTE FORUMINLÄGGEN ## -->
      
        <div class="nf_minibox nf_blue">
          <h4>Foruminlägg</h4>
          <div class="nf_inside nf_stylelist">
            <% If any_Tradar Then %>
              <ul>
                <% For zx = 0 To UBound(list_Tradar, 2) %>
                  <%
                    isTheThread = False
                    If list_Tradar(4,zx) Then isTheThread = True
                  %>
                  <li onclick="location.href='/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>';"><a href="/avdelning/forum/trad.asp<% If isTheThread Then %>?e=<% = list_Tradar(0,zx) %><% Else %>?e=<% = list_Tradar(5,zx) %>&amp;go2=<% = list_Tradar(0,zx) %><% End If %>" title="<% = sEncode(list_Tradar(1,zx)) %>"><% = sEncode(CutText(list_Tradar(1,zx), 32)) %></a><% = list_Tradar(8, zx) %> / <% = DatumReplace(list_Tradar(3,zx)) %></li>
                <% Next %>
              </ul>
            <% Else %>
              <p>Inga inlägg är skrivna valt datum.</p>
            <% End If %>
            <p><a href="/avdelning/forum/nyainlagg.asp">Visa alla foruminlägg</a></p>
          </div>
        </div>
      
      <!-- ## /SENASTE FORUMINLÄGGEN ## -->
      
      <!-- ## SENASTE REGISTRERADE ## -->
      
        <div class="nf_minibox nf_red">
          <h4>Registrerade medlemmar</h4>
          <div class="nf_inside nf_stylelist">
            <% If any_On Then %>
              <ul>
                <% For zx = 0 To UBound(list_On, 2) %>
                  <li onclick="location.href='/avdelning/medlem/?m=<% = sEncode(list_On(0,zx)) %>';"><a href="/avdelning/medlem/?m=<% = sEncode(list_On(0,zx)) %>" title="<% = sEncode(list_On(0,zx)) %>"><% = sEncode(CutText(list_On(0,zx), 32)) %></a> Registrerad: <% = DatumReplace(list_On(3,zx)) %></li>
                <% Next %>
              </ul>
            <% Else %>
              <p>Inga registrerade medlemmar valt datum.</p>
            <% End If %>
            <p><a href="/avdelning/listor/sokmedlem.asp">Sök medlem</a></p>
          </div>
        </div>
      
      <!-- ## /SENASTE REGISTRERADE ## -->
          
    </div>
    
    
    <div class="nf_datablock nf_size_onethird">
      
      <div class="nf_msg nf_msg_third">
        <p><strong>Vad är utgivet detta datum?</strong></p>
      </div>
    
      <!-- ## SENAST UPPDATERADE SPELEN ## -->
      
        <div class="nf_minibox nf_green">
          <h4>Utgivna spel</h4>
          <div class="nf_inside nf_stylelist">
            <% If any_Spel Then %>
              <ul>
                <% For zx = 0 To UBound(list_Spel, 2) %>
                  <%
                  text_UseBox = 0
                  If CLng(list_Spel(8, zx)) > 0 Then text_UseBox = list_Spel(8, zx)
                  If CLng(list_Spel(7, zx)) > 0 Then text_UseBox = list_Spel(7, zx)
                  If CLng(list_Spel(2, zx)) > 0 Then text_UseBox = list_Spel(2, zx)
                  %>
                  <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/spel_lrg.png<% End If %>');" onclick="location.href='/avdelning/spel/spel_visa_info.asp?e=<% = list_Spel(0, zx) %>';"><a href="/avdelning/spel/spel_visa_info.asp?e=<% = list_Spel(0, zx) %>" title="<% = sEncode(list_Spel(1, zx)) %>"><% = sEncode(CutText(list_Spel(1, zx), 32)) %></a><% = lstKonsol(list_Spel(5, zx)) %></li>
                <% Next %>
              </ul>
            <% Else %>
              <p>Inga utgivna spel valt datum.</p>
            <% End If %>
            <p><a href="/avdelning/spel/">Visa alla spel</a></p>
          </div>
        </div>
      
      <!-- ## /SENAST UPPDATERADE SPELEN ## -->
      
      <!-- ## SENAST UPPDATERADE KONSOLERNA ## -->
      
        <div class="nf_minibox nf_green">
          <h4>Utgivna konsoler</h4>
          <div class="nf_inside nf_stylelist">
            <% If any_Konsol Then %>
              <ul>
                <% For zx = 0 To UBound(list_Konsol, 2) %>
                  <%
                  text_UseBox = 0
                  If CLng(list_Konsol(8, zx)) > 0 Then text_UseBox = list_Konsol(8, zx)
                  If CLng(list_Konsol(7, zx)) > 0 Then text_UseBox = list_Konsol(7, zx)
                  If CLng(list_Konsol(2, zx)) > 0 Then text_UseBox = list_Konsol(2, zx)
                  %>
                  <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/konsol_lrg.png<% End If %>');" onclick="location.href='/avdelning/konsol/konsol_visa_info.asp?e=<% = list_Konsol(0, zx) %>';"><a href="/avdelning/konsol/konsol_visa_info.asp?e=<% = list_Konsol(0, zx) %>" title="<% = sEncode(list_Konsol(1, zx)) %>"><% = sEncode(CutText(list_Konsol(1, zx), 32)) %></a><% = lstKonsol(list_Konsol(5, zx)) %></li>
                <% Next %>
              </ul>
            <% Else %>
              <p>Inga utgivna konsoler valt datum.</p>
            <% End If %>
            <p><a href="/avdelning/konsol/">Visa alla konsoler</a></p>
          </div>
        </div>
      
      <!-- ## /SENAST UPPDATERADE KONSOLERNA ## -->
      
      <!-- ## SENAST UPPDATERADE TILLBEHÖREN ## -->
      
        <div class="nf_minibox nf_green">
          <h4>Utgivna tillbehör</h4>
          <div class="nf_inside nf_stylelist">
            <% If any_Tillbehor Then %>
              <ul>
                <% For zx = 0 To UBound(list_Tillbehor, 2) %>
                  <%
                  text_UseBox = 0
                  If CLng(list_Tillbehor(8, zx)) > 0 Then text_UseBox = list_Tillbehor(8, zx)
                  If CLng(list_Tillbehor(7, zx)) > 0 Then text_UseBox = list_Tillbehor(7, zx)
                  If CLng(list_Tillbehor(2, zx)) > 0 Then text_UseBox = list_Tillbehor(2, zx)
                  %>
                  <li style="background-image: url('<% If CLng(text_UseBox) > 0 Then %><% = config_ImageLocation %>?e=<% = text_UseBox %>&amp;w=23&amp;h=23&amp;err=no<% Else %><% = config_GFXLocation %>icons/tillbehor_lrg.png<% End If %>');" onclick="location.href='/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_Tillbehor(0, zx) %>';"><a href="/avdelning/tillbehor/tillbehor_visa_info.asp?e=<% = list_Tillbehor(0, zx) %>" title="<% = sEncode(list_Tillbehor(1, zx)) %>"><% = sEncode(CutText(list_Tillbehor(1, zx), 32)) %></a><% = lstKonsol(list_Tillbehor(5, zx)) %></li>
                <% Next %>
              </ul>
            <% Else %>
              <p>Inga utgivna tillbehör valt datum.</p>
            <% End If %>
            <p><a href="/avdelning/tillbehor/">Visa alla tillbehör</a></p>
          </div>
        </div>
      
      <!-- ## /SENAST UPPDATERADE TILLBEHÖREN ## -->
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
     
      <!-- #### KALENDERN #### -->
        <style type="text/css">
          .calender {
            float:            left;
            background-color: #FFF;
            width:            295px;
          }
          
          .calender .cal_title {
            float:            left;
            width:            295px;
            border-bottom:    solid 2px #FFF;
          }
          
          .calender .cal_title .cal_chooser {
            float:            left;
            width:            30px;
            font:             bold 20px Verdana;
            background-color: #1e5a6d;
          }
          
          .calender .cal_title .cal_chooser:hover {
            background-color: #327388;
          }
          
          .calender .cal_title .cal_chooser_left {
            -webkit-border-top-left-radius: 5px;
            -moz-border-radius-topleft: 5px;
            border-top-left-radius: 5px;
          }
          
          .calender .cal_title .cal_chooser_right {
            -webkit-border-top-right-radius: 5px;
            -moz-border-radius-topright: 5px;
            border-top-right-radius: 5px;
          }
          
          .calender .cal_title .cal_chooser a {
            float:            left;
            width:            100%;
            height:           35px;
            padding:          5px 0 0 0;
            text-align:       center;
            color:            #FFF;
          }
          
          .calender .cal_title .cal_text {
            float:            left;
            width:            191px;
            height:           40px;
            padding:          0 22px 0 22px;
            background-color: #327388;
            overflow:         hidden;
          }
          
          .calender .cal_title .cal_text select {
            font:             16px Arial;
            margin:           8px 5px 0 5px;
          }
          
          .calender .cal_objects {
            float:            left;
            width:            294px;
            border-right:     solid 1px #FFF;
          }
          
          .calender .cal_objects .cal_object {
            float:            left;
            border-left:      solid 1px #FFF;
            border-bottom:    solid 1px #FFF;
            width:            41px;
            background-color: #EEE;
            font:             14px Verdana;
            overflow:         hidden;
          }
          
          .calender .cal_objects .cal_object:hover {
            background-color: #CCC;
          }
          
          .calender .cal_objects .cal_month {
            background-color: #DDD;
          }
          
          .calender .cal_objects .cal_active {
            background-color: #BBB;
            font-weight:      bold;
          }
          
          .calender .cal_objects .cal_object a {
            float:            left;
            width:            100%;
            height:           30px;
            padding:          10px 0 0 0;
            text-align:       center;
            color:            #AAA;
          }
          
          .calender .cal_objects .cal_active a {
            color:            #000;
          }
          
          .calender .cal_objects .cal_month a {
            color:            #333;
          }
        </style>
      
        <div class="calender">
          <div class="cal_title">
            <div class="cal_chooser cal_chooser_left"><a href="?d=<% = lPreviousYear %>-<% = lPreviousMonth %>-01" title="Gå till föregående månad...">«</a></div>
              <div class="cal_text">
                <select onchange="location.href='?d=<% = lYear %>-' + value + '-01';">
                  <% For zz = 1 To 12 %>
                    <option value="<% = Right("00" & zz, 2) %>" <% If CLng(lMonth) = zz Then Response.Write(" selected style='font-weight: bold;'") %>> <% = MonthName(zz) %> </option>
                  <% Next %>
                </select>
                <select onchange="location.href='?d=' + value + '-<% = lMonth %>-01';">
                  <% For zz = 1980 To 2012 %>
                    <option value="<% = zz %>" <% If CLng(lYear) = zz Then Response.Write(" selected style='font-weight: bold;'") %>> <% = zz %> </option>
                  <% Next %>
                </select>
              </div>
            <div class="cal_chooser cal_chooser_right"><a href="?d=<% = lNextYear %>-<% = lNextMonth %>-01" title="Gå till nästa månad...">»</a></div>
          </div>
          
          <div class="cal_objects">
            <% For xx = 1 To 42 %>
              <% If lFirstDayOfMonth <= xx And lLastDayOfMonth > lOnDay Then %>
                <% lOnDay = lOnDay + 1 %>
                <div class="cal_object cal_month <% If lDay = lOnDay Then Response.Write("cal_active") %>"><a href="?d=<% = lYear %>-<% = lMonth %>-<% = lOnDay %>"><% = lOnDay %></a></div>
              <% Else %>
                <% If xx = 36 Then Exit For %>
              
                <% If lOnDay = 0 Then %>
                  <% setDay = lFirstMonth - (lFirstDayOfMonth - (xx + 1)) %>
                  <div class="cal_object"><a href="?d=<% = lPreviousYear %>-<% = lPreviousMonth %>-<% = setDay %>"><% = setDay %></a></div>
                <% Else %>
                  <% lLastMonth = lLastMonth + 1 %>
                  <div class="cal_object"><a href="?d=<% = lNextYear %>-<% = lNextMonth %>-<% = lLastMonth %>"><% = lLastMonth %></a></div>
                <% End if %>
              <% End If %>
            <% Next %>
          </div>
        </div>
      <!-- #### /KALENDERN #### -->
          
      <div class="nf_msg nf_msg_third">
        <p><strong>Information</strong>
        <p>Här kommer du få listat vilka texter som publicerades, inlägg som skrevs, medlemmar som registrerades, spel som gavs ut, konsoler som gavs ut och tillbehör som gavs valt datum.</p>
        <p>Datum för utgivna spel, konsoler och tillbehör är efter vad som finns lagrat i vår databas.</p>
      </div>
      
    </div>
    
  </div>

<!--#INCLUDE FILE="../_page_bottom.asp"-->
<!--#INCLUDE FILE="../__INC/includes_end.asp"-->