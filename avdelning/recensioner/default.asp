<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  sQ           = Trim(MakeLegal(GetQ("q", "ABC", 255)))
  
  If Len(sQ) > config_MinSearch Then
    ' #### FIX TEXT STRÄNG ####
      q = LCase(Trim(sQ))
      
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
  
    sQUrl = Server.URLEncode(sQ) 
    sQ    = sEncode(sQ)
    
    filter_All = "&amp;q=" & sQUrl

    dataSQL = "SELECT TOP 250 rID, rDatumPublicerad, rTitel, rKategori, rStatus, rAnvandarRec, rText, rFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 1 AND cBindID = rID) " & _
              "FROM cms_Recensioner " & _
              "LEFT JOIN CONTAINSTABLE(cms_Recensioner, *, '" & p & "') AS ct ON rID = ct.[KEY] " & _
              "WHERE Rank > 0 AND rStatus = 4 AND rDatumPublicerad <= '" & Now & "' " & _
              "ORDER BY Rank DESC, rTitel ASC"
              
    noMsg   = "Inga träffar på [<strong>" & sEncode(q) & "</strong>], prova att bredda din sökning."
    aSearch = True
  Else
    dataSQL = "SELECT rID, rDatumPublicerad, rTitel, rKategori, rStatus, rAnvandarRec, rText, rFlash, (SELECT COUNT(cID) FROM cms_Kommentarer WHERE cAvdelning = 1 AND cBindID = rID) " & _
              "FROM cms_Recensioner " & _
              "WHERE rStatus = 4 AND rDatumPublicerad <= '" & Now & "' " & _
              "ORDER BY rDatumPublicerad DESC"
              
    noMsg   = "Det finns inga recensioner att visa."
    aSearch = False
  End If

  RS_Open 1, dataSQL, False
  
    If rsDB(1).EOF Then
      any_Recs = False
    Else
      any_Recs = True
      list_Recs = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  ' ### Kommentarer
  If Not config_LockDown_Kommentarer Then
    RS_Open 1, "SELECT TOP 10 cID, cTextM, fsBB_Anv.aAnvNamn, cDatum, cAvdelning, cBindID, cms_Nyheter.nTitel, cms_Recensioner.rTitel, cms_Artiklar.aaTitel FROM cms_Kommentarer " & _
               "LEFT JOIN fsBB_Anv ON cms_Kommentarer.cAnv = fsBB_Anv.aID " & _
               "LEFT JOIN cms_Nyheter ON cms_Kommentarer.cBindID = cms_Nyheter.nID " & _
               "LEFT JOIN cms_Recensioner ON cms_Kommentarer.cBindID = cms_Recensioner.rID " & _
               "LEFT JOIN cms_Artiklar ON cms_Kommentarer.cBindID = cms_Artiklar.aaID " & _
               "WHERE cAvdelning = 1 " & _
               "ORDER BY cDatum DESC", False
    
      If rsDB(1).EOF Then
        any_Kommentarer = False
      Else
        any_Kommentarer = True
        list_Kommentarer = rsDB(1).GetRows
      End If
    
    RS_Close 1
  End If
  
  If any_Recs Then
    CreatePaging 50, UBound(list_Recs, 2)
    CreatePagingChooser
  End If
  
  If pagingOnPage < 1 Then pagingOnPage = 1
%>

<%
  ' ## Globala variabler ##
  
  If aSearch Then
    page_Title    = sEncode(q) & " - Sök - Sida " & pagingOnPage & " - Recensioner"
  Else
    page_Title    = "Recensioner - Sida " & pagingOnPage
  End If
  
  page_Header   = "Nintendo recensioner"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Recensioner&quot; ...'>Recensioner</a> "
  page_SelMenu  = "texter"
  page_Slide    = "recensioner"
  
  page_description    = "Arkiv över alla våra recensioner på N-Forum.se, Nintendo Forum. Du är på sida " & pagingOnPage & "."
  page_keywords       = "recensioner, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/recensioner/">Recensioner</a></span>Nintendo Recensioner</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">

      <div class="nf_msg nf_largesearch">
        <p><strong>Sök efter recensioner...</strong></p>
        <form>
          <input style="width: 564px;" type="text" maxlength=255 name="q" value="<% = sQ %>"> 
          <input style="float: right; width: 80px; font-weight: bold;" type="submit" value="Sök">
        </form>
      </div>
    
      <% If any_Recs Then %>
          
          <ul class="nf_list">
            <% If aSearch Then %><li class="nf_listsplit"> Sökträffar </li><% End If %>
          
            <%
              remPost = CDate("2050-01-01 00:00:00")
              For zx = pagingBOF To pagingEOF
                If zx > UBound(list_Recs, 2) Then Exit For
                
                If Not aSearch Then
                  If DateDiff("m", list_Recs(1, zx), remPost) <> 0 OR DateDiff("yyyy", list_Recs(1, zx), remPost) <> 0 Then 
                    remPost = list_Recs(1, zx)
                    %>
                    <li class="nf_listsplit"> <% = MonthName(Month(list_Recs(1, zx))) %>&nbsp;<% = Year(list_Recs(1, zx)) %> </li>
                    <%
                  End If
                End If
                %>
                  <li> 
                    <% If CLng(list_Recs(7, zx)) > 0 Then %>
                      <div class="nf_front" style="height: 120px; background-image: url('<% = config_ImageLocation %>?e=<% = list_Recs(7, zx) %>&w=180&h=120')"><p>&nbsp;</p></div>
                    <% Else %>
                      <div class="nf_front nf_front_rec"><p><% If list_Recs(5, zx) Then %>Användarrecension<% Else %>Recension<% End If %></p></div>
                    <% End If %>
                    <div class="nf_data">
                      <h3><a href="recension_visa.asp?e=<% = list_Recs(0, zx) %>" title="<% = sEncode(list_Recs(2, zx)) %>"><% = sEncode(list_Recs(2, zx)) %></a></h3>
                      <span class="nf_medium nf_gray nf_bold"><% If list_Recs(5, zx) Then %>Användarrecension<% Else %>Recension<% End If %> / <% = lstKonsol(list_Recs(3, zx)) %> / <% = DatumReplace(list_Recs(1, zx)) %></span>
                      <p style="line-height: 18px;"><% = sEncode(CutText(BBCode_Remove(list_Recs(6, zx)),180)) %></p>
                      
                      <div class="nf_morebtn">
                        <a href="recension_visa.asp?e=<% = list_Recs(0, zx) %>">Läs mer ...</a>
                        <a href="recension_visa.asp?e=<% = list_Recs(0, zx) %>#kommentarer" <% If CLng(list_Recs(8, zx)) > 0 Then %>class="nf_hint"<% End If %>><% If CLng(list_Recs(8, zx)) = 0 Then %>Kommentera texten<% ElseIf CLng(list_Recs(8, zx)) = 1 Then %>1 kommentar<% Else %><% = list_Recs(8, zx) %> kommentarer<% End If %></a>
                      </div>
                    </div>
                  </li>
                <%
              Next
            %>
          </ul>
          
          <div class="nf_paging">
            <a href="default.asp?page=<% = pagingOnPage - 1 %><% = filter_all %>">««</a> |
            
              <% For Each zx In pagingPages %>
                <% If zx = "..." Then %>
                  ... |
                <% Else %>
                  <a href="default.asp?page=<% = zx %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
                <% End If %>
              <% Next %>
              
            | <a href="default.asp?page=<% = pagingOnPage + 1 %><% = filter_all %>">»»</a>
          </div>
      <% Else %>
        <div class="nf_msg"><p><% = noMsg %></p></div>
      <% End If %>
    </div>
      
    <div class="nf_datablock nf_size_onethird">
      
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <div class="nf_minibox nf_blue">
        <h4>Recensioner på N-Forum.se</h4>
        <div class="nf_inside">
          <p>Här finner du alla recensioner som finns på N-Forum.se</p>
        </div>
      </div>
      
      <div class="nf_minibox nf_green">
        <h4>Söktips</h4>
        <div class="nf_inside">
          <p>Sök på minst <strong>3</strong> tecken.</p>
          <p>Om du får för många resultat prova att vara mer specifik och använd fler ord.</p>
          <p>Du kan <strong>INTE</strong> använda termer så som AND, OR och liknande</p>
        </div>
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
        
    </div>
  
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->