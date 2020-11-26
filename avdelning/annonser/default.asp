<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<%
  sQ           = Trim(MakeLegal(GetQ("q", "ABC", 255)))
  filter_kat = GetQ("k","123",0)
  If filter_kat > 0 And filter_kat < lstKSKategori(-1) + 1 Then kat_SQL = "AND ksKategori1 = " & CLng(filter_kat)  
  
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
    
    filter_All = "&amp;q=" & sQUrl & "&amp;k=" & filter_kat
    
    dataSQL = "SELECT TOP 250 ksID, ksSkapadDatum, ksTitel, ksKategori1, ksStatus, ksKategori2, ksSkapadAv, ksTyp, ksTextM " & _
              "FROM cms_KopSalj " & _
              "LEFT JOIN CONTAINSTABLE(cms_KopSalj, *, '" & p & "') AS ct ON ksID = ct.[KEY] " & _
              "WHERE Rank > 0 " & _
              kat_SQL & " " & _
              "AND ksSkapadDatum + " & CLng(config_AdDays) & " > '" & Now & "' " & _
              "AND ksSynlig = 1 " & _
              "ORDER BY Rank DESC, ksTitel ASC"
              
    noMsg   = "Inga träffar på [<strong>" & sEncode(q) & "</strong>], prova att bredda din sökning."
    aSearch = True
  Else
    dataSQL = "SELECT ksID, ksSkapadDatum, ksTitel, ksKategori1, ksStatus, ksKategori2, ksSkapadAv, ksTyp, ksTextM " & _
              "FROM cms_KopSalj WHERE ksID > 0 " & _ 
              "AND ksSkapadDatum + " & CLng(config_AdDays) & " > '" & Now & "' " & _
              "AND ksSynlig = 1 " & _
              "ORDER BY ksSkapadDatum DESC"
              
    noMsg   = "Det finns inga annonser att visa."
    aSearch = False
  End If
  
  RS_Open 1, dataSQL, False
  
    If rsDB(1).EOF Then
      any_Ads = False
    Else
      any_Ads = True
      list_Ads = rsDB(1).GetRows
    End If
  
  RS_Close 1
  
  If any_Ads Then
    CreatePaging 50, UBound(list_Ads, 2)
    CreatePagingChooser
  End If
  
  If pagingOnPage < 1 Then pagingOnPage = 1
%>

<%
  ' ## Globala variabler ##
  If aSearch Then
    page_Title    = sEncode(q) & " - Sök - Sida " & pagingOnPage & " - Annonser"
  Else
    page_Title    = "Annonser - Sida " & pagingOnPage
  End If
  
  page_Header   = "Marknad"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Marknad&quot; ...'>Marknad</a>"
  page_SelMenu  = "buy"
  page_Slide    = "annonser"
  
  page_description    = "Alla upplagda annonser i vår köp- och säljavdelning på N-Forum.se, Nintendo Forum. Du är på sida " & pagingOnPage & "."
  page_keywords       = "annonser, "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/annonser/">Annonser</a></span>Alla Annonser</h1>
    </div>
  
    <div class="nf_datablock nf_size_twothird">
        <div class="nf_msg">
          <p><strong>Sök efter annonser...</strong></p>
          <form>
            <select name="k" style="width: 569px;">
              <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> Alla kategorier </option>
              <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
              <% For zx = 1 To lstKSKategori(-1) %>
                <option value=<% = zx %> style="padding: 1px 0 1px 10px;" <% If filter_kat = zx Then Response.Write(" selected") %>> <% = lstKSKategori(zx) %> </option>
              <% Next %>
            </select> 
          
            <input style="width: 564px;" type="text" maxlength=255 name="q" value="<% = sQ %>"> 
            <input style="float: right; width: 80px; font-weight: bold;" type="submit" value="Sök">
          </form>
        </div>
      
        <% If any_Ads Then %>

            <ul class="nf_list">
              <% If aSearch Then %>
                <li class="nf_listsplit"> Sökträffar </li>
              <% Else %>
                <li class="nf_listsplit"> Alla annonser </li>
              <% End If %>
            
              <%
                For zx = pagingBOF To pagingEOF
                  If zx > UBound(list_Ads, 2) Then Exit For
                  %>
                    <li>
                      <div class="nf_header" style="background-color: <% = getColor(list_Ads(7, zx)) %>;">
                        <span><% = lstKSTyp(list_Ads(7, zx)) %></span>
                        <% If list_Ads(4, zx) = 1 Then %><div class="nf_xtra">Avslutad!</div><% End If %>
                      </div>
                      <div class="nf_data">
                        <h3><a href="annons_visa.asp?e=<% = list_Ads(0, zx) %>" title="<% = sEncode(list_Ads(2, zx)) %>"><% = sEncode(list_Ads(2, zx)) %></a></h3>
                        <span class="nf_medium nf_gray nf_bold"><% = lstKSKategori(list_Ads(3, zx)) %> / <% = DatumReplace(list_Ads(1, zx)) %></span>
                        <p><% = sEncode(CutText(BBCode_Remove(list_Ads(8, zx)),300)) %></p>
                        <span class="nf_small nf_bold">» <a href="annons_visa.asp?e=<% = list_Ads(0, zx) %>">Visa annonsen</a>...</span>
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
      
        <% If CONST_LOGIN Then %>
          <div class="nf_minibox nf_green">
            <h4>Skapa annons</h4>
            <div class="nf_inside">
              <p class="nf_huge nf_center"><strong><a href="ny_annons.asp" title="">Ny annons</a></strong></p>
            </div>
          </div>
        <% Else %>
          <div class="nf_minibox nf_green">
            <h4>Logga in</h4>
            <div class="nf_inside">
              <p style="text-align: center;"><em>Du måste <strong><a href="/avdelning/medlem/loggain.asp">logga in</a></strong> för att kunna kommentera och/eller skapa annonser.</em></p>
              <p style="text-align: center;"><em>Om du inte redan har en användare kan du <strong><a href="/avdelning/medlem/registreradig.asp">registrera dig</a></strong>.</em></p>
            </div>
          </div>
        <% End If %>
      
        <div class="nf_minibox nf_blue">
          <h4>Annonser på N-Forum.se</h4>
          <div class="nf_inside">
            <p>Här finner du alla annonser som finns på N-Forum.se</p>
            <p>Vill du söka bland annonserna gör du det via sökrutan.</p>
          </div>
        </div>

        <div class="nf_minibox nf_blue">
          <h4>Söktips</h4>
          <div class="nf_inside">
            <p>Sök på minst <strong>3</strong> tecken.</p>
            <p>Om du får för många resultat prova att vara mer specifik och använd fler ord.</p>
            <p>Du kan <strong>INTE</strong> använda termer så som AND, OR och liknande</p>
          </div>
        </div>
        
        <div class="nf_minibox nf_red">
          <h4>Observera</h4>
          <div class="nf_inside">
            <p class="nf_bold">N-Forum.se ansvarar inte för vad som säljs i denna avdelning utan det ansvarar aktuell säljare för.</p>
          </div>
        </div>
        
      </div>
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->