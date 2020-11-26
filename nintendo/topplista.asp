<!--#INCLUDE FILE="../__INC/includes.asp"-->

<%

  ' ### Bra spel
  RS_Open 1, "SELECT TOP 10 tID, tTitel, tBoxart_BoxFram, tRegion, rNamn, sKonsol, sTextM, tBoxart_Manual, tBoxart_Kassett, " & _
             "((SELECT SUM(bBetyg) FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID) / (SELECT COUNT(bID) FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID)) AS clBetyg, " & _
             "(SELECT COUNT(*) FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID) AS clBetyg_Antal " & _
             "FROM cms_SpelTitlar " & _
             "LEFT JOIN cms_Region ON tRegion = rID " & _
             "LEFT JOIN cms_Spel ON sID = tSpelID " & _
             "WHERE tSpelID IN(SELECT bSpelID FROM cms_SpelBetyg WHERE bSpelID = cms_SpelTitlar.tSpelID) AND sSynlig = 1 AND tID = sStandard_Titel " & _
             konsol_SQL & _
             "ORDER BY clBetyg DESC, clBetyg_Antal DESC", False
  
    If rsDB(1).EOF Then
      any_Spel     = False
    Else
      any_Spel     = True
      list_Spel    = rsDB(1).GetRows(10)
    End If
  
  RS_Close 1    

%>

<%
  ' ## Globala variabler ##
  page_Title    = "B�sta Nintendo Spelen - Topplista"
  page_Header   = "Topplista"
  page_WhereAmI = "&gt; <a href='default.asp' title='G� till &quot;Hem&quot; ...'>F�rsta sidan</a> "
  page_SelMenu  = "home"
  page_Slide    = "forum"
  
  page_description  = "Spelen med h�gst medlemsbetyg p� N-Forum.se, Nintendo Forum. V�r egna interna topplista."
  page_keywords     = "topplista, "
%>

<!--#INCLUDE FILE="../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../_page_middle.asp"-->

  <div class="content">
    
    <div class="nf_datablock nf_size_full">
     
      <h1>Topplista - B�sta spelen enligt v�ra medlemmar</h1>

    </div>
    
    <div class="nf_datablock nf_size_twothird">
      
      <% If any_Spel Then %>
        <% For zx = 0 To UBound(list_Spel, 2) %>
          
          <% = list_Spel(1, zx) %><br>
        
        <% Next %>
      <% End If %>
    
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      Som inloggad medlem kan du sj�lv vara med och avg�ra de b�sta spelen.
      
      P� varje spel finns det till h�ger om dem sex stj�rnor f�r att ange ditt betyg f�r just det spelet. Klicka bara p� valfri stj�rna s� har du angett ditt betyg.
    </div>
    
  </div>

<!--#INCLUDE FILE="../_page_bottom.asp"-->
<!--#INCLUDE FILE="../__INC/includes_end.asp"-->