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
      konsol_SQL = "AND kKonsol = " & CLng(filter_Konsol)
      konsol_Add = lstKonsol(filter_konsol)
    Else
      filter_konsol = 0
      konsol_Add = "Alla Konsoler"
    End If
    
    filter_All = ""        
  ' ################

  lAnvID = CONST_USERID
  If lAnvID = Empty Then lAnvID = 0

  RS_Open 1, "SELECT kID, tTitel, kKonsol, tSortNo, tBoxart_BoxFram, tBoxart_Manual, tBoxart_Konsol, tRegion, tRelease, tID, " & _ 
             "(SELECT COUNT(biID) FROM cms_Bind_Anv_Konsol WHERE biTitelID = cms_KonsolTitlar.tID AND biAnv = " & CLng(lAnvID) & ") AS tListadAntal " & _
             "FROM cms_KonsolTitlar " & _ 
             "LEFT JOIN cms_Konsol ON cms_KonsolTitlar.tKonsolID = cms_Konsol.kID " & _ 
             "WHERE kSynlig = 1 " & _
             alfa_SQL & _
             region_SQL & _
             konsol_SQL & _
             "ORDER BY tTitel ASC", False
  
    If rsDB(1).EOF Then
      any_Consoles = False
    Else
      any_Consoles = True
      list_Consoles = rsDB(1).GetRows
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
  
  If any_Consoles Then
    CreatePaging config_MaxAntalPosterPerSida, UBound(list_Consoles, 2)
    CreatePagingChooser
  End If
  
  If pagingOnPage < 1 Then pagingOnPage = 1
%>

<%
  ' ## Globala variabler ##
  If CLng(filter_region) > 0 Then text_Region = "utgivna i " & GetRegion(filter_region)
  If Len(filter_alfa) > 0 Then If UCase(filter_alfa) = "NUM" Then text_Alfa   = " - [ # ]" Else text_Alfa   = " - [ " & filter_alfa & " ]"
  
  page_Title    = konsol_Add & " " & text_Region & " " & text_Alfa & " - Sida " & pagingOnPage & " - Konsoler"
  page_Header   = konsol_Add & " - Konsoler - Nintendo"
  page_WhereAmI = "&gt; <a href='default.asp' title='G� till &quot;Konsoler&quot; ...'>Konsoler</a> "
  page_SelMenu  = "databas"
  page_Slide    = "konsoler"
  
  page_description  = konsol_Add & " " & text_Region & " till Nintendo listade p� N-Forum.se, Nintendo Forum. Sida " & pagingOnPage & ". " & text_Alfa
  page_keywords     = konsol_Add & ", "
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
  
    <div class="nf_datablock nf_size_full">
      <h1><span class="nf_extitel"><a href="/avdelning/konsol/">Konsoler</a></span><% = konsol_Add %>&nbsp;<% = text_Region %> <% = text_Alfa %></h1>
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
    
      <% If any_Consoles Then %> 
        
        <ul class="nf_list">
          <%
            For zx = pagingBOF To pagingEOF
              If zx > UBound(list_Consoles, 2) Then Exit For
              
              miniBox = 0
              If CLng(list_Consoles(5, zx)) > 0 Then miniBox = list_Consoles(5, zx)
              If CLng(list_Consoles(6, zx)) > 0 Then miniBox = list_Consoles(6, zx)
              If CLng(list_Consoles(4, zx)) > 0 Then miniBox = list_Consoles(4, zx)
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
                      <img src="<% = config_GFXLocation %>icons/flags/<% = CLng(list_Consoles(7, zx)) %>.png" alt="<% = lstRegion(CLng(list_Consoles(7, zx))) %>" title="Region: <% = lstRegion(CLng(list_Consoles(7, zx))) %>">
                      <a href="konsol_visa_info.asp?e=<% = list_Consoles(9, zx) %>" title="<% = sEncode(list_Consoles(1, zx)) %>"><% = sEncode(list_Consoles(1, zx)) %></a>
                    </h4>
                    <span class="nf_medium nf_gray nf_bold"><% = lstKonsol(list_Consoles(2, zx)) %></span>
                  </div>
                  <div class="nf_extend">
                    <% If CONST_LOGIN Then %>
                      <img src="<% = config_GFXLocation %>icons/plus_lrg.png" style="float: right; cursor: pointer;" alt="+" title="L�gg till titeln i din samling." onclick="OpenCollection('console',<% = list_Consoles(9, zx) %>,0,'list')">
                    <% Else %>
                      <img src="<% = config_GFXLocation %>icons/plus_lrg_bw.png" style="float: right;" alt="+" title="Du m�ste vara inloggad f�r att kunna lista dina konsoler.">
                    <% End If %>
                    <% If CONST_LOGIN Then %><img src="<% = config_GFXLocation %>icons/listed.gif" style="display: <% If CLng(list_Consoles(10, zx)) = 0 Then Response.Write("none") Else Response.Write("block") %>;" id="listicon_<% = list_Consoles(9, zx) %>" alt="LIST" title="Titeln finns i din samling."><% End If %>
                  </div>
                </li>
              <%
            Next
          %>
        </ul>
        
        <div class="nf_paging">
          <a href="default.asp?page=<% = pagingOnPage - 1 %>&amp;k=<% = filter_konsol %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %><% = filter_all %>">��</a> |
          
            <% For Each zx In pagingPages %>
              <% If zx = "..." Then %>
                ... |
              <% Else %>
                <a href="default.asp?page=<% = zx %>&amp;k=<% = filter_konsol %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %><% = filter_all %>" <% If CLng(zx) = CLng(pagingOnPage) Then Response.Write(" class='c'") %>><% = zx %></a> <% If CLng(zx) < pagingNumOfPages Then %> | <% End If %>
              <% End If %>
            <% Next %>
            
          | <a href="default.asp?page=<% = pagingOnPage + 1 %>&amp;k=<% = filter_konsol %>&amp;alfa=<% = filter_alfa %>&amp;region=<% = filter_region %><% = filter_all %>">��</a>
        </div>
      <% Else %>
        <div class="nf_msg nf_red">
          <p style="text-align: center;"><strong>Det finns inga konsoler att visa med aktuella val.</strong></p>
        </div>
      <% End If %>
      
    </div>
    
    <div class="nf_datablock nf_size_onethird">
      
      <!--#INCLUDE FILE="../../__INC/_signup.asp"-->
    
      <!-- ### S�KRUTA ### -->
      <div class="nf_minibox nf_green">
        <h4>S�k efter konsol</h4>
        <div class="nf_inside">
          <form action="sokkonsol.asp" method="GET">
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
            <div class="nf_selfinput_full"><input class="btn" type="submit" value="S�k..."></div>
          </form>
        </div>
      </div>
      <!-- ### /S�KRUTA ### -->
    
      <div class="nf_minibox nf_blue">
        <h4>Ikonf�rklaring</h4>
        <div class="nf_inside">
          <p> <img src="<% = config_GFXLocation %>icons/listed.gif"> Titeln finns i din samling </p>
        </div>
      </div>
      
      <div class="nf_minibox nf_blue">
        <h4>L�gg till i din samling</h4>
        <div class="nf_inside">
          <% If CONST_LOGIN Then %>
            <p>F�r att l�gga till en konsol i din samling klickar du bara p� plusset till h�ger om titeln.</p>
            <p>Du kan lista samma titel flera g�nger.</p>
          <% Else %>
            <p style="text-align: center;"><em>Du m�ste <strong><a href="/avdelning/medlem/loggain.asp">logga in</a></strong> f�r att kunna lista dina konsoler.</em></p>
            <p style="text-align: center;"><em>Om du inte redan har en anv�ndare kan du <strong><a href="/avdelning/medlem/registreradig.asp">registrera dig</a></strong>.</em></p>
          <% End If %>
        </div>
      </div>
        
    </div>

  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->