<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## Hämta all data ##
  lID  = GetQ("e", "123", 0)

  errCode = GetQ("fail","123",0)
  
  If HasAcc(CONST_CMS_RIGHTS,"CMS700") Then
    RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksID = " & CLng(lID), False
  Else
    RS_Open 1, "SELECT * FROM cms_KopSalj WHERE ksSkapadAv = " & CLng(CONST_USERID) & " And ksID = " & CLng(lID), False
  End If
    If Not rsDB(1).EOF Then
      text_ID       = rsDB(1)("ksID")
    
      text_Titel    = rsDB(1)("ksTitel")
      text_TextM    = rsDB(1)("ksTextM")
      text_Typ      = rsDB(1)("ksTyp")
      text_Kategori = rsDB(1)("ksKategori1")
      text_Synlig   = rsDB(1)("ksSynlig")
      
      If rsDB(1)("ksStatus") = 1 Then text_Sold = True Else text_Sold = False
      
      sMetod = "Redigera"
    Else
      text_ID       = 0
      
      sMetod = "Ny"
    End If 
  RS_Close 1
  
  If Session.Value("record_annons") Then  ' ## POSTBACK
    text_ID               = CLng(Session.Value("annons_e"))
    
    text_Titel            = sEncode(Session.Value("annons_aTitel"))
    text_TextM            = sEncode(Session.Value("annons_aTextM"))
    text_Typ              = CLng(Session.Value("annons_aTyp"))
    text_Kategori         = CLng(Session.Value("annons_aKategori"))
    text_Sold             = Session.Value("annons_aSold")
    text_Synlig           = Session.Value("annons_aSynlig")
  End If
  
  Call stop_Rec2Session("annons")
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Lägg upp en ny annons - Annonser"
  page_Header   = "Lägg upp annons"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Marknad&quot; ...'>Marknad</a> " & _
                  "&gt; Lägg upp ny"
  page_SelMenu  = "buy"
  page_Slide    = "annonser"
  
  page_description    = "Skapa din egen annons i vår köp- och sälj-avdelning på N-Forum.se, Nintendo Forum."
  page_keywords       = "skapa annons, "
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <!-- ## SKAPA TEXTEN ## -->
    <div id="edit_preview" style="display: block;">
      <div class="nf_datablock nf_size_full">
        <h1><span class="nf_extitel"><a href="/avdelning/annonser/">Annonser</a></span><% = sMetod %> annons</h1> <a name="EDIT"> </a>
      </div>
    
      <div class="nf_datablock nf_size_full">
        <% If errCode > 0 Then %>
          <div class="nf_msg_full nf_red">
            <% If errCode = 1 Then %><p><strong>Annonsen sparades inte!</strong></p><p>Du har inte angett en nog lång text, den ska innehålla något också.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Annonsen sparades inte!</strong></p><p>Du har angett en kategori.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Annonsen sparades inte!</strong></p><p>Du har angett någon typ.</p><% End If %>
            <% If errCode = 4 Then %><p><strong>Annonsen sparades inte!</strong></p><p>Du har angett en titel.</p><% End If %>
          </div>
        <% End If %>
      </div>
      
      <div class="nf_datablock nf_size_full">
        <div class="nf_forum">
          <form method="POST" action="_action/saveannons.asp" id="txtPost">
            <div class="nf_post_msg">
              <label>Titel:</label> <input class="text" type="text" name="aTitel" value="<% = text_Titel %>" maxlength=255>
              <label>Typ:</label>
              <select name="aTyp" id="aTyp">
                <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> - Välj typ - </option>
                <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                <% For zx = 1 To lstKSTyp(-1) %>
                  <option value=<% = zx %> style="padding: 1px 0 1px 10px;" <% If CLng(text_Typ) = zx Then Response.Write(" selected") %>> <% = lstKSTyp(zx) %> </option>
                <% Next %>
              </select>
              
              <label>Kategori:</label>
              <select name="aKategori" id="aKategori">
                <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> - Välj kategori - </option>
                <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                <% For zx = 1 To lstKSKategori(-1) %>
                  <option value=<% = zx %> style="padding: 1px 0 1px 10px;" <% If CLng(text_Kategori) = zx Then Response.Write(" selected") %>> <% = lstKSKategori(zx) %> </option>
                <% Next %>
              </select>
            
              <label>Såld!</label> <input type="checkbox" class="checkbox" name="aSold" id="aSold" value="YES" <% If text_Sold Then Response.Write(" checked") %>>
              <label>Synlig</label> <input type="checkbox" class="checkbox" name="aSynlig" id="aSynlig" value="YES" <% If text_Synlig Then Response.Write(" checked") %>>
            
              <div class="nf_buttonbar">
                <input onclick="addText('aTextM','b');" type="button" value="B" style="width: 25px; font-weight: bold;">
                <input onclick="addText('aTextM','i');" type="button" value="I" style="width: 25px; font-style: italic;">
                <input onclick="addText('aTextM','u');" type="button" value="U" style="width: 25px; font-decoration: underline;">
                <input onclick="addText('aTextM','s');" type="button" value="S" style="width: 25px; text-decoration: line-through;">
                <div class="nf_buttonsplit"></div>
                <input onclick="addText('aTextM','url');" type="button" value="URL" style="width: 40px;">
                <input onclick="addText('aTextM','img');" type="button" value="IMG" style="width: 40px;">
                <!--
                <div class="nf_buttonsplit"></div>
                <input onclick="addText('aTextM','spoiler');" type="button" value="Spoiler" style="width: 56px;">
                <input onclick="addText('aTextM','indent');" type="button" value="Indenterad" style="width: 80px;">
                <input onclick="addText('aTextM','code');" type="button" value="Monospace" style="width: 80px;">
                -->
              </div>
              
              <textarea name="aTextM" id="aTextM" maxlength="20000" style="height: 600px;" onkeyup="if(this.value==''){document.getElementById('btPreview').disabled=true;}else{document.getElementById('btPreview').disabled=false;}closePreview('fMsg',false);"><% = text_TextM %></textarea>
              
              <input type="hidden" name="e" value="<% = text_ID %>">
            </div>
            <div class="nf_post_btn">
              <input type="submit" style="font-weight: bold;" value="Skicka in...">
              <input type="button" <% If Len(text_TextM) = 0 Then %>disabled<% End If %> id="btPreview" value="Förhandsgranska" onclick="doPreview('aTextM','YES','NO');">
            </div>
          </form>
          
        </div>
      </div>
    </div>
    <!-- ## /SKAPA TRÅDEN ## -->
  
    <!-- ## FÖRHANDSGRANSKA TRÅDEN ## -->
    <div id="post_preview" style="display: none;">
      <div class="nf_datablock nf_size_full">
        <h1>Förhandsgranskning</h1> <a name="PREVIEW"> </a>
      </div>
      
      <div class="nf_datablock nf_size_full">
        <div class="nf_msg_full nf_green" id="warn_preview" style="display: none;">
          <p><img src="<% = config_GFXLocation %>preview_arrow.png" style="float: left; margin-right: 8px;"><strong>Din förhandsgranskning, klicka på [Ändra] för att fortsätta redigera.</strong></p>
        </div>
        
        <div class="nf_forum">
          <div class="nf_post_msg">
            <div class="nf_post_prevtxt">
              <p id="post_preview_text"></p>
            </div>
          </div>
          <div class="nf_post_btn">
            <input type="button" style="font-weight: bold;" value="Skicka in..." onclick="document.getElementById('txtPost').submit();">
            <input type="button" value="Ändra" onclick="closePreview('aTextM',true);">
          </div>
        </div>
      </div>
      
    </div>
    <!-- ## FÖRHANDSGRANSKA TRÅDEN ## -->
    
  </div>

<!--#INCLUDE FILE="../../_page_bottom.asp"-->
<!--#INCLUDE FILE="../../__INC/includes_end.asp"-->