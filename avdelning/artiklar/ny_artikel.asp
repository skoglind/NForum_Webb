<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## Hämta all data ##
  lID  = GetQ("e", "123", 0)

  errCode = GetQ("fail","123",0)
  
  If Session.Value("record_artikel") Then  ' ## POSTBACK
    text_Titel            = sEncode(Session.Value("artikel_aTitel"))
    text_TextM            = sEncode(Session.Value("artikel_aTextM"))
    text_Konsol           = CLng(Session.Value("artikel_aKonsol"))
  End If
  
  Call stop_Rec2Session("artikel")
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Skriv en artikel - Artiklar"
  page_Header   = "Skriv en artikel"
  page_WhereAmI = "&gt; <a href='default.asp' title='Gå till &quot;Artiklar&quot; ...'>Artiklar</a> " & _
                  "&gt; Skriv"
  page_SelMenu  = "texter"
  page_Slide    = "artiklar"
  
  page_description    = "Du kan här skriva din helt egna artikel och få den publicerad på N-Forum.se, Nintendo Forum."
  page_keywords       = "skriv artikel, "
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <!-- ## SKAPA TEXTEN ## -->
    <div id="edit_preview" style="display: block;">
      <div class="nf_datablock nf_size_full">
        <h1><span class="nf_extitel"><a href="/avdelning/artiklar/">Artiklar</a></span>Ny artikel</h1> <a name="EDIT"> </a>
      </div>
    
      <div class="nf_datablock nf_size_full">
        <% If errCode > 0 Then %>
          <div class="nf_msg_full nf_red">
            <% If errCode = 1 Then %><p><strong>Artikeln sparades inte!</strong></p><p>Du har inte angett en nog lång text, den ska innehålla något också.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Artikeln sparades inte!</strong></p><p>Du har angett någon konsol.</p><% End If %>
            <% If errCode = 3 Then %><p><strong>Artikeln sparades inte!</strong></p><p>Du har angett en titel.</p><% End If %>
          </div>
        <% End If %>
      </div>
      
      <div class="nf_datablock nf_size_full">
        <div class="nf_forum">
          <form method="POST" action="_action/postart.asp" id="txtPost">
            <div class="nf_post_msg">
              <label>Titel:</label> <input class="text" type="text" name="aTitel" value="<% = text_Titel %>" maxlength=255>
              
              <label>Konsol:</label>
              <select name="aKonsol" id="aKonsol">
                <option value=0 style="padding: 1px 0 1px 0; font-weight: bold; color: #CCC;"> - Välj konsol - </option>
                <option disabled value=-1 style="border-bottom: dotted 1px #AAA; font-size: 0; height: 1px; margin-bottom: 1px;"> </option>
                <% For zx = 1 To lstKonsol(0) %>
                  <option value=<% = zx %> style="padding: 1px 0 1px 10px;" <% If CLng(text_Konsol) = zx Then Response.Write(" selected") %>> <% = lstKonsol(zx) %> </option>
                <% Next %>
              </select>
            
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