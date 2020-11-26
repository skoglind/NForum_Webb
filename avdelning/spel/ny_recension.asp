<!--#INCLUDE FILE="../../__INC/includes.asp"-->

<% If Not CONST_LOGIN Then Response.Redirect(config_NotLoggedIn) %>

<%
  ' ## Hämta all data ##
  lID  = GetQ("e", "123", 0)

  errCode = GetQ("fail","123",0)
  
  RS_Open 1, "SELECT *, cms_Spel.sKonsol AS sKonsol FROM cms_Speltitlar " & _
             "LEFT JOIN cms_Spel ON tSpelID = cms_Spel.sID " & _
             "WHERE tID = " & CLng(lID), False

    If rsDB(1).EOF Then Response.Redirect("spel.asp")

    text_Titel  = sEncode(rsDB(1)("tTitel"))
    text_Konsol = lstKonsolShort(rsDB(1)("sKonsol"))
    text_KonsolID2 = rsDB(1)("sKonsol")
    text_Region = FixNum(rsDB(1)("tRegion"))
    text_SpelID = FixNum(rsDB(1)("tSpelID"))
  
  RS_Close 1
  
  If Session.Value("record_recension") Then  ' ## POSTBACK
    text_TextM            = sEncode(Session.Value("recension_rTextM"))
    text_Betyg            = CLng(Session.Value("recension_rBetyg"))
  End If
  
  Call stop_Rec2Session("recension")
%>

<%
  ' ## Globala variabler ##
  page_Title    = "Skriv en recension - Spel"
  page_Header   = "Skriv en recension"
  page_WhereAmI = "&gt; <a href='spel.asp' title='Gå till &quot;Spel&quot; ...'>Spel</a> " & _
                  "&gt; <a href='spel_visa_info.asp?e=" & lID & "'>" & text_Titel & "</a> " & _
                  "&gt; <a href='spel_visa_recensioner.asp?e=" & lID & "'>Recensioner</a> " & _
                  "&gt; Skriv"
  page_SelMenu  = "texter"
  page_Slide    = "recensioner"
  
  iFilter       = text_Forum
  
  page_description  = "Skriv en recension för " & text_Titel & " till " & text_Konsol & " på N-Forum.se, Nintendo Forum."
  page_keywords     = "skriv recension, "
  
%>

<!--#INCLUDE FILE="../../_page_top.asp"-->
  <!--#INCLUDE FILE="__menu.asp"-->
<!--#INCLUDE FILE="../../_page_middle.asp"-->

  <div class="content">
    
    <!-- ## SKAPA TEXTEN ## -->
    <div id="edit_preview" style="display: block;">
      <div class="nf_datablock nf_size_full">
        <h1><span class="nf_extitel"><a href="/avdelning/recensioner/">Recensioner</a></span>Ny recension</h1> <a name="EDIT"> </a>
      </div>
    
      <div class="nf_datablock nf_size_full">
        <% If errCode > 0 Then %>
          <div class="nf_msg_full nf_red">
            <% If errCode = 1 Then %><p><strong>Recensionen sparades inte!</strong></p><p>Du har inte angett en nog lång text, den ska innehålla något också.</p><% End If %>
            <% If errCode = 2 Then %><p><strong>Recensionen sparades inte!</strong></p><p>Du har inte satt något betyg.</p><% End If %>
          </div>
        <% End If %>
      </div>
      
      <div class="nf_datablock nf_size_full">
        <div class="nf_forum">
          <form method="POST" action="_action/postrec.asp" id="txtPost">
            <div class="nf_post_msg">
              <label>Titel:</label> <input class="text" type="text" value="<% = text_Titel %>" maxlength=255 style="width: 436px;" disabled>
              <label>Konsol:</label> <input class="text" type="text" value="<% = text_Konsol %>" maxlength=255 style="width: 436px;" disabled>
            
              <label>Betyg:</label>
              <div style="float: left; width: 550px; padding: 4px;">
                <div class="nf_picker"><input type="radio" name="rBetyg" value=1 <% If text_Betyg = 1 Then Response.Write(" checked") %>><span>1</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=2 <% If text_Betyg = 2 Then Response.Write(" checked") %>><span>2</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=3 <% If text_Betyg = 3 Then Response.Write(" checked") %>><span>3</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=4 <% If text_Betyg = 4 Then Response.Write(" checked") %>><span>4</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=5 <% If text_Betyg = 5 Then Response.Write(" checked") %>><span>5</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=6 <% If text_Betyg = 6 Then Response.Write(" checked") %>><span>6</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=7 <% If text_Betyg = 7 Then Response.Write(" checked") %>><span>7</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=8 <% If text_Betyg = 8 Then Response.Write(" checked") %>><span>8</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=9 <% If text_Betyg = 9 Then Response.Write(" checked") %>><span>9</span></div>
                <div class="nf_picker"><input type="radio" name="rBetyg" value=10 <% If text_Betyg = 10 Then Response.Write(" checked") %>><span>10</span></div>
              </div>
              
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
              
              <textarea name="rTextM" id="aTextM" maxlength="20000" style="height: 600px;" onkeyup="if(this.value==''){document.getElementById('btPreview').disabled=true;}else{document.getElementById('btPreview').disabled=false;}closePreview('fMsg',false);"><% = text_TextM %></textarea>
            </div>
            <div class="nf_post_btn">
              <input type="hidden" name="e" value=<% = lID %>>
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